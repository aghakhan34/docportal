/* =====================================================================
   dataProviders.js
   The data layer behind the portal. The UI (app.js) never talks to
   SharePoint directly — it talks to a "provider" with a fixed interface,
   so the same UI works against demo data or a live SharePoint library.

   Provider interface (all async unless noted):
     init()                         -> { mode, account, error }
     signIn() / signOut()           -> connection state
     list()                         -> [doc, ...]
     upload(file, meta, onProgress) -> doc
     remove(doc)                    -> void
     download(doc)                  -> void   (triggers a browser download)
     setStar(doc, starred)          -> void
     get isLive()                   -> boolean

   Normalized "doc" shape used everywhere in the UI:
     { id, name, division(key), dept, status, size(MB),
       date(ISO), starred, uploadedBy, webUrl, _sp{ listItemId, driveItemId, downloadUrl } }
   ===================================================================== */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const FOUR_MB = 4 * 1024 * 1024;
const CHUNK = 12 * 327680; // 3.75 MB — Graph requires upload chunks to be a multiple of 320 KB

/* Helper: build {label <-> key} maps from the configured taxonomy */
function buildDivisionMaps(divisions) {
  const labelToKey = {}, keyToLabel = {};
  Object.entries(divisions).forEach(([key, d]) => {
    labelToKey[d.label.toLowerCase()] = key;
    keyToLabel[key] = d.label;
  });
  return { labelToKey, keyToLabel };
}

/* =====================================================================
   MSAL authentication wrapper
   ===================================================================== */
class GraphAuth {
  constructor(authCfg) {
    this.cfg = authCfg;
    this.account = null;
    this.msal = new msal.PublicClientApplication({
      auth: {
        clientId: authCfg.clientId,
        authority: `https://login.microsoftonline.com/${authCfg.tenantId}`,
        redirectUri: authCfg.redirectUri,
      },
      cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
    });
  }

  async init() {
    await this.msal.initialize();
    // Complete any redirect-based sign-in that is in flight.
    const result = await this.msal.handleRedirectPromise();
    if (result && result.account) this.account = result.account;
    else {
      const accounts = this.msal.getAllAccounts();
      if (accounts.length) this.account = accounts[0];
    }
    return this.account;
  }

  async signIn() {
    try {
      const res = await this.msal.loginPopup({ scopes: this.cfg.scopes, prompt: 'select_account' });
      this.account = res.account;
    } catch (e) {
      // Popups blocked? fall back to a full-page redirect.
      if (e instanceof msal.BrowserAuthError && e.errorCode === 'popup_window_error') {
        await this.msal.loginRedirect({ scopes: this.cfg.scopes });
        return null;
      }
      throw e;
    }
    return this.account;
  }

  async signOut() {
    if (!this.account) return;
    await this.msal.logoutPopup({ account: this.account });
    this.account = null;
  }

  /* Silent-first token acquisition with interactive fallback. */
  async token() {
    if (!this.account) throw new Error('Not signed in');
    try {
      const res = await this.msal.acquireTokenSilent({ scopes: this.cfg.scopes, account: this.account });
      return res.accessToken;
    } catch (e) {
      const res = await this.msal.acquireTokenPopup({ scopes: this.cfg.scopes });
      this.account = res.account;
      return res.accessToken;
    }
  }
}

/* =====================================================================
   SharePoint provider (Microsoft Graph)
   ===================================================================== */
class SharePointProvider {
  constructor(cfg) {
    this.cfg = cfg;
    this.auth = new GraphAuth(cfg.auth);
    this.maps = buildDivisionMaps(cfg.divisions);
    this.cols = cfg.columns;
    this.siteId = null;
    this.listId = null;
    this.driveId = null;
  }

  get isLive() { return !!this.auth.account; }

  async init() {
    const account = await this.auth.init();
    if (account) {
      try { await this._resolveTarget(); }
      catch (e) { return { mode: 'sharepoint', account, error: this._msg(e) }; }
    }
    return { mode: 'sharepoint', account };
  }

  async signIn() {
    const account = await this.auth.signIn();
    if (account) await this._resolveTarget();
    return { mode: 'sharepoint', account };
  }

  async signOut() {
    await this.auth.signOut();
    this.siteId = this.listId = this.driveId = null;
    return { mode: 'sharepoint', account: null };
  }

  /* ---- low-level fetch helper ---- */
  async _api(path, { method = 'GET', headers = {}, body, raw = false, absolute = false } = {}) {
    const token = await this.auth.token();
    const url = absolute ? path : GRAPH + path;
    const res = await fetch(url, {
      method,
      headers: { Authorization: `Bearer ${token}`, ...headers },
      body,
    });
    if (!res.ok) {
      let detail = '';
      try { detail = (await res.json())?.error?.message || ''; } catch {}
      throw new Error(`Graph ${method} ${res.status}: ${detail || res.statusText}`);
    }
    if (raw) return res;
    if (res.status === 204) return null;
    return res.json();
  }

  /* ---- resolve site -> list -> drive once, then cache ---- */
  async _resolveTarget() {
    if (this.siteId && this.listId && this.driveId) return;
    const u = new URL(this.cfg.sharePoint.siteUrl);
    const host = u.hostname;
    const sitePath = u.pathname.replace(/\/$/, ''); // e.g. /sites/Intranet
    const site = await this._api(`/sites/${host}:${sitePath}`);
    this.siteId = site.id;

    const name = this.cfg.sharePoint.libraryName.replace(/'/g, "''");
    const lists = await this._api(
      `/sites/${this.siteId}/lists?$expand=drive&$filter=displayName eq '${encodeURIComponent(name)}'`
    );
    const lib = (lists.value || []).find(l => l.drive) || (lists.value || [])[0];
    if (!lib) throw new Error(`Library "${this.cfg.sharePoint.libraryName}" not found on this site.`);
    this.listId = lib.id;
    this.driveId = lib.drive.id;
  }

  /* ---- READ ---- */
  async list() {
    await this._resolveTarget();
    const data = await this._api(
      `/sites/${this.siteId}/lists/${this.listId}/items` +
      `?$expand=fields,driveItem&$top=500&$orderby=lastModifiedDateTime desc`
    );
    return (data.value || [])
      .filter(it => it.driveItem && it.driveItem.file) // files only, skip folders
      .map(it => this._normalize(it));
  }

  _normalize(item) {
    const di = item.driveItem;
    const f = item.fields || {};
    const divLabel = f[this.cols.division] || '';
    const division = this.maps.labelToKey[String(divLabel).toLowerCase()] || divLabel;
    return {
      id: di.id,
      name: di.name,
      division,
      dept: f[this.cols.department] || '',
      status: f[this.cols.status] || this.cfg.defaultStatus,
      size: +(di.size / (1024 * 1024)).toFixed(2),
      date: di.lastModifiedDateTime,
      starred: !!f[this.cols.starred],
      uploadedBy: di.createdBy?.user?.displayName || 'Unknown',
      webUrl: di.webUrl,
      selected: false,
      _sp: { listItemId: item.id, driveItemId: di.id, downloadUrl: di['@microsoft.graph.downloadUrl'] },
    };
  }

  /* ---- CREATE / UPLOAD ---- */
  async upload(file, meta, onProgress) {
    await this._resolveTarget();
    const driveItem = file.size < FOUR_MB
      ? await this._uploadSmall(file, onProgress)
      : await this._uploadLarge(file, onProgress);

    // Find the list item behind the new file, then write the metadata columns.
    const li = await this._api(`/sites/${this.siteId}/drives/${this.driveId}/items/${driveItem.id}/listItem`);
    const fields = this._metaToFields(meta);
    await this._api(`/sites/${this.siteId}/lists/${this.listId}/items/${li.id}/fields`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(fields),
    });
    if (onProgress) onProgress(1);

    return this._normalize({ id: li.id, driveItem, fields });
  }

  _metaToFields(meta) {
    const out = {};
    out[this.cols.division]   = this.maps.keyToLabel[meta.division] || meta.division;
    out[this.cols.department] = meta.dept;
    out[this.cols.status]     = meta.status;
    return out;
  }

  async _uploadSmall(file, onProgress) {
    if (onProgress) onProgress(0.3);
    const path = encodeURIComponent(file.name);
    const di = await this._api(
      `/sites/${this.siteId}/drives/${this.driveId}/root:/${path}:/content`,
      { method: 'PUT', headers: { 'Content-Type': file.type || 'application/octet-stream' }, body: file }
    );
    if (onProgress) onProgress(0.9);
    return di;
  }

  async _uploadLarge(file, onProgress) {
    const path = encodeURIComponent(file.name);
    const session = await this._api(
      `/sites/${this.siteId}/drives/${this.driveId}/root:/${path}:/createUploadSession`,
      { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) }
    );
    let start = 0, last = null;
    while (start < file.size) {
      const end = Math.min(start + CHUNK, file.size);
      const chunk = file.slice(start, end);
      // Upload-session chunks are unauthenticated PUTs to a pre-signed URL.
      const res = await fetch(session.uploadUrl, {
        method: 'PUT',
        headers: { 'Content-Length': String(end - start), 'Content-Range': `bytes ${start}-${end - 1}/${file.size}` },
        body: chunk,
      });
      if (!res.ok) throw new Error(`Upload failed at byte ${start} (${res.status})`);
      if (onProgress) onProgress(end / file.size);
      if (res.status === 200 || res.status === 201) last = await res.json();
      start = end;
    }
    return last;
  }

  /* ---- DELETE ---- */
  async remove(doc) {
    await this._resolveTarget();
    await this._api(`/sites/${this.siteId}/drives/${this.driveId}/items/${doc._sp.driveItemId}`, { method: 'DELETE' });
  }

  /* ---- DOWNLOAD ---- */
  async download(doc) {
    let url = doc._sp.downloadUrl;
    if (!url) {
      const di = await this._api(`/sites/${this.siteId}/drives/${this.driveId}/items/${doc._sp.driveItemId}`);
      url = di['@microsoft.graph.downloadUrl'];
    }
    const a = document.createElement('a');
    a.href = url; a.download = doc.name;
    document.body.appendChild(a); a.click(); a.remove();
  }

  /* ---- STAR (writes the optional Yes/No column) ---- */
  async setStar(doc, starred) {
    await this._resolveTarget();
    const body = {}; body[this.cols.starred] = starred;
    await this._api(`/sites/${this.siteId}/lists/${this.listId}/items/${doc._sp.listItemId}/fields`, {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
    });
  }

  _msg(e) {
    const m = String(e.message || e);
    if (m.includes('404')) return 'Site or library not found — check siteUrl and libraryName in config.js.';
    if (m.includes('403')) return 'Access denied — the signed-in account lacks permission, or admin consent is pending.';
    if (m.includes('column') || m.includes('field')) return 'A metadata column is missing — check the columns in config.js against the library.';
    return m;
  }
}

/* =====================================================================
   Demo provider (in-memory, no sign-in) — mirrors the original prototype
   ===================================================================== */
class DemoProvider {
  constructor(cfg) {
    this.cfg = cfg;
    this.maps = buildDivisionMaps(cfg.divisions);
    this._id = 0;
    this.docs = [
      ['Q4 Financial Report 2024.pdf','finance','Accounting','Published',2.4,'2025-03-12'],
      ['HR Policy Handbook v3.docx','hr','Employee Relations','Published',1.1,'2025-01-20'],
      ['IT Security Guidelines.pdf','it','Cybersecurity','Review',0.8,'2025-04-02'],
      ['Supply Chain Dashboard.xlsx','operations','Supply Chain','Published',3.2,'2025-02-14'],
      ['Brand Identity Guide.pdf','corporate','Communications','Published',5.7,'2024-11-30'],
      ['GDPR Compliance Checklist.docx','legal','Data Privacy','Review',0.5,'2025-03-28'],
      ['Annual Strategy Deck.pptx','corporate','Strategy','Draft',4.2,'2025-04-10'],
      ['Payroll Summary Jan-Mar.xlsx','finance','Accounting','Published',1.8,'2025-04-01'],
      ['Onboarding Checklist.docx','hr','Talent Acquisition','Published',0.3,'2025-01-05'],
      ['Network Architecture v2.png','it','Infrastructure','Draft',1.5,'2025-03-22'],
      ['Risk Register 2025.xlsx','legal','Risk Management','Review',2.1,'2025-02-28'],
      ['Facilities Maintenance Log.docx','operations','Facilities','Published',0.7,'2025-03-18'],
    ].map(([name, division, dept, status, size, date]) => ({
      id: ++this._id, name, division, dept, status, size, date,
      starred: false, selected: false, uploadedBy: 'John Doe',
    }));
  }
  get isLive() { return false; }
  async init()    { return { mode: 'demo', account: null }; }
  async signIn()  { return { mode: 'demo', account: null }; }
  async signOut() { return { mode: 'demo', account: null }; }
  async list()    { return this.docs.map(d => ({ ...d })); }

  async upload(file, meta, onProgress) {
    if (onProgress) onProgress(1);
    const doc = {
      id: ++this._id, name: file.name, division: meta.division, dept: meta.dept,
      status: meta.status, size: +(file.size / (1024 * 1024)).toFixed(2) || 0.1,
      date: new Date().toISOString(), starred: false, selected: false, uploadedBy: 'You (demo)',
    };
    this.docs.unshift(doc);
    return { ...doc };
  }
  async remove(doc)  { this.docs = this.docs.filter(d => d.id !== doc.id); }
  async download(doc){ /* nothing to download in demo mode */ }
  async setStar(doc, starred) { const d = this.docs.find(x => x.id === doc.id); if (d) d.starred = starred; }
}

/* =====================================================================
   Factory
   ===================================================================== */
function createDataProvider(cfg) {
  if (cfg.mode === 'sharepoint') {
    const placeholderRe = /^0{8}-0{4}-0{4}-0{4}-0{12}$/;
    const unset = placeholderRe.test(cfg.auth.clientId) || cfg.sharePoint.siteUrl.includes('YOURTENANT');
    if (unset) {
      console.warn('[DocPortal] mode is "sharepoint" but config.js still has placeholder values — falling back to demo.');
      return new DemoProvider(cfg);
    }
    if (typeof msal === 'undefined') {
      console.error('[DocPortal] MSAL library not loaded — falling back to demo.');
      return new DemoProvider(cfg);
    }
    return new SharePointProvider(cfg);
  }
  return new DemoProvider(cfg);
}

window.createDataProvider = createDataProvider;
