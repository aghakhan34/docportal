/* =====================================================================
   app.js — UI logic for the Document Portal.
   Loads data from a provider (demo or SharePoint) and renders the views.
   ===================================================================== */

const DIVISIONS = CONFIG.divisions;

const FILE_ICONS = {
  pdf:  { cls:'pdf',   icon:'ti-file-type-pdf' },
  docx: { cls:'word',  icon:'ti-file-type-doc' },
  doc:  { cls:'word',  icon:'ti-file-type-doc' },
  xlsx: { cls:'excel', icon:'ti-file-spreadsheet' },
  xls:  { cls:'excel', icon:'ti-file-spreadsheet' },
  pptx: { cls:'ppt',   icon:'ti-presentation' },
  ppt:  { cls:'ppt',   icon:'ti-presentation' },
  png:  { cls:'img',   icon:'ti-photo' },
  jpg:  { cls:'img',   icon:'ti-photo' },
  jpeg: { cls:'img',   icon:'ti-photo' },
  gif:  { cls:'img',   icon:'ti-photo' },
  txt:  { cls:'txt',   icon:'ti-file-text' },
  zip:  { cls:'zip',   icon:'ti-file-zip' },
};

let provider = null;
let docs = [];
let pendingFiles = [];
let currentView = 'list';
let activeFilter = { type: 'all', division: '', dept: '', ext: '', chip: 'all' };
let sortState = { key: 'date', dir: -1 };

/* ── helpers ── */
function getExt(name) { return name.split('.').pop().toLowerCase(); }
function getIcon(name) { return FILE_ICONS[getExt(name)] || { cls:'other', icon:'ti-file' }; }
function fmtSize(mb) { return mb < 1 ? (mb * 1024).toFixed(0) + ' KB' : mb.toFixed(1) + ' MB'; }
function fmtDate(d) {
  const dt = new Date(d);
  return isNaN(dt) ? '—' : dt.toLocaleDateString('en-US', { year:'numeric', month:'short', day:'numeric' });
}
function esc(s) { return String(s).replace(/[&<>"]/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' }[c])); }
function docById(id) { return docs.find(d => String(d.id) === String(id)); }
function setBusy(on) { document.getElementById('loading-bar').classList.toggle('active', on); }

/* ── build the config-driven chrome (sidebar divisions, selectors) ── */
function buildTaxonomyUI() {
  // Sidebar division list
  const sb = document.getElementById('sb-divisions');
  sb.innerHTML = Object.entries(DIVISIONS).map(([k, d]) =>
    `<div class="sidebar-item" data-div="${k}" onclick="filterDivision('${k}')">
       <i class="ti ${d.icon}" aria-hidden="true"></i> ${esc(d.label)}
       <span class="sidebar-count" id="cnt-${k}">0</span>
     </div>`).join('');

  // Division <select>s (context bar + upload modal)
  const opts = '<option value="">— Select Division —</option>' +
    Object.entries(DIVISIONS).map(([k, d]) => `<option value="${k}">${esc(d.label)}</option>`).join('');
  document.getElementById('sel-division').innerHTML = opts;
  document.getElementById('up-division').innerHTML =
    '<option value="">Select Division</option>' +
    Object.entries(DIVISIONS).map(([k, d]) => `<option value="${k}">${esc(d.label)}</option>`).join('');
}

/* ── connection / auth UI ── */
function renderConnection(state) {
  const el = document.getElementById('conn');
  if (!el) return;
  if (state.mode === 'demo') {
    el.className = 'conn-status demo';
    el.innerHTML = `<span class="dot"></span> Demo data`;
  } else if (state.error) {
    el.className = 'conn-status error';
    el.innerHTML = `<span class="dot"></span> ${esc(state.error)} ` +
      `<button class="conn-signin" onclick="doSignIn()" style="margin-left:6px"><i class="ti ti-refresh"></i> Retry</button>`;
  } else if (state.account) {
    el.className = 'conn-status live';
    el.innerHTML = `<span class="dot"></span> ${esc(state.account.name || state.account.username)} ` +
      `<button class="topnav-btn" title="Sign out" onclick="doSignOut()" style="padding:2px 4px;color:#fff"><i class="ti ti-logout"></i></button>`;
    const av = document.getElementById('avatar');
    if (av) av.textContent = initials(state.account.name || state.account.username);
  } else {
    el.className = 'conn-status';
    el.innerHTML = `<span class="dot"></span> Not signed in ` +
      `<button class="conn-signin" onclick="doSignIn()"><i class="ti ti-login"></i> Sign in</button>`;
  }
}
function initials(name) {
  const p = String(name).replace(/@.*/, '').split(/[ .]/).filter(Boolean);
  return ((p[0]?.[0] || '') + (p[1]?.[0] || '')).toUpperCase() || 'U';
}

async function doSignIn() {
  try {
    setBusy(true);
    const state = await provider.signIn();
    renderConnection(state);
    await reload();
  } catch (e) {
    renderConnection({ mode: 'sharepoint', error: 'Sign-in failed' });
    showToast('Sign-in failed: ' + e.message, 'error');
  } finally { setBusy(false); }
}

async function doSignOut() {
  try { await provider.signOut(); } catch {}
  renderConnection({ mode: 'sharepoint', account: null });
  docs = [];
  renderDocs();
  showToast('Signed out');
}

/* ── data load ── */
async function reload() {
  if (!provider.isLive && CONFIG.mode === 'sharepoint') { docs = []; renderDocs(); return; }
  try {
    setBusy(true);
    docs = await provider.list();
    renderDocs();
  } catch (e) {
    showToast('Could not load documents: ' + e.message, 'error');
    renderConnection({ mode: CONFIG.mode, account: provider.auth?.account, error: provider._msg ? provider._msg(e) : e.message });
  } finally { setBusy(false); }
}

/* ── filtering & sorting ── */
function getFilteredDocs() {
  const q = document.getElementById('search-input').value.toLowerCase();
  const now = Date.now();
  return docs.filter(d => {
    if (activeFilter.type === 'starred' && !d.starred) return false;
    if (activeFilter.type === 'recent') {
      const age = now - new Date(d.date).getTime();
      if (!(age <= 30 * 864e5)) return false;
    }
    if (activeFilter.division && d.division !== activeFilter.division) return false;
    if (activeFilter.dept && d.dept !== activeFilter.dept) return false;
    if (activeFilter.ext) {
      const ext = getExt(d.name);
      const map = { pdf:['pdf'], word:['docx','doc'], excel:['xlsx','xls'], image:['png','jpg','jpeg','gif'] };
      if (!map[activeFilter.ext]?.includes(ext)) return false;
    }
    if (activeFilter.chip !== 'all' && d.status !== activeFilter.chip) return false;
    if (q && !d.name.toLowerCase().includes(q) && !(d.dept || '').toLowerCase().includes(q)) return false;
    return true;
  }).sort((a, b) => {
    const k = sortState.key;
    let av = a[k] ?? '', bv = b[k] ?? '';
    if (k === 'size') { av = a.size; bv = b.size; }
    if (k === 'date') { av = new Date(a.date).getTime(); bv = new Date(b.date).getTime(); }
    if (av < bv) return -1 * sortState.dir;
    if (av > bv) return  1 * sortState.dir;
    return 0;
  });
}

/* ── rendering ── */
function renderDocs() {
  const list = getFilteredDocs();
  updateStats(list);
  updateSidebarCounts();

  const tbody = document.getElementById('doc-tbody');
  const emptyState = document.getElementById('empty-state');
  const gridView = document.getElementById('doc-grid-view');

  if (list.length === 0) {
    tbody.innerHTML = '';
    emptyState.style.display = 'block';
    gridView.innerHTML = '';
    return;
  }
  emptyState.style.display = 'none';

  const statusMap = { Published:'published', Draft:'draft', Review:'review' };
  const statusIcons = { Published:'ti-circle-check', Draft:'ti-pencil', Review:'ti-clock' };

  tbody.innerHTML = list.map(d => {
    const ic = getIcon(d.name);
    const divLabel = DIVISIONS[d.division]?.label || d.division || '—';
    const sCls = statusMap[d.status] || 'review';
    const sIcon = statusIcons[d.status] || 'ti-file';
    return `<tr class="${d.selected?'selected':''}">
      <td onclick="event.stopPropagation()"><input type="checkbox" aria-label="Select ${esc(d.name)}" ${d.selected?'checked':''} onchange="toggleRow('${d.id}', this)"></td>
      <td>
        <div class="file-icon">
          <div class="file-icon-badge ${ic.cls}"><i class="ti ${ic.icon}" aria-hidden="true"></i></div>
          <div>
            <button class="file-name" onclick="openDoc('${d.id}')">${esc(d.name)}</button>
            <div class="file-meta">Uploaded by ${esc(d.uploadedBy || '—')} · ${fmtDate(d.date)}</div>
          </div>
        </div>
      </td>
      <td><span class="dept-tag">${esc(divLabel)}</span></td>
      <td style="font-size:12px;color:var(--text-muted)">${esc(d.dept || '—')}</td>
      <td><span class="status-badge ${sCls}"><i class="ti ${sIcon}" aria-hidden="true"></i>${esc(d.status)}</span></td>
      <td style="font-size:12px;color:var(--text-muted);white-space:nowrap">${fmtDate(d.date)}</td>
      <td style="font-size:12px;color:var(--text-muted)">${fmtSize(d.size)}</td>
      <td>
        <div class="action-menu">
          <button class="action-btn" title="Download" onclick="event.stopPropagation();downloadDoc('${d.id}')"><i class="ti ti-download" aria-hidden="true"></i></button>
          <button class="action-btn" title="${d.starred?'Remove star':'Add star'}" onclick="event.stopPropagation();toggleStar('${d.id}')"><i class="ti ti-star${d.starred?'-filled':''}" style="${d.starred?'color:#f8a800':''}" aria-hidden="true"></i></button>
          <button class="action-btn" title="Delete" onclick="event.stopPropagation();deleteDoc('${d.id}')" style="color:var(--danger)"><i class="ti ti-trash" aria-hidden="true"></i></button>
        </div>
      </td>
    </tr>`;
  }).join('');

  gridView.innerHTML = list.map(d => {
    const ic = getIcon(d.name);
    return `<div class="doc-card" title="${esc(d.name)}" onclick="openDoc('${d.id}')">
      <div class="file-icon-badge ${ic.cls}" style="width:44px;height:44px;font-size:24px"><i class="ti ${ic.icon}" aria-hidden="true"></i></div>
      <div class="doc-card-name">${esc(d.name)}</div>
      <div class="doc-card-meta">${fmtSize(d.size)} · ${fmtDate(d.date)}</div>
    </div>`;
  }).join('');
}

function updateStats(list) {
  document.getElementById('st-total').textContent = list.length;
  document.getElementById('st-pub').textContent = list.filter(d=>d.status==='Published').length;
  document.getElementById('st-rev').textContent = list.filter(d=>d.status==='Review').length;
  const total = list.reduce((s,d)=>s+d.size,0);
  document.getElementById('st-size').textContent = total < 1 ? (total*1024).toFixed(0)+' KB' : total.toFixed(1)+' MB';
  document.getElementById('cnt-all').textContent = docs.length;
}

function updateSidebarCounts() {
  Object.keys(DIVISIONS).forEach(k => {
    const el = document.getElementById('cnt-' + k);
    if (el) el.textContent = docs.filter(d=>d.division===k).length;
  });
}

/* ── selectors / breadcrumb ── */
function onDivisionChange() {
  const val = document.getElementById('sel-division').value;
  const dept = document.getElementById('sel-dept');
  dept.disabled = !val;
  dept.innerHTML = '<option value="">All Departments</option>';
  if (val && DIVISIONS[val]) DIVISIONS[val].depts.forEach(d => dept.innerHTML += `<option value="${esc(d)}">${esc(d)}</option>`);
  activeFilter.division = val;
  activeFilter.dept = '';
  syncSidebarActive();
  updateBreadcrumb(); updateContextBadges(); renderDocs();
}
function onDeptChange() {
  activeFilter.dept = document.getElementById('sel-dept').value;
  updateBreadcrumb(); updateContextBadges(); renderDocs();
}
function updateBreadcrumb() {
  const div = document.getElementById('sel-division').value;
  const dept = document.getElementById('sel-dept').value;
  let path = 'All Documents';
  if (div) path = DIVISIONS[div].label;
  if (dept) path += ' / ' + dept;
  document.getElementById('bc-path').textContent = path;
  document.getElementById('content-title').textContent = dept || (div ? DIVISIONS[div].label : 'All Documents');
  document.getElementById('content-subtitle').textContent = dept
    ? `Documents in ${dept}`
    : (div ? `All documents in ${DIVISIONS[div].label}` : 'Select a division and department above to filter documents');
}
function updateContextBadges() {
  const div = document.getElementById('sel-division').value;
  const dept = document.getElementById('sel-dept').value;
  let html = '';
  if (div) html += `<span class="ctx-badge"><i class="ti ti-building" aria-hidden="true"></i>${esc(DIVISIONS[div].label)}</span>`;
  if (dept) html += `<span class="ctx-badge"><i class="ti ti-users" aria-hidden="true"></i>${esc(dept)}</span>`;
  document.getElementById('ctx-badges').innerHTML = html;
}
function syncSidebarActive() {
  const div = activeFilter.division;
  document.querySelectorAll('.sidebar-item[data-div]').forEach(el =>
    el.classList.toggle('active', el.getAttribute('data-div') === div));
}

/* ── sidebar filters ── */
function filterDivision(div) {
  document.getElementById('sel-division').value = div;
  onDivisionChange();
}
function filterType(t) {
  activeFilter.type = t;
  document.querySelectorAll('.sidebar-item[data-type]').forEach(el =>
    el.classList.toggle('active', el.getAttribute('data-type') === t));
  document.querySelectorAll('.sidebar-item[data-div]').forEach(el => el.classList.remove('active'));
  renderDocs();
}
function filterExt(ext) { activeFilter.ext = activeFilter.ext === ext ? '' : ext; renderDocs(); }
function chipFilter(val, ev) {
  activeFilter.chip = val;
  document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
  (ev?.currentTarget)?.classList.add('active');
  renderDocs();
}
function sortBy(key) {
  if (sortState.key === key) sortState.dir *= -1;
  else { sortState.key = key; sortState.dir = 1; }
  renderDocs();
}
function sortDocs() { sortState.dir *= -1; renderDocs(); showToast('Sort order updated'); }
function openFilterMenu() {
  const bar = document.getElementById('filter-bar');
  bar.style.display = bar.style.display === 'none' ? 'flex' : 'none';
}
function setView(v) {
  currentView = v;
  document.getElementById('doc-list-view').style.display = v === 'list' ? 'block' : 'none';
  document.getElementById('doc-grid-view').style.display = v === 'grid' ? 'grid' : 'none';
  document.getElementById('vbtn-list').classList.toggle('active', v === 'list');
  document.getElementById('vbtn-grid').classList.toggle('active', v === 'grid');
}
function toggleView(v) { setView(v); }

/* ── row interactions ── */
function openDoc(id) {
  const d = docById(id);
  if (d?.webUrl) window.open(d.webUrl, '_blank', 'noopener');
  else showToast(provider.isLive ? 'Opening…' : 'Preview is available once connected to SharePoint');
}
function toggleRow(id, el) {
  const d = docById(id);
  if (d) d.selected = el.checked;
  refreshBulkButtons();
  renderDocs();
}
function toggleAll(el) {
  const ids = new Set(getFilteredDocs().map(d => String(d.id)));
  docs.forEach(d => { if (ids.has(String(d.id))) d.selected = el.checked; });
  refreshBulkButtons();
  renderDocs();
}
function refreshBulkButtons() {
  const any = docs.some(d => d.selected);
  document.getElementById('btn-download').style.display = any ? '' : 'none';
  document.getElementById('btn-delete').style.display = any ? '' : 'none';
}

async function toggleStar(id) {
  const d = docById(id);
  if (!d) return;
  const next = !d.starred;
  d.starred = next;            // optimistic
  renderDocs();
  try { await provider.setStar(d, next); }
  catch (e) { d.starred = !next; renderDocs(); showToast('Could not update star: ' + e.message, 'error'); }
}

async function downloadDoc(id) {
  const d = docById(id);
  if (!d) return;
  if (!provider.isLive && CONFIG.mode === 'demo') { showToast('Download works once connected to SharePoint'); return; }
  try { setBusy(true); await provider.download(d); showToast('Download started', 'success'); }
  catch (e) { showToast('Download failed: ' + e.message, 'error'); }
  finally { setBusy(false); }
}

async function deleteDoc(id) {
  const d = docById(id);
  if (!d) return;
  if (!confirm(`Delete "${d.name}"? This removes it from SharePoint.`)) return;
  try {
    setBusy(true);
    await provider.remove(d);
    docs = docs.filter(x => x.id !== d.id);
    renderDocs();
    showToast('Document deleted', 'success');
  } catch (e) { showToast('Delete failed: ' + e.message, 'error'); }
  finally { setBusy(false); }
}

async function bulkAction(action) {
  const selected = docs.filter(d => d.selected);
  if (!selected.length) return;
  if (action === 'delete') {
    if (!confirm(`Delete ${selected.length} document(s) from SharePoint?`)) return;
    try {
      setBusy(true);
      for (const d of selected) await provider.remove(d);
      const ids = new Set(selected.map(d => d.id));
      docs = docs.filter(d => !ids.has(d.id));
      showToast(`${selected.length} document(s) deleted`, 'success');
    } catch (e) { showToast('Delete failed: ' + e.message, 'error'); await reload(); }
    finally { setBusy(false); }
  } else { // download
    try { setBusy(true); for (const d of selected) await provider.download(d); showToast(`Downloading ${selected.length} document(s)`, 'success'); }
    catch (e) { showToast('Download failed: ' + e.message, 'error'); }
    finally { setBusy(false); }
  }
  document.getElementById('btn-download').style.display = 'none';
  document.getElementById('btn-delete').style.display = 'none';
  renderDocs();
}

/* ── upload modal ── */
function openUpload() {
  if (CONFIG.mode === 'sharepoint' && !provider.isLive) {
    showToast('Sign in to upload to SharePoint', 'error');
    return;
  }
  const div = document.getElementById('sel-division').value;
  const dept = document.getElementById('sel-dept').value;
  if (div) { document.getElementById('up-division').value = div; onUpDivisionChange(); }
  if (dept) setTimeout(() => { document.getElementById('up-dept').value = dept; }, 10);
  document.getElementById('upload-panel').classList.add('open');
}
function closeUpload() {
  document.getElementById('upload-panel').classList.remove('open');
  pendingFiles = [];
  document.getElementById('file-list').innerHTML = '';
  document.getElementById('up-title').value = '';
  document.getElementById('btn-upload-submit').disabled = true;
}
function onUpDivisionChange() {
  const val = document.getElementById('up-division').value;
  const dept = document.getElementById('up-dept');
  dept.innerHTML = '<option value="">Select Department</option>';
  if (val && DIVISIONS[val]) DIVISIONS[val].depts.forEach(d => dept.innerHTML += `<option value="${esc(d)}">${esc(d)}</option>`);
}
function onDragOver(e) { e.preventDefault(); document.getElementById('drop-zone').classList.add('dragover'); }
function onDragLeave(e) { document.getElementById('drop-zone').classList.remove('dragover'); }
function onDrop(e) { e.preventDefault(); document.getElementById('drop-zone').classList.remove('dragover'); addFiles([...e.dataTransfer.files]); }
function onFileSelect(e) { addFiles([...e.target.files]); }
function addFiles(files) {
  files.forEach(f => { if (!pendingFiles.find(p => p.name === f.name)) pendingFiles.push(f); });
  renderFileList();
}
function renderFileList() {
  const list = document.getElementById('file-list');
  list.innerHTML = pendingFiles.map((f, i) => {
    const ic = getIcon(f.name);
    const size = f.size < 1024*1024 ? (f.size/1024).toFixed(0)+' KB' : (f.size/(1024*1024)).toFixed(1)+' MB';
    return `<div class="file-item">
      <i class="ti ${ic.icon}" aria-hidden="true"></i>
      <div class="file-item-col">
        <span class="file-item-name">${esc(f.name)}</span>
        <div class="file-item-progress" id="prog-${i}"><span></span></div>
      </div>
      <span class="file-item-size">${size}</span>
      <button class="file-item-remove" onclick="removeFile(${i})" title="Remove"><i class="ti ti-x" aria-hidden="true"></i></button>
    </div>`;
  }).join('');
  document.getElementById('btn-upload-submit').disabled = pendingFiles.length === 0;
}
function removeFile(i) { pendingFiles.splice(i, 1); renderFileList(); }

async function submitUpload() {
  const div = document.getElementById('up-division').value;
  const dept = document.getElementById('up-dept').value;
  const status = document.getElementById('up-status').value;
  const titleOverride = document.getElementById('up-title').value.trim();

  if (!div) { showToast('Select a division', 'error'); return; }
  if (!dept) { showToast('Select a department', 'error'); return; }
  if (!pendingFiles.length) { showToast('Add at least one file', 'error'); return; }

  const submitBtn = document.getElementById('btn-upload-submit');
  submitBtn.disabled = true;
  setBusy(true);

  let ok = 0;
  for (let i = 0; i < pendingFiles.length; i++) {
    const f = pendingFiles[i];
    const renamed = (i === 0 && titleOverride) ? renameWithExt(f, titleOverride) : f;
    const bar = document.querySelector(`#prog-${i} > span`);
    try {
      const newDoc = await provider.upload(renamed, { division: div, dept, status }, (p) => { if (bar) bar.style.width = Math.round(p*100) + '%'; });
      if (provider.isLive) docs.unshift(newDoc); // demo provider already inserts
      else if (!provider.isLive && CONFIG.mode === 'demo') { /* DemoProvider.upload already unshifted to its store */ }
      ok++;
    } catch (e) {
      showToast(`"${f.name}" failed: ${e.message}`, 'error');
      if (bar) { bar.style.width = '100%'; bar.style.background = 'var(--danger)'; }
    }
  }

  setBusy(false);
  if (ok) {
    showToast(`${ok} file(s) uploaded`, 'success');
    closeUpload();
    document.getElementById('sel-division').value = div;
    onDivisionChange();
    setTimeout(() => { document.getElementById('sel-dept').value = dept; onDeptChange(); }, 10);
    await reload();
  } else {
    submitBtn.disabled = false;
  }
}
function renameWithExt(file, title) {
  const ext = getExt(file.name);
  const name = /\.[^.]+$/.test(title) ? title : `${title}.${ext}`;
  try { return new File([file], name, { type: file.type }); }
  catch { file._displayName = name; return file; } // very old browsers
}

/* ── toasts ── */
function showToast(msg, type = '') {
  const icons = { success:'ti-circle-check', error:'ti-alert-circle', '':'ti-info-circle' };
  const container = document.getElementById('toasts');
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.innerHTML = `<i class="ti ${icons[type]}" aria-hidden="true"></i>${esc(msg)}`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 3600);
}

/* ── boot ── */
async function boot() {
  buildTaxonomyUI();
  setView('list');
  document.getElementById('topnav-site').textContent =
    CONFIG.mode === 'sharepoint' ? CONFIG.sharePoint.siteUrl.replace(/^https?:\/\//, '') : 'Corporate Document Management';

  provider = createDataProvider(CONFIG);
  let state;
  try { state = await provider.init(); }
  catch (e) { state = { mode: CONFIG.mode, error: provider._msg ? provider._msg(e) : e.message }; }
  renderConnection(state);

  await reload();
}

document.addEventListener('DOMContentLoaded', boot);
