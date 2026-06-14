# Document Portal — SharePoint-connected build

A Microsoft 365–styled document portal (divisions → departments, list/grid views,
upload, search, filters, stats) that reads and writes a real **SharePoint document
library** through the **Microsoft Graph API**.

It ships with two modes so it runs today and connects when you're ready:

| Mode | What it does | Needs setup? |
|------|--------------|--------------|
| `demo` (default) | Runs entirely in the browser with sample data. Nothing is saved. | No |
| `sharepoint` | Lists, uploads, downloads, deletes, and tags real files in a SharePoint library. | Yes — see below |

You switch modes by editing one line in `config.js`.

---

## Files

```
index.html         The page
styles.css         Styles (Microsoft 365 / SharePoint look)
config.js          ← the only file you edit to connect SharePoint
dataProviders.js   Data layer: MSAL sign-in + Microsoft Graph calls (and demo data)
app.js             UI logic
```

---

## Quick start (demo mode)

Because the app signs in with Microsoft, it must be **served over http(s)** — opening
`index.html` from the file system will break authentication later. For a quick local look:

```bash
cd docportal
python3 -m http.server 8080
# open http://localhost:8080
```

You'll see sample documents and can click around. Uploads/downloads are simulated in demo mode.

---

## Connecting to SharePoint

### 1. Register an app in Microsoft Entra ID (Azure AD)

1. Go to **Entra ID → App registrations → New registration**.
2. Name it (e.g. *Document Portal*).
3. Under **Redirect URI**, choose platform **Single-page application (SPA)** and enter the
   exact URL the app will be served from (e.g. `https://contoso.sharepoint.com/sites/Intranet/SiteAssets/docportal/index.html`,
   or `http://localhost:8080` for testing). The redirect URI must match `auth.redirectUri` in `config.js` character-for-character.
4. Register, then copy the **Application (client) ID** and **Directory (tenant) ID**.

### 2. Grant Microsoft Graph permissions

1. In the app registration → **API permissions → Add a permission → Microsoft Graph → Delegated permissions**.
2. Add: `User.Read`, `Sites.ReadWrite.All`, `Files.ReadWrite.All`.
3. Click **Grant admin consent** (a Global/Application admin must do this once).

> **Least privilege option:** `Sites.ReadWrite.All` grants access to every site. If your
> tenant uses **`Sites.Selected`**, an admin can instead grant this app write access to only
> the one site (via Graph `POST /sites/{id}/permissions`). The app code is unchanged; just
> swap the scope in `config.js`.

### 3. Prepare the document library

In the target SharePoint library, add these columns so the portal can store its metadata.
Column **internal** names should match `columns` in `config.js` (create them without spaces
and the internal name equals the display name):

| Column | Type | Notes |
|--------|------|-------|
| `Division` | Choice | Choices must equal the division **labels** in `config.js` (Corporate, Operations, Finance, Human Resources, IT & Technology, Legal & Compliance) |
| `Department` | Choice or Single line of text | Department names |
| `Status` | Choice | Exactly: `Published`, `Draft`, `Review` |
| `Starred` | Yes/No | Optional. Omit it and remove the line in `config.js` if you don't want starring |

### 4. Fill in `config.js`

```js
window.CONFIG = {
  mode: 'sharepoint',                          // ← switch from 'demo'
  auth: {
    clientId: '<Application (client) ID>',
    tenantId: '<Directory (tenant) ID>',
    redirectUri: window.location.origin + window.location.pathname,  // or a fixed URL
    scopes: ['User.Read', 'Sites.ReadWrite.All', 'Files.ReadWrite.All'],
  },
  sharePoint: {
    siteUrl: 'https://contoso.sharepoint.com/sites/Intranet',
    libraryName: 'Documents',
  },
  // columns / divisions: adjust to match your library
};
```

### 5. Host the app

The files are static — host them anywhere that serves over HTTPS at the registered redirect URI:

- **Inside SharePoint** (no extra infra): upload the four files to a folder in **Site Assets**,
  then link to `index.html`. Quick, but SharePoint may apply its own headers; test sign-in.
- **Azure Static Web Apps / Azure Storage static site / any static host:** drop the folder in,
  point the redirect URI at the resulting URL.

That's it. Open the page, click **Sign in**, and the portal will load live documents.

---

## How the mapping works

| Portal field | SharePoint source |
|--------------|-------------------|
| File name, size, modified date, uploaded-by | The file (Graph `driveItem`) |
| Division | `Division` column (label ↔ key translated automatically) |
| Department | `Department` column |
| Status badge | `Status` column |
| Star | `Starred` column |
| Open | The file's `webUrl` (opens in SharePoint / Office online) |

Uploads under 4 MB use a direct `PUT`; larger files use a resumable Graph **upload session**
and are sent in 3.75 MB chunks. After upload, the metadata columns are set on the new item.

---

## Troubleshooting

| Message in the connection badge | Cause / fix |
|---------------------------------|-------------|
| *Falling back to demo* (console) | `config.js` still has placeholder IDs or `YOURTENANT`. Fill them in. |
| **Site or library not found** | Check `siteUrl` and `libraryName` (use the library's display name). |
| **Access denied / consent pending** | Admin hasn't granted consent, or the account lacks site permission. |
| **A metadata column is missing** | A column in `config.js` doesn't exist in the library, or the internal name differs. |
| Sign-in popup closes / blocked | Allow popups; otherwise the app falls back to a full-page redirect automatically. |
| Redirect loop / `AADSTS` error | `auth.redirectUri` must exactly equal a SPA redirect URI on the app registration. |

---

## Alternative: SharePoint Framework (SPFx) web part

This build is a **standalone SPA** — the simplest way to point an HTML app at SharePoint, and
it can live outside SharePoint too. If your goal is to embed the portal *as a web part on a
SharePoint page* with sign-in handled automatically by the page context, the enterprise-standard
approach is an **SPFx web part** (built with the Yeoman SharePoint generator, packaged as a
`.sppkg`, deployed to the tenant App Catalog). It removes the separate app registration and
gets Graph tokens from the page, but requires a Node build toolchain and tenant deployment. The
data-mapping design here (library columns ↔ portal fields) carries over directly if you migrate.

---

## Security notes

- No secrets live in this code. Authentication uses MSAL with the **authorization-code +
  PKCE** flow for SPAs; tokens are held in `sessionStorage` and cleared on sign-out.
- The app only ever has the permissions you consent to, scoped to the signed-in user's own
  access (delegated). It cannot see anything the user couldn't already see in SharePoint.
- Prefer `Sites.Selected` if your security team wants to limit the app to a single site.
