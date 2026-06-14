/* =====================================================================
   config.js  —  EDIT THIS FILE to connect the portal to SharePoint.
   This is the only file you need to change for a basic deployment.
   See README.md for step-by-step setup.

   ---------------------------------------------------------------------
   SETUP CHECKLIST  (to go live against SharePoint)
     [✅] siteUrl filled in        → intlexp.sharepoint.com/sites/Corporate
     [✅] libraryName              → 'Documents' (default library)
     [✅] clientId                 → app registration
     [✅] tenantId                 → app registration
     [✅] add 4 library columns    → Division, Department, Status, Starred
     [ ] confirm SPA redirect URI  → must match where you host the page
                                     (e.g. http://localhost:8080/ for local test)
     [✅] flip mode to 'sharepoint' → LIVE
   ---------------------------------------------------------------------
   ===================================================================== */

window.CONFIG = {

  /* ---------------------------------------------------------------
     MODE
     'demo'        → runs entirely in the browser with sample data.
                     No sign-in, nothing is saved. Good for trying the UI.
     'sharepoint'  → reads/writes a real SharePoint document library
                     through Microsoft Graph. Requires the settings below.
     --------------------------------------------------------------- */
  mode: 'sharepoint',      // ✅ LIVE — reading/writing the Corporate Documents library

  /* ---------------------------------------------------------------
     AZURE / ENTRA APP REGISTRATION  (required for 'sharepoint' mode)
     Create an app registration in the Azure portal (Entra ID).
     See README → "1. Register the app".
     --------------------------------------------------------------- */
  auth: {
    clientId: '911c89fb-4db5-4523-9d4a-2f32e5e21d31',   // ✅ DONE — Application (client) ID
    tenantId: '0fbc929f-2e2c-4db7-9e8c-e12707a9540b',   // ✅ DONE — Directory (tenant) ID
    // Where Microsoft sends the user back after sign-in. Must EXACTLY match a
    // Redirect URI (type: SPA) registered on the app. Usually the page's own URL.
    redirectUri: window.location.origin + window.location.pathname,
    // Delegated Microsoft Graph permissions the app asks the user to consent to.
    scopes: ['User.Read', 'Sites.ReadWrite.All', 'Files.ReadWrite.All'],
  },

  /* ---------------------------------------------------------------
     SHAREPOINT TARGET  (required for 'sharepoint' mode)
     --------------------------------------------------------------- */
  sharePoint: {
    // Full URL of the SharePoint site that holds the library.
    siteUrl: 'https://intlexp.sharepoint.com/sites/Corporate',   // ✅ DONE
    // Display name of the document library. 'Documents' is the default
    // library on every site (its URL path is /Shared Documents).
    libraryName: 'Documents',                                    // ✅ leave as-is unless you target a different library
  },

  /* ---------------------------------------------------------------
     COLUMN MAPPING
     Internal names of the library columns that hold the portal's metadata.
     For columns created without spaces, the internal name == the display name.
     If a column already exists with a different internal name, set it here.
     --------------------------------------------------------------- */
  columns: {
    division:   'Division',     // Choice column (values must match the labels below)
    department: 'Department',   // Choice or single-line text column
    status:     'Status',       // Choice column: Published / Draft / Review
    starred:    'Starred',      // Yes/No (boolean) column  — optional
  },

  /* When a document has no Status set in SharePoint, treat it as: */
  defaultStatus: 'Published',

  /* ---------------------------------------------------------------
     TAXONOMY  (drives the Division / Department selectors)
     Keep these labels in sync with the choices in your SharePoint
     Division and Department columns.
     --------------------------------------------------------------- */
  divisions: {
    corporate:  { label: 'Corporate',           icon: 'ti-building',       depts: ['Executive Office','Communications','Strategy','Investor Relations'] },
    operations: { label: 'Operations',          icon: 'ti-settings-cog',   depts: ['Supply Chain','Logistics','Quality Assurance','Facilities','Manufacturing'] },
    finance:    { label: 'Finance',             icon: 'ti-chart-bar',      depts: ['Accounting','Treasury','Tax','Financial Planning','Audit'] },
    hr:         { label: 'Human Resources',     icon: 'ti-users',          depts: ['Talent Acquisition','Learning & Development','Compensation & Benefits','Employee Relations','Diversity & Inclusion'] },
    it:         { label: 'IT & Technology',     icon: 'ti-device-laptop',  depts: ['Infrastructure','Software Development','Cybersecurity','Data & Analytics','IT Support'] },
    legal:      { label: 'Legal & Compliance',  icon: 'ti-scale',          depts: ['Corporate Legal','Contract Management','Regulatory Compliance','Data Privacy','Risk Management'] },
  },
};
