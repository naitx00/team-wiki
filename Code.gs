/* Code.gs
 * OIS Team Wiki — Apps Script web app: entry point, configuration, identity.
 * Page and folder operations live in Api.gs. See README.md for deployment.
 *
 * Storage is one Google Drive folder (ROOT_FOLDER_ID in Script Properties):
 *   <root>/page.html                     a page
 *   <root>/page.html.annotations.json    its highlights and comments
 *   <root>/QA/checklist.html             a subfolder is a wiki folder
 * This is the same layout as ois-toolhub/wiki-pages, so the existing pages
 * migrate by uploading that folder to Drive as-is.
 */

// Script Properties used:
//   ROOT_FOLDER_ID  Drive folder holding the pages. `setup()` creates one.
//   ACCESS          Who can open the wiki and what each person may do —
//                   managed on the wiki's Settings page (see Access.gs).
const PROP_ROOT = 'ROOT_FOLDER_ID';
const DEFAULT_FOLDER_NAME = 'OIS Team Wiki';

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  let boot;
  try {
    boot = getBootstrap();
  } catch (err) {
    // Render the page anyway so the reader sees what is wrong, not a blank
    // Google error screen.
    boot = { error: String(err && err.message || err), user: whoAmI(), tree: null };
  }
  const brand = brand_();
  boot.brand = { name: brand.name, org: brand.org };
  // Embedded in a <script> block: escape "<" so page content can't close it.
  t.bootJson = JSON.stringify(boot).replace(/</g, '\\u003c');
  t.brand = brand;
  t.brandLogo = brandLogoHtml_(brand);
  t.brandCss = brand.css ? '<style>' + brand.css.replace(/<\//g, '') + '</style>' : '';
  return t.evaluate()
    .setTitle(brand.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // Lets the ToolHub embed the wiki in an iframe.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Branding ─────────────────────────────────────────────────
// Name, logos and colours come from BRAND in Brand.gs, which each deployment
// keeps locally (it is gitignored). Without it the wiki looks neutral.
function brand_() {
  const b = (typeof BRAND !== 'undefined' && BRAND) || {};
  const name = String(b.name || 'Team Wiki');
  return {
    name: name,
    org: String(b.org || ''),
    title: String(b.title || name),
    logoLight: /^data:image\//.test(b.logoLight || '') ? b.logoLight : '',
    logoDark: /^data:image\//.test(b.logoDark || '') ? b.logoDark : '',
    css: String(b.css || '')
  };
}

// The header logo: the brand's two images (one per theme), or a neutral mark.
function brandLogoHtml_(brand) {
  const alt = (brand.org || brand.name).replace(/[<>&"]/g, '');
  if (brand.logoLight || brand.logoDark) {
    const light = brand.logoLight || brand.logoDark, dark = brand.logoDark || brand.logoLight;
    return '<img class="wk-logo wk-logo-light" alt="' + alt + '" width="90" height="18" src="' + light + '" />' +
           '<img class="wk-logo wk-logo-dark" alt="" aria-hidden="true" width="90" height="18" src="' + dark + '" />';
  }
  return '<span class="wk-mark" aria-hidden="true"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" ' +
    'stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
    '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20V3H6.5A2.5 2.5 0 0 0 4 5.5z"/><path d="M4 19.5A2.5 2.5 0 0 0 6.5 22H20v-5"/></svg></span>';
}

// Template helper: <?!= include('Styles'); ?>
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// Everything the first paint needs, in one call — embedded into the page by
// doGet so the wiki shows its sidebar without waiting for a round trip.
// People not on the access list get their identity and nothing else.
function getBootstrap() {
  const user = whoAmI();
  if (!user.canView) return { user: user, denied: true, tree: null };
  return { user: user, tree: listTree() };
}

// ── One-time setup ─────────────────────────────────────────
// Run from the Apps Script editor. Creates the Drive folder when none is
// configured and logs its link.
function setup() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(PROP_ROOT);
  if (existing) {
    const f = DriveApp.getFolderById(existing);
    Logger.log('Already configured: "%s" %s', f.getName(), f.getUrl());
    return;
  }
  const folder = DriveApp.createFolder(DEFAULT_FOLDER_NAME);
  props.setProperty(PROP_ROOT, folder.getId());
  Logger.log('Created "%s" %s — share it with the team, then deploy.', folder.getName(), folder.getUrl());
}

// Points the wiki at an existing folder, e.g. an uploaded copy of wiki-pages.
// Run from the editor after filling in the ID from the folder's URL.
function useExistingFolder() {
  const id = 'PASTE_FOLDER_ID_HERE';
  const f = DriveApp.getFolderById(id);
  PropertiesService.getScriptProperties().setProperty(PROP_ROOT, id);
  Logger.log('Wiki now uses "%s" %s', f.getName(), f.getUrl());
}

function rootFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_ROOT);
  if (!id) throw new Error('The wiki has no Drive folder yet — run setup() in the Apps Script editor.');
  try {
    return DriveApp.getFolderById(id);
  } catch (err) {
    throw new Error("Can't open the wiki's Drive folder — ask the owner to share it with you.");
  }
}

// Identity and permissions (whoAmI, assertPerm_, Settings) live in Access.gs.

// Serialises writes across all users so a conflict check and the write that
// follows it can't interleave with someone else's save.
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The wiki is busy saving another change — try again in a moment.');
  try { return fn(); } finally { lock.releaseLock(); }
}
