/* Access.gs
 * OIS Team Wiki — who can open the wiki, what each person may do, and the
 * Settings page admins use to change that.
 *
 * Stored as JSON in Script Property ACCESS:
 *   {
 *     admins:  ['admin@example.com'],   // manage Settings; can do everything
 *     viewers: 'listed' | 'domain',                   // who can open the wiki at all
 *     users:   { 'name@example.com': { add, edit, upload, link, delete } }
 *   }
 * The permissions:
 *   add     create pages and folders
 *   edit    change pages (and Docs inline), move pages, rename folders,
 *           highlights and comments on wiki pages
 *   upload  upload files (converted to Docs, Sheets, Slides)
 *   link    link existing Google or Office files
 *   delete  delete pages and folders, remove links
 * Everyone on the list can open and read the wiki; `viewers: 'domain'` also
 * lets anyone else at the admins' domain read it.
 *
 * There is always at least one admin, and admins can't remove their own
 * admin rights, so a save can't lock everyone out of Settings. (Should it
 * happen anyway, delete the ACCESS property in Project Settings → Script
 * Properties; the next run starts from the seed below.) The old EDITORS
 * property is ignored once ACCESS exists.
 */

const PROP_ACCESS = 'ACCESS';
const PERMS = ['add', 'edit', 'upload', 'link', 'delete'];

// The starting access list (used once, when ACCESS doesn't exist yet) lives
// in AccessSeed.gs, which is kept out of the repository — it names real
// people. See AccessSeed.gs.example. Without it, run setupAccess() once.
function seedAdmins_() { return typeof SEED_ADMINS !== 'undefined' ? SEED_ADMINS.slice() : []; }
function seedUsers_() { return typeof SEED_USERS !== 'undefined' ? SEED_USERS.slice() : []; }
const EMAIL_RE = /^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/;

function allPerms_(on) {
  const p = {};
  PERMS.forEach(k => { p[k] = !!on; });
  return p;
}

// The organisation "anyone at … can view" refers to: the admins' domain.
function orgDomain_(access) {
  const a = (access.admins[0] || seedAdmins_()[0] || '');
  return a.split('@')[1] || '';
}

function readAccess_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_ACCESS);
  if (raw) {
    try { return normaliseAccess_(JSON.parse(raw)); } catch (e) { /* fall through to the seed */ }
  }
  const seeded = { admins: seedAdmins_(), viewers: 'listed', users: {} };
  seedUsers_().forEach(e => { seeded.users[e] = allPerms_(true); });
  // With no seed there is nobody to let in yet — don't store an empty list;
  // setupAccess() creates the first admin.
  if (seeded.admins.length) props.setProperty(PROP_ACCESS, JSON.stringify(seeded));
  return seeded;
}

// Run once from the Apps Script editor on a fresh install without
// AccessSeed.gs: makes the person running it the first admin. Does nothing
// once an access list exists.
function setupAccess() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_ACCESS)) {
    Logger.log('An access list already exists — manage people on the wiki\'s Settings page.');
    return;
  }
  const me = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!me) throw new Error('Could not tell who you are — run this from the Apps Script editor.');
  const access = { admins: [me], viewers: 'listed', users: {} };
  access.users[me] = allPerms_(true);
  props.setProperty(PROP_ACCESS, JSON.stringify(access));
  Logger.log('%s is now the wiki\'s admin. Add people on the Settings page.', me);
}

// Lower-case, valid emails only; unknown keys dropped; every permission a boolean.
function normaliseAccess_(a) {
  const clean = e => String(e || '').trim().toLowerCase();
  const out = { admins: [], viewers: a && a.viewers === 'domain' ? 'domain' : 'listed', users: {} };
  ((a && a.admins) || []).map(clean).forEach(e => {
    if (EMAIL_RE.test(e) && out.admins.indexOf(e) === -1) out.admins.push(e);
  });
  const users = (a && a.users) || {};
  Object.keys(users).forEach(k => {
    const e = clean(k);
    if (!EMAIL_RE.test(e)) return;
    const perms = {};
    PERMS.forEach(p => { perms[p] = !!(users[k] && users[k][p]); });
    out.users[e] = perms;
  });
  return out;
}

// Everything the client needs to know about the person using it.
function whoAmI() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const access = readAccess_();
  const isAdmin = !!email && access.admins.indexOf(email) !== -1;
  const listed = access.users[email];
  const perms = isAdmin ? allPerms_(true) : Object.assign(allPerms_(false), listed || {});
  const domain = orgDomain_(access);
  const canView = !!email && (isAdmin || !!listed ||
    (access.viewers === 'domain' && !!domain && email.endsWith('@' + domain)));
  return {
    email: email,
    isAdmin: isAdmin,
    canView: canView,
    perms: perms,
    canEdit: PERMS.some(p => perms[p]),          // any permission at all
    contact: access.admins.slice(0, 3)           // who to ask for access
  };
}

// Reads call this; so does doGet.
function assertCanView_() {
  const me = whoAmI();
  if (!me.canView) {
    throw new Error(me.email
      ? me.email + " doesn't have access to this wiki yet — " + (me.contact[0]
          ? 'ask ' + me.contact[0] + ' to add you.'
          : 'it has no admin yet: the owner should run setupAccess() in the Apps Script editor.')
      : 'Sign in with your ' + (brand_().org ? brand_().org + ' ' : 'work ') + 'Google account to open the wiki.');
  }
  return me;
}

// Every write calls this with the permission it needs. The client hides the
// controls too, but this is the check that counts.
const PERM_WORDS = {
  add: 'create pages or folders', edit: 'edit pages', upload: 'upload files',
  link: 'link files', delete: 'delete pages or folders'
};
function assertPerm_(perm) {
  const me = assertCanView_();
  if (!me.perms[perm]) {
    throw new Error(me.email + " doesn't have permission to " + PERM_WORDS[perm] + ' — ask ' +
      (me.contact[0] || 'an admin') + '.');
  }
  return me;
}

// Kept for anything that only needs "may change something".
function assertCanEdit_() {
  const me = assertCanView_();
  if (!me.canEdit) throw new Error(me.email + ' has read-only access to this wiki.');
  return me;
}

// ── Settings (admins only) ──────────────────────────────────

function assertAdmin_() {
  const me = assertCanView_();
  if (!me.isAdmin) throw new Error('Only wiki admins can change settings.');
  return me;
}

function getSettings() {
  const me = assertAdmin_();
  const access = readAccess_();
  // Admins appear in the list too, so they can be seen and managed in one place.
  const emails = Object.keys(access.users);
  access.admins.forEach(a => { if (emails.indexOf(a) === -1) emails.push(a); });
  emails.sort();
  return {
    me: me.email,
    domain: orgDomain_(access),
    viewers: access.viewers,
    perms: PERMS.slice(),
    people: emails.map(e => ({
      email: e,
      admin: access.admins.indexOf(e) !== -1,
      perms: Object.assign(allPerms_(false), access.users[e] || {})
    }))
  };
}

// `settings` is { viewers, people: [{ email, admin, perms }] } as the page
// sends it. Returns the saved settings.
function saveSettings(settings) {
  const me = assertAdmin_();
  const next = { admins: [], viewers: settings && settings.viewers === 'domain' ? 'domain' : 'listed', users: {} };
  const bad = [];
  ((settings && settings.people) || []).forEach(p => {
    const e = String(p && p.email || '').trim().toLowerCase();
    if (!e) return;
    if (!EMAIL_RE.test(e)) { bad.push(p.email); return; }
    const perms = {};
    PERMS.forEach(k => { perms[k] = !!(p.perms && p.perms[k]); });
    next.users[e] = perms;
    if (p.admin && next.admins.indexOf(e) === -1) next.admins.push(e);
  });
  if (bad.length) throw new Error('Not valid email addresses: ' + bad.join(', '));
  // Someone must always be able to reach Settings.
  if (next.admins.indexOf(me.email) === -1) {
    throw new Error("You can't remove your own admin access — ask another admin to do it.");
  }
  withLock_(() => {
    PropertiesService.getScriptProperties().setProperty(PROP_ACCESS, JSON.stringify(normaliseAccess_(next)));
  });
  return getSettings();
}
