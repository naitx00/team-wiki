/* Api.gs
 * OIS Team Wiki — every read and write the page makes, via google.script.run.
 *
 * Pages are addressed by Drive file ID; folders by their path under the root
 * ("QA/Templates"), the same folder paths the client already works in.
 *
 * Each page's title, snippet and revision number are kept in its Drive
 * description as a small JSON record, so listing the wiki never has to open
 * every page. The revision is what catches two people saving the same page.
 */

const PAGE_RE = /\.(html?|md|markdown)$/i;
const SIDECAR = '.annotations.json';
const META_TAG = 'ois-wiki';
const SNIPPET_LEN = 160;
const MAX_PAGE_CHARS = 9500000;   // File.setContent refuses anything over 10 MB
const TEXT_CACHE_SECS = 21600;    // CacheService maximum (6 hours)

// ── Reading ─────────────────────────────────────────────────

// Folders and page metadata only — never page content. One imported Word
// document can be over a megabyte, so content comes one page at a time
// through getPage().
function listTree() {
  assertCanView_();
  const root = rootFolder_();
  const out = { space: root.getName(), folders: [], pages: [] };
  walk_(root, '', out);
  return out;
}

function walk_(folder, path, out) {
  const subs = folder.searchFolders('trashed = false');
  while (subs.hasNext()) {
    const f = subs.next();
    const name = f.getName();
    if (name.charAt(0) === '.') continue;
    const child = path ? path + '/' + name : name;
    out.folders.push(child);
    walk_(f, child, out);
  }
  const files = folder.searchFiles('trashed = false');
  while (files.hasNext()) {
    const file = files.next();
    const kind = pageKind_(file);
    if (kind) out.pages.push(pageInfo_(file, path, null, kind));
  }
}

function getPage(id) {
  assertCanView_();
  const file = pageFile_(id);
  const kind = pageKind_(file);
  const meta = readMeta_(file);
  if (kind !== 'html' && (appOfPage_(file, kind) !== 'doc' || officeOfPage_(file, kind))) {
    return { id: id, kind: kind, content: '', snippet: meta ? meta.s : '', rev: 0,
             modified: docModified_(file, kind), annotations: { highlights: [], comments: [] } };
  }
  if (kind !== 'html') {
    const docId = docIdOf_(file, kind);
    const content = docHtml_(docId);
    // Keep the snippet the sidebar and contents list show in step with the
    // Doc, which is edited outside the wiki.
    const snippet = truncate_(plainText_(content), SNIPPET_LEN);
    if (!meta || meta.s !== snippet) {
      try { writeMeta_(file, { t: meta ? meta.t : '', s: snippet, r: 0 }); } catch (e) { /* read-only */ }
    }
    return {
      id: id, kind: kind, content: content, snippet: snippet, rev: 0,
      modified: docModified_(file, kind),
      annotations: { highlights: [], comments: [] }
    };
  }
  return {
    id: id,
    kind: kind,
    content: readText_(file),
    rev: meta ? meta.r : 0,
    modified: file.getLastUpdated().getTime(),
    annotations: readAnnotations_(file)
  };
}

// Searches every page's text. A page matches, best first, when it has:
//   3  the exact phrase typed ("issues/suggestions/updates")
//   2  the same words in the same order, whatever punctuation or spacing sits
//      between them ("Issues / Suggestions / Updates", a line break, commas)
//   1  every word somewhere on the page
// Each page's searchable text is cached per revision, so only pages changed
// since the last search are read from Drive; Docs that aren't cached are
// exported in parallel. A page that can't be read is reported back in
// `unreadable` (so the reader isn't told "no matches" in silence) and is
// never cached as empty — the next search tries it again.
function searchPages(query) {
  assertCanView_();
  const q = normalizeText_(String(query || '')).replace(/\s+/g, ' ').trim().toLowerCase();
  const empty = { hits: [], searched: 0, unreadable: [] };
  if (!q) return empty;
  const qLoose = looseText_(q);
  const words = qLoose ? qLoose.split(' ') : [];
  const pages = listTree().pages;
  const cache = CacheService.getScriptCache();
  const keyOf = p => 'txt:' + p.id + ':' + p.rev + ':' + p.modified;
  const texts = readTextCache_(cache, pages.map(keyOf));
  const problems = {};

  const missing = pages.filter(p => texts[keyOf(p)] === undefined);
  const remember = (p, text) => { texts[keyOf(p)] = text; writeTextCache_(cache, keyOf(p), text); };

  missing.filter(p => p.kind === 'html').forEach(p => {
    try { remember(p, searchText_(readText_(DriveApp.getFileById(p.id)))); }
    catch (e) { problems[p.id] = 'could not be read'; }
  });

  // Docs: export as HTML, many at once. What fails is retried once if Google
  // was only busy, and otherwise read through DocumentApp instead — that
  // route has no export size limit, which long-running Docs (a weekly agenda
  // gaining a page every week) eventually outgrow.
  missing.filter(p => p.kind !== 'html' && p.office).forEach(p => {
    try { remember(p, normalizeText_(officePlainText_(p.docId, p.office)).replace(/\s+/g, ' ').trim()); }
    catch (e) { problems[p.id] = String(e && e.message || e).replace(/^Exception:\s*/, ''); }
  });

  missing.filter(p => p.kind !== 'html' && !p.office && p.app === 'sheet').forEach(p => {
    try { remember(p, normalizeText_(sheetPlainText_(p.docId)).replace(/\s+/g, ' ').trim()); }
    catch (e) {
      // Drive can at least export the first tab.
      try {
        const req = exportRequest_(p.docId, 'text/csv');
        const res = UrlFetchApp.fetch(req.url, req);
        if (res.getResponseCode() === 200) remember(p, searchText_(res.getContentText('UTF-8')));
        else problems[p.id] = exportFailure_(res).reason;
      } catch (e2) { problems[p.id] = docAppProblem_(e); }
    }
  });

  const exportMime = p => p.app === 'slide' ? 'text/plain' : 'text/html';
  const docs = missing.filter(p => p.kind !== 'html' && !p.office && p.app !== 'sheet');
  let retry = [];
  const fallback = [];
  const lastFailure = {};
  for (let i = 0; i < docs.length; i += 15) {
    const batch = docs.slice(i, i + 15);
    let responses;
    try { responses = UrlFetchApp.fetchAll(batch.map(p => exportRequest_(p.docId, exportMime(p)))); }
    catch (e) { retry = retry.concat(batch); continue; }
    responses.forEach((res, j) => {
      const p = batch[j];
      if (res.getResponseCode() === 200) { remember(p, searchText_(res.getContentText('UTF-8'))); return; }
      const failure = exportFailure_(res);
      lastFailure[p.id] = failure;
      (failure.retryable ? retry : fallback).push(p);
    });
  }
  if (retry.length) Utilities.sleep(700);   // rate limits are usually gone a moment later
  retry.forEach(p => {
    try {
      const req = exportRequest_(p.docId, exportMime(p));
      const res = UrlFetchApp.fetch(req.url, req);
      if (res.getResponseCode() === 200) { remember(p, searchText_(res.getContentText('UTF-8'))); return; }
      lastFailure[p.id] = exportFailure_(res);
    } catch (e) { lastFailure[p.id] = { reason: String(e && e.message || e) }; }
    fallback.push(p);
  });
  fallback.forEach(p => {
    if (p.app !== 'doc') { problems[p.id] = (lastFailure[p.id] && lastFailure[p.id].reason) || 'could not be read'; return; }
    try { remember(p, normalizeText_(docPlainText_(p.docId)).replace(/\s+/g, ' ').trim()); }
    catch (e) {
      problems[p.id] = (lastFailure[p.id] && lastFailure[p.id].reason) || docAppProblem_(e);
    }
  });

  const hits = [];
  pages.forEach(p => {
    const text = texts[keyOf(p)] || '';
    const score = Math.max(matchScore_(normalizeText_(p.title).toLowerCase(), q, qLoose, words),
                           matchScore_(text.toLowerCase(), q, qLoose, words));
    if (!score) return;
    hits.push({ id: p.id, score: score, snippet: snippetFor_(text, q, words) });
  });
  hits.sort((a, b) => b.score - a.score);
  return {
    hits: hits,
    searched: pages.length,
    unreadable: pages.filter(p => problems[p.id]).map(p => ({ id: p.id, title: p.title, reason: problems[p.id] }))
  };
}

// Punctuation and spacing folded to single spaces: "Issues / Suggestions,
// Updates" and "issues/suggestions/updates" both become
// "issues suggestions updates".
function looseText_(s) {
  return String(s).toLowerCase().replace(/[^\p{L}\p{N}]+/gu, ' ').trim();
}

function matchScore_(lower, q, qLoose, words) {
  if (!lower) return 0;
  if (lower.indexOf(q) !== -1) return 3;
  if (!qLoose) return 0;
  const loose = ' ' + looseText_(lower) + ' ';
  if (loose.indexOf(' ' + qLoose + ' ') !== -1 || loose.indexOf(qLoose) !== -1) return 2;
  return words.every(w => loose.indexOf(w) !== -1) ? 1 : 0;
}

// Text around the best place in the page: the exact phrase, else its first
// word.
function snippetFor_(text, q, words) {
  const lower = text.toLowerCase();
  let at = lower.indexOf(q);
  let len = q.length;
  if (at === -1 && words.length) {
    at = lower.indexOf(words[0]);
    len = words[0].length;
  }
  return at === -1 ? truncate_(text, 140) : around_(text, at, len);
}

// Why Google refused an export, in words, and whether trying again soon could
// help. Google's own reason code tells apart the cases that all arrive as 403.
function exportFailure_(res) {
  const code = res.getResponseCode();
  let reason = '';
  let message = '';
  try {
    const err = JSON.parse(res.getContentText()).error || {};
    message = err.message || '';
    reason = (err.errors && err.errors[0] && err.errors[0].reason) || '';
  } catch (e) { /* not JSON */ }
  if (code === 429 || code >= 500 || /rateLimit|userRateLimit|backendError/i.test(reason)) {
    return { retryable: true, reason: 'Google was busy — search again in a moment' };
  }
  if (/exportSizeLimit/i.test(reason) || /too large/i.test(message)) {
    return { reason: 'too large for Google to export' };
  }
  if (/cannotDownload|cannotCopy|cannotExport/i.test(reason) || /download|copy/i.test(message)) {
    return { reason: "its owner turned off downloading and copying for viewers — ask them to allow it, or to make you an editor" };
  }
  if (code === 404 || /notFound/i.test(reason)) {
    return { reason: "not shared with the wiki's owner, or deleted" };
  }
  if (code === 403) return { reason: "the wiki's owner doesn't have permission to read it" + (message ? ' (' + message + ')' : '') };
  return { reason: 'Google Drive error ' + code + (message ? ': ' + message : '') };
}

function docAppProblem_(e) {
  const msg = String(e && e.message || e);
  if (/not have permission|access denied|not found|missing/i.test(msg)) return "not shared with the wiki's owner, or deleted";
  return msg;
}

// ── Writing pages ───────────────────────────────────────────

// `baseRev` is the revision the editor started from. If someone else has
// saved since, nothing is written and the caller gets { conflict: true } so
// the reader can decide; `force` overwrites anyway.
function savePage(id, content, baseRev, force) {
  assertPerm_('edit');
  checkSize_(content);
  return withLock_(() => guard_(() => {
    const file = pageFile_(id);
    if (pageKind_(file) !== 'html') throw new Error('This page is a Google Doc — edit it in Google Docs.');
    const current = readMeta_(file);
    const rev = current ? current.r : 0;
    if (!force && baseRev !== null && baseRev !== undefined && rev !== baseRev) {
      return { conflict: true, rev: rev, modified: file.getLastUpdated().getTime() };
    }
    file.setContent(content);
    const meta = computeMeta_(content, file.getName(), rev + 1);
    writeMeta_(file, meta);
    const info = pageInfo_(file, folderPathOf_(file), meta);
    writeTextCache_(CacheService.getScriptCache(), 'txt:' + id + ':' + info.rev + ':' + info.modified, searchText_(content));
    info.conflict = false;
    return info;
  }));
}

// Creates a page in `folderPath` (created if missing). The filename is made
// unique within that folder; the one actually used comes back in the result.
function createPage(folderPath, filename, content) {
  assertPerm_('add');
  checkSize_(content);
  return withLock_(() => guard_(() => {
    const folder = resolveFolder_(folderPath, true);
    const name = uniqueName_(folder, cleanFileName_(filename));
    const mime = /\.html?$/i.test(name) ? MimeType.HTML : MimeType.PLAIN_TEXT;
    const file = folder.createFile(Utilities.newBlob(content, mime, name));
    const meta = computeMeta_(content, name, 1);
    writeMeta_(file, meta);
    return pageInfo_(file, normPath_(folderPath), meta);
  }));
}

// Moves to Drive's trash rather than deleting outright, so a mistake can be
// undone from Drive for 30 days. For a linked Doc only the shortcut goes —
// the Doc itself belongs to someone else and is never touched.
function deletePage(id) {
  assertPerm_('delete');
  return withLock_(() => guard_(() => {
    const file = pageFile_(id);
    if (pageKind_(file) === 'html') {
      const side = sidecarOf_(file);
      if (side) side.setTrashed(true);
    }
    file.setTrashed(true);
    return true;
  }));
}

function movePage(id, destPath) {
  assertPerm_('edit');
  return withLock_(() => guard_(() => {
    const file = pageFile_(id);
    const kind = pageKind_(file);
    const dest = resolveFolder_(destPath, true);
    // Docs are named by title and Drive allows duplicates, so only file
    // pages need a unique name (and carry a sidecar along).
    const side = kind === 'html' ? sidecarOf_(file) : null;
    const name = kind === 'html' ? uniqueName_(dest, file.getName()) : file.getName();
    if (name !== file.getName()) file.setName(name);
    file.moveTo(dest);
    if (side) {
      side.setName(name + SIDECAR);
      side.moveTo(dest);
    }
    return pageInfo_(file, normPath_(destPath), null, kind);
  }));
}

function saveAnnotations(id, annotations) {
  assertPerm_('edit');
  return withLock_(() => guard_(() => {
    const file = pageFile_(id);
    if (pageKind_(file) !== 'html') throw new Error('Use Google Docs comments on this page.');
    const clean = {
      highlights: Array.isArray(annotations && annotations.highlights) ? annotations.highlights : [],
      comments: Array.isArray(annotations && annotations.comments) ? annotations.comments : []
    };
    const side = sidecarOf_(file);
    const empty = !clean.highlights.length && !clean.comments.length;
    if (empty) {
      if (side) side.setTrashed(true);
      return true;
    }
    const json = JSON.stringify(clean, null, 2);
    if (side) side.setContent(json);
    else parentOf_(file).createFile(Utilities.newBlob(json, 'application/json', file.getName() + SIDECAR));
    return true;
  }));
}

// ── Writing folders ─────────────────────────────────────────

function createFolder(path) {
  assertPerm_('add');
  return withLock_(() => guard_(() => {
    resolveFolder_(path, true);
    return normPath_(path);
  }));
}

function renameFolder(path, newName) {
  assertPerm_('edit');
  const clean = cleanFolderName_(newName);
  if (!normPath_(path)) throw new Error("The top level can't be renamed.");
  if (!clean) throw new Error("That name can't be used for a folder.");
  return withLock_(() => guard_(() => {
    const parts = normPath_(path).split('/');
    const parent = resolveFolder_(parts.slice(0, -1).join('/'), false);
    if (childFolder_(parent, clean)) throw new Error('"' + clean + '" already exists.');
    resolveFolder_(path, false).setName(clean);
    return parts.slice(0, -1).concat(clean).join('/');
  }));
}

function deleteFolder(path) {
  assertPerm_('delete');
  if (!normPath_(path)) throw new Error("The top level can't be deleted.");
  return withLock_(() => guard_(() => {
    resolveFolder_(path, false).setTrashed(true);
    return true;
  }));
}

// ── Page metadata ───────────────────────────────────────────

function pageInfo_(file, folderPath, knownMeta, knownKind) {
  const kind = knownKind || pageKind_(file);
  if (kind !== 'html') return docPageInfo_(file, folderPath, kind);
  const name = file.getName();
  let meta = knownMeta || readMeta_(file);
  if (!meta) {
    // A page added straight to Drive, or migrated from wiki-pages. Work the
    // metadata out once and store it, so the next listing is cheap.
    meta = computeMeta_(readText_(file), name, 0);
    if (!file.getDescription()) {
      try { writeMeta_(file, meta); } catch (e) { /* read-only visitor */ }
    }
  }
  const folder = folderPath || '';
  return {
    id: file.getId(),
    kind: 'html',
    file: name,
    folder: folder,
    path: folder ? folder + '/' + name : name,
    title: meta.t,
    snippet: meta.s,
    rev: meta.r,
    modified: file.getLastUpdated().getTime()
  };
}

// A Google Doc or a shortcut to one. The title is always the Doc's current
// name; the snippet is whatever getPage last saw (empty until first opened,
// since working it out means exporting the Doc).
function docPageInfo_(file, folderPath, kind) {
  const folder = folderPath || '';
  const docId = docIdOf_(file, kind);
  let title = file.getName();
  let modified = file.getLastUpdated().getTime();
  let broken = false;
  if (kind === 'link') {
    try {
      const target = DriveApp.getFileById(docId);
      title = target.getName();
      modified = target.getLastUpdated().getTime();
    } catch (e) { broken = true; }
  }
  const meta = readMeta_(file);
  const app = appOfPage_(file, kind);
  const office = officeOfPage_(file, kind);
  if (office) title = cleanTitle_(title);
  return {
    id: file.getId(),
    kind: kind,
    app: app,
    office: office || null,
    docId: docId,
    docUrl: fileUrl_(app, docId, office),
    broken: broken,
    file: file.getName(),
    folder: folder,
    path: (folder ? folder + '/' : '') + file.getName() + '#' + file.getId(),
    title: title,
    snippet: meta ? meta.s : '',
    rev: 0,
    modified: modified
  };
}

// 'html' for .html/.md page files, 'gdoc' for a Google Doc, 'link' for a
// shortcut to a Google Doc; null for anything else in the folder.
function pageKind_(file) {
  const mime = file.getMimeType();
  if (googleAppOf_(mime)) return 'gdoc';
  if (mime === SHORTCUT_MIME) {
    try { return googleAppOf_(file.getTargetMimeType()) ? 'link' : null; } catch (e) { return null; }
  }
  const name = file.getName();
  return name.charAt(0) !== '.' && PAGE_RE.test(name) ? 'html' : null;
}

// 'pptx', 'docx', … for an Office file, '' for a Google-native one.
function officeOfPage_(file, kind) {
  let mime = file.getMimeType();
  if (kind === 'link') { try { mime = file.getTargetMimeType(); } catch (e) { return ''; } }
  const o = OFFICE_APPS[mime];
  return o ? o.office : '';
}

function appOfPage_(file, kind) {
  let mime = file.getMimeType();
  if (kind === 'link') { try { mime = file.getTargetMimeType(); } catch (e) { /* unknown */ } }
  const g = googleAppOf_(mime);
  return g ? g.app : 'doc';
}

function docIdOf_(file, kind) {
  return kind === 'link' ? file.getTargetId() : file.getId();
}

function docModified_(file, kind) {
  if (kind === 'link') {
    try { return DriveApp.getFileById(file.getTargetId()).getLastUpdated().getTime(); } catch (e) { /* fall through */ }
  }
  return DriveApp.getFileById(file.getId()).getLastUpdated().getTime();
}

function readMeta_(file) {
  const d = file.getDescription();
  if (!d || d.indexOf(META_TAG) === -1) return null;
  try {
    const m = JSON.parse(d);
    if (m[META_TAG] !== 1) return null;
    return { t: String(m.t || ''), s: String(m.s || ''), r: Number(m.r) || 0 };
  } catch (e) { return null; }
}

function writeMeta_(file, meta) {
  const record = {};
  record[META_TAG] = 1;
  record.t = meta.t;
  record.s = meta.s;
  record.r = meta.r;
  file.setDescription(JSON.stringify(record));
}

function computeMeta_(content, filename, rev) {
  return {
    t: titleFromContent_(content, titleFromFilename_(filename)),
    s: truncate_(plainText_(content), SNIPPET_LEN),
    r: rev
  };
}

// The same rules as the client: a page's title is its first heading, as
// Markdown or as an <h1>.
function titleFromContent_(content, fallback) {
  const text = String(content || '');
  const md = text.match(/^\s*#\s+(.+)$/m);
  if (md) return cleanTitle_(md[1].trim());
  const h1 = text.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  if (h1) {
    const plain = decodeEntities_(h1[1].replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim();
    if (plain) return cleanTitle_(plain);
  }
  return fallback;
}

function titleFromFilename_(name) {
  return String(name || '').replace(/\.[A-Za-z0-9]{1,8}$/, '').replace(/[-_]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function cleanTitle_(t) {
  return String(t || '').replace(/\.(html?|md|markdown|docx?|pdf|txt|rtf|xlsx?|pptx?)$/i, '').trim();
}

// Readable text without the title or any markup — for snippets.
function plainText_(content) {
  return markupToText_(String(content || '')
    .replace(/^\s*#\s+.*\n+/, '')
    .replace(/<h1[^>]*>[\s\S]*?<\/h1>/i, ''));
}

// Everything a reader could search for, headings included.
function searchText_(content) {
  return markupToText_(String(content || '').replace(/<head[\s\S]*?<\/head>/i, ' '));
}

// Tags that only change formatting are removed outright. Google's export
// wraps every run of differently formatted text in its own <span>, even in
// the middle of a word, so turning those into spaces split "on-call" into
// "on- call" and searches missed it. Block tags still become spaces.
const INLINE_TAGS_RE = /<\/?(span|a|b|i|u|em|strong|sup|sub|font|s|strike|del|ins|mark|code|small|big|abbr|cite|q)\b[^>]*>/gi;

function markupToText_(html) {
  const text = decodeEntities_(html
    .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, ' ')
    .replace(INLINE_TAGS_RE, '')
    .replace(/<[^>]+>/g, ' '))
    .replace(/(\*\*|__|~~|`)/g, '');          // Markdown emphasis in legacy .md pages
  return normalizeText_(text).replace(/\s+/g, ' ').trim();
}

// Characters people can't type the way Docs stores them. Each replacement
// is one character for one, so positions in the text don't move.
function normalizeText_(s) {
  return String(s)
    .replace(/[\u2018\u2019\u201A\u201B\u2032]/g, "'")
    .replace(/[\u201C\u201D\u201E\u201F\u2033]/g, '"')
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g, '-')
    .replace(/[\u00A0\u2007\u202F]/g, ' ');
}

function decodeEntities_(s) {
  return s
    .replace(/&#(\d+);/g, (m, n) => safeChar_(parseInt(n, 10)))
    .replace(/&#x([0-9a-f]+);/gi, (m, n) => safeChar_(parseInt(n, 16)))
    .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&rsquo;|&lsquo;/g, "'")
    .replace(/&rdquo;|&ldquo;/g, '"').replace(/&ndash;|&mdash;/g, '-').replace(/&hellip;/g, '…')
    .replace(/&amp;/g, '&');
}

function safeChar_(code) {
  try { return String.fromCodePoint(code); } catch (e) { return ' '; }
}

function truncate_(s, n) { return s.length > n ? s.slice(0, n) + '…' : s; }

function around_(text, at, len) {
  const start = Math.max(0, at - 60);
  const end = Math.min(text.length, at + len + 60);
  return (start > 0 ? '…' : '') + text.slice(start, end) + (end < text.length ? '…' : '');
}

// A cache value tops out at 100 KB, so a page's text is stored in pieces:
// `key` holds how many, `key#0`, `key#1`… hold the text. A page too long for
// the cap isn't cached at all, so it is always searched in full.
const TEXT_PIECE = 20000;          // chars; under 100 KB even at 4 bytes a char
const TEXT_MAX_PIECES = 40;

function readTextCache_(cache, keys) {
  const heads = getAllBatched_(cache, keys);
  const pieceKeys = [];
  keys.forEach(k => {
    const n = parseInt(heads[k], 10);
    for (let i = 0; i < n; i++) pieceKeys.push(k + '#' + i);
  });
  const pieces = getAllBatched_(cache, pieceKeys);
  const out = {};
  keys.forEach(k => {
    const n = parseInt(heads[k], 10);
    if (!(n >= 0)) return;
    const parts = [];
    for (let i = 0; i < n; i++) {
      const part = pieces[k + '#' + i];
      if (part === undefined) return;       // a piece expired — treat as uncached
      parts.push(part);
    }
    out[k] = parts.join('');
  });
  return out;
}

function writeTextCache_(cache, key, text) {
  const parts = [];
  for (let i = 0; i < text.length; i += TEXT_PIECE) parts.push(text.slice(i, i + TEXT_PIECE));
  if (parts.length > TEXT_MAX_PIECES) return;
  const values = {};
  parts.forEach((part, i) => { values[key + '#' + i] = part; });
  values[key] = String(parts.length);
  try { cache.putAll(values, TEXT_CACHE_SECS); } catch (e) { /* best effort */ }
}

function getAllBatched_(cache, keys) {
  const out = {};
  for (let i = 0; i < keys.length; i += 100) Object.assign(out, cache.getAll(keys.slice(i, i + 100)));
  return out;
}

// ── Drive helpers ───────────────────────────────────────────

function readText_(file) {
  return file.getBlob().getDataAsString('UTF-8');
}

function readAnnotations_(file) {
  const empty = { highlights: [], comments: [] };
  const side = sidecarOf_(file);
  if (!side) return empty;
  try {
    const parsed = JSON.parse(readText_(side));
    return {
      highlights: Array.isArray(parsed.highlights) ? parsed.highlights : [],
      comments: Array.isArray(parsed.comments) ? parsed.comments : []
    };
  } catch (e) { return empty; }
}

// A page by ID — but only one inside the wiki's folder. google.script.run can
// be called with any ID, and when the app runs as its owner that would
// otherwise reach every file the owner can open.
function pageFile_(id) {
  let file;
  try { file = DriveApp.getFileById(String(id)); }
  catch (e) { throw new Error('That page no longer exists — reload the wiki.'); }
  if (file.isTrashed() || !pageKind_(file) || !isInsideRoot_(file)) {
    throw new Error('That page no longer exists — reload the wiki.');
  }
  return file;
}

function isInsideRoot_(file) {
  const rootId = rootFolder_().getId();
  let level = [file];
  for (let depth = 0; depth < 25 && level.length; depth++) {
    const next = [];
    for (const item of level) {
      const parents = item.getParents();
      while (parents.hasNext()) {
        const p = parents.next();
        if (p.getId() === rootId) return true;
        next.push(p);
      }
    }
    level = next;
  }
  return false;
}

function parentOf_(file) {
  const parents = file.getParents();
  if (!parents.hasNext()) throw new Error('That page has no folder.');
  return parents.next();
}

// The page's folder path relative to the root, walking up by first parent.
function folderPathOf_(file) {
  const rootId = rootFolder_().getId();
  const parts = [];
  let f = parentOf_(file);
  for (let i = 0; i < 25 && f.getId() !== rootId; i++) {
    parts.unshift(f.getName());
    const up = f.getParents();
    if (!up.hasNext()) break;
    f = up.next();
  }
  return parts.join('/');
}

function sidecarOf_(file) {
  return childFile_(parentOf_(file), file.getName() + SIDECAR);
}

function childFile_(folder, name) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return null;
}

function childFolder_(folder, name) {
  const it = folder.getFoldersByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return null;
}

// Walks down from the root. With `create`, missing folders are made on the way.
function resolveFolder_(path, create) {
  let folder = rootFolder_();
  normPath_(path).split('/').filter(Boolean).forEach(part => {
    const name = cleanFolderName_(part);
    if (!name) throw new Error('"' + part + '" can\'t be used as a folder name.');
    const found = childFolder_(folder, name);
    if (found) folder = found;
    else if (create) folder = folder.createFolder(name);
    else throw new Error('The folder "' + path + '" no longer exists — reload the wiki.');
  });
  return folder;
}

function normPath_(path) {
  return String(path || '').split('/').map(s => s.trim()).filter(Boolean).join('/');
}

// Human-readable folder names — strip only what a file system would refuse,
// so a folder downloaded from Drive still works as wiki-pages locally.
function cleanFolderName_(raw) {
  return String(raw || '').replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, ' ').trim();
}

function cleanFileName_(raw) {
  let name = String(raw || '').split(/[\\/]/).pop().replace(/[:*?"<>|]/g, '').trim();
  if (!name || name.charAt(0) === '.') name = 'untitled' + name;
  if (!PAGE_RE.test(name)) name += '.html';
  return name;
}

function uniqueName_(folder, name) {
  if (!childFile_(folder, name)) return name;
  const ext = (name.match(/\.[^.]+$/) || ['.html'])[0];
  const base = name.slice(0, name.length - ext.length);
  let i = 2;
  while (childFile_(folder, base + '-' + i + ext)) i++;
  return base + '-' + i + ext;
}

function checkSize_(content) {
  if (String(content || '').length > MAX_PAGE_CHARS) {
    throw new Error('This page is over 10 MB, which is more than Drive will save in one go — split it into smaller pages or use smaller images.');
  }
}

// Drive's own "Access denied" says nothing about what to do. When the app
// runs as the visitor, it means their Drive access to the folder is view-only.
function guard_(fn) {
  try {
    return fn();
  } catch (err) {
    const msg = String(err && err.message || err);
    if (/access denied|permission|insufficient/i.test(msg)) {
      throw new Error("You don't have edit access to the wiki's Drive folder — ask the owner to share it with you as an editor.");
    }
    throw err;
  }
}
