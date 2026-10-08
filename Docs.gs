/* Docs.gs
 * OIS Team Wiki — pages that are Google Docs, Sheets or Slides.
 *
 * Two kinds of wiki page live here, alongside the older .html/.md pages:
 *   gdoc  a Google Doc, Sheet or Slides file inside the wiki folder
 *         (new pages and uploads)
 *   link  a Drive shortcut in the wiki folder pointing at one that lives
 *         somewhere else — the file keeps its owner and sharing
 * Which Google app a page belongs to is its `app`: 'doc', 'sheet' or 'slide'.
 *
 * The wiki shows a Doc by exporting it to HTML; editing happens in Google
 * Docs itself. Everything goes through the Drive REST API with UrlFetchApp,
 * so no advanced service has to be switched on.
 */

const GDOC_MIME = 'application/vnd.google-apps.document';
const GSHEET_MIME = 'application/vnd.google-apps.spreadsheet';
const GSLIDES_MIME = 'application/vnd.google-apps.presentation';
// Google app per file type, and the path segment its URLs use.
const GOOGLE_APPS = {};
GOOGLE_APPS[GDOC_MIME] = { app: 'doc', path: 'document', noun: 'Doc' };
GOOGLE_APPS[GSHEET_MIME] = { app: 'sheet', path: 'spreadsheets', noun: 'Sheet' };
GOOGLE_APPS[GSLIDES_MIME] = { app: 'slide', path: 'presentation', noun: 'Slides file' };
// Office files kept in Drive as they are (a .pptx someone uploaded, say).
// They can be linked without converting: Drive's viewer shows them, Google's
// Office-compatibility editor edits them in place, and search reads the text
// straight out of the file (see officePlainText_).
const OFFICE_APPS = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': { app: 'doc', office: 'docx' },
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': { app: 'sheet', office: 'xlsx' },
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': { app: 'slide', office: 'pptx' },
  'application/msword': { app: 'doc', office: 'doc' },
  'application/vnd.ms-excel': { app: 'sheet', office: 'xls' },
  'application/vnd.ms-powerpoint': { app: 'slide', office: 'ppt' }
};
const SHORTCUT_MIME = 'application/vnd.google-apps.shortcut';
const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';
const UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3/files';

// Upload types and what Google converts each into. Its converters keep Word,
// Excel and PowerPoint formatting better than anything done in the browser.
const IMPORT_TYPES = {
  docx: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', GDOC_MIME],
  doc: ['application/msword', GDOC_MIME],
  odt: ['application/vnd.oasis.opendocument.text', GDOC_MIME],
  rtf: ['application/rtf', GDOC_MIME],
  pdf: ['application/pdf', GDOC_MIME],
  txt: ['text/plain', GDOC_MIME],
  html: ['text/html', GDOC_MIME],
  htm: ['text/html', GDOC_MIME],
  xlsx: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', GSHEET_MIME],
  xls: ['application/vnd.ms-excel', GSHEET_MIME],
  ods: ['application/vnd.oasis.opendocument.spreadsheet', GSHEET_MIME],
  csv: ['text/csv', GSHEET_MIME],
  tsv: ['text/tab-separated-values', GSHEET_MIME],
  pptx: ['application/vnd.openxmlformats-officedocument.presentationml.presentation', GSLIDES_MIME],
  ppt: ['application/vnd.ms-powerpoint', GSLIDES_MIME],
  odp: ['application/vnd.oasis.opendocument.presentation', GSLIDES_MIME]
};

// Google-native or Office file the wiki can show, or null.
function googleAppOf_(mime) { return GOOGLE_APPS[mime] || OFFICE_APPS[mime] || null; }

// Office files open in Google's Office-compatibility mode (rtpof) — edits
// save back into the .pptx/.docx/.xlsx itself.
function fileUrl_(app, id, office) {
  const path = app === 'sheet' ? 'spreadsheets' : app === 'slide' ? 'presentation' : 'document';
  return 'https://docs.google.com/' + path + '/d/' + id + '/edit' + (office ? '?rtpof=true&sd=true' : '');
}
function docUrl_(docId) { return fileUrl_('doc', docId); }

// ── Creating ────────────────────────────────────────────────

// New page: a Google Doc made from a little starter HTML.
function createDoc(folderPath, title, html) {
  assertPerm_('add');
  const name = cleanDocName_(title);
  return withLock_(() => guard_(() => {
    const folder = resolveFolder_(folderPath, true);
    const blob = Utilities.newBlob(html || '<h1>' + escapeHtml_(name) + '</h1>', 'text/html', name + '.html');
    const file = convertToDoc_(folder, name, blob);
    return pageInfo_(file, normPath_(folderPath), null, 'gdoc');
  }));
}

// Upload: the original file's bytes, converted by Google into a Doc, a
// Sheet (Excel, CSV) or Slides (PowerPoint).
function importFile(folderPath, filename, base64) {
  assertPerm_('upload');
  const ext = String(filename || '').split('.').pop().toLowerCase();
  const type = IMPORT_TYPES[ext];
  if (!type) {
    throw new Error('Unsupported file type ".' + ext + '" — upload Word, PDF, text or HTML (becomes a Doc), ' +
      'Excel or CSV (a Sheet), or PowerPoint (Slides).');
  }
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 45 * 1024 * 1024) throw new Error('That file is over 45 MB — too large to convert.');
  const name = cleanDocName_(String(filename).replace(/\.[^.]+$/, ''));
  return withLock_(() => guard_(() => {
    const folder = resolveFolder_(folderPath, true);
    const file = convertToDoc_(folder, name, Utilities.newBlob(bytes, type[0], filename), type[1]);
    return pageInfo_(file, normPath_(folderPath), null, 'gdoc');
  }));
}

// Link: a shortcut to an existing Doc, Sheet or Slides file, found by its URL
// or ID.
function linkDoc(folderPath, urlOrId) {
  assertPerm_('link');
  const id = docIdFromUrl_(urlOrId);
  if (!id) throw new Error("That doesn't look like a Google Docs, Sheets or Slides link.");
  let target;
  try { target = DriveApp.getFileById(id); }
  catch (e) { throw new Error("Can't open that file — make sure the wiki's owner has at least view access to it."); }
  if (!googleAppOf_(target.getMimeType())) {
    throw new Error('Only Google Docs, Sheets and Slides, and Word, Excel or PowerPoint files, can be linked.');
  }
  if (isInsideRoot_(target)) throw new Error('"' + target.getName() + '" is already in the wiki.');
  return withLock_(() => guard_(() => {
    const folder = resolveFolder_(folderPath, true);
    const shortcut = folder.createShortcut(id);
    return pageInfo_(shortcut, normPath_(folderPath), null, 'link');
  }));
}

function docIdFromUrl_(s) {
  const text = String(s || '').trim();
  const m = text.match(/\/d\/([A-Za-z0-9_-]{20,})/) || text.match(/[?&]id=([A-Za-z0-9_-]{20,})/);
  if (m) return m[1];
  return /^[A-Za-z0-9_-]{20,}$/.test(text) ? text : null;
}

// Drive's multipart upload with conversion: metadata says "make this a Doc"
// (or a Sheet, or Slides), the media part is whatever we have.
function convertToDoc_(folder, name, blob, targetMime) {
  const boundary = 'oiswiki' + Utilities.getUuid().replace(/-/g, '');
  const meta = { name: name, mimeType: targetMime || GDOC_MIME, parents: [folder.getId()] };
  const head = Utilities.newBlob(
    '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
    '\r\n--' + boundary + '\r\nContent-Type: ' + blob.getContentType() + '\r\n\r\n').getBytes();
  const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
  const res = UrlFetchApp.fetch(UPLOAD_API + '?uploadType=multipart&supportsAllDrives=true&fields=id', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    payload: head.concat(blob.getBytes(), tail),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('Google Drive could not convert the file: ' + apiError_(res));
  return DriveApp.getFileById(JSON.parse(res.getContentText()).id);
}

// ── Reading ─────────────────────────────────────────────────

// A Doc as HTML the wiki can show: its body, and its stylesheet rewritten so
// every rule only applies inside the page (the export styles `body`, `p`,
// `li`… globally, which would otherwise restyle the whole wiki).
// The export call as a request object, so search can send many at once with
// UrlFetchApp.fetchAll.
function exportRequest_(docId, mime) {
  return {
    url: DRIVE_API + '/' + docId + '/export?mimeType=' + encodeURIComponent(mime || 'text/html'),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
}

function docHtml_(docId) {
  const req = exportRequest_(docId);
  const res = UrlFetchApp.fetch(req.url, req);
  const code = res.getResponseCode();
  if (code === 403 || code === 404) {
    throw new Error("The wiki can't open this Google Doc — it may have been deleted, or its owner hasn't shared it with the wiki's owner.");
  }
  if (code !== 200) throw new Error('Google Drive could not export the Doc: ' + apiError_(res));
  const raw = res.getContentText('UTF-8');

  const css = [];
  raw.replace(/<style[^>]*>([\s\S]*?)<\/style>/gi, (m, c) => { css.push(c); return m; });
  const bodyMatch = raw.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  let body = bodyMatch ? bodyMatch[1] : raw;

  // Links come wrapped in a google.com/url?q=… redirect — point them at the
  // real address.
  body = body.replace(/href="https:\/\/www\.google\.com\/url\?q=([^"&]*)[^"]*"/g, (m, q) => {
    let real = q;
    try { real = decodeURIComponent(q.replace(/&amp;/g, '&')); } catch (e) { /* keep as-is */ }
    return 'href="' + real.replace(/"/g, '%22') + '" target="_blank" rel="noopener"';
  });

  return '<style>' + scopeCss_(css.join('\n')) + '</style><div class="gdoc">' + body + '</div>';
}

// Every tab of a Sheet as text, read through SpreadsheetApp. Drive can only
// export the first tab, so this is the main route; big sheets are capped so
// one huge log sheet can't stall a search.
function sheetPlainText_(id) {
  const ss = SpreadsheetApp.openById(id);
  const parts = [];
  let cells = 0;
  ss.getSheets().forEach(sh => {
    if (cells > 200000) return;
    const rows = sh.getLastRow(), cols = sh.getLastColumn();
    parts.push(sh.getName());
    if (!rows || !cols) return;
    const take = Math.min(rows, Math.max(1, Math.floor((200000 - cells) / cols)));
    sh.getRange(1, 1, take, cols).getDisplayValues().forEach(r => parts.push(r.filter(String).join(' ')));
    cells += take * cols;
  });
  return parts.join('\n');
}

// Text inside an Office file, read from the file itself — Drive can't export
// these. .docx/.pptx/.xlsx are zip archives of XML; the text sits in <w:t>,
// <a:t> and <t> runs. Runs of one paragraph are joined without spaces (Word
// splits words across runs at every formatting change), paragraphs with one.
// The older binary .doc/.ppt/.xls formats can't be read this way.
function officePlainText_(id, office) {
  if (!/^(docx|pptx|xlsx)$/.test(office)) {
    throw new Error('an older Office format (.' + office + ') — open it and save as .' + office + 'x, or convert it to Google, to search inside it');
  }
  const blob = DriveApp.getFileById(id).getBlob();
  blob.setContentType('application/zip');
  const entries = Utilities.unzip(blob);
  const want = office === 'docx' ? /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/
             : office === 'pptx' ? /^ppt\/(slides\/slide\d+|notesSlides\/notesSlide\d+)\.xml$/
             : /^xl\/(sharedStrings|worksheets\/sheet\d+)\.xml$/;
  // slide2 before slide10: compare the name without its number, then the number.
  const stem = name => name.replace(/\d+\.xml$/, '');
  const num = name => { const m = name.match(/(\d+)\.xml$/); return m ? +m[1] : 0; };
  const parts = [];
  entries.filter(e => want.test(e.getName()))
    .sort((a, b) => stem(a.getName()).localeCompare(stem(b.getName())) || num(a.getName()) - num(b.getName()))
    .forEach(e => parts.push(officeXmlText_(e.getDataAsString('UTF-8'))));
  return parts.join('\n');
}

function officeXmlText_(xml) {
  const out = [];
  // A paragraph (<w:p>, <a:p>) or a shared string / inline cell string (<si>, <is>).
  const blockRe = /<(w:p|a:p|si|is)\b[^>]*>([\s\S]*?)<\/\1>/g;
  const runRe = /<(?:w:t|a:t|t)(?:\s[^>]*)?>([^<]*)<\/(?:w:t|a:t|t)>/g;
  let block;
  while ((block = blockRe.exec(xml))) {
    let text = '', run;
    runRe.lastIndex = 0;
    while ((run = runRe.exec(block[2]))) text += run[1];
    if (text) out.push(decodeEntities_(text));
  }
  // Numbers typed into cells live in <v> elements, not strings.
  const numRe = /<c\b[^>]*?(?:t="n"[^>]*)?>\s*<v>([^<]*)<\/v>/g;
  let num;
  while ((num = numRe.exec(xml))) if (!/t="s"/.test(num[0])) out.push(num[1]);
  return out.join(' ');
}

// A Doc's text read through DocumentApp — the fallback when Google won't
// export it. Every tab is included (Docs can have several, and child tabs).
function docPlainText_(docId) {
  const doc = DocumentApp.openById(docId);
  const parts = [];
  const walk = tabs => (tabs || []).forEach(tab => {
    try { parts.push(tab.asDocumentTab().getBody().getText()); } catch (e) { /* not a document tab */ }
    walk(tab.getChildTabs && tab.getChildTabs());
  });
  if (typeof doc.getTabs === 'function') walk(doc.getTabs());
  if (!parts.length) parts.push(doc.getBody().getText());
  return parts.join('\n');
}

// ── Comments ────────────────────────────────────────────────

// Every comment on a Doc page, with its replies, newest first. Plain text
// only — the panel escapes it, so nothing from a comment can inject markup.
function getDocComments(id) {
  assertCanView_();
  const file = pageFile_(id);
  const kind = pageKind_(file);
  if (kind === 'html') throw new Error('This page keeps its comments in the wiki.');
  const docId = docIdOf_(file, kind);
  const fields = 'nextPageToken,comments(id,content,createdTime,modifiedTime,resolved,deleted,' +
    'quotedFileContent/value,author(displayName,emailAddress,me),' +
    'replies(id,content,createdTime,deleted,action,author(displayName,emailAddress)))';
  const out = [];
  let pageToken = '';
  for (let guard = 0; guard < 20; guard++) {
    const url = DRIVE_API + '/' + docId + '/comments?pageSize=100&includeDeleted=false&fields=' +
      encodeURIComponent(fields) + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 403 || code === 404) {
      throw new Error("The wiki can't read this Doc's comments — its owner hasn't shared it with the wiki's owner.");
    }
    if (code !== 200) throw new Error('Google Drive could not list the comments: ' + apiError_(res));
    const data = JSON.parse(res.getContentText());
    (data.comments || []).forEach(c => {
      if (c.deleted) return;
      out.push({
        id: c.id,
        text: c.content || '',
        quote: c.quotedFileContent && c.quotedFileContent.value ? decodeEntities_(c.quotedFileContent.value) : '',
        author: personName_(c.author),
        created: Date.parse(c.createdTime) || 0,
        resolved: !!c.resolved,
        replies: (c.replies || []).filter(r => !r.deleted).map(r => ({
          id: r.id,
          // A reply that only resolves or reopens has an action and no text.
          text: r.content || (r.action === 'resolve' ? 'Marked as resolved' : r.action === 'reopen' ? 'Reopened' : ''),
          action: r.action || '',
          author: personName_(r.author),
          created: Date.parse(r.createdTime) || 0
        }))
      });
    });
    pageToken = data.nextPageToken || '';
    if (!pageToken) break;
  }
  out.sort((a, b) => b.created - a.created);
  return { docUrl: fileUrl_(appOfPage_(file, kind), docId), comments: out };
}

function personName_(p) {
  if (!p) return 'Someone';
  return p.displayName || (p.emailAddress ? p.emailAddress.split('@')[0] : 'Someone');
}

function scopeCss_(css) {
  return css
    .replace(/@import[^;]*;/g, '')
    .replace(/([^{}]+)\{([^{}]*)\}/g, (m, sel, decl) => {
      const s = sel.trim();
      if (!s || s.charAt(0) === '@') return '';
      const scoped = s.split(',').map(x => {
        x = x.trim();
        return /^(html|body)$/i.test(x) ? '.gdoc' : '.gdoc ' + x;
      }).join(',');
      return scoped + '{' + decl + '}';
    });
}

function apiError_(res) {
  try {
    const j = JSON.parse(res.getContentText());
    return (j.error && j.error.message) || res.getResponseCode();
  } catch (e) { return 'HTTP ' + res.getResponseCode(); }
}

function cleanDocName_(raw) {
  const name = String(raw || '').replace(/\s+/g, ' ').trim();
  return name || 'Untitled page';
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
