/* mock-apps-script.js — an in-memory stand-in for the Apps Script services the
 * wiki uses (DriveApp, the Drive REST API via UrlFetchApp, Cache, Lock,
 * Properties, Session). Runs in Node (dev/test-server.js) and in the browser
 * (dev/build-preview.js). Defines globals; `__mock` exposes the store.
 */
(function (g) {
  const GDOC = 'application/vnd.google-apps.document';
  const GSHEET = 'application/vnd.google-apps.spreadsheet';
  const GSLIDES = 'application/vnd.google-apps.presentation';
  const GOOGLE = [GDOC, GSHEET, GSLIDES];
  const SHORTCUT = 'application/vnd.google-apps.shortcut';
  let nextId = 1, clock = Date.now();
  const items = {};
  const enc = s => Array.from(new TextEncoder().encode(s));
  const dec = b => new TextDecoder().decode(new Uint8Array(b.map(x => x & 255)));

  function node(kind, name, parent, extra) {
    const n = Object.assign({ id: 'id' + (nextId++) + 'abcdefghijklmnopqrstu', kind, name,
      parents: parent ? [parent.id] : [], trashed: false, desc: '', content: '', updated: ++clock }, extra || {});
    items[n.id] = n; return n;
  }
  const iter = arr => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
  const kids = (n, kind) => Object.values(items).filter(x => x.kind === kind && x.parents.includes(n.id));

  function blob(data, mime, name, zipEntries) {
    const bytes = typeof data === 'string' ? enc(data) : data;
    return { getBytes: () => bytes.slice(), getContentType: () => mime, getName: () => name,
             setContentType: t => { mime = t; }, zipEntries: zipEntries || null,
             getDataAsString: () => dec(bytes), data: typeof data === 'string' ? data : dec(bytes), name };
  }
  function wf(n) { return {
    getId: () => n.id, getName: () => n.name, setName: v => { n.name = v; }, getUrl: () => '#' + n.id,
    isTrashed: () => n.trashed, setTrashed: v => { n.trashed = v; },
    createFolder: name => wf(node('folder', name, n)),
    createFile: b => { const f = node('file', b.name, n); f.content = b.data; return wfile(f); },
    createShortcut: targetId => {
      const t = items[targetId];
      return wfile(node('file', t ? t.name : 'shortcut', n, { mime: SHORTCUT, targetId, targetMime: t && t.mime }));
    },
    searchFolders: q => { if (q !== 'trashed = false') throw new Error('unexpected query ' + q); return iter(kids(n, 'folder').filter(x => !x.trashed).map(wf)); },
    searchFiles: q => { if (q !== 'trashed = false') throw new Error('unexpected query ' + q); return iter(kids(n, 'file').filter(x => !x.trashed).map(wfile)); },
    // Like the real DriveApp, these include trashed items.
    getFoldersByName: name => iter(kids(n, 'folder').filter(x => x.name === name).map(wf)),
    getFilesByName: name => iter(kids(n, 'file').filter(x => x.name === name).map(wfile)),
    getParents: () => iter(n.parents.map(id => wf(items[id])))
  }; }
  function wfile(n) { return {
    getId: () => n.id, getName: () => n.name, setName: v => { n.name = v; n.updated = ++clock; },
    getMimeType: () => n.mime || (/\.html?$/i.test(n.name) ? 'text/html' : /\.json$/.test(n.name) ? 'application/json' : 'text/plain'),
    getTargetId: () => n.targetId || null,
    // The shortcut records its target's type, readable without access to the target.
    getTargetMimeType: () => n.targetMime || null,
    // Office files carry `zip`: { 'ppt/slides/slide1.xml': '<xml…>' }.
    getBlob: () => blob(n.content, 'text/plain', n.name, n.zip),
    setContent: c => { n.content = c; n.updated = ++clock; },
    getDescription: () => n.desc, setDescription: d => { n.desc = d; n.updated = ++clock; },
    getLastUpdated: () => new Date(n.updated),
    getParents: () => iter(n.parents.map(id => wf(items[id]))),
    moveTo: f => { n.parents = [f.getId()]; },
    isTrashed: () => n.trashed, setTrashed: v => { n.trashed = v; }
  }; }

  const props = {}, cache = {};
  const mock = { items, node, props, calls: [], user: 'editor@example.com', GDOC, GSHEET, GSLIDES, SHORTCUT, fetches: [] };
  g.__mock = mock;

  g.DriveApp = {
    getFolderById: id => { if (!items[id] || items[id].kind !== 'folder') throw new Error('No item with the given ID'); return wf(items[id]); },
    getFileById: id => { const n = items[id]; if (!n || n.kind !== 'file' || n.noAccess) throw new Error('No item with the given ID'); return wfile(n); },
    createFolder: name => wf(node('folder', name, null))
  };
  g.Utilities = {
    newBlob: (data, mime, name) => blob(data, mime, name),
    base64Decode: s => typeof Buffer !== 'undefined'
      ? Array.from(Buffer.from(s, 'base64'))
      : Array.from(atob(s), c => c.charCodeAt(0)),
    getUuid: () => 'uuid-' + (++clock),
    unzip: b => {
      if (!b.zipEntries) throw new Error('Could not unzip: not a zip file');
      return Object.keys(b.zipEntries).map(k => ({ getName: () => k, getDataAsString: () => b.zipEntries[k] }));
    },
    sleep: () => {}
  };
  g.MimeType = { HTML: 'text/html', PLAIN_TEXT: 'text/plain', GOOGLE_DOCS: GDOC, GOOGLE_SHEETS: GSHEET, GOOGLE_SLIDES: GSLIDES, SHORTCUT };
  g.ScriptApp = { getOAuthToken: () => 'mock-token' };

  const resp = (code, body) => ({ getResponseCode: () => code, getContentText: () => body });
  g.UrlFetchApp = { fetchAll: reqs => reqs.map(r => g.UrlFetchApp.fetch(r.url, r)), fetch: (url, opts) => {
    mock.fetches.push(url);
    // Upload with conversion to a Google Doc (multipart/related).
    if (url.indexOf('/upload/drive/v3/files') !== -1) {
      const boundary = /boundary=(.+)$/.exec(opts.contentType)[1];
      const text = dec(opts.payload);
      const parts = text.split('--' + boundary).slice(1, -1);
      const meta = JSON.parse(parts[0].split('\r\n\r\n')[1]);
      const mediaHead = parts[1].split('\r\n\r\n')[0];
      const media = parts[1].slice(mediaHead.length + 4, -2);
      const type = /Content-Type: (.+)/.exec(mediaHead)[1].trim();
      if (GOOGLE.indexOf(meta.mimeType) === -1) return resp(400, JSON.stringify({ error: { message: 'bad target' } }));
      const parent = items[meta.parents[0]];
      const content = type === 'text/html' ? media : '<h1>' + meta.name + '</h1><p>Converted from ' + type + ' by Google.</p>';
      const extra = { mime: meta.mimeType, content, sourceType: type };
      if (meta.mimeType === GSHEET) extra.sheets = [{ name: 'Sheet1', rows: [['Converted from', type]] }];
      const doc = node('file', meta.name, parent, extra);
      return resp(200, JSON.stringify({ id: doc.id }));
    }
    // Export a Google Doc as HTML, shaped like Google's real export.
    const m = /\/drive\/v3\/files\/([^/]+)\/export\?mimeType=([^&]+)/.exec(url);
    if (m) {
      const want = decodeURIComponent(m[2]);
      const n = items[m[1]];
      if (!n || n.noAccess || n.noExport || n.trashed) return resp(404, JSON.stringify({ error: { message: 'File not found', errors: [{ reason: 'notFound' }] } }));
      // e.g. { code: 403, reason: 'exportSizeLimitExceeded', message: 'This file is too large to be exported.' }
      if (n.exportError) return resp(n.exportError.code, JSON.stringify({ error: { code: n.exportError.code,
        message: n.exportError.message, errors: [{ reason: n.exportError.reason, message: n.exportError.message }] } }));
      if (GOOGLE.indexOf(n.mime) === -1) return resp(403, JSON.stringify({ error: { message: 'Export only supports Docs Editors files.' } }));
      if (n.mime === GSHEET) {
        if (want !== 'text/csv') return resp(400, JSON.stringify({ error: { message: 'bad export type for a sheet' } }));
        const first = (n.sheets || [{ rows: [] }])[0];         // Drive exports only the first tab
        return resp(200, first.rows.map(r => r.join(',')).join('\n'));
      }
      if (n.mime === GSLIDES) {
        if (want !== 'text/plain') return resp(400, JSON.stringify({ error: { message: 'bad export type for slides' } }));
        return resp(200, String(n.content || '').replace(/<[^>]+>/g, ' '));
      }
      if (want !== 'text/html') return resp(400, JSON.stringify({ error: { message: 'bad export type for a doc' } }));
      return resp(200, '<html><head><meta content="text/html; charset=UTF-8" http-equiv="content-type">' +
        '<style type="text/css">@import url(https://themes.googleusercontent.com/fonts/css?kit=abc);' +
        'ol{margin:0;padding:0}table td,table th{padding:0}.c1{color:#cc0000}body{background:#fff;max-width:468pt}h1{font-size:20pt}</style>' +
        '</head><body class="c3 doc-content">' + n.content + '</body></html>');
    }
    // Comments on a Doc, in the Drive v3 response shape. Two per page, to
    // exercise paging.
    const cm = /\/drive\/v3\/files\/([^/]+)\/comments\?(.*)$/.exec(url);
    if (cm) {
      const n = items[cm[1]];
      if (!n || n.noAccess) return resp(404, JSON.stringify({ error: { message: 'File not found' } }));
      const list = n.comments || [];
      const tok = /pageToken=(\d+)/.exec(cm[2]);
      const start = tok ? +tok[1] : 0;
      const page = { comments: list.slice(start, start + 2) };
      if (start + 2 < list.length) page.nextPageToken = String(start + 2);
      return resp(200, JSON.stringify(page));
    }
    return resp(404, 'unexpected url ' + url);
  } };
  // DocumentApp: reads a Doc's text (all tabs) whatever its size.
  const plainOf = html => String(html || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  g.DocumentApp = { openById: id => {
    const n = items[id];
    if (!n || n.noAccess || n.noDocApp || n.trashed || n.mime !== GDOC) {
      throw new Error("Document is missing (perhaps it was deleted, or you don't have read access?)");
    }
    mock.docAppOpens = (mock.docAppOpens || 0) + 1;
    const mkTab = t => ({ asDocumentTab: () => ({ getBody: () => ({ getText: () => t.text }) }),
                          getChildTabs: () => (t.children || []).map(mkTab) });
    const tabs = n.tabs || [{ text: plainOf(n.content) }];
    return { getTabs: () => tabs.map(mkTab), getBody: () => ({ getText: () => tabs[0].text }) };
  } };
  // SpreadsheetApp: every tab's cells, as display values.
  g.SpreadsheetApp = { openById: id => {
    const n = items[id];
    if (!n || n.noAccess || n.noSheetApp || n.trashed || n.mime !== GSHEET) {
      throw new Error("Unexpected error while getting the method or property openById on object SpreadsheetApp.");
    }
    return { getSheets: () => (n.sheets || []).map(sh => ({
      getName: () => sh.name,
      getLastRow: () => sh.rows.length,
      getLastColumn: () => sh.rows.reduce((m, r) => Math.max(m, r.length), 0),
      getRange: (r, c, nr, nc) => ({ getDisplayValues: () => sh.rows.slice(r - 1, r - 1 + nr).map(row => {
        const out = row.slice(c - 1, c - 1 + nc).map(String); while (out.length < nc) out.push(''); return out; }) })
    })) };
  } };
  g.PropertiesService = { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) };
  g.CacheService = { getScriptCache: () => ({
    getAll: ks => { const o = {}; ks.forEach(k => { if (k in cache) o[k] = cache[k]; }); return o; },
    put: (k, v) => { cache[k] = v; },
    putAll: (o) => { Object.keys(o).forEach(k => { cache[k] = o[k]; }); } }) };
  mock.cache = cache;
  g.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  // getEffectiveUser: whoever the script runs as (the editor's user for
  // setupAccess); defaults to the visitor.
  g.Session = { getActiveUser: () => ({ getEmail: () => mock.user }),
                getEffectiveUser: () => ({ getEmail: () => mock.owner || mock.user }) };
  g.Logger = { log: () => {} };
})(typeof window !== 'undefined' ? window : globalThis);
