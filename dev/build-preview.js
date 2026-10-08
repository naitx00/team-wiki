// Local preview without deploying: node dev/build-preview.js, then open
// dev/preview.html in a browser (add ?viewer to see the read-only view).
//
// Builds preview.html: Index.html with its includes expanded, Code.gs, Api.gs
// and Docs.gs running in the page on the mock services in
// mock-apps-script.js, and google.script.run faked on top.
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..') + '/';
const out = path.join(__dirname, 'preview.html');
// Brand.gs is local (gitignored). `--no-brand` previews the neutral look the
// public repository ships with, even when Brand.gs exists.
const useBrand = !process.argv.includes('--no-brand') && fs.existsSync(dir + 'Brand.gs');
const serverFiles = (useBrand ? ['Brand.gs'] : []).concat(['Code.gs', 'Access.gs', 'Api.gs', 'Docs.gs']);

const mock = `
<script>${fs.readFileSync(path.join(__dirname, 'mock-apps-script.js'), 'utf8')}</script>
<script>${serverFiles.map(f => fs.readFileSync(dir + f, 'utf8')).join('\n')}</script>
<script>
  // Seed: a root folder with a few pages, like an uploaded wiki-pages.
  setup();
  // Access: the preview user is an admin; ?viewer = on the list with no
  // permissions; ?denied = not on the list.
  window.__mock.props.ACCESS = JSON.stringify({ admins: ['editor@example.com'], viewers: 'listed', users: {
    'reader@example.com': { add: false, edit: false, upload: false, link: false, delete: false },
    'member@example.com': { add: true, edit: true, upload: true, link: true, delete: true },
    'limited@example.com': { add: true, edit: true, upload: false, link: true, delete: false } } });
  if (location.search.includes('viewer')) window.__mock.user = 'reader@example.com';
  if (location.search.includes('denied')) window.__mock.user = 'stranger@example.com';
  const asUser = /[?&]as=([^&]+)/.exec(location.search);        // ?as=limited@example.com
  if (asUser) window.__mock.user = decodeURIComponent(asUser[1]);
  (function () {
    const m = window.__mock, root = m.items[m.props.ROOT_FOLDER_ID];
    m.node('file', 'handbook.md', root).content = '# Team Handbook\\n\\nWelcome to the **OIS** team. See [[QA Checklist]].';
    const qa = m.node('folder', 'QA', root);
    m.node('file', 'qa-checklist.html', qa).content = '<h1>QA Checklist</h1><p>Check the banner copy before release.</p>';
    m.node('folder', 'Sub', qa);
    // A page that is a Google Doc, and a link to a Doc kept elsewhere in Drive.
    m.node('file', 'Escalation Process', root, { mime: m.GDOC,
      content: '<h1 class="c1">Escalation Process</h1><p>Call the on-call lead first. ' +
               '<a href="https://www.google.com/url?q=https://example.com&amp;sa=D">Example</a></p>' });
    Object.values(m.items).find(x => x.name === 'Escalation Process').comments = [
      { id: 'AAA1', content: 'Who is the on-call lead this week?', createdTime: '2026-09-29T09:00:00Z', resolved: false,
        quotedFileContent: { value: 'Call the on-call lead first.' }, author: { displayName: 'Maria Santos' },
        replies: [{ id: 'r1', content: 'See the Team Calendar.', createdTime: '2026-09-29T10:00:00Z', author: { displayName: 'Jo Cruz' } }] },
      { id: 'AAA2', content: 'Add the escalation phone number.', createdTime: '2026-09-30T08:00:00Z', resolved: false,
        author: { displayName: 'Jo Cruz' }, replies: [] },
      { id: 'AAA3', content: 'Typo in heading', createdTime: '2026-09-20T08:00:00Z', resolved: true,
        quotedFileContent: { value: 'Escalation Proces' }, author: { displayName: 'Ana Reyes' },
        replies: [{ id: 'r2', content: '', action: 'resolve', createdTime: '2026-09-21T08:00:00Z', author: { displayName: 'Maria Santos' } }] },
      { id: 'AAA4', content: 'deleted one', createdTime: '2026-09-19T08:00:00Z', deleted: true, author: { displayName: 'X' } }
    ];
    m.node('file', 'Weekly Sync', root, { mime: m.GDOC,
      content: '<h2>Issues / Suggestions / Updates</h2><p>Banner QA backlog is growing.</p>' });
    m.node('file', 'Private Notes', root, { mime: m.GDOC, noExport: true, noDocApp: true, content: '<p>secret</p>' });
    const budget = m.node('file', 'Budget 2026', null, { mime: m.GSHEET,
      sheets: [{ name: 'Summary', rows: [['Team', 'Total']] }, { name: 'Vendors', rows: [['Acme Consulting', 'retainer']] }] });
    m.node('file', 'Budget 2026', qa, { mime: m.SHORTCUT, targetId: budget.id, targetMime: m.GSHEET });
    m.node('file', 'Kickoff Deck', root, { mime: m.GSLIDES, content: '<p>Welcome to the OIS kickoff</p>' });
    window.__deckSourceId = m.node('file', 'Roadmap Slides', null, { mime: m.GSLIDES, content: '<p>2027 roadmap</p>' }).id;
    window.__pptxId = m.node('file', 'OIS Onboarding Deck.pptx', null, {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      zip: { 'ppt/slides/slide1.xml': '<p:sld><a:p><a:r><a:t>Welcome aboard</a:t></a:r></a:p></p:sld>' } }).id;
    const cal = m.node('file', 'Team Calendar', null, { mime: m.GDOC, content: '<p>Holidays and on-call rota.</p>' });
    m.node('file', 'Team Calendar', qa, { mime: m.SHORTCUT, targetId: cal.id, targetMime: m.GDOC });
    window.__externalDocId = cal.id;
  })();
  // google.script.run: async, results round-tripped through JSON like the real thing.
  window.google = { script: { get run() {
    let ok = () => {}, fail = () => {};
    const runner = new Proxy({}, { get(_, prop) {
      if (prop === 'withSuccessHandler') return f => { ok = f; return runner; };
      if (prop === 'withFailureHandler') return f => { fail = f; return runner; };
      return (...args) => setTimeout(() => {
        window.__mock.calls.push(prop);
        let res;
        try { res = window[prop](...JSON.parse(JSON.stringify(args))); }
        catch (e) { fail(new Error(e.message)); return; }
        ok(res === undefined ? null : JSON.parse(JSON.stringify(res)));
      }, 30);
    } });
    return runner;
  } } };
</script>`;

// HtmlService strips // and /* */ comments from inline scripts, and the
// stripper only understands '...' and "..." strings. It doesn't recognise
// regex or template literals, so a regex containing // or a quote character
// gets cut apart. That broke the first deployment with "Invalid regular
// expression: missing /". Doing the same here makes the preview fail the same
// way Google would.
const BS = String.fromCharCode(92);
function stripLikeHtmlService(src) {
  let out = '', i = 0, q = null;
  while (i < src.length) {
    const c = src[i], n = src[i + 1];
    if (q) {
      out += c;
      if (c === BS) { out += n || ''; i += 2; continue; }
      if (c === q || c === '\n') q = null;
      i++; continue;
    }
    if (c === '"' || c === "'") { q = c; out += c; i++; continue; }
    if (c === '/' && n === '/') { while (i < src.length && src[i] !== '\n') i++; continue; }
    if (c === '/' && n === '*') { const e = src.indexOf('*/', i + 2); i = e < 0 ? src.length : e + 2; out += ' '; continue; }
    out += c; i++;
  }
  return out;
}
function includeFile(name) {
  const text = fs.readFileSync(dir + name + '.html', 'utf8');
  if (!/^<script>/.test(text)) return text;
  return text.replace(/^<script>([\s\S]*)<\/script>\s*$/, (m, js) => '<script>' + stripLikeHtmlService(js) + '</script>\n');
}

let html = fs.readFileSync(dir + 'Index.html', 'utf8');
html = html.replace(/<\?!= include\('(\w+)'\); \?>/g, (m, name) => includeFile(name));
// Mock services must exist before the boot JSON is computed.
html = html.replace('<script>\n  // Server data rendered', mock + '\n<script>\n  // Server data rendered');
// The parts doGet() fills in: brand logo, colours, name, and boot data.
const vm = require('vm');
const sandbox = {};
vm.createContext(sandbox);
serverFiles.filter(f => f !== 'Api.gs' && f !== 'Docs.gs' && f !== 'Access.gs')
  .forEach(f => vm.runInContext(fs.readFileSync(dir + f, 'utf8'), sandbox, { filename: f }));
const brand = vm.runInContext('brand_()', sandbox);
const escHtml = v => String(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
html = html.replace('<?!= brandLogo ?>', () => vm.runInContext('brandLogoHtml_(brand_())', sandbox));
html = html.replace('<?!= brandCss ?>', () => brand.css ? '<style>' + brand.css + '</style>' : '');
html = html.replace(/<\?= brand\.name \?>/g, () => escHtml(brand.name));
html = html.replace('<title>', '<title>');
html = html.replace('</head>', () => '<title>' + escHtml(brand.title) + '</title>\n</head>');
html = html.replace('<?!= bootJson ?>', () =>
  'Object.assign(JSON.parse(JSON.stringify(getBootstrap())), { brand: ' + JSON.stringify({ name: brand.name, org: brand.org }) + ' })');
if (/<\?(!=|=) (bootJson|brand)/.test(html) || /^<\?!= include/m.test(html)) throw new Error('unexpanded template tag');
fs.writeFileSync(out, html);
console.log('wrote', out, useBrand ? '(with Brand.gs)' : '(neutral, no brand)');
