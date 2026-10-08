// Run: node dev/test-server.js   (plain Node, no dependencies)
// Loads Code.gs, Api.gs and Docs.gs on the mock services in
// mock-apps-script.js and runs every server function.
const fs = require('fs'), vm = require('vm'), assert = require('assert'), path = require('path');
const dir = path.join(__dirname, '..') + '/';

const ctx = { console, TextEncoder, TextDecoder, Buffer };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, 'mock-apps-script.js'), 'utf8'), ctx);
// A made-up starting list in place of the real, untracked AccessSeed.gs.
const TEST_SEED = [
  "const SEED_ADMINS = ['admin@example.com'];",
  "const SEED_USERS = ['admin@example.com', 'member@example.com', 'teammate@example.com'];"
].join(String.fromCharCode(10));
vm.runInContext(TEST_SEED, ctx, { filename: 'AccessSeed.gs' });
for (const f of ['Code.gs', 'Access.gs', 'Api.gs', 'Docs.gs']) vm.runInContext(fs.readFileSync(dir + f, 'utf8'), ctx, { filename: f });
const { items, node, props } = ctx.__mock;
const call = (fn, ...args) => ctx[fn](...args);
const plain = v => JSON.parse(JSON.stringify(v));
const search = q => plain(call('searchPages', q)).hits;
const ids = q => search(q).map(h => h.id);

// ── not configured ──
// The seeded access list, checked before anything else runs.
ctx.__mock.user = 'admin@example.com';
const seeded = plain(call('whoAmI'));
assert.strictEqual(seeded.isAdmin, true, 'the seed admin is an admin');
const seededAccess = JSON.parse(props.ACCESS);
assert.strictEqual(Object.keys(seededAccess.users).length, 3, 'everyone in the seed added');
assert.ok(Object.values(seededAccess.users).every(p => p.add && p.edit && p.upload && p.link && p.delete), 'all permissions on');
assert.deepStrictEqual(seededAccess.admins, ['admin@example.com']);
ctx.__mock.user = 'member@example.com';
let jess = plain(call('whoAmI'));
assert.ok(jess.canView && !jess.isAdmin && jess.perms.edit, 'listed people can view and edit, not admin');
ctx.__mock.user = 'stranger@example.com';
assert.strictEqual(plain(call('whoAmI')).canView, false, 'people not on the list are kept out');
// The rest of the suite runs as an admin.
props.ACCESS = JSON.stringify({ admins: ['editor@example.com'], viewers: 'listed', users: {} });
ctx.__mock.user = 'editor@example.com';
console.log('seeded access ok');

// A fresh install without AccessSeed.gs: nobody gets in until the owner runs
// setupAccess() in the editor, which makes them the first admin.
{
  const fresh = { console, TextEncoder, TextDecoder, Buffer };
  vm.createContext(fresh);
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'mock-apps-script.js'), 'utf8'), fresh);
  for (const f of ['Code.gs', 'Access.gs', 'Api.gs', 'Docs.gs']) vm.runInContext(fs.readFileSync(dir + f, 'utf8'), fresh, { filename: f });
  fresh.__mock.user = 'someone@example.com';
  const before = JSON.parse(JSON.stringify(fresh.whoAmI()));
  assert.strictEqual(before.canView, false, 'no seed: nobody in');
  assert.strictEqual(fresh.__mock.props.ACCESS, undefined, 'empty list not stored');
  assert.throws(() => fresh.listTree(), /no admin yet: the owner should run setupAccess\(\)/);
  fresh.__mock.owner = 'owner@example.com';            // running it from the editor
  fresh.setupAccess();
  fresh.__mock.user = 'owner@example.com';
  const after = JSON.parse(JSON.stringify(fresh.whoAmI()));
  assert.ok(after.isAdmin && after.canView && after.perms.delete, 'owner is the first admin');
  fresh.setupAccess();                                  // a second run changes nothing
  assert.deepStrictEqual(JSON.parse(fresh.__mock.props.ACCESS).admins, ['owner@example.com']);
  console.log('fresh install + setupAccess ok');
}

assert.throws(() => call('listTree'), /run setup\(\)/);

call('setup');
const rootId = props.ROOT_FOLDER_ID;
assert.ok(rootId);

// ── migrated content: files dropped straight into Drive, no metadata ──
const root = items[rootId];
const legacy = node('file', 'handbook.md', root); legacy.content = '# Team Handbook\n\nWelcome to the **OIS** team.';
const qa = node('folder', 'QA', root);
const cl = node('file', 'checklist.html', qa); cl.content = '<h1>QA &amp; Checklist</h1><p>Check the banner copy.</p>';
const side = node('file', 'checklist.html.annotations.json', qa); side.content = JSON.stringify({ highlights: [{ id: 'h1' }], comments: [] });
node('file', 'pages.js', root).content = 'window.OIS_WIKI_PAGES={}';

let tree = plain(call('listTree'));
assert.deepStrictEqual(tree.folders, ['QA']);
assert.strictEqual(tree.pages.length, 2, 'pages.js and sidecar are not pages');
const byTitle = t => tree.pages.find(p => p.title === t);
assert.ok(byTitle('Team Handbook'), 'md title');
assert.ok(byTitle('QA & Checklist'), 'html title with entity');
assert.strictEqual(byTitle('QA & Checklist').path, 'QA/checklist.html');
assert.strictEqual(byTitle('QA & Checklist').kind, 'html');
assert.strictEqual(byTitle('Team Handbook').snippet, 'Welcome to the OIS team.');
assert.ok(legacy.desc.includes('ois-wiki'), 'metadata stored on first listing');
console.log('listTree ok');

// ── getPage ──
const pg = call('getPage', cl.id);
assert.ok(pg.content.includes('Check the banner'));
assert.strictEqual(pg.annotations.highlights.length, 1);
assert.strictEqual(pg.rev, 0);
assert.throws(() => call('getPage', 'nope'), /no longer exists/);
const outside = node('file', 'secret.html', null);
assert.throws(() => call('getPage', outside.id), /no longer exists/, 'files outside the root are refused');
assert.throws(() => call('getPage', side.id), /no longer exists/, 'sidecars are not pages');
console.log('getPage ok');

// ── save + conflict ──
let r = call('savePage', cl.id, '<h1>QA Checklist v2</h1><p>New text</p>', 0, false);
assert.strictEqual(r.conflict, false); assert.strictEqual(r.rev, 1); assert.strictEqual(r.title, 'QA Checklist v2');
r = call('savePage', cl.id, '<h1>Stale</h1>', 0, false);
assert.strictEqual(r.conflict, true, 'stale revision is refused');
assert.ok(cl.content.includes('v2'), 'nothing written on conflict');
r = call('savePage', cl.id, '<h1>Forced</h1>', 0, true);
assert.strictEqual(r.conflict, false); assert.strictEqual(r.rev, 2);
assert.throws(() => call('savePage', cl.id, 'x'.repeat(9600000), 2, false), /10 MB/);
console.log('savePage ok');

// ── create (legacy html pages), unique names, new folders ──
const a = call('createPage', '', 'banner-qa.html', '<h1>Banner QA</h1>');
const b = call('createPage', '', 'banner-qa.html', '<h1>Banner QA again</h1>');
assert.strictEqual(a.file, 'banner-qa.html'); assert.strictEqual(b.file, 'banner-qa-2.html');
const c = call('createPage', 'Guides/Onboarding', '../../evil.php', '<h1>Deep</h1>');
assert.strictEqual(c.file, 'evil.php.html', 'path stripped, page extension forced');
assert.strictEqual(c.path, 'Guides/Onboarding/evil.php.html');
assert.strictEqual(call('getPage', c.id).rev, 1);
console.log('createPage ok');

// ── annotations ──
call('saveAnnotations', a.id, { highlights: [], comments: [{ id: 'c1', text: 'hi' }] });
assert.strictEqual(call('getPage', a.id).annotations.comments[0].text, 'hi');
call('saveAnnotations', a.id, { highlights: [], comments: [] });
assert.strictEqual(call('getPage', a.id).annotations.comments.length, 0, 'empty annotations trash the sidecar');
call('saveAnnotations', a.id, { highlights: [{ id: 'h9' }], comments: [] });
console.log('saveAnnotations ok');

// ── move (name clash + sidecar follows) ──
call('createPage', 'QA', 'banner-qa.html', '<h1>Other</h1>');
const moved = call('movePage', a.id, 'QA');
assert.strictEqual(moved.path, 'QA/banner-qa-2.html');
assert.strictEqual(call('getPage', a.id).annotations.highlights[0].id, 'h9', 'sidecar moved and renamed with the page');
console.log('movePage ok');

// ── Google Docs: create, import, export ──
const d1 = plain(call('createDoc', 'QA', 'Release Steps', '<h1>Release Steps</h1><p>Step one.</p>'));
assert.strictEqual(d1.kind, 'gdoc');
assert.strictEqual(d1.title, 'Release Steps');
assert.strictEqual(d1.folder, 'QA');
assert.ok(d1.docUrl.startsWith('https://docs.google.com/document/d/' + d1.id));
assert.strictEqual(items[d1.id].mime, ctx.__mock.GDOC);
let dp = call('getPage', d1.id);
assert.strictEqual(dp.kind, 'gdoc');
assert.ok(dp.content.includes('Step one.'));
assert.ok(dp.content.includes('<div class="gdoc">'), 'wrapped for scoping');
assert.ok(!/(^|[}\s])body\{/.test(dp.content) && dp.content.includes('.gdoc{background'), 'body rule scoped to .gdoc');
assert.ok(dp.content.includes('.gdoc .c1{') && dp.content.includes('.gdoc ol{') && dp.content.includes('.gdoc table td,.gdoc table th{'), 'every selector scoped');
assert.ok(!dp.content.includes('@import'), 'remote font imports dropped');
assert.strictEqual(dp.snippet, 'Step one.');
assert.throws(() => call('savePage', d1.id, '<p>x</p>', 0, true), /edit it in Google Docs/);
assert.throws(() => call('saveAnnotations', d1.id, { comments: [{}] }), /Google Docs comments/);
tree = plain(call('listTree'));
assert.strictEqual(tree.pages.find(p => p.id === d1.id).snippet, 'Step one.', 'snippet remembered after first open');
console.log('createDoc + export ok');

// Links inside a Doc lose Google's redirect wrapper.
items[d1.id].content = '<p><a href="https://www.google.com/url?q=https://example.com/a%3Fb%3D1&amp;sa=D&amp;source=editors">site</a></p>';
dp = call('getPage', d1.id);
assert.ok(dp.content.includes('href="https://example.com/a?b=1" target="_blank"'), dp.content);

const docxB64 = Buffer.from('PK fake docx bytes').toString('base64');
const imp = plain(call('importFile', '', 'Onboarding Guide.docx', docxB64));
assert.strictEqual(imp.kind, 'gdoc'); assert.strictEqual(imp.title, 'Onboarding Guide');
const upload = ctx.__mock.fetches.filter(u => u.includes('/upload/')).length;
assert.ok(upload >= 2);
assert.ok(call('getPage', imp.id).content.includes('wordprocessingml'), 'sent with the Word mime type');
assert.throws(() => call('importFile', '', 'virus.exe', docxB64), /Unsupported file type/);
console.log('importFile ok');

// ── Google Docs: link an existing Doc ──
const external = node('file', 'Team Calendar', null, { mime: ctx.__mock.GDOC, content: '<p>Holidays are listed here.</p>' });
const link = plain(call('linkDoc', 'QA', 'https://docs.google.com/document/d/' + external.id + '/edit?usp=sharing'));
assert.strictEqual(link.kind, 'link'); assert.strictEqual(link.title, 'Team Calendar'); assert.strictEqual(link.docId, external.id);
assert.ok(call('getPage', link.id).content.includes('Holidays'));
assert.throws(() => call('linkDoc', '', 'https://docs.google.com/document/d/' + d1.id + '/edit'), /already in the wiki/);
assert.throws(() => call('linkDoc', '', 'not a link'), /doesn't look like/);
const pdf = node('file', 'Scan.pdf', null, { mime: 'application/pdf' });
assert.throws(() => call('linkDoc', '', 'https://drive.google.com/file/d/' + pdf.id + '/view'), /Only Google Docs, Sheets and Slides/);
external.name = 'Team Calendar 2026';
assert.strictEqual(plain(call('listTree')).pages.find(p => p.id === link.id).title, 'Team Calendar 2026', 'title follows the Doc');
console.log('linkDoc ok');

// Comments on a Doc: every page of results, deleted ones dropped, newest first.
external.comments = [
  { id: 'c1', content: 'First', createdTime: '2026-09-01T00:00:00Z', resolved: false, author: { displayName: 'Ann' },
    quotedFileContent: { value: 'Holidays &amp; rota' }, replies: [{ id: 'r1', content: 'Reply', createdTime: '2026-09-02T00:00:00Z', author: { emailAddress: 'bo@example.com' } }] },
  { id: 'c2', content: 'Gone', deleted: true, createdTime: '2026-09-03T00:00:00Z', author: { displayName: 'X' } },
  { id: 'c3', content: 'Newest', createdTime: '2026-09-05T00:00:00Z', resolved: true, author: {},
    replies: [{ id: 'r2', content: '', action: 'resolve', createdTime: '2026-09-06T00:00:00Z', author: { displayName: 'Ann' } }] }
];
let dc = plain(call('getDocComments', link.id));
assert.deepStrictEqual(dc.comments.map(c => c.id), ['c3', 'c1'], 'all pages, no deleted, newest first');
assert.strictEqual(dc.comments[1].quote, 'Holidays & rota');
assert.strictEqual(dc.comments[1].replies[0].author, 'bo', 'name falls back to the email');
assert.strictEqual(dc.comments[0].author, 'Someone');
assert.strictEqual(dc.comments[0].replies[0].text, 'Marked as resolved');
assert.ok(dc.docUrl.includes(external.id));
assert.throws(() => call('getDocComments', cl.id), /keeps its comments in the wiki/);
console.log('getDocComments ok');

// Search finds text inside Docs.
assert.ok(search('holidays').some(h => h.id === link.id));

// Removing a link trashes only the shortcut.
call('deletePage', link.id);
assert.strictEqual(items[link.id].trashed, true);
assert.strictEqual(external.trashed, false, 'the linked Doc itself is untouched');
// A Doc the wiki lost access to is listed as broken and explains itself when opened.
const lost = node('file', 'Private', null, { mime: ctx.__mock.GDOC, content: '<p>x</p>' });
const link2 = plain(call('linkDoc', '', lost.id));
lost.noAccess = true;
tree = plain(call('listTree'));
const broken = tree.pages.find(p => p.id === link2.id);
assert.ok(broken && broken.broken, 'unreadable link is listed, flagged broken');
assert.throws(() => call('getPage', link2.id), /can't open this Google Doc/);
call('deletePage', link2.id);                      // and can still be removed
assert.strictEqual(items[link2.id].trashed, true);
console.log('link delete + broken ok');

// Moving a Doc keeps its name; deleting a Doc trashes it.
const mv = plain(call('movePage', d1.id, 'Guides'));
assert.strictEqual(mv.folder, 'Guides'); assert.strictEqual(mv.title, 'Release Steps'); assert.strictEqual(mv.kind, 'gdoc');
call('deletePage', imp.id);
assert.strictEqual(items[imp.id].trashed, true);
console.log('doc move/delete ok');

// ── Google Sheets and Slides: link, upload, search, comments ──
const GS = ctx.__mock.GSHEET, GP = ctx.__mock.GSLIDES;
const budget = node('file', 'Budget 2026', null, { mime: GS, sheets: [
  { name: 'Summary', rows: [['Team', 'Total'], ['OIS', '1200']] },
  { name: 'Vendors', rows: [['Acme Consulting', 'retainer'], ['Zeta', 'licence']] }] });
const deck = node('file', 'Quarterly Review', null, { mime: GP, content: '<p>Q3 wins and the roadmap</p>' });
const lb = plain(call('linkDoc', 'QA', 'https://docs.google.com/spreadsheets/d/' + budget.id + '/edit#gid=0'));
assert.strictEqual(lb.kind, 'link'); assert.strictEqual(lb.app, 'sheet');
assert.strictEqual(lb.docUrl, 'https://docs.google.com/spreadsheets/d/' + budget.id + '/edit');
const ld = plain(call('linkDoc', 'QA', 'https://docs.google.com/presentation/d/' + deck.id + '/edit?usp=sharing'));
assert.strictEqual(ld.app, 'slide');
assert.ok(ld.docUrl.startsWith('https://docs.google.com/presentation/d/'));
tree = plain(call('listTree'));
assert.strictEqual(tree.pages.find(p => p.id === lb.id).app, 'sheet');
// Search reaches every tab of a Sheet (Drive's CSV export only has the first) and Slides text.
assert.ok(search('acme consulting').some(h => h.id === lb.id), 'second tab of a Sheet searched');
assert.ok(search('roadmap').some(h => h.id === ld.id), 'Slides text searched');
// A Sheet SpreadsheetApp can't open still gets its first tab searched via export.
const ledger = node('file', 'Ledger', root, { mime: GS, noSheetApp: true, sheets: [{ name: 'A', rows: [['walrus', 'x']] }, { name: 'B', rows: [['narwhal']] }] });
assert.ok(search('walrus').some(h => h.id === ledger.id), 'first tab via CSV export');
assert.ok(!search('narwhal').some(h => h.id === ledger.id));
// getPage doesn't try to HTML-export a Sheet.
assert.strictEqual(call('getPage', lb.id).content, '');
// Upload: Excel -> Sheet, PowerPoint -> Slides, CSV -> Sheet.
const b64 = t => Buffer.from(t).toString('base64');
const xl = plain(call('importFile', '', 'Headcount.xlsx', b64('PK xlsx')));
assert.strictEqual(xl.app, 'sheet'); assert.strictEqual(items[xl.id].mime, GS);
assert.ok(items[xl.id].sourceType.includes('spreadsheetml'));
const pp = plain(call('importFile', '', 'Kickoff.pptx', b64('PK pptx')));
assert.strictEqual(pp.app, 'slide'); assert.strictEqual(items[pp.id].mime, GP);
const cv = plain(call('importFile', '', 'contacts.csv', b64('a,b')));
assert.strictEqual(cv.app, 'sheet'); assert.strictEqual(items[cv.id].sourceType, 'text/csv');
assert.throws(() => call('importFile', '', 'movie.mp4', b64('x')), /Unsupported file type/);
// Comments work the same way on Sheets.
budget.comments = [{ id: 's1', content: 'Check Q4 totals', createdTime: '2026-09-01T00:00:00Z', author: { displayName: 'Ann' }, replies: [] }];
const sc = plain(call('getDocComments', lb.id));
assert.strictEqual(sc.comments[0].text, 'Check Q4 totals');
assert.ok(sc.docUrl.includes('/spreadsheets/d/'));
// Removing a link leaves the Sheet alone.
call('deletePage', lb.id);
assert.strictEqual(budget.trashed, false);
console.log('sheets + slides ok');

// ── Office files kept as-is in Drive (a .pptx uploaded to My Drive) ──
const PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const slideXml = runs => '<p:sld><p:txBody><a:p>' + runs.map(t => '<a:r><a:rPr b="1"/><a:t>' + t + '</a:t></a:r>').join('') + '</a:p></p:txBody></p:sld>';
const pptx = node('file', 'OIS Onboarding Deck.pptx', null, { mime: PPTX, zip: {
  'ppt/slides/slide1.xml': slideXml(['Welcome to ', 'O', 'IS']),
  'ppt/slides/slide10.xml': slideXml(['Tenth slide']),
  'ppt/slides/slide2.xml': slideXml(['Escalation ', 'path &amp; contacts']),
  'ppt/notesSlides/notesSlide1.xml': slideXml(['Speaker note: mention the badge']),
  'ppt/presentation.xml': '<p:presentation><a:t>not text</a:t></p:presentation>' } });
// The link exactly as Drive hands it out for a .pptx.
const lp = plain(call('linkDoc', '', 'https://docs.google.com/presentation/d/' + pptx.id + '/edit?usp=drive_link&ouid=105873697318845861159&rtpof=true&sd=true'));
assert.strictEqual(lp.kind, 'link'); assert.strictEqual(lp.app, 'slide'); assert.strictEqual(lp.office, 'pptx');
assert.strictEqual(lp.title, 'OIS Onboarding Deck', 'extension dropped from the title');
assert.strictEqual(lp.docUrl, 'https://docs.google.com/presentation/d/' + pptx.id + '/edit?rtpof=true&sd=true');
assert.strictEqual(call('getPage', lp.id).content, '', 'no HTML export attempted');
assert.ok(ids('welcome to ois').includes(lp.id), 'runs joined: O + IS = OIS');
assert.ok(ids('escalation path & contacts').includes(lp.id), 'entities decoded');
assert.ok(ids('mention the badge').includes(lp.id), 'speaker notes searched');
assert.ok(!ids('not text').includes(lp.id), 'only slide text, not the package');
const pptSnippet = plain(call('searchPages', 'tenth slide')).hits.find(h => h.id === lp.id).snippet;
assert.ok(pptSnippet.indexOf('Escalation') < pptSnippet.indexOf('Tenth'), 'slide 2 before slide 10: ' + pptSnippet);

const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const docx = node('file', 'Policy.docx', root, { mime: DOCX, zip: {
  'word/document.xml': '<w:document><w:body><w:p><w:r><w:t>Data retention is </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">ninety</w:t></w:r><w:r><w:t> days</w:t></w:r></w:p></w:body></w:document>',
  'word/footer1.xml': '<w:ftr><w:p><w:r><w:t>Confidential footer</w:t></w:r></w:p></w:ftr>' } });
assert.ok(ids('retention is ninety days').includes(docx.id), 'Word file dropped in the wiki folder is a page, runs joined');
assert.ok(ids('confidential footer').includes(docx.id));
assert.strictEqual(plain(call('listTree')).pages.find(p => p.id === docx.id).title, 'Policy');

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const xlsx = node('file', 'Rates.xlsx', root, { mime: XLSX, zip: {
  'xl/sharedStrings.xml': '<sst><si><t>Hourly rate</t></si><si><r><t>Senior </t></r><r><t>consultant</t></r></si></sst>',
  'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row><c r="A1" t="s"><v>0</v></c><c r="B1"><v>175</v></c></row></sheetData></worksheet>' } });
assert.ok(ids('senior consultant').includes(xlsx.id));
assert.ok(ids('175').includes(xlsx.id), 'typed numbers searched');

const ppt = node('file', 'Old Deck.ppt', root, { mime: 'application/vnd.ms-powerpoint' });
const oldRes = plain(call('searchPages', 'anything at all'));
assert.ok(/older Office format/.test(oldRes.unreadable.find(u => u.id === ppt.id).reason), 'legacy .ppt explained');
ppt.trashed = true;   // keep later tests' unreadable lists as they were

pptx.comments = [{ id: 'p1', content: 'Update slide 3', createdTime: '2026-09-01T00:00:00Z', author: { displayName: 'Jo' }, replies: [] }];
assert.strictEqual(plain(call('getDocComments', lp.id)).comments[0].text, 'Update slide 3');
call('deletePage', lp.id);
assert.strictEqual(pptx.trashed, false, 'the .pptx itself is kept');
console.log('office files ok');

// ── folders ──
assert.strictEqual(call('renameFolder', 'Guides/Onboarding', 'New Starters'), 'Guides/New Starters');
assert.throws(() => call('renameFolder', 'Guides', 'QA'), /already exists/);
assert.throws(() => call('renameFolder', '', 'x'), /top level/);
tree = plain(call('listTree'));
assert.ok(tree.pages.some(p => p.path === 'Guides/New Starters/evil.php.html'));
call('createFolder', 'Empty');
assert.ok(call('listTree').folders.includes('Empty'), 'empty folders listed');
call('deleteFolder', 'Guides');
tree = plain(call('listTree'));
assert.ok(!tree.folders.some(f => f.startsWith('Guides')));
assert.ok(!tree.pages.some(p => (p.folder || '').startsWith('Guides')));
assert.throws(() => call('deleteFolder', ''), /top level/);
console.log('folders ok');

// ── delete page ──
call('deletePage', b.id);
assert.ok(!call('listTree').pages.some(p => p.id === b.id));
assert.throws(() => call('getPage', b.id), /no longer exists/);
console.log('deletePage ok');

// ── search ──
assert.ok(search('banner').length >= 1);
let hits = search('WELCOME');
assert.strictEqual(hits.length, 1); assert.ok(hits[0].snippet.includes('Welcome'));
assert.strictEqual(search('  ').length, 0);
assert.strictEqual(search('zzzz-not-there').length, 0);
console.log('searchPages ok');

// ── Search: text Google splits, curly quotes, headings, failures, long pages ──
const runbook = node('file', 'Runbook', root, { mime: ctx.__mock.GDOC,
  content: '<h1>Paging rules</h1><p>Call the on-<span class="c1">call</span> lead. The team’s rota — see HR &#8212; now.</p>' });
// (ids is defined near the top)
assert.ok(ids('on-call').includes(runbook.id), 'formatting span mid-word does not split it');
assert.ok(ids("team's").includes(runbook.id), 'straight quote matches curly');
assert.ok(ids('rota - see').includes(runbook.id), 'hyphen matches em dash');
assert.ok(ids('paging rules').includes(runbook.id), 'first heading is searchable');
assert.ok(ids('  ON-CALL   lead ').includes(runbook.id), 'case and spacing ignored');
legacy.content = ['# Team Handbook', '', 'Welcome to the **OIS** team.'].join(String.fromCharCode(10));
assert.ok(ids('the ois team').includes(legacy.id), 'Markdown emphasis ignored');

// A Doc that can't be exported right now is retried and never cached as empty.
const okapi = node('file', 'Zoo', root, { mime: ctx.__mock.GDOC, content: '<p>okapi escalation</p>' });
const realFetch = ctx.UrlFetchApp.fetch;
let failing = true, attempts = 0;
ctx.UrlFetchApp.fetch = (u, o) => {
  if (u.includes(okapi.id) && u.includes('/export')) { attempts++; if (failing) return { getResponseCode: () => 429, getContentText: () => '{}' }; }
  return realFetch(u, o);
};
// While Google is busy: one parallel try, one retry, then DocumentApp reads it anyway.
okapi.noDocApp = true;
assert.strictEqual(ids('okapi').length, 0, 'unreadable every way');
assert.strictEqual(attempts, 2, 'one parallel try plus one retry');
failing = false;
assert.deepStrictEqual(ids('okapi'), [okapi.id], 'found once Google recovers — the failure was not cached');
ctx.UrlFetchApp.fetch = realFetch;

// Long pages are cached in full (in pieces), so repeat searches still reach the end.
const big = node('file', 'big.html', root);
big.content = '<h1>Big</h1><p>' + 'lorem ipsum '.repeat(12000) + 'needle-at-the-end</p>';
assert.ok(ids('needle-at-the-end').includes(big.id), 'first search');
const bigKeys = Object.keys(ctx.__mock.cache).filter(k => k.includes(big.id));
assert.ok(bigKeys.length > 3, 'stored in pieces: ' + bigKeys.length);
const reads = { n: 0 }; const origGet = ctx.DriveApp.getFileById;
ctx.DriveApp.getFileById = id => { if (id === big.id) reads.n++; return origGet(id); };
assert.ok(ids('needle-at-the-end').includes(big.id), 'second search, from cache');
assert.strictEqual(reads.n, 0, 'served from the cache, not Drive');
ctx.DriveApp.getFileById = origGet;
console.log('search robustness ok');

// ── Search: punctuation and spacing don't have to match what was typed ──
const layouts = {
  exact: '<p>Issues/Suggestions/Updates</p>',
  spaced: '<p>Issues / Suggestions / Updates</p>',
  split: '<table><tr><td><p>Issues/</p><p>Suggestions/Updates</p></td></tr></table>',
  words: '<p>Issues, Suggestions and Updates</p>',
  other: '<p>Team lunch schedule</p>'
};
const layoutOf = {};
Object.keys(layouts).forEach(k => { layoutOf[node('file', 'Layout ' + k, root, { mime: ctx.__mock.GDOC, content: layouts[k] }).id] = k; });
const res = plain(call('searchPages', 'Issues/Suggestions/Updates'));
const found = res.hits.filter(h => layoutOf[h.id]).map(h => layoutOf[h.id] + ':' + h.score);
assert.deepStrictEqual(found, ['exact:3', 'spaced:2', 'split:2', 'words:1'], 'ranked best first: ' + found);
assert.ok(res.searched > 5 && Array.isArray(res.unreadable));
console.log('loose matching ok');

// A Doc the wiki can't read is reported, not silently skipped.
const locked = node('file', 'Locked Doc', root, { mime: ctx.__mock.GDOC, content: '<p>Issues/Suggestions/Updates</p>' });
const lockedFetch = ctx.UrlFetchApp.fetch;
ctx.UrlFetchApp.fetch = (u, o) => (u.includes(locked.id) && u.includes('/export'))
  ? { getResponseCode: () => 403, getContentText: () => '{"error":{"message":"forbidden"}}' } : lockedFetch(u, o);
locked.noDocApp = true;                    // unreadable through DocumentApp too
const lockedRes = plain(call('searchPages', 'issues suggestions'));
ctx.UrlFetchApp.fetch = lockedFetch;
assert.ok(!lockedRes.hits.some(h => h.id === locked.id));
assert.deepStrictEqual(lockedRes.unreadable.map(u => u.title), ['Locked Doc']);
assert.ok(/permission/.test(lockedRes.unreadable[0].reason), lockedRes.unreadable[0].reason);
console.log('unreadable reported ok');

// ── Docs Google won't export are read through DocumentApp instead ──
const agenda = node('file', 'OIS Weekly Team Meeting Agenda', root, { mime: ctx.__mock.GDOC,
  exportError: { code: 403, reason: 'exportSizeLimitExceeded', message: 'This file is too large to be exported.' },
  tabs: [{ text: 'Week 40 agenda' }, { text: 'Archive', children: [{ text: 'Week 12: Issues/Suggestions/Updates from QA' }] }] });
let before = ctx.__mock.docAppOpens || 0;
let r2 = plain(call('searchPages', 'Issues/Suggestions/Updates'));
assert.ok(r2.hits.some(h => h.id === agenda.id), 'too-large Doc found via DocumentApp, child tab included');
assert.ok(!r2.unreadable.some(u => u.id === agenda.id));
assert.strictEqual((ctx.__mock.docAppOpens || 0) - before, 1, 'no pointless export retry for a size error');
before = ctx.__mock.docAppOpens;
plain(call('searchPages', 'week 40'));
assert.strictEqual(ctx.__mock.docAppOpens, before, 'cached afterwards');

const blocked = node('file', 'No Download Doc', root, { mime: ctx.__mock.GDOC,
  exportError: { code: 403, reason: 'cannotDownloadFile', message: 'The user does not have sufficient permissions for this file.' },
  content: '<p>penguin roster</p>' });
assert.ok(ids('penguin').includes(blocked.id), 'download-blocked Doc found via DocumentApp');

// When DocumentApp can't read it either, Google's actual reason is shown.
const huge = node('file', 'Huge And Locked', root, { mime: ctx.__mock.GDOC, noDocApp: true,
  exportError: { code: 403, reason: 'exportSizeLimitExceeded', message: 'This file is too large to be exported.' }, content: '<p>x</p>' });
const nodl = node('file', 'Nodl And Locked', root, { mime: ctx.__mock.GDOC, noDocApp: true,
  exportError: { code: 403, reason: 'cannotDownloadFile', message: 'x' }, content: '<p>x</p>' });
const gone = node('file', 'Gone Doc', root, { mime: ctx.__mock.GDOC, noDocApp: true, noExport: true, content: '<p>x</p>' });
const why = {};
plain(call('searchPages', 'anything')).unreadable.forEach(u => { why[u.title] = u.reason; });
assert.ok(/too large/.test(why['Huge And Locked']), why['Huge And Locked']);
assert.ok(/turned off downloading/.test(why['Nodl And Locked']), why['Nodl And Locked']);
assert.ok(/not shared|deleted/.test(why['Gone Doc']), why['Gone Doc']);
console.log('DocumentApp fallback ok');

// ── permissions ──
const setAccess = a => { props.ACCESS = JSON.stringify(a); };
const none = { add: false, edit: false, upload: false, link: false, delete: false };
setAccess({ admins: ['boss@example.com'], viewers: 'listed', users: {
  'reader@example.com': none,
  'adder@example.com': Object.assign({}, none, { add: true }),
  'editorish@example.com': Object.assign({}, none, { edit: true }),
  'uploader@example.com': Object.assign({}, none, { upload: true }),
  'linker@example.com': Object.assign({}, none, { link: true }),
  'deleter@example.com': Object.assign({}, none, { delete: true })
} });
const as = (who, fn) => { ctx.__mock.user = who; try { return fn(); } finally { ctx.__mock.user = 'boss@example.com'; } };
const refused = (who, fn, what) => assert.throws(() => as(who, fn), new RegExp("doesn't have permission to " + what), who + ' ' + what);
const allowed = (who, fn) => as(who, fn);

// Reading needs only to be on the list.
allowed('reader@example.com', () => assert.ok(call('listTree').pages.length > 0));
allowed('reader@example.com', () => call('searchPages', 'welcome'));
assert.strictEqual(as('reader@example.com', () => plain(call('whoAmI'))).canEdit, false);
// Each action needs its own permission.
refused('reader@example.com', () => call('createDoc', '', 'x', ''), 'create pages');
refused('editorish@example.com', () => call('createFolder', 'Nope'), 'create pages or folders');
allowed('adder@example.com', () => call('createFolder', 'Adders'));
allowed('adder@example.com', () => call('createDoc', 'Adders', 'By adder', '<h1>x</h1>'));
refused('adder@example.com', () => call('savePage', cl.id, '<h1>x</h1>', null, true), 'edit pages');
allowed('editorish@example.com', () => call('savePage', cl.id, '<h1>Edited by editor</h1>', null, true));
refused('editorish@example.com', () => call('importFile', '', 'a.docx', Buffer.from('x').toString('base64')), 'upload files');
allowed('uploader@example.com', () => call('importFile', '', 'up.docx', Buffer.from('x').toString('base64')));
refused('uploader@example.com', () => call('linkDoc', '', external.id), 'link files');
const linked = allowed('linker@example.com', () => plain(call('linkDoc', 'Adders', external.id)));
refused('linker@example.com', () => call('deletePage', linked.id), 'delete pages');
refused('editorish@example.com', () => call('deleteFolder', 'Adders'), 'delete pages or folders');
allowed('deleter@example.com', () => call('deletePage', linked.id));
allowed('deleter@example.com', () => call('deleteFolder', 'Adders'));
// Not on the list: no reading either; no email: sign in.
assert.throws(() => as('stranger@example.com', () => call('listTree')), /doesn't have access to this wiki yet — ask boss@example.com/);
assert.throws(() => as('', () => call('listTree')), /Sign in/);
assert.strictEqual(as('stranger@example.com', () => plain(call('getBootstrap'))).denied, true);
// "Anyone at the domain can view" lets colleagues read, not change.
setAccess(Object.assign(JSON.parse(props.ACCESS), { viewers: 'domain' }));
allowed('stranger@example.com', () => call('listTree'));
refused('stranger@example.com', () => call('createFolder', 'x'), 'create pages or folders');
assert.throws(() => as('someone@gmail.com', () => call('listTree')), /doesn't have access/);
// Admins can do everything; case doesn't matter.
allowed('BOSS@example.com', () => assert.strictEqual(plain(call('whoAmI')).isAdmin, true));
console.log('permissions ok');

// ── Settings ──
assert.throws(() => as('reader@example.com', () => call('getSettings')), /Only wiki admins/);
let st = plain(call('getSettings'));
assert.ok(st.people.some(p => p.email === 'boss@example.com' && p.admin));
assert.deepStrictEqual(st.perms, ['add', 'edit', 'upload', 'link', 'delete']);
st.people.push({ email: '  New.Person@Example.com ', admin: false, perms: { add: true, edit: true } });
st.people = st.people.filter(p => p.email !== 'deleter@example.com');
st.viewers = 'listed';
st = plain(call('saveSettings', st));
const np = st.people.find(p => p.email === 'new.person@example.com');
assert.ok(np && np.perms.add && np.perms.edit && !np.perms.delete, 'added, lower-cased, partial permissions');
assert.ok(!st.people.some(p => p.email === 'deleter@example.com'), 'removed');
assert.strictEqual(st.viewers, 'listed');
assert.throws(() => call('saveSettings', { viewers: 'listed', people: [{ email: 'not-an-email', perms: {} }, { email: 'boss@example.com', admin: true, perms: {} }] }), /Not valid email/);
assert.throws(() => call('saveSettings', { viewers: 'listed', people: [{ email: 'boss@example.com', admin: false, perms: {} }] }), /can't remove your own admin/);
assert.throws(() => as('reader@example.com', () => call('saveSettings', st)), /Only wiki admins/);
// A newly added admin can manage settings.
st.people.find(p => p.email === 'new.person@example.com').admin = true;
call('saveSettings', st);
allowed('new.person@example.com', () => assert.ok(plain(call('getSettings')).people.length > 1));
console.log('settings ok');

// ── Drive access denied is translated ──
ctx.__mock.user = 'boss@example.com';
const orig = ctx.DriveApp.getFileById;
ctx.DriveApp.getFileById = id => { const f = orig(id); f.setContent = () => { throw new Error('Access denied: DriveApp.'); }; return f; };
assert.throws(() => call('savePage', cl.id, '<h1>x</h1>', null, true), /edit access to the wiki's Drive folder/);
ctx.DriveApp.getFileById = orig;
console.log('guard ok');

// ── bootstrap JSON is safe inside <script> ──
node('file', 'xss.html', root).content = '<h1></script><script>alert(1)</script></h1>';
const boot = JSON.stringify(call('getBootstrap')).replace(/</g, '\\u003c');
assert.ok(!boot.includes('</script'));
console.log('\nALL SERVER TESTS PASSED');
