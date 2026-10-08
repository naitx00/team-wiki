# OIS Team Wiki on Google Apps Script

This is the ToolHub wiki (`assets/js/wiki/process-guides.js`) as a standalone
Apps Script web app. Pages live in one Google Drive folder, and everyone on the
team reads and writes that same folder. There's no local copy, no FileZilla
upload and no `pages.js` index to keep in sync.

Folders, drag to reorder, search and find-in-page work as they do in the hub.
Here is what's different:

| | Hub (file system) | Apps Script |
|---|---|---|
| Where pages live | `wiki-pages/` on one PC | A shared Drive folder |
| Who can edit | Whoever types an email in the chip | Real Google account; server-enforced |
| Page list | Whole wiki loaded at once | Titles only; a page's content loads when opened |
| Search | In-memory | On the server, cached per page revision |
| Two people save one page | Last save wins silently | Second saver is warned and chooses |
| Delete | Permanent | Drive trash, restorable for 30 days |
| Comments | Anonymous | Carry the author's account |
| New pages and uploads | `.html` files, edited in the wiki | **Google Docs**, edited in Google Docs |
| Existing Google Docs | Not possible | **Link Google Doc** adds one without moving it |

## Pages that are Google Docs, Sheets or Slides

Everything below applies to all three. Sheets and Slides get their own icon
(green, amber) in the sidebar and title, open in Google's own Sheets or Slides
viewer and editor, and their comments show in the panel. Uploads are converted
by Google: Word, PDF, text and HTML become a Doc; Excel (`.xlsx`, `.xls`,
`.ods`) and CSV/TSV become a Sheet; PowerPoint (`.pptx`, `.ppt`, `.odp`)
becomes Slides. **Link** accepts a Google Docs, Sheets or Slides URL. Search
reads every tab of a Sheet (via SpreadsheetApp; up to 200,000 cells) and the
text of every slide. **New page** still creates a Doc.

**Word, Excel and PowerPoint files kept as-is in Drive** (a `.pptx` someone
uploaded, opened with `rtpof=true` in its URL) can be linked without
converting. They show in Drive's file viewer, **Edit** opens Google's
Office-compatibility editor (which saves back into the `.pptx` itself), and
search reads the text straight out of the file — slides and speaker notes,
Word body/headers/footers, Excel strings and numbers. The older binary formats
(`.ppt`, `.doc`, `.xls`) can be linked and viewed but not searched inside.

- **New page** creates a Google Doc in the folder you're browsing and opens it
  in the embedded Google Docs editor.
- **Edit** on a Doc page swaps the viewer for the real Google Docs editor inside
  the wiki. Google saves as you type; **Done editing** returns to the viewer,
  which shows the saved version. **↗ Open in Google Docs** opens it in its own
  tab instead.
- **Upload document** hands the file to Google, which converts it to a Doc.
  Word (`.docx`/`.doc`), `.odt`, `.rtf`, PDF, text and HTML are all supported.
  Markdown is rendered to HTML first.
- **Link Google Doc** takes a Doc's link and adds a Drive shortcut to it in the
  wiki folder. The Doc keeps its location, owner and sharing. Removing the link
  from the wiki deletes only the shortcut, never the Doc.
- **Showing a Doc:** the wiki embeds Google's own Docs viewer (`/preview`),
  the same renderer Google Docs uses. Fonts, spacing, sizes, page layout,
  headers and footers, and images all look exactly as they do in Docs.
  After editing in another tab, the viewer reloads when you come back to the
  wiki. The wiki's own 🔍 Find can't reach inside Google's
  viewer, so on Doc pages Ctrl+F uses the browser's find. Search across the
  wiki still finds text inside Docs, because the server reads each Doc's text
  for it.
- **Comments:** the 💬 Comments panel lists every comment on the Doc, read
  from Google Drive: quoted text, author, time, replies, and a toggle for
  resolved ones. The badge counts the open comments. Each comment has an
  **Open in Docs** link that jumps to it. Adding, replying and resolving happen
  in Google Docs (**✎ Edit** mode works), and the panel refreshes when you
  finish editing or press ↻ Refresh. The wiki's own highlights and
  comments stay on the older `.html` pages, which also keep the wiki's built-in
  editor.
- **Access:** the viewer shows each Doc with the **reader's own** Google
  access, so readers need at least view access to the Docs. Sharing the wiki's
  Drive folder with the team covers pages created in the wiki; a linked Doc
  keeps its own sharing. Search and linking use the app owner's access. A linked Doc
  the wiki can no longer open is marked ⚠ in the sidebar.

## Files

| File | What it is |
|---|---|
| `Code.gs` | `doGet`, configuration, `setup()` |
| `Access.gs` | Who can open the wiki, the five permissions, the Settings page's server side |
| `AccessSeed.gs` | **Local only (gitignored).** The starting access list; see `AccessSeed.gs.example` |
| `Brand.gs` | **Local only (gitignored).** Name, logos and colours; see `Brand.gs.example` and Branding below |
| `Api.gs` | Every read and write: pages, folders, annotations, search |
| `Docs.gs` | Google Doc pages: create, import, link, and export to HTML (Drive REST API via `UrlFetchApp`) |
| `Index.html` | Page shell |
| `Styles.html` | Hub tokens (`01-tokens.css`) and wiki styles (`06-wiki.css`) |
| `Helpers.html` | `esc`, `loadScriptOnce`, theme, account chip, `pgCall` (the `google.script.run` bridge) |
| `WikiClient.html` | The wiki UI. It's a copy of `process-guides.js` with the storage layer rewritten |
| `appsscript.json` | Manifest: V8, Drive scope, web app settings |
| `dev/` | Local preview, server tests, and the mock Apps Script services they share. Never pushed to Apps Script |

## Setting it up

### 1. Create the project

Using [clasp](https://github.com/google/clasp):

```sh
npm install -g @google/clasp
clasp login
cd apps-script/wiki
clasp create --type webapp --title "OIS Team Wiki" --rootDir .
clasp push
```

`clasp create` writes `.clasp.json`, which is gitignored. To link an existing
project, copy `.clasp.json.example` to `.clasp.json` and fill in the script ID.

Without clasp: go to [script.google.com](https://script.google.com) and create a
project. Add each `.gs` and `.html` file under the same name without the
extension, and paste the contents in. Then turn on **Project Settings → Show
"appsscript.json"** and replace that file's contents with ours.

### 2. Point it at a Drive folder

**If you're moving the existing wiki**, upload the whole `wiki-pages` folder to
Drive by dragging it in. Subfolders and `.annotations.json` files come across
as they are. `pages.js` is ignored. Open `Code.gs`, paste the folder's ID
(the part after `/folders/` in its URL) into `useExistingFolder()`, and run
that function once.

**If you're starting fresh**, run `setup()` once. It creates an "OIS Team Wiki"
folder in your Drive and logs its link.

Either way, the folder ID ends up in **Project Settings → Script Properties**
as `ROOT_FOLDER_ID`, and you can change it there.

The first listing after a migration reads every page once to work out its
title. After that, each page's title, snippet and revision number are stored in
its Drive **description**. Don't edit those descriptions by hand.

### 3. Decide who can edit

Choose one of these deployment models:

`appsscript.json` currently has **B, restricted to you**
(`"executeAs": "USER_DEPLOYING"`, `"access": "MYSELF"`), which is good for
trying it out.

**A. Execute as the user accessing the app** (`"USER_ACCESSING"`).
- Share the Drive folder with the team: **Viewer** for readers, **Editor** for
  editors. Drive sharing becomes the permission system, and Drive's version
  history shows who changed what.
- Each person approves Drive access once, the first time they open the app.
- Known Google problem: if someone's browser is signed into more than one
  Google account, the app can open as "Sorry, unable to open the file at this
  time". Opening it in a window signed into only the work account works.

**B. Execute as me** (`"USER_DEPLOYING"`).
- Nobody else needs access to the folder, and there's no per-person approval.
- Every change is made as you in Drive. The wiki still records comment authors.
- The wiki's own access list (below) decides who gets in and what they can do.

**The access list (⚙ Settings, admins only).** Each person on it can open the
wiki, plus any of five permissions: **Add** (pages and folders), **Edit**
(pages and Google files, moving pages, renaming folders, highlights and
comments), **Upload**, **Link** and **Delete**. Admins manage the list and can
do everything. People not on it see a "you don't have access yet" page, unless
Settings allows anyone at your domain to read. The server checks the
permission on every call; hiding buttons is only a convenience. The list lives
in the `ACCESS` Script Property. The first run seeds it with the OIS team, and
admins can't remove their own admin rights. If it is ever broken, delete
`ACCESS` and the next run re-seeds it. See `Access.gs`.

For other people to reach the wiki at all, the deployment's access must be
**Anyone within your domain** (`"access": "DOMAIN"`); with `"MYSELF"` only you can
open it, whatever the list says.

### 4. Deploy

**Deploy → New deployment → Web app**
- **Execute as:** matches the choice in step 3
- **Who has access:** Only myself while testing; **Anyone within your domain**
  (`"access": "DOMAIN"`) to open it to the team

Share the `/exec` URL. After later changes, deploy with:

```sh
bash dev/deploy.sh "what changed"
```

It pushes and publishes a new version to the same deployment, so the URL
doesn't change. It also **refuses to deploy unless `appsscript.json` says
`"access": "DOMAIN"` (Anyone within your domain) and `"executeAs":
"USER_DEPLOYING"` (Execute as: Me)**, and unless the server tests pass. Every
push uploads `appsscript.json`, and that file decides the web app settings, so
change them there rather than in the Deploy dialog. Otherwise the next push
silently reverts them. (The script uses `clasp push --force`, which Git Bash
needs because clasp can't show its "overwrite the manifest?" prompt there.)

To put the wiki inside the ToolHub, point an iframe or a Quicklink at that URL.
The app allows framing (`XFrameOptionsMode.ALLOWALL`). The reader must be
signed in to their work Google account in that browser.

## Branding

Out of the box the wiki is neutral: a book mark for a logo, "Team Wiki" as its
name, an indigo accent, and "your work Google account" in sign-in messages.
To brand it, copy `Brand.gs.example` to `Brand.gs` and fill in any of: the
wiki's name, the organisation's name, the browser-tab title, a logo for each
theme (as `data:image/png;base64,…` URIs), and CSS that overrides the colour
variables (`--accent`, `--accent-2`, `--accent-glow`, `--navy`, `--on-accent`).

`Brand.gs` is gitignored, so a public copy of the code carries no one's
branding, but clasp still pushes it, so the deployed wiki has it. Preview the
neutral look with `node dev/build-preview.js --no-brand`.

The same goes for `AccessSeed.gs`, the starting access list, which names real
people. Without it, run `setupAccess()` once in the Apps Script editor to
become the first admin, then add people on the Settings page.

## Working on it locally

```sh
node dev/build-preview.js   # then open dev/preview.html
node dev/test-server.js     # exercises every server function on a mock Drive
```

The preview runs the real `Code.gs` and `Api.gs` in the browser against an
in-memory Drive, with sample pages. It also strips the inline scripts the way
HtmlService does, so a script that would break once deployed breaks in the
preview too.

**When editing `WikiClient.html` or `Helpers.html`:** HtmlService removes
`//` and `/* */` comments from inline scripts, and its stripper doesn't
recognise regex or template literals. Inside a regex, write `//` as `\/{2}`,
and write quote characters as `\x22` (`"`), `\x27` (`'`) and `\x60` (`` ` ``).
Otherwise the regex gets cut apart and the page stays blank with "Invalid
regular expression: missing /". Add `?viewer` to the URL to see the
read-only view. Changes are lost on reload.

## Limits worth knowing

- **Page size:** up to about 10 MB, which is Drive's limit for one write.
  Imported Word documents embed their images as base64, so a document with
  many large images can hit this. The error message explains what to do.
- **Search:** matches substrings in titles and page text, like the hub's
  search did. The first search after 6 hours without one reads every page from
  Drive again. That's fine for hundreds of pages; it would be slow for
  thousands.
- **Comments and highlights:** these save as a whole list per page. If two
  people comment on the same page at the same moment, the later save wins.
  Page content itself is protected by the revision check.
- **Drag-to-reorder:** the order is still saved per browser in `localStorage`.
- **Accounts outside your domain:** they get no email from Apps Script, so they're
  read-only.

## Keeping it in step with the hub

`WikiClient.html` and `Styles.html` are copies, and UI changes made to the hub
wiki don't carry over automatically. Everything outside the storage layer is
the same code, so porting a fix is usually a matter of copying the function
across. The functions that differ are `pgLoadPages`, `pgEnsureContent`,
`pgOpenPage`, `pgSaveAnnotations`, `pgCommitSave`, `pgDoDeletePage`,
`pgMovePageToFolder`, the folder functions, `pgDoCreateNewPage`,
`pgHandleUpload`, and search (`pgRunSearch`, `pgMatchesSearch`,
`pgRenderSearchResults`).
