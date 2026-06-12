# Markdown for Word — your own add-in

A Word task-pane add-in that converts in both directions:

- **Export to MD** — turns the whole document (or just your selection) into Markdown you can copy or download as a `.md` file. Supports headings, bold/italic, links, lists, tables (GitHub style), and code blocks.
- **Image extraction** — embedded pictures are pulled out automatically on export. The Markdown references them by relative path (`![image 1](images/image-1.png)`), and the **Download .zip** button gives you the `.md` file plus an `images/` folder in one bundle — ready to drop into a GitHub repo or wiki.
- **Insert MD** — paste Markdown and insert it as properly formatted Word content, either at your cursor or at the end of the document. Image links (`![alt](url)`) are downloaded and embedded in the document automatically; any that can't be fetched are kept as links and reported.
- **Preview** — see a rendered Markdown view of your document.
- **Save to OneDrive** — one click uploads the exported `.md` straight to your OneDrive (no download/re-upload dance). If images were extracted, it creates a folder with the `.md` plus an `images/` subfolder so the relative links keep working. Requires a one-time setup — see below.

## What's in this folder

| File | Purpose |
|------|---------|
| `manifest.xml` | Tells Word about your add-in (name, icon, where the task pane lives) — points at localhost for local testing |
| `manifest.github.xml` | Same manifest, pre-wired for GitHub Pages hosting (see *Host it on GitHub Pages*) |
| `taskpane.html` | The entire add-in — UI and conversion logic in one self-contained file |
| `assets/` | Ribbon icons (16/32/80 px) |

There is **no build step**. The conversion libraries (marked, turndown, JSZip) load from a CDN.

## Step 1 — Serve the files over HTTPS

Office add-ins must load from an HTTPS address. The easiest local option uses the
official Office dev certificates (requires Node.js):

```bash
cd word-markdown-addin
npx office-addin-dev-certs install     # one-time: creates a trusted localhost cert
npx http-server -S -p 3000 \
  -C ~/.office-addin-dev-certs/localhost.crt \
  -K ~/.office-addin-dev-certs/localhost.key
```

Verify it works by opening https://localhost:3000/taskpane.html in your browser.

> **No Node.js?** Skip this step entirely and host on GitHub Pages instead —
> see *Host it on GitHub Pages* below.

## Host it on GitHub Pages (recommended)

GitHub Pages serves your files over HTTPS for free, so the add-in works from
anywhere with no local server running.

1. **Create the repo** — at github.com, click **New repository**, name it
   `word-markdown-addin`, set it to **Public** (Pages is free on public repos;
   note this makes the code visible to anyone), and create it.
2. **Upload the files** — on the new repo's page, choose **uploading an existing
   file** (or **Add file → Upload files**) and drag the whole project folder in:
   `taskpane.html`, both manifests, and the `assets` folder. Commit.
3. **Turn on Pages** — repo **Settings → Pages** → under *Build and deployment*,
   set Source to **Deploy from a branch**, branch **main**, folder **/ (root)**.
   Save and wait a minute or two.
4. **Verify** — open
   `https://<your-username>.github.io/word-markdown-addin/taskpane.html`
   in a browser. You should see the task pane UI.
5. **Update the manifest** — open `manifest.github.xml` and replace every
   `YOUR-GITHUB-USERNAME` with your actual GitHub username (it appears 9 times —
   find & replace in any text editor). If you named the repo something other than
   `word-markdown-addin`, replace that part of the URLs too.
6. **Sideload `manifest.github.xml`** into Word using the same steps as Step 2
   below — upload this file instead of `manifest.xml`.

> **Using Save to OneDrive?** Your Entra app registration's redirect URIs are tied
> to where the add-in is hosted. In the app registration, add these SPA redirect
> URIs alongside the localhost ones:
> `brk-multihub://<your-username>.github.io` and
> `https://<your-username>.github.io/word-markdown-addin/taskpane.html`.
> (The client ID baked into `taskpane.html` is not a secret — it's safe in a
> public repo.)

> **Updating later:** edit files directly on github.com or push changes with git;
> Pages redeploys automatically in a minute or so. Close and reopen the task pane
> to pick up the new version.

## Step 2 — Sideload the manifest into Word

### Word on the web (easiest)
1. Open a document at office.com in Word.
2. **Home** ribbon → **Add-ins** → **More Add-ins** (or **Add-ins** → **More Settings** → **Upload My Add-in**, depending on your tenant's UI).
3. Choose **Upload My Add-in** and select `manifest.xml`.
4. A **Markdown** button appears on the Home tab — click it to open the pane.

### Word on Windows
1. Share the folder containing `manifest.xml` as a network share (it can be your own machine: right-click folder → Properties → Sharing), e.g. `\\YOURPC\addins`.
2. In Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
3. Add the share path as a catalog URL, check **Show in Menu**, click OK, and restart Word.
4. **Insert → My Add-ins → Shared Folder** tab → pick **Markdown for Word**.

### Word on Mac
1. Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the `wef` folder if needed).
2. Restart Word → **Insert → My Add-ins** → your add-in appears under **Developer Add-ins**.

## Step 3 — Use it

- **Export**: open the pane → *Convert document* → copy or download the Markdown.
- **Insert**: paste Markdown into the *Insert MD* tab → *Insert into document*.
- **Preview**: see your document rendered as Markdown HTML.

## Save to OneDrive setup (one-time)

Uploading to OneDrive means the add-in must sign you in to Microsoft, and Microsoft
requires every app that signs users in to be registered first. It's free and takes
about five minutes:

1. Go to [entra.microsoft.com](https://entra.microsoft.com) → **Identity → Applications → App registrations** → **New registration**.
2. Name it (e.g. `Markdown for Word`) and choose who can use it:
   - *Accounts in this organizational directory only* — just your work tenant
   - *...and personal Microsoft accounts* — also lets personal OneDrive work
3. Under **Redirect URI**, pick **Single-page application (SPA)** and enter:
   `brk-multihub://localhost:3000`
   After creating the app, open **Authentication → Add URI** and also add
   `https://localhost:3000/taskpane.html` (the popup fallback).
   *(If you host the add-in somewhere other than localhost:3000, use that origin instead.)*
4. Open **API permissions** → confirm **Microsoft Graph → Delegated → Files.ReadWrite**
   is listed (add it if not). No admin consent is needed for this permission in most tenants.
5. Copy the **Application (client) ID** from the Overview page.
6. Open `taskpane.html` and paste it into the `GRAPH_CLIENT_ID` constant near the
   top of the script:

   ```js
   var GRAPH_CLIENT_ID = "00000000-0000-0000-0000-000000000000";
   ```

Reload the task pane. The first click on **Save to OneDrive** asks you to sign in
and approve file access; after that it's silent. Files land in the root of your
OneDrive (a folder named after the file when images are included), and a link to
the uploaded file appears in the pane.

## How it works (the 2-minute tour)

Office add-ins are just web pages that Word loads in an embedded browser, plus a
manifest that tells Word where the page lives. The page talks to the document
through the **Office JavaScript API** (`office.js`):

- `Word.run(...)` opens a batch against the document.
- `body.getRange().getHtml()` pulls the document content as HTML →
  **turndown** converts that HTML to Markdown.
- `marked` converts Markdown to HTML → `insertHtml(...)` pushes it back into
  the document with formatting intact.

## Ideas for v2

- Front-matter support (export a YAML block with title/author).
- A folder picker for OneDrive saves (currently files go to the OneDrive root).
- Publish to AppSource so colleagues can install it from the store.

## Troubleshooting

- **Pane is blank** → the HTTPS server isn't running or the certificate isn't trusted. Re-check Step 1 in a browser first.
- **"Upload My Add-in" missing** → your Microsoft 365 admin may have disabled sideloading; ask them to allow user-uploaded add-ins.
- **Tables don't convert** → the GFM plugin script must load; check your network allows cdn.jsdelivr.net.
- **An inserted image stayed a link** → the image host blocked the download (offline, login-required, or CORS-restricted server). The Markdown link is preserved so nothing is lost.
- **"OneDrive isn't set up yet"** → paste your client ID into `GRAPH_CLIENT_ID` in `taskpane.html` (see *Save to OneDrive setup*).
- **Sign-in window never appears or closes instantly** → the redirect URIs in your app registration must exactly match where the add-in is hosted; re-check step 3 of the OneDrive setup.
- **Upload fails with HTTP 403** → your account doesn't have the Files.ReadWrite permission consented; re-check step 4, or your admin may need to approve it.
