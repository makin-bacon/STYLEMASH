# StyleRipper

StyleRipper is a **local-only** browser tool for standardizing text formatting in Microsoft Word documents. Drag in a `.docx`/`.dotx` file, see every distinct text style and appearance it contains (whether from a named Word style or ad-hoc direct formatting), merge groups of them into clean named styles, and save the result — all without ever installing Word or uploading the file anywhere. Everything runs in your browser; the document never leaves your machine.

## Requirements

You need **Node.js** and **npm** installed to run this project. If you already have them, skip to [Install and run](#install-and-run).

### Installing Node.js

StyleRipper needs a current LTS release of Node — **Node 18, 20, or 22+**. (An odd-numbered release like Node 19, 21, or 23 will hit dependency issues — see [Troubleshooting](#troubleshooting).)

Check what you already have:

```bash
node --version
npm --version
```

If that fails, or shows an unsupported version, install Node with one of these:

- **macOS**: `brew install node` (via [Homebrew](https://brew.sh)), or download the installer from [nodejs.org](https://nodejs.org)
- **Windows**: download the installer from [nodejs.org](https://nodejs.org)
- **Any OS (recommended if you'll switch Node versions often)**: use [nvm](https://github.com/nvm-sh/nvm) (macOS/Linux) or [nvm-windows](https://github.com/coreybutler/nvm-windows), then run:
  ```bash
  nvm install --lts
  nvm use --lts
  ```

Confirm it worked by re-running `node --version` — you should see `v18.x`, `v20.x`, or `v22.x` (or higher).

## Install and run

1. Open a terminal and navigate into this project folder:

   ```bash
   cd path/to/STYLERIPPER
   ```

2. Install dependencies (only needed once, or whenever `package.json` changes):

   ```bash
   npm install
   ```

3. Start the local dev server:

   ```bash
   npm run dev
   ```

4. Open the URL it prints — typically **http://localhost:5173** — in your browser (Chrome, Firefox, Edge, or Safari all work).

5. When you're done, press `Ctrl+C` in the terminal to stop the server.

That's it — StyleRipper is now running entirely on your machine. No account, no sign-in, no network access required beyond the initial `npm install`.

## Using the app

1. **Drag and drop** a `.docx` or `.dotx` file onto the browser window (or click to browse for one).
2. StyleRipper reads the document and shows a **Style Report**: every distinct text appearance it found, most common first, each with a live preview, its resolved attributes (font, size, color, bold/italic/underline/strikethrough), and where it came from (a named style, or direct formatting with no style at all).
3. **Select** one or more entries (checkboxes) and click **Do it** to fold them into a single named style — pick an existing style to merge into, or create a new one.
4. Use **+ New Style** in the "User-Created Styles" panel to define a style from scratch, or **Edit** to tweak one you've already created.
5. Need to fix something the UI doesn't expose? Click **Edit XML** on any entry to hand-edit its underlying `<w:rPr>` XML directly.
6. Click **Save locally** at any time to download the result. The file keeps its original format and filename, with `-RIPPED` appended (e.g. `Report.docx` → `Report-RIPPED.docx`).

**Supported formats**: `.docx` and `.dotx` only (Office Open XML). Legacy `.doc` (pre-2007 binary format) isn't supported. Headers, footers, and footnotes/endnotes aren't scanned in this version — only the main document body.

## Available scripts

Run these from the project folder:

| Command | What it does |
|---|---|
| `npm run dev` | Starts the local dev server with hot reload (main way to use the app) |
| `npm run build` | Type-checks the project and builds an optimized static bundle into `dist/` |
| `npm run preview` | Serves the built `dist/` bundle locally, to sanity-check a production build |
| `npm test` | Runs the automated test suite (Vitest) |
| `npm run lint` | Checks the code for lint issues (Oxlint) |

## Troubleshooting

**`npm install` or `npm run dev` fails with an `EBADENGINE` warning or "Cannot find native binding" error.**
This usually means Node is on an unsupported odd-numbered version (e.g. 19, 21, 23). Switch to an LTS version (18, 20, or 22+; see [Installing Node.js](#installing-nodejs) above), then reinstall cleanly:
```bash
rm -rf node_modules package-lock.json
npm install
```

**Port 5173 is already in use.**
Another process (maybe a previous `npm run dev` you forgot to stop) is using it. Either stop that process, or let Vite pick a different port automatically — it will prompt you and tell you the new URL.

**The browser shows a blank page.**
Check the terminal running `npm run dev` for errors, and check your browser's developer console (F12) for errors. Make sure you're opening the exact URL printed in the terminal.

**Word shows a "needs repair" prompt when opening a saved file.**
This shouldn't happen — please open an issue with the document (or a minimal reproduction) attached if you hit it.

## Privacy

StyleRipper does all of its work client-side, in your browser. Your document is never uploaded, transmitted, or stored anywhere outside your own machine.
