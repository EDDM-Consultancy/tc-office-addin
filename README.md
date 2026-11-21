# Teamcenter add-in for Word and Excel

A **Teamcenter** button on the Home tab of Word and Excel opens a task pane where users can:

- **Open** — search items, browse revisions, open an `MSWordX` / `MSExcelX` dataset in a new window (read-only or *Check out & open*).
- **Save** — save the current document back to its dataset as a new dataset version (auto check-out if needed, optional check-in after save).
- **Save as new** — store an unsaved document in Teamcenter, either as a **new Item** (e.g. `Document`) or attached to an **existing Item Revision**.
- **Check out / Check in / Cancel check-out** for the linked dataset.

Target: Teamcenter 2412 (SOA JSON services on the web tier + FSC), Word/Excel desktop (Microsoft 365 or Office 2021+) and Office on the web.

## How it works (one paragraph)

The task pane is a plain HTML/JS Office web add-in. It is served by a small Node.js server (`server.js`) that also **proxies** `/tc/*` to the Teamcenter web tier and `/fms/*` to the FSC, so the browser never makes cross-origin calls and the Teamcenter session cookie just works. Documents opened from Teamcenter get the dataset UID written into their custom document properties (`TC_DatasetUID` etc.), so any window knows where to save back. Details: `HANDOVER.md`.

## Try it without Teamcenter (mock)

```bash
npm install
npm test                          # end-to-end test against the mock TC server
npm run mock                      # mock Teamcenter on :7001 (login erik / test)
npm run certs                     # one-time: trusted https://localhost dev certificate
TC_WEB_URL=http://localhost:7001 FSC_URL=http://localhost:7001 npm start
```
(Windows PowerShell: `$env:TC_WEB_URL="http://localhost:7001"; $env:FSC_URL="http://localhost:7001"; npm start`)

Sideload in Word/Excel desktop: **Insert → Add-ins → My Add-ins → Upload My Add-in → `manifest.xml`** (or use a shared-folder catalog, below).

## Connect to the real Teamcenter 2412

1. Edit `config.json` → `server`:
   - `tcWebUrl`: web tier base URL, e.g. `http://tcweb01:7001` (no `/tc`)
   - `tcContext`: web app context, normally `tc`
   - `fscUrl`: FSC that serves your volume, e.g. `http://tcfsc01:4544`
2. Check the `client` section against your BMIDE / site (dataset types, named references `word`/`excel`, item types, saved query `Item...` and its entry names — these are **localised**; a Dutch-locale server may call them differently).
3. `npm start` and test sign-in from the task pane with a Teamcenter user.

## Deploy for all users

1. Run `server.js` on an internal server (Windows service via NSSM, or Docker) with a **certificate trusted by the client PCs** (`certFile`/`keyFile` in `config.json`). Office add-ins require HTTPS.
2. Replace every `https://localhost:3000` in `manifest.xml` with that server's URL (e.g. `https://tcoffice.ectm.local`).
3. Distribute the manifest:
   - **Shared folder catalog** (simple, on-prem): put `manifest.xml` on a file share, add the share under *File → Options → Trust Center → Trusted Add-in Catalogs* (or push it by GPO). Users find it under *My Add-ins → Shared Folder*.
   - **Microsoft 365 admin center → Integrated apps → Upload custom app** (pushes it to everyone, also on Office on the web).

## Requirements

| | Needed for |
|---|---|
| WordApi 1.3 | open in new window, custom properties (Word 2016+ / M365) |
| ExcelApi 1.8 / 1.12 | open in new window / custom properties (**Excel 2021 or M365**; Excel 2019 perpetual is too old) |
| Node.js 18+ | add-in server |
| Network | add-in server → TC web tier and FSC ports |

## Files

| Path | What |
|---|---|
| `manifest.xml` | Office add-in manifest (Word + Excel ribbon button) |
| `server.js` | HTTPS host + `/tc` and `/fms` reverse proxy |
| `config.json` | server URLs + Teamcenter data-model settings |
| `src/tc-client.js` | Teamcenter SOA JSON + FMS client |
| `src/office-doc.js` | Word/Excel abstraction (file bytes, custom props, open) |
| `src/docprops.js` | writes the TC link into a .docx/.xlsx package |
| `src/taskpane.*` | UI |
| `mock/mock-tc.js` | in-memory Teamcenter for demo/tests |
| `test/flow.test.mjs` | end-to-end tests (client → proxy → mock) |
| `HANDOVER.md` | developer handover / code review notes |
