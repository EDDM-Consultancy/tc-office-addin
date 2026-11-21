# Developer handover — Teamcenter Office add-in (v1.0.0)

Audience: a developer who needs to maintain, debug or extend this add-in. Read `README.md` first for install/deploy.

## 1. Architecture

```
 Word / Excel ──(task pane, HTTPS)──► add-in server (server.js, Node/Express)
   office.js                              ├── static: src/, assets/, manifest.xml, /config.json
   taskpane.js                            ├── /tc/*  ──proxy──► TC web tier  /<tcContext>/JsonRestServices/...
   tc-client.js ──fetch same-origin──►    └── /fms/* ──proxy──► FSC          /fms/fmsdownload, /fms/fmsupload
```

* **Why a proxy:** the task pane runs in a browser sandbox. Calling the TC web tier directly would need CORS on Tomcat/WebLogic and a cross-site session cookie. Proxying keeps everything same-origin; the JSESSIONID cookie path is rewritten to `/tc`.
* **No build step:** plain ES modules served as-is (`src/package.json` marks them as modules for Node tests). `office.js` comes from Microsoft's CDN (required by Microsoft); JSZip is served from `node_modules`.
* **Stateless server:** the Node server keeps no state; the TC session lives in the TC web tier, keyed by the browser's cookie.

## 2. Document ↔ dataset link

A document knows its dataset through **custom document properties**:

| Property | Content |
|---|---|
| `TC_DatasetUID` | dataset UID (the key used for save/check-in) |
| `TC_DatasetType` | `MSWordX` / `MSExcelX` |
| `TC_ItemID`, `TC_RevID`, `TC_DatasetName` | display only |

* **Open:** the file is downloaded, `docprops.js` injects the properties into the OOXML package (`docProps/custom.xml` + content type + root rel) *before* opening it with `Word.createDocument` / `Excel.createWorkbook`, because the new window has no other way to learn the link.
* **Save as new:** the properties are written via the Office API *before* `getFileAsync`, so the uploaded file carries the link too.
* Side effect: the properties stay in files that are later copied out of Teamcenter. *Unlink* deletes them.

## 3. Teamcenter operations used

All via `POST /tc/JsonRestServices/<Service>/<operation>` with envelope `{header:{state,policy}, body}` (`TcClient.call`).

| Function | SOA operation(s) |
|---|---|
| login / session | `Core-2011-06-Session/login`, `Core-2007-01-Session/getTCSessionInfo`, `Core-2006-03-Session/logout` |
| properties | `Core-2006-03-DataManagement/getProperties` |
| search | `Query-2010-04-SavedQuery/findSavedQueries` + `Query-2006-03-SavedQuery/executeSavedQuery` (query `Item...`, by ID and by Name, merged) |
| browse | `revision_list` on Item; relation properties (`IMAN_specification`, `TC_Attaches`, `IMAN_reference`) on revision |
| check-out/in | `Core-2006-03-Reservation/checkout`, `checkin`, `cancelCheckout` |
| download | `ref_list` → `Core-2006-03-FileManagement/getFileReadTickets` → `GET /fms/fmsdownload/<name>?ticket=` |
| upload | `getDatasetWriteTickets` (createNewVersion=true) → `POST /fms/fmsupload/` multipart (`fmsFile`, `fmsTicket`) → `commitDatasetFiles` |
| new item | `generateItemIdsAndInitialRevisionIds` → `createItems` |
| new dataset | `Core-2008-06-DataManagement/createDatasets2` (relation from `createRelation`) |

Error handling: a body whose `.QName` contains `Exception` or a ServiceData with `partialErrors` raises `TcError` with the TC message; the UI shows it. Session-related errors return the user to the sign-in form.

## 4. ⚠ Verify against the real TC 2412 before go-live

The code was tested end-to-end against `mock/mock-tc.js`, which mimics the documented JSON shapes. It has **not yet run against a real Teamcenter**. Check these points first (use the browser dev tools in the task pane: right-click → Inspect, or F12 with Office on the web):

1. **JSON REST endpoint enabled** on the web tier (`/tc/JsonRestServices/...`). Quick test: sign in from the pane.
2. **Map serialisation** — TC returns maps as `[[keys],[values]]`; `mapEntries()` accepts that and plain objects. Affects `getFileReadTickets.tickets` and `generateItemIdsAndInitialRevisionIds`.
3. **FMS URLs** — download/upload paths `/fms/fmsdownload/…` and `/fms/fmsupload/` on the FSC; the add-in server must reach the FSC port (default 4544). If you front FMS through an Active Workspace gateway, point `fscUrl` at the gateway.
4. **Saved query entry names** are localised (`Item ID`/`Name`). Set `itemQueryIdEntry`/`itemQueryNameEntry` to what the `Item...` query shows in the rich client for the server locale.
5. **Named references** `word` (MSWordX) and `excel` (MSExcelX) — confirm in BMIDE.
6. **Item types** in `newItemTypes` must exist and allow ID generation by naming rules; mandatory properties on create (e.g. custom `Document` attributes) will make `createItems` fail → add them as `extendedAttributes` in `TcClient.createItem`.
7. **Check-out granularity** — the add-in checks out the **dataset**. If your process expects the Item Revision to be checked out, extend `saveLinked()`.
8. **Authentication** — Teamcenter user/password via SOA login. TcSS/SSO is **not** implemented (would need the SSO login service + token login `Core-2011-06-Session/login` with SSO credentials).

## 5. Security notes

* Credentials travel browser → add-in server over HTTPS, then server → TC web tier over whatever `tcWebUrl` uses (HTTP on the LAN by default). Use HTTPS to the web tier if the network segment is not trusted.
* The server is an open reverse proxy to the TC web tier and FSC for anyone who can reach it — run it on the internal network only (firewall), not exposed to the internet.
* No credentials are stored; the password field is cleared after sign-in.
* All values rendered in the UI are HTML-escaped (`esc()` in `taskpane.js`).

## 6. Known limitations / ideas

* Opening always creates a **new window**; the task pane must be opened there with the Teamcenter button (auto-open is not possible for documents created from Teamcenter files).
* No local file cache; the Office window holds the only working copy until saved to TC. Closing without saving loses edits (Word/Excel will prompt to save locally, which does **not** update Teamcenter).
* No released-status guard in the UI: saving to a dataset on a released revision fails with the TC access error (by design of TC rules).
* PowerPoint not included; adding it means a `Presentation` host in the manifest, a `PowerPoint` entry in `config.client.hosts` (`MSPowerPointX`/`powerpoint`), and an `openNew` branch (PowerPoint can only *insert* slides from base64, not open a new presentation).
* Possible extensions: Teamcenter properties → Word content controls / Excel named cells (title block sync), revise item from the pane, "My checked-out documents" list via `checked_out_user` query.

## 7. Tests

`npm test` → `test/flow.test.mjs` spins up the mock TC + the real `server.js` app and drives `TcClient` through: bad password, no-session rejection, static/config serving, open → check out → download → stamp props → upload → re-download (byte-equal) → check in, new item + dataset + upload, attach to existing revision, double check-out rejection. Office-specific code (`office-doc.js`, UI) needs manual testing in Word/Excel.

Manual test script (real TC):
1. Sign in. 2. Open tab → search a known item → *Check out & open* a Word dataset → edit → *Save to Teamcenter* with *Check in after save* → verify new dataset version in the rich client. 3. New blank Excel → *Save as new* → New item → verify Item/Rev/dataset. 4. Same with *Existing item revision*. 5. Have a colleague check a dataset out → verify you see "Checked out by …" and Save is blocked.

## 8. Change record

| Version | Date | Change | Impact |
|---|---|---|---|
| 1.0.0 | 2026-11-21 | Initial version: Word + Excel, open/save/save-as-new/check-out/check-in, proxy server, mock + tests | New component; no changes to Teamcenter data model |

When this goes to GitHub, add a change-record entry per merged PR (what changed, impact on users/TC).
