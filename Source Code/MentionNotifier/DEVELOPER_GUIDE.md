# MentionNotifier — Developer Setup, Debug, Build & Deploy Guide

Audience: a developer who received this project as a zip **without `node_modules`** and has to
set it up, debug it, build it and host it (front end + backend) on their own.

---

## 1. What the solution is

MentionNotifier is a **Word task-pane add-in** (Office.js) plus a small **Node/Express backend**.

| Part | Runs where | Job |
|---|---|---|
| Add-in (`src/`) | Inside Word (browser engine), served as static files by IIS | Scans document comments every 8 s, detects `@username` mentions, shows a comment-management pane, asks the backend to send emails |
| Backend (`backend/server.js`) | Node.js on the same IIS site (via IISNode) | (1) Calls SharePoint `_api/web/siteusers` over NTLM with a **service account** to get the @mention list. (2) Sends notification mail through **SMTP** |

Flow:

```
Word (taskpane.js)
  ├─ GET  /api/siteusers?docUrl=<Office.context.document.url>  ─▶ backend ─NTLM─▶ SharePoint
  └─ POST /api/send-email {to, subject, html}                  ─▶ backend ─SMTP──▶ mail server
```

The add-in calls the backend on the **same origin** (`CONFIG.relayBaseUrl: ""`), which is why the
backend lives under the same IIS site and is reached through the `/api/*` rewrite rule.

---

## 2. Prerequisites

On the **developer machine**:

- Windows 10/11 + **desktop Word** (Microsoft 365 / 2016+). Word on the web also works for
  sideloading but desktop is what `npm start` targets.
- **Node.js LTS** (18 or 20) and npm — check with `node -v`, `npm -v`.
- **Git** (optional), **VS Code** (recommended; `.vscode/` is preconfigured).
- Network access from the dev machine to your SharePoint server and SMTP server if you want the
  backend to work end-to-end locally.

On the **IIS server** (production hosting):

- Windows Server + IIS
- **Node.js LTS** installed at the path named in `web.config` (`C:\Program Files\nodejs\node.exe`)
- **IIS URL Rewrite Module**
- **IISNode** (https://github.com/Azure/iisnode)

No Gulp/Heft/SPFx is needed — this is plain Webpack.

---

## 3. File-by-file reference

```
MentionNotifier/
├─ manifest.xml
├─ package.json / package-lock.json
├─ webpack.config.js
├─ babel.config.json
├─ web.config
├─ .eslintrc.json / .gitignore
├─ .vscode/ (launch.json, tasks.json, settings.json, extensions.json)
├─ assets/                 icons used by the manifest and the pane
├─ src/
│  ├─ taskpane/ (taskpane.html, taskpane.css, taskpane.js)
│  └─ commands/ (commands.html, commands.js)
└─ backend/
   ├─ server.js
   ├─ package.json / package-lock.json
   ├─ .env                 (secrets — NOT in git, you create it)
   └─ .gitignore
```

### 3.1 Front-end / build files

| File | What it does and why it exists |
|---|---|
| **manifest.xml** | The add-in's identity card that Word loads. Contains: `Id` (GUID, never change it once deployed), `Version` (**bump on every release**), `SourceLocation` (the URL of `taskpane.html` — the page Word shows in the pane), `IconUrl`/`HighResolutionIconUrl`, `SupportUrl`, `AppDomains` (extra domains the pane may navigate to), `Hosts` (Document = Word) and `Permissions` (`ReadWriteDocument`, needed to read and edit comments). Every URL here must be reachable by the **end user's** Word. |
| **package.json** | Front-end dependencies and npm scripts (section 5). `config.dev_server_port` (3000) is the dev-server port. Everything under `devDependencies` is build/debug tooling; only `core-js` and `regenerator-runtime` ship in the bundle (polyfills). |
| **package-lock.json** | Pins exact dependency versions so `npm install` gives the recipient the same tree. Keep it in the zip. |
| **webpack.config.js** | The build. Bundles `src/taskpane/taskpane.js` + `src/commands/commands.js`, runs Babel, generates `taskpane.html`/`commands.html`, copies `assets/`, `manifest*.xml` and `web.config` into `dist/`. Also starts the HTTPS dev server on port 3000 using certs from `office-addin-dev-certs`. Holds `urlDev` and `urlProd` (see warning in 3.3). |
| **babel.config.json** | Babel preset-env so modern JS is transpiled for the older embedded browser engines Word uses (browserslist in `package.json` also includes IE 11). |
| **web.config** | IIS configuration, copied to `dist/` root by webpack. Registers IISNode for `backend/server.js`, the **URL rewrite rule** `^api/.*` → `backend/server.js`, IISNode logging (`iisnode/` folder), `watchedFiles` (restart on backend `.js` change) and `httpErrors existingResponse="PassThrough"` so the backend's JSON errors are not replaced by IIS HTML pages. |
| **.eslintrc.json** | Lint rules (Office add-in preset) used by `npm run lint`. |
| **.gitignore** | Excludes `node_modules`, `dist`, logs, etc. |
| **.vscode/launch.json** | F5 debug configurations that attach the VS Code debugger to Word's embedded browser (Edge Chromium / Edge Legacy). |
| **.vscode/tasks.json** | Tasks for build, dev-server, `Debug: Word Desktop`, stop debug, lint. `Debug: Word Desktop` is the pre-launch task for F5. |
| **assets/** | Icons (16–128 px) and images. Copied to `dist/assets/`. The manifest icon URLs point here. |

### 3.2 Source files

| File | What it does |
|---|---|
| **src/taskpane/taskpane.html** | The pane UI markup: comment input with `@` autocomplete, search bar (paste `[ref:XXXXXX]` to find a comment), comment list. Loads `office.js` from Microsoft's CDN (the machine must have internet access to `appsforoffice.microsoft.com`). |
| **src/taskpane/taskpane.css** | Pane styling. |
| **src/taskpane/taskpane.js** (~1000 lines) | **All client logic.** See 3.4. |
| **src/commands/commands.html / commands.js** | Standard Office template leftovers for ribbon commands. `commands.js` calls Outlook `mailbox` APIs and is **not used** — the manifest declares no ribbon commands. Safe to leave; do not rely on it. |

### 3.3 Things you must know about `manifest.xml` and `webpack.config.js`

`webpack.config.js` only rewrites URLs in the manifest during a **production build**, and only if
the manifest contains the exact string `https://localhost:3000/` (`urlDev`) — it then replaces it
with `urlProd`.

**The manifest in this repo already contains hard-coded production URLs**
(`http://mennotifier.regdocs365.com/...`), so:

1. The webpack replacement is currently a **no-op** — `urlProd` (`https://subs-tst.regdocs365.com/`)
   is not what ends up in `dist/manifest.xml`; whatever is typed in `manifest.xml` is.
2. `npm start` (debug) sideloads this manifest as-is, which points Word at **production**, not at
   your local dev server.

Recommended setup for a new developer (pick one):

- **Option A (simplest):** keep two manifests. Copy `manifest.xml` to `manifest.dev.xml`, replace
  every host URL with `https://localhost:3000/`, and run `office-addin-debugging start manifest.dev.xml`.
  Note webpack's `from: "manifest*.xml"` would then copy both to `dist/`; delete
  `dist/manifest.dev.xml` before deploying or tighten the pattern.
- **Option B:** change `manifest.xml` to use `https://localhost:3000/` everywhere and let the
  production build rewrite it to `urlProd`. Then set `urlProd` to your real hosting URL (with
  trailing slash, matching what you typed after `localhost:3000/`).

Either way, keep `urlProd`, `SourceLocation`, icon URLs and `AppDomains` consistent with the URL
users will actually open.

### 3.4 `src/taskpane/taskpane.js` — code guide

| Area (approx. lines) | Purpose |
|---|---|
| Globals + `CONFIG` (1–26) | In-memory caches (`allUsersCache`, `emailCache`, `notifiedMentions`, `cachedDocumentComments`), UI state, `isScanning` lock. `CONFIG.scanInterval` = 8000 ms; `CONFIG.relayBaseUrl` = `""` (same origin). **For local dev against a locally running backend set it to `"http://localhost:5000"` and set it back to `""` before building for production.** |
| `Office.onReady` (28) | Entry point. For Word: preloads users, wires UI, starts the scan loop (first run after 2 s, then every `scanInterval`). |
| `normalizeSiteUsers`, `fetchSiteUsersViaRelay`, `preloadSiteUsers` (48–110) | Calls `GET /api/siteusers?docUrl=...` with `Office.context.document.url`; maps SharePoint JSON to the cache shape. Errors from the backend (`not_sharepoint`, `no_access`, `not_found`) are shown to the user via `showError`. |
| `initAutocomplete`, `renderSuggestions`, `applySuggestionHighlight` (112–245) | `@` mention dropdown with keyboard navigation. |
| `handleFormSubmission`, `createNativeCommentFromTaskpane`, `createReplyToCommentInWord`, `updateNativeCommentInWord`, `deleteNativeCommentFromWord`, `navigateToCommentInDoc` (247–445) | Word JS API calls that create / reply / edit / delete native Word comments and scroll to one. |
| `enterEditMode`, `enterReplyMode`, `resetFormState`, `formatCommentTimestamp`, `computeRefLabel`, `buildCommentNodeHTML`, `renderCommentsList` (445–700) | Pane UI rendering and form state. |
| `scanAllComments` (729) | Background loop: loads all comments + replies via `body.getComments()`, runs `processComment` on each, rebuilds `cachedDocumentComments`. Skips read-only documents and uses `isScanning` to avoid overlapping runs (prevents duplicate emails). |
| `isKnownUsername`, `processComment` (813–877) | Finds `@mentions` with a regex, ignores names not in `allUsersCache`, appends a unique `[ref:XXXXXX]` token to the comment text, and tracks who was already notified in `notifiedMentions` so each person is mailed **once per comment**. Mentions already present on first load are treated as already notified (no re-spam after reopening). |
| `sendNotificationSandbox` (879) | Builds the HTML email and subject, resolves the recipient, calls the backend. |
| `sendEmailViaRelay` (940) | `POST /api/send-email`. |
| `resolveUserEmailLocal` (956) | Maps `@username` → email using SharePoint `LoginName`/`Title`, preferring a record that actually has an email. |
| `buildCleanDocUrl` (993) | Converts the WOPI/Office-web URL into a plain document link for the email. |

Note: `notifiedMentions` is in memory only. Closing and reopening the pane/document resets it; the
`[ref:...]` token in the comment text is what prevents re-notifying already-processed comments.

### 3.5 `backend/server.js` — code guide

| Piece | Purpose |
|---|---|
| `dotenv` + required env vars | Loads `backend/.env`. **Process exits** if `SP_DOMAIN`, `SP_USERNAME`, `SP_PASSWORD` or `TRUSTED_SP_HOSTS` is missing. |
| Logging + `uncaughtException` / `unhandledRejection` handlers | Timestamped logs that end up in IISNode's `iisnode/` log files; keeps `httpntlm` errors from silently killing the process. |
| Access-log middleware | Logs method, path, status, duration for every request. |
| `TRUSTED_SP_HOSTS` + `isTrustedSharePointUrl` | **Security control (SSRF/NTLM-relay protection).** Only documents whose host is in the allow-list are queried. |
| `ntlmRequest` | Promise wrapper over `httpntlm` using the service account. Sends `X-FORMS_BASED_AUTH_ACCEPTED: f` so multi-provider claims zones treat it as a Windows-auth client. |
| `findSiteUsersForUrl` | Given a document URL, tries `<prefix>/_api/web/siteusers` starting at the full path and trimming one segment at a time until a site answers. 401/403 stops immediately (`no_access`); nothing resolves → `not_found`. This is why new sites need no configuration. |
| `mailTransporter` (nodemailer) | SMTP transport built from `SMTP_*` env vars; auth only if user+pass set; `verify()` logs readiness at startup. `tls.rejectUnauthorized:false` is set to avoid handshake failures on internal networks — tighten this if your SMTP server has a valid certificate. |
| `GET /api/health` | Returns `{status:"ok", trustedHosts:[...]}`. First thing to test after any deployment. |
| `GET /api/siteusers` | Validates `docUrl`, resolves site, returns `{success, users}`. Errors: 400 `not_sharepoint`, 403 `no_access`, 404 `not_found`, 502 other. |
| `POST /api/send-email` | Body `{to, subject, html}` → sends via SMTP. 400 if a field is missing, 502 on SMTP failure. |
| Final error middleware | Logs stack traces for anything unhandled, returns JSON 500. |
| `app.listen(PORT \|\| 5000)` | Under IISNode `PORT` is a named pipe supplied by IIS; 5000 is used for standalone runs only. |

---

## 4. First-time setup from the zip

Open a terminal in the project root (the folder containing `manifest.xml`).

```powershell
# 1. Front-end dependencies (creates node_modules/)
npm install

# 2. Backend dependencies (separate package.json)
cd backend
npm install
cd ..
```

If `npm install` fails on a corporate network, configure the npm registry/proxy
(`npm config set proxy ...`) and retry.

### 4.1 Create `backend/.env`

`.env` is excluded from git and from a clean zip — create it yourself at `backend/.env`:

```ini
# Standalone/local testing only. IIS ignores this and supplies a named pipe.
PORT=5000

# SharePoint service account (NTLM). Needs READ access on every location where
# documents live. Easiest for many site collections: a Web Application User
# Policy (Central Admin) so current and future sites are covered.
SP_DOMAIN=YOURDOMAIN
SP_USERNAME=svc-mentionnotifier
SP_PASSWORD=change-me

# Hostnames (comma separated) of the SharePoint servers whose documents are allowed.
# Host only - no protocol, no path. Must match the host in the document URL
# exactly (case-insensitive), e.g. spse01 and spse01.corp.local if both are used.
TRUSTED_SP_HOSTS=spse01,spse01.corp.local

# SMTP for notification mail
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURE=false            # true only for port 465
SMTP_USER=notifications@yourdomain.com
SMTP_PASS=app-password-here  # Gmail needs an App Password
SMTP_FROM=notifications@yourdomain.com
SMTP_TIMEOUT=30000
```

Rules:

- Leave `SMTP_USER`/`SMTP_PASS` empty for relays that allow anonymous submission.
- Never commit `.env`; never put secrets in `src/` or `manifest.xml`.
- When sharing this project with someone, send `.env` separately or share a sanitized
  `.env.example`.

---

## 5. npm scripts (root `package.json`)

| Command | What it does |
|---|---|
| `npm run dev-server` | Webpack dev server on `https://localhost:3000` (hot rebuild, dev certs). |
| `npm start` | `office-addin-debugging start manifest.xml` — starts dev server, sideloads the manifest into desktop Word and opens it. |
| `npm stop` | Stops debugging and unregisters the sideloaded manifest. |
| `npm run build:dev` | Unminified build into `dist/`, manifest URLs unchanged. |
| `npm run build` | **Production build** into `dist/` (minified, manifest URLs rewritten per section 3.3). |
| `npm run watch` | Rebuild on file change (dev mode). |
| `npm run validate` | Validates `manifest.xml` against Office schema. Run before every release. |
| `npm run lint` / `lint:fix` | ESLint check / auto-fix. |
| `npm run signin` / `signout` | Microsoft 365 account login for M365 sideloading (rarely needed here). |

Backend (`cd backend`): `npm start` → `node server.js`.

---

## 6. Local development and debugging

### 6.1 One-time: trust the dev certificate

```powershell
npx office-addin-dev-certs install
```

Accept the Windows prompt. Without this Word will refuse to load `https://localhost:3000`.

### 6.2 Run the backend locally

```powershell
cd backend
npm start
```

Expected log: `Mention Notifier backend active on port: 5000` and
`[SMTP] Server is ready to send messages.`
Check `http://localhost:5000/api/health`.

The local machine must be able to reach SharePoint and SMTP. If it exits immediately, a required
`.env` value is missing (the log says which).

### 6.3 Point the add-in at the local backend

In `src/taskpane/taskpane.js` set:

```js
relayBaseUrl: "http://localhost:5000",   // dev only — revert to "" for production builds
```

The backend already uses `cors()`, so calls from `https://localhost:3000` work.
(Mixed content: an https page calling `http://localhost` is allowed by Chromium-based engines;
if Word's engine blocks it, run the backend behind HTTPS or test against the deployed backend.)

### 6.4 Run the add-in in Word

1. Use a dev manifest (see 3.3) pointing to `https://localhost:3000/`.
2. `npm start` (or in VS Code: **Run → "Word Desktop (Edge Chromium)"**, F5).
3. Word opens; open a document **from a SharePoint library** (the backend rejects non-SharePoint
   URLs with "not recognized as opened from a SharePoint site"). Open the pane from the Home
   ribbon (or Insert → My Add-ins → Developer Add-ins).
4. When finished: `npm stop`.

### 6.5 Debugging the front end

- **VS Code:** F5 using `.vscode/launch.json`; set breakpoints in `src/taskpane/taskpane.js`
  (source maps are on via `devtool: "source-map"`).
- **Without VS Code:** right-click inside the pane → **Inspect** (Edge Chromium/WebView2) for
  Console, Network and Sources. `console.log` lines are prefixed `[MentionNotifier]`.
- If the pane is blank/white: check Network tab for failing `taskpane.js`/`office.js` and verify
  the dev certificate is trusted.
- If you changed the manifest, run `npm stop` then `npm start` so Word reloads it.
- Word caches add-in pages: clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` if changes do not
  appear.

### 6.6 Debugging the backend

- Locally: watch the terminal; or run `node --inspect server.js` and attach VS Code
  ("Attach to Node Process").
- Useful endpoints: `/api/health`, `/api/siteusers?docUrl=<encoded document URL>`.
- Test email: 

  ```powershell
  curl -X POST http://localhost:5000/api/send-email -H "Content-Type: application/json" `
       -d '{"to":"you@domain.com","subject":"test","html":"<b>hi</b>"}'
  ```

- Log interpretation: `[NTLM] Candidate ... -> HTTP xxx` shows each URL level tried;
  `SPRequestGuid` can be searched in SharePoint ULS logs; `[SMTP ERROR]` shows mail failures.

---

## 7. Building a release

1. Make the code change.
2. Make sure `relayBaseUrl` is `""` in `taskpane.js`.
3. If URLs/hosting changed: update `urlProd` in `webpack.config.js` and the URLs in `manifest.xml`.
4. **Bump `<Version>` in `manifest.xml`** (e.g. `1.0.0.3` → `1.0.0.4`) — Office only refreshes a
   deployed add-in when the version increases.
5. Validate and build:

   ```powershell
   npm run validate
   npm run build
   ```

6. Output is in **`dist/`**: `taskpane.html`, `taskpane.js`, `commands.html/js`, `polyfill.js`,
   `assets/`, `manifest.xml`, `web.config`. Check `dist/manifest.xml` contains the correct URLs.

`dist/` is regenerated (and cleaned) on every build — never hand-edit it.

---

## 8. Hosting on IIS

### 8.1 One-time server setup

1. Install Node.js LTS (default path `C:\Program Files\nodejs\`; otherwise edit
   `nodeProcessCommandLine` in `web.config`).
2. Install the IIS URL Rewrite Module, then IISNode.
3. Create an IIS site whose physical path is a new folder (e.g. `C:\inetpub\MentionNotifier`) and
   bind your hostname/port/certificate. This host is the one in `manifest.xml`.
4. Give the app pool identity **modify** rights on the `iisnode` log folder (IISNode creates it
   under the site root) and read rights everywhere else.
5. Ensure the server can reach SharePoint and the SMTP host, and that Word clients can reach the
   site URL (HTTPS is strongly recommended; Word on the web and many tenants require it).

### 8.2 Deployed folder layout

```
<IIS site root>/
├─ taskpane.html, taskpane.js, commands.html, commands.js, polyfill.js
├─ assets/
├─ manifest.xml
├─ web.config                    <- must be at the ROOT, not inside backend/
├─ iisnode/                      <- auto-created logs
└─ backend/
   ├─ server.js
   ├─ package.json
   ├─ node_modules/              <- run npm install --omit=dev on the server (or copy)
   └─ .env                       <- real values
```

`backend/` is **not** part of the webpack output; copy it separately, keeping the relative path
`backend/server.js`.

### 8.3 Deploying a front-end-only change

1. `npm run build`
2. Copy the contents of `dist/` over the site root (keep `backend/` and `iisnode/`).
3. Browse `https://<host>/taskpane.html` — the UI should render.
4. If `manifest.xml` changed (version bump), redeploy the manifest (section 9).

### 8.4 Deploying a backend change

1. Edit `backend/server.js` (and `backend/package.json` if dependencies changed).
2. Copy `backend/server.js` (+ `package.json`) to `<site root>\backend\`.
   If dependencies changed run `npm install --omit=dev` inside the server's `backend` folder.
3. Update `backend/.env` on the server if new variables are needed — **do not overwrite it with a
   dev copy**.
4. Restart: `web.config` has `watchedFiles="backend/*.js"`, so replacing `server.js` recycles the
   Node process automatically. For `.env` or `node_modules` changes recycle the app pool:
   `Restart-WebAppPool <poolname>` or `iisreset`.
5. Verify `https://<host>/api/health`. If it 404s, the URL Rewrite/IISNode setup is wrong; if it
   500s, read `iisnode/*.txt`.

### 8.5 Production troubleshooting

| Symptom | Check |
|---|---|
| `/api/health` 404 | URL Rewrite module installed? `web.config` at site root? IISNode installed? |
| `/api/health` 500 / process dies | Latest file in `iisnode/` — usually a missing `.env` value (`[CRITICAL] Missing ...`) or missing `node_modules`. |
| Pane says "not recognized as opened from a SharePoint site" | Document host not in `TRUSTED_SP_HOSTS`, or document was not opened from SharePoint. |
| Pane says "contact your admin ... does not have access" | Service account has no Read on that site; grant via Web App User Policy. See `SPRequestGuid` in logs. |
| "No SharePoint site could be found" | `_api/web/siteusers` failed at every URL level — wrong host binding/zone or NTLM not enabled on that zone. |
| Mentions work but no email | `[SMTP ERROR]` in logs; check `SMTP_*`, Gmail app password, firewall to SMTP port; check the user has an Email in SharePoint. |
| Old UI still shown | Browser/Office cache, or manifest version not bumped. |
| Browser shows HTML 502 instead of JSON | `httpErrors existingResponse="PassThrough"` missing from `web.config`. |

---

## 9. Distributing the add-in to users

`manifest.xml` is what users/admins install. Options:

- **Microsoft 365 admin center → Integrated apps → Upload custom apps** (Centralized Deployment).
- **Shared network folder catalog:** put `manifest.xml` in a shared folder; each user adds it in
  Word → File → Options → Trust Center → Trusted Add-in Catalogs; then Insert → My Add-ins →
  Shared Folder.
- **SharePoint app catalog** for on-prem setups.

On every release, upload the manifest with the **bumped version**. Users may need to restart Word.

---

## 10. Quick checklist for a new environment

- [ ] `npm install` in root and in `backend/`
- [ ] `backend/.env` created with service account, `TRUSTED_SP_HOSTS`, SMTP values
- [ ] Decide hosting URL; update `urlProd` (webpack) and every URL in `manifest.xml`
- [ ] `npm run validate` passes
- [ ] `relayBaseUrl` is `""`
- [ ] `npm run build` succeeds; `dist/manifest.xml` has the right URLs
- [ ] IIS prerequisites installed; site bound and (ideally) HTTPS
- [ ] `dist/` contents + `backend/` (with `node_modules`, `.env`) copied to site root
- [ ] `/taskpane.html` renders; `/api/health` returns `ok`
- [ ] Manifest deployed to a test user; open a SharePoint document in Word; type `@name` in a
      comment; confirm the email arrives

## 11. Security notes

- The service-account password and SMTP password live only in `backend/.env`; restrict NTLM
  service account to Read-only.
- Keep `TRUSTED_SP_HOSTS` limited to your real SharePoint servers.
- `/api/send-email` is unauthenticated and can send mail to any address with any content if
  reached; restrict the site to your internal network/users or add authentication before exposing
  it publicly.
- `tls.rejectUnauthorized:false` on SMTP disables certificate validation; remove it if your mail
  server presents a valid certificate.
