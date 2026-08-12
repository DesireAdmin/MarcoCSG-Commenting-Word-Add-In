# MentionNotifier — Setup, Configuration & Deployment Guide

This guide is for the developer who will take over this codebase. It covers everything needed
to install, configure, build, deploy, and **debug** the add-in without help from the original
development team.

---

## 1. What this project is

**MentionNotifier** is a Microsoft Word task-pane add-in that adds `@mention` support to native
Word comments. When someone types `@username` in a comment and that username matches a real
SharePoint site user, the add-in emails that person a notification with a link back to the
document and a reference code that jumps straight to the comment.

The project has **two independent halves** that are built/run/deployed separately:

| Part | Location | What it is | Runs as |
|---|---|---|---|
| **Add-in front end** | project root (`src/`, `manifest.xml`, `webpack.config.js`) | The task-pane UI loaded inside Word (HTML/CSS/JS, bundled by Webpack) | Static files served by IIS (or the Webpack dev server while developing) |
| **Backend relay** | `backend/` | A small Node/Express API that does the two things the browser can't do directly: NTLM-authenticate to SharePoint to fetch site users, and send the notification emails over SMTP | Node process under **IISNode**, or standalone with `node server.js` |

Why a backend exists at all: the browser sandbox that hosts the task pane cannot do NTLM
authentication against SharePoint, and SharePoint's own outgoing-email utility has TLS
compatibility problems talking to Gmail's SMTP. The backend is a thin relay that solves both
problems server-side using one shared service account.

---

## 2. Prerequisites

Install these on the machine that will build/develop/debug the code, and on whichever machine
will host it:

| Requirement | Notes |
|---|---|
| **Node.js 18.x or 20.x LTS** (with npm) | Used to install dependencies, run Webpack, and run the backend. Download from https://nodejs.org. Run `node -v` and `npm -v` to confirm. |
| **Git** | To clone/manage the source (optional if you only received a zip). |
| **Windows Server with IIS** | Target hosting environment. Needs the **URL Rewrite** module and **IISNode** installed (see §7). |
| **Microsoft Word (desktop, Microsoft 365, Windows)** | For sideloading and testing/debugging the add-in. The task pane runs inside Word's built-in **Edge WebView2** runtime — on a normal, up-to-date Windows 10/11 + Microsoft 365 machine this is already installed alongside Office, nothing extra to add. |
| A code editor | VS Code recommended — this repo already ships `.vscode/launch.json` and `.vscode/tasks.json` for one-click sideload/debug (see §8.3). |
| **SharePoint on-premises farm access** | You need a **service account** with at least Read permission on the site collections where documents live (ideally granted once via a Web Application User Policy in Central Admin, rather than per-site). |
| **SMTP credentials** | An account/relay that can send outbound mail (the current build uses a Gmail account with an app password; any standard SMTP server works). |

### About build tooling — no Gulp / no Heft needed

If you've worked with **SharePoint Framework (SPFx)** projects before, you may expect to see
`gulp` or `heft` here — **this project does not use either.** SPFx web parts are a different
technology from Office JS task-pane add-ins. This project is a plain Yeoman-style Office Add-in
(`office-addin-taskpane-js` template), and its **only** build tool is **Webpack** (bundling) with
**Babel** (transpiling), both driven purely through the `npm run ...` scripts in `package.json` —
see §4 and §6. There is nothing else to install for building beyond Node.js/npm itself; Webpack,
Babel, and the `office-addin-*` CLI tools (debugging, dev-certs, manifest validation) are all
already listed as `devDependencies` and get installed locally into `node_modules` by
`npm install` — no global installs required.

---

## 3. What's in the zip / repo layout

```
MentionNotifier/
├─ src/
│  ├─ taskpane/
│  │  ├─ taskpane.html      UI markup for the task pane
│  │  ├─ taskpane.css       Styling
│  │  └─ taskpane.js        All add-in logic (see §3.1)
│  └─ commands/
│     ├─ commands.html      Required stub page for ribbon-button commands
│     └─ commands.js        Required stub script (Office add-in boilerplate, not used for logic)
├─ assets/                  Icons used by the manifest (16/32/64/80/128 px, etc.)
├─ manifest.xml             The add-in manifest — this is what you deploy to end users
├─ webpack.config.js        Build configuration (bundles src/ into dist/)
├─ web.config               IIS config — copied into dist/ on every build (see §7)
├─ package.json             Front-end npm scripts/dependencies
├─ backend/
│  ├─ server.js             The Express relay API (SharePoint NTLM + SMTP email)
│  ├─ package.json          Backend npm dependencies
│  └─ .env                  Backend secrets/config, shipped with placeholder values — replace
│                            with your real SharePoint/SMTP values before running (see §5)
├─ dist/                    Build OUTPUT (created by `npm run build`) — this is what actually gets deployed
└─ USER_GUIDE.md            This file
```

> **Note on what was excluded from the shared zip:** `node_modules` (both root and `backend/`)
> and the previously built `dist/` folder were removed, since both regenerate automatically
> (`npm install` and `npm run build` respectively — see §4 and §7.3). `backend/.env` **is**
> included, but with its real SharePoint/SMTP values replaced by placeholders — see §5 for what
> each placeholder means and what to put in its place.

### 3.1 `src/taskpane/taskpane.js` — what it actually does

This is the only file with real application logic. Roughly, top to bottom:

- **`CONFIG.relayBaseUrl`** — base URL the browser uses to reach the backend (`""` = same origin,
  i.e. assume the backend is reachable at `/api/...` on the same host that serves the task pane).
  This is the one setting you may need to change per environment — see §6.3.
- **`preloadSiteUsers()`** — on load, calls the backend's `/api/siteusers` to fetch the SharePoint
  site's user list (name/email/login) for whichever site the open document belongs to.
- **`initAutocomplete()`** — wires up the `@name` autocomplete dropdown in the comment input box.
- **`scanAllComments()`** — runs every `CONFIG.scanInterval` ms (default 8000) via `setInterval`.
  Reads all native Word comments/replies through the Office.js Comments API, and for any comment
  containing a `@username` that matches a known SharePoint user, tags it with a hidden
  `[ref:XXXXXX]` code and fires an email (once per user per comment — tracked in-memory via
  `notifiedMentions`).
- **`createNativeCommentFromTaskpane` / `updateNativeCommentInWord` / `createReplyToCommentInWord`
  / `deleteNativeCommentFromWord`** — the task pane's own mini comment-management UI, backed
  entirely by Word's native Comments API (so comments created here are ordinary Word comments,
  visible in Word itself, not something proprietary).
- **`sendNotificationSandbox()` / `sendEmailViaRelay()`** — builds the HTML email and POSTs it to
  the backend's `/api/send-email`.

You should not need to touch this file for a standard deployment — everything environment-specific
is isolated to `CONFIG.relayBaseUrl` (see §6.3).

### 3.2 `backend/server.js` — what it actually does

- Loads config from `backend/.env` via `dotenv`. Refuses to start if `SP_DOMAIN` / `SP_USERNAME`
  / `SP_PASSWORD` / `TRUSTED_SP_HOSTS` are missing (check the log if it exits immediately).
- `GET /api/health` — quick liveness check, also echoes back the configured trusted hosts.
- `GET /api/siteusers?docUrl=...` — validates that `docUrl`'s hostname is in `TRUSTED_SP_HOSTS`
  (this allowlist exists specifically to prevent the relay being pointed at an arbitrary server —
  otherwise it's an SSRF hole, and worse with NTLM since a rogue server could capture the
  credential handshake). It then walks up the URL path segment by segment, calling SharePoint's
  `_api/web/siteusers` REST endpoint at each level with the shared service account (NTLM), until
  one resolves to an actual site.
- `POST /api/send-email` — sends the notification email via `nodemailer`/SMTP using the configured
  account.

---

## 4. Installing dependencies

There are **two separate `node_modules` folders** to install — one for the add-in front end
(project root), one for the backend relay (`backend/`). `npm install` must be run in **both**
places.

From the project root:

```bash
npm install
```

Then into the backend folder:

```bash
cd backend
npm install
cd ..
```

This needs to be done once initially, and again any time `package.json` (root or backend)
changes.

---

## 5. Configuring the backend (`backend/.env`)

`backend/.env` ships in the zip with placeholder values in place of our real credentials. Open it
and replace each placeholder with your own values:

```ini
# Port for standalone/local testing only. Under IISNode in production, IIS
# supplies the port via a named pipe automatically and this value is ignored.
PORT=5000

# --- SharePoint service account (used for NTLM calls in /api/siteusers) ---
# This account needs Read access wherever documents actually live. For farms
# with many site collections, granting it once via a Web Application User
# Policy in Central Admin (covers the whole web app, current and future
# sites) is much less maintenance than doing it site by site.
SP_DOMAIN=YOUR_DOMAIN
SP_USERNAME=svc-your-service-account
SP_PASSWORD=your-service-account-password

# --- Trusted SharePoint front-end server hostnames ---
# Comma-separated list of the actual SharePoint web-front-end server names.
# /api/siteusers checks a document's URL host against this list before doing
# anything with it - this is a security control, not a convenience setting.
# List server hostnames only (not site URLs, not web app URLs). Add an entry
# only when documents start being opened from an additional physical server.
TRUSTED_SP_HOSTS=your-sp-server,your-sp-server.yourdomain.local

# --- Outgoing SMTP for notification emails ---
SMTP_HOST=smtp.yourprovider.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your-smtp-username
SMTP_PASS=your-smtp-password-or-app-password
SMTP_FROM=notifications@yourdomain.com
SMTP_TIMEOUT=30000
```

Notes:

- If your SMTP relay allows anonymous submission, you can leave `SMTP_USER`/`SMTP_PASS` blank —
  the backend only attaches SMTP auth if both are set.
- `SMTP_SECURE=true` is only needed for port 465 (implicit TLS); for the common 587/STARTTLS case
  leave it `false`.
- **Never commit this file to git with real values filled in** — `backend/.gitignore` already
  excludes `.env` from version control. Treat the service-account password and SMTP password as
  real secrets.
- This is the **only** file where secrets live. No credentials belong in `taskpane.js`,
  `manifest.xml`, or anywhere under `src/`.

---

## 6. Configuring the front end for your environment

You will be hosting the built add-in files (from `dist/`) on your own IIS server under your own
hostname/port. A few things need to be updated so the add-in points at *your* server instead of
ours. Do this **before** running a production build.

### 6.1 `manifest.xml`

This is a plain XML file — open it in any text/XML editor. Replace every occurrence of our
placeholder host (currently `spse01`, ports `8088`/`8080`) with your actual server's hostname and
the port IIS will serve the add-in on:

| Element | Purpose |
|---|---|
| `<IconUrl>`, `<HighResolutionIconUrl>`, `<SupportUrl>` | Where the icon/help assets are served from |
| `<AppDomains><AppDomain>` | Domain(s) the add-in is allowed to navigate to |
| `<DefaultSettings><SourceLocation>` | URL of `taskpane.html` |
| `<bt:Image>` entries under `<Resources><bt:Images>` | Icon URLs again, for each size |
| `<bt:Url>` entries under `<Resources><bt:Urls>` | `commands.html` and `taskpane.html` URLs, plus the "Learn more" link |

Also update, once, for branding:
- `<ProviderName>` — your company name (currently `DesireInfoweb`)
- `<Id>` — a GUID that uniquely identifies this add-in. **Leave this as-is** unless you are
  intentionally publishing this as a brand-new, unrelated add-in — changing it means Word treats
  it as a different add-in entirely (existing installs won't update, they'll duplicate).

Everything in this file must use the **same** hostname/port your IIS site is actually configured
on (see §7) — a mismatch here is the most common cause of the task pane failing to load.

> This is now the **only** manifest in the project — the separate localhost/ngrok development
> manifests that used to ship alongside it have been removed as unnecessary. If you ever want a
> faster local dev loop against the Webpack dev server (`https://localhost:3000`), see the tip in
> §8.1 — you don't need a second manifest file for that, just a temporary local edit.

### 6.2 `webpack.config.js`

```js
const urlDev = "https://localhost:3000/";
const urlProd = "http://spse01:8080"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION
```

Update `urlProd` to match your IIS site's actual base URL. This value is only used as a
find/replace target: during a **production** build (`npm run build`), the build copies
`manifest.xml` into `dist/` and replaces any literal `https://localhost:3000/` text with
`urlProd`. Since `manifest.xml` doesn't currently contain `localhost:3000` (it already hardcodes
real server URLs, per §6.1), this has no visible effect today — it only matters if you later start
pointing `manifest.xml` at localhost for local testing (§8.1) and then forget to change it back
before a production build. Keep it accurate regardless, as a safety net.

### 6.3 `src/taskpane/taskpane.js` — `CONFIG.relayBaseUrl`

```js
const CONFIG = {
  scanInterval: 8000,
  relayBaseUrl: "",
};
```

- Leave this as `""` for production. It means "call `/api/...` on whatever host served this page",
  which is correct once both the add-in files and the backend are deployed together under the
  same IIS site (see §7 — that's exactly how `web.config`'s rewrite rule is set up).
- Only change this if you deliberately run the backend on a **different host/port** from the
  front end (e.g. temporarily during local development, pointing a locally-served task pane at a
  standalone backend on `http://localhost:5000` — see §8.1/§8.4). Remember to change it back
  before building for production.

### 6.4 `web.config`

Usually no changes needed. The one thing to verify is this path, which must match where Node.js
is actually installed on your IIS server:

```xml
<iisnode nodeProcessCommandLine="&quot;C:\Program Files\nodejs\node.exe&quot;" ... />
```

---

## 7. Deploying to IIS

### 7.1 One-time IIS server setup

Install, in this order, on the Windows Server running IIS:

1. **Node.js LTS** — must be installed on the server itself (not just your dev machine), at the
   path referenced in `web.config`.
2. **IIS URL Rewrite Module** — https://www.iis.net/downloads/microsoft/url-rewrite
3. **IISNode** — https://github.com/Azure/iisnode (lets IIS host a Node.js process as a site)
4. Create an IIS site (or app) pointing its physical path at the folder you will deploy `dist/`
   into, bound to the hostname/port you used in `manifest.xml` (§6.1).

### 7.2 What gets deployed

The **`dist/`** folder is the deployable unit for the front end — never deploy `src/` directly,
it isn't bundled/minified and browsers won't run it as-is.

Deploy this layout to the IIS site's physical folder:

```
<IIS site root>/
├─ (everything from dist/: taskpane.html, taskpane.js, commands.html, commands.js,
│   polyfill.js, assets/, manifest.xml, web.config, ...)
└─ backend/
   ├─ server.js
   ├─ package.json
   ├─ node_modules/     ← run `npm install` here, on the server, or copy it over
   └─ .env               ← your real values filled in (§5); never leave placeholders in production
```

- `web.config` **must** end up at the site root (not inside `backend/`) — `dist/`'s build already
  copies it there for you. IIS only reads routing/iisnode config starting from the site root, so a
  copy nested in `backend/` would not intercept `/api/*` requests.
- The `backend/` folder is **not** part of the Webpack build output — copy it into the deployed
  site alongside `dist/`'s contents manually (or via your own deploy script), keeping the same
  relative path (`backend/server.js` next to `taskpane.html`).
- With `web.config`'s rewrite rule in place, requests to `/api/*` on the site are transparently
  routed to `backend/server.js` running under IISNode — that's what makes `relayBaseUrl: ""` work
  in production.

### 7.3 Build & deploy steps

```bash
# 1. Install/refresh dependencies (see §4) if you haven't already
npm install
cd backend && npm install && cd ..

# 2. Production build - bundles src/ into dist/, copies assets/manifest/web.config
npm run build
```

`npm run build` runs `webpack --mode production`. Confirm it finished with no errors, then:

3. Copy the contents of `dist/` to the IIS site's physical folder.
4. Copy the `backend/` folder (including its own `node_modules` and your filled-in `.env`) into
   that same physical folder, as a subfolder named `backend`.
5. Recycle the IIS application pool (or `iisreset`) so IISNode picks up the new backend code.
6. Browse to `http://<your-host>:<port>/taskpane.html` directly in a browser first — you should
   see the task pane UI render (even outside Word) and no 404/500. This confirms IIS is serving
   the static files correctly before you even involve Word.
7. Hit `http://<your-host>:<port>/api/health` — should return
   `{"status":"ok","trustedHosts":[...]}`. If this 404s, the IIS URL Rewrite rule or IISNode setup
   is the problem, not the add-in.
8. Distribute `manifest.xml` to end users through your normal Office add-in deployment method —
   Centralized Deployment via the Microsoft 365 admin center, a network share catalog, or
   SharePoint app catalog, depending on how your organization deploys Word add-ins. This is a
   Microsoft 365 admin task, not something this project controls.

There is no `npm run build:dev` deployment path — `build:dev` is strictly for local
testing/debugging without minification (see §8).

---

## 8. Local development & debugging workflow

### 8.1 Two ways to run this locally

**Option A — against your real/test IIS server (simplest, matches production):**
Since `manifest.xml` already points at a real hostname, the most friction-free loop is:
run `npm run watch` (rebuilds `dist/` automatically on every save, unminified with source maps),
copy/sync `dist/` to your test IIS site as you go, and debug against that (§8.2). No manifest
edits needed at all.

**Option B — against the Webpack dev server on `https://localhost:3000` (faster inner loop, no
IIS redeploy per change):**

```bash
npm run dev-server
```

This starts Webpack's dev server on `https://localhost:3000` with a self-signed cert
auto-provisioned by `office-addin-dev-certs` (accept/trust it the first time Windows prompts).
To sideload against it you need a manifest whose `SourceLocation`/`AppDomains`/icon URLs point at
`https://localhost:3000` instead of your real server. Since the dedicated `manifest.local.xml`
has been removed, just **make your own local, uncommitted copy** for this — e.g. duplicate
`manifest.xml` to something like `manifest.dev.xml`, point its URLs at `https://localhost:3000`,
sideload *that* copy (§8.2), and never deploy or hand this file to anyone — it's a personal dev
convenience, not part of the shipped project.

If you're also testing backend changes at the same time, run the backend separately
(`cd backend && node server.js`) and temporarily set `CONFIG.relayBaseUrl` in
`src/taskpane/taskpane.js` to `"http://localhost:5000"` (§6.3) — then set it back to `""` before
committing/building for real.

### 8.2 Opening Word and debugging the task pane (`taskpane.js`)

This is the main day-to-day debugging loop, and it works the same way whether the task pane is
being served from your real IIS server, or from `https://localhost:3000` — the mechanism is
identical, you're just pointed at a different manifest/URL.

1. **Build in dev mode** so the bundle stays readable (this matters — a `npm run build` production
   bundle is minified and much harder to set breakpoints in):
   ```bash
   npm run build:dev
   # or, to keep rebuilding automatically as you edit files:
   npm run watch
   ```
2. **Sideload the add-in into Word:**
   - Easiest: `npm start` — this builds, sideloads `manifest.xml`, and launches Word automatically
     (equivalent to VS Code's "Debug: Word Desktop" task, §8.3).
   - Manual alternative (works with any manifest, including a local dev copy from §8.1 Option B):
     open Word → **Insert** tab → **My Add-ins** → the dropdown/gear icon → **Upload My Add-in** →
     browse to your manifest file. The task pane button then appears on the **Home** ribbon tab.
3. **Open the task pane** by clicking its ribbon button ("Show Task Pane").
4. **Open DevTools for the task pane itself** — right-click anywhere *inside the task pane* and
   choose **Inspect**. This opens a separate Edge DevTools window attached specifically to the
   task pane's WebView2 process (not the same as debugging Word itself — Word has no DevTools of
   its own; the task pane, being a web page, does).
   - If **Inspect** doesn't appear on the right-click menu, sideloaded/dev-mode add-ins should show
     it by default in current Microsoft 365 builds. If it's missing, developer tools support may be
     disabled centrally — see Microsoft's "Enable developer tools trust settings" guidance for the
     relevant registry/trust setting, or ask your Microsoft 365 admin.
5. **In the DevTools window:**
   - **Sources tab** — find `taskpane.js` under the page's file tree, click a line number to set a
     breakpoint, then trigger the code path (type a comment with `@username`, wait for the
     8-second `scanAllComments` interval, click a comment card, etc.). Execution pauses right there
     with full variable inspection, call stack, and step controls — exactly like debugging any web
     page.
   - **Console tab** — the code already logs extensively with a `[MentionNotifier]` prefix
     (`preloadSiteUsers`, `scanAllComments`, mention/email failures, etc.) — check here first before
     reaching for breakpoints.
   - **Network tab** — inspect the actual `GET /api/siteusers` and `POST /api/send-email` requests/
     responses. This is the fastest way to tell whether a problem is in the front-end code or the
     backend relay: if the request never fires, it's front-end logic; if it fires and comes back
     with an error status, the problem is in the backend (jump to §8.4/§9).
6. **After editing code:** if you're on `npm run watch`, the bundle rebuilds automatically — just
   reload the task pane (close and reopen it from the ribbon button, or use the DevTools reload)
   to pick up the new code. A one-off `build:dev` needs to be re-run manually before reloading.

### 8.3 Debugging via VS Code (optional, integrated breakpoints)

This repo already includes `.vscode/launch.json` and `.vscode/tasks.json`, which let you set
breakpoints directly in the VS Code editor instead of in a separate DevTools window.

1. Open the **Run and Debug** panel in VS Code.
2. Choose **"Word Desktop (Edge Chromium)"** from the dropdown and press **F5** (or click the
   green run arrow).
3. This runs the **"Debug: Word Desktop"** task behind the scenes (`npm run start -- desktop --app
   word`, i.e. `office-addin-debugging start manifest.xml`), which builds, sideloads, and launches
   Word for you, then attaches the debugger.
4. Set breakpoints directly in `src/taskpane/taskpane.js` in the editor (not the bundled `dist/`
   copy) — VS Code maps them through the source map automatically.
5. Trigger the relevant code path in Word as in §8.2 step 5 — VS Code will pause on your breakpoint
   just like the DevTools Sources tab does.
6. Stop debugging (red square, or Shift+F5) — this runs the **"Stop Debug"** task automatically,
   which unsideloads the add-in.

Note the checked-in `launch.json`'s "Word Desktop (Edge Legacy)" configuration is hardcoded to
attach at `https://localhost:3000/taskpane.html...` — that one only works if `manifest.xml` (or
whatever you sideload) is actually serving from `localhost:3000` (§8.1, Option B). The **"Word
Desktop (Edge Chromium)"** configuration is the one that works regardless of which URL you're
actually sideloading from, so it's the one to prefer day to day.

### 8.4 Debugging the backend (`server.js`) locally

The backend is a plain Node/Express app, so it debugs like any other Node script:

```bash
cd backend
node server.js
```

- Console output (the timestamped `[INFO]`/`[ERROR]` lines) prints straight to that terminal —
  often enough on its own, since `server.js` already logs every request and every SharePoint/SMTP
  outcome.
- For breakpoints: in VS Code, open a **JavaScript Debug Terminal** (Command Palette → "Debug:
  JavaScript Debug Terminal") and run `node server.js` in *that* terminal instead of a plain one —
  VS Code auto-attaches its debugger, so any breakpoint you set in `backend/server.js` in the
  editor will be hit with no extra configuration needed.
- Alternatively, run `node --inspect-brk server.js` and use VS Code's "Attach to Node Process" (or
  open `edge://inspect` / `chrome://inspect` in a Chromium browser) to attach manually.
- To exercise it without going through Word/the task pane at all, hit the endpoints directly, e.g.:
  ```bash
  curl "http://localhost:5000/api/health"
  curl "http://localhost:5000/api/siteusers?docUrl=http://your-sp-server/sites/YourSite/Shared%20Documents/file.docx"
  ```
  This isolates whether an issue is in the NTLM/SharePoint call itself versus something in the
  front end's handling of the response.

---

## 9. Debugging in production / on the client's server

- **Front-end errors** (task pane not loading, JS exceptions): same technique as §8.2 step 4 —
  right-click inside the task pane → **Inspect**. Works identically whether the add-in is
  sideloaded for dev or deployed for real; the only difference is production's `dist/` bundle is
  minified, so stack traces/line numbers are less readable (rebuild with `build:dev` and swap it
  in temporarily if you need to debug something specific to the production deployment).
- **Backend errors**: IISNode writes logs to an `iisnode` subfolder next to `backend/server.js` on
  the server (path/rotation controlled by `web.config`'s `<iisnode>` element — currently capped at
  40 log files, 1 MB each). `server.js` timestamps and labels every log line (`[INFO]`/`[ERROR]`),
  and every request is logged with method/path/status/duration regardless of which route handled
  it — start there.
- **"Missing service account configuration" / "Missing TRUSTED_SP_HOSTS" and the backend exits
  immediately**: `backend/.env` is missing, misplaced, or missing required keys — see §5. Confirm
  the file is actually at `backend/.env` (sibling to `server.js`), not the project root.
- **`/api/siteusers` returns `not_sharepoint`**: The document's URL host isn't in
  `TRUSTED_SP_HOSTS`. Confirm the hostname Word reports for the open document
  (`Office.context.document.url`) matches one of the entries — check for a mismatch like an IP
  address vs. hostname, or a load-balancer/alternate access mapping name that isn't listed.
- **`/api/siteusers` returns `no_access` (403)**: The NTLM service account doesn't have Read
  permission on that specific site. Grant it there, or better, via a Web Application User Policy.
- **`/api/siteusers` returns `not_found` (404)**: No SharePoint site could be resolved for any
  prefix of the document's URL — usually means the document isn't actually on a SharePoint site
  the service account/relay can see, or `TRUSTED_SP_HOSTS` allowed a host that isn't a real
  SharePoint front end.
- **Emails never arrive**: Check the backend log for `[SMTP]` lines — `mailTransporter.verify()`
  logs a clear success/failure message every time the backend starts, which will surface bad SMTP
  credentials or connectivity immediately without needing to trigger a real mention.
- **Task pane loads but shows a generic HTTP error page for a backend failure**: Confirm
  `<httpErrors existingResponse="PassThrough" />` is present in `web.config` at the site root —
  without it, IIS substitutes its own generic HTML error page over the backend's actual JSON error
  body, which is confusing when debugging.
- **`office-addin-manifest validate manifest.xml`** — run this (via `npx office-addin-manifest
  validate manifest.xml`) any time you hand-edit the manifest, to catch XML/schema mistakes before
  deploying.

---

## 10. Quick-reference checklist for a new environment

Before your first production build/deploy in a new environment, confirm you've touched all of:

- [ ] `backend/.env` — placeholders replaced with your real SharePoint service account, trusted
  hosts, and SMTP settings (§5)
- [ ] `manifest.xml` — all URLs point at your IIS hostname/port, `ProviderName` updated (§6.1)
- [ ] `webpack.config.js` — `urlProd` matches your IIS hostname/port (§6.2)
- [ ] `src/taskpane/taskpane.js` — `CONFIG.relayBaseUrl` is `""` for production (§6.3)
- [ ] `web.config` — Node.js path matches your server's install (§6.4)
- [ ] IIS server has URL Rewrite + IISNode installed (§7.1)
- [ ] `npm install` run in both project root and `backend/` (§4)
- [ ] `npm run build` completed with no errors, `dist/` deployed to IIS root, `backend/` (with its
  own `node_modules` and filled-in `.env`) deployed alongside it (§7.3)
- [ ] `/taskpane.html` and `/api/health` both load correctly in a plain browser before sideloading
  into Word (§7.3)
