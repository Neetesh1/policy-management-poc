# ONLYOFFICE + Node.js POC

A Proof of Concept web application for uploading, editing, and versioning `.docx` documents using **ONLYOFFICE Document Server** embedded in a **Node.js / Express** backend.

---

## Features

| Feature | Details |
|---|---|
| Upload `.docx` | Validates type and size (max 50 MB) |
| Browser editing | ONLYOFFICE editor embedded via Docs API |
| Callback save | ONLYOFFICE calls back; new version is persisted |
| Version history | Every save creates a new immutable version |
| Download | Latest or any specific version |
| Audit log | Every upload and save action is recorded |
| JWT security | Files are served with short-lived JWT tokens |

---

## Architecture

```
Browser
  │  ← loads UI from Node.js
  │  ← loads ONLYOFFICE editor script from ONLYOFFICE server (port 8080)
  ▼
Node.js / Express  (port 3000)
  │  ← stores .docx files under  data/documents/{docId}/versions/
  │  ← stores metadata in         data/db.json
  ▼
ONLYOFFICE Document Server  (Docker, port 8080)
  │  → fetches document file    GET  /file/:docId?token=<jwt>
  │  → sends save callback      POST /onlyoffice/callback/:docId
```

---

## Quick Start

### Prerequisites

- [Docker Desktop](https://www.docker.com/products/docker-desktop/) (Docker + Compose)

### 1 — Clone and configure

```bash
git clone <repo-url>
cd onlyoffice-node-poc
cp .env.example .env          # review defaults; change JWT_SECRET
```

Place your ONLYOFFICE Docs Developer Edition trial license as `license.lic` in the
project root — `docker-compose.yml` mounts it read-only into the Document Server
container at `/var/www/onlyoffice/Data/license.lic` (per the
[activation guide](https://helpcenter.onlyoffice.com/docs/installation/docs-developer-activation.aspx)).
It is git-ignored and never committed.

### 2 — Start all services

```bash
docker compose up --build
```

| Service | URL |
|---|---|
| Document Manager (Node.js) | http://localhost:3000 |
| ONLYOFFICE Document Server | http://localhost:8080 |

### 3 — Use the app

1. Open http://localhost:3000
2. Upload a `.docx` file
3. Click **Edit** — the ONLYOFFICE editor opens in a new tab
4. Make changes and close the editor (ONLYOFFICE triggers an auto-save callback)
5. The document list shows the incremented version
6. Click **History** to view all versions and the audit log

---

## Deploy to Firebase

Yes, this POC can be deployed with Firebase by using **Firebase Hosting + Cloud Run**:

- `onlyoffice-app` (Node.js app) on Cloud Run
- `onlyoffice-docs` (ONLYOFFICE Document Server) on Cloud Run
- Firebase Hosting rewrite to `onlyoffice-app`

Use the full deployment guide in [docs/firebase-deployment.md](docs/firebase-deployment.md).

Included helper files:

- `firebase.json`
- `.firebaserc.example`

### Deploy From GitHub (recommended)

If you do not want local `gcloud` / `firebase` CLIs, use GitHub Actions deployment.

- Workflow: `.github/workflows/deploy-firebase.yml`
- Setup guide: [docs/github-actions-deploy.md](docs/github-actions-deploy.md)

---

## Local Development (without Docker)

1. Install and run [ONLYOFFICE Document Server](https://helpcenter.onlyoffice.com/installation/docs-community-install-ubuntu.aspx) on port 8080.
2. Install dependencies:
   ```bash
   npm install
   ```
3. Copy and edit the environment file:
   ```bash
   cp .env.example .env
   # Set ONLYOFFICE_APP_URL=http://localhost:3000 (same as APP_URL for local dev)
   ```
4. Start the server:
   ```bash
   npm run dev      # with auto-reload (nodemon)
   # or
   npm start        # production mode
   ```

---

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `PORT` | `3000` | Port the Node.js app listens on |
| `APP_URL` | `http://localhost:3000` | Public URL browsers use to reach the app |
| `ONLYOFFICE_APP_URL` | `http://app:3000` | Internal URL ONLYOFFICE server uses to fetch files / send callbacks (use container name in Docker) |
| `ONLYOFFICE_SERVER_URL` | `http://localhost:8080` | URL browser uses to load the ONLYOFFICE editor API script |
| `JWT_SECRET` | `poc-secret-change-in-production` | Shared JWT secret — **must match** the ONLYOFFICE `JWT_SECRET` env var |
| `DATA_PATH` | `<project>/data` | Override data directory (DB + document storage) |

---

## API Reference

| Method | Path | Description |
|---|---|---|
| `GET` | `/` | Document Manager UI |
| `POST` | `/upload` | Upload a `.docx` file |
| `GET` | `/api/documents` | List all documents |
| `GET` | `/editor/:documentId` | ONLYOFFICE editor page |
| `GET` | `/api/editor-config/:documentId` | Signed ONLYOFFICE config (JWT) |
| `GET` | `/file/:documentId?token=<jwt>` | Serve document file to ONLYOFFICE |
| `POST` | `/onlyoffice/callback/:documentId` | ONLYOFFICE save callback |
| `GET` | `/versions/:documentId` | Version history + audit log |
| `GET` | `/download/:documentId` | Download latest version |
| `GET` | `/download-version/:documentId/:versionNo` | Download a specific version |
| `GET` | `/api/config` | Frontend configuration (ONLYOFFICE URL) |
| `GET` | `/automation/policy-info-panel.html` | Connector-window panel opened from the "Policy Tools" toolbar tab (Automation API) |
| `POST` | `/api/audit-event/:documentId` | Records a client-reported Automation API event (tag change, comment lifecycle, review accept/reject) into the audit trail |

---

## White-label branding vs. Automation API

Two different premium ONLYOFFICE Docs Developer features are wired into this POC — they are **not** interchangeable:

| | White-label branding | Automation API |
|---|---|---|
| Config surface | `editorConfig.customization` (`about`, `loaderName`, `loaderLogo`, `layout.*`) | `docEditor.createConnector()` client-side |
| What it does | Toggle/hide native ONLYOFFICE chrome (logo, loader text, toolbar tabs) | Add **new** custom UI (toolbar tabs, context menu items, modal windows) and read/write document content from outside the editor |
| License required | Extended white-label license | Automation API license (separate add-on) |
| Behavior without the license | Silently ignored — no error | `createConnector()`/`addToolbarMenuItem()` may throw or no-op; wrapped in try/catch with a console warning |
| Implemented here | `server.js` sets `about`, `loaderName`, optional `loaderLogo` via env vars | `public/editor.html` adds a "Policy Tools" toolbar tab that opens `public/automation/policy-info-panel.html` in a connector window |

**Automation API cannot hide/replace native toolbar tabs or the loading logo** — that part still requires the white-label license. If neither license is present on your trial server, check the browser console for `[Automation API]` warnings to confirm which parts are actually active.

### Compact toolbar & hiding toolbar tabs

`server.js` sets two more `customization` fields on top of the above:

- `compactToolbar: true` — **standard branding**, works on every edition (no license needed). Collapses the ribbon to just the tab row; clicking a tab expands its buttons for that session, matching the "Home only by default, expand on demand" behavior.
- `layout.toolbar.{plugins,protect,view}: false` — hides the Plugins/Protection/View tabs entirely. This is **white-label** config (same license as above) — it no-ops on an unlicensed or expired-license server. `home` cannot be hidden per the ONLYOFFICE docs. There is no standard "AI" tab in `layout.toolbar` because it isn't native chrome — it's a bundled **plugin** (see below).

The visible **AI** tab comes from ONLYOFFICE's built-in AI plugin, bundled since Docs 9.0.4 (guid `{9DC93CDB-B576-4F0C-B55E-FCC9C48DD007}`). It's disabled via `editorConfig.plugins.disable`, which requires **no license** — it fully blocks the plugin rather than just hiding its UI.

### Revision tracking & paragraph tagging via the connector

Building on the [Connector class](https://api.onlyoffice.com/docs/docs-api/usage-api/automation-api/connector-class/) and [Document API Events](https://api.onlyoffice.com/docs/plugins/interacting-with-editors/document-api/Events/) docs, `public/editor.html` wires the connector to:

- `attachEvent('onChangeContentControl', ...)` — detects edits to tagged paragraphs (content controls) in real time and reports `CONTENT_CONTROL_CHANGED` to `/api/audit-event/:documentId`.
- `attachEvent('onAddComment' | 'onChangeCommentData' | 'onRemoveComment', ...)` — logs the full comment lifecycle live, instead of only on document save.
- "Accept All Changes" / "Reject All Changes" toolbar buttons (editor/reviewer roles only) — call `executeMethod('AcceptReviewChanges' | 'RejectReviewChanges', [])` and log `REVIEW_CHANGES_ACCEPTED`/`REVIEW_CHANGES_REJECTED`.

**Known limitation**: there is no Document API event that fires per individual tracked-change edit (author/timestamp/diff text), only comment events and bulk accept/reject/navigate methods (`AcceptReviewChanges`, `RejectReviewChanges`, `MoveToNextReviewChange`). A structured, change-by-change revision feed ("who changed exactly what text, when") is **not** exposed by this API surface and would still require diffing saved versions server-side.

---

## Data Layout

```
data/
├── db.json                        ← JSON document registry
└── documents/
    └── {uuid}/
        └── versions/
            ├── v1.docx            ← original upload
            ├── v2.docx            ← first edit save
            └── v3.docx            ← second edit save
```

### Document record in `db.json`

```json
{
  "id": "uuid",
  "title": "example.docx",
  "createdAt": "ISO-date",
  "updatedAt": "ISO-date",
  "currentVersion": 2,
  "versions": [
    { "versionNo": 1, "filePath": "...", "fileHash": "sha256", "createdBy": "user", "createdAt": "..." },
    { "versionNo": 2, "filePath": "...", "fileHash": "sha256", "createdBy": "user", "createdAt": "..." }
  ],
  "audit": [
    { "action": "DOCUMENT_UPLOADED", "versionNo": 1, "userId": "user", "timestamp": "..." },
    { "action": "DOCUMENT_SAVED",    "versionNo": 2, "userId": "user", "timestamp": "..." }
  ]
}
```

---

## ONLYOFFICE Callback Flow

```
User closes editor / auto-save triggers
        │
        ▼
ONLYOFFICE Document Server
  POST /onlyoffice/callback/:documentId
  Body: { status: 2, url: "<download-url>", users: [...] }
        │
        ▼
Node.js downloads updated file from ONLYOFFICE
Saves as v{n+1}.docx  →  updates db.json
```

ONLYOFFICE status codes handled:

| Status | Meaning |
|---|---|
| `1` | Document is being edited (no action) |
| `2` | All editors closed — ready to save |
| `6` | Force-saved with error — save anyway |

---

## Known Limitations (POC)

- No authentication or RBAC
- Single-user ownership (all files attributed to `user-001`)
- `db.json` is not safe for concurrent writes at high load
- Local filesystem storage only (not S3 / Azure Blob)
- No production TLS / secure headers

## Future Enhancements

- Role-based access (Editor / Reviewer / Approver)
- Approval workflow
- S3 / Azure Blob storage backend
- PostgreSQL metadata store
- SharePoint integration
- Version diff / comparison UI
- AI translation pipeline
