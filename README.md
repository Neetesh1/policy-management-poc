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
