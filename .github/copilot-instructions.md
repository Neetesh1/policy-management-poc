# Copilot Instructions — ONLYOFFICE Node.js POC

## Project Overview
This is a Node.js/Express + ONLYOFFICE Document Server POC for uploading, editing, and versioning `.docx` files.

## Tech Stack
- **Backend**: Node.js 18, Express 4, `multer`, `jsonwebtoken`, `axios`, `uuid`
- **Frontend**: Vanilla HTML/JS + Bootstrap 5 (CDN)
- **Editor**: ONLYOFFICE Document Server (Docker image `onlyoffice/documentserver`)
- **Storage**: Local filesystem under `data/documents/{docId}/versions/`
- **DB**: `data/db.json` (flat JSON array of document records)
- **Container**: Docker + Docker Compose

## Key Files
| File | Purpose |
|---|---|
| `server.js` | All Express routes and business logic |
| `public/index.html` | Document Manager UI (upload + list + version history) |
| `public/editor.html` | ONLYOFFICE editor page |
| `docker-compose.yml` | Runs Node.js app + ONLYOFFICE together |
| `Dockerfile` | Node.js app container |
| `data/db.json` | Auto-created at startup; persists document metadata |

## API Endpoints
- `POST /upload` — upload `.docx` (multer, 50 MB limit)
- `GET /api/documents` — list all documents
- `GET /editor/:documentId` — serves editor.html
- `GET /api/editor-config/:documentId` — returns JWT-signed ONLYOFFICE config
- `GET /file/:documentId?token=<jwt>` — serve file to ONLYOFFICE server
- `POST /onlyoffice/callback/:documentId` — ONLYOFFICE save callback (status 2 / 6)
- `GET /versions/:documentId` — version history + audit log
- `GET /download/:documentId` — download latest version
- `GET /download-version/:documentId/:versionNo` — download specific version
- `GET /api/config` — returns `{ onlyofficeServerUrl }` for frontend

## Conventions
- All document IDs are UUIDv4; validated with `uuidValidate()` before use in file paths
- File paths are checked with `safePath()` (defense-in-depth against path traversal)
- JWT secret is shared between Node.js app and ONLYOFFICE via `JWT_SECRET` env var
- `ONLYOFFICE_APP_URL` is the internal URL ONLYOFFICE uses to reach the Node.js app (container name in Docker)
- `ONLYOFFICE_SERVER_URL` is the browser-facing URL to load the editor API script
- `DATA_PATH` env var overrides the default `<project>/data` directory (used in tests)
- Version numbering: v1 = original upload; v2+ = subsequent saves

## Docker Network Notes
- In Docker Compose, `ONLYOFFICE_APP_URL=http://app:3000` (service name)
- In local dev (no Docker), `ONLYOFFICE_APP_URL=http://localhost:3000`
- ONLYOFFICE must be able to reach the callback and file endpoints from within Docker

## Out of Scope (POC)
- RBAC / authentication
- Production security hardening
- S3 / Azure Blob storage
- SharePoint integration
- AI/translation features
