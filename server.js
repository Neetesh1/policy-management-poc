'use strict';

const express = require('express');
const multer = require('multer');
const jwt = require('jsonwebtoken');
const axios = require('axios');
const { v4: uuidv4, validate: uuidValidate } = require('uuid');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

// ── Configuration ──────────────────────────────────────────────────────────────
const JWT_SECRET = process.env.JWT_SECRET || 'poc-secret-change-in-production';
const APP_URL = process.env.APP_URL || 'http://localhost:3000';
// ONLYOFFICE_APP_URL: internal URL that ONLYOFFICE server uses to reach this app
// (in Docker use container name, e.g. http://app:3000; default = APP_URL for local dev)
const ONLYOFFICE_APP_URL = process.env.ONLYOFFICE_APP_URL || APP_URL;
// ONLYOFFICE_SERVER_URL: URL browser uses to load the ONLYOFFICE editor API script
const ONLYOFFICE_SERVER_URL = process.env.ONLYOFFICE_SERVER_URL || 'http://localhost:8080';
const PORT = parseInt(process.env.PORT || '3000', 10);

const DATA_PATH = path.resolve(process.env.DATA_PATH || path.join(__dirname, 'data'));
const DB_PATH = path.join(DATA_PATH, 'db.json');
const STORAGE_PATH = path.join(DATA_PATH, 'documents');
const TEMP_PATH = path.resolve(process.env.TEMP_PATH || path.join(__dirname, 'temp'));

// ── Directory & DB initialization ──────────────────────────────────────────────
fs.mkdirSync(DATA_PATH, { recursive: true });
fs.mkdirSync(STORAGE_PATH, { recursive: true });
fs.mkdirSync(TEMP_PATH, { recursive: true });
if (!fs.existsSync(DB_PATH)) {
  fs.writeFileSync(DB_PATH, JSON.stringify({ documents: [] }, null, 2));
}

// ── DB helpers ─────────────────────────────────────────────────────────────────
function readDB() {
  return JSON.parse(fs.readFileSync(DB_PATH, 'utf8'));
}

function writeDB(data) {
  fs.writeFileSync(DB_PATH, JSON.stringify(data, null, 2));
}

// ── Path safety helper (defense-in-depth against traversal) ───────────────────
function safePath(base, ...parts) {
  const resolved = path.resolve(base, ...parts);
  const normalizedBase = base.endsWith(path.sep) ? base : base + path.sep;
  if (!resolved.startsWith(normalizedBase) && resolved !== base) {
    throw new Error('Unsafe path detected');
  }
  return resolved;
}

// ── Express app ────────────────────────────────────────────────────────────────
const app = express();
app.use(express.json());

// ONLYOFFICE Document Server fetches plugin assets cross-origin (browser-side),
// so CORS must be allowed explicitly or plugins.js fails to load config.json.
app.use('/plugins', (req, res, next) => {
  res.header('Access-Control-Allow-Origin', ONLYOFFICE_SERVER_URL);
  res.header('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

app.use(express.static(path.join(__dirname, 'public')));

// ── File-type helpers ─────────────────────────────────────────────────────────
const ALLOWED_EXTENSIONS = new Set(['.docx', '.doc', '.pdf']);
const VALID_ROLES = new Set(['editor', 'reviewer', 'commenter']);

// ── POC Users ──────────────────────────────────────────────────────────────────
const POC_USERS = {
  'poc-user-1': { id: 'poc-user-1', name: 'POC User 1' },
  'poc-user-2': { id: 'poc-user-2', name: 'POC User 2' },
  'poc-user-3': { id: 'poc-user-3', name: 'POC User 3' }
};
const DEFAULT_USER = POC_USERS['poc-user-1'];

/**
 * Maps a role name to ONLYOFFICE permissions + mode.
 *
 * editor    — full editing (default)
 * reviewer  — tracked changes only; changes are highlighted for the editor to accept/reject
 * commenter — add comments only; cannot change document content
 */
function getRoleConfig(role) {
  switch (role) {
    case 'reviewer':
      return {
        mode: 'edit',
        permissions: { edit: false, review: true, comment: true, download: true, print: true, fillForms: false }
      };
    case 'commenter':
      return {
        mode: 'edit',
        permissions: { edit: false, review: false, comment: true, download: true, print: true, fillForms: false }
      };
    default: // editor
      return {
        mode: 'edit',
        permissions: { edit: true, review: true, comment: true, download: true, print: true, fillForms: true }
      };
  }
}

function getDocMeta(ext) {
  switch (ext) {
    case '.pdf':
      return { fileType: 'pdf', documentType: 'pdf', mimeType: 'application/pdf' };
    case '.doc':
      return { fileType: 'doc', documentType: 'word', mimeType: 'application/msword' };
    default: // .docx
      return {
        fileType: 'docx',
        documentType: 'word',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      };
  }
}

// ── Multer setup ───────────────────────────────────────────────────────────────
const upload = multer({
  dest: TEMP_PATH,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50 MB
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (!ALLOWED_EXTENSIONS.has(ext)) {
      return cb(
        Object.assign(
          new Error('Only .docx, .doc and .pdf files are allowed'),
          { code: 'INVALID_TYPE' }
        )
      );
    }
    cb(null, true);
  }
});

// ── Middleware ─────────────────────────────────────────────────────────────────
function validateDocId(req, res, next) {
  if (!uuidValidate(req.params.documentId)) {
    return res.status(400).json({ error: 'Invalid document ID' });
  }
  next();
}

// ── Routes ─────────────────────────────────────────────────────────────────────

/**
 * GET /api/config
 * Returns app configuration consumed by the frontend (e.g. ONLYOFFICE server URL).
 */
app.get('/api/config', (_req, res) => {
  res.json({ onlyofficeServerUrl: ONLYOFFICE_SERVER_URL });
});

/**
 * GET /api/documents
 * Returns the list of all uploaded documents.
 */
app.get('/api/documents', (_req, res) => {
  const db = readDB();
  res.json(
    db.documents.map(d => ({
      id: d.id,
      title: d.title,
      currentVersion: d.currentVersion,
      createdAt: d.createdAt,
      updatedAt: d.updatedAt
    }))
  );
});

/**
 * POST /upload
 * Upload a .docx file (max 50 MB). Stores as v1 and creates a DB entry.
 */
app.post('/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file provided' });
  }

  try {
    const docId = uuidv4();
    const ext = path.extname(req.file.originalname).toLowerCase();
    const docDir = safePath(STORAGE_PATH, docId, 'versions');
    fs.mkdirSync(docDir, { recursive: true });

    const destPath = path.join(docDir, `v1${ext}`);
    fs.copyFileSync(req.file.path, destPath);
    fs.unlinkSync(req.file.path);

    const fileBuffer = fs.readFileSync(destPath);
    const fileHash = crypto.createHash('sha256').update(fileBuffer).digest('hex');
    const now = new Date().toISOString();

    const doc = {
      id: docId,
      title: req.file.originalname,
      extension: ext,
      createdAt: now,
      updatedAt: now,
      currentVersion: 1,
      versions: [
        {
          versionNo: 1,
          filePath: destPath,
          fileHash,
          createdBy: 'user',
          createdAt: now
        }
      ],
      audit: [
        {
          action: 'DOCUMENT_UPLOADED',
          versionNo: 1,
          userId: 'user',
          timestamp: now
        }
      ]
    };

    const db = readDB();
    db.documents.push(doc);
    writeDB(db);

    res.json({ success: true, documentId: docId, title: req.file.originalname });
  } catch (err) {
    // Clean up temp file on failure
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    throw err;
  }
});

/**
 * GET /editor/:documentId
 * Serves the ONLYOFFICE editor HTML page.
 */
app.get('/editor/:documentId', validateDocId, (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'editor.html'));
});

/**
 * GET /api/editor-config/:documentId
 * Returns the signed ONLYOFFICE DocEditor configuration object.
 */
app.get('/api/editor-config/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);

  if (!doc) return res.status(404).json({ error: 'Document not found' });

  // Sanitise & resolve role — defaults to 'editor' for unknown/missing values
  const rawRole = req.query.role;
  const role = VALID_ROLES.has(rawRole) ? rawRole : 'editor';
  const { mode, permissions } = getRoleConfig(role);

  // Resolve POC user — defaults to poc-user-1 for unknown/missing values
  const rawUser = req.query.user;
  const pocUser = POC_USERS[rawUser] || DEFAULT_USER;

  // Short-lived token so ONLYOFFICE can fetch the file
  const fileToken = jwt.sign(
    { documentId, action: 'read', version: doc.currentVersion },
    JWT_SECRET,
    { expiresIn: '2h' }
  );

  const { fileType, documentType } = getDocMeta(doc.extension || '.docx');

  const configPayload = {
    document: {
      fileType,
      title: doc.title,
      url: `${ONLYOFFICE_APP_URL}/file/${documentId}?token=${encodeURIComponent(fileToken)}`,
      // Key must be unique per version; changing it forces ONLYOFFICE to reload the doc
      key: `${documentId.replace(/-/g, '')}_v${doc.currentVersion}`,
      permissions
    },
    documentType,
    editorConfig: {
      mode,
      callbackUrl: `${ONLYOFFICE_APP_URL}/onlyoffice/callback/${documentId}`,
      user: { id: pocUser.id, name: pocUser.name },
      customization: {
        autosave: true,
        forcesave: false,
        // NOTE: hiding toolbar tabs (customization.layout.toolbar.*) requires a commercial
        // ONLYOFFICE branding license — Community Edition ignores it silently, so it's omitted here.
      },
      plugins: {
        // Both plugins must autostart: non-visual context-menu plugins still need their init()
        // to run once to attach the onContextMenuShow listener, otherwise the menu never appears.
        autostart: [
          'asc.{C36DDFB5-08F0-4A68-B829-5FB1F7D49331}',
          'asc.{B2C4D6E8-F0A1-4B3C-9D5E-7F8A9B0C1D2E}'
        ],
        pluginsData: [
          `${APP_URL}/plugins/content-controls-tags/config.json`,
          `${APP_URL}/plugins/policy-context-menu/config.json`
        ]
      }
    }
  };

  // Sign the entire payload so ONLYOFFICE can verify it hasn't been tampered with
  const token = jwt.sign(configPayload, JWT_SECRET);
  res.json({ ...configPayload, token, role, user: pocUser });
});

/**
 * GET /file/:documentId?token=<jwt>
 * Serves the latest version of the document to ONLYOFFICE Document Server.
 */
app.get('/file/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const rawToken = req.query.token;

  if (!rawToken) return res.status(401).json({ error: 'Token required' });

  try {
    jwt.verify(rawToken, JWT_SECRET);
  } catch {
    return res.status(401).json({ error: 'Invalid or expired token' });
  }

  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const latest = doc.versions.find(v => v.versionNo === doc.currentVersion);
  if (!latest) return res.status(404).json({ error: 'File metadata not found' });

  const ext = doc.extension || '.docx';
  const filePath = safePath(STORAGE_PATH, documentId, 'versions', `v${latest.versionNo}${ext}`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found on disk' });

  const { mimeType } = getDocMeta(ext);
  res.setHeader('Content-Type', mimeType);
  res.download(filePath, doc.title);
});

/**
 * POST /onlyoffice/callback/:documentId
 * Receives save callbacks from ONLYOFFICE Document Server.
 * Status 2 = ready to save; Status 6 = force-saved with error.
 */
app.post('/onlyoffice/callback/:documentId', validateDocId, async (req, res) => {
  const { documentId } = req.params;
  const body = req.body;

  // Validate JWT sent by ONLYOFFICE in Authorization header
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    try {
      jwt.verify(authHeader.slice(7), JWT_SECRET);
    } catch {
      console.error(`[callback] JWT verification failed for document ${documentId}`);
      return res.json({ error: 1 });
    }
  }

  console.log(`[callback] documentId=${documentId} status=${body.status}`);

  if (body.status === 2 || body.status === 6) {
    if (!body.url) {
      console.error('[callback] Missing download URL in payload');
      return res.json({ error: 1 });
    }

    try {
      const fileResponse = await axios.get(body.url, {
        responseType: 'arraybuffer',
        timeout: 30000,
        maxContentLength: 50 * 1024 * 1024
      });

      const db = readDB();
      const doc = db.documents.find(d => d.id === documentId);
      if (!doc) return res.json({ error: 1 });

      const newVersionNo = doc.currentVersion + 1;
      const docDir = safePath(STORAGE_PATH, documentId, 'versions');
      fs.mkdirSync(docDir, { recursive: true });

      const ext = doc.extension || '.docx';
      const newFilePath = path.join(docDir, `v${newVersionNo}${ext}`);
      fs.writeFileSync(newFilePath, fileResponse.data);

      const fileHash = crypto.createHash('sha256').update(fileResponse.data).digest('hex');
      const now = new Date().toISOString();
      const userId = (body.users && body.users[0]) || 'user';

      doc.currentVersion = newVersionNo;
      doc.updatedAt = now;
      doc.versions.push({
        versionNo: newVersionNo,
        filePath: newFilePath,
        fileHash,
        createdBy: userId,
        createdAt: now
      });
      doc.audit.push({
        action: 'DOCUMENT_SAVED',
        versionNo: newVersionNo,
        userId,
        timestamp: now
      });

      writeDB(db);
      console.log(`[callback] Saved v${newVersionNo} for document ${documentId}`);
    } catch (err) {
      console.error('[callback] Save failed:', err.message);
      return res.json({ error: 1 });
    }
  }

  res.json({ error: 0 });
});

/**
 * GET /versions/:documentId
 * Returns the full version history and audit log for a document.
 */
app.get('/versions/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  res.json({
    documentId,
    title: doc.title,
    currentVersion: doc.currentVersion,
    versions: doc.versions.map(v => ({
      versionNo: v.versionNo,
      createdAt: v.createdAt,
      createdBy: v.createdBy,
      fileHash: v.fileHash,
      downloadUrl: `/download-version/${documentId}/${v.versionNo}`
    })),
    audit: doc.audit
  });
});

/**
 * GET /download/:documentId
 * Downloads the latest version of a document.
 */
app.get('/download/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const latest = doc.versions.find(v => v.versionNo === doc.currentVersion);
  if (!latest) return res.status(404).json({ error: 'File not found' });

  const ext = doc.extension || '.docx';
  const filePath = safePath(STORAGE_PATH, documentId, 'versions', `v${latest.versionNo}${ext}`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found on disk' });

  res.download(filePath, doc.title);
});

/**
 * GET /download-version/:documentId/:versionNo
 * Downloads a specific version of a document.
 */
app.get('/download-version/:documentId/:versionNo', validateDocId, (req, res) => {
  const { documentId, versionNo } = req.params;
  const versionNumber = parseInt(versionNo, 10);

  if (isNaN(versionNumber) || versionNumber < 1) {
    return res.status(400).json({ error: 'Invalid version number' });
  }

  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const version = doc.versions.find(v => v.versionNo === versionNumber);
  if (!version) return res.status(404).json({ error: 'Version not found' });

  const ext = doc.extension || '.docx';
  const filePath = safePath(STORAGE_PATH, documentId, 'versions', `v${versionNumber}${ext}`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Version file not found on disk' });

  const baseName = path.basename(doc.title, ext);
  res.download(filePath, `${baseName}_v${versionNumber}${ext}`);
});

// ── Error handler ──────────────────────────────────────────────────────────────
// eslint-disable-next-line no-unused-vars
app.use((err, _req, res, _next) => {
  if (err.code === 'INVALID_TYPE') {
    return res.status(400).json({ error: err.message });
  }
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(400).json({ error: 'File size exceeds 50MB limit' });
  }
  console.error('Unhandled error:', err.message);
  res.status(500).json({ error: 'Internal server error' });
});

// ── Start / Export ─────────────────────────────────────────────────────────────
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`ONLYOFFICE POC server running at http://localhost:${PORT}`);
    console.log(`ONLYOFFICE Document Server expected at: ${ONLYOFFICE_SERVER_URL}`);
    console.log(`Data directory: ${DATA_PATH}`);
  });
}

module.exports = app;
