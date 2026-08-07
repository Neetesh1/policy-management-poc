'use strict';

const express = require('express');
const multer = require('multer');
const jwt = require('jsonwebtoken');
const axios = require('axios');
const { v4: uuidv4, validate: uuidValidate } = require('uuid');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

// ── Configuration ──────────────────────────────────────────────────────────────
const JWT_SECRET = process.env.JWT_SECRET || 'poc-secret-change-in-production';
const APP_URL = process.env.APP_URL || 'http://localhost:3000';
// ONLYOFFICE_APP_URL: internal URL that ONLYOFFICE server uses to reach this app
// (in Docker use container name, e.g. http://app:3000; default = APP_URL for local dev)
const ONLYOFFICE_APP_URL = process.env.ONLYOFFICE_APP_URL || APP_URL;
// ONLYOFFICE_SERVER_URL: URL browser uses to load the ONLYOFFICE editor API script
const ONLYOFFICE_SERVER_URL = process.env.ONLYOFFICE_SERVER_URL || 'http://localhost:8080';
// ONLYOFFICE_DOCSERVER_INTERNAL_URL: URL this Node app uses to call the ONLYOFFICE
// Conversion API (POST /converter) directly — in Docker use the service name
// (e.g. http://onlyoffice), default = ONLYOFFICE_SERVER_URL for local dev.
const ONLYOFFICE_DOCSERVER_INTERNAL_URL = process.env.ONLYOFFICE_DOCSERVER_INTERNAL_URL || ONLYOFFICE_SERVER_URL;
const PORT = parseInt(process.env.PORT || '3000', 10);

// White-label branding (https://api.onlyoffice.com/docs/docs-api/usage-api/config/editor/customization/customization-white-label/)
// NOTE: these fields require the "extended white label license" add-on for ONLYOFFICE Docs Developer —
// on a trial/Community-tier server they are silently ignored (canBranding check fails server-side).
const WHITE_LABEL_LOADER_NAME = process.env.WHITE_LABEL_LOADER_NAME || 'Loading policy document…';
const WHITE_LABEL_LOADER_LOGO_URL = process.env.WHITE_LABEL_LOADER_LOGO_URL || '';
const WHITE_LABEL_SHOW_ABOUT = process.env.WHITE_LABEL_SHOW_ABOUT !== 'false';

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

// ── Clean export helper ─────────────────────────────────────────────────────────
// Strips our policy/regulation content control tags (and their paragraph highlight)
// plus all comments from a .docx buffer, without touching the stored/live version.
async function stripTagsAndComments(fileBuffer) {
  const zip = await JSZip.loadAsync(fileBuffer);
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) return fileBuffer;

  const xmlText = await docXmlFile.async('string');
  const doc = new DOMParser().parseFromString(xmlText, 'text/xml');

  // Drop comment anchors/references so no comments remain visible/linked.
  ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference'].forEach(tagName => {
    Array.from(doc.getElementsByTagName(tagName)).forEach(node => {
      if (node.parentNode) node.parentNode.removeChild(node);
    });
  });

  // Unwrap only the content controls our plugins created ({policy:...} / {regulation:...}),
  // clearing the paragraph highlight too, and leaving any other native content controls intact.
  Array.from(doc.getElementsByTagName('w:sdt')).forEach(sdt => {
    const sdtPr = sdt.getElementsByTagName('w:sdtPr')[0];
    const tagEl = sdtPr && sdtPr.getElementsByTagName('w:tag')[0];
    const tagVal = tagEl && tagEl.getAttribute('w:val');
    if (!tagVal || !/^\{(policy|regulation):/.test(tagVal)) return;

    const sdtContent = sdt.getElementsByTagName('w:sdtContent')[0];
    if (!sdtContent) return;

    Array.from(sdtContent.getElementsByTagName('w:shd')).forEach(shd => {
      if (shd.parentNode) shd.parentNode.removeChild(shd);
    });

    const parent = sdt.parentNode;
    while (sdtContent.firstChild) {
      parent.insertBefore(sdtContent.firstChild, sdt);
    }
    parent.removeChild(sdt);
  });

  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
  return zip.generateAsync({ type: 'nodebuffer' });
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

// Static catalog of regulations that can be assigned to a paragraph/title (POC data).
const REGULATIONS = [
  { code: 'GDPR', name: 'General Data Protection Regulation' },
  { code: 'HIPAA', name: 'Health Insurance Portability and Accountability Act' },
  { code: 'CCPA', name: 'California Consumer Privacy Act' },
  { code: 'SOX', name: 'Sarbanes-Oxley Act' },
  { code: 'PCI-DSS', name: 'Payment Card Industry Data Security Standard' },
  { code: 'ISO-27001', name: 'ISO/IEC 27001 Information Security Management' }
];

// ── Routes ─────────────────────────────────────────────────────────────────────

/**
 * GET /api/config
 * Returns app configuration consumed by the frontend (e.g. ONLYOFFICE server URL).
 */
app.get('/api/config', (_req, res) => {
  res.json({ onlyofficeServerUrl: ONLYOFFICE_SERVER_URL });
});

/**
 * GET /api/regulations
 * Returns the catalog of regulations that can be assigned to a paragraph/title.
 */
app.get('/api/regulations', (_req, res) => {
  res.json(REGULATIONS);
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
      extension: d.extension || '.docx',
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
        // Standard branding (no premium license needed): collapses the ribbon to just the
        // tab row by default; clicking a tab expands its buttons for that session.
        compactToolbar: true,
        // White-label branding — requires the extended white-label license; no-ops otherwise.
        about: WHITE_LABEL_SHOW_ABOUT,
        loaderName: WHITE_LABEL_LOADER_NAME,
        ...(WHITE_LABEL_LOADER_LOGO_URL ? { loaderLogo: WHITE_LABEL_LOADER_LOGO_URL } : {}),
        // Hiding specific toolbar tabs (Protection/View) requires the same commercial
        // white-label license as the fields above — silently ignored without it (e.g. an
        // expired/missing license.lic). Home cannot be hidden per the ONLYOFFICE docs.
        // Plugins tab is kept visible on purpose: with autostart disabled, it's the only
        // way to manually launch Tag List / Policy Actions from the ribbon.
        layout: {
          toolbar: {
            protect: false,
            view: false
          }
        }
      },
      plugins: {
        // Neither plugin autostarts: both icons still show in the Plugins tab, but their
        // panels stay hidden until the user clicks one — nothing opens automatically on load.
        // Trade-off: Policy Actions' right-click "Policy Tagging" menu only attaches once its
        // icon has been clicked at least once per session (init() is what wires the listener).
        autostart: [],
        // Disables the built-in AI plugin/toolbar tab (bundled since ONLYOFFICE Docs 9.0.4).
        // plugins.disable requires no license — it fully blocks the plugin, not just hides UI.
        disable: [
          'asc.{9DC93CDB-B576-4F0C-B55E-FCC9C48DD007}'
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
      // body.url points at ONLYOFFICE_SERVER_URL (the browser-facing host), which isn't
      // reachable from inside this container — rewrite it to the internal Docker service
      // so we can actually fetch the saved file instead of getting ECONNREFUSED, which
      // silently fails every save and leaves the doc server's cache in a broken state.
      const internalFileUrl = body.url.replace(/^https?:\/\/[^/]+/, ONLYOFFICE_DOCSERVER_INTERNAL_URL);
      const fileResponse = await axios.get(internalFileUrl, {
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

// Automation API-reported events we allow into the audit trail (see public/editor.html
// initAutomationConnector). Whitelisted to prevent arbitrary/unbounded client input.
const ALLOWED_AUTOMATION_ACTIONS = new Set([
  'CONTENT_CONTROL_CHANGED',
  'COMMENT_ADDED',
  'COMMENT_CHANGED',
  'COMMENT_REMOVED',
  'REVIEW_CHANGES_ACCEPTED',
  'REVIEW_CHANGES_REJECTED'
]);

/**
 * POST /api/audit-event/:documentId
 * Records a client-side Automation API event (paragraph tag change, comment
 * lifecycle, review accept/reject) into the document's audit trail, so it shows
 * up in the /versions revision history alongside upload/save events.
 * https://api.onlyoffice.com/docs/docs-api/usage-api/automation-api/connector-class/
 * https://api.onlyoffice.com/docs/plugins/interacting-with-editors/document-api/Events/
 */
app.post('/api/audit-event/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const { action, userId, meta } = req.body || {};

  if (!ALLOWED_AUTOMATION_ACTIONS.has(action)) {
    return res.status(400).json({ error: 'Invalid or unsupported action' });
  }

  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  // Re-serialize meta through JSON to strip functions/prototypes and cap size.
  let safeMeta;
  try {
    const str = JSON.stringify(meta);
    if (str && str.length <= 2000) safeMeta = JSON.parse(str);
  } catch {
    safeMeta = undefined;
  }

  doc.audit.push({
    action,
    versionNo: doc.currentVersion,
    userId: typeof userId === 'string' ? userId.slice(0, 100) : 'unknown',
    meta: safeMeta,
    timestamp: new Date().toISOString()
  });
  writeDB(db);

  res.json({ success: true });
});

/**
 * POST /api/tags/:documentId
 * Persists structured paragraph-tag metadata (key, color, content control id) in
 * the custom DB instead of a Word comment, linked to the document via documentId
 * and to the specific paragraph via the content control's InternalId.
 */
app.post('/api/tags/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const { tagKey, colorKey, contentControlId, userId, tagType, label } = req.body || {};

  if (typeof tagKey !== 'string' || !tagKey.trim()) {
    return res.status(400).json({ error: 'tagKey is required' });
  }

  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const tag = {
    id: uuidv4(),
    tagKey: tagKey.slice(0, 200),
    tagType: typeof tagType === 'string' ? tagType.slice(0, 50) : 'policy',
    label: typeof label === 'string' ? label.slice(0, 200) : tagKey.slice(0, 200),
    colorKey: typeof colorKey === 'string' ? colorKey.slice(0, 50) : 'custom',
    contentControlId: contentControlId != null ? String(contentControlId).slice(0, 100) : null,
    userId: typeof userId === 'string' ? userId.slice(0, 100) : 'unknown',
    createdAt: new Date().toISOString()
  };

  doc.tags = doc.tags || [];
  doc.tags.push(tag);
  writeDB(db);

  res.json({ success: true, tag });
});

/**
 * GET /api/tags/:documentId
 * Returns all tags recorded against a document.
 */
app.get('/api/tags/:documentId', validateDocId, (req, res) => {
  const db = readDB();
  const doc = db.documents.find(d => d.id === req.params.documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });
  res.json(doc.tags || []);
});

/**
 * GET /history-file/:documentId/:versionNo?token=<jwt>
 * Serves a specific historical version's file to ONLYOFFICE's native Version
 * History panel (used by refreshHistory/setHistoryData "url" fields).
 */
app.get('/history-file/:documentId/:versionNo', validateDocId, (req, res) => {
  const { documentId, versionNo } = req.params;
  const rawToken = req.query.token;
  if (!rawToken) return res.status(401).json({ error: 'Token required' });
  try {
    jwt.verify(rawToken, JWT_SECRET);
  } catch {
    return res.status(401).json({ error: 'Invalid or expired token' });
  }

  const versionNumber = parseInt(versionNo, 10);
  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const version = doc.versions.find(v => v.versionNo === versionNumber);
  if (!version) return res.status(404).json({ error: 'Version not found' });

  const ext = doc.extension || '.docx';
  const filePath = safePath(STORAGE_PATH, documentId, 'versions', `v${versionNumber}${ext}`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found on disk' });

  const { mimeType } = getDocMeta(ext);
  res.setHeader('Content-Type', mimeType);
  res.download(filePath, doc.title);
});

/**
 * GET /api/history/:documentId
 * Returns the document's version history shaped for ONLYOFFICE's native Version
 * History panel (docEditor.refreshHistory / setHistoryData).
 * https://api.onlyoffice.com/docs/docs-api/usage-api/config/editor/events/onrequesthistory/
 */
app.get('/api/history/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const ext = doc.extension || '.docx';
  const { fileType } = getDocMeta(ext);

  const history = doc.versions.map(v => {
    const fileToken = jwt.sign(
      { documentId, action: 'read', version: v.versionNo },
      JWT_SECRET,
      { expiresIn: '2h' }
    );
    const userInfo = { id: v.createdBy, name: v.createdBy };
    return {
      version: v.versionNo,
      key: `${documentId.replace(/-/g, '')}_v${v.versionNo}`,
      created: v.createdAt,
      user: userInfo,
      changes: [{ created: v.createdAt, user: userInfo }],
      fileType,
      // Extension appended after the version number (parseInt ignores it) so ONLYOFFICE's
      // "direct link to file" validation accepts the URL (it checks for a recognizable extension).
      url: `${ONLYOFFICE_APP_URL}/history-file/${documentId}/${v.versionNo}${ext}?token=${encodeURIComponent(fileToken)}`
    };
  });

  res.json({ currentVersion: doc.currentVersion, history });
});

/**
 * GET /api/history-data/:documentId/:versionNo
 * Returns a single version shaped + signed for docEditor.setHistoryData(). ONLYOFFICE
 * verifies this payload's own JWT (not just the file URL's), so it must be signed here.
 */
app.get('/api/history-data/:documentId/:versionNo', validateDocId, (req, res) => {
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
  const { fileType } = getDocMeta(ext);
  const sorted = [...doc.versions].sort((a, b) => a.versionNo - b.versionNo);
  const idx = sorted.findIndex(v => v.versionNo === versionNumber);
  const previousVersion = idx > 0 ? sorted[idx - 1] : null;

  function fileUrlFor(v) {
    const fileToken = jwt.sign(
      { documentId, action: 'read', version: v.versionNo },
      JWT_SECRET,
      { expiresIn: '2h' }
    );
    // Extension appended after the version number (parseInt ignores it) so ONLYOFFICE's
    // "direct link to file" validation accepts the URL (it checks for a recognizable extension).
    return `${ONLYOFFICE_APP_URL}/history-file/${documentId}/${v.versionNo}${ext}?token=${encodeURIComponent(fileToken)}`;
  }

  const payload = {
    version: version.versionNo,
    key: `${documentId.replace(/-/g, '')}_v${version.versionNo}`,
    fileType,
    url: fileUrlFor(version)
  };
  if (previousVersion) {
    payload.previous = {
      fileType,
      key: `${documentId.replace(/-/g, '')}_v${previousVersion.versionNo}`,
      url: fileUrlFor(previousVersion)
    };
  }

  // The whole payload (not just the file URL) must be JWT-signed for setHistoryData.
  const token = jwt.sign(payload, JWT_SECRET);
  res.json({ ...payload, token });
});

/**
 * POST /api/restore-version/:documentId
 * Restores an older version by copying it forward as a brand-new latest version —
 * non-destructive, every prior version stays on disk and in the history list.
 */
app.post('/api/restore-version/:documentId', validateDocId, (req, res) => {
  const { documentId } = req.params;
  const { versionNo, userId } = req.body || {};
  const versionNumber = parseInt(versionNo, 10);

  if (isNaN(versionNumber) || versionNumber < 1) {
    return res.status(400).json({ error: 'Invalid version number' });
  }

  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const sourceVersion = doc.versions.find(v => v.versionNo === versionNumber);
  if (!sourceVersion) return res.status(404).json({ error: 'Version not found' });

  const ext = doc.extension || '.docx';
  const sourcePath = safePath(STORAGE_PATH, documentId, 'versions', `v${versionNumber}${ext}`);
  if (!fs.existsSync(sourcePath)) return res.status(404).json({ error: 'Version file not found on disk' });

  const newVersionNo = doc.currentVersion + 1;
  const destPath = safePath(STORAGE_PATH, documentId, 'versions', `v${newVersionNo}${ext}`);
  fs.copyFileSync(sourcePath, destPath);

  const fileHash = crypto.createHash('sha256').update(fs.readFileSync(destPath)).digest('hex');
  const now = new Date().toISOString();
  const restoredBy = typeof userId === 'string' ? userId.slice(0, 100) : 'user';

  doc.versions.push({ versionNo: newVersionNo, filePath: destPath, fileHash, createdBy: restoredBy, createdAt: now });
  doc.currentVersion = newVersionNo;
  doc.updatedAt = now;
  doc.audit.push({
    action: 'VERSION_RESTORED',
    versionNo: newVersionNo,
    userId: restoredBy,
    meta: { restoredFrom: versionNumber },
    timestamp: now
  });
  writeDB(db);

  res.json({ success: true, newVersionNo });
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
 * GET /download-clean/:documentId
 * Downloads the latest version with policy/regulation tags and comments stripped out —
 * generated on the fly from the stored file; the original version on disk is untouched.
 */
app.get('/download-clean/:documentId', validateDocId, async (req, res) => {
  const { documentId } = req.params;
  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const latest = doc.versions.find(v => v.versionNo === doc.currentVersion);
  if (!latest) return res.status(404).json({ error: 'File not found' });

  const ext = doc.extension || '.docx';
  if (ext !== '.docx') {
    return res.status(400).json({ error: 'Clean export is only supported for .docx documents' });
  }

  const filePath = safePath(STORAGE_PATH, documentId, 'versions', `v${latest.versionNo}${ext}`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found on disk' });

  try {
    const cleaned = await stripTagsAndComments(fs.readFileSync(filePath));
    const baseName = path.basename(doc.title, ext);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}_clean${ext}"`);
    res.send(cleaned);
  } catch (err) {
    console.error('[download-clean] Failed:', err.message);
    res.status(500).json({ error: 'Failed to generate clean export' });
  }
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

/**
 * POST /api/convert-to-word/:documentId
 * Converts a PDF document to .docx via the ONLYOFFICE Conversion API
 * (https://api.onlyoffice.com/docs/docs-api/additional-api/conversion-api/)
 * and stores the result as a brand-new document (leaves the original PDF untouched).
 */
app.post('/api/convert-to-word/:documentId', validateDocId, async (req, res) => {
  const { documentId } = req.params;
  const db = readDB();
  const doc = db.documents.find(d => d.id === documentId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  if ((doc.extension || '.docx') !== '.pdf') {
    return res.status(400).json({ error: 'Only PDF documents can be converted to Word' });
  }

  try {
    // Short-lived token so the Document Server can fetch the source PDF
    const fileToken = jwt.sign(
      { documentId, action: 'read', version: doc.currentVersion },
      JWT_SECRET,
      { expiresIn: '10m' }
    );

    const conversionPayload = {
      async: false,
      filetype: 'pdf',
      key: crypto.randomBytes(16).toString('hex'), // must be unique per conversion request
      outputtype: 'docx',
      title: doc.title,
      url: `${ONLYOFFICE_APP_URL}/file/${documentId}?token=${encodeURIComponent(fileToken)}`
    };
    // Conversion API requires the whole request body signed as a JWT (server has JWT_ENABLED=true)
    const convertToken = jwt.sign(conversionPayload, JWT_SECRET);

    const convertResponse = await axios.post(
      `${ONLYOFFICE_DOCSERVER_INTERNAL_URL}/converter`,
      { ...conversionPayload, token: convertToken },
      {
        headers: { Authorization: `Bearer ${convertToken}`, Accept: 'application/json' },
        timeout: 60000
      }
    );

    const result = convertResponse.data;
    if (result.error) {
      console.error(`[convert] ConvertService error code ${result.error} for document ${documentId}`);
      return res.status(502).json({ error: `Conversion failed (error code ${result.error})` });
    }
    if (!result.endConvert || !result.fileUrl) {
      // Synchronous request should always finish, but guard against a slow/async response anyway
      return res.status(202).json({ pending: true, percent: result.percent || 0 });
    }

    const fileResponse = await axios.get(result.fileUrl, {
      responseType: 'arraybuffer',
      timeout: 30000,
      maxContentLength: 50 * 1024 * 1024
    });

    const newDocId = uuidv4();
    const docDir = safePath(STORAGE_PATH, newDocId, 'versions');
    fs.mkdirSync(docDir, { recursive: true });
    const destPath = path.join(docDir, 'v1.docx');
    fs.writeFileSync(destPath, fileResponse.data);

    const fileHash = crypto.createHash('sha256').update(fileResponse.data).digest('hex');
    const now = new Date().toISOString();
    const newTitle = `${path.basename(doc.title, path.extname(doc.title))}.docx`;

    const newDoc = {
      id: newDocId,
      title: newTitle,
      extension: '.docx',
      createdAt: now,
      updatedAt: now,
      currentVersion: 1,
      versions: [
        { versionNo: 1, filePath: destPath, fileHash, createdBy: 'user', createdAt: now }
      ],
      audit: [
        { action: 'CONVERTED_FROM_PDF', versionNo: 1, userId: 'user', timestamp: now, sourceDocumentId: documentId }
      ]
    };

    db.documents.push(newDoc);
    writeDB(db);

    res.json({ success: true, documentId: newDocId, title: newTitle });
  } catch (err) {
    console.error('[convert] Failed:', err.message);
    res.status(500).json({ error: 'Conversion failed' });
  }
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
