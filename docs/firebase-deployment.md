# Firebase Deployment (Hosting + Cloud Run)

This app can be deployed with Firebase, but because it is an Express server and depends on ONLYOFFICE Document Server, the practical setup is:

- `onlyoffice-app` (this Node.js app) on Cloud Run
- `onlyoffice-docs` (ONLYOFFICE image) on Cloud Run
- Firebase Hosting in front of `onlyoffice-app` using a rewrite

## Important POC caveat

Current storage is local filesystem (`data/`) and `db.json`.
On Cloud Run, local disk is ephemeral and instance-local, so data will not be durable across restarts/scaling.

For production, move files and metadata to managed storage (for example: Cloud Storage + Firestore/Cloud SQL).

## 1) Prerequisites

- Google Cloud project with billing enabled
- Firebase project linked to the same Google Cloud project
- Installed CLIs:
  - `gcloud`
  - `firebase`

Login:

```bash
gcloud auth login
gcloud config set project YOUR_PROJECT_ID
firebase login
firebase use YOUR_PROJECT_ID
```

Enable required APIs:

```bash
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com firebasehosting.googleapis.com
```

## 2) Deploy ONLYOFFICE to Cloud Run

Set a JWT secret once and reuse it for both services:

```bash
export JWT_SECRET="replace-with-a-strong-secret"
```

Deploy ONLYOFFICE:

```bash
gcloud run deploy onlyoffice-docs \
  --image onlyoffice/documentserver:latest \
  --region us-central1 \
  --platform managed \
  --allow-unauthenticated \
  --port 80 \
  --cpu 2 \
  --memory 4Gi \
  --set-env-vars JWT_ENABLED=true,JWT_SECRET=$JWT_SECRET,JWT_HEADER=Authorization,JWT_PREFIX=Bearer
```

Capture ONLYOFFICE URL:

```bash
export ONLYOFFICE_SERVER_URL="$(gcloud run services describe onlyoffice-docs --region us-central1 --format='value(status.url)')"
```

## 3) Deploy Node app to Cloud Run

Deploy from source (uses project Dockerfile):

```bash
gcloud run deploy onlyoffice-app \
  --source . \
  --region us-central1 \
  --platform managed \
  --allow-unauthenticated \
  --set-env-vars NODE_ENV=production,JWT_SECRET=$JWT_SECRET,ONLYOFFICE_SERVER_URL=$ONLYOFFICE_SERVER_URL
```

Capture app URL:

```bash
export APP_URL="$(gcloud run services describe onlyoffice-app --region us-central1 --format='value(status.url)')"
```

Update app env so callback/file URLs are reachable by ONLYOFFICE:

```bash
gcloud run services update onlyoffice-app \
  --region us-central1 \
  --set-env-vars NODE_ENV=production,JWT_SECRET=$JWT_SECRET,ONLYOFFICE_SERVER_URL=$ONLYOFFICE_SERVER_URL,APP_URL=$APP_URL,ONLYOFFICE_APP_URL=$APP_URL
```

## 4) Configure Firebase Hosting

1. Copy `.firebaserc.example` to `.firebaserc` and set your project id.
2. Confirm `firebase.json` rewrite points to:
   - serviceId: `onlyoffice-app`
   - region: `us-central1`

Deploy hosting:

```bash
firebase deploy --only hosting
```

## 5) Verify

- Open your Firebase Hosting URL.
- Upload a `.docx`/`.doc`/`.pdf`.
- Open editor and verify it loads ONLYOFFICE and saves versions.

## Troubleshooting

- Editor script fails to load:
  - verify `ONLYOFFICE_SERVER_URL` on `onlyoffice-app` points to `onlyoffice-docs` URL.
- Callback save fails:
  - confirm both services share same `JWT_SECRET`.
  - confirm `APP_URL` and `ONLYOFFICE_APP_URL` are the `onlyoffice-app` URL.
- Data disappears after restart/scale:
  - expected with current POC storage; migrate to managed storage for persistence.
