# Deploy from GitHub to Firebase (No local CLI required)

This setup deploys directly from GitHub Actions:

- Cloud Run service: `onlyoffice-docs` (ONLYOFFICE Document Server)
- Cloud Run service: `onlyoffice-app` (Node.js app)
- Firebase Hosting rewrite to `onlyoffice-app`

Workflow file: `.github/workflows/deploy-firebase.yml`

## 1) Create a deploy service account in Google Cloud

In Google Cloud Console, create a service account (example name: `github-deployer`).

Grant these roles on your project:

- Cloud Run Admin
- Service Account User
- Cloud Build Editor
- Artifact Registry Writer
- Firebase Hosting Admin
- Service Usage Admin

Create a JSON key for this service account and copy its content.

## 2) Configure GitHub repository secrets and variable

In your GitHub repo: Settings -> Secrets and variables -> Actions

Add secrets:

- `GCP_PROJECT_ID`: your Google Cloud / Firebase project id
- `JWT_SECRET`: shared JWT secret for app + ONLYOFFICE
- `GCP_SA_KEY`: full JSON key for the deploy service account

Optional repository variable:

- `GCP_REGION`: Cloud Run region (default used by workflow is `us-central1`)

## 3) Ensure Firebase Hosting rewrite targets Cloud Run

This repo already includes `firebase.json` configured for:

- serviceId: `onlyoffice-app`
- region: `us-central1`

If you use a different region, update `firebase.json` to match.

## 4) Trigger deployment

- Push to `main`, or
- Run workflow manually from GitHub Actions tab

The workflow will:

1. Authenticate to GCP using service account key
2. Enable required APIs
3. Deploy ONLYOFFICE to Cloud Run
4. Deploy app to Cloud Run and set runtime env vars
5. Deploy Firebase Hosting (live channel)

## 5) Notes for this POC

Current persistence uses local files (`data/`) and `db.json`.
On Cloud Run this storage is ephemeral, so data is not durable across restarts/scaling.

For production, move storage to managed services (for example Cloud Storage + Firestore/Cloud SQL).
