# Ganatra Clinic — web app

The clinic's login screen, dashboard, case records, collections, and
financial statements, built as a normal static website that talks to the
`ganatra-clinic-backend` API.

## Deploy on Render (Static Site)

1. Upload this folder to a new GitHub repository (same way as the backend —
   drag the *contents* of this folder in, not the folder itself, so
   `package.json` sits at the top level of the repo).
2. On Render: **New → Static Site**, connect the repo.
3. Build Command: `npm install && npm run build`
4. Publish Directory: `dist`
5. Deploy. Render gives you a normal web address like
   `https://ganatra-clinic-app.onrender.com` — open it in any browser.
6. On the login screen, paste your backend's address
   (`https://ganatra-clinic.onrender.com`, no trailing slash) into the
   "API server URL" field, then log in or tap the sample-data button.

## Local development

```bash
npm install
npm run dev
```
