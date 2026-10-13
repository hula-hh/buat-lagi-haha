# YUSEN — Smart Attendance System

A responsive attendance web app with QR scanning and Google Sheets storage through Google Apps Script.

## Features
- Responsive dashboard and attendance statistics
- QR scanning with the device camera
- Check-in / check-out, duplicate prevention, and attendance history
- Google Sheets as the data store
- Search/filter-ready attendance data and CSV export
- Apps Script backend starter in `apps-script/Code.gs`

## Quick start

1. Open the Google Sheet that will store attendance data.
2. Open **Extensions → Apps Script**.
3. Copy `apps-script/Code.gs` into the Apps Script editor.
4. In Apps Script, run `setupYusen()` once and approve permissions. This creates the required sheets and headers.
5. In **Project Settings → Script Properties**, set `ADMIN_KEY` to a long random value. Do not commit this value to GitHub.
6. Deploy → **New deployment** → **Web app**. Execute as **Me**. Set access according to your organization's needs. Copy the deployment URL ending in `/exec`.
7. Open the YUSEN site, choose **Pengaturan**, paste the deployment URL and the admin key (admin key is stored only in that browser's local storage). For a production deployment, use an authenticated server-side admin flow instead of exposing a shared admin key in a public client.
8. For GitHub Pages: Repository **Settings → Pages → Deploy from a branch → main → /(root)**.

## Important security notes
- A public static website cannot keep a secret in frontend code. If the site is public, do not use a shared admin key in the browser as a production security boundary.
- Restrict access to the Apps Script deployment where possible, and use organization sign-in / an authenticated server for real employee data.
- QR codes should contain a random user token, not a name or student ID. Rotate tokens if they are shared.
- Apps Script web apps and browser cross-origin policies can vary. The starter uses JSONP-compatible GET requests for the static GitHub Pages frontend. GET requests are convenient for a demo but are not ideal for sensitive production mutations; use an authenticated POST/proxy or serve the UI from Apps Script for production.
- Do not put personal or sensitive data in a public repository.

## Sheets created
- `Users`: userId, name, department, qrToken, active
- `Attendance`: timestamp, date, time, userId, name, action, status
- `LeaveRequests`: timestamp, userId, name, type, date, note, status
- `Settings`: key, value

## QR format
QR codes contain the user's random `qrToken`. Create/manage user tokens from a trusted admin environment. This starter provides the UI and core API scaffold; harden identity and authorization before using it for official attendance.
