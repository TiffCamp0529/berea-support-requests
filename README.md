# Berea College — Administrative Support Request System

A secure, web-based intake and tracking system for the Office of the Provost and Dean of Faculty Development at Berea College. Supervisors submit support requests through a detailed branching form; the executive administrative assistant manages them from an admin dashboard, tracks personal tasks alongside incoming requests, and everyone stays in sync through status updates, a notes thread, and email notifications.

Built with [Streamlit](https://streamlit.io) and deployable free on Streamlit Community Cloud.

---

## Features

- **Password-protected access** with separate logins for the administrator and each supervisor
- **13-branch request form** covering events, meetings, travel, documents, communications, room reservations, catering, guests, budget, research, faculty development, presentations, and general support
- **Admin dashboard** — view, filter, and search every request; update statuses; export to CSV
- **Supervisor portal** — read-only view of a supervisor's own requests with live status
- **Notes thread** on every request for two-way updates without editing the original submission
- **Personal task list** for the administrator, including recurring tasks, shown alongside requests as one combined workload
- **Accept & convert** — accepting a request turns it into a tracked task automatically and notifies the submitter
- **Email notifications** for new submissions, status changes, acceptances, and new notes, with per-user preferences
- **Persistent storage** via Google Sheets (falls back to session memory if not configured)
- Styled to the Berea College 2026 brand standards

---

## Quick start (deploy to the web)

Full click-by-click instructions for a non-technical user are in **[DEPLOYMENT.md](DEPLOYMENT.md)**. In brief:

1. Create a free [GitHub](https://github.com) account and a new repository
2. Add `app.py`, `requirements.txt`, and `.streamlit/config.toml` (the guide shows how)
3. Create a free [Streamlit Community Cloud](https://streamlit.io/cloud) account and deploy the repository
4. Add your passwords and settings in the Streamlit **Secrets** panel
5. (Optional) Connect a Google Sheet for permanent storage and Gmail for email notifications

**Setup time:** about 45–60 minutes, one time.

---

## Run it locally (optional)

If you'd like to test the app on your own computer first:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL it prints (usually `http://localhost:8501`). Without a secrets file it runs in session mode — fully functional, but data resets when you stop the app. To test with your own passwords locally, create a file at `.streamlit/secrets.toml` using `secrets.toml.example` as a template. Never commit that file to GitHub.

---

## Files in this repository

| File | What it is |
|------|-----------|
| `app.py` | The entire application |
| `requirements.txt` | Python libraries the app needs |
| `.streamlit/config.toml` | App theme and settings (Berea brand colors) |
| `.streamlit/secrets.toml.example` | Template for passwords and settings — copy from this, never commit the real one |
| `DEPLOYMENT.md` | Step-by-step setup guide for non-technical users |
| `.gitignore` | Keeps secrets and local files out of the repository |

---

## Configuration

All passwords, supervisor accounts, email settings, and the Google Sheet connection are stored as **secrets** — never in the code. Locally they live in `.streamlit/secrets.toml`; in production they go in the Streamlit Cloud Secrets panel. See `secrets.toml.example` for the format and `DEPLOYMENT.md` for a walkthrough.

Required at minimum:

- `ADMIN_PASSWORD` — the administrator's login
- `ADMIN_EMAIL` — where the administrator receives notifications
- At least one `[supervisors.*]` block

Optional:

- `GMAIL_USER` and `GMAIL_APP_PASSWORD` — enable email notifications
- `GOOGLE_SHEET_ID` and `GOOGLE_SERVICE_ACCOUNT` — enable permanent storage

---

## Security notes

- Access is gated by password; each supervisor sees only their own requests
- All traffic is encrypted over HTTPS (enforced by Streamlit Cloud)
- Secrets are stored in Streamlit's encrypted vault, never in the repository
- `.gitignore` prevents the real `secrets.toml` from ever being committed
- Personal tasks are visible only to the administrator
- For FERPA-sensitive data, consult Berea College IT about self-hosted options

---

## Maintenance

- **Update the app:** edit `app.py` on GitHub and commit — Streamlit redeploys automatically
- **Add a supervisor:** add a `[supervisors.*]` block in the Streamlit Secrets panel
- **Retrieve data:** download a CSV from the admin dashboard, or open the connected Google Sheet directly

---

*Office of the Provost and Dean of Faculty Development, Berea College.*
