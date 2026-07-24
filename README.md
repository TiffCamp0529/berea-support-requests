# Berea College — Administrative Support Request System

A secure, web-based intake and tracking system for the Office of the Provost and Dean of Faculty Development at Berea College. Supervisors submit support requests through a detailed branching form; staff (the executive administrative assistant, colleagues, and support personnel) manage them from a shared dashboard, track and delegate tasks, and everyone stays in sync through status updates, a notes thread, and email notifications.

Built with [Streamlit](https://streamlit.io) and deployable free on Streamlit Community Cloud.

---

## Features

- **Password-protected access** with four roles: Admin, Colleague, Support, and Supervisor
- **13-branch request form** covering events, meetings, travel, documents, communications, room reservations, catering, guests, budget, research, faculty development, presentations, and general support — with multiple-location support and an optional requestor checklist
- **Admin dashboard** with nav cards that summarize workload at a glance
- **Supervisor portal** — supervisors see all tasks with a filter to toggle between their own requests and everyone's; can cancel/recall their own requests
- **Task management** — list and card (kanban) views, category/priority/owner filters, recurring tasks, interactive checklists with progress bars
- **Today's Focus** — a consolidated daily checklist pulling unchecked items from every active task
- **Weekly Planner** — a Monday–Friday grid of tasks by deadline, with overdue and upcoming sections
- **Staff-to-staff delegation** — colleagues and admins can assign tasks to other staff, who accept or decline
- **Accept & convert** — accepting a request turns it into a tracked task, generates an optional AI checklist, sets calendar fields, and notifies the submitter
- **Edit-locking** — once a task is accepted, only the assigned person (and admins) can edit it; everyone else is notes-only
- **Hidden activity log** — every task change is logged; admins can reveal a task's full history
- **User management** (admin only) — add, edit, deactivate, and reset passwords, all from the app
- **Self-service password change** for every user
- **Email notifications** with per-user preferences
- **Calendar integration** — download .ics files per task, or auto-sync to Outlook, Planner, and To Do via Power Automate
- **Persistent storage** via Google Sheets (falls back to session memory if not configured)
- **Confidentiality notices** throughout, reminding users not to enter PII
- Styled to the Berea College 2026 brand standards

---

## Quick start (deploy to the web)

Full click-by-click instructions are in **[DEPLOYMENT.md](DEPLOYMENT.md)**. In brief:

1. Create a free [GitHub](https://github.com) account and a new repository
2. Add `app.py`, `requirements.txt`, and `.streamlit/config.toml`
3. Create a free [Streamlit Community Cloud](https://streamlit.io/cloud) account and deploy the repository
4. Add your passwords and settings in the Streamlit **Secrets** panel
5. (Optional) Connect a Google Sheet for permanent storage and Gmail for email notifications

**Setup time:** about 45–60 minutes, one time.

---

## Roles

| Role | Can do |
|------|--------|
| **Admin** | Everything: dashboard, all tasks, accept/deny, delegate, user management, activity logs |
| **Colleague** | See and filter all tasks, delegate to and receive from staff, create own tasks, receive supervisor requests — no user management |
| **Support** | Dashboard, manage tasks, accept/deny requests — no user management |
| **Supervisor** | Submit requests, view all tasks (filterable), cancel their own requests |

---

## Storage (Google Sheets tabs — created automatically)

| Tab | Holds |
|-----|-------|
| `submissions` | Every support request |
| `tasks` | All tasks (personal, accepted, and delegated) |
| `notes` | The notes/comments threads |
| `users` | User accounts (passwords stored hashed) |
| `notification_prefs` | Per-user email preferences |
| `audit_log` | The hidden activity log |
| `sessions` | Short-lived tokens for 10-minute auto-relogin |
| `delegations` | Staff-to-staff task requests |

You never create these by hand — the app adds them and keeps their columns in sync automatically.

---

## Run it locally (optional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Without a secrets file it runs in session mode — fully functional, but data resets when you stop the app.

---

## Files in this repository

| File | What it is |
|------|-----------|
| `app.py` | The entire application |
| `requirements.txt` | Python libraries the app needs |
| `.streamlit/config.toml` | App theme and settings (Berea brand colors) |
| `DEPLOYMENT.md` | Step-by-step setup guide |
| `.gitignore` | Keeps secrets and local files out of the repository |

---

*Office of the Provost & Dean of Faculty Development · Berea College*
