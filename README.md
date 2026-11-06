# agent-performance-tracking
# Daily Performance Tracker (Flask)

A local Flask + SQLite performance tracking app based on the uploaded 20-page specification.

## Features included

- Three roles: Admin, Team Lead and Agent.
- Secure password hashing and role-protected pages.
- Admin user creation, agent-to-Team-Lead assignment, activation/deactivation.
- Monthly targets for average calls, approval percentage and error percentage.
- Agent daily performance entry and update using a calendar.
- Monthly KPI calculations and catch-up reminders.
- Week-off roster management; roster week-offs are excluded from catch-up calculations.
- Team Lead team summary, agent-level detail and charts.
- Entry change history / audit trail.
- Excel export of monthly daily performance.
- Dark blue dashboard theme.

## Requirements

- Windows 10/11, macOS or Linux
- Python 3.10 or newer
- VS Code (recommended)

## Run in VS Code (Windows)

1. Extract this ZIP file to a folder, for example `C:\Projects\Daily_Performance_Tracker`.
2. Open that folder in VS Code.
3. Open **Terminal → New Terminal**.
4. Create a virtual environment:

   ```powershell
   py -m venv .venv
   ```

5. Activate it:

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

   If PowerShell blocks activation, run this once in the same terminal:

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\.venv\Scripts\Activate.ps1
   ```

6. Install dependencies:

   ```powershell
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

7. Start the app:

   ```powershell
   python app.py
   ```

8. Open this address in your browser:

   `http://127.0.0.1:5000`

## First login

- Username: `admin`
- Password: `admin123`

**Important:** Change the admin password immediately after login. The initial password is intended only for local first-run setup.

## Suggested setup sequence

1. Sign in as Admin and change the password.
2. Create Team Lead accounts.
3. Create Agent accounts and assign each agent to a Team Lead.
4. Configure monthly targets.
5. Add week-off dates under **Week-off roster**.
6. Agents sign in and enter their daily calls, approvals and errors.
7. Team Leads monitor the team dashboard and export the monthly report.

## Database

- Local development uses SQLite. The `tracker.db` file is created automatically in the project folder.
- When the `DATABASE_URL` environment variable is configured, the app uses PostgreSQL instead.
- On Render, create a PostgreSQL database and add its **Internal Database URL** to the web service as the `DATABASE_URL` environment variable. The app accepts Render's `postgres://` / `postgresql://` URL formats and uses the Psycopg 3 driver.
- Existing SQLite records are **not copied automatically** when you switch to PostgreSQL. Export or migrate your existing users, daily entries, entry history, roster and targets before switching the live service. Keep a backup of `tracker.db`.
- Database tables are created automatically at application startup. This creates missing tables; it does not copy records from SQLite or alter existing table structures.

## Notes and boundaries

- This is a local development application. It binds to `127.0.0.1` and is not exposed to the public internet.
- Excel export is included. The PDF mentions Google Sheets and roster Excel imports as upcoming parts, but does not provide their implementation details; those integrations are therefore not represented as completed features here.
- The included charts load Chart.js from a CDN, so charts require an internet connection. The rest of the app runs locally.
- Before public deployment, set a strong random `SECRET_KEY`, use HTTPS, configure secure cookies, add CSRF protection and review backups, access controls and operational security.
