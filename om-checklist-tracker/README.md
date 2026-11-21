# O and M Checklist Tracker

A local-first web app for daily solar-site O&M observations. Operators complete the checklist in a browser; administrators provision users/sites, review daily status, and download submitted checklists as Excel or PDF.

## What is included

- Admin and operations-user workflows, with no public account registration.
- One assigned operations user per site; each operator only sees their own site's checklist/history.
- Browser checklist imported from `Trafo_Checklist_Nov_2026_Populated.xlsx` (28 checklist items); seed daily readings in the source workbook are intentionally not imported as live submissions.
- Numeric temperature inputs, status controls and text observations, comments, and overall remarks.
- One admin-configurable daily cutoff shared across sites/users, fixed to India Standard Time. Default is 17:00 IST; submissions are locked at or after the cutoff.
- Admin dashboard for today's site status, history, and server-protected Excel/PDF downloads.

## Local setup (Windows PowerShell)

1. Install Python 3.11+ (Python 3.14 is also suitable) and open PowerShell in this folder.
2. Create and activate a virtual environment:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

3. Install dependencies:

   ```powershell
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

4. Prepare the local SQLite database and import the checklist:

   ```powershell
   python manage.py migrate
   python manage.py import_checklist
   ```

   The workbook is bundled in the project root. To import another workbook, pass its path, e.g. `python manage.py import_checklist "C:\\path\\Trafo.xlsx" --site "Site name"`.

5. Create the first administrator. The command prompts for the password without echoing it:

   ```powershell
   python manage.py create_admin admin --email admin@example.com
   ```

   For operational use, create a fresh administrator with your own strong password. Do not reuse any temporary password used for local smoke testing.

6. Start the development server:

   ```powershell
   python manage.py runserver
   ```

7. Browse to http://127.0.0.1:8000/login/ and sign in as admin. Add sites, then create one operator account per site. Set a shared cutoff in **Settings**.

## Tests

```powershell
python manage.py test checklist
```

Tests cover cutoff behavior, input validation, unique site/day submissions, admin exports, and rejection of user export access.

## Data and permissions

- The original workbook is preserved. Its populated November sample readings are reference data only; importing it creates checklist definitions and does not mark any site as submitted.
- User/site and role checks happen on the server. Export routes require administrator/staff status; downloads are generated in memory and are not published under static or media paths.
- The daily cutoff is enforced when a user opens and submits the form. The server uses `Asia/Kolkata`; at the cutoff instant and afterward, submissions are rejected.
- SQLite data is local to the project. Back up `db.sqlite3` while the server is stopped or using a consistent SQLite backup procedure.

## Deployment recommendation (after local acceptance)

Render is a practical first host for a small team: deploy the Django app with a managed PostgreSQL database, HTTPS, a long `DJANGO_SECRET_KEY`, `DJANGO_DEBUG=0`, and `DJANGO_ALLOWED_HOSTS` set to the production host. Supply `DATABASE_URL` via the provider's environment settings; configure automated database backups. Check organizational data-residency requirements and select the region before deployment. Do not deploy with the development secret. For persistent user-uploaded or generated artifacts in a future version, use private object storage; the current exports are on-demand and in-memory.

## Scope notes

Checklist edit/versioning screens, notification/reminder delivery, and post-cutoff admin override are not part of the first version. To change checklist definitions, create a new template version/import workflow rather than modifying historical submitted data.
