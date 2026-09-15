# GSM Management System

A Flask web application for managing GSM/SIM inventory records stored in Microsoft SQL Server. It provides secure role-based access for administrators, electrical users, and electronics users.

## Included capabilities

- User login and role-based field permissions.
- Search, view, add, update, and delete GSM records.
- Excel (`.xlsx`) upload with row validation and duplicate handling.
- Downloadable Excel reports and an upload-template download.
- Company-brand image management and configurable interface appearance.

## Technology

- Python / Flask
- Microsoft SQL Server via ODBC Driver 18
- HTML, CSS, and JavaScript templates

## Local setup

1. Install Python 3.10 or later and Microsoft ODBC Driver 18 for SQL Server.
2. Create and activate a virtual environment.
3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Copy `.env.example` to `.env`, then set the Flask secret and SQL Server values.
5. Run the application:

   ```bash
   flask --app app run
   ```

6. Open `http://127.0.0.1:5000` in a browser.

## Security notes

`.env`, uploaded branding files, and local appearance settings are intentionally excluded from version control. Do not commit production credentials. Configure SQL Server access and the application secret through environment variables in each deployment environment.

## Project structure

```text
app.py              Flask routes, authentication, data access, reporting, and upload logic
Templates/          Application pages (login, dashboard, GSM master, upload)
.env.example        Safe configuration template
requirements.txt    Python dependencies
```
