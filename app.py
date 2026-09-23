import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import json
from datetime import date, datetime, timedelta
from functools import wraps
from html import escape

from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
)
import pyodbc
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import traceback

load_dotenv()

app = Flask(__name__)

# ============================================================
# FLASK CONFIGURATION
# ============================================================

app.secret_key = os.environ["SECRET_KEY"]


# ============================================================
# SQL SERVER CONFIGURATION
# ============================================================

SQL_SERVER = os.environ["DB_SERVER"]
SQL_DATABASE = os.environ["DB_NAME"]
SQL_USERNAME = os.environ["DB_USER"]
SQL_PASSWORD = os.environ["DB_PASSWORD"]


# ============================================================
# DATABASE CONNECTION
# ============================================================


def get_db_connection():
    connection_string = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DATABASE};"
        f"UID={SQL_USERNAME};"
        f"PWD={SQL_PASSWORD};"
        "Encrypt=no;"
        "TrustServerCertificate=yes;"
        "Connection Timeout=15;"
    )
    return pyodbc.connect(connection_string)


GSM_COLUMNS = (
    "Sim_Received_Date",
    "SIM_NO",
    "Mobile_No",
    "IMEI_Number",
    "Branch",
    "LO_No_MOD_no",
    "Site_Name_",
    "CARD_NO",
    "Main_CARD_SR_No",
    "COP_CARD_SR_No",
    "LPI_CARD_Sr_No",
    "FAS_CARD_Sr_No",
    "Type",
    "Database_Type",
    "Drive_Serial_number",
    "Drive_Barcode",
    "Sim_Status",
    "Reject_Date",
)

SEARCH_FIELDS = (
    "SIM_NO",
    "IMEI_Number",
    "LO_No_MOD_no",
    "CARD_NO",
    "Main_CARD_SR_No",
    "COP_CARD_SR_No",
    "LPI_CARD_Sr_No",
    "FAS_CARD_Sr_No",
)

ROLE_FIELDS = {
    "admin": GSM_COLUMNS,
    "electrical": (
        "Branch",
        "Main_CARD_SR_No",
        "COP_CARD_SR_No",
        "LPI_CARD_Sr_No",
        "FAS_CARD_Sr_No",
        "Drive_Serial_number",
        "Drive_Barcode",
    ),
    "electronics": (
        "IMEI_Number",
        "Branch",
        "LO_No_MOD_no",
        "Site_Name_",
        "CARD_NO",
        "Main_CARD_SR_No",
        "COP_CARD_SR_No",
        "LPI_CARD_Sr_No",
        "FAS_CARD_Sr_No",
        "Type",
        "Database_Type",
        "Sim_Status",
        "Reject_Date",
    ),
}

COMPANY_IMAGE_DIR = os.path.join(app.root_path, "uploads")
COMPANY_IMAGE_NAME = "company-brand"
ALLOWED_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "webp", "bmp", "svg"}
APPEARANCE_PATH = os.path.join(app.root_path, "appearance.json")
MASTER_USERNAME = "master"
GSM_FIELD_CONTROL_TYPES = {"text", "date", "dropdown"}
GSM_FIELD_SETTINGS_DEFAULTS = {
    field: {
        "type": "date" if field in {"Sim_Received_Date", "Reject_Date"} else "text",
        "options": [],
    }
    for field in GSM_COLUMNS
}

# SIM_NO identifies a record and is deliberately not included in any update set.
REPORT_FIELDS = {
    role: (
        tuple(fields)
        if role == "admin"
        else ("SIM_NO",) + tuple(field for field in fields if field != "SIM_NO")
    )
    for role, fields in ROLE_FIELDS.items()
}


def row_to_gsm_record(row, columns=GSM_COLUMNS):
    """Convert a pyodbc row to the JSON shape used by the GSM page."""
    record = {}
    for column, value in zip(columns, row):
        if value is None:
            record[column] = ""
        elif column in {"Sim_Received_Date", "Reject_Date"}:
            if isinstance(value, (datetime, date)):
                record[column] = value.strftime("%Y-%m-%d")
            else:
                text = str(value).strip()
                if text:
                    for format_string in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
                        try:
                            record[column] = datetime.strptime(
                                text[:10], format_string
                            ).strftime("%Y-%m-%d")
                            break
                        except ValueError:
                            record[column] = text[:10]
                else:
                    record[column] = ""
        elif column == "SIM_NO":
            record[column] = str(value).strip()
        elif hasattr(value, "isoformat"):
            record[column] = value.isoformat()
        else:
            record[column] = str(value)
    return record


def request_sim_no(data):
    """Read the SIM number while accepting the existing UI and schema names."""
    value = data.get("SIM_NO", data.get("Sim_No", ""))
    return str(value).strip()


def current_role():
    return str(session.get("role", "")).strip().lower()


def role_for_login(username, database_role):
    username_role = str(username).strip().lower()
    if username_role == MASTER_USERNAME:
        return "admin"
    if username_role in {"admin", "electrical", "electronics"}:
        return username_role
    return str(database_role).strip().lower()


def is_master_user():
    return str(session.get("username", "")).strip().lower() == MASTER_USERNAME


def fields_for_role(role=None):
    return ROLE_FIELDS.get(role or current_role(), ())


def company_image_path():
    for extension in ALLOWED_IMAGE_EXTENSIONS:
        path = os.path.join(COMPANY_IMAGE_DIR, "{}.{}".format(COMPANY_IMAGE_NAME, extension))
        if os.path.isfile(path):
            return path
    return None


def _load_appearance_document():
    try:
        with open(APPEARANCE_PATH, encoding="utf-8") as appearance_file:
            saved = json.load(appearance_file)
        return saved if isinstance(saved, dict) else {}
    except FileNotFoundError:
        return {}
    except (OSError, ValueError):
        app.logger.exception("Unable to read appearance settings.")
        return {}


def load_appearance():
    defaults = {"field_color": "#253858", "field_background": "#ffffff", "font_family": "Poppins", "font_size": "13px"}
    saved = _load_appearance_document()
    return {**defaults, **{key: str(value) for key, value in saved.items() if key in defaults}}


def load_field_settings():
    """Load safe control types and dropdown options for GSM fields."""
    settings = {
        field: {"type": config["type"], "options": list(config["options"])}
        for field, config in GSM_FIELD_SETTINGS_DEFAULTS.items()
    }
    saved = _load_appearance_document().get("field_settings", {})
    if not isinstance(saved, dict):
        return settings
    for field, config in saved.items():
        if field not in settings or not isinstance(config, dict):
            continue
        control_type = str(config.get("type", settings[field]["type"])).lower()
        if control_type not in GSM_FIELD_CONTROL_TYPES:
            control_type = settings[field]["type"]
        options = config.get("options", [])
        if not isinstance(options, list):
            options = []
        settings[field] = {
            "type": control_type,
            "options": [str(option).strip() for option in options if str(option).strip()][:100],
        }
    return settings


def report_fields_for_role(role=None):
    return REPORT_FIELDS.get(role or current_role(), ())


def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if "username" not in session:
            return redirect("/login")
        return view(*args, **kwargs)

    return wrapped_view


def role_required(*roles):
    allowed_roles = {role.lower() for role in roles}

    def decorator(view):
        @wraps(view)
        def wrapped_view(*args, **kwargs):
            if "username" not in session:
                return jsonify({"success": False, "message": "Authentication required."}), 401
            if current_role() not in allowed_roles:
                return jsonify({"success": False, "message": "You are not authorized for this action."}), 403
            return view(*args, **kwargs)

        return wrapped_view

    return decorator


def editable_update_data(data, role):
    """Return only fields the role may update, rejecting attempted changes."""
    allowed = set(fields_for_role(role))
    supplied = set(data) - {"SIM_NO"}
    unauthorized = supplied - allowed
    if unauthorized:
        return None, unauthorized
    return {field: data.get(field) for field in allowed if field in data}, set()


def make_xlsx(columns, rows):
    """Create a dependency-free XLSX workbook for report downloads."""
    def cell(value):
        text = "" if value is None else str(value)
        return '<c t="inlineStr"><is><t>{}</t></is></c>'.format(escape(text))

    sheet_rows = ["<row>{}</row>".format("".join(cell(column) for column in columns))]
    sheet_rows.extend(
        "<row>{}</row>".format("".join(cell(value) for value in row))
        for row in rows
    )
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>{}</sheetData></worksheet>".format("".join(sheet_rows))
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="GSM Report" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
    output.seek(0)
    return output


UPLOAD_COLUMNS = GSM_COLUMNS
UPLOAD_MAX_BYTES = 10 * 1024 * 1024
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _xlsx_cell_value(cell, shared_strings):
    """Return the displayed value for a cell in a small, dependency-free XLSX reader."""
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(
            text.text or ""
            for text in cell.iter()
            if text.tag.rsplit("}", 1)[-1] == "t"
        )

    value = next(
        (child.text for child in cell if child.tag.rsplit("}", 1)[-1] == "v"),
        None,
    )
    if value is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError) as exc:
            raise ValueError("The workbook contains an invalid shared string.") from exc
    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"
    return value


def read_upload_xlsx(file_storage):
    """Read the first worksheet, preserving typed Excel date cells."""
    raw = file_storage.read(UPLOAD_MAX_BYTES + 1)
    if len(raw) > UPLOAD_MAX_BYTES:
        raise ValueError("The uploaded file is too large (maximum 10 MB).")

    try:
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    except (InvalidFileException, OSError, ValueError, zipfile.BadZipFile) as exc:
        raise ValueError("Upload a valid .xlsx Excel workbook.") from exc

    try:
        if not workbook.worksheets:
            raise ValueError("The workbook does not contain a first worksheet.")
        parsed_rows = []
        for excel_row, cells in enumerate(workbook.worksheets[0].iter_rows(), start=1):
            values = tuple(cell.value for cell in cells)
            while values and values[-1] is None:
                values = values[:-1]
            parsed_rows.append((excel_row, values))
    finally:
        workbook.close()

    if not parsed_rows:
        raise ValueError("The workbook must contain a header row.")
    return parsed_rows[0][1], parsed_rows[1:]


def _normalise_upload_value(value):
    return "" if value is None else str(value).strip()


def _parse_upload_date(value):
    """Convert Excel and text date values to a value pyodbc binds as SQL DATE."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= value <= 2_958_465:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()

    text = str(value).strip()
    if not text:
        return None
    for format_string in (
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%d-%b-%Y", "%d %b %Y", "%d-%B-%Y", "%d %B %Y",
    ):
        try:
            return datetime.strptime(text, format_string).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise ValueError(
            "must be an Excel date or a date such as YYYY-MM-DD, DD/MM/YYYY, or MM/DD/YYYY"
        ) from exc


def _prepare_upload_row(row):
    """Normalise text fields and preserve Sim_Received_Date as a Python date."""
    if len(row) > len(UPLOAD_COLUMNS):
        raise ValueError("Too many columns; use the downloaded template.")
    values = list(row) + [""] * (len(UPLOAD_COLUMNS) - len(row))
    date_index = UPLOAD_COLUMNS.index("Sim_Received_Date")
    values[date_index] = _parse_upload_date(values[date_index])
    return tuple(
        value if index == date_index else _normalise_upload_value(value)
        for index, value in enumerate(values)
    )


def _assert_sim_received_date_is_date_column(cursor):
    """Prevent dates being silently stored in a text-typed SQL column."""
    cursor.execute(
        """
        SELECT DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo'
          AND TABLE_NAME = 'revisedGSM'
          AND COLUMN_NAME = 'Sim_Received_Date'
        """
    )
    row = cursor.fetchone()
    data_type = str(row[0]).strip().lower() if row else ""
    if data_type not in {"date", "datetime", "datetime2"}:
        raise ValueError(
            "dbo.revisedGSM.Sim_Received_Date must be DATE, DATETIME, or DATETIME2. "
            "Convert the existing column before uploading."
        )


def _upload_report(inserted, skipped_existing, skipped_uploaded, invalid_rows):
    skipped = skipped_existing + skipped_uploaded
    invalid = len(invalid_rows)
    return {
        "inserted": inserted,
        "skipped": skipped,
        "invalid": invalid,
        "skipped_existing": skipped_existing,
        "skipped_uploaded": skipped_uploaded,
        "invalid_rows": invalid_rows,
    }


def _existing_sim_numbers(cursor, sim_numbers):
    """Find existing SIM_NO values in batches below SQL Server's parameter limit."""
    existing = set()
    for start in range(0, len(sim_numbers), 1000):
        batch = sim_numbers[start:start + 1000]
        placeholders = ", ".join("?" for _ in batch)
        cursor.execute(
            "SELECT SIM_NO FROM dbo.revisedGSM WHERE SIM_NO IN ({})".format(
                placeholders
            ),
            tuple(batch),
        )
        for row in cursor.fetchall():
            value = row[0] if not isinstance(row, str) else row
            existing.add(_normalise_upload_value(value).casefold())
    return existing


def process_gsm_upload(file_storage):
    """Validate and insert an upload, returning a count-rich report."""
    header, numbered_rows = read_upload_xlsx(file_storage)
    actual_header = tuple(_normalise_upload_value(value) for value in header)
    expected_header = tuple(UPLOAD_COLUMNS)
    if actual_header != expected_header:
        raise ValueError(
            "Invalid headers. The first row must contain the dbo.revisedGSM "
            "columns in the required order."
        )

    valid_rows = []
    invalid_rows = []
    seen_uploaded = set()
    skipped_uploaded = 0
    for excel_row, row in numbered_rows:
        try:
            values = _prepare_upload_row(row)
        except ValueError as exc:
            reason = str(exc)
            if not reason.startswith("Too many columns"):
                reason = "Sim_Received_Date {}".format(reason)
            invalid_rows.append({"row": excel_row, "reason": reason})
            continue
        if not any(values):
            continue
        sim_no = values[UPLOAD_COLUMNS.index("SIM_NO")]
        if not sim_no:
            invalid_rows.append({"row": excel_row, "reason": "SIM_NO is required."})
            continue

        sim_key = sim_no.casefold()
        if sim_key in seen_uploaded:
            skipped_uploaded += 1
            continue
        seen_uploaded.add(sim_key)
        valid_rows.append(values)

    if not valid_rows:
        return _upload_report(0, 0, skipped_uploaded, invalid_rows)

    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        _assert_sim_received_date_is_date_column(cursor)
        existing = _existing_sim_numbers(
            cursor,
            [row[UPLOAD_COLUMNS.index("SIM_NO")] for row in valid_rows],
        )
        rows_to_insert = [
            row for row in valid_rows
            if row[UPLOAD_COLUMNS.index("SIM_NO")].casefold() not in existing
        ]
        skipped_existing = len(valid_rows) - len(rows_to_insert)

        insert_sql = "INSERT INTO dbo.revisedGSM ({}) VALUES ({})".format(
            ", ".join(UPLOAD_COLUMNS),
            ", ".join("?" for _ in UPLOAD_COLUMNS),
        )
        for row in rows_to_insert:
            cursor.execute(
                insert_sql,
                tuple(value if value != "" else None for value in row),
            )
        conn.commit()
        cursor.close()
        conn.close()
        return _upload_report(
            len(rows_to_insert),
            skipped_existing,
            skipped_uploaded,
            invalid_rows,
        )
    except Exception:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        raise


# ============================================================
# HOME
# ============================================================

@app.route("/")
def index():
    return redirect("/dashboard" if "username" in session else "/login")

@app.route("/healthz")
def health_check():
    """Lightweight health endpoint for the hosting platform."""
    return jsonify({"status": "ok"}), 200


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = None
        database_error = False
        try:
            conn = get_db_connection()
            row = conn.cursor().execute(
                """
                SELECT User_Name, User_Password, Branch_Name, Role
                FROM dbo.userlogin
                WHERE LOWER(LTRIM(RTRIM(User_Name))) = LOWER(?)
                  AND User_Password = ?
                """,
                (username, password),
            ).fetchone()
        except pyodbc.Error:
            app.logger.exception("Login database query failed")
            row = None
            database_error = True
            flash("Unable to verify login right now.")
        finally:
            if conn:
                conn.close()

        role = role_for_login(username, row.Role) if row else ""
        if row and role in ROLE_FIELDS:
            session["username"] = row.User_Name
            session["branch"] = row.Branch_Name
            session["role"] = role
            session.permanent = request.form.get("remember") == "on"
            return redirect("/dashboard")

        if not row and not database_error:
            flash("Invalid username or password.")

    return render_template("login.html")


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/dashboard")
@login_required
def dashboard():

    return render_template("dashboard.html")


@app.context_processor
def branding_context():
    image_path = company_image_path()
    return {
        "company_image_url": "/company-image" if image_path else None,
    }


# ============================================================
# GSM MASTER PAGE
# ============================================================

@app.route("/GSM")
@app.route("/gsm")
@login_required
def gsmmaster():
    role = current_role()
    return render_template(
        "gsm.html",
        role=role,
        appearance=load_appearance(),
        field_settings=load_field_settings(),
        gsm_columns=list(GSM_COLUMNS),
        editable_fields=list(fields_for_role(role)),
        can_save=role == "admin" and not is_master_user(),
        can_delete=role == "admin" and not is_master_user(),
        is_master=is_master_user(),
        can_search=role in {"admin", "electrical", "electronics"},
    )


@app.route("/company-image")
@login_required
def company_image():
    image_path = company_image_path()
    if not image_path:
        return jsonify({"success": False, "message": "No company image uploaded."}), 404
    return send_file(image_path)


# ============================================================
# SQL CONNECTION TEST
# ============================================================

@app.route("/gsm/test-connection")
@role_required("admin")
def test_connection():

    conn = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        # Test database
        cursor.execute("SELECT DB_NAME()")

        database_name = cursor.fetchone()[0]

        # Test table
        cursor.execute("""
            SELECT COUNT(*)
            FROM dbo.revisedGSM
        """)

        record_count = cursor.fetchone()[0]

        cursor.close()
        conn.close()

        return jsonify({
            "success": True,
            "message": "SQL Server connection successful.",
            "database": database_name,
            "table": "dbo.revisedGSM",
            "records": record_count
        })

    except Exception:

        print("")
        print("========================================")
        print("SQL CONNECTION ERROR")
        print("========================================")
        app.logger.exception("SQL connection test failed")
        print("========================================")
        print("")

        if conn:

            try:
                conn.close()
            except:
                pass

        return jsonify({
            "success": False,
            "message": "Unable to connect to SQL Server."
        }), 500


# ============================================================
# SEARCH GSM RECORD
# ============================================================

@app.route("/gsm/search", methods=["GET"])
@role_required("admin", "electrical", "electronics")
def search_gsm():
    search_field = request.args.get("field", "SIM_NO").strip()
    search_value = request.args.get("value", request.args.get("sim_no", "")).strip()

    if search_field not in SEARCH_FIELDS:
        return jsonify({
            "success": False,
            "message": "That search field is not supported."
        }), 400

    if not search_value:

        return jsonify({
            "success": False,
            "message": "A search value is required."
        }), 400

    if search_field == "LO_No_MOD_no":
        search_value = re.sub(r"\D", "", search_value)
        if not search_value:
            return jsonify({
                "success": False,
                "message": "Enter numbers only for LO No / MOD No search."
            }), 400

    conn = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        # Search returns the complete record; role permissions are applied to
        # editing, not to viewing the record.
        columns = GSM_COLUMNS
        if search_field in {"SIM_NO", "LO_No_MOD_no"}:
            sql = (
                "SELECT {} FROM dbo.revisedGSM "
                "WHERE LTRIM(RTRIM(REPLACE({}, '''', ''))) LIKE ?"
            ).format(", ".join(columns), search_field)
            query_value = (
                "%{}".format(search_value)
                if search_field == "SIM_NO"
                else "%{}%".format(search_value)
            )
        else:
            sql = (
                "SELECT {} FROM dbo.revisedGSM "
                "WHERE LTRIM(RTRIM(REPLACE({}, '''', ''))) = ?"
            ).format(", ".join(columns), search_field)
            query_value = search_value

        cursor.execute(sql, (query_value,))

        rows = cursor.fetchall()

        if not rows:
            cursor.close()
            conn.close()

            return jsonify({
                "success": False,
                "message": "No matching GSM records found."
            }), 404

        records = [row_to_gsm_record(row, columns) for row in rows]
        if current_role() != "admin":
            for record in records:
                record["Mobile_No"] = ""

        cursor.close()
        conn.close()

        return jsonify({
            "success": True,
            "data": records,
            "count": len(records)
        })

    except Exception:

        print("")
        print("========================================")
        print("GSM SEARCH ERROR")
        print("========================================")
        traceback.print_exc()
        print("========================================")
        print("")

        if conn:

            try:
                conn.close()
            except:
                pass

        return jsonify({
            "success": False,
            "message": "Database error while searching."
        }), 500


# ============================================================
# GET ALL GSM RECORDS
# ============================================================

@app.route("/gsm/all", methods=["GET"])
@role_required("admin", "electrical", "electronics")
def get_all_gsm():

    conn = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        columns = GSM_COLUMNS
        cursor.execute(
            "SELECT {} FROM dbo.revisedGSM ORDER BY SIM_NO".format(
                ", ".join(columns)
            )
        )

        rows = cursor.fetchall()

        records = [row_to_gsm_record(row, columns) for row in rows]
        if current_role() != "admin":
            for record in records:
                record["Mobile_No"] = ""

        cursor.close()
        conn.close()

        return jsonify({
            "success": True,
            "data": records,
            "count": len(records)
        })

    except Exception:

        print("")
        print("========================================")
        print("GSM ALL RECORDS ERROR")
        print("========================================")
        traceback.print_exc()
        print("========================================")
        print("")

        if conn:

            try:
                conn.close()
            except:
                pass

        return jsonify({
            "success": False,
            "message": "Unable to fetch GSM records."
        }), 500


# ============================================================
# SAVE GSM RECORD
# ============================================================

@app.route("/gsm/save", methods=["POST"])
@role_required("admin")
def save_gsm():

    if is_master_user():
        return jsonify({
            "success": False,
            "message": "Save access is disabled for the master user."
        }), 403

    data = request.get_json(silent=True)

    if not data:

        return jsonify({
            "success": False,
            "message": "No data received."
        }), 400

    sim_no = request_sim_no(data)

    if not sim_no:

        return jsonify({
            "success": False,
            "message": "SIM No. is required."
        }), 400

    conn = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        # Check duplicate SIM
        cursor.execute("""
            SELECT COUNT(*)
            FROM dbo.revisedGSM
            WHERE SIM_NO = ?
        """, (sim_no,))

        count = cursor.fetchone()[0]

        if count > 0:

            cursor.close()
            conn.close()

            return jsonify({
                "success": True,
                "inserted": False,
                "message": "A GSM record with this SIM No. already exists; the duplicate was skipped."
            })

        sql = """
            INSERT INTO dbo.revisedGSM
            (
                Sim_Received_Date,
                SIM_NO,
                Mobile_No,
                IMEI_Number,
                Branch,
                LO_No_MOD_no,
                Site_Name_,
                CARD_NO,
                Main_CARD_SR_No,
                COP_CARD_SR_No,
                LPI_CARD_Sr_No,
                FAS_CARD_Sr_No,
                Type,
                Database_Type,
                Drive_Serial_number,
                Drive_Barcode,
                Sim_Status,
                Reject_Date
            )
            VALUES
            (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?, ?
            )
        """

        values = (
            data.get("Sim_Received_Date") or None,
            sim_no,
            data.get("Mobile_No") or None,
            data.get("IMEI_Number") or None,
            data.get("Branch") or None,
            data.get("LO_No_MOD_no") or None,
            data.get("Site_Name_") or None,
            data.get("CARD_NO") or None,
            data.get("Main_CARD_SR_No") or None,
            data.get("COP_CARD_SR_No") or None,
            data.get("LPI_CARD_Sr_No") or None,
            data.get("FAS_CARD_Sr_No") or None,
            data.get("Type") or None,
            data.get("Database_Type") or None,
            data.get("Drive_Serial_number") or None,
            data.get("Drive_Barcode") or None,
            data.get("Sim_Status") or None,
            data.get("Reject_Date") or None
        )

        cursor.execute(sql, values)

        conn.commit()

        cursor.close()
        conn.close()

        return jsonify({
            "success": True,
            "message": "GSM record saved successfully."
        })

    except Exception:

        print("")
        print("========================================")
        print("GSM SAVE ERROR")
        print("========================================")
        traceback.print_exc()
        print("========================================")
        print("")

        if conn:

            try:
                conn.rollback()
                conn.close()
            except:
                pass

        return jsonify({
            "success": False,
            "message": "Unable to save GSM record."
        }), 500


# ============================================================
# UPDATE GSM RECORD
# ============================================================

@app.route("/gsm/update", methods=["PUT"])
@role_required("admin", "electrical", "electronics")
def update_gsm():

    if is_master_user():
        return jsonify({
            "success": False,
            "message": "Update access is disabled for the master user."
        }), 403

    data = request.get_json(silent=True)

    if not data:

        return jsonify({
            "success": False,
            "message": "No data received."
        }), 400

    sim_no = request_sim_no(data)

    if not sim_no:

        return jsonify({
            "success": False,
            "message": "SIM No. is required."
        }), 400

    role = current_role()
    update_data, unauthorized = editable_update_data(data, role)
    if unauthorized:
        return jsonify({
            "success": False,
            "message": "Fields not allowed for your role: {}.".format(
                ", ".join(sorted(unauthorized))
            )
        }), 403
    if not update_data:
        return jsonify({
            "success": False,
            "message": "At least one permitted field is required."
        }), 400

    conn = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        update_fields = [field for field in fields_for_role(role) if field in update_data]

        # Do not block status/date edits because an existing LO number happens
        # to be shared. Validate a conflict only when this update changes LO.
        lo_no_mod_no = str(update_data.get("LO_No_MOD_no") or "").strip()
        if lo_no_mod_no and "LO_No_MOD_no" in update_fields:
            cursor.execute(
                "SELECT TOP 1 LO_No_MOD_no FROM dbo.revisedGSM WHERE SIM_NO = ?",
                (sim_no,),
            )
            current_record = cursor.fetchone()
            if not current_record:
                cursor.close()
                conn.close()
                return jsonify({
                    "success": False,
                    "message": "SIM No. not found."
                }), 404

            current_lo_no_mod_no = str(current_record[0] or "").strip()
            if lo_no_mod_no != current_lo_no_mod_no:
                cursor.execute(
                    """
                    SELECT TOP 1 SIM_NO, LO_No_MOD_no
                    FROM dbo.revisedGSM
                    WHERE LTRIM(RTRIM(REPLACE(LO_No_MOD_no, '''', ''))) = ?
                      AND LTRIM(RTRIM(REPLACE(SIM_NO, '''', ''))) <> ?
                    """,
                    (lo_no_mod_no, sim_no),
                )
                duplicate = cursor.fetchone()
                if duplicate:
                    cursor.close()
                    conn.close()
                    return jsonify({
                        "success": False,
                        "message": (
                            "SIM_NO {} + LO_No_MOD_no {} already has an entry "
                            "in the system."
                        ).format(str(duplicate[0]).strip(), str(duplicate[1]).strip())
                    }), 409

        sql = "UPDATE dbo.revisedGSM SET {} WHERE SIM_NO = ?".format(
            ", ".join("{} = ?".format(field) for field in update_fields)
        )
        values = tuple(update_data[field] or None for field in update_fields) + (sim_no,)

        cursor.execute(sql, values)

        affected = cursor.rowcount

        conn.commit()

        cursor.close()
        conn.close()

        if affected == 0:

            return jsonify({
                "success": False,
                "message": "SIM No. not found."
            }), 404

        return jsonify({
            "success": True,
            "message": "GSM record updated successfully."
        })

    except Exception:

        print("")
        print("========================================")
        print("GSM UPDATE ERROR")
        print("========================================")
        traceback.print_exc()
        print("========================================")
        print("")

        if conn:

            try:
                conn.rollback()
                conn.close()
            except:
                pass

        return jsonify({
            "success": False,
            "message": "Unable to update GSM record."
        }), 500


# ============================================================
# DELETE GSM RECORD
# ============================================================

@app.route("/gsm/delete", methods=["DELETE"])
@role_required("admin")
def delete_gsm():

    if is_master_user():
        return jsonify({
            "success": False,
            "message": "Delete access is disabled for the master user."
        }), 403

    data = request.get_json(silent=True)

    sim_no = ""

    if data:

        sim_no = request_sim_no(data)

    if not sim_no:

        return jsonify({
            "success": False,
            "message": "SIM No. is required."
        }), 400

    conn = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        cursor.execute("""
            DELETE FROM dbo.revisedGSM
            WHERE SIM_NO = ?
        """, (sim_no,))

        affected = cursor.rowcount

        conn.commit()

        cursor.close()
        conn.close()

        if affected == 0:

            return jsonify({
                "success": False,
                "message": "SIM No. not found."
            }), 404

        return jsonify({
            "success": True,
            "message": "GSM record deleted successfully."
        })

    except Exception:

        print("")
        print("========================================")
        print("GSM DELETE ERROR")
        print("========================================")
        traceback.print_exc()
        print("========================================")
        print("")

        if conn:

            try:
                conn.rollback()
                conn.close()
            except:
                pass

        return jsonify({
            "success": False,
            "message": "Unable to delete GSM record."
        }), 500


# ============================================================
# DOWNLOAD ROLE-SCOPED EXCEL REPORT
# ============================================================

@app.route("/gsm/report", methods=["GET"])
@app.route("/gsm/download", methods=["GET"])
@app.route("/gsm/report/download", methods=["GET"])
@role_required("admin", "electrical", "electronics")
def download_gsm_report():
    conn = None
    role = current_role()
    columns = report_fields_for_role(role)
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT {} FROM dbo.revisedGSM ORDER BY SIM_NO".format(
                ", ".join(columns)
            )
        )
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        report = make_xlsx(columns, rows)
        return send_file(
            report,
            as_attachment=True,
            download_name="gsm_{}_report.xlsx".format(role),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception:
        app.logger.exception("GSM report generation failed")
        if conn:
            try:
                conn.close()
            except Exception:
                pass
        return jsonify({
            "success": False,
            "message": "Unable to generate GSM report."
        }), 500


# ============================================================
# ADMIN GSM UPLOAD
# ============================================================

@app.route("/gsm/upload", methods=["GET", "POST"])
@role_required("admin")
def upload_gsm():
    if request.method == "GET":
        return render_template("upload.html", columns=UPLOAD_COLUMNS, appearance=load_appearance())

    uploaded_file = request.files.get("file")
    if not uploaded_file or not uploaded_file.filename:
        return jsonify({
            "success": False,
            "message": "Choose an .xlsx file to upload.",
            "inserted": 0,
            "skipped": 0,
            "invalid": 0,
        }), 400
    if not uploaded_file.filename.lower().endswith(".xlsx"):
        return jsonify({
            "success": False,
            "message": "Only .xlsx Excel workbooks are supported.",
            "inserted": 0,
            "skipped": 0,
            "invalid": 0,
        }), 400

    try:
        report = process_gsm_upload(uploaded_file)
    except ValueError as exc:
        return jsonify({
            "success": False,
            "message": str(exc),
            "inserted": 0,
            "skipped": 0,
            "invalid": 0,
        }), 400
    except Exception:
        app.logger.exception("GSM upload failed")
        return jsonify({
            "success": False,
            "message": "Unable to process the GSM upload.",
            "inserted": 0,
            "skipped": 0,
            "invalid": 0,
        }), 500

    return jsonify({
        "success": True,
        "message": (
            "Upload complete: {inserted} inserted, {skipped} skipped, "
            "{invalid} invalid."
        ).format(**report),
        "report": report,
        **report,
    })


@app.route("/admin", methods=["GET"])
@role_required("admin")
def admin_page():
    return render_template(
        "admin.html",
        appearance=load_appearance(),
        field_settings=load_field_settings(),
        gsm_columns=list(GSM_COLUMNS),
    )


@app.route("/admin/company-image", methods=["POST", "DELETE"])
@role_required("admin")
def manage_company_image():
    os.makedirs(COMPANY_IMAGE_DIR, exist_ok=True)
    if request.method == "DELETE":
        image_path = company_image_path()
        if image_path:
            os.remove(image_path)
        return jsonify({"success": True, "message": "Company image deleted."})

    uploaded_file = request.files.get("image")
    filename = (uploaded_file.filename if uploaded_file else "").lower()
    extension = filename.rsplit(".", 1)[-1] if "." in filename else ""
    if not uploaded_file or not uploaded_file.filename or extension not in ALLOWED_IMAGE_EXTENSIONS:
        return jsonify({
            "success": False,
            "message": "Choose a PNG, JPG, JPEG, GIF, WEBP, BMP, or SVG image."
        }), 400
    for existing_extension in ALLOWED_IMAGE_EXTENSIONS:
        existing_path = os.path.join(
            COMPANY_IMAGE_DIR, "{}.{}".format(COMPANY_IMAGE_NAME, existing_extension)
        )
        if os.path.isfile(existing_path):
            os.remove(existing_path)
    uploaded_file.save(os.path.join(COMPANY_IMAGE_DIR, "{}.{}".format(COMPANY_IMAGE_NAME, extension)))
    return jsonify({"success": True, "message": "Company image uploaded."})


@app.route("/admin/appearance", methods=["GET", "POST"])
@role_required("admin")
def admin_appearance():
    if request.method == "GET":
        return jsonify({"success": True, "data": load_appearance()})
    data = request.get_json(silent=True) or {}
    current = load_appearance()
    allowed = {"field_color", "field_background", "font_family", "font_size"}
    updated = {key: str(data[key]) for key in allowed if key in data}
    current.update(updated)
    with open(APPEARANCE_PATH, "w", encoding="utf-8") as appearance_file:
        json.dump(current, appearance_file, indent=2)
    return jsonify({"success": True, "data": current, "message": "Appearance settings saved."})


@app.route("/admin/field-settings", methods=["GET", "POST"])
@role_required("admin")
def admin_field_settings():
    if request.method == "GET":
        return jsonify({"success": True, "data": load_field_settings()})

    data = request.get_json(silent=True) or {}
    incoming = data.get("field_settings")
    if not isinstance(incoming, dict):
        return jsonify({"success": False, "message": "Field settings are required."}), 400

    settings = load_field_settings()
    for field in GSM_COLUMNS:
        config = incoming.get(field)
        if not isinstance(config, dict):
            continue
        control_type = str(config.get("type", settings[field]["type"])).lower()
        if control_type not in GSM_FIELD_CONTROL_TYPES:
            return jsonify({"success": False, "message": "Unsupported control type for {}.".format(field)}), 400
        options = config.get("options", [])
        if not isinstance(options, list):
            return jsonify({"success": False, "message": "Dropdown options for {} must be a list.".format(field)}), 400
        settings[field] = {
            "type": control_type,
            "options": [str(option).strip() for option in options if str(option).strip()][:100],
        }

    saved = _load_appearance_document()
    saved["field_settings"] = settings
    try:
        with open(APPEARANCE_PATH, "w", encoding="utf-8") as appearance_file:
            json.dump(saved, appearance_file, indent=2)
    except OSError:
        app.logger.exception("Unable to save GSM field settings.")
        return jsonify({"success": False, "message": "Unable to save field settings."}), 500
    return jsonify({"success": True, "data": settings, "message": "GSM field settings saved."})


@app.route("/gsm/upload/template", methods=["GET"])
@app.route("/gsm/upload/template.xlsx", methods=["GET"])
@role_required("admin")
def download_upload_template():
    template = make_xlsx(UPLOAD_COLUMNS, [])
    return send_file(
        template,
        as_attachment=True,
        download_name="revisedGSM_upload_template.xlsx",
        mimetype=XLSX_MIMETYPE,
    )


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect("/login")


# ============================================================
# 404 ERROR
# ============================================================

@app.errorhandler(404)
def page_not_found(error):
    del error

    return jsonify({
        "success": False,
        "message": "Page not found."
    }), 404


# ============================================================
# APPLICATION START
# ============================================================

if __name__ == "__main__":

    print("================================")
    print("GSM APPLICATION")
    print("DATABASE: GSMMaster")
    print("TABLE: dbo.revisedGSM")
    print("================================")
    print("http://127.0.0.1:5001")
    print("================================")

    app.run(
        host="127.0.0.1",
        port=5001,
        debug=True
    )
