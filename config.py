import os
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

# Pacific time zone (handles PST/PDT daylight saving automatically).
PACIFIC_TZ = ZoneInfo("America/Los_Angeles")


def now_pacific() -> datetime:
    """Return the current time in US Pacific time (PST/PDT)."""
    return datetime.now(PACIFIC_TZ)


def is_weekend_pacific() -> bool:
    """Return True if the current Pacific date is Saturday or Sunday."""
    # Monday=0 ... Saturday=5, Sunday=6
    return now_pacific().weekday() >= 5

class NoDataError(Exception):
    """Raised when the report has no data for the selected date."""


BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"

EXCEL_REPORT_FOLDER = BASE_DIR / "excel_reports"
EXCEL_ARCHIVE_FOLDER = EXCEL_REPORT_FOLDER / "Archive"
LOG_FOLDER = BASE_DIR / "log_files"
AUTH_STATE_PATH = BASE_DIR / ".tableau_auth.json"


def _parse_env_file(env_path: Path) -> dict[str, str]:
    """Parse .env entries written as KEY = value or Key : value."""
    values: dict[str, str] = {}
    if not env_path.is_file():
        return values

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        if "=" in line:
            key, value = line.split("=", 1)
        elif ":" in line:
            key, value = line.split(":", 1)
        else:
            continue

        key = key.strip().lower().replace(" ", "_")
        value = value.strip()
        if value:
            values[key] = value
    return values


_env_values = _parse_env_file(ENV_PATH)


def _env(*keys: str, default: str = "") -> str:
    """Return the first non-empty value from os.environ or the .env file.

    Only UPPER_CASE keys are looked up in os.environ. This is important on
    Windows, where os.environ is case-insensitive: a friendly key like
    'username' would otherwise collide with the Windows 'USERNAME' variable
    (the logged-in employee ID) and silently override the .env value.
    The .env file is checked for every key using its lower-cased form.
    """
    for key in keys:
        if key.isupper():
            val = os.environ.get(key)
            if val and val.strip():
                return val.strip()
    for key in keys:
        val = _env_values.get(key.lower().replace(" ", "_"))
        if val and val.strip():
            return val.strip()
    return default


def _env_int(*keys: str, default: int) -> int:
    raw = _env(*keys, default="")
    try:
        return int(raw) if raw else default
    except ValueError:
        return default


# --- Defaults (used only when the value is absent from .env / environment) ---
DEFAULT_LANDING_URL = (
    "https://viz.mckesson.com/#/site/SupplyChainAnalytics/views/"
    "DSCSAAnalyticsandMonitoringLandingPage_Old/"
    "DSCSAAnalyticsandMonitoringLandingPage?:iid=1&:linktarget=_blank"
)

DEFAULT_DB_PATH = r"D:\Yugal\Incident_SLA_Service_Auto\mckesson_db.db"
DEFAULT_TABLE_NAME = "gr1_gr3_data"
DEFAULT_LANDING_REPORT_TITLE = "DSCSA Ops Only IB Summary and Detail Report"
DEFAULT_SEARCH_KEYWORD = "Ops only"
DEFAULT_VIEW_LABEL = "Ops Only IB Summary Report"
DEFAULT_BROWSER = "msedge"

# --- Timeouts (milliseconds); overridable via .env ---
LOGIN_TIMEOUT_MS = _env_int("TABLEAU_LOGIN_TIMEOUT_MS", "login_timeout_ms", default=5 * 60 * 1000)
PAGE_LOAD_TIMEOUT_MS = _env_int("TABLEAU_PAGE_LOAD_TIMEOUT_MS", "page_load_timeout_ms", default=120_000)
DOWNLOAD_TIMEOUT_MS = _env_int("TABLEAU_DOWNLOAD_TIMEOUT_MS", "download_timeout_ms", default=180_000)
DATA_REFRESH_TIMEOUT_MS = _env_int(
    "TABLEAU_DATA_REFRESH_TIMEOUT_MS", "data_refresh_timeout_ms", default=10 * 60 * 1000
)

TABLE_COLUMNS = [
    "Index",
    "BU",
    "Name",
    "Instance_ID",
    "MSGUID",
    "Message",
    "DC",
    "DC_Name",
    "Sender_GLN",
    "Receiving_Location",
    "Receiving_Owning_Party",
    "Dropship_Flag",
    "Creation_Date",
    "Creation_Time",
    "Time_Intervel",
    "Created_BY",
    "Process_Start_TS",
    "Process_End_TS",
    "Processing_Time",
    "SLA_Flag",
    "BUP_Type",
    "LOT_Number",
    "PO_Number",
    "Dispatch_Advice",
    "Success",
    "Processed_First_Pass",
    "Reprocess_Time",
    "MSG_No",
    "Message_Type",
]

EXCEL_COLUMN_MAP = {
    "Index": "Index",
    "Sr No": "Index",
    "Sr. No": "Index",
    "Sr No.": "Index",
    "Sr. No.": "Index",
    "BU": "BU",
    "Name": "Name",
    "Instance ID": "Instance_ID",
    "Instance_ID": "Instance_ID",
    "MSGUID": "MSGUID",
    "Message": "Message",
    "DC": "DC",
    "DC Name": "DC_Name",
    "DC_Name": "DC_Name",
    "Sender GLN": "Sender_GLN",
    "Sender_GLN": "Sender_GLN",
    "Receiving Location": "Receiving_Location",
    "Receiving_Location": "Receiving_Location",
    "Receiving Owning Party": "Receiving_Owning_Party",
    "Receiving_Owning_Party": "Receiving_Owning_Party",
    "Dropship Flag": "Dropship_Flag",
    "Dropship_Flag": "Dropship_Flag",
    "Creation Date": "Creation_Date",
    "Creation_Date": "Creation_Date",
    "Creation Time": "Creation_Time",
    "Creation_Time": "Creation_Time",
    "Time Intervel": "Time_Intervel",
    "Time Interval": "Time_Intervel",
    "Time_Intervel": "Time_Intervel",
    "Created BY": "Created_BY",
    "Created By": "Created_BY",
    "Created_BY": "Created_BY",
    "Process Start TS": "Process_Start_TS",
    "Process_Start_TS": "Process_Start_TS",
    "Process End TS": "Process_End_TS",
    "Process_End_TS": "Process_End_TS",
    "Processing Time": "Processing_Time",
    "Processing_Time": "Processing_Time",
    "SLA Flag": "SLA_Flag",
    "SLA_Flag": "SLA_Flag",
    "BUP Type": "BUP_Type",
    "BUP_Type": "BUP_Type",
    "LOT Number": "LOT_Number",
    "LOT_Number": "LOT_Number",
    "PO Number": "PO_Number",
    "PO_Number": "PO_Number",
    "Dispatch Advice": "Dispatch_Advice",
    "Dispatch_Advice": "Dispatch_Advice",
    "Success": "Success",
    "Processed First Pass": "Processed_First_Pass",
    "Processed_First_Pass": "Processed_First_Pass",
    "Reprocess Time": "Reprocess_Time",
    "Reprocess_Time": "Reprocess_Time",
    "MSG No #": "MSG_No",
    "MSG No. #": "MSG_No",
    "MSG_No_#": "MSG_No",
    "MSG_No": "MSG_No",
    "Message Type": "Message_Type",
    "Message_Type": "Message_Type",
}

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS gr1_gr3_data (
    "Index" INTEGER PRIMARY KEY AUTOINCREMENT,
    BU TEXT,
    Name TEXT,
    Instance_ID TEXT,
    MSGUID TEXT,
    Message TEXT,
    DC INTEGER,
    DC_Name TEXT,
    Sender_GLN TEXT,
    Receiving_Location TEXT,
    Receiving_Owning_Party TEXT,
    Dropship_Flag TEXT,
    Creation_Date TEXT,
    Creation_Time TEXT,
    Time_Intervel TEXT,
    Created_BY TEXT,
    Process_Start_TS TEXT,
    Process_End_TS TEXT,
    Processing_Time TEXT,
    SLA_Flag TEXT,
    BUP_Type INTEGER,
    LOT_Number TEXT,
    PO_Number TEXT,
    Dispatch_Advice TEXT,
    Success TEXT,
    Processed_First_Pass TEXT,
    Reprocess_Time TEXT,
    MSG_No TEXT,
    Message_Type TEXT
);
"""


def _normalize_date(value: str) -> str:
    """Return a Tableau-friendly M/D/YYYY date string."""
    value = value.strip()
    if not value:
        return now_pacific().strftime("%-m/%-d/%Y") if os.name != "nt" else now_pacific().strftime("%#m/%#d/%Y")

    for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed = datetime.strptime(value, fmt)
            if os.name == "nt":
                return parsed.strftime("%#m/%#d/%Y")
            return parsed.strftime("%-m/%-d/%Y")
        except ValueError:
            continue
    return value


def get_tableau_username() -> str:
    """Return the required Tableau/Okta username."""
    # NOTE: do NOT fall back to the OS 'USERNAME' env var - on Windows that is
    # the Windows login (e.g. an employee ID) and would override the .env value.
    username = _env("TABLEAU_USERNAME", "username")
    if not username:
        raise ValueError(
            "Username is required. Set Username in .env or TABLEAU_USERNAME in the environment."
        )
    return username


def get_tableau_password() -> str:
    """Return an optional password for service-account login. Empty means Okta SSO."""
    return _env("TABLEAU_PASSWORD", "PASSWORD", "password").rstrip(";")


def get_tableau_credentials() -> tuple[str, str]:
    return get_tableau_username(), get_tableau_password()


def get_report_url() -> str:
    return _env("TABLEAU_REPORT_URL", "report_url", default=DEFAULT_LANDING_URL)


def get_search_keyword() -> str:
    return _env("TABLEAU_SEARCH_KEYWORD", "search_keyword", default=DEFAULT_SEARCH_KEYWORD)


def get_landing_report_title() -> str:
    return _env(
        "TABLEAU_REPORT_TITLE",
        "report_title",
        "landing_report_title",
        default=DEFAULT_LANDING_REPORT_TITLE,
    )


def get_view_label() -> str:
    return _env("TABLEAU_VIEW_LABEL", "view_label", default=DEFAULT_VIEW_LABEL)


def get_browser_channels() -> tuple:
    """Return an ordered tuple of browser channels to try launching.

    ``None`` means bundled Chromium. Configurable via the .env 'browser' key
    (msedge / chrome / chromium).
    """
    choice = _env("TABLEAU_BROWSER", "browser", default=DEFAULT_BROWSER).lower()
    if choice in ("chrome", "google chrome"):
        return ("chrome", "msedge", None)
    if choice in ("chromium", "default", "bundled"):
        return (None,)
    # Default / msedge / edge
    return ("msedge", "chrome", None)


def get_browser_executable() -> str:
    """Return an explicit browser executable path (e.g. a portable Chrome), or ''.

    When set (via .env 'chrome_path' / 'browser_path'), it takes precedence over
    the browser channel and is launched directly.
    """
    return _env("TABLEAU_BROWSER_PATH", "chrome_path", "browser_path", default="")


def get_date_range() -> tuple[str, str]:
    today = now_pacific()
    if os.name == "nt":
        default_date = today.strftime("%#m/%#d/%Y")
    else:
        default_date = today.strftime("%-m/%-d/%Y")

    from_date = _env("TABLEAU_FROM_DATE", "from_date", default=default_date)
    to_date = _env("TABLEAU_TO_DATE", "to_date", default=default_date)
    return _normalize_date(from_date), _normalize_date(to_date)


def get_db_path() -> str:
    return _env("TABLEAU_DB_PATH", "database_name", "database", default=DEFAULT_DB_PATH)


def get_table_name() -> str:
    return _env("TABLEAU_TABLE_NAME", "table_name", default=DEFAULT_TABLE_NAME)


def timestamp_label() -> str:
    return now_pacific().strftime("%d%m%Y_%H%M%S")


def ensure_directories() -> None:
    EXCEL_REPORT_FOLDER.mkdir(parents=True, exist_ok=True)
    EXCEL_ARCHIVE_FOLDER.mkdir(parents=True, exist_ok=True)
    LOG_FOLDER.mkdir(parents=True, exist_ok=True)
