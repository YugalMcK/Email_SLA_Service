"""
Load Ops Only IB Summary Excel exports into SQLite.
"""

import argparse
import logging
import re
import sqlite3
import sys
import time
from pathlib import Path

import pandas as pd

from config import (
    CREATE_TABLE_SQL,
    EXCEL_COLUMN_MAP,
    EXCEL_REPORT_FOLDER,
    TABLE_COLUMNS,
    get_db_path,
    get_table_name,
)

logger = logging.getLogger("GR_Tableau.db")

DB_CONNECT_TIMEOUT = 60.0
DB_MAX_RETRIES = 5
DB_RETRY_DELAY = 3


def _normalize_header(value: object) -> str:
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def _find_header_row(raw_df: pd.DataFrame) -> int:
    expected = set(EXCEL_COLUMN_MAP.keys()) | set(TABLE_COLUMNS)
    best_row = 0
    best_score = 0
    for row_idx in range(min(15, len(raw_df))):
        row_values = {_normalize_header(v) for v in raw_df.iloc[row_idx].tolist() if str(v).strip()}
        score = len(row_values & expected)
        if score > best_score:
            best_score = score
            best_row = row_idx
    if best_score < 3:
        raise ValueError("Could not detect a valid header row in the Excel export.")
    return best_row


def _prepare_dataframe(excel_path: Path) -> pd.DataFrame:
    raw_df = pd.read_excel(excel_path, header=None, engine="openpyxl")
    header_row = _find_header_row(raw_df)
    headers = [_normalize_header(value) for value in raw_df.iloc[header_row].tolist()]
    data_df = raw_df.iloc[header_row + 1 :].copy()
    data_df.columns = headers
    data_df = data_df.dropna(how="all")

    rename_map = {}
    for column in data_df.columns:
        normalized = _normalize_header(column)
        if normalized in EXCEL_COLUMN_MAP:
            rename_map[column] = EXCEL_COLUMN_MAP[normalized]
        else:
            fallback = normalized.replace(" ", "_").replace(".", "").replace("#", "")
            rename_map[column] = fallback

    data_df = data_df.rename(columns=rename_map)

    for column in TABLE_COLUMNS:
        if column not in data_df.columns:
            data_df[column] = None

    data_df = data_df[TABLE_COLUMNS]
    data_df = data_df.replace({pd.NA: None, "": None})
    data_df = data_df.dropna(how="all")

    for int_col in ("Index", "DC", "BUP_Type"):
        data_df[int_col] = pd.to_numeric(data_df[int_col], errors="coerce").astype("Int64")

    logger.info("Prepared %s rows from '%s'.", len(data_df), excel_path.name)
    return data_df


def _connect_db(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path, timeout=DB_CONNECT_TIMEOUT)
    conn.execute("PRAGMA busy_timeout = 60000")
    try:
        conn.execute("PRAGMA journal_mode = WAL")
    except sqlite3.OperationalError:
        pass
    return conn


def _table_columns(conn: sqlite3.Connection, table_name: str) -> list[str]:
    rows = conn.execute(f'PRAGMA table_info("{table_name}")').fetchall()
    return [row[1] for row in rows]


def _autoincrement_pk_columns(conn: sqlite3.Connection, table_name: str) -> set[str]:
    """Return integer primary-key columns that SQLite auto-populates.

    A single INTEGER PRIMARY KEY is a rowid alias (optionally AUTOINCREMENT), so
    we must NOT insert values into it - let SQLite generate them. PRAGMA row
    format: (cid, name, type, notnull, dflt_value, pk).
    """
    rows = conn.execute(f'PRAGMA table_info("{table_name}")').fetchall()
    pk_cols = [row for row in rows if row[5]]
    if len(pk_cols) == 1 and (pk_cols[0][2] or "").upper() == "INTEGER":
        return {pk_cols[0][1]}
    return set()


def ensure_table(db_path: str, table_name: str) -> None:
    last_error = None
    for attempt in range(1, DB_MAX_RETRIES + 1):
        try:
            conn = _connect_db(db_path)
            try:
                conn.execute(CREATE_TABLE_SQL)
                conn.commit()
                return
            finally:
                conn.close()
        except sqlite3.OperationalError as exc:
            last_error = exc
            if "locked" not in str(exc).lower():
                raise
            logger.warning(
                "Database locked while ensuring table (attempt %s/%s). Retrying.",
                attempt,
                DB_MAX_RETRIES,
            )
            time.sleep(DB_RETRY_DELAY)

    raise sqlite3.OperationalError(f"Database is locked after {DB_MAX_RETRIES} attempts: {last_error}")


def _to_sqlite_value(value):
    """Convert a pandas/numpy scalar to a native Python type SQLite accepts.

    numpy.int64 (from pandas 'Int64' columns) is bound by sqlite3 as a non-native
    type and is rejected by an INTEGER PRIMARY KEY column with 'datatype
    mismatch', so numpy scalars must be converted to plain int/float first.
    """
    try:
        if value is None or pd.isna(value):
            return None
    except (TypeError, ValueError):
        # pd.isna can raise on non-scalar values; fall through and return as-is.
        pass

    # numpy / pandas scalars expose .item() to return the native Python value.
    item = getattr(value, "item", None)
    if callable(item):
        try:
            return value.item()
        except (ValueError, AttributeError):
            pass
    return value


def _insert_dataframe(conn: sqlite3.Connection, df: pd.DataFrame, table_name: str) -> None:
    table_cols = _table_columns(conn, table_name)
    if not table_cols:
        raise RuntimeError(f"Table '{table_name}' does not exist in the database.")

    missing_cols = [col for col in df.columns if col not in table_cols]
    if missing_cols:
        raise RuntimeError(
            f"Database table '{table_name}' is missing columns: {missing_cols}. "
            f"Existing columns: {table_cols}"
        )

    # Skip auto-increment integer PK columns so SQLite assigns them itself.
    auto_pk = _autoincrement_pk_columns(conn, table_name)
    if auto_pk:
        logger.info("Letting SQLite auto-generate primary key column(s): %s", ", ".join(auto_pk))

    columns = [col for col in df.columns if col in table_cols and col not in auto_pk]
    placeholders = ", ".join("?" for _ in columns)
    columns_sql = ", ".join(f'"{col}"' for col in columns)
    insert_sql = f'INSERT INTO "{table_name}" ({columns_sql}) VALUES ({placeholders})'

    rows = [
        tuple(_to_sqlite_value(value) for value in row)
        for row in df[columns].itertuples(index=False, name=None)
    ]
    conn.executemany(insert_sql, rows)


def load_excel_to_db(
    excel_path: Path,
    *,
    db_path: str | None = None,
    table_name: str | None = None,
    replace_existing: bool = True,
) -> int:
    db_path = db_path or get_db_path()
    table_name = table_name or get_table_name()
    ensure_table(db_path, table_name)

    df = _prepare_dataframe(excel_path)
    if df.empty:
        logger.warning("No rows found in '%s'; skipping database load.", excel_path)
        return 0

    last_error = None
    for attempt in range(1, DB_MAX_RETRIES + 1):
        try:
            conn = _connect_db(db_path)
            try:
                if replace_existing:
                    conn.execute(f'DELETE FROM "{table_name}"')
                    logger.info("Cleared existing rows from '%s'.", table_name)
                    # Reset AUTOINCREMENT so the PK restarts at 1 on a full reload.
                    try:
                        conn.execute(
                            "DELETE FROM sqlite_sequence WHERE name = ?", (table_name,)
                        )
                    except sqlite3.OperationalError:
                        # sqlite_sequence only exists when a table uses AUTOINCREMENT.
                        pass

                _insert_dataframe(conn, df, table_name)
                conn.commit()
                logger.info("Loaded %s rows into '%s'.", len(df), table_name)
                return len(df)
            finally:
                conn.close()
        except sqlite3.OperationalError as exc:
            last_error = exc
            if "locked" not in str(exc).lower():
                raise
            logger.warning(
                "Database locked while loading data (attempt %s/%s). Retrying.",
                attempt,
                DB_MAX_RETRIES,
            )
            time.sleep(DB_RETRY_DELAY)

    raise sqlite3.OperationalError(f"Database is locked after {DB_MAX_RETRIES} attempts: {last_error}")


def _latest_excel_in_reports() -> Path | None:
    """Return the most recently modified .xlsx in the excel_reports folder."""
    candidates = sorted(
        EXCEL_REPORT_FOLDER.glob("*.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Load an Ops Only IB Summary Excel export into SQLite (standalone)."
    )
    parser.add_argument(
        "excel_file",
        nargs="?",
        help="Path to the .xlsx file to load. "
        "If omitted, the latest file in the excel_reports folder is used.",
    )
    parser.add_argument("--db-path", help="Override the database path (defaults to .env).")
    parser.add_argument("--table-name", help="Override the table name (defaults to .env).")
    parser.add_argument(
        "--append",
        action="store_true",
        help="Append rows instead of replacing existing table contents.",
    )
    args = parser.parse_args()

    if args.excel_file:
        report_path = Path(args.excel_file)
    else:
        report_path = _latest_excel_in_reports()
        if report_path is None:
            logger.error(
                "No .xlsx file provided and none found in '%s'. "
                "Pass a file path explicitly.",
                EXCEL_REPORT_FOLDER,
            )
            return 1
        logger.info("No file specified; using latest export '%s'.", report_path)

    if not report_path.is_file():
        logger.error("Excel file not found: '%s'.", report_path)
        return 1

    try:
        row_count = load_excel_to_db(
            report_path,
            db_path=args.db_path,
            table_name=args.table_name,
            replace_existing=not args.append,
        )
    except Exception as exc:  # noqa: BLE001 - surface a clean error + exit code
        logger.exception("Failed to load '%s': %s", report_path, exc)
        return 1

    if row_count == 0:
        logger.warning("No rows were loaded from '%s'.", report_path)
        return 0

    logger.info("Done. Loaded %s rows from '%s'.", row_count, report_path)
    return 0


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        stream=sys.stdout,
    )
    sys.exit(main())
