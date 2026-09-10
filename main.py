"""
GR1/GR3 Tableau automation service.

Downloads the Ops Only IB Summary report from McKesson Tableau and loads it into SQLite.
"""

import argparse
import logging
import shutil
import sys
from pathlib import Path

from config import (
    EXCEL_ARCHIVE_FOLDER,
    EXCEL_REPORT_FOLDER,
    LOG_FOLDER,
    NoDataError,
    ensure_directories,
    is_weekend_pacific,
    now_pacific,
    timestamp_label,
)
from db_loader import load_excel_to_db
from tableau_automation import download_tableau_report

# Create the log folder up front so a log file is always written.
LOG_FOLDER.mkdir(parents=True, exist_ok=True)
LOG_FILE_PATH = LOG_FOLDER / f"gr1_gr3_{timestamp_label()}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE_PATH, encoding="utf-8"),
    ],
)
logger = logging.getLogger("GR_Tableau")
logger.info("Logging to file: %s", LOG_FILE_PATH)


class WeekendSkip(Exception):
    """Raised when the service is invoked on a weekend (Pacific time)."""


def run(*, skip_download: bool = False, excel_file: str | None = None, force: bool = False) -> int:
    ensure_directories()

    today = now_pacific()
    if is_weekend_pacific() and not force:
        raise WeekendSkip(
            f"Service does not run on weekends. Today is {today.strftime('%A, %m/%d/%Y')} "
            f"(Pacific time). Use --force to override."
        )

    if skip_download:
        if excel_file:
            report_path = Path(excel_file)
        else:
            candidates = sorted(
                EXCEL_REPORT_FOLDER.glob("*.xlsx"),
                key=lambda path: path.stat().st_mtime,
                reverse=True,
            )
            if not candidates:
                raise FileNotFoundError(
                    f"No Excel files found in '{EXCEL_REPORT_FOLDER}'. Run without --skip-download first."
                )
            report_path = candidates[0]
        logger.info("Skipping download; using '%s'.", report_path)
    else:
        report_path = download_tableau_report()
        logger.info("Report downloaded to '%s'.", report_path)

    row_count = load_excel_to_db(report_path)
    if row_count == 0:
        raise NoDataError(
            f"No data available to download/load from '{report_path.name}'. "
            f"Terminating without updating the database."
        )

    archive_target = EXCEL_ARCHIVE_FOLDER / report_path.name
    shutil.move(str(report_path), archive_target)
    logger.info("Archived report to '%s'.", archive_target)
    logger.info("Completed successfully. Loaded %s rows.", row_count)
    return row_count


def main() -> int:
    parser = argparse.ArgumentParser(description="McKesson Tableau GR1/GR3 automation")
    parser.add_argument(
        "--skip-download",
        action="store_true",
        help="Load the latest existing Excel file without opening Tableau.",
    )
    parser.add_argument(
        "--excel-file",
        help="Specific Excel file to load when --skip-download is set.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Run even on weekends (Saturday/Sunday).",
    )
    args = parser.parse_args()

    try:
        run(skip_download=args.skip_download, excel_file=args.excel_file, force=args.force)
        return 0
    except WeekendSkip as exc:
        logger.info("%s", exc)
        return 0
    except NoDataError as exc:
        logger.warning("%s", exc)
        return 0
    except Exception as exc:
        logger.exception("GR1/GR3 Tableau automation failed: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
