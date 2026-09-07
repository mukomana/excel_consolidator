"""
main.py

Main entry point for the Weekly Report Automation project.

Author: Freedom Mukomana
"""

from pathlib import Path
import json
import logging
import time
import argparse

from report_selector import select_latest_report
from consolidator import WeeklyReportConsolidator
from pathlib import Path
from settings import (
    DATA_SOURCE,
    LOCAL_REPORTS_FOLDER
)
from settings import (
    PROCESS_LATEST_ONLY,
    FILTER_TRIMESTER,
    FILTER_WEEK,
    MANUAL_REPORTS
)

##############################################################################
# CONFIGURATION
##############################################################################

ROOT = Path(__file__).resolve().parent.parent

CONFIG_DIR = ROOT / "config"
OUTPUT_DIR = ROOT / "output"
ARCHIVE_DIR = ROOT / "archive"
LOG_DIR = ROOT / "logs"

OUTPUT_DIR.mkdir(exist_ok=True)
ARCHIVE_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

CONFIG_FILE = CONFIG_DIR / "nodes.json"

##############################################################################
# LOGGING
##############################################################################

logging.basicConfig(
    filename=LOG_DIR / "automation.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger(__name__)

##############################################################################
# LOAD CONFIGURATION
##############################################################################

def load_sites():

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

##############################################################################
# DOWNLOAD REPORTS
##############################################################################

def download_reports():

    from auth import get_context
    from sharepoint import SharePointManager

    reports = []

    sites = load_sites()

    for site in sites:

        logger.info(f"Processing {site['name']}")

        try:

            context = get_context(site["site"])

            sp = SharePointManager(context)

            local_folder = ARCHIVE_DIR / site["name"]

            workbook = sp.download_latest_report(
                library="Shared Documents",
                reports_folder="Weekly Reports",
                destination_folder=local_folder
            )

            reports.append(
                {
                    "site": site["name"],
                    "workbook": workbook
                }
            )

        except Exception as ex:

            logger.exception(ex)

            reports.append(
                {
                    "site": site["name"],
                    "workbook": None
                }
            )

    return reports

##############################################################################
# DOWNLOAD REPORTS FROM LOCAL FOLDER
##############################################################################

def get_local_reports(root_folder):
    reports = []

    root_folder = Path(root_folder)

    # Allow the caller to point directly at a Weekly Reports folder, or at a
    # parent directory containing one or more site folders with a Weekly Reports
    # subfolder. Ignore any unrelated directories in the root.
    candidate_dirs = []

    if root_folder.name.lower() == "weekly reports" and root_folder.is_dir():
        candidate_dirs = [root_folder]
    else:
        candidate_dirs = [
            child for child in root_folder.iterdir()
            if child.is_dir() and (child / "Weekly Reports").exists()
        ]

    for site_folder in candidate_dirs:

        weekly_reports = site_folder / "Weekly Reports"

        print(f"Processing {site_folder.name}", flush=True)

        if not weekly_reports.exists():
            continue

        # Search recursively for all Excel files
        excel_files = []

        for ext in ("*.xlsx", "*.xlsm", "*.xls"):
            excel_files.extend(weekly_reports.rglob(ext))

        # Ignore temporary Excel files
        excel_files = [
            f for f in excel_files
            if not f.name.startswith("~$")
        ]
        #######################################################################
        # Manual file selection
        #######################################################################

        if MANUAL_REPORTS:
            manual = {

                Path(f).resolve()

                for f in MANUAL_REPORTS

            }

            excel_files = [

                f

                for f in excel_files

                if f.resolve() in manual

            ]

        #######################################################################
        # Trimester filter
        #######################################################################

        if FILTER_TRIMESTER:
            excel_files = [

                f

                for f in excel_files

                if FILTER_TRIMESTER.lower()

                   in str(f.parent).lower()

            ]

        #######################################################################
        # Week filter
        #######################################################################

        if FILTER_WEEK:
            excel_files = [

                f

                for f in excel_files

                if FILTER_WEEK.lower()

                   in str(f.parent).lower()

            ]

        #######################################################################
        # Nothing found
        #######################################################################

        if not excel_files:
            print(f"{site_folder.name:<20} No matching reports", flush=True)

            continue

        print(f"Processing {len(excel_files)} excel files", flush=True)

        ###############################################################################
        # Latest only
        ###############################################################################

        if PROCESS_LATEST_ONLY:

            workbook, warnings = select_latest_report(

                excel_files,

                site_folder.name

            )

            if workbook is not None:
                reports.append({

                    "site": site_folder.name,

                    "workbook": workbook,

                    "warnings": warnings

                })

        ###############################################################################
        # All reports
        ###############################################################################

        else:

            for workbook in excel_files:
                reports.append({

                    "site": site_folder.name,

                    "workbook": workbook,

                    "warnings": []

                })

    print(flush=True)

    print("=" * 60, flush=True)

    print("REPORT SELECTION", flush=True)

    print("=" * 60, flush=True)

    print(f"Latest Only : {PROCESS_LATEST_ONLY}", flush=True)

    print(f"Trimester   : {FILTER_TRIMESTER}", flush=True)

    print(f"Week        : {FILTER_WEEK}", flush=True)

    print(f"Manual      : {len(MANUAL_REPORTS)} file(s)", flush=True)

    print(f"Reports     : {len(reports)}", flush=True)

    print("=" * 60, flush=True)

    return reports

##############################################################################
# MAIN
##############################################################################

def main():

    start = time.time()

    logger.info("=" * 70)
    logger.info("Weekly Report Automation Started")
    logger.info("=" * 70)

    ###############################################################
    # Download reports
    ###############################################################

    parser = argparse.ArgumentParser(
        description="Weekly Report Automation"
    )

    parser.add_argument(
        "--mode",
        choices=["sharepoint", "local"],
        default=DATA_SOURCE,
        help="Source of the reports. Defaults to DATA_SOURCE in settings.py."
    )

    parser.add_argument(
        "--folder",
        default=str(LOCAL_REPORTS_FOLDER),
        help="Root folder containing site folders (local mode only)."
    )

    args = parser.parse_args()

    ###########################################################################
    # Determine report source
    ###########################################################################
    print(flush=True)
    print(f"Data Source : {args.mode}", flush=True)
    print(f"Path : {args.folder}", flush=True)

    if args.mode == "local":

        reports = get_local_reports(
            Path(args.folder)
        )

    else:

        reports = download_reports()

    ###############################################################
    # Consolidate
    ###############################################################
    print("Starting consolidation...", flush=True)
    consolidator = WeeklyReportConsolidator()

    print(f"Consolidating {len(reports)} reports...", flush=True)
    print("Processing reports...", flush=True)
    consolidator.process_reports(reports)

    ###############################################################
    # Export outputs
    ###############################################################
    print("Exporting reports...", flush=True)
    consolidator.export(OUTPUT_DIR)

    ###############################################################
    # Save history
    ###############################################################
    print("Update database...", flush=True)
    consolidator.save_database()

    print("Export history data...", flush=True)
    consolidator.export_history(OUTPUT_DIR)

    ###############################################################
    # Summary
    ###############################################################
    print("Print summary...", flush=True)
    consolidator.print_summary()

    ###############################################################
    # Upload outputs
    ###############################################################
    print("Upload to sharepoint...", flush=True)
    UPLOAD_TO_SHAREPOINT = False

    if UPLOAD_TO_SHAREPOINT:

        from auth import get_context
        from sharepoint import SharePointManager

        management = load_sites()[0]["management_site"]

        context = get_context(management)

        sp = SharePointManager(context)

        sp.upload_folder(
            OUTPUT_DIR,
            "Shared Documents/Consolidated Reports"
        )

    ###############################################################
    # Finish
    ###############################################################
    print("Ending consolidation...", flush=True)
    elapsed = round(time.time() - start, 2)

    logger.info(f"Completed in {elapsed} seconds.")

    print(flush=True)
    print("=" * 70, flush=True)
    print("WEEKLY REPORT AUTOMATION COMPLETE", flush=True)
    print("=" * 70, flush=True)
    print(f"Execution Time : {elapsed} seconds", flush=True)
    print(f"Reports Processed : {len(reports)}", flush=True)
    print(f"Output Folder : {OUTPUT_DIR}", flush=True)
    print("=" * 70, flush=True)


##############################################################################
# ENTRY
##############################################################################

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        logger.warning("Execution cancelled by user.")

    except Exception as ex:

        logger.exception(ex)

        print(flush=True)
        print("Automation failed.", flush=True)
        print(ex, flush=True)