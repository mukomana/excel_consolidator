"""
report_selector.py

Selects the correct weekly report from a collection of Excel files.

The filename is ignored. The report with the latest reporting
period (End Date) is selected.

Author: Freedom Mukomana
"""

from pathlib import Path
from typing import List, Optional, Tuple
import logging

from extractor import extract_report

logger = logging.getLogger(__name__)


##############################################################################
# Helpers
##############################################################################

VALID_EXTENSIONS = [".xlsx", ".xlsm"]


def is_valid_excel(filename: str) -> bool:
    """
    Returns True if the file is a valid Excel workbook.
    """

    name = Path(filename).name

    if name.startswith("~$"):
        return False

    if Path(name).suffix.lower() not in VALID_EXTENSIONS:
        return False

    return True


##############################################################################
# Scan folder
##############################################################################

def get_excel_files(folder: Path) -> List[Path]:
    """
    Returns every Excel workbook inside a folder and all subfolders.
    """

    files = []

    for ext in ("*.xlsx", "*.xlsm", "*.xls"):

        files.extend(folder.rglob(ext))

    files = [

        f for f in files

        if f.is_file()

        and not f.name.startswith("~$")

    ]

    return files


##############################################################################
# Select latest report
##############################################################################

def select_latest_report(
    folder: Path,
    site_name: str
) -> Tuple[Optional[Path], List[str]]:
    """
    Returns

        latest workbook

        warning list

    """

    warnings = []

    excel_files = get_excel_files(folder)

    if len(excel_files) == 0:

        warnings.append("No Excel reports found.")

        return None, warnings

    reports = []

    for workbook in excel_files:

        try:

            data = extract_report(
                workbook,
                site_name
            )

            reports.append(

                {
                    "file": workbook,
                    "start": data.get("start_date"),
                    "end": data.get("end_date")
                }

            )

        except Exception as ex:

            logger.exception(ex)

            warnings.append(
                f"Unable to read {workbook.name}"
            )

    if len(reports) == 0:

        warnings.append(
            "No readable reports."
        )

        return None, warnings

    reports.sort(

        key=lambda x: x["end"]

        if x["end"] is not None

        else "",

        reverse=True

    )

    latest = reports[0]

    duplicates = [

        x

        for x in reports

        if x["end"] == latest["end"]

    ]

    if len(duplicates) > 1:

        warnings.append(

            f"{len(duplicates)} reports found for "

            f"{latest['end']}."

        )

        duplicates.sort(

            key=lambda x: x["file"].stat().st_mtime,

            reverse=True

        )

        latest = duplicates[0]

    logger.info(

        f"{site_name}: "

        f"{latest['file'].name} selected"

    )

    return latest["file"], warnings


##############################################################################
# Build processing queue
##############################################################################

def build_queue(root_folder: Path):
    """
    Creates a processing queue.

    Folder structure

    root/

        Johannesburg/

            Weekly Reports/

        Durban/

            Weekly Reports/

    """

    queue = []

    for site in root_folder.iterdir():

        if not site.is_dir():
            continue

        reports = site / "Weekly Reports"

        if not reports.exists():
            continue

        workbook, warnings = select_latest_report(

            reports,

            site.name

        )

        queue.append(

            {

                "site": site.name,

                "folder": reports,

                "workbook": workbook,

                "warnings": warnings

            }

        )

    return queue


##############################################################################
# Summary
##############################################################################

def print_summary(queue):

    print()

    print("-" * 60)

    print("Processing Queue")

    print("-" * 60)

    for item in queue:

        print()

        print(item["site"])

        if item["workbook"]:

            print(

                "Workbook:",

                item["workbook"].name

            )

        else:

            print(

                "Workbook: NONE"

            )

        if item["warnings"]:

            print(

                "Warnings:"

            )

            for warning in item["warnings"]:

                print(

                    "  -",

                    warning

                )

    print()

    print("-" * 60)


##############################################################################
# Test
##############################################################################

if __name__ == "__main__":

    ROOT = Path(r"C:\WeeklyReports")

    queue = build_queue(ROOT)

    print_summary(queue)