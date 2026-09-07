"""
settings.py

Central configuration for the Weekly Report Automation project.

Author: Freedom Mukomana
"""

from pathlib import Path

###############################################################################
# PROJECT PATHS
###############################################################################

# Project root folder
ROOT_DIR = Path(__file__).resolve().parent

# Configuration
CONFIG_DIR = ROOT_DIR / "config"
CONFIG_FILE = CONFIG_DIR / "nodes.json"

# Working folders
INPUT_DIR = ROOT_DIR / "input"
OUTPUT_DIR = ROOT_DIR / "output"
ARCHIVE_DIR = ROOT_DIR / "archive"
LOG_DIR = ROOT_DIR / "logs"
DATABASE_DIR = ROOT_DIR / "database"

# Create folders if they don't exist
for folder in [
    INPUT_DIR,
    OUTPUT_DIR,
    ARCHIVE_DIR,
    LOG_DIR,
    DATABASE_DIR
]:
    folder.mkdir(parents=True, exist_ok=True)

###############################################################################
# DATA SOURCE
###############################################################################

# Options:
#   "local"
#   "sharepoint"

DATA_SOURCE = "local"

# Used when DATA_SOURCE = "local"
LOCAL_REPORTS_FOLDER = r"C:\Users\freedom.mukomana\OneDrive - South African Medical Research Council\SAPRIN - South African Population Research Infrastructure Network-Reporting - Documents"

###############################################################################
# SHAREPOINT
###############################################################################

# Only used when DATA_SOURCE = "sharepoint"

MANAGEMENT_SITE = "https://YOURTENANT.sharepoint.com/sites/Management"

MANAGEMENT_LIBRARY = "Shared Documents"

MANAGEMENT_REPORT_FOLDER = "Consolidated Reports"

###############################################################################
# DATABASE
###############################################################################

DATABASE_FILE = DATABASE_DIR / "weekly_reports.db"

###############################################################################
# OUTPUT FILES
###############################################################################

CONSOLIDATED_REPORT = OUTPUT_DIR / "Weekly_Consolidated_Report.xlsx"

VALIDATION_REPORT = OUTPUT_DIR / "Validation_Report.xlsx"

FAILED_REPORTS = OUTPUT_DIR / "Failed_Reports.xlsx"

WARNING_REPORT = OUTPUT_DIR / "Warnings.xlsx"

HISTORY_REPORT = OUTPUT_DIR / "Historical_Weekly_Data.xlsx"

###############################################################################
# LOGGING
###############################################################################

LOG_FILE = LOG_DIR / "automation.log"

LOG_LEVEL = "INFO"

###############################################################################
# VALIDATION
###############################################################################

REQUIRED_COLUMNS = [

    "site",
    "start_date",
    "end_date",
    "allocated",
    "completed"

]

PERCENTAGE_COLUMNS = [

    "completion_rate",
    "contact_rate",
    "dbs_rate",
    "hiv_rate",
    "height_weight_rate",
    "blood_glucose_rate",
    "health_utilisation_rate"

]

###############################################################################
# EXCEL SETTINGS
###############################################################################

EXCEL_EXTENSIONS = [

    ".xlsx",
    ".xlsm",
    ".xls"

]

IGNORE_PREFIXES = [

    "~$"

]

###############################################################################
# REPORT SETTINGS
###############################################################################

DATE_FORMAT = "%Y-%m-%d"

MAX_REPORT_AGE_DAYS = 30

###############################################################################
# PERFORMANCE
###############################################################################

MAX_DOWNLOAD_THREADS = 5

MAX_UPLOAD_RETRIES = 3

###############################################################################
# DEBUG
###############################################################################

DEBUG = True

VERBOSE = True

###############################################################################
# REPORT PROCESSING OPTIONS
###############################################################################

# True  -> Process only the latest report for each site
# False -> Process all reports found
PROCESS_LATEST_ONLY = False

###############################################################################
# OPTIONAL FILTERS
###############################################################################

# Examples:
# FILTER_TRIMESTER = "Trimester 1"
# FILTER_WEEK = "Week 09"

FILTER_TRIMESTER = None
FILTER_WEEK = None

###############################################################################
# MANUAL FILE SELECTION
###############################################################################

# If this list contains one or more files,
# ONLY these files will be processed.
#
# Leave empty to ignore.

MANUAL_REPORTS = [

    # ROOT_DIR.parent /
    # "Archive/Johannesburg/Weekly Reports/Trimester 1/Week 09/JHB.xlsx",

]