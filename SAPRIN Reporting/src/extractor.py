"""
extractor.py

Reads an IHS Weekly Report workbook and extracts KPIs.

Author: Freedom Mukomana
"""

from pathlib import Path
from typing import Dict, Optional

import re
from datetime import datetime

from openpyxl import load_workbook

###########################################################################
# KPI LABELS
###########################################################################

KPI_LABELS = {
    "start_date": [
        "Start Date"
    ],

    "end_date": [
        "End Date"
    ],

    "fca": [
        "FCA"
    ],

    "allocated": [
        "Individuals allocated"
    ],

    "completed": [
        "Completed"
    ],

    "completion_rate": [
        "Completion Rate"
    ],

    "contact_rate": [
        "Contact Rate"
    ],

    "dbs_rate": [
        "Participation Rate-DBS"
    ],

    "hiv_rate": [
        "Participation Rate-HIV"
    ],

    "height_weight_rate": [
        "Participation Rate-Height & Weight"
    ],

    "blood_glucose_rate": [
        "Participation Rate-Blood Glucose"
    ],

    "health_utilisation_rate": [
        "Participation Rate-Health Utilisation"
    ]
}

###############################################################################
# COMMON KPIs
###############################################################################

COMMON_KPI_LABELS = {

    "report_start_date": [
        "Which week are you reporting on (indicate start and end date)"
    ],

    "report_end_date": [
        "Which week are you reporting on (indicate start and end date)"
    ],

    "weeks_elapsed": [
        "How many weeks have gone since starting"
    ],

    "time_elapsed_pct": [
        "What percentage of time have you now covered"
    ],

    "allocated": [
        "allocated"
    ],

    "completed": [
        "completed"
    ],

    "completion_rate": [
        "Completion Rate"
    ],

    "contact_rate": [
        "Contact  Rate"
    ],

    "participation_rate": [
        "Participation Rate"
    ]

}

###############################################################################
# FIELD + TELEPHONIC
###############################################################################

FIELD_TELEPHONE_LABELS = {

    "fca": [
        "Which FCA /Trimester Are you currently working in?"
    ],

    "period_start": [
        "On which date did you start this 15 week FCA/Trimester?"
    ],

    "period_end": [
        "On which date did you start this 15 week FCA/Trimester?"
    ],

    "non_contact": [
        "Non-contact"
    ],

    "non_contact_rate": [
        "Non-contact"
    ],

    "passive_refusal": [
        "Passive Refusal"
    ],

    "passive_refusal_rate": [
        "Passive Refusal"
    ],

    "contacted": [
        "Contacted"
    ],

    "refused": [
        "Refused"
    ],

    "refusal_rate": [
        "Refusal Rate"
    ],

    "participated": [
        "Participated",
        "Particpated"
    ],

    "consented_yes": [
        "Consented YES"
    ],

    "consented_no": [
        "Consented No"
    ]

}

###############################################################################
# VERBAL AUTOPSY
###############################################################################

VA_LABELS = {

    "reporting_year": [
        "Which year you currently working in?"
    ],

    "year_start_date": [
        "On which date did you start working on cases in this calendar year?"
    ],

    "year_end_date": [
        "On which date did you start working on cases in this calendar year?"
    ],

    "premature": [
        "Premature"
    ],

    "premature_rate": [
        "Premature"
    ],

    "non_contact": [
        "Non-contact"
    ],

    "non_contact_rate": [
        "Non-contact"
    ],

    "contacted": [
        "Contacted"
    ],

    "refused": [
        "Refused"
    ],

    "participated": [
        "Participated",
        "Particpated"
    ],

    "consented_yes": [
        "Yes"
    ],

    "consented_no": [
        "No"
    ]

}
###########################################################################
# HELPERS
###########################################################################

def normalise(value):

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\n", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip().lower()


def percentage_to_float(value):

    if value is None:
        return None

    if isinstance(value, (int, float)):
        if value <= 1:
            return value * 100
        return value

    value = str(value).replace("%", "").strip()

    try:
        return float(value)
    except Exception:
        return value


def parse_date(value):

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    return value


###########################################################################
# FIND LABEL
###########################################################################

def find_value(ws, labels):

    """
    Search worksheet for one of the labels.

    Returns the value immediately to the right.
    """

    labels = [normalise(x) for x in labels]

    for row in ws.iter_rows():

        for cell in row:

            cell_value = normalise(cell.value)

            if cell_value in labels:

                # Right cell

                value = ws.cell(
                    row=cell.row,
                    column=cell.column + 1
                ).value

                if value is None:

                    # Try two columns to the right

                    value = ws.cell(
                        row=cell.row,
                        column=cell.column + 2
                    ).value

                return value

    return None

def extract_sheet(sheet, survey_type, labels):

    record = {

        "survey_type": survey_type

    }

    ###########################################################################
    # Common KPIs
    ###########################################################################

    for field, aliases in COMMON_KPI_LABELS.items():

        value = find_value(sheet, aliases)

        if value is None:
            continue

        if "rate" in field:
            value = percentage_to_float(value)

        if "date" in field:
            value = parse_date(value)

        record[field] = value

    ###########################################################################
    # Survey-specific KPIs
    ###########################################################################

    for field, aliases in labels.items():

        value = find_value(sheet, aliases)

        if value is None:
            continue

        if "rate" in field:
            value = percentage_to_float(value)

        if "date" in field:
            value = parse_date(value)

        record[field] = value

    return record

###########################################################################
# EXTRACT
###########################################################################

def extract_report(filepath, site_name):

    workbook = load_workbook(
        filepath,
        data_only=True
    )

    records = []

    ###########################################################################
    # Field Household Surveillance
    ###########################################################################

    if "Field-HouseholdSurveillance" in workbook.sheetnames:

        record = {

            "site": site_name,

            "source_file": Path(filepath).name

        }

        record.update(

            extract_sheet(

                workbook["Field-HouseholdSurveillance"],

                "Field",

                FIELD_TELEPHONE_LABELS

            )

        )

        records.append(record)

    ###########################################################################
    # Telephonic Household Surveillance
    ###########################################################################

    if "Telephonic HouseholdSurveillance" in workbook.sheetnames:

        record = {

            "site": site_name,

            "source_file": Path(filepath).name

        }

        record.update(

            extract_sheet(

                workbook["Telephonic HouseholdSurveillance"],

                "Telephonic",

                FIELD_TELEPHONE_LABELS

            )

        )

        records.append(record)

    ###########################################################################
    # Verbal Autopsy
    ###########################################################################

    if "VerbalAutopsy" in workbook.sheetnames:

        record = {

            "site": site_name,

            "source_file": Path(filepath).name

        }

        record.update(

            extract_sheet(

                workbook["VerbalAutopsy"],

                "Verbal Autopsy",

                VA_LABELS

            )

        )

        records.append(record)

    workbook.close()

    return records

###########################################################################
# VALIDATION
###########################################################################

REQUIRED_FIELDS = [

    "start_date",

    "end_date",

    "allocated",

    "completed"

]


def validate_record(record):

    missing = []

    for field in REQUIRED_FIELDS:

        if field not in record:

            missing.append(field)

    return missing


###########################################################################
# TEST
###########################################################################

if __name__ == "__main__":

    report = extract_report(

        "Sample Weekly Report.xlsx",

        "Johannesburg"

    )

    print(report)

    print()

    print(validate_record(report))