"""
validator.py

Validation functions for the Weekly Report Automation project.

Author: Freedom Mukomana
"""

from pathlib import Path
import pandas as pd


###############################################################################
# Validation Result Object
###############################################################################

class ValidationResult:

    def __init__(self):

        self.errors = []
        self.warnings = []

    @property
    def passed(self):

        return len(self.errors) == 0

    def add_error(self, message):

        self.errors.append(message)

    def add_warning(self, message):

        self.warnings.append(message)


###############################################################################
# Validate a Single Record
###############################################################################

def validate_record(record):

    """
    Validate one extracted report.

    Parameters
    ----------
    record : dict

    Returns
    -------
    ValidationResult
    """

    result = ValidationResult()

    ###########################################################################
    # Required fields
    ###########################################################################

    required = [

        "site",
        "report_start_date",
        "report_end_date"
        #,
        #"allocated",
        #"completed"

    ]

    for field in required:

        if field not in record:

            result.add_error(f"Missing field: {field}")

            continue

        if record[field] in [None, ""]:

            result.add_error(f"{field} is empty")

    ###########################################################################
    # Stop if required fields are missing
    ###########################################################################

    if not result.passed:

        return result

    ###########################################################################
    # Dates
    ###########################################################################

    if record["report_start_date"] > record["report_end_date"]:

        result.add_error(

            "Start date occurs after End date"

        )

    ###########################################################################
    # Counts
    ###########################################################################

    allocated_raw = record.get("allocated")
    completed_raw = record.get("completed")

    both_blank = allocated_raw in [None, ""] and completed_raw in [None, ""]

    if both_blank:

        # Nothing was reported this week -- e.g. a Verbal Autopsy tab with
        # no cases that week, or a Field/Telephonic tab with no household
        # activity. This is a legitimate "no activity" week, not a data
        # error, so the numeric checks below are skipped rather than
        # failing on missing/blank values.

        result.add_warning(

            "No allocated/completed data reported this week (nothing to report)"

        )

    else:

        try:

            allocated = float(allocated_raw)

            completed = float(completed_raw)

            if allocated < 0:

                result.add_error(

                    "Allocated cannot be negative"

                )

            if completed < 0:

                result.add_error(

                    "Completed cannot be negative"

                )

            if completed > allocated:

                result.add_error(

                    "Completed exceeds Allocated"

                )

        except Exception:

            # Only one of allocated/completed was blank, or a non-numeric
            # value was entered -- this is a genuine partial-data anomaly,
            # not a "nothing to report" week, so it stays flagged.

            result.add_error(

                "Allocated/Completed are not numeric"

            )

    ###########################################################################
    # Percentage Fields
    ###########################################################################

    percentage_fields = [

        "completion_rate",
        "contact_rate",
        "dbs_rate",
        "hiv_rate",
        "height_weight_rate",
        "blood_glucose_rate",
        "health_utilisation_rate"

    ]

    for field in percentage_fields:

        value = record.get(field)

        if value is None:

            continue

        try:

            value = float(value)

        except Exception:

            result.add_error(

                f"{field} is not numeric"

            )

            continue

        if value < 0:

            result.add_error(

                f"{field} below 0%"

            )

        if value > 100:

            result.add_error(

                f"{field} exceeds 100%"

            )

    ###########################################################################
    # Warnings
    ###########################################################################

    if record.get("completion_rate") == 0:

        result.add_warning(

            "Completion rate is 0%"

        )

    if record.get("contact_rate") == 0:

        result.add_warning(

            "Contact rate is 0%"

        )

    return result


###############################################################################
# Validate Entire DataFrame
###############################################################################

def validate_dataframe(df):

    """
    Validate the consolidated dataframe.

    Returns
    -------
    pandas.DataFrame
    """

    issues = []

    ###########################################################################
    # Duplicate Site + Week
    ###########################################################################

    duplicates = df.duplicated(

        subset=["site", "report_end_date"],

        keep=False

    )

    for _, row in df[duplicates].iterrows():

        issues.append({

            "Site": row["site"],

            "Severity": "WARNING",

            "Issue": "Duplicate weekly report"

        })

    ###########################################################################
    # Individual Validation
    ###########################################################################

    for _, row in df.iterrows():

        result = validate_record(

            row.to_dict()

        )

        for error in result.errors:

            issues.append({

                "Site": row["site"],

                "Severity": "ERROR",

                "Issue": error

            })

        for warning in result.warnings:

            issues.append({

                "Site": row["site"],

                "Severity": "WARNING",

                "Issue": warning

            })

    return pd.DataFrame(issues)


###############################################################################
# Validation Summary
###############################################################################

def validation_summary(validation_df):

    if validation_df.empty:

        return {

            "Errors": 0,

            "Warnings": 0,

            "Total": 0

        }

    return {

        "Errors":

            len(

                validation_df[

                    validation_df["Severity"] == "ERROR"

                ]

            ),

        "Warnings":

            len(

                validation_df[

                    validation_df["Severity"] == "WARNING"

                ]

            ),

        "Total":

            len(validation_df)

    }


###############################################################################
# Save Validation Report
###############################################################################

def save_validation_report(

    validation_df,

    filename

):

    filename = Path(filename)

    with pd.ExcelWriter(

        filename,

        engine="openpyxl"

    ) as writer:

        validation_df.to_excel(

            writer,

            sheet_name="Validation",

            index=False

        )


###############################################################################
# Test
###############################################################################

if __name__ == "__main__":

    sample = {

        "site": "Johannesburg",

        "start_date": pd.Timestamp("2026-07-01"),

        "end_date": pd.Timestamp("2026-07-07"),

        "allocated": 100,

        "completed": 95,

        "completion_rate": 95,

        "contact_rate": 90

    }

    result = validate_record(sample)

    print(result.errors)

    print(result.warnings)