"""
consolidate.py

Consolidates weekly reports from all sites into a single dataset.

Author: Freedom Mukomana
"""

from pathlib import Path
import logging

import pandas as pd

from extractor import extract_report
from validator import (
    validate_record,
    validate_dataframe,
    validation_summary,
    save_validation_report
)

from database import (
    initialise_database,
    insert_dataframe,
    load_history
)

from weekly_deltas import add_weekly_columns

logger = logging.getLogger(__name__)


###############################################################################
# Consolidator
###############################################################################

class WeeklyReportConsolidator:

    def __init__(self):

        self.survey_records = []

        self.failed = []

        self.warnings = []

    ###########################################################################
    # Process a single workbook
    ###########################################################################
    def process_report(
            self,
            filepath: Path,
            site_name: str
    ):

        logger.info(f"Reading {filepath}")

        print("Entered process_reports()")
        try:

            #######################################################################
            # Extract all worksheets
            #######################################################################

            survey_records = extract_report(
                filepath,
                site_name
            )
            print(f"Number of survey records: {len(survey_records)}")
            #######################################################################
            # Validate each extracted record
            #######################################################################

            for survey_record in survey_records:

                result = validate_record(survey_record)

                if not result.passed:
                    self.failed.append({

                        "Site": site_name,

                        "Survey": survey_record.get("survey_type",),

                        "File": filepath.name,

                        "Errors": "; ".join(result.errors)

                    })

                    continue

                if result.warnings:
                    self.warnings.append({

                        "Site": site_name,

                        "Survey": survey_record.get("survey_type"),

                        "File": filepath.name,

                        "Warnings": "; ".join(result.warnings)

                    })

                self.survey_records.append(survey_record)

        except Exception as ex:

            logger.exception(ex)

            self.failed.append({

                "Site": site_name,

                "Survey": "",

                "File": filepath.name,

                "Errors": str(ex)

            })

    ###########################################################################
    # Process multiple reports
    ###########################################################################

    def process_reports(self, reports):
        print("Entered process_reports()")
        print(f"Reports received: {len(reports)}")
        """
        reports

        [

            {

                "site":"Johannesburg",

                "workbook":Path(...)

            }

        ]
        """

        for report in reports:

            workbook = report.get("workbook")

            if workbook is None:

                self.failed.append({

                    "Site": report["site"],

                    "File": "",

                    "Errors": "No report"

                })

                continue

            self.process_report(

                workbook,

                report["site"]

            )

    ###########################################################################
    # Build dataframe
    ###########################################################################

    def dataframe(self):

        return pd.DataFrame(self.survey_records)

    ###########################################################################
    # Validation
    ###########################################################################

    def validate(self):

        df = self.dataframe()

        if len(df) == 0:

            return pd.DataFrame()

        return validate_dataframe(df)

    ###########################################################################
    # Export
    ###########################################################################

    def export(self, output_folder):

        output_folder = Path(output_folder)

        output_folder.mkdir(

            exist_ok=True

        )

        df = self.dataframe()

        if df.empty:
            logger.warning("No data to export.")
            return

        #######################################################################
        # Weekly Excel
        #######################################################################

        excel = output_folder / "Weekly_Consolidated_Report.xlsx"

        with pd.ExcelWriter(

            excel,

            engine="xlsxwriter"

        ) as writer:
            if "survey_type" in df.columns:
                print(df.columns.tolist())
                for survey in sorted(df["survey_type"].unique()):

                    subset = df[

                        df["survey_type"] == survey
                        ]

                    subset.to_excel(

                        writer,

                        sheet_name=survey[:31],

                        index=False

                    )

        #######################################################################
        # CSV
        #######################################################################

        csv = output_folder / "Weekly_Consolidated_Report.csv"

        df.to_csv(

            csv,

            index=False

        )

        #######################################################################
        # Validation
        #######################################################################

        validation = self.validate()

        validation_file = output_folder / "Validation_Report.xlsx"

        save_validation_report(

            validation,

            validation_file

        )

        #######################################################################
        # Failed Reports
        #######################################################################

        failed = pd.DataFrame(

            self.failed

        )

        failed.to_excel(

            output_folder / "Failed_Reports.xlsx",

            index=False

        )

        #######################################################################
        # Warnings
        #######################################################################

        warnings = pd.DataFrame(

            self.warnings

        )

        warnings.to_excel(

            output_folder / "Warnings.xlsx",

            index=False

        )

        logger.info("Outputs written.")

    ###########################################################################
    # Save to Database
    ###########################################################################

    def save_database(self):

        initialise_database()

        df = self.dataframe()

        insert_dataframe(df)

    ###########################################################################
    # Historical Export
    ###########################################################################

    def export_history(self, output_folder):

        history = load_history()

        if not history.empty:
            history = add_weekly_columns(history)

        history.to_excel(

            Path(output_folder)

            / "Historical_Weekly_Data.xlsx",

            index=False

        )

    ###########################################################################
    # Summary
    ###########################################################################

    def summary(self):

        df = self.dataframe()

        validation = self.validate()

        summary = validation_summary(validation)
        print(df.columns.tolist())
        summary["Workbooks Processed"] = len(df["source_file"].unique())

        summary["Survey Records"] = len(df)

        summary["Failed Reports"] = len(self.failed)

        summary["Warnings"] = len(self.warnings)

        return summary

    ###########################################################################
    # Print Summary
    ###########################################################################

    def print_summary(self):

        summary = self.summary()

        print()

        print("=" * 60)

        print("WEEKLY REPORT CONSOLIDATION")

        print("=" * 60)

        for key, value in summary.items():

            print(f"{key:<25}{value}")

        print("=" * 60)


###############################################################################
# Test
###############################################################################

if __name__ == "__main__":

    reports = [

        {

            "site": "Johannesburg",

            "workbook": Path("Johannesburg.xlsx")

        },

        {

            "site": "Durban",

            "workbook": Path("Durban.xlsx")

        }

    ]

    consolidator = WeeklyReportConsolidator()

    consolidator.process_reports(reports)

    consolidator.export("output")

    consolidator.save_database()

    consolidator.export_history("output")

    consolidator.print_summary()