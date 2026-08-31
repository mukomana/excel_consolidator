"""
database.py

Database functions for the Weekly Report Automation project.

Author: Freedom Mukomana
"""

from datetime import datetime
from pathlib import Path
import logging

import pandas as pd

from sqlalchemy import (
    create_engine,
    Column,
    Integer,
    Float,
    String,
    Date,
    DateTime,
    UniqueConstraint
)

from sqlalchemy.orm import declarative_base
from sqlalchemy.orm import sessionmaker

###############################################################################
# CONFIGURATION
###############################################################################

DATABASE_FOLDER = Path("database")
DATABASE_FOLDER.mkdir(exist_ok=True)

DATABASE_FILE = DATABASE_FOLDER / "weekly_reports.db"

DATABASE_URL = f"sqlite:///{DATABASE_FILE}"

logger = logging.getLogger(__name__)

###############################################################################
# SQLAlchemy
###############################################################################

Base = declarative_base()

engine = create_engine(
    DATABASE_URL,
    echo=False
)

Session = sessionmaker(bind=engine)


###############################################################################
# FIELD MAPPING
#
# extract_report() (extractor.py) never produces keys literally called
# "start_date" / "end_date" -- it produces "report_start_date" and
# "report_end_date" (present on every record, regardless of survey type --
# see COMMON_KPI_LABELS in extractor.py), plus survey-specific period
# fields (period_start/period_end for Field & Telephonic, year_start_date/
# year_end_date for VA) that describe the FCA/Trimester or calendar year,
# not the reporting week itself.
#
# For a WEEKLY report table, "the week this row covers" should be the
# reporting week -- report_start_date / report_end_date -- not the
# trimester/year boundary. This mapping makes that explicit instead of
# silently reading a key that was never being set (record.get("start_date")
# / record.get("end_date") previously always returned None).
###############################################################################

def map_record_to_db_fields(record: dict) -> dict:
    """
    Translate a record produced by extract_report() into the flat dict
    of column values expected by the WeeklyReport table.
    """

    return {
        "site": record.get("site"),
        "survey_type": record.get("survey_type"),

        "start_date": record.get("report_start_date"),
        "end_date": record.get("report_end_date"),

        "allocated": record.get("allocated"),
        "completed": record.get("completed"),

        "non_contact": record.get("non_contact"),
        "passive_refusal": record.get("passive_refusal"),
        "contacted": record.get("contacted"),
        "refused": record.get("refused"),
        "participated": record.get("participated"),
        "consented_yes": record.get("consented_yes"),
        "consented_no": record.get("consented_no"),
        "premature": record.get("premature"),  # VA-specific

        "completion_rate": record.get("completion_rate"),
        "contact_rate": record.get("contact_rate"),
        "participation_rate": record.get("participation_rate"),

        # KPI_LABELS-derived rates -- only populated if/when extract_report()
        # is extended to read the KPI_LABELS sheet; left in place so the
        # column exists and older rows / future data aren't lost.
        "dbs_rate": record.get("dbs_rate"),
        "hiv_rate": record.get("hiv_rate"),
        "height_weight_rate": record.get("height_weight_rate"),
        "blood_glucose_rate": record.get("blood_glucose_rate"),
        "health_utilisation_rate": record.get("health_utilisation_rate"),

        "source_file": record.get("source_file"),
    }


###############################################################################
# TABLE
###############################################################################

class WeeklyReport(Base):

    __tablename__ = "weekly_reports"

    id = Column(Integer, primary_key=True)

    site = Column(String(100), nullable=False)

    # Distinguishes Field / Telephonic / Verbal Autopsy records for the
    # same site + week. Without this, the uniqueness constraint below
    # would treat different survey types in the same week as duplicates
    # and silently drop all but the first one inserted.
    survey_type = Column(String(50), nullable=False)

    start_date = Column(Date)

    end_date = Column(Date)

    allocated = Column(Integer)

    completed = Column(Integer)

    # Cumulative count fields (these were previously missing entirely --
    # only allocated/completed and the rate columns were being stored,
    # so weekly deltas for contact/refusal/participation/consent had no
    # underlying data to diff against).
    non_contact = Column(Integer)

    passive_refusal = Column(Integer)

    contacted = Column(Integer)

    refused = Column(Integer)

    participated = Column(Integer)

    consented_yes = Column(Integer)

    consented_no = Column(Integer)

    premature = Column(Integer)  # VA-specific

    completion_rate = Column(Float)

    contact_rate = Column(Float)

    participation_rate = Column(Float)

    dbs_rate = Column(Float)

    hiv_rate = Column(Float)

    height_weight_rate = Column(Float)

    blood_glucose_rate = Column(Float)

    health_utilisation_rate = Column(Float)

    source_file = Column(String(255))

    imported_at = Column(
        DateTime,
        default=datetime.utcnow
    )

    __table_args__ = (

        UniqueConstraint(

            "site",

            "survey_type",

            "end_date",

            name="uq_site_survey_week"

        ),

    )


###############################################################################
# INITIALISE DATABASE
###############################################################################

def initialise_database():
    """
    Create database tables if they do not exist.

    NOTE: this only creates tables that don't exist yet -- it does NOT
    alter an existing table's columns. If weekly_reports.db already exists
    from before this schema change, the new columns (survey_type,
    contacted, refused, participated, consented_yes, consented_no,
    non_contact, passive_refusal, premature, participation_rate) will be
    missing from it and inserts will fail. For a dev database it's usually
    simplest to delete database/weekly_reports.db and let it be recreated
    fresh; for a production database with data worth keeping, this needs
    a proper migration (e.g. Alembic) instead of just deleting the file.
    """

    Base.metadata.create_all(engine)

    logger.info("Database initialised.")


###############################################################################
# INSERT
###############################################################################

def insert_record(record: dict):
    """
    Insert one weekly report. Accepts a raw record as produced by
    extract_report() and maps it to the database's column names/keys.
    """

    session = Session()

    try:

        fields = map_record_to_db_fields(record)

        exists = session.query(WeeklyReport).filter(

            WeeklyReport.site == fields["site"],

            WeeklyReport.survey_type == fields["survey_type"],

            WeeklyReport.end_date == fields["end_date"]

        ).first()

        if exists:

            logger.warning(

                f"{fields['site']} "

                f"{fields['survey_type']} "

                f"{fields['end_date']} "

                "already exists."

            )

            return False

        row = WeeklyReport(**fields)

        session.add(row)

        session.commit()

        logger.info(

            f"Inserted "

            f"{fields['site']} "

            f"{fields['survey_type']}"

        )

        return True

    except Exception as ex:

        session.rollback()

        logger.exception(ex)

        return False

    finally:

        session.close()


###############################################################################
# INSERT DATAFRAME
###############################################################################

def insert_dataframe(df: pd.DataFrame):
    """
    Insert every row from a dataframe.
    """

    inserted = 0

    skipped = 0

    for _, row in df.iterrows():

        success = insert_record(

            row.to_dict()

        )

        if success:

            inserted += 1

        else:

            skipped += 1

    logger.info(

        f"Inserted={inserted} "

        f"Skipped={skipped}"

    )


###############################################################################
# READ
###############################################################################

def load_history():

    query = """

    SELECT *

    FROM weekly_reports

    ORDER BY

        end_date,

        site,

        survey_type

    """

    return pd.read_sql(

        query,

        engine

    )


###############################################################################
# SITE HISTORY
###############################################################################

def load_site(site):

    query = """

    SELECT *

    FROM weekly_reports

    WHERE site = :site

    ORDER BY end_date

    """

    return pd.read_sql(

        query,

        engine,

        params={

            "site": site

        }

    )


###############################################################################
# DELETE WEEK
###############################################################################

def delete_week(site, end_date, survey_type=None):

    session = Session()

    try:

        rows = session.query(

            WeeklyReport

        ).filter(

            WeeklyReport.site == site,

            WeeklyReport.end_date == end_date

        )

        if survey_type is not None:

            rows = rows.filter(

                WeeklyReport.survey_type == survey_type

            )

        deleted = rows.delete()

        session.commit()

        logger.info(

            f"Deleted {deleted}"

        )

    except Exception:

        session.rollback()

        raise

    finally:

        session.close()


###############################################################################
# EXPORT
###############################################################################

def export_excel(filename):

    df = load_history()

    df.to_excel(

        filename,

        index=False

    )


###############################################################################
# SUMMARY
###############################################################################

def database_summary():

    df = load_history()

    summary = {

        "Sites":

            df["site"].nunique(),

        "Weeks":

            df["end_date"].nunique(),

        "Records":

            len(df)

    }

    return summary


###############################################################################
# TEST
###############################################################################

if __name__ == "__main__":

    initialise_database()

    print(

        database_summary()

    )