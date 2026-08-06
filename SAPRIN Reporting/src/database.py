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
# TABLE
###############################################################################

class WeeklyReport(Base):

    __tablename__ = "weekly_reports"

    id = Column(Integer, primary_key=True)

    site = Column(String(100), nullable=False)

    start_date = Column(Date)

    end_date = Column(Date)

    allocated = Column(Integer)

    completed = Column(Integer)

    completion_rate = Column(Float)

    contact_rate = Column(Float)

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

            "end_date",

            name="uq_site_week"

        ),

    )


###############################################################################
# INITIALISE DATABASE
###############################################################################

def initialise_database():
    """
    Create database tables if they do not exist.
    """

    Base.metadata.create_all(engine)

    logger.info("Database initialised.")


###############################################################################
# INSERT
###############################################################################

def insert_record(record: dict):
    """
    Insert one weekly report.
    """

    session = Session()

    try:

        exists = session.query(WeeklyReport).filter(

            WeeklyReport.site == record["site"],

            WeeklyReport.end_date == record["end_date"]

        ).first()

        if exists:

            logger.warning(

                f"{record['site']} "

                f"{record['end_date']} "

                "already exists."

            )

            return False

        row = WeeklyReport(

            site=record.get("site"),

            start_date=record.get("start_date"),

            end_date=record.get("end_date"),

            allocated=record.get("allocated"),

            completed=record.get("completed"),

            completion_rate=record.get("completion_rate"),

            contact_rate=record.get("contact_rate"),

            dbs_rate=record.get("dbs_rate"),

            hiv_rate=record.get("hiv_rate"),

            height_weight_rate=record.get("height_weight_rate"),

            blood_glucose_rate=record.get("blood_glucose_rate"),

            health_utilisation_rate=record.get("health_utilisation_rate"),

            source_file=record.get("source_file")

        )

        session.add(row)

        session.commit()

        logger.info(

            f"Inserted "

            f"{record['site']}"

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

        site

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

def delete_week(site, end_date):

    session = Session()

    try:

        rows = session.query(

            WeeklyReport

        ).filter(

            WeeklyReport.site == site,

            WeeklyReport.end_date == end_date

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