"""
weekly_deltas.py

Reads a folder of weekly IHS report files (one .xlsx per week, same site)
and converts the cumulative/aggregate values each file reports into
week-on-week incremental numbers.

Depends on extractor.py (extract_report) living alongside this file.

Author: Freedom Mukomana
"""

from pathlib import Path

import pandas as pd

from extractor import extract_report

###########################################################################
# WHICH FIELDS ARE CUMULATIVE COUNTS (diffable) VS RATES (recomputed)
###########################################################################

# Raw counts that accumulate week over week within an FCA/Trimester (or
# VA calendar year). Weekly value = this_week - previous_week.
CUMULATIVE_COUNT_FIELDS = [
    "completed",
    "non_contact",
    "passive_refusal",
    "contacted",
    "refused",
    "participated",
    "consented_yes",
    "consented_no",
    "premature",  # VA-specific
]

# Fields that should NOT be diffed -- they're a fixed target/denominator
# for the whole reporting period, not a running total.
STATIC_FIELDS = [
    "allocated",
]

# Rates are recomputed FROM the weekly incremental counts, not subtracted
# from each other (a rate isn't additive across weeks). Maps the output
# weekly-rate field name -> (numerator field, denominator field), both
# referring to the *weekly* (already-diffed) count fields above.
WEEKLY_RATE_FORMULAS = {
    "weekly_contact_rate": ("weekly_contacted", "weekly_completed"),
    "weekly_refusal_rate": ("weekly_refused", "weekly_contacted"),
    "weekly_participation_rate": ("weekly_participated", "weekly_contacted"),
    "weekly_consent_rate": ("weekly_consented_yes", "weekly_participated"),
}

###########################################################################
# GROUPING / SORTING KEY
###########################################################################

def group_key(record):
    """
    Records are grouped by (site, survey_type) so that Field, Telephonic,
    and Verbal Autopsy numbers -- which run on independent cumulative
    counters -- are never diffed against each other.
    """
    return (record.get("site"), record.get("survey_type"))


def sort_key(record):
    """
    Sort chronologically within a group. Checks, in order:
    - "start_date" -- the column name used once records come FROM the
      database (load_history()), after database.py's field mapping
    - "report_start_date" -- the raw key extract_report() produces when
      called directly (e.g. process_weekly_folder(), before DB insertion)
    - "period_start" / "year_start_date" -- last-resort fallbacks
    Falls back to source filename if no date is found, so a bad/blank
    date cell doesn't crash the sort.
    """
    date = (
        record.get("start_date")
        or record.get("report_start_date")
        or record.get("period_start")
        or record.get("year_start_date")
    )
    if date is not None:
        return (0, date)
    return (1, record.get("source_file", ""))


###########################################################################
# DELTA COMPUTATION
###########################################################################

def compute_weekly_deltas(records):
    """
    Given a list of extracted weekly records (already all belonging to the
    same site + survey_type, any order), sort them chronologically and
    return a new list of records with added "weekly_<field>" keys holding
    the incremental value for that week.

    Handles blank/"nothing to report" weeks correctly: a workbook can
    still produce a row for a week where a given field simply wasn't
    filled in (find_value() returns None, so the field is missing/None
    on that record). Diffing only against the IMMEDIATELY PREVIOUS record
    would be wrong in that case -- if last week was blank, this week would
    get compared against None and get treated as "week one", reporting
    its entire cumulative total as a single week's number.

    Instead, each field tracks its own last-known non-blank value across
    the whole ordered sequence, so a blank week is skipped over rather
    than resetting the baseline for the following week.

    The first week a field has real data has nothing valid to diff
    against, so its weekly_<field> is set equal to its cumulative value
    (i.e. assumed to start from zero).
    """

    ordered = sorted(records, key=sort_key)

    last_known_value = {}  # field -> last non-None cumulative value seen
    output = []

    for record in ordered:

        weekly_record = dict(record)  # shallow copy, keep original cumulative fields intact

        for field in CUMULATIVE_COUNT_FIELDS:

            current_value = record.get(field)

            if current_value is None:
                # Nothing reported this week for this field -- we can't
                # compute a weekly delta, and we deliberately do NOT
                # update last_known_value, so the next week that DOES
                # have real data diffs against the last week that
                # actually reported one, not this blank week.
                continue

            if field not in last_known_value:
                weekly_value = current_value  # first time this field has real data: nothing to subtract
            else:
                weekly_value = current_value - last_known_value[field]

                if weekly_value < 0:
                    # Cumulative counter went DOWN -- almost always means a
                    # data correction/re-entry happened, not a real negative
                    # week. Flag it instead of silently reporting nonsense.
                    weekly_record.setdefault("_warnings", []).append(
                        f"{field} decreased vs last known value "
                        f"({last_known_value[field]} -> {current_value}); "
                        f"likely a data correction, not a true weekly count."
                    )

            weekly_record[f"weekly_{field}"] = weekly_value
            last_known_value[field] = current_value

        if not any(record.get(f) is not None for f in CUMULATIVE_COUNT_FIELDS):
            # None of the tracked fields had any data this week -- flag the
            # whole row as a non-reporting week so it's easy to spot/filter
            # downstream rather than just silently having no weekly_* values.
            weekly_record["_no_data_reported"] = True

        # Recompute rates from this week's incremental counts (not diffed)
        for rate_field, (numerator_field, denominator_field) in WEEKLY_RATE_FORMULAS.items():

            numerator = weekly_record.get(numerator_field)
            denominator = weekly_record.get(denominator_field)

            if numerator is None or denominator is None:
                continue

            if denominator == 0:
                weekly_record[rate_field] = None
            else:
                weekly_record[rate_field] = round((numerator / denominator) * 100, 1)

        output.append(weekly_record)

    return output


###########################################################################
# FOLDER PROCESSING
###########################################################################

def process_weekly_folder(folder_path, site_name):
    """
    Reads every .xlsx file in `folder_path` (one file per week, same site),
    extracts each with extract_report(), groups the resulting records by
    (site, survey_type), and computes weekly deltas within each group.

    Returns a flat list of records (dicts), each with both the original
    cumulative fields and the added weekly_<field> values.
    """

    folder = Path(folder_path)

    all_records = []

    for filepath in sorted(folder.glob("*.xlsx")):

        all_records.extend(extract_report(str(filepath), site_name))

    groups = {}

    for record in all_records:
        groups.setdefault(group_key(record), []).append(record)

    results = []

    for key, group_records in groups.items():
        results.extend(compute_weekly_deltas(group_records))

    return results


###########################################################################
# DATAFRAME ENTRY POINT (for use with the database / load_history() flow)
###########################################################################

def add_weekly_columns(df):
    """
    Takes the full historical DataFrame (as returned by database.load_history())
    and returns a new DataFrame with weekly_<field> and weekly_<rate> columns
    added, computed across each (site, survey_type) group's full history.

    This is the entry point meant to be called from
    WeeklyReportConsolidator.export_history() -- deltas need the FULL
    history, not just the current run's batch, since this week's delta
    depends on last week's cumulative value which may have been inserted
    into the database on a previous run.
    """

    if df.empty:
        return df

    records = df.to_dict("records")

    #######################################################################
    # Dedup safeguard
    #
    # If the same (site, survey_type, week) somehow ended up in the
    # database more than once -- e.g. a report got reprocessed before a
    # new week's file existed -- diffing against a duplicate would corrupt
    # the delta (report a real week as 0, or double-count it depending on
    # ordering). Collapse duplicates on (site, survey_type, date), keeping
    # the LAST occurrence under the assumption that a later insert is a
    # correction/re-run of the same week's data, not a genuinely new week.
    #
    # NOTE: verify this assumption against how insert_dataframe() actually
    # behaves in database.py -- if it already enforces uniqueness on
    # insert, this safeguard is a no-op and harmless either way.
    #######################################################################

    deduped = {}

    for record in records:
        key = (group_key(record), sort_key(record))
        deduped[key] = record  # later occurrence wins

    groups = {}

    for record in deduped.values():
        groups.setdefault(group_key(record), []).append(record)

    output = []

    for key, group_records in groups.items():
        output.extend(compute_weekly_deltas(group_records))

    return pd.DataFrame(output)


###########################################################################
# TEST
###########################################################################

if __name__ == "__main__":

    weekly_records = process_weekly_folder(
        "weekly_reports/Johannesburg",  # <-- folder containing one .xlsx per week
        "Johannesburg"
    )

    for r in weekly_records:
        print(
            r.get("survey_type"),
            r.get("report_start_date"),
            "completed(cum)=", r.get("completed"),
            "completed(weekly)=", r.get("weekly_completed"),
            r.get("_warnings", "")
        )