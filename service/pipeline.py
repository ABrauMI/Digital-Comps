"""Turn a Drive file_id into a finished .xlsx report."""
import datetime
import json
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from scripts.build_report import build  # noqa: E402

REQUIRED_SPEND_FIELDS = ["spend", "sentiment", "language", "updated_at"]


class NoDataError(ValueError):
    """Raised when a source sheet has no usable rows yet (e.g. a race with
    no tracked ad spend so far) -- expected and skippable, not a crash."""


def derive_title(file_name):
    title = re.sub(r"raw\s*data", "", file_name, flags=re.IGNORECASE).strip()
    return title or file_name


def _clean_creative_rows(rows):
    cleaned = []
    for r in rows:
        try:
            r["spend"] = float(r["spend"])
            r["first_ran"] = str(datetime.date.fromisoformat(r["first_ran"][:10]))
            r["last_ran"] = str(datetime.date.fromisoformat(r["last_ran"][:10]))
        except (KeyError, ValueError):
            continue
        cleaned.append(r)
    return cleaned


def _clean_spend_rows(rows):
    cleaned = []
    for r in rows:
        if not all(r.get(f) not in (None, "") for f in REQUIRED_SPEND_FIELDS):
            continue
        try:
            r["spend"] = float(r["spend"])
            r["target_date"] = str(datetime.date.fromisoformat(r["target_date"][:10]))
        except (KeyError, ValueError):
            continue
        cleaned.append(r)
    return cleaned


def latest_updated_at(spend_rows):
    """Newest updated_at timestamp among a sheet's (raw, uncleaned) daily-
    spend rows, or None if there's nothing to compare yet. These are plain
    "YYYY-MM-DD HH:MM:SS" strings, which sort correctly as-is. This -- not
    Drive's own modifiedTime -- is the change-detection signal: sheets fed
    by an IMPORTRANGE-style formula pull don't reliably bump modifiedTime
    when the source data refreshes, but Ad Hawk does stamp each row with
    when it was actually written."""
    stamps = [r["updated_at"] for r in spend_rows if r.get("updated_at")]
    return max(stamps) if stamps else None


def _has_two_party_race(spend_rows):
    party_total = {}
    for r in spend_rows:
        party_total[r["party"]] = party_total.get(r["party"], 0.0) + r["spend"]
    return party_total.get("D", 0) > 0 and party_total.get("R", 0) > 0


def _date_floors():
    """Per-race data cutoffs, e.g. a client only wanting to see spend after
    a contested primary resolved. Configured via RACE_DATE_FLOORS as
    comma-separated 'name substring:YYYY-MM-DD' pairs, e.g.
    'MT-01:2026-06-11'. Matching is a case-insensitive substring against
    the source file's name."""
    raw = os.environ.get("RACE_DATE_FLOORS", "")
    floors = []
    for part in raw.split(","):
        part = part.strip()
        if not part or ":" not in part:
            continue
        name, date_str = part.rsplit(":", 1)
        try:
            floors.append((name.strip().lower(), date_str.strip()))
        except ValueError:
            continue
    return floors


def _date_floor_for(file_name):
    name_lower = file_name.lower()
    for substr, floor_str in _date_floors():
        if substr in name_lower:
            return floor_str
    return None


def _race_excluded_advertisers():
    """Per-race advertiser exclusions -- e.g. keeping a generic
    institutional advertiser out of one race's report without hiding it
    everywhere. Configured via RACE_EXCLUDED_ADVERTISERS as comma-separated
    'race substring:advertiser1|advertiser2' groups, e.g.
    'MT-01:MT WAY|MT SECRETARY OF STATE, PA-GOV:SOME PAC'. Race matching is
    a case-insensitive substring against the source file's name; advertiser
    matching is an exact (case-insensitive) match against the advertiser
    column, not a substring -- so excluding one committee doesn't
    accidentally also swallow a similarly-named one."""
    raw = os.environ.get("RACE_EXCLUDED_ADVERTISERS", "")
    rules = []
    for part in raw.split(","):
        part = part.strip()
        if not part or ":" not in part:
            continue
        race_substr, advertisers = part.split(":", 1)
        names = {a.strip().lower() for a in advertisers.split("|") if a.strip()}
        if names:
            rules.append((race_substr.strip().lower(), names))
    return rules


def _excluded_advertisers_for(file_name):
    name_lower = file_name.lower()
    excluded = set()
    for race_substr, advertiser_names in _race_excluded_advertisers():
        if race_substr in name_lower:
            excluded |= advertiser_names
    return excluded


def generate_report(drive_client, file_id, file_name=None, creative_rows=None, spend_rows=None):
    """Pulls the sheet (unless rows are already provided -- e.g. the caller
    just read them to check updated_at, no need to fetch twice), builds the
    report, returns (xlsx_path, title)."""
    if file_name is None:
        file_name = drive_client.get_file_name(file_id)

    if creative_rows is None and spend_rows is None:
        creative_rows, spend_rows = drive_client.fetch_sheet_data(file_id)

    creative_clean = _clean_creative_rows(creative_rows)
    spend_clean = _clean_spend_rows(spend_rows)

    floor_str = _date_floor_for(file_name)
    if floor_str:
        creative_clean = [r for r in creative_clean if r["first_ran"] >= floor_str]
        spend_clean = [r for r in spend_clean if r["target_date"] >= floor_str]

    excluded_advs = _excluded_advertisers_for(file_name)
    if excluded_advs:
        creative_clean = [r for r in creative_clean if r["advertiser"].strip().lower() not in excluded_advs]
        spend_clean = [r for r in spend_clean if r["advertiser"].strip().lower() not in excluded_advs]

    if not creative_clean and not spend_clean:
        raise NoDataError(f"'{file_name}' has no usable data rows yet -- nothing to report.")

    title = derive_title(file_name)
    include_dvr = _has_two_party_race(spend_clean)

    tmpdir = tempfile.mkdtemp(prefix="digicomp_")
    json.dump(creative_clean, open(os.path.join(tmpdir, "raw_creative_clean.json"), "w"))
    json.dump(spend_clean, open(os.path.join(tmpdir, "daily_spend_clean.json"), "w"))

    output_path = os.path.join(tmpdir, f"{title.replace(' ', '_')}_Digital_Competitive_Report.xlsx")
    build(tmpdir, title, output_path, datetime.date.today(), include_dvr)
    return output_path, title
