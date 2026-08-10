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


def generate_report(drive_client, file_id, file_name=None):
    """Pulls the sheet, builds the report, returns (xlsx_path, title)."""
    if file_name is None:
        file_name = drive_client.get_file_name(file_id)

    creative_rows, spend_rows = drive_client.fetch_sheet_data(file_id)
    creative_clean = _clean_creative_rows(creative_rows)
    spend_clean = _clean_spend_rows(spend_rows)

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
