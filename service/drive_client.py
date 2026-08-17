"""Google Drive + Sheets access for the Ad Hawk source folder.

Reads sheet data straight through the Sheets API (values.batchGet), so
there's no markdown-table parsing involved -- each row comes back as a
clean dict keyed by its header row.
"""
import json
import os
import random
import time

from google.oauth2 import service_account
from googleapiclient.discovery import build as build_google_client
from googleapiclient.errors import HttpError

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]


def _load_credentials():
    raw = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(raw)
    return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)


def _excluded_names():
    """Source files to skip entirely, by case-insensitive name substring --
    e.g. multi-race rollup sheets whose numbers don't match the standalone
    per-race sheets (see "Anchor Raw Data"). Configured via
    EXCLUDED_SOURCE_NAMES, comma-separated."""
    raw = os.environ.get("EXCLUDED_SOURCE_NAMES", "")
    return [n.strip().lower() for n in raw.split(",") if n.strip()]


def _retry(request, max_attempts=5, base_delay=2.0):
    """Execute a googleapiclient request, retrying with exponential backoff
    on HTTP 429 (rate limit). Drive/Sheets quotas are per-minute, so a
    short wait and retry recovers cleanly instead of failing the file."""
    for attempt in range(max_attempts):
        try:
            return request.execute()
        except HttpError as e:
            if e.resp.status == 429 and attempt < max_attempts - 1:
                time.sleep(base_delay * (2 ** attempt) + random.uniform(0, 1))
                continue
            raise


class DriveClient:
    def __init__(self):
        creds = _load_credentials()
        self.drive = build_google_client("drive", "v3", credentials=creds)
        self.sheets = build_google_client("sheets", "v4", credentials=creds)

    def list_source_files(self, folder_ids):
        """Every Google Sheet across one or more source folders. folder_ids
        is a single Drive folder id, or several comma-separated (e.g.
        DRIVE_FOLDER_ID="folder1,folder2") -- lets the bot watch more than
        one folder (say, single-race sheets in one and multi-district
        rollups in another) without having to merge them in Drive itself.
        Results are de-duplicated by resolved file id in case the same
        sheet is reachable from more than one folder."""
        ids = [f.strip() for f in folder_ids.split(",") if f.strip()]
        seen, merged = set(), []
        for folder_id in ids:
            for f in self._list_source_files_one(folder_id):
                if f["id"] not in seen:
                    seen.add(f["id"])
                    merged.append(f)
        return merged

    def _list_source_files_one(self, folder_id):
        """Every Google Sheet directly inside one source folder, including
        shortcuts to sheets that live elsewhere (resolved to the real
        target's id, since that's what actually needs reading)."""
        query = (
            f"'{folder_id}' in parents and trashed=false and "
            "(mimeType='application/vnd.google-apps.spreadsheet' "
            "or mimeType='application/vnd.google-apps.shortcut')"
        )
        raw_files, page_token = [], None
        while True:
            resp = _retry(self.drive.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, shortcutDetails)",
                pageToken=page_token,
            ))
            raw_files.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break

        excluded = _excluded_names()
        raw_files = [f for f in raw_files if not any(n in f["name"].lower() for n in excluded)]

        resolved = []
        for f in raw_files:
            if f["mimeType"] == "application/vnd.google-apps.shortcut":
                details = f.get("shortcutDetails", {})
                if details.get("targetMimeType") != "application/vnd.google-apps.spreadsheet":
                    continue  # shortcut to something that isn't a sheet
                target = _retry(self.drive.files().get(
                    fileId=details["targetId"], fields="id, name, modifiedTime"
                ))
                resolved.append({"id": target["id"], "name": f["name"], "modifiedTime": target["modifiedTime"]})
            else:
                resolved.append({"id": f["id"], "name": f["name"], "modifiedTime": f["modifiedTime"]})
        return resolved

    def get_file_name(self, file_id):
        return _retry(self.drive.files().get(fileId=file_id, fields="name"))["name"]

    def get_modified_time(self, file_id):
        return _retry(self.drive.files().get(fileId=file_id, fields="modifiedTime"))["modifiedTime"]

    def fetch_sheet_data(self, file_id):
        """Returns (creative_rows, spend_rows) as lists of dicts, detected by
        header shape rather than by tab name (Ad Hawk's tab names aren't
        guaranteed consistent across races). Uses a single batchGet for all
        tabs' values instead of one values.get call per tab, since each
        source file is now read on every scheduled check rather than only
        when Drive's own modifiedTime changes -- fewer requests per file
        matters a lot more than it used to."""
        meta = _retry(self.sheets.spreadsheets().get(spreadsheetId=file_id))
        titles = [s["properties"]["title"] for s in meta["sheets"]]
        # A1-notation ranges need a sheet title with spaces/special chars
        # single-quoted, or the API silently returns no values instead of
        # erroring.
        quoted_titles = ["'" + t.replace("'", "''") + "'" for t in titles]
        result = _retry(self.sheets.spreadsheets().values().batchGet(
            spreadsheetId=file_id, ranges=quoted_titles
        ))

        creative_rows, spend_rows = [], []
        for value_range in result.get("valueRanges", []):
            values = value_range.get("values", [])
            if not values:
                continue
            header = [h.strip().lower() for h in values[0]]
            # some exports (e.g. multi-race rollup sheets) use "platform"
            # where single-race sheets use "source_platform" -- normalize so
            # every downstream row has the same key regardless of source.
            header = ["source_platform" if h == "platform" else h for h in header]
            records = []
            for row in values[1:]:
                padded = row + [""] * (len(header) - len(row))
                records.append(dict(zip(header, padded)))
            if "transcript" in header and "first_ran" in header:
                creative_rows = records
            elif "target_date" in header:
                spend_rows = records
        return creative_rows, spend_rows
