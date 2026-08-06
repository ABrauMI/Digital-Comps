"""Google Drive + Sheets access for the Ad Hawk source folder.

Reads sheet data straight through the Sheets API (values.get), so there's
no markdown-table parsing involved -- each row comes back as a clean dict
keyed by its header row.
"""
import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build as build_google_client

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]


def _load_credentials():
    raw = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(raw)
    return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)


class DriveClient:
    def __init__(self):
        creds = _load_credentials()
        self.drive = build_google_client("drive", "v3", credentials=creds)
        self.sheets = build_google_client("sheets", "v4", credentials=creds)

    def list_source_files(self, folder_id):
        """Every Google Sheet directly inside the source folder, including
        shortcuts to sheets that live elsewhere (resolved to the real
        target's id/modifiedTime, since that's what actually needs reading
        and change-tracking -- a shortcut's own modifiedTime barely moves)."""
        query = (
            f"'{folder_id}' in parents and trashed=false and "
            "(mimeType='application/vnd.google-apps.spreadsheet' "
            "or mimeType='application/vnd.google-apps.shortcut')"
        )
        raw_files, page_token = [], None
        while True:
            resp = self.drive.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, shortcutDetails)",
                pageToken=page_token,
            ).execute()
            raw_files.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break

        resolved = []
        for f in raw_files:
            if f["mimeType"] == "application/vnd.google-apps.shortcut":
                details = f.get("shortcutDetails", {})
                if details.get("targetMimeType") != "application/vnd.google-apps.spreadsheet":
                    continue  # shortcut to something that isn't a sheet
                target = self.drive.files().get(
                    fileId=details["targetId"], fields="id, name, modifiedTime"
                ).execute()
                resolved.append({"id": target["id"], "name": f["name"], "modifiedTime": target["modifiedTime"]})
            else:
                resolved.append({"id": f["id"], "name": f["name"], "modifiedTime": f["modifiedTime"]})
        return resolved

    def get_file_name(self, file_id):
        return self.drive.files().get(fileId=file_id, fields="name").execute()["name"]

    def get_modified_time(self, file_id):
        return self.drive.files().get(fileId=file_id, fields="modifiedTime").execute()["modifiedTime"]

    def fetch_sheet_data(self, file_id):
        """Returns (creative_rows, spend_rows) as lists of dicts, detected by
        header shape rather than by tab name (Ad Hawk's tab names aren't
        guaranteed consistent across races)."""
        meta = self.sheets.spreadsheets().get(spreadsheetId=file_id).execute()
        creative_rows, spend_rows = [], []
        for sheet in meta["sheets"]:
            title = sheet["properties"]["title"]
            # A1-notation ranges need a sheet title with spaces/special chars
            # single-quoted, or the API silently returns no values instead of
            # erroring.
            quoted_title = "'" + title.replace("'", "''") + "'"
            result = self.sheets.spreadsheets().values().get(
                spreadsheetId=file_id, range=quoted_title
            ).execute()
            values = result.get("values", [])
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
