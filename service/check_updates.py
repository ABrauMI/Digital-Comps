"""Run on a schedule (e.g. 8am/3pm ET, via Railway's Cron Job): checks
every source sheet in the Drive folder, and for any whose data changed,
posts one "Comp Drop" summary message with each updated report attached
as a threaded reply underneath it (rather than one message per file).

Change detection is based on the data's own "updated_at" column, not
Drive's file-level modifiedTime -- sheets fed by an IMPORTRANGE-style
formula pull recalculate without registering as a "modification" in
Drive's revision history, so modifiedTime alone misses real updates."""
import os
import time
from datetime import datetime
from zoneinfo import ZoneInfo

from slack_sdk import WebClient

from drive_client import DriveClient
from pipeline import generate_report, latest_updated_at, NoDataError
from state_store import load_state, save_state

FOLDER_ID = os.environ["DRIVE_FOLDER_ID"]
CHANNEL = os.environ["SLACK_DEFAULT_CHANNEL"]
# Spread requests out to stay well under Sheets' per-minute read quota
# rather than firing every file's read back-to-back.
PAUSE_BETWEEN_FILES_SECONDS = 1.5


def main():
    drive = DriveClient()
    slack = WebClient(token=os.environ["SLACK_BOT_TOKEN"])
    state = load_state()

    files = drive.list_source_files(FOLDER_ID)
    updated, skipped, failed = [], 0, 0

    for f in files:
        time.sleep(PAUSE_BETWEEN_FILES_SECONDS)
        try:
            creative_rows, spend_rows = drive.fetch_sheet_data(f["id"])
        except Exception as e:
            print(f"ERROR reading {f['name']}: {e}")
            failed += 1
            continue

        latest = latest_updated_at(spend_rows)
        if latest is None or state.get(f["id"]) == latest:
            continue  # no data yet, or no new data since last check

        print(f"Change detected: {f['name']} (updated_at {latest})")
        try:
            xlsx_path, title = generate_report(
                drive, f["id"], f["name"], creative_rows=creative_rows, spend_rows=spend_rows
            )
        except NoDataError as e:
            print(f"Skipping {f['name']}: {e}")
            skipped += 1
            state[f["id"]] = latest
            save_state(state)
            continue
        except Exception as e:
            print(f"ERROR generating {f['name']}: {e}")
            failed += 1
            continue  # leave state alone so this file gets retried next run
        updated.append({"file_id": f["id"], "latest": latest, "title": title, "xlsx_path": xlsx_path})

    if updated:
        now_et = datetime.now(ZoneInfo("America/New_York"))
        label = now_et.strftime("Comp Drop %m/%d/%y %-I%p").replace("AM", "am").replace("PM", "pm")
        summary = f"*{label}*\n" + "\n".join(f"• {u['title']}" for u in updated)
        parent = slack.chat_postMessage(channel=CHANNEL, text=summary)
        thread_ts = parent["ts"]

        for u in updated:
            try:
                slack.files_upload_v2(
                    channel=CHANNEL,
                    thread_ts=thread_ts,
                    file=u["xlsx_path"],
                    title=f"{u['title']} Digital Competitive Report",
                )
            except Exception as e:
                print(f"ERROR posting {u['title']}: {e}")
                failed += 1
                continue
            state[u["file_id"]] = u["latest"]
            save_state(state)

    print(f"Checked {len(files)} source file(s): {len(updated)} updated, {skipped} skipped (no data), {failed} failed.")


if __name__ == "__main__":
    main()
