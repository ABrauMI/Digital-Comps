"""Run on a schedule (e.g. hourly, via Railway's Cron Job): checks every
source sheet in the Drive folder, regenerates and posts to Slack any
report whose source data changed since the last run."""
import os

from slack_sdk import WebClient

from drive_client import DriveClient
from pipeline import generate_report
from state_store import load_state, save_state

FOLDER_ID = os.environ["DRIVE_FOLDER_ID"]
CHANNEL = os.environ["SLACK_DEFAULT_CHANNEL"]


def main():
    drive = DriveClient()
    slack = WebClient(token=os.environ["SLACK_BOT_TOKEN"])
    state = load_state()

    files = drive.list_source_files(FOLDER_ID)
    changed = 0
    for f in files:
        if state.get(f["id"]) == f["modifiedTime"]:
            continue
        print(f"Change detected: {f['name']}")
        xlsx_path, title = generate_report(drive, f["id"], f["name"])
        slack.files_upload_v2(
            channel=CHANNEL,
            file=xlsx_path,
            title=f"{title} Digital Competitive Report",
            initial_comment=f"Updated comp for *{title}* — source data changed.",
        )
        state[f["id"]] = f["modifiedTime"]
        changed += 1

    save_state(state)
    print(f"Checked {len(files)} source file(s), {changed} updated.")


if __name__ == "__main__":
    main()
