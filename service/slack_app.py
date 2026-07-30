"""Always-on Socket Mode listener for the /digicomp slash command:
prompts for a source file, generates the comp, and posts it back."""
import os

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

from drive_client import DriveClient
from pipeline import generate_report

FOLDER_ID = os.environ["DRIVE_FOLDER_ID"]

app = App(token=os.environ["SLACK_BOT_TOKEN"], signing_secret=os.environ["SLACK_SIGNING_SECRET"])
drive = DriveClient()


@app.command("/digicomp")
def handle_digicomp(ack, body, client):
    ack()
    files = drive.list_source_files(FOLDER_ID)
    options = [
        {"text": {"type": "plain_text", "text": f["name"]}, "value": f["id"]}
        for f in files
    ]
    if not options:
        client.chat_postEphemeral(
            channel=body["channel_id"], user=body["user_id"],
            text="No source files found in the Ad Hawk Drive folder."
        )
        return

    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "digicomp_select",
            "private_metadata": body["channel_id"],
            "title": {"type": "plain_text", "text": "Generate Comp"},
            "submit": {"type": "plain_text", "text": "Generate"},
            "blocks": [
                {
                    "type": "input",
                    "block_id": "source_block",
                    "label": {"type": "plain_text", "text": "Choose a source file"},
                    "element": {
                        "type": "static_select",
                        "action_id": "source_select",
                        "options": options,
                    },
                }
            ],
        },
    )


@app.view("digicomp_select")
def handle_submission(ack, view, client):
    ack()
    channel_id = view["private_metadata"]
    file_id = view["state"]["values"]["source_block"]["source_select"]["selected_option"]["value"]

    client.chat_postMessage(channel=channel_id, text="Generating your comp, one moment…")
    xlsx_path, title = generate_report(drive, file_id)
    client.files_upload_v2(
        channel=channel_id,
        file=xlsx_path,
        title=f"{title} Digital Competitive Report",
        initial_comment=f"Here's the latest comp for *{title}*.",
    )


if __name__ == "__main__":
    handler = SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"])
    handler.start()
