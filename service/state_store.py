"""Tracks the last-seen modifiedTime per source file, so the scheduled
checker knows what's actually new. Stored on a Railway volume so it
survives between cron runs."""
import json
import os

STATE_PATH = os.environ.get("STATE_PATH", "/data/state.json")


def load_state():
    if not os.path.exists(STATE_PATH):
        return {}
    with open(STATE_PATH) as f:
        return json.load(f)


def save_state(state):
    os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
    with open(STATE_PATH, "w") as f:
        json.dump(state, f, indent=1)
