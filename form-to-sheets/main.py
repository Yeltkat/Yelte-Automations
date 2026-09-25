"""
Form to Google Sheets: every form submission becomes a new row in a spreadsheet.

What it does:
  - Serves a simple quote-request form at http://localhost:5050
  - Accepts submissions from that form, or as JSON on /webhook
    (so a website, Tally, Typeform, Webflow, a webshop, ... can post to it too)
  - Validates and cleans the input, blocks simple spam bots
  - Appends a row to a Google Sheet (or to output/submissions.csv when
    Google isn't configured yet, so you can try it right away)
  - Optionally posts a notification to a Discord channel

Usage:
  python main.py
  Then open http://localhost:5050 and submit the form.
  See README.md for the Google Sheets setup.
"""

import csv
import os
import re
import threading
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, redirect, render_template, request, url_for

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / ".env")

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
GOOGLE_CREDENTIALS_FILE = BASE_DIR / os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME", "Submissions")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "")
DISCORD_WEBHOOK_URL = os.getenv("DISCORD_WEBHOOK_URL", "")
LOCAL_CSV = BASE_DIR / "output" / "submissions.csv"
# Not 5000: on macOS that port is taken by the AirPlay Receiver
PORT = int(os.getenv("PORT", "5050"))

COLUMNS = ["Received", "Name", "Email", "Phone", "Service", "Budget", "Message", "Source"]
SERVICES = ["Spreadsheet cleanup", "Web scraping", "Automation / integration", "Other"]
BUDGETS = ["< €50", "€50 - €150", "€150 - €500", "€500+"]
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$")


# ---------------------------------------------------------------------------
# Storage: Google Sheets, or a local CSV as a fallback
# ---------------------------------------------------------------------------

class GoogleSheetStorage:
    def __init__(self, sheet_id, credentials_file, worksheet_name):
        import gspread  # only needed when Google Sheets is actually used

        client = gspread.service_account(filename=str(credentials_file))
        spreadsheet = client.open_by_key(sheet_id)
        try:
            self.worksheet = spreadsheet.worksheet(worksheet_name)
        except gspread.WorksheetNotFound:
            self.worksheet = spreadsheet.add_worksheet(worksheet_name, rows=1000, cols=len(COLUMNS))
        if not self.worksheet.row_values(1):
            self.worksheet.append_row(COLUMNS)
            self.worksheet.format("1:1", {"textFormat": {"bold": True}})
        self.description = f"Google Sheet '{spreadsheet.title}' / {worksheet_name}"

    def append(self, row):
        # RAW stops Sheets from interpreting input like "=HYPERLINK(...)" as a formula
        self.worksheet.append_row(row, value_input_option="RAW")


class LocalCSVStorage:
    def __init__(self, path):
        self.path = path
        self.lock = threading.Lock()
        self.description = f"local file {path.relative_to(BASE_DIR)}"

    def append(self, row):
        with self.lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            is_new = not self.path.exists()
            with open(self.path, "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                if is_new:
                    writer.writerow(COLUMNS)
                writer.writerow(row)


def create_storage():
    if GOOGLE_SHEET_ID and GOOGLE_CREDENTIALS_FILE.exists():
        return GoogleSheetStorage(GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE, WORKSHEET_NAME)
    print("Google Sheets not configured (see README), saving to a local CSV instead.")
    return LocalCSVStorage(LOCAL_CSV)


# ---------------------------------------------------------------------------
# Validation and notifications
# ---------------------------------------------------------------------------

def clean_submission(data):
    """Return (row, errors). Accepts form data or JSON with the same field names."""
    def field(name, max_len=200):
        return re.sub(r"\s+", " ", str(data.get(name, "") or "")).strip()[:max_len]

    submission = {
        "name": field("name", 100),
        "email": field("email", 150).lower(),
        "phone": field("phone", 30),
        "service": field("service"),
        "budget": field("budget"),
        "message": str(data.get("message", "") or "").strip()[:2000],
    }

    errors = {}
    if len(submission["name"]) < 2:
        errors["name"] = "Please enter your name."
    if not EMAIL_RE.match(submission["email"]):
        errors["email"] = "Please enter a valid email address."
    if submission["phone"] and not re.fullmatch(r"[+\d][\d\s\-()]{6,}", submission["phone"]):
        errors["phone"] = "Please enter a valid phone number, or leave it empty."
    if submission["service"] not in SERVICES:
        errors["service"] = "Please choose a service."
    if submission["budget"] and submission["budget"] not in BUDGETS:
        errors["budget"] = "Please choose a budget from the list."
    if len(submission["message"]) < 10:
        errors["message"] = "Please describe what you need in a few words."
    return submission, errors


def to_row(submission, source):
    return [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        submission["name"],
        submission["email"],
        submission["phone"],
        submission["service"],
        submission["budget"],
        submission["message"],
        source,
    ]


def notify_discord(submission):
    """Post a short message to a Discord channel. Failures never block the submission."""
    if not DISCORD_WEBHOOK_URL:
        return
    content = (f"**New request: {submission['service']}**\n"
               f"{submission['name']} ({submission['email']})"
               f"{' | budget ' + submission['budget'] if submission['budget'] else ''}\n"
               f"> {submission['message'][:300]}")
    try:
        requests.post(DISCORD_WEBHOOK_URL, json={"content": content}, timeout=10)
    except requests.RequestException as exc:
        print(f"Discord notification failed: {exc}")


# ---------------------------------------------------------------------------
# Web app
# ---------------------------------------------------------------------------

def create_app(storage=None):
    app = Flask(__name__)
    app.config["storage"] = storage or create_storage()

    def save(submission, source):
        app.config["storage"].append(to_row(submission, source))
        notify_discord(submission)

    @app.get("/")
    def form():
        return render_template("form.html", services=SERVICES, budgets=BUDGETS,
                               values={}, errors={})

    @app.post("/submit")
    def submit():
        # Honeypot: this field is hidden from humans, so only bots fill it in
        if request.form.get("website"):
            return redirect(url_for("thanks"))
        submission, errors = clean_submission(request.form)
        if errors:
            return render_template("form.html", services=SERVICES, budgets=BUDGETS,
                                   values=submission, errors=errors), 400
        save(submission, "website form")
        return redirect(url_for("thanks"))

    @app.get("/thanks")
    def thanks():
        return render_template("thanks.html")

    @app.post("/webhook")
    def webhook():
        """For other tools posting JSON: {"name": ..., "email": ..., "service": ..., "message": ...}"""
        if WEBHOOK_SECRET and request.headers.get("X-Webhook-Secret") != WEBHOOK_SECRET:
            return jsonify(error="unauthorized"), 401
        data = request.get_json(silent=True)
        if not isinstance(data, dict):
            return jsonify(error="expected a JSON object"), 400
        submission, errors = clean_submission(data)
        if errors:
            return jsonify(error="validation failed", fields=errors), 422
        save(submission, str(data.get("source") or "webhook")[:50])
        return jsonify(status="saved"), 201

    return app


if __name__ == "__main__":
    app = create_app()
    print(f"Saving submissions to: {app.config['storage'].description}")
    print(f"Open http://localhost:{PORT} to fill in the form")
    app.run(host="127.0.0.1", port=PORT, debug=False)
