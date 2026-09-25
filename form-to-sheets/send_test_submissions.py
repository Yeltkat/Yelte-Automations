"""
Sends a few example submissions to the running app, so you can see rows
appear in the sheet without typing them in by hand.

Usage (with main.py running in another terminal):
  python send_test_submissions.py
"""

import os

import requests
from dotenv import load_dotenv

load_dotenv()
URL = f"http://127.0.0.1:{os.getenv('PORT', '5050')}/webhook"
HEADERS = {"X-Webhook-Secret": os.getenv("WEBHOOK_SECRET", "")}

EXAMPLES = [
    {"name": "Sophie Bakker", "email": "sophie@example.com", "phone": "+31 6 12345678",
     "service": "Spreadsheet cleanup", "budget": "€50 - €150",
     "message": "Every Monday I merge three webshop exports by hand. Can this be automated?",
     "source": "test script"},
    {"name": "Lucas Peeters", "email": "LUCAS@EXAMPLE.COM ",
     "service": "Web scraping", "budget": "€150 - €500",
     "message": "I need a weekly price list from two public supplier catalogs.",
     "source": "test script"},
    # This one is invalid on purpose, to show the validation
    {"name": "X", "email": "not-an-email", "service": "Something else", "message": "hi"},
]

for example in EXAMPLES:
    response = requests.post(URL, json=example, headers=HEADERS, timeout=10)
    print(response.status_code, response.json())
