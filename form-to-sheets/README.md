# Form to Google Sheets

Every form submission automatically becomes **a new row in a Google Sheet**, no copy-pasting. You can also get a Discord ping for each new request.

This demo is a small web app with a quote-request form. It can also receive submissions from other tools (a website, Tally, Typeform, Webflow, a webshop) through a JSON webhook.

```
 Website form  ─┐
                ├──►  validate + clean  ──►  Google Sheet (new row)
 Other tools   ─┘     block spam bots        └─► Discord notification (optional)
 (JSON webhook)
```

## Features

- A clean, mobile-friendly form at `http://localhost:5050`
- A `/webhook` endpoint that accepts JSON from any other tool, protected with a shared secret
- Validation with friendly error messages (name, email, phone, service, message length)
- Cleans the input: trims spaces, lowercases emails, caps field lengths
- A honeypot field that silently drops spam bots
- Writes values as plain text, so input like `=IMPORTXML(...)` can't run as a formula in the sheet
- Creates the worksheet and a bold header row automatically on first use
- **Works without any setup**: if Google isn't configured yet, rows go to `output/submissions.csv`

Example of the resulting rows:

| Received | Name | Email | Phone | Service | Budget | Message | Source |
|---|---|---|---|---|---|---|---|
| 2026-09-25 13:47:27 | Sophie Bakker | sophie@example.com | +31 6 12345678 | Spreadsheet cleanup | €50 - €150 | Every Monday I merge three webshop exports by hand... | test script |
| 2026-09-25 13:47:27 | Lucas Peeters | lucas@example.com | | Web scraping | €150 - €500 | I need a weekly price list from two public supplier catalogs. | test script |
| 2026-09-25 13:47:31 | Emma de Vries | emma@example.com | | Automation / integration | | Connect my Shopify orders to a sheet please | website form |

## Run it (no Google account needed)

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt

python main.py
```

Open http://localhost:5050 and submit the form. Rows appear in `output/submissions.csv`.

To send a few example submissions (including one invalid one), run this in a second terminal:

```bash
python send_test_submissions.py
```

```
201 {'status': 'saved'}
201 {'status': 'saved'}
422 {'error': 'validation failed', 'fields': {'email': 'Please enter a valid email address.', ...}}
```

## Connect a real Google Sheet

1. Go to [Google Cloud Console](https://console.cloud.google.com/), create a project, and enable the **Google Sheets API** and **Google Drive API**.
2. Under *IAM & Admin → Service Accounts*, create a service account. Then open it, go to *Keys → Add key → JSON*, and save the file in this folder as `credentials.json`.
3. Create a Google Sheet and **share it** (Editor) with the service account's email address. It looks like `name@project.iam.gserviceaccount.com`.
4. Copy `.env.example` to `.env` and fill in `GOOGLE_SHEET_ID`. That's the long ID in the sheet's URL: `docs.google.com/spreadsheets/d/<GOOGLE_SHEET_ID>/edit`.
5. Run `python main.py`. It should print `Saving submissions to: Google Sheet '...'`.

`credentials.json` and `.env` are in `.gitignore`. Never commit them.

### Optional settings (`.env`)

| Setting | What it does |
|---|---|
| `WORKSHEET_NAME` | Tab name to write to (default `Submissions`, created if missing) |
| `WEBHOOK_SECRET` | If set, `/webhook` requires the header `X-Webhook-Secret: <value>` |
| `DISCORD_WEBHOOK_URL` | Posts a message in a Discord channel for every new request (*Channel settings → Integrations → Webhooks*) |
| `PORT` | Server port (default 5050, because macOS uses 5000 for AirPlay) |

## Sending data from another tool

```bash
curl -X POST http://localhost:5050/webhook \
  -H "Content-Type: application/json" \
  -H "X-Webhook-Secret: your-secret" \
  -d '{"name": "Anna", "email": "anna@example.com", "service": "Web scraping",
       "message": "Need a list of public events every week", "source": "Tally"}'
```

| Response | Meaning |
|---|---|
| `201` | Saved |
| `401` | Wrong or missing secret |
| `422` | Validation failed; the response says which fields |

## Going live

Running locally is great for testing. For real use, the app can be deployed as-is to a small host such as Render, Railway or PythonAnywhere. Set the same values as environment variables there and upload `credentials.json` as a secret file.

## Files

```
form-to-sheets/
├── main.py                    # the web app + Google Sheets logic
├── send_test_submissions.py   # posts example data to the running app
├── templates/                 # form and thank-you page
├── static/style.css
├── requirements.txt
├── .env.example               # copy to .env and fill in
└── .gitignore                 # keeps credentials out of git
```
