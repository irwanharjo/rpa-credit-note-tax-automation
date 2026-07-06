# Credit Note & Tax (PPN) Recalculation Automation

An end-to-end automation for a recurring billing-reconciliation workflow: recalculating VAT/PPN after a credit note is applied, cross-checking disputed line items against a web claims portal, and notifying the client by email — all from a batch of incoming Excel spreadsheets. Genericized for portfolio purposes (client identifiers, portal URL, spreadsheet ID, and email addresses have been replaced with placeholders/environment variables).

## What it does

1. **Ingests** a batch of billing spreadsheets, auto-detecting the header row and relevant columns regardless of minor formatting differences.
2. **Recalculates PPN (VAT)** per line item when a credit note affects a "farmasi" (pharmacy) charge, correctly netting negative credit-note items against the largest matching charge in the same block before recomputing DPP/PPN — a real edge case in Indonesian tax-invoice adjustments.
3. **Detects disputed items**: rows flagged "not pre-authorized" are cross-checked live against a web claims portal (headless Selenium) to see if the patient is actually a VVIP member exempt from pre-authorization. If so, the whole invoice is held and a dispute email (with screenshot evidence) is sent instead of being processed normally.
4. **Formats and exports** a clean, bordered, landscape-oriented PDF of the recalculated invoice via `win32com` (driving a real Excel instance for pixel-perfect print output).
5. **Emails the result** via the Gmail API — replying in the same thread as the original billing email when one exists, otherwise starting a new one, with the PDF attached.
6. **Logs a daily recap** spreadsheet of every invoice processed that day, and appends normal (non-disputed) PPN requests to a shared Google Sheet used for tax-invoice tracking, without disturbing existing formulas.

## Setup

```bash
pip install -r requirements.txt
copy .env.example .env   # then fill in your own values
```

You'll also need:
- A Google Cloud OAuth `credentials.json` (Gmail + Sheets API enabled) placed in this folder — never committed, see `.gitignore`.
- Your own `ESELON_MAP` entries in `otomasi_cn_tax.py` matching your client/entity batch-code prefixes (the values here are illustrative placeholders).

## Run

```bash
python otomasi_cn_tax.py
```

## Stack

Python · pandas · openpyxl · pywin32 (Excel COM automation) · Selenium · Gmail API · Google Sheets API
