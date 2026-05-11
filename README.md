# amzn-feedback-manager

Scrapes Amazon Seller Central feedback for each seller account, classifies entries as positive or negative, stores them in a SQL Server database, and emails a formatted Excel report. Runs on a weekday schedule via APScheduler.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install git+https://github.com/dominicci13/shared-python-utils.git
```

### 2. Configure environment

```bash
cp .env.example .env
```

Edit `.env` with your credentials and SQL table names.

### 3. Configure accounts

Account names and Seller Central URLs are loaded from `fc_utils.accounts` via `config/accounts.json` in the shared-python-utils installation.

## Run

```bash
python run_amzn_feedback_manager.py
```

The script runs automatically at 10:00 Mon–Fri via APScheduler.

## Environment Variables

| Variable | Description |
|---|---|
| `AMZN_email` | Amazon Seller Central login email |
| `AMZN_pass` | Amazon Seller Central password |
| `CHROME_USER_DATA_DIR` | Path to Chrome automation profile directory |
| `SENDER_EMAIL` | Outlook account used to send the report email |
| `TO_EMAIL` | Comma-separated list of recipient email addresses |
| `CC_EMAIL` | Comma-separated list of CC email addresses |
| `DB_TABLE_NEGATIVE` | SQL Server table name for negative feedback (default: `FedManNegative`) |
| `DB_TABLE_POSITIVE` | SQL Server table name for positive feedback (default: `FedManPositive`) |
