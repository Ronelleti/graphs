# Metabase Screenshot Automation

Logs into Metabase, captures a dashboard and a question's chart, and saves
them as PNG files.

## What it produces

Saved into `metabase_screenshots/`:

- `dashboard_1288.png` — the dashboard at `/dashboard/1288`
- `question_17017_notebook.png` — the question at `/question/17017/notebook`,
  after clicking "Visualize" to show the resulting chart rather than the
  query builder view

## Requirements

- **Python 3**, installed with "Add python.exe to PATH" checked during setup
- **Network access** to `http://10.11.2.32:3000` (Metabase) already working
  on the machine
- A Metabase username and password

## Installation

```powershell
pip install playwright
playwright install chromium
```

(If you've already installed these for the CNOM screenshot automation on this
same PC, you don't need to install them again — they're shared.)

## Configuration

Set these two environment variables before running (PowerShell session):

```powershell
$env:METABASE_USERNAME = "your_username"
$env:METABASE_PASSWORD = "your_password"
```

## Usage

```powershell
python metabase_screenshots.py
```

No separate login step needed — on first run (or whenever the saved session
has expired), it logs in automatically using the two environment variables
above and saves a session file (`metabase_session.json`) next to the script,
reused on later runs to skip login when possible.

The script currently runs **headed** (a visible browser window opens) so you
can watch it work — this is useful while confirming everything's correct.
Once confirmed, this should be switched to headless for scheduled/unattended
runs (change `headless=False` to `headless=True` in the script).

## How it works, briefly

1. Logs into Metabase (or reuses a saved session).
2. Navigates to the dashboard URL and screenshots it directly.
3. Navigates to the question's `/notebook` URL (the query-builder view),
   clicks the **Visualize** button to switch to the actual chart, waits for
   it to render, then screenshots that.

## Known considerations

- **Not yet combined with the CNOM email pipeline.** These two screenshots
  currently just save to a local folder — they aren't attached to the same
  email as the CNOM report yet. That's a natural next step once this is
  confirmed working reliably (see the "Known considerations" section of the
  CNOM README for how that pipeline's email step is structured, if you want
  to merge them).
- **Login selectors were adjusted once already** (via `get_by_placeholder`
  and `get_by_role`, matching Metabase's actual login form) — if Metabase is
  ever upgraded and its login page changes, `do_login()` is the function most
  likely to need a re-check.
- **The "Visualize" click is specific to `/notebook` URLs.** If you add more
  question URLs later, decide per-URL whether they need this same click
  (query-builder view) or are already showing the chart directly.