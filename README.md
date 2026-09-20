# CNOM Status Report Automation

Automates logging into CNOM (Core Network Operations Manager), capturing a set of
Status Overview screenshots and graphs, and saving them to a timestamped folder —
ready to be emailed or picked up by a scheduler.

## What it produces

Each run creates a new folder under `screenshots/<timestamp>/` containing 11 files:

- `node_monitor.png` — full Node Monitor status table
- `RegisteredUsers.png` — CSCF Registrations (HFCSCF01 + RHCSCF01)
- `AbnormalBGFTerminations_HF.png` / `AbnormalBGFTerminations_RH.png` — SBG H.248 abnormal terminations
- `ActiveBGFCalls_HF.png` / `ActiveBGFCalls_RH.png` — SBG H.248 active calls
- `AKARegUsers.png` — SBG Registrations (AKA registered users)
- `SGW_HFEPG01.png` / `SGW_RHEPG01.png` — EPG SGW Traffic usage + Throughput
- `PGW_HFEPG01.png` / `PGW_RHEPG01.png` — EPG PGW Traffic usage + Throughput

## Requirements

- **Python 3**, installed with "Add python.exe to PATH" checked during setup
- **Network/VPN access** to `https://10.21.32.4:8585` (CNOM) already working on the machine — the script cannot set this up itself
- A CNOM username and password
- A [Brevo](https://brevo.com) account (free tier: 300 emails/day) with a verified sender address and an API key

## Installation

```powershell
pip install playwright pillow requests
playwright install chromium
```

## Configuration

Set these five environment variables before running (PowerShell session):

```powershell
$env:CNOM_USERNAME = "your_cnom_username"
$env:CNOM_PASSWORD = "your_cnom_password"
$env:BREVO_API_KEY = "your_brevo_api_key"
$env:EMAIL_FROM = "your_verified_sender@address.com"
$env:EMAIL_TO = "recipient1@company.com,recipient2@company.com"
```

`EMAIL_FROM` must exactly match an address you've verified as a sender in Brevo
(Settings → Senders, Domains & Dedicated IPs → Senders) — Brevo rejects sending
from an unverified address. `EMAIL_TO` accepts multiple recipients, comma-separated.

None of these are saved anywhere by the script — they need to be set each time
the terminal/session starts, or configured as permanent system environment
variables if this will run unattended (e.g. via Task Scheduler).

## Usage

```powershell
python cnom_screenshots.py
```

That's it — no separate setup/login step is needed. On first run (or whenever the
saved session has expired), the script logs in automatically using the two
environment variables above and saves a session file (`session.json`) next to the
script, which it reuses on subsequent runs to skip login when possible.

The script runs **headless** by default (no visible browser window) — this is the
mode intended for scheduled/unattended runs. A full run currently takes a few
minutes, since each of the 11 outputs is captured via its own fresh page
navigation for reliability.

## How it works, briefly

1. Logs into CNOM (or reuses a saved session).
2. Screenshots the Node Monitor page.
3. For the IMS-scoped composites (CSCF/SBG graphs): navigates to Status Overview,
   selects the IMS scope in the Dashboard Tree View, opens Graph Comparison, and
   fills in Node type / Node name / KPI group / KPI for each panel.
4. For the EPG-scoped composites (SGW/PGW graphs): switches the Dashboard Tree
   View scope to Node type → EPG (checking HFEPG01 and RHEPG01), which opens a
   different "KPI timeline" view with its own Node name / KPI group fields.
5. Each multi-node graph is built by capturing one node's panel at a time and
   stitching the images together vertically (via Pillow), rather than relying on
   the page's own multi-panel UI, which proved unreliable for this purpose.

## Known considerations

- **Email is session-independent**: since it's sent via Brevo's API (a plain
  HTTP request), it works the same whether anyone is logged into the PC or
  not — unlike a desktop mail client. This makes the script suitable for
  running as a scheduled task or Windows Service without anyone needing to be
  logged in.
- **CNOM's front-end UI is fragile to automate**: several custom UI components
  (checkboxes, dropdowns) required specific workarounds discovered via trial
  and error (see comments throughout the script, particularly around
  `select_dropdown`, `select_epg_scope`, and `select_kpitimeline_dropdown`). If
  CNOM's UI changes in a future update, these are the functions most likely to
  need adjustment.
- **Brevo free tier**: 300 emails/day, more than enough for 3x/day runs. The
  `EMAIL_FROM` address must stay verified in Brevo — if it's ever removed or
  the verification lapses, sending will fail silently in the log (the script
  prints the error rather than crashing).

## Scheduling

To run this automatically at set times (e.g. 07:00 / 15:00 / 23:00), use Windows
Task Scheduler pointed at:

```
python C:\path\to\cnom_screenshots.py
```

with the `CNOM_USERNAME`, `CNOM_PASSWORD`, `BREVO_API_KEY`, `EMAIL_FROM`, and
`EMAIL_TO` environment variables set as system environment variables (not just
PowerShell session variables) so they're available regardless of how the task
is triggered.