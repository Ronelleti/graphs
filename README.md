# All-In-One Shift Report Automation

One program that does everything: asks for the shift date/name, captures 11
CNOM screenshots, captures 2 Metabase screenshots, and emails all 14 files
together in one message.

## What it does, in order

1. A small popup window appears asking for the check date (pre-filled with
   today) and your name. Click **Save** to continue, or close the window to
   skip this part (the run continues, just without the report attached).
2. Logs into CNOM and captures the same 11 files as before: Node Monitor,
   RegisteredUsers, AbnormalBGFTerminations (HF/RH), ActiveBGFCalls (HF/RH),
   AKARegUsers, and the 4 SGW/PGW EPG graphs.
3. Logs into Metabase and captures the dashboard and the question's chart.
4. Sends one email via Brevo with everything attached (up to 14 files).

## Requirements

- **Python 3**, with "Add python.exe to PATH" checked during install
- Network access to both CNOM (`https://10.21.32.4:8585`) and Metabase
  (`http://10.11.2.32:3000`), already working on the machine
- CNOM login, Metabase login, and a Brevo account (verified sender + API key)

## Installation

```powershell
pip install playwright pillow requests python-docx
playwright install chromium
```

## Configuration

Set these 7 environment variables (PowerShell session, or as permanent
SYSTEM variables — see the exe section below):

```powershell
$env:CNOM_USERNAME = "your_cnom_username"
$env:CNOM_PASSWORD = "your_cnom_password"
$env:METABASE_USERNAME = "your_metabase_username"
$env:METABASE_PASSWORD = "your_metabase_password"
$env:BREVO_API_KEY = "your_brevo_api_key"
$env:EMAIL_FROM = "your_verified_sender@address.com"
$env:EMAIL_TO = "recipient1@company.com,recipient2@company.com"
```

## Folder layout required

```
your-folder/
  all_in_one.py
  phonechecks/
    19_09_26.docx      <- master template (name/date get overwritten each run)
  session.json          <- created automatically after first CNOM login
  metabase_session.json <- created automatically after first Metabase login
```

## Usage

```powershell
python all_in_one.py
```

The popup appears first — fill it in, then the browser automation runs on
its own (a few minutes) and sends the email at the end.

Currently runs **headed** (a visible browser window) so you can watch it
work. For unattended runs later, change `headless=False` to `headless=True`
near the bottom of the script.

## Known considerations

- **Two separate logins, two separate sessions**: CNOM and Metabase are
  unrelated systems with their own credentials and their own saved session
  files (`session.json` and `metabase_session.json`). Either can expire
  independently — the script re-logs into whichever one needs it.
- **The popup can't be skipped for unattended/scheduled runs** without further
  changes: since it needs a person to type their name, this script (as-is)
  needs someone to actually be there to click through the popup — it's not
  yet suitable for a fully automatic overnight schedule. If that's needed
  later, the popup step would need to be made optional or replaced.
- **This is a first merge** of three previously-separate, individually-tested
  scripts — treat the first few runs as a new testing round even though each
  piece worked on its own before.

---

## Building it as a standalone .exe

```powershell
pip install pyinstaller
pyinstaller --onefile all_in_one.py
```

(No `--windowed` flag this time — we want the console window to stay visible
so you can see the automation's progress, in addition to the popup.)

The finished `all_in_one.exe` appears in the `dist\` folder. Move it out to
your main project folder (next to `phonechecks\`, `session.json`, etc.) —
same as with the earlier exe files.

## Testing it on another PC

This is the real test of whether the exe is truly portable. On the second
PC:

### 1. Copy the whole folder over
Copy the entire folder containing `all_in_one.exe`, `phonechecks\` (with the
template `.docx` inside), and (if they already exist) `session.json` /
`metabase_session.json` — though those two will just get recreated on first
run if missing, since login is automatic.

### 2. Copy the Chromium browser folder
The exe does **not** bundle Playwright's browser binary. Copy your
`ms-playwright` folder from:
```
C:\Users\<your_username>\AppData\Local\ms-playwright
```
to the second PC — anywhere is fine, just note the path.

### 3. Set the 8 environment variables as SYSTEM variables
Same 7 as before, **plus** `PLAYWRIGHT_BROWSERS_PATH` pointing at wherever
you put the copied `ms-playwright` folder on that PC. Set these as **System**
variables (not User) so they work regardless of which Windows account is
logged in — see `SETUP_FOR_TEAM.md` from the CNOM project for the exact
Windows menu steps (same process, just 8 variables instead of 6).

### 4. Confirm network access works on that PC
Both CNOM's and Metabase's addresses need to actually be reachable from that
second PC — open them in a regular browser there first to confirm, before
testing the exe.

### 5. Run it
```
.\all_in_one.exe
```
The popup should appear, then the browser automation should run through both
CNOM and Metabase, then send the email — all without Python or pip installed
on that second machine at all.

If something fails on the second PC that worked fine on the first, the two
most likely causes (based on everything we've hit so far) are: the Chromium
folder path being wrong, or that PC's network/VPN not actually reaching
CNOM/Metabase the way the first one does.