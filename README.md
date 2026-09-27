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

## Configuration (once per PC, by the tool owner only)

All settings live in **`config.ini` next to `allinone.exe`**. No environment
variables are needed, so users without permission to edit them can still run
the tool. Everyone shares the same CNOM and Metabase account, one Brevo
sender and one receiver, so nobody using the tool has to enter anything.

1. Copy `config.example.ini` to `config.ini` in the same folder as the exe.
2. Fill in the values (no quotes; `$`, `%` and `"` in passwords are fine):

```ini
[cnom]
username = shared_cnom_username
password = shared_cnom_password

[metabase]
username = shared_metabase_username
password = shared_metabase_password

[email]
brevo_api_key = your_brevo_api_key
from = verified_sender@company.com
to = receiver@company.com
```

### Graphs inside the email body (optional)

Brevo's HTTP API can only send the graphs as attachments. To show them in
the email body instead, add Brevo's SMTP login to `[email]`:

```ini
smtp_login = xxxxxx@smtp-brevo.com
smtp_key = your_smtp_key
```

Both values are on Brevo's **SMTP & API → SMTP** tab (the SMTP key is
different from the API key). The email then goes through
`smtp-relay.brevo.com` on port 587. If that port is blocked or the login
fails, the program falls back to the API and sends the graphs as
attachments, as before. Set `smtp_port = 2525` if only 587 is blocked.

If a value is left empty, the matching environment variable (`CNOM_USERNAME`,
`EMAIL_TO`, ...) is used instead, so older setups keep working.

Every email goes from the one verified sender to the one receiver. The
subject shows the check date and the examiner's name, and the body shows
which PC and Windows user ran it, so the reports can be told apart.

`config.ini` is in `.gitignore`. Never commit it.

> Anyone who can run the tool can open `config.ini`, so treat the Brevo key
> as known to all users. In Brevo, restrict the key to your office's public
> IP (Security → Authorized IPs) and use a key made only for this tool.

## Folder layout

```
C:\dailychecks\             <- set up by the tool owner, same on both PCs
  allinone.exe
  config.ini                 <- your real settings (not in git)
  ms-playwright\             <- Chromium, copied here (exe only)
  phonechecks\
    19_09_26.docx            <- master template (read only, never overwritten)
```

Each Windows user gets their own folder, created automatically:

```
%LOCALAPPDATA%\dailychecks\
  session.json               <- CNOM login session
  metabase_session.json      <- Metabase login session
  screenshots\<date_time>\    <- captured images
  reports\<date>.docx         <- filled-in shift report
```

So several people can use the same PC under different Windows accounts
without sharing sessions or running into "access denied" on each other's
files. The program finds `config.ini` and the template from the exe's own
folder, so it works from a shortcut or Task Scheduler no matter what the
"Start in" folder is.

## Usage

```powershell
python allinone.py
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
pyinstaller --onefile allinone.py
```

(No `--windowed` flag this time — we want the console window to stay visible
so you can see the automation's progress, in addition to the popup.)

The finished `allinone.exe` appears in the `dist\` folder. Move it out to
your main project folder (next to `phonechecks\`, `session.json`, etc.) —
same as with the earlier exe files.

## Testing it on another PC

This is the real test of whether the exe is truly portable. On the second
PC:

### 1. Copy the whole folder over
Copy the folder containing `allinone.exe`, `config.ini` and `phonechecks\`
(with the template `.docx` inside). The session files are per-user and get
created on first run, so there's nothing else to copy.

### 2. Copy the Chromium browser folder next to the exe
The exe does **not** bundle Playwright's browser binary. Copy your
`ms-playwright` folder from:
```
C:\Users\<your_username>\AppData\Local\ms-playwright
```
into the tool folder, so it sits at `C:\dailychecks\ms-playwright\`. The
program finds it there by itself for every Windows user. No environment
variable is needed.

### 3. Create config.ini
Copy `config.ini` from the first PC (or fill in `config.example.ini`), as
described under **Configuration** above.

### 4. Confirm network access works on that PC
Both CNOM's and Metabase's addresses need to actually be reachable from that
second PC — open them in a regular browser there first to confirm, before
testing the exe.

### 5. Run it
```
.\allinone.exe
```
The popup should appear, then the browser automation should run through both
CNOM and Metabase, then send the email — all without Python or pip installed
on that second machine at all.

If something fails on the second PC that worked fine on the first, the two
most likely causes (based on everything we've hit so far) are: the Chromium
folder missing from `C:\dailychecks\ms-playwright\`, or that PC's network/VPN not actually reaching
CNOM/Metabase the way the first one does.