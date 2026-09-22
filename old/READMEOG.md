# CNOM Report Automation — Setup Guide

This gives you two programs you can just double-click, no Python or technical
setup needed beyond the one-time steps below.

- **fill_report.exe** — pops up a small window to enter today's date and your
  name, and saves a filled-in copy of the shift report.
- **screenshots.exe** — logs into CNOM, captures 11 status screenshots, and
  emails them (plus today's report, if you already ran fill_report.exe) to the
  team.

## One-time setup (do this once per PC / per Windows account)

### 1. Copy these files into one folder

Put all of these together in the same folder (e.g. `C:\CNOM\`):

- `screenshots.exe`
- `fill_report.exe`
- `phonechecks` folder (contains the master template `.docx` file)
- `ms-playwright` folder (the browser CNOM automation uses — see below)

### 2. Set the environment variables

Press the Windows key, type **"environment variables"**, open **"Edit the
system environment variables"** → **Environment Variables**. Under **System
variables** (not User variables — this makes it work for every account on
this PC), click **New** and add each of these:

| Variable name | Value |
|---|---|
| `CNOM_USERNAME` | your CNOM login username |
| `CNOM_PASSWORD` | your CNOM login password |
| `BREVO_API_KEY` | (ask whoever set up the Brevo account) |
| `EMAIL_FROM` | the verified sender address (ask if unsure) |
| `EMAIL_TO` | the recipient list, comma-separated, e.g. `person1@company.com,person2@company.com` |
| `PLAYWRIGHT_BROWSERS_PATH` | full path to the `ms-playwright` folder you copied in step 1, e.g. `C:\CNOM\ms-playwright` |

Click OK on everything to save. **Restart the PC** (or at least fully log out
and back in) so the new variables take effect everywhere.

### 3. Confirm CNOM's VPN/network access works

`screenshots.exe` needs this PC to already be able to reach CNOM
(`https://10.21.32.4:8585`) the normal way you'd access it in a browser. If
you can't open that address in Chrome/Edge on this PC, the exe won't work
either — that's a network/VPN issue to sort out first, not something the exe
can fix.

## Daily use

1. Double-click **fill_report.exe**. A small window appears with today's date
   pre-filled — type your name and click **Save**.
2. Double-click **screenshots.exe**. A window opens and runs through the
   capture automatically (takes a few minutes) — you'll see it working, no
   need to click anything. When it finishes, an email goes out automatically
   with all the screenshots plus today's report.

That's it — no typing commands, no Python.

## If something goes wrong

- **A window flashes and closes immediately**: run it from a Command Prompt
  instead of double-clicking, so the error message stays visible: open
  Command Prompt, `cd` into the folder, then type `screenshots.exe` and press
  Enter.
- **"Executable doesn't exist" mentioning chromium**: the
  `PLAYWRIGHT_BROWSERS_PATH` variable is missing or points to the wrong
  folder — recheck step 2.
- **No email arrives but the screenshots ran fine**: check that `EMAIL_TO`,
  `EMAIL_FROM`, and `BREVO_API_KEY` are all set correctly, and check your spam
  folder.
- **fill_report.exe can't find the template**: make sure the `phonechecks`
  folder (with the master `.docx` template inside it) is in the same folder
  as `fill_report.exe`.

For anything else, contact whoever set this up originally.