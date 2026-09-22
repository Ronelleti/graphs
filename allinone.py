"""
All-in-one shift report automation.

  1. Pops up a small window asking for the check date and examiner name,
     and saves a filled-in copy of the shift report.
  2. Logs into CNOM and captures 11 status screenshots.
  3. Logs into Metabase and captures a dashboard + a question's chart.
  4. Emails all of it together (14 attachments: 11 CNOM PNGs + 2 Metabase
     PNGs + the shift report .docx) via Brevo.

Set credentials once per terminal (PowerShell), or as permanent SYSTEM
environment variables for unattended/shared-PC use:
    $env:CNOM_USERNAME = "your_cnom_username"
    $env:CNOM_PASSWORD = "your_cnom_password"
    $env:METABASE_USERNAME = "your_metabase_username"
    $env:METABASE_PASSWORD = "your_metabase_password"
    $env:BREVO_API_KEY = "your_brevo_api_key"
    $env:EMAIL_FROM = "your_verified_sender@address.com"
    $env:EMAIL_TO = "recipient1@company.com,recipient2@company.com"

Requires: pip install playwright pillow requests python-docx
          playwright install chromium

Run: python all_in_one.py
"""

import os
import re
import sys
import time
import base64
import argparse
import subprocess
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from pathlib import Path


def ensure_package(pip_name: str, import_name: str = None):
    """Check that a package is importable; if not, pip-install it. Only
    useful when running as a raw Python script — once packaged as an exe,
    all packages are already bundled and this becomes a no-op."""
    import_name = import_name or pip_name
    try:
        __import__(import_name)
    except ImportError:
        print(f"Installing missing package: {pip_name}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])


for _pip_name, _import_name in [
    ("playwright", "playwright"),
    ("pillow", "PIL"),
    ("requests", "requests"),
    ("python-docx", "docx"),
]:
    ensure_package(_pip_name, _import_name)


def ensure_chromium():
    """Check that Playwright's Chromium browser is actually installed;
    if not, download it. Only works if playwright's pip package is
    present with a working sys.executable — same caveat as above."""
    from playwright.sync_api import sync_playwright
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
    except Exception as e:
        if "executable doesn't exist" in str(e).lower() or "playwright install" in str(e).lower():
            print("Chromium browser not found — installing (this can take a minute)...")
            subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
        else:
            raise


ensure_chromium()

import docx
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from PIL import Image

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

CNOM_BASE_URL = "https://10.21.32.4:8585/"
CNOM_SESSION_FILE = "session.json"

METABASE_BASE_URL = "http://10.11.2.32:3000"
METABASE_SESSION_FILE = "metabase_session.json"
METABASE_URLS = {
    "dashboard_1288.png": f"{METABASE_BASE_URL}/dashboard/1288",
    "question_17017_notebook.png": f"{METABASE_BASE_URL}/question/17017/notebook",
}

OUTPUT_ROOT = Path("screenshots")
PHONECHECKS_DIR = Path("phonechecks")
REPORT_TEMPLATE_PATH = PHONECHECKS_DIR / "19_09_26.docx"  # master template

CNOM_USERNAME = os.environ.get("CNOM_USERNAME", "")
CNOM_PASSWORD = os.environ.get("CNOM_PASSWORD", "")
METABASE_USERNAME = os.environ.get("METABASE_USERNAME", "")
METABASE_PASSWORD = os.environ.get("METABASE_PASSWORD", "")

BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "")
EMAIL_FROM = os.environ.get("EMAIL_FROM", "")
EMAIL_TO = os.environ.get("EMAIL_TO", "")

CNOM_LOGGED_IN_MARKER = "text=Select item"
NODE_MONITOR_LINK = "text=Node Monitor"
NODE_MONITOR_READY_MARKER = "text=Node list"
GRAPH_COMPARISON_TAB = "text=Graph comparison"


# ---------------------------------------------------------------------------
# Step 1: shift report popup (from fill_report.py)
# ---------------------------------------------------------------------------

def run_report_popup() -> Path:
    """Blocks until the user fills in date/name and clicks Save (or closes
    the window). Returns the saved file path, or None if cancelled."""
    result = {"path": None}

    def on_submit():
        date_str = date_entry.get().strip()
        name_str = name_entry.get().strip()

        if not date_str or not name_str:
            messagebox.showwarning("Missing info", "Please fill in both the date and the name.")
            return

        if not REPORT_TEMPLATE_PATH.exists():
            messagebox.showerror("Template not found", f"Could not find {REPORT_TEMPLATE_PATH.resolve()}")
            return

        try:
            doc = docx.Document(REPORT_TEMPLATE_PATH)
            table = doc.tables[0]
            table.rows[0].cells[1].text = date_str  # תאריך בדיקה
            table.rows[0].cells[3].text = name_str  # שם הבודק
            safe_date = date_str.replace(".", "_").replace("/", "_")
            out_path = PHONECHECKS_DIR / f"{safe_date}.docx"
            doc.save(out_path)
            result["path"] = out_path
            root.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    from tkinter import ttk

    root = tk.Tk()
    root.title("Shift Report")
    root.configure(bg="#f4f6f9")
    root.resizable(False, False)

    # Center the window on screen
    win_w, win_h = 380, 300
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = (screen_w // 2) - (win_w // 2)
    y = (screen_h // 2) - (win_h // 2)
    root.geometry(f"{win_w}x{win_h}+{x}+{y}")

    FONT_TITLE = ("Segoe UI", 15, "bold")
    FONT_LABEL = ("Segoe UI", 10)
    FONT_ENTRY = ("Segoe UI", 11)
    ACCENT = "#2f6fed"

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TEntry", font=FONT_ENTRY, padding=8, relief="flat")
    style.configure(
        "Accent.TButton",
        font=("Segoe UI", 10, "bold"),
        foreground="white",
        background=ACCENT,
        padding=(10, 8),
        borderwidth=0,
    )
    style.map("Accent.TButton", background=[("active", "#255ecb")])

    # Header bar
    header = tk.Frame(root, bg=ACCENT, height=56)
    header.pack(fill="x")
    tk.Label(header, text="📋  Shift Report", font=FONT_TITLE, bg=ACCENT, fg="white").pack(pady=12)

    # Body
    body = tk.Frame(root, bg="#f4f6f9", padx=30, pady=25)
    body.pack(fill="both", expand=True)

    tk.Label(body, text="Check date (DD.MM.YY)", font=FONT_LABEL, bg="#f4f6f9", fg="#333").pack(anchor="w")
    date_entry = ttk.Entry(body, font=FONT_ENTRY, justify="center")
    date_entry.insert(0, datetime.now().strftime("%d.%m.%y"))
    date_entry.pack(fill="x", pady=(4, 18))

    tk.Label(body, text="Examiner name", font=FONT_LABEL, bg="#f4f6f9", fg="#333").pack(anchor="w")
    name_entry = ttk.Entry(body, font=FONT_ENTRY, justify="center")
    name_entry.pack(fill="x", pady=(4, 24))
    name_entry.focus()

    save_btn = tk.Button(
        body,
        text="Start Shift  ➜",
        font=("Segoe UI", 12, "bold"),
        fg="white",
        bg=ACCENT,
        activebackground="#255ecb",
        activeforeground="white",
        relief="flat",
        bd=0,
        cursor="hand2",
        pady=10,
        command=on_submit,
    )
    save_btn.pack(fill="x")
    save_btn.bind("<Enter>", lambda e: save_btn.config(bg="#255ecb"))
    save_btn.bind("<Leave>", lambda e: save_btn.config(bg=ACCENT))

    root.bind("<Return>", lambda event: on_submit())

    root.mainloop()
    return result["path"]


# ---------------------------------------------------------------------------
# Step 2: CNOM screenshots
# ---------------------------------------------------------------------------

def cnom_do_login(page):
    if not CNOM_USERNAME or not CNOM_PASSWORD:
        print("ERROR: CNOM_USERNAME / CNOM_PASSWORD environment variables are not set.")
        sys.exit(1)
    print("[CNOM] Filling in login form automatically...")
    page.fill("#username", CNOM_USERNAME)
    page.fill("#password", CNOM_PASSWORD)
    page.click("#button")
    page.wait_for_selector(CNOM_LOGGED_IN_MARKER, timeout=20000)
    print("[CNOM] Login successful.")


def cnom_session_is_valid(page) -> bool:
    page.goto(CNOM_BASE_URL, wait_until="networkidle")
    try:
        page.wait_for_selector(CNOM_LOGGED_IN_MARKER, timeout=8000)
        return True
    except PWTimeout:
        return False


def cnom_screenshot_node_monitor(page, out_dir: Path):
    print("[CNOM] Opening Node Monitor...")
    time.sleep(1.5)
    page.locator(NODE_MONITOR_LINK).evaluate("el => el.click()")
    page.wait_for_selector(NODE_MONITOR_READY_MARKER, timeout=15000)
    time.sleep(1)
    page.screenshot(path=str(out_dir / "node_monitor.png"), full_page=True)
    print(" -> saved node_monitor.png")


def cnom_select_ims_scope(page):
    print("[CNOM] Expanding Network service and selecting IMS...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    page.get_by_text("Network service", exact=True).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    page.get_by_text("IMS", exact=True).last.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.3)
    page.get_by_role("button", name=re.compile(r"^Select")).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(1.5)


def cnom_open_graph_comparison(page):
    print("[CNOM] Returning to Status Overview...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    page.locator("text=Status Overview").first.evaluate("el => el.click()")
    time.sleep(1.5)

    if page.locator(GRAPH_COMPARISON_TAB).count() == 0:
        cnom_select_ims_scope(page)
    else:
        print("[CNOM] IMS scope already active — skipping selection.")

    print("[CNOM] Opening Graph comparison...")
    page.locator(GRAPH_COMPARISON_TAB).first.evaluate("el => el.click()")
    time.sleep(1.5)


def cnom_set_time_range(page, label_text: str):
    print(f"[CNOM] Setting time range to {label_text}...")
    time_btn = page.get_by_text(re.compile(r"^Last"), exact=False).first
    time_btn.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    page.get_by_text(label_text, exact=True).last.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(1)


def cnom_select_dropdown(page, panel_index: int, label_text: str, option_text: str):
    label = page.locator(f"div.selection-label:text-is('{label_text}')").nth(panel_index)
    container = label.locator("xpath=..")
    dropdown = container.locator("eui-base-v0-dropdown")
    dropdown.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    menu = dropdown.locator("eui-base-v0-menu")

    option = menu.get_by_text(option_text, exact=True).first
    option.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)

    if label_text == "KPI":
        for _ in range(3):
            time.sleep(0.5)
            checked_boxes = menu.locator("eui-base-v0-checkbox[checked]")
            names = checked_boxes.evaluate_all("els => els.map(el => el.getAttribute('name'))")
            stray_names = [n for n in names if n and n != option_text]
            if not stray_names:
                break
            for name in stray_names:
                stray = menu.get_by_text(name, exact=True).first
                stray.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
                time.sleep(0.3)
        page.keyboard.press("Escape")
        time.sleep(0.3)


def cnom_capture_single_panel(page, node_type: str, node_name: str, kpi_group: str, kpi: str,
                               time_range: str, tmp_path: Path) -> Path:
    cnom_open_graph_comparison(page)
    if time_range:
        cnom_set_time_range(page, time_range)

    cnom_select_dropdown(page, 0, "Node type", node_type)
    cnom_select_dropdown(page, 0, "Node name", node_name)
    cnom_select_dropdown(page, 0, "KPI group", kpi_group)
    cnom_select_dropdown(page, 0, "KPI", kpi)
    time.sleep(2)

    full_path = tmp_path.with_suffix(".full.png")
    page.screenshot(path=str(full_path), full_page=True)

    header = page.get_by_text("Select data", exact=True).first
    footer = page.get_by_text("Add new KPI selection", exact=False).first
    top = max(0, header.bounding_box()["y"] - 20)
    footer_box = footer.bounding_box()
    bottom = footer_box["y"] + footer_box["height"] + 20

    img = Image.open(full_path)
    cropped = img.crop((0, top, img.width, min(img.height, bottom)))
    cropped.save(tmp_path)
    full_path.unlink(missing_ok=True)
    print(f"[CNOM]    captured panel: {node_type}/{node_name}/{kpi_group}/{kpi}")
    return tmp_path


def cnom_stack_images(image_paths: list, output_path: Path):
    images = [Image.open(p) for p in image_paths]
    width = max(im.width for im in images)
    total_height = sum(im.height for im in images)
    combined = Image.new("RGB", (width, total_height), "white")
    y = 0
    for im in images:
        combined.paste(im, (0, y))
        y += im.height
    combined.save(output_path)
    for p in image_paths:
        Path(p).unlink(missing_ok=True)
    print(f" -> saved {output_path.name}")


def cnom_build_composite(page, out_dir: Path, panels: list, time_range: str, filename: str):
    print(f"[CNOM] Building {filename} from {len(panels)} panel(s)...")
    tmp_paths = []
    for i, (node_type, node_name, kpi_group, kpi) in enumerate(panels):
        tmp_path = out_dir / f"_tmp_panel_{i}.png"
        cnom_capture_single_panel(page, node_type, node_name, kpi_group, kpi, time_range, tmp_path)
        tmp_paths.append(tmp_path)
    cnom_stack_images(tmp_paths, out_dir / filename)


def cnom_is_tree_item_checked(page, item_label: str) -> bool:
    tree = page.locator("eui-base-v0-tree")
    item = tree.locator(f"e-tree-view-item[label='{item_label}']").first
    cb = item.locator("input[type='checkbox']").first
    return cb.count() > 0 and cb.is_checked()


def cnom_select_epg_scope(page):
    print("[CNOM] Selecting EPG (HFEPG01 + RHEPG01)...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    tree = page.locator("eui-base-v0-tree")

    if tree.get_by_text("IMS", exact=True).count() == 0 and tree.get_by_text("CORE", exact=True).count() == 0:
        tree.get_by_text("Network service", exact=True).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
        time.sleep(0.5)
    if tree.get_by_text("EPG", exact=True).count() == 0:
        tree.get_by_text("Node type", exact=True).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
        time.sleep(0.5)

    for name in ["IMS", "CORE"]:
        if cnom_is_tree_item_checked(page, name):
            print(f"[CNOM]    clearing: {name}")
            tree.get_by_text(name, exact=True).last.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
            time.sleep(0.5)

    if tree.get_by_text("HFEPG01", exact=True).count() == 0:
        tree.get_by_text("EPG", exact=True).first.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
        time.sleep(0.5)
    for node_name in ["HFEPG01", "RHEPG01"]:
        if not cnom_is_tree_item_checked(page, node_name):
            item = tree.get_by_text(node_name, exact=True).first
            item.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
            time.sleep(0.5)

    time.sleep(0.5)
    select_btn = page.get_by_role("button", name=re.compile(r"^Select"))
    if select_btn.count() > 0:
        select_btn.first.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
        time.sleep(1.5)
    else:
        print("[CNOM]    no Select button to confirm — assuming already applied.")


def cnom_open_epg_graph_comparison(page):
    print("[CNOM] Returning to Status Overview...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    page.locator("text=Status Overview").first.evaluate("el => el.click()")
    time.sleep(1.5)

    print("[CNOM] Reopening item selection...")
    page.get_by_text("Select item", exact=True).first.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(1)

    cnom_select_epg_scope(page)

    print("[CNOM] Opening Graph comparison...")
    page.locator(GRAPH_COMPARISON_TAB).first.evaluate("el => el.click()")
    time.sleep(1.5)


def cnom_select_kpitimeline_dropdown(page, panel_index: int, dropdown_position: int, option_text: str):
    panel = page.locator("e-cnom-lib-node-and-kpi-selector").nth(panel_index)
    dropdown = panel.locator("eui-base-v0-dropdown").nth(dropdown_position)
    dropdown.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    menu = dropdown.locator("eui-base-v0-menu")
    option = menu.get_by_text(option_text, exact=True).first
    option.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)


def cnom_capture_epg_panel(page, node_name: str, kpi_group: str, panel_index: int):
    cnom_select_kpitimeline_dropdown(page, panel_index, 0, node_name)
    cnom_select_kpitimeline_dropdown(page, panel_index, 1, kpi_group)
    time.sleep(2)


def cnom_build_epg_composites(page, out_dir: Path):
    cnom_open_epg_graph_comparison(page)
    cnom_set_time_range(page, "Last 7 days")

    cnom_capture_epg_panel(page, "HFEPG01", "SGW Traffic usage", 0)
    cnom_capture_epg_panel(page, "HFEPG01", "SGW Throughput", 1)
    time.sleep(2)
    page.screenshot(path=str(out_dir / "SGW_HFEPG01.png"), full_page=True)
    print(" -> saved SGW_HFEPG01.png")

    cnom_capture_epg_panel(page, "RHEPG01", "SGW Traffic usage", 0)
    cnom_capture_epg_panel(page, "RHEPG01", "SGW Throughput", 1)
    time.sleep(2)
    page.screenshot(path=str(out_dir / "SGW_RHEPG01.png"), full_page=True)
    print(" -> saved SGW_RHEPG01.png")

    cnom_capture_epg_panel(page, "HFEPG01", "PGW Traffic usage", 0)
    cnom_capture_epg_panel(page, "HFEPG01", "PGW Throughput", 1)
    time.sleep(2)
    page.screenshot(path=str(out_dir / "PGW_HFEPG01.png"), full_page=True)
    print(" -> saved PGW_HFEPG01.png")

    cnom_capture_epg_panel(page, "RHEPG01", "PGW Traffic usage", 0)
    cnom_capture_epg_panel(page, "RHEPG01", "PGW Throughput", 1)
    time.sleep(2)
    page.screenshot(path=str(out_dir / "PGW_RHEPG01.png"), full_page=True)
    print(" -> saved PGW_RHEPG01.png")


def run_cnom_captures(browser, out_dir: Path):
    if Path(CNOM_SESSION_FILE).exists():
        context = browser.new_context(ignore_https_errors=True, storage_state=CNOM_SESSION_FILE, viewport={"width": 1920, "height": 1080})
    else:
        context = browser.new_context(ignore_https_errors=True, viewport={"width": 1920, "height": 1080})

    page = context.new_page()

    if not cnom_session_is_valid(page):
        print("[CNOM] Session missing or expired — logging in automatically...")
        cnom_do_login(page)
        context.storage_state(path=CNOM_SESSION_FILE)
        print("[CNOM] Session refreshed.")

    cnom_screenshot_node_monitor(page, out_dir)

    cnom_build_composite(
        page, out_dir,
        [("CSCF", "HFCSCF01", "Registrations", "IMSCSCFRegisteredUsers"),
         ("CSCF", "RHCSCF01", "Registrations", "IMSCSCFRegisteredUsers")],
        time_range="Last 7 days",
        filename="RegisteredUsers.png",
    )
    cnom_build_composite(
        page, out_dir,
        [("SBG", "HFIBCF01", "H.248", "IMSSBGAbnormalBGFTerminations"),
         ("SBG", "HFSBG01", "H.248", "IMSSBGAbnormalBGFTerminations")],
        time_range="Last 7 days",
        filename="AbnormalBGFTerminations_HF.png",
    )
    cnom_build_composite(
        page, out_dir,
        [("SBG", "RHIBCF01", "H.248", "IMSSBGAbnormalBGFTerminations"),
         ("SBG", "RHSBG01", "H.248", "IMSSBGAbnormalBGFTerminations")],
        time_range="Last 7 days",
        filename="AbnormalBGFTerminations_RH.png",
    )
    cnom_build_composite(
        page, out_dir,
        [("SBG", "HFIBCF01", "H.248", "IMSSBGActiveBGFCalls"),
         ("SBG", "HFSBG01", "H.248", "IMSSBGActiveBGFCalls")],
        time_range="Last 7 days",
        filename="ActiveBGFCalls_HF.png",
    )
    cnom_build_composite(
        page, out_dir,
        [("SBG", "RHIBCF01", "H.248", "IMSSBGActiveBGFCalls"),
         ("SBG", "RHSBG01", "H.248", "IMSSBGActiveBGFCalls")],
        time_range="Last 7 days",
        filename="ActiveBGFCalls_RH.png",
    )
    cnom_build_composite(
        page, out_dir,
        [("SBG", "HFSBG01", "Registrations", "IMSASBGAKARegUsers"),
         ("SBG", "RHSBG01", "Registrations", "IMSASBGAKARegUsers")],
        time_range="Last 7 days",
        filename="AKARegUsers.png",
    )

    cnom_build_epg_composites(page, out_dir)
    context.close()


# ---------------------------------------------------------------------------
# Step 3: Metabase screenshots
# ---------------------------------------------------------------------------

def metabase_do_login(page):
    if not METABASE_USERNAME or not METABASE_PASSWORD:
        print("ERROR: METABASE_USERNAME / METABASE_PASSWORD environment variables are not set.")
        sys.exit(1)
    print("[Metabase] Going to login page...")
    page.goto(f"{METABASE_BASE_URL}/auth/login")
    page.wait_for_load_state("networkidle")

    print("[Metabase] Filling in login form...")
    page.get_by_placeholder("nicetoseeyou@email.com").fill(METABASE_USERNAME)
    page.get_by_placeholder("Shhh...").fill(METABASE_PASSWORD)
    page.get_by_role("button", name="Sign in").click()
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    if "/auth/login" in page.url:
        print("[Metabase] WARNING: still on the login page after submitting — login likely failed.")
    else:
        print("[Metabase] Login successful.")


def metabase_session_is_valid(page) -> bool:
    page.goto(METABASE_BASE_URL, wait_until="networkidle")
    return "/auth/login" not in page.url


def metabase_capture(page, out_dir: Path, name: str, url: str, click_visualize: bool = False):
    print(f"[Metabase] Opening {url}")
    page.goto(url, wait_until="networkidle")
    time.sleep(3)

    if click_visualize:
        print("[Metabase] Clicking Visualize...")
        page.get_by_role("button", name="Visualize").click()
        page.wait_for_load_state("networkidle")
        time.sleep(3)

    out_path = out_dir / name
    page.screenshot(path=str(out_path), full_page=True)
    print(f" -> saved {name}")


def run_metabase_captures(browser, out_dir: Path):
    if Path(METABASE_SESSION_FILE).exists():
        context = browser.new_context(storage_state=METABASE_SESSION_FILE, viewport={"width": 1920, "height": 1080})
    else:
        context = browser.new_context(viewport={"width": 1920, "height": 1080})

    page = context.new_page()

    if not metabase_session_is_valid(page):
        print("[Metabase] Session missing or expired — logging in...")
        metabase_do_login(page)
        if "/auth/login" not in page.url:
            context.storage_state(path=METABASE_SESSION_FILE)

    for name, url in METABASE_URLS.items():
        metabase_capture(page, out_dir, name, url, click_visualize=(name == "question_17017_notebook.png"))

    context.close()


# ---------------------------------------------------------------------------
# Step 4: Email
# ---------------------------------------------------------------------------

def send_report_email(out_dir: Path, report_path: Path):
    images = sorted(out_dir.glob("*.png"))
    all_files = images + ([report_path] if report_path and report_path.exists() else [])

    if not all_files:
        print("No files found to email.")
        return
    if not BREVO_API_KEY or not EMAIL_FROM or not EMAIL_TO:
        print("BREVO_API_KEY / EMAIL_FROM / EMAIL_TO not set — skipping email (files are still saved).")
        return

    print(f"Emailing {len(all_files)} file(s) via Brevo to {EMAIL_TO}...")

    to_list = [{"email": addr.strip()} for addr in EMAIL_TO.split(",") if addr.strip()]

    attachments = []
    for path in all_files:
        content = base64.b64encode(path.read_bytes()).decode()
        attachments.append({"content": content, "name": path.name})

    payload = {
        "sender": {"name": "Shift Report Automation", "email": EMAIL_FROM},
        "to": to_list,
        "subject": f"Shift Report — {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "htmlContent": (
            "<p>Automated shift report attached.</p>"
            f"<p>Files: {', '.join(p.name for p in all_files)}</p>"
        ),
        "attachment": attachments,
    }

    try:
        response = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            headers={
                "api-key": BREVO_API_KEY,
                "Content-Type": "application/json",
                "accept": "application/json",
            },
            json=payload,
            timeout=60,
        )
        if response.status_code in (200, 201):
            print("Email sent via Brevo.")
        else:
            print(f"Email failed: {response.status_code} {response.text}")
    except Exception as e:
        print(f"Email failed: {e}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Opening shift report form...")
    report_path = run_report_popup()
    if report_path:
        print(f"Shift report saved: {report_path}")
    else:
        print("Shift report was not saved (window closed/cancelled) — continuing without it.")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = OUTPUT_ROOT / timestamp
    out_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # switch to True for unattended runs

        run_cnom_captures(browser, out_dir)
        run_metabase_captures(browser, out_dir)

        browser.close()

    print(f"\nDone. Files saved in {out_dir}/")

    send_report_email(out_dir, report_path)


if __name__ == "__main__":
    main()