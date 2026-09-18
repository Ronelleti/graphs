"""
CNOM Status Overview screenshot automation.

  python cnom_screenshots.py --setup      -> headed, manual/auto login, saves session.json
  python cnom_screenshots.py              -> headless scheduled run (what n8n calls)

Set credentials once per terminal (PowerShell):
    $env:CNOM_USERNAME = "your_username"
    $env:CNOM_PASSWORD = "your_password"

Requires: pip install playwright openpyxl pillow
          playwright install chromium
"""

import os
import re
import sys
import time
import argparse
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from PIL import Image

BASE_URL = "https://10.21.32.4:8585/"
SESSION_FILE = "session.json"
OUTPUT_ROOT = Path("screenshots")

USERNAME = os.environ.get("CNOM_USERNAME", "")
PASSWORD = os.environ.get("CNOM_PASSWORD", "")

LOGGED_IN_MARKER = "text=Select item"
NODE_MONITOR_LINK = "text=Node Monitor"
NODE_MONITOR_READY_MARKER = "text=Node list"
GRAPH_COMPARISON_TAB = "text=Graph comparison"


def debug_shot(page, name):
    page.screenshot(path=f"debug_{name}.png", full_page=True)


def do_login(page):
    if not USERNAME or not PASSWORD:
        print("ERROR: CNOM_USERNAME / CNOM_PASSWORD environment variables are not set.")
        sys.exit(1)
    print("Filling in login form automatically...")
    page.fill("#username", USERNAME)
    page.fill("#password", PASSWORD)
    page.click("#button")
    page.wait_for_selector(LOGGED_IN_MARKER, timeout=20000)
    print("Login successful.")


def run_setup():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(ignore_https_errors=True, viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        page.goto(BASE_URL)
        do_login(page)
        context.storage_state(path=SESSION_FILE)
        print(f"Session saved to {SESSION_FILE}.")
        browser.close()


def session_is_valid(page) -> bool:
    page.goto(BASE_URL, wait_until="networkidle")
    try:
        page.wait_for_selector(LOGGED_IN_MARKER, timeout=8000)
        return True
    except PWTimeout:
        return False


def screenshot_node_monitor(page, out_dir: Path):
    print("Opening Node Monitor...")
    time.sleep(1.5)
    page.locator(NODE_MONITOR_LINK).evaluate("el => el.click()")
    page.wait_for_selector(NODE_MONITOR_READY_MARKER, timeout=15000)
    time.sleep(1)
    page.screenshot(path=str(out_dir / "node_monitor.png"), full_page=True)
    print(" -> saved node_monitor.png")


def select_ims_scope(page):
    print("Expanding Network service and selecting IMS...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    page.get_by_text("Network service", exact=True).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    page.get_by_text("IMS", exact=True).last.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.3)
    page.get_by_role("button", name=re.compile(r"^Select")).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(1.5)


def open_graph_comparison(page):
    """Navigate fresh to Graph comparison with a single default panel."""
    print("Returning to Status Overview...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    page.locator("text=Status Overview").first.evaluate("el => el.click()")
    time.sleep(1.5)

    if page.locator(GRAPH_COMPARISON_TAB).count() == 0:
        select_ims_scope(page)
    else:
        print("IMS scope already active — skipping selection.")

    print("Opening Graph comparison...")
    page.locator(GRAPH_COMPARISON_TAB).first.evaluate("el => el.click()")
    time.sleep(1.5)


def set_time_range(page, label_text: str):
    print(f"Setting time range to {label_text}...")
    time_btn = page.get_by_text(re.compile(r"^Last"), exact=False).first
    time_btn.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    page.get_by_text(label_text, exact=True).last.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(1)


def select_dropdown(page, panel_index: int, label_text: str, option_text: str):
    """Select an option in the panel_index-th occurrence of a dropdown.
    We only ever use panel_index=0 now, since each panel is captured fresh."""
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
        # A default KPI can end up checked with a delay, arriving AFTER we
        # click our target. So check a few times, after a short wait each
        # time, and uncheck anything that isn't our target.
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


def capture_single_panel(page, node_type: str, node_name: str, kpi_group: str, kpi: str,
                          time_range: str, tmp_path: Path) -> Path:
    """Fresh navigation, select one node's panel, crop out just that
    Select-data-plus-graph block, save it, and return the cropped path."""
    open_graph_comparison(page)
    if time_range:
        set_time_range(page, time_range)

    select_dropdown(page, 0, "Node type", node_type)
    select_dropdown(page, 0, "Node name", node_name)
    select_dropdown(page, 0, "KPI group", kpi_group)
    select_dropdown(page, 0, "KPI", kpi)
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
    print(f"   captured panel: {node_type}/{node_name}/{kpi_group}/{kpi}")
    return tmp_path


def stack_images(image_paths: list, output_path: Path):
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


def build_composite(page, out_dir: Path, panels: list, time_range: str, filename: str):
    """panels: list of (node_type, node_name, kpi_group, kpi) tuples.
    Each is captured in its own fresh navigation, then stacked vertically."""
    print(f"Building {filename} from {len(panels)} panel(s)...")
    tmp_paths = []
    for i, (node_type, node_name, kpi_group, kpi) in enumerate(panels):
        tmp_path = out_dir / f"_tmp_panel_{i}.png"
        capture_single_panel(page, node_type, node_name, kpi_group, kpi, time_range, tmp_path)
        tmp_paths.append(tmp_path)
    stack_images(tmp_paths, out_dir / filename)


def run_scheduled():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = OUTPUT_ROOT / timestamp
    out_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        if Path(SESSION_FILE).exists():
            context = browser.new_context(ignore_https_errors=True, storage_state=SESSION_FILE, viewport={"width": 1920, "height": 1080})
        else:
            context = browser.new_context(ignore_https_errors=True, viewport={"width": 1920, "height": 1080})

        page = context.new_page()

        if not session_is_valid(page):
            print("Session missing or expired — logging in automatically...")
            do_login(page)
            context.storage_state(path=SESSION_FILE)
            print("Session refreshed.")

        screenshot_node_monitor(page, out_dir)

        build_composite(
            page, out_dir,
            [("CSCF", "HFCSCF01", "Registrations", "IMSCSCFRegisteredUsers"),
             ("CSCF", "RHCSCF01", "Registrations", "IMSCSCFRegisteredUsers")],
            time_range="Last 7 days",
            filename="RegisteredUsers.png",
        )

        build_composite(
            page, out_dir,
            [("SBG", "HFIBCF01", "H.248", "IMSSBGAbnormalBGFTerminations"),
             ("SBG", "HFSBG01", "H.248", "IMSSBGAbnormalBGFTerminations")],
            time_range="Last 7 days",
            filename="AbnormalBGFTerminations_HF.png",
        )

        build_composite(
            page, out_dir,
            [("SBG", "RHIBCF01", "H.248", "IMSSBGAbnormalBGFTerminations"),
             ("SBG", "RHSBG01", "H.248", "IMSSBGAbnormalBGFTerminations")],
            time_range="Last 7 days",
            filename="AbnormalBGFTerminations_RH.png",
        )

        build_composite(
            page, out_dir,
            [("SBG", "HFIBCF01", "H.248", "IMSSBGActiveBGFCalls"),
             ("SBG", "HFSBG01", "H.248", "IMSSBGActiveBGFCalls")],
            time_range="Last 7 days",
            filename="ActiveBGFCalls_HF.png",
        )

        build_composite(
            page, out_dir,
            [("SBG", "RHIBCF01", "H.248", "IMSSBGActiveBGFCalls"),
             ("SBG", "RHSBG01", "H.248", "IMSSBGActiveBGFCalls")],
            time_range="Last 7 days",
            filename="ActiveBGFCalls_RH.png",
        )

        build_composite(
            page, out_dir,
            [("SBG", "HFSBG01", "Registrations", "IMSASBGAKARegUsers"),
             ("SBG", "RHSBG01", "Registrations", "IMSASBGAKARegUsers")],
            time_range="Last 7 days",
            filename="AKARegUsers.png",
        )

        browser.close()

    print(f"\nDone. Screenshots saved in {out_dir}/")
    print(str(out_dir))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true")
    args = parser.parse_args()

    if args.setup:
        run_setup()
    else:
        run_scheduled()