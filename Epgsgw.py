"""
Standalone test: EPG (HFEPG01/RHEPG01) SGW Traffic usage + SGW Throughput.

Reuses the same login/session approach as the main cnom_screenshots.py.
Run: python epg_sgw.py
"""

import os
import re
import sys
import time
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


def session_is_valid(page) -> bool:
    page.goto(BASE_URL, wait_until="networkidle")
    try:
        page.wait_for_selector(LOGGED_IN_MARKER, timeout=8000)
        return True
    except PWTimeout:
        return False


def tree_checkbox(page, item_label: str):
    """The real native checkbox input for a tree item, found via its
    e-tree-view-item[label] attribute — a genuine <input type="checkbox">
    two shadow-roots deep, confirmed via devtools."""
    tree = page.locator("eui-base-v0-tree")
    item = tree.locator(f"e-tree-view-item[label='{item_label}']").first
    return item.locator("input[type='checkbox']").first


def set_checkbox(checkbox_locator, checked: bool):
    """The input is visually hidden, so set the property directly and fire
    the events the component listens for, instead of a real click."""
    checkbox_locator.evaluate(
        """(el, checked) => {
            el.checked = checked;
            el.dispatchEvent(new Event('click', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('input', { bubbles: true }));
        }""",
        checked,
    )


def clear_any_scope(page):
    """Uncheck whatever is currently selected in the tree, whatever it is,
    so we start from a clean slate before selecting EPG."""
    tree = page.locator("eui-base-v0-tree")
    checked = tree.locator("input[type='checkbox']:checked")
    count = checked.count()
    for i in range(count):
        cb = checked.nth(0)  # re-query each time since unchecking shifts the list
        set_checkbox(cb, False)
        time.sleep(0.2)


def select_epg_scope(page):
    print("Selecting EPG (HFEPG01 + RHEPG01)...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    tree = page.locator("eui-base-v0-tree")

    # Read the breadcrumb to know what's currently selected, so we clear
    # exactly that instead of guessing/checking an unreliable property.
    breadcrumb = page.locator("text=Select item").first.locator("xpath=preceding-sibling::*[1]")
    current_text = breadcrumb.inner_text() if breadcrumb.count() > 0 else ""
    print(f"   currently selected (per breadcrumb): '{current_text.strip()}'")
    current_names = [n.strip() for n in current_text.split(",") if n.strip()]

    # Expand both sections so whatever's currently checked is clickable.
    tree.get_by_text("Network service", exact=True).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    tree.get_by_text("Node type", exact=True).evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)

    for name in current_names:
        item = tree.get_by_text(name, exact=True).last
        if item.count() > 0:
            print(f"   clearing: {name}")
            item.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
            time.sleep(0.3)

    # Expand EPG and check both its children by clicking their visible text
    # (the mechanism that's reliably worked for checking items all along).
    tree.get_by_text("EPG", exact=True).first.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    for node_name in ["HFEPG01", "RHEPG01"]:
        item = tree.get_by_text(node_name, exact=True).first
        item.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
        time.sleep(0.3)

    time.sleep(0.5)
    new_breadcrumb = page.locator("text=Select item").first.locator("xpath=preceding-sibling::*[1]")
    print(f"   breadcrumb now shows: '{new_breadcrumb.inner_text().strip() if new_breadcrumb.count() > 0 else '?'}'")

    select_btn = page.get_by_role("button", name=re.compile(r"^Select"))
    if select_btn.count() > 0:
        select_btn.first.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
        time.sleep(1.5)
    else:
        print("   no Select button to confirm — assuming already applied.")


def open_epg_graph_comparison(page):
    print("Returning to Status Overview...")
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.3)
    page.locator("text=Status Overview").first.evaluate("el => el.click()")
    time.sleep(1.5)

    print("Reopening item selection...")
    page.get_by_text("Select item", exact=True).first.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(1)

    select_epg_scope(page)

    print("Opening Graph comparison...")
    page.screenshot(path="debug_epg_before_graph_tab.png", full_page=True)
    page.locator(GRAPH_COMPARISON_TAB).first.evaluate("el => el.click()")
    time.sleep(1.5)


def set_time_range(page, label_text: str):
    print(f"Setting time range to {label_text}...")
    time_btn = page.get_by_text(re.compile(r"^Last"), exact=False).first
    time_btn.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    page.get_by_text(label_text, exact=True).last.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(1)


def select_dropdown(page, panel_index: int, label_text: str, option_text):
    """option_text: a specific string, or "__ALL__" to check every item."""
    label = page.locator(f"div.selection-label:text-is('{label_text}')").nth(panel_index)
    container = label.locator("xpath=..")
    dropdown = container.locator("eui-base-v0-dropdown")
    dropdown.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    menu = dropdown.locator("eui-base-v0-menu")

    if label_text == "KPI" and option_text == "__ALL__":
        time.sleep(0.3)
        item_labels = menu.locator("eui-base-v0-menu-item").evaluate_all(
            "els => els.map(el => el.getAttribute('label'))"
        )
        for lbl in item_labels:
            if lbl:
                item = menu.get_by_text(lbl, exact=True).first
                item.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
                time.sleep(0.2)
        page.keyboard.press("Escape")
        time.sleep(0.3)
        return

    option = menu.get_by_text(option_text, exact=True).first
    option.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)


def select_kpitimeline_dropdown(page, panel_index: int, dropdown_position: int, option_text: str):
    """For the 'KPI timeline' widget's Node name / KPI group dropdowns.
    dropdown_position: 0 for Node name, 1 for KPI group — found by position
    within the panel's own e-cnom-lib-node-and-kpi-selector wrapper, which
    stays stable even after a value is selected (unlike the dropdown's own
    'label' attribute, which changes to show the picked value)."""
    panel = page.locator("e-cnom-lib-node-and-kpi-selector").nth(panel_index)
    dropdown = panel.locator("eui-base-v0-dropdown").nth(dropdown_position)
    dropdown.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)
    menu = dropdown.locator("eui-base-v0-menu")
    option = menu.get_by_text(option_text, exact=True).first
    option.evaluate("el => { el.scrollIntoView({block: 'center'}); el.click(); }")
    time.sleep(0.5)


def capture_epg_panel(page, node_name: str, kpi_group: str, panel_index: int):
    select_kpitimeline_dropdown(page, panel_index, 0, node_name)
    select_kpitimeline_dropdown(page, panel_index, 1, kpi_group)
    time.sleep(2)
    debug_shot(page, f"epg_panel_{panel_index}_after_kpi_group")


def add_new_kpi_selection(page):
    print(" Adding another panel...")
    page.get_by_text("Add new KPI selection", exact=False).last.evaluate("el => el.click()")
    time.sleep(1)


def main():
    out_dir = OUTPUT_ROOT / datetime.now().strftime("%Y%m%d_%H%M_epgtest")
    out_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # headed for now, so you can watch it work

        if Path(SESSION_FILE).exists():
            context = browser.new_context(ignore_https_errors=True, storage_state=SESSION_FILE, viewport={"width": 1920, "height": 1080})
        else:
            context = browser.new_context(ignore_https_errors=True, viewport={"width": 1920, "height": 1080})

        page = context.new_page()

        if not session_is_valid(page):
            print("Session missing or expired — logging in automatically...")
            do_login(page)
            context.storage_state(path=SESSION_FILE)

        try:
            open_epg_graph_comparison(page)
            set_time_range(page, "Last 7 days")

            # Two panels already exist by default (one per scoped node) —
            # set both to HFEPG01 with the two different KPI groups.
            capture_epg_panel(page, "HFEPG01", "SGW Traffic usage", 0)
            capture_epg_panel(page, "HFEPG01", "SGW Throughput", 1)

            time.sleep(2)
            full_path = out_dir / "SGW_HFEPG01.png"
            page.screenshot(path=str(full_path), full_page=True)
            print(f" -> saved {full_path}")

            # Now switch both panels' Node name to RHEPG01, same KPI groups.
            capture_epg_panel(page, "RHEPG01", "SGW Traffic usage", 0)
            capture_epg_panel(page, "RHEPG01", "SGW Throughput", 1)

            time.sleep(2)
            full_path2 = out_dir / "SGW_RHEPG01.png"
            page.screenshot(path=str(full_path2), full_page=True)
            print(f" -> saved {full_path2}")

            # Same pattern again, but with the PGW KPI groups.
            capture_epg_panel(page, "HFEPG01", "PGW Traffic usage", 0)
            capture_epg_panel(page, "HFEPG01", "PGW Throughput", 1)

            time.sleep(2)
            full_path3 = out_dir / "PGW_HFEPG01.png"
            page.screenshot(path=str(full_path3), full_page=True)
            print(f" -> saved {full_path3}")

            capture_epg_panel(page, "RHEPG01", "PGW Traffic usage", 0)
            capture_epg_panel(page, "RHEPG01", "PGW Throughput", 1)

            time.sleep(2)
            full_path4 = out_dir / "PGW_RHEPG01.png"
            page.screenshot(path=str(full_path4), full_page=True)
            print(f" -> saved {full_path4}")
        except Exception as e:
            print(f"FAILED: {e}")
            print("Browser stays open — go look at it now, then come back here.")

        input("Press Enter to close the browser...")
        browser.close()


if __name__ == "__main__":
    main()