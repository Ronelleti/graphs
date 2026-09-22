"""
Standalone test: Metabase dashboard + question screenshots.

Set credentials first (PowerShell):
    $env:METABASE_USERNAME = "your_username"
    $env:METABASE_PASSWORD = "your_password"

Run: python metabase_screenshots.py
"""

import os
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE_URL = "http://10.11.2.32:3000"
SESSION_FILE = "metabase_session.json"
OUTPUT_DIR = Path("metabase_screenshots")
OUTPUT_DIR.mkdir(exist_ok=True)

USERNAME = os.environ.get("METABASE_USERNAME", "")
PASSWORD = os.environ.get("METABASE_PASSWORD", "")

URLS = {
    "dashboard_1288.png": f"{BASE_URL}/dashboard/1288",
    "question_17017_notebook.png": f"{BASE_URL}/question/17017/notebook",
}


def do_login(page):
    if not USERNAME or not PASSWORD:
        print("ERROR: METABASE_USERNAME / METABASE_PASSWORD environment variables are not set.")
        sys.exit(1)
    print("Going to Metabase login page...")
    page.goto(f"{BASE_URL}/auth/login")
    page.wait_for_load_state("networkidle")

    print("Filling in login form...")
    page.get_by_placeholder("nicetoseeyou@email.com").fill(USERNAME)
    page.get_by_placeholder("Shhh...").fill(PASSWORD)
    page.get_by_role("button", name="Sign in").click()
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    if "/auth/login" in page.url:
        print("WARNING: still on the login page after submitting — login likely failed.")
        page.screenshot(path="debug_metabase_after_login.png", full_page=True)
    else:
        print("Login looks successful.")


def session_is_valid(page) -> bool:
    page.goto(BASE_URL, wait_until="networkidle")
    return "/auth/login" not in page.url


def capture(page, name, url, click_visualize=False):
    print(f"Opening {url}")
    page.goto(url, wait_until="networkidle")
    time.sleep(3)  # let the page render

    if click_visualize:
        print("Clicking Visualize...")
        page.get_by_role("button", name="Visualize").click()
        page.wait_for_load_state("networkidle")
        time.sleep(3)  # let the resulting chart render

    out_path = OUTPUT_DIR / name
    page.screenshot(path=str(out_path), full_page=True)
    print(f" -> saved {out_path}")


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # headed so you can watch it work

        if Path(SESSION_FILE).exists():
            context = browser.new_context(storage_state=SESSION_FILE, viewport={"width": 1920, "height": 1080})
        else:
            context = browser.new_context(viewport={"width": 1920, "height": 1080})

        page = context.new_page()

        if not session_is_valid(page):
            print("Session missing or expired — logging in...")
            do_login(page)
            if "/auth/login" not in page.url:
                context.storage_state(path=SESSION_FILE)

        for name, url in URLS.items():
            capture(page, name, url, click_visualize=(name == "question_17017_notebook.png"))

        input("Press Enter to close the browser...")
        browser.close()


if __name__ == "__main__":
    main()