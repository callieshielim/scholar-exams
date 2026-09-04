#!/usr/bin/env python3
"""
Delete (archive) questions from admin.11plusonline.co.uk by their display ID.

The site has no bulk delete: for each question you open the row's 3-dot menu,
click "Remove", then confirm "Archive" in the dialog. This script automates that
for a range of IDs (e.g. Q0775 .. Q0792).

Usage:
    python delete_questions.py                 # deletes Q0775..Q0792 (asks to confirm)
    python delete_questions.py --start 775 --end 792
    python delete_questions.py --dry-run       # show what would be deleted, do nothing
    python delete_questions.py --yes           # skip the confirmation prompt
"""

import argparse
import sys
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── Config ────────────────────────────────────────────────────────────────────
LOGIN_URL = "https://admin.11plusonline.co.uk/login"
LIST_URL  = "https://admin.11plusonline.co.uk/question-bank/question"
EMAIL     = "admin1@11plusonline.co.uk"
PASSWORD  = "password1"

# Range of question IDs to delete (inclusive). Displayed as Q#### (zero-padded).
DEFAULT_START = 775
DEFAULT_END   = 792

DELAY_MS = 800   # pause between deletions


def label(n: int) -> str:
    return f"Q{n:04d}"


# ── Browser helpers ───────────────────────────────────────────────────────────

def login(page):
    page.goto(LOGIN_URL)
    page.wait_for_load_state("networkidle")
    page.fill('input[type="email"], input[name="email"]', EMAIL)
    page.fill('input[type="password"], input[name="password"]', PASSWORD)
    page.click('button[type="submit"]')
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    if "login" in page.url:
        raise RuntimeError("Login failed — check credentials")
    print("[ok] Logged in\n")


def find_row(page, qid: str):
    """Filter the list by ID and return the <tr> whose first cell == qid (or None)."""
    search = page.query_selector('input[placeholder="Search by ID..."]')
    if search is None:
        raise RuntimeError("Could not find the 'Search by ID...' input on the list page")
    search.click()
    search.fill("")            # clear previous filter
    search.fill(qid)
    page.wait_for_timeout(1800)  # let the list re-filter

    for tr in page.query_selector_all("table tbody tr"):
        first_cell = tr.query_selector("td")
        if first_cell and (first_cell.inner_text() or "").strip() == qid:
            return tr
    return None


def delete_one(page, qid: str, dry_run: bool) -> str:
    """Returns 'deleted', 'missing', or 'error:<msg>'."""
    try:
        row = find_row(page, qid)
    except Exception as e:
        return f"error:{e}"

    if row is None:
        return "missing"

    if dry_run:
        return "deleted"  # would-delete

    # open the row's 3-dot menu
    trigger = row.query_selector('button[data-slot="dropdown-menu-trigger"]')
    if trigger is None:
        return "error:no 3-dot menu button in row"
    trigger.click()

    # menu renders in a portal at body level → click the "Remove" item
    try:
        page.wait_for_selector('[role="menuitem"]', timeout=5000)
    except PWTimeout:
        return "error:menu did not open"
    remove = None
    for item in page.query_selector_all('[role="menuitem"]'):
        if (item.inner_text() or "").strip().lower() == "remove":
            remove = item
            break
    if remove is None:
        return "error:'Remove' item not found"
    remove.click()

    # confirm dialog → click the red "Archive" button
    try:
        dialog = page.wait_for_selector('[role="alertdialog"]', timeout=5000)
    except PWTimeout:
        return "error:confirm dialog did not appear"
    archive = None
    for b in dialog.query_selector_all("button"):
        if (b.inner_text() or "").strip().lower() == "archive":
            archive = b
            break
    if archive is None:
        return "error:'Archive' button not found"
    archive.click()

    # wait for the dialog to close (deletion committed)
    try:
        page.wait_for_selector('[role="alertdialog"]', state="detached", timeout=8000)
    except PWTimeout:
        return "error:dialog did not close after Archive"

    page.wait_for_timeout(DELAY_MS)
    return "deleted"


# ── Main ──────────────────────────────────────────────────────────────────────

def main(args):
    ids = [label(n) for n in range(args.start, args.end + 1)]

    print("=" * 56)
    print("DELETE (ARCHIVE) QUESTIONS")
    print("=" * 56)
    print(f"  Range   : {ids[0]} .. {ids[-1]}  ({len(ids)} questions)")
    print(f"  Mode    : {'DRY RUN (nothing will be deleted)' if args.dry_run else 'LIVE — will archive'}")
    print()

    if not args.dry_run and not args.yes:
        ans = input(f"Type 'yes' to archive these {len(ids)} questions: ").strip().lower()
        if ans != "yes":
            print("Aborted.")
            return

    results = {"deleted": [], "missing": [], "error": []}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        login(page)

        page.goto(LIST_URL)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)

        for qid in ids:
            print(f"-> {qid} ...", end=" ", flush=True)
            status = delete_one(page, qid, args.dry_run)
            if status == "deleted":
                print("would delete" if args.dry_run else "archived")
                results["deleted"].append(qid)
            elif status == "missing":
                print("not found (skipped)")
                results["missing"].append(qid)
            else:
                msg = status.split("error:", 1)[-1]
                print(f"ERROR ({msg})")
                results["error"].append(qid)

        browser.close()

    print("\n" + "=" * 56)
    verb = "Would delete" if args.dry_run else "Archived"
    print(f"{verb}: {len(results['deleted'])}/{len(ids)}")
    if results["missing"]:
        print(f"Not found: {', '.join(results['missing'])}")
    if results["error"]:
        print(f"Errors:    {', '.join(results['error'])}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", type=int, default=DEFAULT_START, help="first numeric ID (default 775)")
    ap.add_argument("--end",   type=int, default=DEFAULT_END,   help="last numeric ID (default 792)")
    ap.add_argument("--dry-run", action="store_true", help="show what would be deleted, do nothing")
    ap.add_argument("--yes", action="store_true", help="skip the confirmation prompt")
    args = ap.parse_args()
    if args.start > args.end:
        print("ERROR: --start must be <= --end")
        sys.exit(1)
    main(args)
