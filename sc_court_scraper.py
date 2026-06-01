"""
SC Public Index Court Records Scraper
======================================
Scrapes all 46 South Carolina counties from publicindex.sccourts.org
for App/Trans/Rts cases with Case Filed dates in the last X days.

Requirements:
    python -m pip install playwright python-dateutil
    python -m playwright install chromium

Usage:
    python sc_court_scraper.py                       <- headless
    python sc_court_scraper.py --visible             <- shows browser
    python sc_court_scraper.py --visible --test aiken  <- one county, visible
"""

import pandas as pd
import sys
import time
from datetime import date, timedelta
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("ERROR: playwright not installed.")
    print("Run:  python -m pip install playwright && python -m playwright install chromium")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
BASE_URL   = "https://publicindex.sccourts.org/{county}/publicindex/"
TODAY      = date.today()
BEGIN_DATE = TODAY - timedelta(days=30) #CHANGE DATE RANGE right herere
DATE_FMT   = "%m/%d/%Y"
OUTPUT_XLSX = "sc_court_records.xlsx"

#Greenville link is different and tried to get past authentication, but cannot, so 45/46 counties covered
COUNTIES = [
    "abbeville", "aiken", "allendale", "anderson", "bamberg",
    "barnwell", "beaufort", "berkeley", "calhoun", "charleston",
    "cherokee", "chester", "chesterfield", "clarendon", "colleton",
    "darlington", "dillon", "dorchester", "edgefield", "fairfield",
    "florence", "georgetown", "greenwood", "hampton",
    "horry", "jasper", "kershaw", "lancaster", "laurens",
    "lee", "lexington", "marion", "marlboro", "mccormick",
    "newberry", "oconee", "orangeburg", "pickens", "richland",
    "saluda", "spartanburg", "sumter", "union", "williamsburg",
    "york",
]

# ---------------------------------------------------------------------------
# Parse command-line args
# ---------------------------------------------------------------------------
HEADLESS    = False
TEST_COUNTY = None

args = sys.argv[1:]
if "--visible" in args:
    HEADLESS = False
    args.remove("--visible")
if "--test" in args:
    idx = args.index("--test")
    TEST_COUNTY = args[idx + 1].lower()
    COUNTIES = [TEST_COUNTY]
    print(f"TEST MODE: running only '{TEST_COUNTY}'")

# ---------------------------------------------------------------------------
# Helper: select by index
# ---------------------------------------------------------------------------
def set_select(page, select_index: int, desired_text_choices: list):
    selects = page.locator("select")
    count = selects.count()
    if select_index >= count:
        raise RuntimeError(f"Expected at least {select_index+1} <select> elements, found {count}")

    sel = selects.nth(select_index)
    sel.scroll_into_view_if_needed()
    sel.wait_for(state="visible", timeout=20_000)

    options = sel.evaluate("el => Array.from(el.options).map(o => o.text.trim())")
    print(f"      select[{select_index}] options: {options}")

    chosen = None
    for desired in desired_text_choices:
        match = next((o for o in options if o.strip() == desired), None)
        if not match:
            match = next((o for o in options if desired.lower() in o.lower()), None)
        if match:
            chosen = match
            break

    if chosen is None:
        raise ValueError(f"None of {desired_text_choices} found. Available: {options}")

    print(f"      → Selecting '{chosen}'")

    # Use JS dispatchEvent to avoid triggering a page navigation via the
    # ASP.NET __doPostBack that causes the 406 error
    sel.evaluate(f"""
        el => {{
            for (let i = 0; i < el.options.length; i++) {{
                if (el.options[i].text.trim() === {repr(chosen)}) {{
                    el.selectedIndex = i;
                    break;
                }}
            }}
            el.dispatchEvent(new Event('change', {{ bubbles: true }}));
        }}
    """)

    # Give the page a moment to react without doing a full navigation
    time.sleep(2)


# ---------------------------------------------------------------------------
# Helper: fill input by index
# ---------------------------------------------------------------------------
def set_input(page, input_index: int, value: str):
    inputs = page.locator(
        "input[type='text'], input[type='date'], input:not([type])"
    )
    count = inputs.count()
    if input_index >= count:
        raise RuntimeError(f"Expected at least {input_index+1} text inputs, found {count}")

    field = inputs.nth(input_index)
    field.scroll_into_view_if_needed()
    field.wait_for(state="visible", timeout=20_000)
    field.click()
    field.fill(value)
    field.dispatch_event("change")
    print(f"      input[{input_index}] → '{value}'")


# ---------------------------------------------------------------------------
# Debug dump
# ---------------------------------------------------------------------------
def debug_dump(page):
    print("\n  ── DEBUG: <select> elements ──")
    selects = page.locator("select")
    for i in range(selects.count()):
        s = selects.nth(i)
        try:
            opts = s.evaluate("el => Array.from(el.options).map(o => o.text.trim())")
            cur  = s.evaluate("el => el.options[el.selectedIndex]?.text.trim()")
            print(f"    select[{i}] current='{cur}' options={opts[:5]}")
        except Exception:
            print(f"    select[{i}]: (could not read)")

    print("  ── DEBUG: text <input> elements ──")
    inputs = page.locator("input[type='text'], input:not([type])")
    for i in range(min(inputs.count(), 12)):
        inp = inputs.nth(i)
        try:
            name = inp.get_attribute("name") or inp.get_attribute("id") or "?"
            val  = inp.input_value()
            print(f"    input[{i}] name='{name}' value='{val}'")
        except Exception:
            print(f"    input[{i}]: (could not read)")


# ---------------------------------------------------------------------------
# Extract result rows
# ---------------------------------------------------------------------------
def extract_table_rows(page) -> list:
    rows = []
    try:
        page.wait_for_selector("table", timeout=30_000)
    except PWTimeout:
        print("      No table appeared after search.")
        return rows

    tables = page.locator("table").all()
    target = None
    best_col_count = 0
    for tbl in tables:
        try:
            first_row_cells = tbl.locator("tr").first.locator("th, td").all_text_contents()
            if len(first_row_cells) > best_col_count:
                best_col_count = len(first_row_cells)
                target = tbl
        except Exception:
            continue

    if target is None or best_col_count < 2:
        print("      Could not identify a results table.")
        return rows

    all_trs = target.locator("tr").all()
    if not all_trs:
        return rows

    header_cells = all_trs[0].locator("th, td").all_text_contents()
    headers = [h.strip() for h in header_cells]
    print(f"      Headers: {headers}")

    def col_idx(keywords):
        for kw in keywords:
            for i, h in enumerate(headers):
                if kw.lower() in h.lower():
                    return i
        return None

    idx_name  = col_idx(["name"])
    idx_party = col_idx(["party type", "party"])
    idx_case  = col_idx(["case number", "case no", "case #", "case"])
    idx_filed = col_idx(["filed date", "date filed", "filed"])

    if all(x is None for x in [idx_name, idx_party, idx_case, idx_filed]):
        print("      WARNING: headers unrecognized, using columns 0-3")
        idx_name, idx_party, idx_case, idx_filed = 0, 1, 2, 3

    for tr in all_trs[1:]:
        cells = tr.locator("td").all_text_contents()
        if not cells:
            continue

        def safe(idx):
            if idx is None or idx >= len(cells):
                return ""
            return cells[idx].strip()

        row = {
            "Name"        : safe(idx_name),
            "Case Number" : safe(idx_case),
            "Filed Date"  : safe(idx_filed),
        }
        if any(row.values()):
            rows.append(row)

    return rows


def handle_pagination(page) -> list:
    all_rows = []
    page_num = 1
    while True:
        print(f"      → reading page {page_num} …")
        batch = extract_table_rows(page)
        all_rows.extend(batch)

        clicked = False
        next_candidates = page.locator(
            "a:has-text('Next'), a[title='Next'], a:has-text('>'), a:has-text('»')"
        ).all()
        for link in next_candidates:
            try:
                txt = link.text_content().strip()
                if link.is_visible() and txt in ("Next", ">", "»"):
                    link.click()
                    page.wait_for_load_state("networkidle", timeout=40_000)
                    page_num += 1
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            break

    return all_rows


# ---------------------------------------------------------------------------
# Wait for search outcome: results table OR no-results popup
# Returns: "results" | "no_results" | "unknown"
# ---------------------------------------------------------------------------
def wait_for_search_outcome(page, timeout_ms: int = 60_000) -> str:
    """
    After clicking Search, wait for one of two things to appear:
      1. A results table (county has matching cases)
      2. A popup / alert-style element saying no results were found

    If neither appears within timeout_ms, returns "unknown" so the
    caller knows to retry rather than silently move on.
    """
    deadline = time.time() + timeout_ms / 1000
    poll_interval = 0.5

    while time.time() < deadline:
        # ── Check for the specific no-results popup ────────────────────────
        # The site shows a popup with the exact text:
        #   "Your search did not return any results"
        # and an "Ok" button beneath it. Both must be visible.
        try:
            msg = page.locator("text=Your search did not return any results")
            ok_btn = page.locator("button:has-text('OK'), input[value='OK']")
            if msg.first.is_visible() and ok_btn.first.is_visible():
                print("      No-results popup confirmed ('Your search did not return any results').")
                # Dismiss the popup so the page is clean for any retry
                ok_btn.first.click()
                return "no_results"
        except Exception:
            pass

        # ── Check for a results table with actual data rows ────────────────
        try:
            tables = page.locator("table").all()
            for tbl in tables:
                rows = tbl.locator("tr").all()
                if len(rows) > 1:          # header + at least one data row
                    cells = rows[0].locator("th, td").all_text_contents()
                    headers_lower = [c.strip().lower() for c in cells]
                    if any("name" in h for h in headers_lower):
                        print("      Results table detected.")
                        return "results"
        except Exception:
            pass

        time.sleep(poll_interval)

    return "unknown"


# ---------------------------------------------------------------------------
# Per-county scrape
# ---------------------------------------------------------------------------
def scrape_county(browser, county: str) -> list:
    context = browser.new_context(
        # Spoof a real Chrome on Windows — avoids Accept-header-based 406
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        extra_http_headers={
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;"
                "q=0.9,image/avif,image/webp,image/apng,*/*;"
                "q=0.8,application/signed-exchange;v=b3;q=0.7"
            ),
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
        },
        # Remove the 'webdriver' flag that sites use to detect automation
        ignore_https_errors=True,
    )

    # Patch navigator.webdriver to undefined so the site can't detect automation
    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        });
    """)

    page = context.new_page()

    try:
        # ── 1. Load disclaimer ─────────────────────────────────────────────
        if(county == "Greenville"):
            url = "https://www2.greenvillecounty.org/scjd/publicindex/"
        else:
            url = BASE_URL.format(county=county)
        print(f"  [1/4] {url}")
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)

        # ── 2. Accept ──────────────────────────────────────────────────────
        print("  [2/4] Accepting disclaimer …")
        accept = page.locator(
            "input[value='Accept'], input[value='I Accept'], "
            "button:has-text('Accept'), a:has-text('Accept')"
        ).first
        accept.wait_for(state="visible", timeout=30_000)
        accept.click()

        # Wait for the SC Judicial 'verifying your browser' screen to pass.
        # Poll for the first <select> to appear using Playwright locators
        # (avoids eval-based wait_for_function which CSP blocks).
        print("      Waiting for search form to load (browser verification may appear) …")
        page.locator("select").first.wait_for(state="visible", timeout=60_000)
        page.wait_for_load_state("networkidle", timeout=40_000)

        # ── 3. Fill form ───────────────────────────────────────────────────
        # Dropdowns in order:
        #   [0] Court Type      leave alone
        #   [1] Agency          leave alone
        #   [2] Case Type       → App/Trans/Rts
        #   [3] Case SubType    leave alone
        #   [4] Party Type      leave alone
        #   [5] Action Type     leave alone
        #   [6] Date Type       → Case Filed
        #
        # Text inputs in order:
        #   [0] Case #          leave alone
        #   [1] Last Name       leave alone
        #   [2] First           leave alone
        #   [3] Middle          leave alone
        #   [4] Suffix          leave alone
        #   [5] CDR Code        leave alone
        #   [6] Indictment #    leave alone
        #   [7] Beginning       → 60 days ago
        #   [8] Ending          → today

        print("  [3/4] Setting filters …")

        # Case Type (index 2) — selecting this may trigger an ASP.NET
        # postback that briefly wipes the DOM, so we wait for the page
        # to fully settle before touching anything else.
        set_select(page, 2, ["App/Trans/Rts", "App/Trans/Rt", "App/Trns/Rts", "App/Trns/Rt"])
        
        

        # Wait for selects to reappear after the ASP.NET postback
        # Use locator-based wait — CSP blocks eval-based wait_for_function
        page.locator("select").nth(6).wait_for(state="visible", timeout=30_000)
        page.wait_for_load_state("networkidle", timeout=60_000)

        # Date Type (index 6)
        set_select(page, 6, ["Case Filed"])
        page.wait_for_load_state("networkidle", timeout=20_000)

        # Fill date fields by their confirmed ASP.NET name attributes
        for name_attr, value in [
            ("TextBoxDateFrom", BEGIN_DATE.strftime(DATE_FMT)),
            ("TextBoxDateTo",   TODAY.strftime(DATE_FMT)),
        ]:
            field = page.locator(f"input[name*='{name_attr}']").first
            field.wait_for(state="visible", timeout=20_000)
            field.click()
            field.fill(value)
            field.dispatch_event("change")
            print(f"      {name_attr} → '{value}'")

        # ── 4. Search (with outcome verification & retry) ────────────────
        MAX_SEARCH_RETRIES = 3
        rows = None

        for attempt in range(1, MAX_SEARCH_RETRIES + 1):
            print(f"  [4/4] Searching (attempt {attempt}/{MAX_SEARCH_RETRIES}) …")

            # Dismiss any JS alert that may appear (no-results popup)
            dismissed_alert = [False]
            def _on_dialog(dialog):
                msg = dialog.message.lower()
                print(f"      JS alert: '{dialog.message}'")
                dismissed_alert[0] = True
                dialog.accept()
            page.on("dialog", _on_dialog)

            search_btn = page.locator(
                "input[value='Search'], button:has-text('Search')"
            ).first
            search_btn.wait_for(state="visible", timeout=20_000)
            search_btn.click()
            page.wait_for_load_state("networkidle", timeout=60_000)

            # If a JS alert fired, that IS the no-results confirmation
            if dismissed_alert[0]:
                print("      No-results confirmed via JS alert.")
                rows = []
                break

            outcome = wait_for_search_outcome(page, timeout_ms=30_000)

            if outcome == "results":
                rows = handle_pagination(page)
                break
            elif outcome == "no_results":
                print("      No-results popup confirmed — no cases for this county.")
                rows = []
                break
            else:  # unknown — neither table nor popup appeared
                print(f"      WARNING: Could not confirm search outcome (attempt {attempt}).")
                if attempt < MAX_SEARCH_RETRIES:
                    print("      Retrying search …")
                    time.sleep(3)
                    # Re-click search; form values are still set
                    continue
                else:
                    print("      All retries exhausted — treating as no records.")
                    rows = []

        page.remove_listener("dialog", _on_dialog)
        return rows if rows is not None else []

    except Exception as exc:
        print(f"  !! ERROR for {county}: {exc}")
        debug_dump(page)
        if not HEADLESS:
            print("     (pausing 15s — check the browser window)")
            time.sleep(15)
        return []

    finally:
        context.close()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print(" SC Public Index — App/Trans/Rts Scraper")
    print(f" Date range : {BEGIN_DATE.strftime(DATE_FMT)}  →  {TODAY.strftime(DATE_FMT)}")
    print(f" Counties   : {len(COUNTIES)}")
    print(f" Output     : {OUTPUT_XLSX}")
    print(f" Headless   : {HEADLESS}")
    print("=" * 60)

    all_records = []
    county_summary = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=HEADLESS,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
        )

        for i, county in enumerate(COUNTIES, 1):
            print(f"\n[{i:02d}/{len(COUNTIES)}] {county.upper()}")
            rows = scrape_county(browser, county)

            if rows:
                for r in rows:
                    r["County"] = county.title()
                all_records.extend(rows)
                print(f"  ✓  {len(rows)} record(s) found")
            else:
                print("  –  No records")

            county_summary.append((county.title(), len(rows)))
            time.sleep(10)

        browser.close()
        # Dedupe within the current run by Case Number
        seen = set()
        deduped_current = []
        for r in all_records:
            key = r["Case Number"].strip()
            if key not in seen:
                seen.add(key)
                deduped_current.append(r)

        intra_run_dupes = len(all_records) - len(deduped_current)
        if intra_run_dupes:
            print(f"🔁  Removed {intra_run_dupes} intra-run duplicate(s)")

        all_records = deduped_current
    if all_records:
        output_path = Path(OUTPUT_XLSX)

        # Load existing workbook if present
        existing_records = []
        existing_case_numbers = set()

        if output_path.exists():
            existing_df = pd.read_excel(output_path, dtype=str)
            existing_df = existing_df.fillna("")
            existing_records = existing_df.to_dict("records")

            if "Case Number" in existing_df.columns:
                existing_case_numbers = set(
                    existing_df["Case Number"].astype(str).str.strip()
                )

            print(f"\n📂  Loaded {len(existing_records)} existing record(s) from {output_path.name}")

        # Filter out duplicates already present in workbook
        new_records = [
            r for r in all_records
            if r["Case Number"].strip() not in existing_case_numbers
        ]

        duplicate_count = len(all_records) - len(new_records)

        if duplicate_count:
            print(f"🔁  Skipped {duplicate_count} duplicate(s) already in Excel")

        combined = existing_records + new_records

        if combined:
            df = pd.DataFrame(combined)

            desired_cols = ["County", "Name", "Case Number", "Filed Date"]
            df = df[[c for c in desired_cols if c in df.columns]]

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Records")

                ws = writer.sheets["Records"]

                # Freeze header row
                ws.freeze_panes = "A2"

                # Auto-size columns
                for column in ws.columns:
                    max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
                    ws.column_dimensions[column[0].column_letter].width = min(max_len + 2, 60)

            print(f"✅  Saved {len(combined)} total record(s) ({len(new_records)} new) → {output_path.resolve()}")
        else:
            print("\n⚠️  No records found across any county and no existing file. No Excel file written.")
    else:
        print("\n⚠️  No records found across any county. No Excel file written.")

    print("\n── County Summary ──────────────────────────")
    for name, count in county_summary:
        marker = "✓" if count else "–"
        print(f"  {marker}  {name:<18} {count:>5} record(s)")
    print("────────────────────────────────────────────")


if __name__ == "__main__":
    main()