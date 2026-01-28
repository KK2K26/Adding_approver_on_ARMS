import time
import os
import json
from datetime import datetime
import pandas as pd  # Excel I/O
from urllib.parse import urljoin

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    NoSuchWindowException,
    WebDriverException,
)

REQUESTS_URL   = "https://bat.bats.kyndryl.net/arms2/unit-owner/packages"
ARMS_HOST      = "bat.bats.kyndryl.net"

BROWSER        = input("Enter the browser in which have ARMS open (chrome/edge): ").strip().lower()

EXCEL_PATH     = "account.xlsx"
SHEET_NAME     = "Sheet1"
REMOTE_DEBUG   = "localhost:9222"

PER_ITEM_DELAY = 0.3
MATCH_MODE     = "equals"  # equals|startswith|plain

OU_ID_COLUMN = "id"
ACCOUNT_NAME_COLUMN = "Account name"

PROGRESS_FILE  = "progress.json"
RESUME_MODE    = True
STOP_ON_ERROR  = True

APPROVERS_INPUT = input("Enter 3 approvers (comma-separated): ").strip()
APPROVER_LIST = [a.strip() for a in APPROVERS_INPUT.split(",") if a.strip()]
if len(APPROVER_LIST) != 3:
    raise ValueError(f"Please enter exactly 3 approvers separated by commas. You entered {len(APPROVER_LIST)}.")

AUTOMATION_HANDLE = None


def _row_key(ou_id, account_name):
    """Return a normalized unique key for a row."""
    return f"{str(ou_id).strip().lower()}||{str(account_name).strip().lower()}"


def load_progress():
    """Load progress from file or return default structure."""
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "completed_keys" not in data:
                data["completed_keys"] = []
            if "in_progress" not in data:
                data["in_progress"] = {}
            return data
        except Exception:
            return {"completed_keys": [], "in_progress": {}}
    return {"completed_keys": [], "in_progress": {}}


def save_progress(progress):
    """Write progress safely to disk (atomic replace)."""
    tmp = PROGRESS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(progress, f, indent=2)
    os.replace(tmp, PROGRESS_FILE)


def update_in_progress(progress, key, excel_row, link_index, approver_index):
    """Update 'in_progress' checkpoint for resume."""
    progress.setdefault("in_progress", {})
    progress["in_progress"][key] = {
        "excel_row": int(excel_row),
        "link_index": int(link_index),
        "approver_index": int(approver_index),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    save_progress(progress)


def mark_row_completed(progress, key):
    """Mark a key as completed and persist."""
    if key in progress.get("in_progress", {}):
        del progress["in_progress"][key]

    completed = set(progress.get("completed_keys", []))
    completed.add(key)
    progress["completed_keys"] = sorted(completed)
    progress["completed_at"] = datetime.now().isoformat(timespec="seconds")
    save_progress(progress)


def get_edge_driver_attached(debugger_addr=REMOTE_DEBUG):
    """Attach Selenium to an existing Edge session via remote debugger."""
    opts = EdgeOptions()
    opts.add_experimental_option("debuggerAddress", debugger_addr)
    return webdriver.Edge(options=opts)


def get_chrome_driver_attached(debugger_addr=REMOTE_DEBUG):
    """Attach Selenium to an existing Chrome session via remote debugger."""
    opts = ChromeOptions()
    opts.add_experimental_option("debuggerAddress", debugger_addr)
    return webdriver.Chrome(options=opts)


def ensure_automation_tab(driver):
    """
    Ensure Selenium controls a dedicated ARMS tab; re-select or create if needed.
    Returns the window handle.
    """
    global AUTOMATION_HANDLE
    try:
        if AUTOMATION_HANDLE and AUTOMATION_HANDLE in driver.window_handles:
            driver.switch_to.window(AUTOMATION_HANDLE)
            return AUTOMATION_HANDLE
    except (NoSuchWindowException, WebDriverException):
        pass

    try:
        for h in driver.window_handles:
            try:
                driver.switch_to.window(h)
                cur = driver.current_url or ""
                if ARMS_HOST in cur:
                    AUTOMATION_HANDLE = h
                    return AUTOMATION_HANDLE
            except Exception:
                continue
    except Exception:
        pass

    try:
        handles = driver.window_handles
        if handles:
            driver.switch_to.window(handles[0])
        driver.execute_script("window.open('about:blank','_blank');")
        AUTOMATION_HANDLE = driver.window_handles[-1]
        driver.switch_to.window(AUTOMATION_HANDLE)
        driver.get(REQUESTS_URL)
        return AUTOMATION_HANDLE
    except Exception as e:
        raise RuntimeError(f"Could not create/switch to automation tab: {e}")


def run_with_retries(fn, attempts=3, base_sleep=1.0, recover=None):
    """Run a callable with retries for transient Selenium errors."""
    retry_on = (
        TimeoutException,
        StaleElementReferenceException,
        ElementClickInterceptedException,
        NoSuchWindowException,
        WebDriverException,
    )

    last_exc = None
    for i in range(1, attempts + 1):
        try:
            return fn()
        except retry_on as e:
            last_exc = e
            if recover:
                try:
                    recover(e, i)
                except Exception:
                    pass
            time.sleep(base_sleep * i)
    raise last_exc


def safe_get(driver, url, attempts=3):
    """Navigate to URL with retries while maintaining the automation tab."""
    def _go():
        ensure_automation_tab(driver)
        driver.get(url)

    run_with_retries(_go, attempts=attempts, base_sleep=1.0, recover=lambda e, n: ensure_automation_tab(driver))


def wait_for_processing_to_finish(driver, timeout=30):
    """Wait until the table processing overlay becomes invisible."""
    ensure_automation_tab(driver)
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.invisibility_of_element_located((By.ID, "packages_table_processing")))
    except TimeoutException:
        pass


def set_datatable_page_length(driver, length=-1, timeout=20):
    """Set DataTables page length using API or dropdown (-1 = All)."""
    ensure_automation_tab(driver)
    js = r"""
(function(len){
  var tableEl = document.querySelector('#packages_table');
  if (!tableEl) return {ok:false, msg:'table not found'};

  // Try DataTables API
  var dt = null;
  try {
    dt = (window.jQuery && window.jQuery.fn && window.jQuery.fn.dataTable)
         ? window.jQuery(tableEl).DataTable()
         : null;
  } catch(e){ dt = null; }

  if (dt) {
    dt.page.len(len).draw(false);
    return {ok:true, msg:'set via DataTables API', len:len};
  }

  // Fallback: dropdown
  var sel = document.querySelector('#packages_table_length select');
  if (sel) {
    sel.value = String(len);
    sel.dispatchEvent(new Event('change', {bubbles:true}));
    return {ok:true, msg:'set via dropdown', len:len};
  }
  return {ok:false, msg:'no API and no dropdown'};
})(arguments[0]);
"""
    driver.execute_script(js, length)
    wait_for_processing_to_finish(driver, timeout=timeout)
    time.sleep(0.2)


def apply_global_search(driver, target_text, match_mode="equals", timeout=40):
    """Apply global search on DataTables using regex/plain based on mode."""
    ensure_automation_tab(driver)
    js = r"""
(function(query, mode){
  function escapeRegex(s){
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  var tableEl = document.querySelector('#packages_table');
  if (!tableEl) return {usedApi:false, usedInput:false, message:'table not found'};

  var dt = null;
  try {
    dt = (window.jQuery && window.jQuery.fn && window.jQuery.fn.dataTable)
         ? window.jQuery(tableEl).DataTable()
         : null;
  } catch (e) { dt = null; }

  if (dt) {
    var settings   = dt.settings ? dt.settings()[0] : null;
    var serverSide = settings ? !!settings.oFeatures.bServerSide : false;

    if (!serverSide) {
      if (mode === 'equals') {
        var rx = '^' + escapeRegex(query) + '$';
        dt.search(rx, true, false).draw(false);
        return {usedApi:true, regex:true, serverSide:false, message:'client-side regex equals search'};
      } else if (mode === 'startswith') {
        var rx2 = '^' + escapeRegex(query);
        dt.search(rx2, true, false).draw(false);
        return {usedApi:true, regex:true, serverSide:false, message:'client-side regex prefix search'};
      }
    }
    dt.search(query).draw(false);
    return {usedApi:true, regex:false, serverSide:serverSide, message:'api plain search applied'};
  }

  var input = document.querySelector('div.dataTables_filter input[type="search"]') ||
              document.querySelector('input[type="search"]');
  if (input) {
    input.value = query;
    input.dispatchEvent(new Event('input',  { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    return {usedApi:false, usedInput:true, message:'input search applied'};
  }
  return {usedApi:false, usedInput:false, message:'no search control'};
})(arguments[0], arguments[1]);
"""
    res = driver.execute_script(js, target_text, match_mode)
    wait_for_processing_to_finish(driver, timeout=timeout)
    time.sleep(0.2)
    print("Search result:", res)


def get_new_approver_links_for_account_name(driver, account_name, timeout=30):
    """Collect unique 'New approver' links from visible rows."""
    ensure_automation_tab(driver)
    wait_for_processing_to_finish(driver, timeout=timeout)

    wait = WebDriverWait(driver, timeout)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#packages_table tbody")))

    tbody = driver.find_element(By.CSS_SELECTOR, "#packages_table tbody")
    rows = tbody.find_elements(By.CSS_SELECTOR, "tr")

    target_norm = str(account_name).strip().lower()  # retained for future filter use if needed

    links = []
    seen = set()

    for r in rows:
        if not r.is_displayed():
            continue

        tds = r.find_elements(By.CSS_SELECTOR, "td")
        if not tds:
            continue

        try:
            a = r.find_element(By.XPATH, ".//a[normalize-space()='New approver' or contains(normalize-space(.),'New approver')]")
        except NoSuchElementException:
            continue

        href = a.get_attribute("href")
        if not href:
            continue

        abs_url = urljoin(driver.current_url, href)
        if abs_url not in seen:
            seen.add(abs_url)
            links.append(abs_url)

    if not links:
        raise NoSuchElementException(f"No 'New approver' links found for Account name: {account_name}")

    return links


def wait_for_add_approver_page(driver, timeout=30):
    """Wait until approver input is present on Add Approver page."""
    ensure_automation_tab(driver)
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.presence_of_element_located((By.ID, "approver_value_input")))


def select_from_suggestions(driver, typed_query, timeout=20):
    """Select an autocomplete item that contains the typed query."""
    ensure_automation_tab(driver)
    wait = WebDriverWait(driver, timeout)

    try:
        wait.until(
            EC.any_of(
                EC.visibility_of_any_elements_located((By.CSS_SELECTOR, "ul.suggest-list li")),
                EC.visibility_of_any_elements_located((By.CSS_SELECTOR, "ul.ui-autocomplete li"))
            )
        )
    except TimeoutException:
        time.sleep(0.5)

    items = driver.find_elements(By.CSS_SELECTOR, "ul.suggest-list li, ul.ui-autocomplete li")
    if not items:
        return False

    typed_norm = typed_query.strip().lower()
    chosen = None
    for it in items:
        txt = (it.text or "").strip().lower()
        if typed_norm and typed_norm in txt:
            chosen = it
            break
    if chosen is None:
        chosen = items[0]

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", chosen)
        time.sleep(0.1)
        chosen.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", chosen)
        except Exception:
            inp = driver.find_element(By.ID, "approver_value_input")
            inp.send_keys(Keys.ARROW_DOWN)
            inp.send_keys(Keys.ENTER)

    try:
        # ensure hidden 'approver_value' is populated
        wait.until(lambda d: d.execute_script("""
            var el = document.querySelector("input[name='approver_value']");
            return !!(el && el.value && el.value.trim().length > 0);
        """))
        return True
    except TimeoutException:
        return False


def fill_and_submit_approver(driver, approver_query, timeout=40):
    """Populate approver field from suggestions and submit."""
    ensure_automation_tab(driver)

    wait_for_add_approver_page(driver, timeout=timeout)
    wait = WebDriverWait(driver, timeout)

    inp = wait.until(EC.element_to_be_clickable((By.ID, "approver_value_input")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)

    inp.clear()
    time.sleep(0.1)
    inp.send_keys(approver_query)
    time.sleep(0.4)

    if not select_from_suggestions(driver, approver_query, timeout=timeout):
        raise RuntimeError("No suggestions found / selection failed.")

    submit_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @name='submit' and @value='Submit']"))
    )

    try:
        submit_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", submit_btn)

    try:
        wait.until(
            EC.any_of(
                EC.url_contains("/arms2/unit-owner/packages"),
                EC.presence_of_element_located((By.CSS_SELECTOR, ".alert-success, .alert-info, .alert-warning, .alert-danger")),
                EC.invisibility_of_element_located((By.ID, "approver_value_input"))
            )
        )
    except TimeoutException:
        pass


def process_one_record(driver, ou_id, account_name, approver_list,
                       progress, excel_row, match_mode="equals", timeout=50):
    """Process one Excel row: search, open each link, submit all approvers."""
    ensure_automation_tab(driver)

    key = _row_key(ou_id, account_name)

    # Resume from saved indices if present
    state = progress.get("in_progress", {}).get(key, {})
    saved_start_link_idx = int(state.get("link_index", 0))
    saved_start_approver_idx = int(state.get("approver_index", 0))

    def work():
        """Inner worker with retry wrapper."""
        ensure_automation_tab(driver)
        safe_get(driver, REQUESTS_URL)
        wait_for_processing_to_finish(driver, timeout=timeout)
        apply_global_search(driver, target_text=str(ou_id).strip(), match_mode=match_mode, timeout=timeout)
        set_datatable_page_length(driver, length=-1, timeout=timeout)

        links = get_new_approver_links_for_account_name(driver, account_name, timeout=timeout)
        print(f"Found {len(links)} row(s) for Account '{account_name}' (searched by OU ID '{ou_id}')")

        link_begin = saved_start_link_idx if saved_start_link_idx < len(links) else 0
        if saved_start_link_idx >= len(links):
            print("[WARN] Saved link_index out of range. Resetting to 0.")

        for link_idx in range(link_begin, len(links)):
            link = links[link_idx]
            print(f"  --> Processing link {link_idx+1}/{len(links)}: {link}")

            approver_begin = saved_start_approver_idx if link_idx == link_begin else 0

            for appr_idx in range(approver_begin, len(approver_list)):
                approver = approver_list[appr_idx]

                update_in_progress(progress, key, excel_row, link_idx, appr_idx)
                ensure_automation_tab(driver)
                safe_get(driver, link)

                run_with_retries(
                    lambda: fill_and_submit_approver(driver, approver, timeout=timeout),
                    attempts=3,
                    base_sleep=1.0,
                    recover=lambda e, n: ensure_automation_tab(driver)
                )

                time.sleep(PER_ITEM_DELAY)

        mark_row_completed(progress, key)
        print(f"[DONE] Completed OU ID='{ou_id}', Account='{account_name}'")

    run_with_retries(work, attempts=2, base_sleep=2.0, recover=lambda e, n: ensure_automation_tab(driver))


def main():
    """Attach to browser, load Excel, iterate rows, process, and persist progress."""
    print("Starting …")

    if BROWSER == "chrome":
        driver = get_chrome_driver_attached(REMOTE_DEBUG)
        print("Attached to existing Chrome session.")
    elif BROWSER == "edge":
        driver = get_edge_driver_attached(REMOTE_DEBUG)
        print("Attached to existing Edge session.")
    else:
        raise ValueError("Unsupported browser. Enter 'chrome' or 'edge'.")

    ensure_automation_tab(driver)

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    if OU_ID_COLUMN not in df.columns:
        raise ValueError(f"Excel column '{OU_ID_COLUMN}' not found. Available columns: {list(df.columns)}")
    if ACCOUNT_NAME_COLUMN not in df.columns:
        raise ValueError(f"Excel column '{ACCOUNT_NAME_COLUMN}' not found. Available columns: {list(df.columns)}")

    progress = load_progress()
    completed = set(progress.get("completed_keys", []))

    for idx, row in df.iterrows():
        excel_row = idx + 2  # header offset
        ou_id = str(row[OU_ID_COLUMN]).strip() if pd.notna(row[OU_ID_COLUMN]) else ""
        account_name = str(row[ACCOUNT_NAME_COLUMN]).strip() if pd.notna(row[ACCOUNT_NAME_COLUMN]) else ""

        if not ou_id:
            continue

        key = _row_key(ou_id, account_name)

        if RESUME_MODE and key in completed:
            print(f"[SKIP] Already completed: OU ID='{ou_id}', Account='{account_name}'")
            continue

        try:
            process_one_record(
                driver=driver,
                ou_id=ou_id,
                account_name=account_name,
                approver_list=APPROVER_LIST,
                progress=progress,
                excel_row=excel_row,
                match_mode=MATCH_MODE,
                timeout=50
            )
            completed.add(key)

        except Exception as e:
            print(f"[ERROR] Failed at Excel row {excel_row} (OU ID='{ou_id}', Account='{account_name}'): {e}")

            progress["last_error"] = {
                "excel_row": excel_row,
                "ou_id": ou_id,
                "account_name": account_name,
                "error": str(e),
                "time": datetime.now().isoformat(timespec="seconds"),
            }
            save_progress(progress)

            if STOP_ON_ERROR:
                break
            else:
                continue

    print("Done. Verify UI for success/alerts.")


if __name__ == "__main__":
    main()