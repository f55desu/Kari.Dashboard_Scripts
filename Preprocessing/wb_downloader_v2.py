import os
import sys
import time
import shutil
import glob
import logging
from datetime import date, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# Setup paths
PROJECT_DIR = Path(__file__).parent
downloads_folder = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"
DEFAULT_DOWNLOADS_DIR = Path(downloads_folder)
# DEFAULT_DOWNLOADS_DIR(exist_ok=True)

LOG_FILE = PROJECT_DIR / "config" / "wb_downloader.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

WB_REPORT_URL = "https://seller.wildberries.ru/content-analytics/interactive-report/main"

# CDP Connection config (from DASHBOARD/ad_expenses.py)
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_PATH_SECOND = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA_DIR = r"C:\automation_profile"

def _is_cdp_available(port: int) -> bool:
    import socket
    try:
        with socket.create_connection(("localhost", port), timeout=2):
            return True
    except (ConnectionRefusedError, OSError, TimeoutError):
        return False

def _ensure_chrome_running(cdp_port: int = 9222) -> bool:
    import subprocess
    if _is_cdp_available(cdp_port):
        log.info(f"Chrome CDP is already available on port {cdp_port}")
        return True

    log.info(f"Chrome CDP not found. Starting Chrome with a dedicated profile...")
    if Path(CHROME_PATH).is_file():
        print("Chrome is okay...")
        chrome_cmd = [
            CHROME_PATH,
            f"--remote-debugging-port={cdp_port}",
            f"--user-data-dir={CHROME_USER_DATA_DIR}",
            "--no-first-run",
            "--disable-default-apps",
            WB_REPORT_URL
        ]
    elif Path(CHROME_PATH_SECOND).is_file():
        print("Chrome is okay...")
        chrome_cmd = [
            CHROME_PATH_SECOND,
            f"--remote-debugging-port={cdp_port}",
            f"--user-data-dir={CHROME_USER_DATA_DIR}",
            "--no-first-run",
            "--disable-default-apps",
            WB_REPORT_URL
        ]
    else:
        raise Exception("Chrome was not found in both default folders!")

    try:
        subprocess.Popen(
            chrome_cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except FileNotFoundError:
        log.error(f"ERROR: Chrome not found at {CHROME_PATH}")
        return False

    for i in range(15):
        time.sleep(1)
        if _is_cdp_available(cdp_port):
            log.info(f"Chrome started and CDP ready in {i + 1}s")
            return True

    log.error(f"ERROR: Chrome did not start CDP within 15 seconds")
    return False

def safe_click(page, locator_strategies, name="element"):
    for loc in locator_strategies:
        try:
            if loc.is_visible(timeout=3000):
                log.info(f"Found {name} using standard locator.")
                loc.click()
                return True
        except:
            pass
    return False

def download_wb_report_v2(target_date: date = None, output_dir: Path = None):
    log.info(f"Starting WB Interactive Report Download V2 (UI Flow)...")
    
    if target_date is None:
        target_date = date.today() - timedelta(days=1)
    
    # Allow caller to override output directory
    DOWNLOADS_DIR = Path(output_dir) if output_dir else DEFAULT_DOWNLOADS_DIR
    if output_dir:
        DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    else:
        # Fail-fast, если к сетевой шару нет подключения.
        try:
            DEFAULT_DOWNLOADS_DIR.stat()
        except Exception as e:
            print(f"[ERROR] Cannot access downloads folder: {downloads_folder}. {e}")
            return False
    # DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
        
    date_str = target_date.strftime("%d.%m.%Y")
    short_date_name = target_date.strftime("%d.%m")
    
    log.info(f"Target date is: {date_str} (name: {short_date_name})")

    if not _ensure_chrome_running(9222):
        print("Failed to ensure Chrome is running with CDP.")
        return False

    with sync_playwright() as p:
        log.info(f"Connecting to existing Chrome on port 9222...")
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
        except Exception as e:
            log.error(f"ERROR: Could not connect to Chrome despite checking port. Error: {e}")
            return False

        # Find WB tab or create a new one
        page = None
        for pg in context.pages:
            if "seller.wildberries.ru/content-analytics/interactive-report" in pg.url:
                page = pg
                break
        
        if not page:
            log.info(f"Opening new WB tab...")
            page = context.new_page()
            
        log.info(f"Navigating to {WB_REPORT_URL}...")
        page.goto(WB_REPORT_URL, wait_until="domcontentloaded")
        
        log.info(f"Letting page render (wait 5s)...")
        page.wait_for_timeout(5000)
        
        # 1. Click the top-most calendar picker
        log.info(f"Step 1: Clicking the top calendar (by visual position)...")
        
        try:
            # We look for typical calendar icons or buttons.
            # Wildberries often uses .Date-input__icon-button or just buttons with an SVG inside.
            # Sometimes there's a big clickable div with a 'date' class.
            selectors = [
                 "[class*='Date-input__icon-button']",  # WB uses hashed class names like Date-input__icon-button__+gq8I+3N8o
            ]
            
            found_elements = []
            for sel in selectors:
                try:
                    elements = page.locator(sel).all()
                    for el in elements:
                        box = el.bounding_box()
                        # The user confirmed the target calendar is around Y=340, wrong one is Y=431
                        # We exclude top nav/random buttons (y < 300) and footers
                        if box and box['width'] > 0 and box['height'] > 0 and 300 < box['y'] < 800:
                            found_elements.append((box['y'], box['x'], el))
                except:
                    pass
            
            if found_elements:
                # Deduplicate elements
                unique_elements = {}
                for y, x, el in found_elements:
                    key = f"{int(y)}_{int(x)}"
                    if key not in unique_elements:
                        unique_elements[key] = (y, el)
                
                sorted_by_y = sorted(unique_elements.values(), key=lambda item: item[0])
                
                top_y, target_el = sorted_by_y[0]
                log.info(f"Identified {len(sorted_by_y)} distinct calendar-like elements. Uppermost is at Y={top_y:.1f}")
                target_el.click(force=True)
            else:
                log.error(f"ERROR: Could not find any visually rendered calendar elements matching the constraints.")
                
        except Exception as e:
             log.info(f"Failed to open calendar: {e}")
             
        page.wait_for_timeout(2000)

        log.info(f"Step 2: Entering dates '{date_str}' into popup inputs...")
        try:
            date_inputs = page.locator("input[type='text'], input:not([type])").all()
            popup_inputs = []
            
            for input_el in date_inputs:
                try:
                    box = input_el.bounding_box()
                    # Popup inputs are below the calendar button (Y~450–800)
                    if box and box['width'] > 0 and box['height'] > 0 and 450 < box['y'] < 800:
                        # Skip readonly inputs (e.g. Select dropdowns)
                        if input_el.get_attribute("readonly") is not None:
                            continue
                        popup_inputs.append((box['y'], box['x'], input_el))
                except:
                    pass
            
            # Sort by Y, then X — top = Start, bottom = End
            popup_inputs.sort(key=lambda x: (x[0], x[1]))
            
            log.info(f"Found {len(popup_inputs)} popup inputs in Y=[450,800].")
            
            if len(popup_inputs) >= 2:
                # [0] = Начало периода (Start), [1] = Конец периода (End)
                start_el = popup_inputs[0][2]
                end_el   = popup_inputs[1][2]
                
                # Fill Start date
                start_el.fill("")
                start_el.fill(date_str)
                page.wait_for_timeout(300)
                
                # Fill End date — do NOT press Enter, it closes the popup
                end_el.fill("")
                end_el.fill(date_str)
                page.wait_for_timeout(300)
                
            elif len(popup_inputs) == 1:
                popup_inputs[0][2].fill("")
                popup_inputs[0][2].fill(date_str)
                page.wait_for_timeout(300)
            else:
                log.error(f"ERROR: Could not find popup inputs in Y=[450,800].")
                
        except Exception as e:
            log.info(f"Entering dates failed: {e}")

            
        page.wait_for_timeout(1000)
        
        # 3. Click "Сохранить"
        log.info(f"Step 3: Clicking 'Сохранить'...")
        save_strategies = [
            page.get_by_role("button", name="Сохранить"),
            page.locator("button:has-text('Сохранить')")
        ]
        if not safe_click(page, save_strategies, name="'Сохранить' button"):
            page.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll('button')).filter(b => b.textContent === 'Сохранить');
                if (btns.length > 0) btns[0].click();
            }''')
            
        log.info(f"Waiting 4 seconds for graph to update...")
        page.wait_for_timeout(4000)

        # 4. Click "Создать Excel" button — leftmost of 3 icon-only action buttons at Y~1199
        # Class: `button__... onlyIcon__M+CwJTc1EP` sorted by X: [0]=Excel, [1]=History, [2]=3rd
        log.info(f"Step 4: Clicking 'Создать Excel' (icon button with class onlyIcon)...")
        clicked_excel = page.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll("button"));
            const actionBtns = btns.filter(b => b.className.includes("onlyIcon__M+CwJTc1EP"));
            if (actionBtns.length === 0) return false;
            actionBtns.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
            actionBtns[0].click();  // leftmost = Создать Excel
            return true;
        }''')
        log.info(f"Excel button click result: {clicked_excel}")

        
        page.wait_for_timeout(2000)
        
        # 5. Enter filename in modal and click "Сформировать"
        log.info(f"Step 5: Checking for filename modal...")
        try:
            # The filename modal is a React Portal — appended to END of body DOM.
            # So we must use .last, NOT .first (first = calendar date inputs).
            filename_input = page.locator("input[placeholder*='Без детализации'], input[type='text']").last
            if filename_input.is_visible(timeout=3000):
                log.info(f"Filename modal appeared. Entering name '{short_date_name}'...")
                filename_input.fill(short_date_name)
                page.wait_for_timeout(300)
                
                # Click "Сформировать" button (NOT Enter — Enter goes somewhere wrong)
                confirm_strategies = [
                    page.get_by_role("button", name="Сформировать"),
                    page.locator("button:has-text('Сформировать')"),
                    page.locator("button:has-text('Создать')"),
                    page.locator("button:has-text('Ок')"),
                    page.locator("button:has-text('OK')"),
                ]
                clicked_confirm = safe_click(page, confirm_strategies, name="modal confirm button")
                if not clicked_confirm:
                    # Fallback: JS click any button in a portal/modal near the input
                    page.evaluate('''() => {
                        const btns = Array.from(document.querySelectorAll('button'));
                        const confirm_btn = btns.find(b => 
                            b.textContent && (
                                b.textContent.includes('Сформировать') ||
                                b.textContent.includes('Создать') ||
                                b.textContent.includes('Ок') ||
                                b.textContent.includes('OK')
                            )
                        );
                        if (confirm_btn) confirm_btn.click();
                    }''')
                page.wait_for_timeout(1000)
        except Exception as e:
            log.info(f"Filename modal step error: {e}")

        
        # 6 & 7. Polling downloads

        target_zip_name = f"{short_date_name} с {date_str} по {date_str}.zip"
        target_path = DOWNLOADS_DIR / target_zip_name
        
        log.info(f"Target Zip Name is expected to be: {target_zip_name}")
        log.info(f"Step 6 & 7: Polling downloads menu...")
        
        # --- INITIAL CHECK ---
        log.info(f"Performing initial check for 'Формируется' status...")
        try:
            page.evaluate('''() => {
                // Click the History/Downloads button (2nd of 3 icon-only action buttons, sorted by X)
                const btns = Array.from(document.querySelectorAll("button"))
                    .filter(b => b.className.includes("onlyIcon__M+CwJTc1EP"));
                btns.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
                if (btns.length >= 2) btns[1].click();
                else if (btns.length === 1) btns[0].click();
            }''')
        except:
            pass
            
        page.wait_for_timeout(2000)
        
        # Check initial state
        initial_state = page.evaluate(f'''(args) => {{
            const fileName = args[0];
            const dateStr = args[1];
            const elements = Array.from(document.querySelectorAll('div, li'));
            // Filter elements containing our filename and the exact date range
            const targetEl = elements.find(el => el.textContent && el.textContent.includes(fileName) && el.textContent.includes(dateStr));
            
            if (!targetEl) return "NOT_FOUND";
            if (targetEl.textContent.includes('Формируется')) return "GENERATING";
            if (targetEl.textContent.includes('МБ')) return "READY";
            return "UNKNOWN";
        }}''', [short_date_name, date_str])
        
        if initial_state == "NOT_FOUND":
            log.error(f"ERROR: Report '{short_date_name}' for {date_str} was NOT FOUND in the downloads list!")
            return False
            
        log.info(f"Report initial status: {initial_state}. Closing menu and waiting 30 seconds...")
        page.mouse.click(0, 0)
        time.sleep(30)
        
        # --- POLLING LOOP — ждём готовности файла, затем перехватываем download ---
        for attempt in range(30):  # max 30 attempts ~ 300 seconds (10s interval)
            log.info(f"Checking downloads (Attempt {attempt+1}/30)...")
            
            # Click the "Downloads/History" icon (2nd of 3 action buttons)
            try:
                page.evaluate('''() => {
                    const btns = Array.from(document.querySelectorAll("button"))
                        .filter(b => b.className.includes("onlyIcon__M+CwJTc1EP"));
                    btns.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
                    if (btns.length >= 2) btns[1].click();
                    else if (btns.length === 1) btns[0].click();
                }''')
            except:
                pass
                
            page.wait_for_timeout(2000)
            
            # Проверяем готов ли файл (без клика!)
            is_ready = page.evaluate(f'''(args) => {{
                const fileName = args[0];
                const dateStr = args[1];
                const elements = Array.from(document.querySelectorAll('div, li'));
                const targetEl = elements.find(el => el.textContent && el.textContent.includes(fileName) && !el.textContent.includes('Формируется') && el.textContent.includes('МБ') && el.textContent.includes(dateStr));
                return !!targetEl;
            }}''', [short_date_name, date_str])
            
            if is_ready:
                log.info(f"File is ready — intercepting download via expect_download()...")
                # Перехватываем загрузку через Playwright — не зависит от папки Downloads
                try:
                    with page.expect_download(timeout=60000) as dl_info:
                        # Теперь кликаем кнопку — Playwright перехватит событие
                        page.evaluate(f'''(args) => {{
                            const fileName = args[0];
                            const dateStr = args[1];
                            const elements = Array.from(document.querySelectorAll('div, li'));
                            const targetEl = elements.find(el => el.textContent && el.textContent.includes(fileName) && !el.textContent.includes('Формируется') && el.textContent.includes('МБ') && el.textContent.includes(dateStr));
                            if (targetEl) {{
                                const btns = targetEl.querySelectorAll('button');
                                for(let btn of btns) {{
                                    if (btn.textContent && btn.textContent.includes('МБ')) {{
                                        btn.click(); return;
                                    }}
                                }}
                                if(btns.length > 0) btns[0].click();
                            }}
                        }}''', [short_date_name, date_str])
                    
                    download = dl_info.value
                    log.info(f"Download intercepted: {download.suggested_filename}")
                    download.save_as(str(target_path))
                    log.info(f"SUCCESS: Saved to {target_path}")
                    break
                except Exception as e:
                    log.error(f"expect_download failed: {e}")
                    return False
                
            # Если не готов — закрываем попап и ждём
            page.mouse.click(0, 0)
            time.sleep(10)
            
        else:
            log.error("Timeout: file never became READY in WB downloads list.")
            return False

    return True

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Download WB Interactive Report")
    parser.add_argument("--date", type=str, default=None,
                        help="Target date in YYYY-MM-DD format (default: yesterday)")
    parser.add_argument("--output-dir", type=str, default=None,
                        help="Directory to save the downloaded ZIP (default: ./downloads)")
    args = parser.parse_args()
    
    target = None
    if args.date:
        from datetime import datetime
        target = datetime.strptime(args.date, "%Y-%m-%d").date()
    
    success = download_wb_report_v2(target_date=target, output_dir=args.output_dir)
    os._exit(0 if success else 1)
