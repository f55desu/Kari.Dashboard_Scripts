"""
ozon_ad_report_downloader.py — Скачивает отчёт по рекламным кампаниям Ozon.

URL: https://seller.ozon.ru/app/advertisement/product/overview

Шаги:
  1. Навигация на URL
  2. Клик "Скачать отчёт" (ё!)
  3. В попапе: клик на поле "Выберите период" (label поверх input)
  4. Клик "Вчера" в datepicker
  5. Клик "Сформировать"
  6. Ждём скачивания файла, сохраняем в output_dir

Запуск:
  python ozon_ad_report_downloader.py [--output-dir PATH] [--date YYYY-MM-DD]
"""
import sys
import os
import re
import time
import logging
from datetime import date, timedelta, datetime
from pathlib import Path
from playwright.sync_api import sync_playwright

# sys.stdout.reconfigure(encoding="utf-8")
# sys.stderr.reconfigure(encoding="utf-8")

PROJECT_DIR          = Path(__file__).parent
downloads_folder = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"
DEFAULT_DOWNLOADS_DIR = Path(downloads_folder)

LOG_FILE = os.path.join(PROJECT_DIR, "assembly_ozon.log")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8' # Для корректного отображения кириллицы
)
log = logging.getLogger(__name__)

CDP_PORT    = 9222
REPORT_URL  = "https://seller.ozon.ru/app/advertisement/product/overview"
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


def _ensure_chrome(port: int = CDP_PORT) -> bool:
    if _is_cdp_available(port):
        return True
    import subprocess
    if Path(CHROME_PATH).is_file():
        print("Chrome is okay...")
        subprocess.Popen([
        CHROME_PATH,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={CHROME_USER_DATA_DIR}",
        "--no-first-run", "--no-default-browser-check",
    ])
    elif Path(CHROME_PATH_SECOND).is_file():
        print("Chrome is okay...")
        subprocess.Popen([
        CHROME_PATH_SECOND,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={CHROME_USER_DATA_DIR}",
        "--no-first-run", "--no-default-browser-check",
    ])
    else:
        raise Exception("Chrome was not found in both default folders!")
    
    for _ in range(15):
        time.sleep(1)
        if _is_cdp_available(port):
            return True
    log.info(f"[ERROR] Chrome CDP не поднялся за 15 секунд")
    return False


def download_ozon_ad_report(target_date: date = None, output_dir: Path = None) -> bool:
    log.info(f"Starting Ozon Ad Report Downloader...")

    if target_date is None:
        target_date = date.today() - timedelta(days=1)

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

    log.info(f"URL: {REPORT_URL}")

    if not _ensure_chrome():
        return False
    log.info("Chrome CDP ready")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        context = browser.contexts[0]

        # Find or open tab
        page = None
        for pg in context.pages:
            if "advertisement/product/overview" in pg.url:
                page = pg
                break
        if not page:
            page = context.new_page()

        # Navigate
        log.info("Navigating...")
        page.goto(REPORT_URL, wait_until="domcontentloaded", timeout=60000)

        log.info("Waiting for page (up to 15s)...")
        try:
            # Button text: Скачать отчёт (ё!)
            page.wait_for_selector(
                "button:has-text('Скачать отчёт'), button:has-text('Скачать отчет')",
                timeout=15000
            )
        except Exception:
            page.wait_for_timeout(10000)

        # ── Step 1: Click "Скачать отчёт" ─────────────────────────────────────
        log.info("Step 1: Clicking 'Скачать отчёт'...")
        for loc in [
            page.locator("button:has-text('Скачать отчёт')").first,
            page.locator("button:has-text('Скачать отчет')").first,
        ]:
            try:
                if loc.is_visible(timeout=2000):
                    loc.click()
                    log.info("Popup opened")
                    break
            except Exception as e:
                    log.warning(f"Strategy failed: {e}")

        page.wait_for_timeout(800)

        # ── Step 2: Open the period datepicker ────────────────────────────────
        # Popup may be off-screen when run as 4th script.
        # Strategy: JS-scroll into view first → then Playwright native click (triggers React events).
        log.info("Step 2: Opening period selector...")
        opened_picker = False

        picker_selectors = [
            "label:has-text('Выберите период')",
            "input[placeholder*='дд.мм']",
        ]
        for css in picker_selectors:
            try:
                loc = page.locator(css).first
                loc.wait_for(state="attached", timeout=2000)
                # Scroll element into view so Playwright can click it
                loc.evaluate("el => el.scrollIntoView({block: 'center', inline: 'center'})")
                page.wait_for_timeout(300)  # let scroll settle
                # Now use Playwright native click — properly triggers React events
                loc.click(timeout=5000, force=False)
                opened_picker = True
                log.info(f"Period picker opened ({css})")
                break
            except Exception as e:
                log.warning(f"Strategy failed ({css}): {e}")
                try:
                    # Last resort: dispatch synthetic pointer events
                    loc.dispatch_event("pointerdown")
                    loc.dispatch_event("mousedown")
                    loc.dispatch_event("pointerup")
                    loc.dispatch_event("mouseup")
                    loc.dispatch_event("click")
                    opened_picker = True
                    log.info(f"Period picker opened via dispatch ({css})")
                    break
                except Exception as e2:
                    log.warning(f"Dispatch fallback failed: {e2}")

        if not opened_picker:
            log.error("ERROR: Could not open period picker")
            return False

        page.wait_for_timeout(500)

        # ── Step 3: Click "Вчера" ─────────────────────────────────────────────
        log.info("Step 3: Clicking 'Вчера'...")
        clicked_yesterday = False
        for loc in [
            page.get_by_text("Вчера", exact=True),
            page.locator("button:has-text('Вчера')").first,
            page.locator("span:has-text('Вчера')").first,
            page.locator("div:has-text('Вчера')").last,
        ]:
            try:
                if loc.is_visible(timeout=2000):
                    loc.click()
                    clicked_yesterday = True
                    log.info("'Вчера' selected")
                    break
            except Exception as e:
                    log.warning(f"Вчера strategy failed: {e}")

        if not clicked_yesterday:
            # Debug: dump what appeared in the picker area
            items = page.evaluate("""() => {
                return Array.from(document.querySelectorAll("button, div, span, li"))
                    .filter(el => {
                        const t = (el.innerText||'').trim();
                        const r = el.getBoundingClientRect();
                        return t && t.length < 20 && r.width > 0 && r.height > 0 && r.top < 400;
                    })
                    .map(el => ({tag: el.tagName, text: (el.innerText||'').trim(),
                                 top: Math.round(el.getBoundingClientRect().top)}))
                    .filter((x,i,a) => a.findIndex(y=>y.text===x.text)===i)
                    .sort((a,b) => a.top-b.top);
            }""")
            log.info(f"Datepicker items: {items}")
            log.error("ERROR: Could not click 'Вчера'")
            return False

        page.wait_for_timeout(500)

        # ── Step 4: Click "Сформировать" and capture download ─────────────────
        log.info("Step 4: Clicking 'Сформировать'...")
        click_time = datetime.now()
        try:
            with page.expect_download(timeout=120000) as dl_info:
                page.locator("button:has-text('Сформировать')").first.click()
            download = dl_info.value
            suggested = download.suggested_filename
            log.info(f"Download: {suggested}")
            dest = DOWNLOADS_DIR / suggested
            download.save_as(str(dest))
            log.info(f"SUCCESS: saved to {dest}")
        except Exception as e:
            log.error(f"ERROR downloading: {e}")
            return False

    return True


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Download Ozon Advertisement Report")
    parser.add_argument("--date", type=str, default=None,
                        help="Target date YYYY-MM-DD (informational)")
    parser.add_argument("--output-dir", type=str, default=None,
                        help="Directory to save the downloaded file (default: ./downloads)")
    
    # Используем parse_known_args вместо parse_args
    args, unknown = parser.parse_known_args()

    target = None
    if args.date:
        target = datetime.strptime(args.date, "%Y-%m-%d").date()

    success = download_ozon_ad_report(target_date=target, output_dir=args.output_dir)
    os._exit(0 if success else 1)
