"""
ozon_analytics_downloader.py — Скачивает аналитику продаж Ozon.

URL: https://seller.ozon.ru/app/analytics/graphs-old?__rr=9&abt_att=1&datePreset=yesterday
Дата уже вшита в URL (datePreset=yesterday), менять ничего не нужно.

Шаги:
  1. Навигация на URL
  2. Ждём загрузки страницы
  3. Кликаем кнопку "Скачать"
  4. В выпавшем dropdown кликаем "Отчет за период"
  5. Ждём загрузки файла, сохраняем в output_dir

Запуск:
  python ozon_analytics_downloader.py [--output-dir PATH] [--date YYYY-MM-DD]
  (--date не влияет на URL, но используется для именования файла в конфиге)
"""
import sys
import re
import os
import time
import logging
from datetime import date, timedelta, datetime
from pathlib import Path
from playwright.sync_api import sync_playwright

# sys.stdout.reconfigure(encoding="utf-8")
# sys.stderr.reconfigure(encoding="utf-8")

PROJECT_DIR = Path(__file__).parent
downloads_folder = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"
DEFAULT_DOWNLOADS_DIR = Path(downloads_folder)

# LOG_FILE = PROJECT_DIR / "config" / "ozon_analytics.log"

logging.basicConfig(
    filename='assembly_ozon.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8' # Для корректного отображения кириллицы
)
log = logging.getLogger(__name__)

CDP_PORT    = 9222
REPORT_URL  = "https://seller.ozon.ru/app/analytics/graphs-old?__rr=9&abt_att=1&datePreset=yesterday"
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
    print(f"[ERROR] Chrome CDP не поднялся за 15 секунд")
    return False


def _validate_filename_timestamp(filename: str, click_time: datetime, tolerance_seconds: int = 60) -> bool:
    """
    Парсит метку времени из имени файла Ozon и проверяет,
    что она находится в пределах ±tolerance_seconds от click_time.

    Ожидаемый формат: analytics_report_YYYY-MM-DD_HH_MM[_SS].xlsx
    Пример:           analytics_report_2026-03-18_10_02.xlsx
    """
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})_(\d{2})_(\d{2})(?:_(\d{2}))?", filename)
    if not m:
        log.warning(f"Не удалось распарсить timestamp из '{filename}' — валидация пропущена")
        return True  # Не можем проверить — считаем OK

    year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    hour, minute     = int(m.group(4)), int(m.group(5))
    second           = int(m.group(6)) if m.group(6) else 0

    file_ts = datetime(year, month, day, hour, minute, second)
    delta   = abs((file_ts - click_time).total_seconds())

    if delta <= tolerance_seconds:
        log.info(f"Timestamp файла {file_ts.strftime('%H:%M:%S')} "
                  f"в пределах {int(delta)}s от click_time {click_time.strftime('%H:%M:%S')} (±{tolerance_seconds}s) OK")
        return True
    else:
        log.warning(f"Timestamp файла {file_ts.strftime('%H:%M:%S')} "
                     f"отличается на {int(delta)}s от click_time {click_time.strftime('%H:%M:%S')} "
                     f"(допуск ±{tolerance_seconds}s) — возможно скачан старый отчёт!")
        return False


def download_ozon_analytics_report(target_date: date = None, output_dir: Path = None) -> bool:
    log.info("Starting Ozon Analytics Downloader...")

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

    log.info(f"Target date: {target_date}")
    log.info(f"URL: {REPORT_URL}")

    if not _ensure_chrome():
        return False
    log.info(f"Chrome CDP ready on port {CDP_PORT}")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        context = browser.contexts[0]

        # Find or open tab
        page = None
        for pg in context.pages:
            if "seller.ozon.ru/app/analytics" in pg.url:
                page = pg
                break
        if not page:
            page = context.new_page()

        # Navigate — всегда с полным reload чтобы не было stale-state от предыдущих скриптов
        log.info("Navigating (full reload)...")
        page.goto(REPORT_URL, wait_until="domcontentloaded", timeout=60000)

        # Wait for page to render — увеличен до 45s (runner 3-й скрипт = чистый старт)
        log.info("Waiting for page content (up to 45s)...")
        try:
            page.wait_for_selector(
                "button:has-text('Скачать'), [class*='download'], button:has-text('Download')",
                timeout=45000
            )
            log.info("Page loaded — download button found")
        except Exception:
            log.warning("Download button not found in 45s, waiting 10s flat...")
            page.wait_for_timeout(10000)

        # Step 1: Click the "Скачать" button to open dropdown
        log.info("Step 1: Clicking 'Скачать' button...")
        download_btn_strategies = [
            page.get_by_role("button", name="Скачать"),
            page.locator("button:has-text('Скачать')").first,
        ]

        opened_dropdown = False
        for loc in download_btn_strategies:
            try:
                if loc.is_visible(timeout=3000):
                    loc.click()
                    opened_dropdown = True
                    log.info("Dropdown opened")
                    break
            except Exception as e:
                log.warning(f"Strategy failed: {e}")

        if not opened_dropdown:
            # JS fallback
            opened_dropdown = page.evaluate('''() => {
                const btn = Array.from(document.querySelectorAll("button"))
                    .find(b => b.innerText && b.innerText.trim() === "Скачать");
                if (btn) { btn.click(); return true; }
                return false;
            }''')
            if opened_dropdown:
                log.info("Dropdown opened via JS fallback")

        if not opened_dropdown:
            log.error("ERROR: Could not find 'Скачать' button")
            # Debug
            info = page.evaluate('''() => {
                return Array.from(document.querySelectorAll("button"))
                    .filter(b => b.innerText)
                    .map(b => b.innerText.trim().substring(0, 40))
                    .filter(t => t.length > 0);
            }''')
            log.error(f"Visible buttons: {info[:15]}")
            # Скриншот для диагностики
            scr = PROJECT_DIR / "config" / f"debug_ozon_analytics_{int(time.time())}.png"
            page.screenshot(path=str(scr))
            log.error(f"Screenshot saved: {scr}")
            return False

        # Wait for dropdown to appear
        page.wait_for_timeout(800)

        # Step 2: Click "Отчёт за период" in the dropdown to trigger generation
        log.info("Step 2: Clicking 'Отчёт за период' to trigger generation...")
        report_strategies = [
            page.get_by_text("Отчёт за период", exact=True),
            page.locator("div:has-text('Отчёт за период')").last,
            page.locator("[role='menuitem']:has-text('Отчёт за период')"),
            page.locator("li:has-text('Отчёт за период')"),
        ]

        clicked_report = False
        click_time = None
        for loc in report_strategies:
            try:
                if loc.is_visible(timeout=2000):
                    click_time = datetime.now()
                    log.info(f"Click time for generation: {click_time.strftime('%H:%M:%S')}")
                    loc.click()
                    clicked_report = True
                    break
            except Exception as e:
                log.error(f"Strategy failed: {e}")

        if not clicked_report:
            log.error("ERROR: Could not click 'Отчёт за период'")
            return False

        log.info("Report generation triggered. Waiting 5s before checking Download Manager...")
        page.wait_for_timeout(5000)

        # Step 3: Open Download Manager
        log.info("Step 3: Opening Download Manager...")
        opened_dm = False
        dm_locators = [
            page.get_by_text("Открыть менеджер загрузок", exact=True),
            page.locator("text=Открыть менеджер загрузок")
        ]
        for loc in dm_locators:
            if loc.is_visible(timeout=1000):
                loc.click()
                opened_dm = True
                break
                
        if not opened_dm:
            log.info("Re-opening 'Скачать' dropdown...")
            download_btn_strategies[0].click(timeout=3000)
            page.wait_for_timeout(1000)
            for loc in dm_locators:
                if loc.is_visible(timeout=2000):
                    loc.click()
                    opened_dm = True
                    break

        if not opened_dm:
            log.error("Could not open Download Manager")
            return False

        # Step 4: Wait for report to be Ready and Download
        log.info("Step 4: Waiting for report to be 'Готов' in Download Manager...")
        
        downloaded_file = False
        for attempt in range(12):
            page.wait_for_timeout(5000)
            log.info(f"Checking Download Manager... (attempt {attempt+1}/12)")
            try:
                # Find the row containing the report and status 'Готов'
                import re
                download_btn = (
                    page.locator("div")
                    .filter(has_text=re.compile(r"analytics_report"))
                    .filter(has_text="Готов")
                    .get_by_role("button", name="Скачать")
                    .first
                )
                
                if download_btn.is_visible():
                    with page.expect_download(timeout=10000) as dl_info:
                        download_btn.click()
                else:
                    raise Exception("Report not ready or button not found")
                             
                download = dl_info.value
                suggested = download.suggested_filename
                log.info(f"Download started from Manager: {suggested}")
                dest = DOWNLOADS_DIR / suggested
                download.save_as(str(dest))
                log.info(f"SUCCESS: saved to {dest}")
                downloaded_file = True
                break
            except Exception as e:
                log.info(f"Waiting... ({e})")
                
        if not downloaded_file:
            log.error("Failed to download report after 60s")
            return False

    return True


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Download Ozon Analytics Report")
    parser.add_argument("--date", type=str, default=None,
                        help="Target date YYYY-MM-DD (informational, URL uses datePreset=yesterday)")
    parser.add_argument("--output-dir", type=str, default=None,
                        help="Directory to save the downloaded file (default: ./downloads)")
    args = parser.parse_args()

    target = None
    if args.date:
        from datetime import datetime
        target = datetime.strptime(args.date, "%Y-%m-%d").date()

    success = download_ozon_analytics_report(target_date=target, output_dir=args.output_dir)
    os._exit(0 if success else 1)
