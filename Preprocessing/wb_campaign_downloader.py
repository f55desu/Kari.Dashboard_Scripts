"""
wb_campaign_downloader.py — Скачивает отчёт по кампаниям (WB Promotions / Finances).

URL формат:
  https://cmp.wildberries.ru/campaigns/finances?from=YYYY-MM-DD&to=YYYY-MM-DD

Логика дат:
  to   = вчера
  from = вчера минус 2 месяца

Запуск:
  python wb_campaign_downloader.py [--date YYYY-MM-DD] [--output-dir PATH]

  --date       дата конца периода (по умолчанию: вчера)
  --output-dir куда класть скачанный XLSX (по умолчанию: UNC-путь из `downloads_folder`)
"""
import os
import sys
import time
import shutil
import glob
from datetime import date, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr is not None:
    sys.stderr.reconfigure(encoding="utf-8")

PROJECT_DIR = Path(__file__).parent

# Сетевой путь, который должен быть доступен всегда.
# Директория существует “как есть”, поэтому не создаём её автоматически.
downloads_folder = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"
DEFAULT_DOWNLOADS_DIR = Path(downloads_folder)

CDP_PORT   = 9222
BASE_URL   = "https://cmp.wildberries.ru/campaigns/finances"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_PATH_SECOND = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA_DIR = r"C:\automation_profile"


# ─── helpers ──────────────────────────────────────────────────────────────────

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


def _date_minus_2_months(d: date) -> date:
    """Вычитает ровно 2 месяца из даты (с учётом конца месяца)."""
    month = d.month - 2
    year  = d.year
    if month <= 0:
        month += 12
        year  -= 1
    # Обрезаем день если новый месяц короче (31→28 и т.д.)
    import calendar
    max_day = calendar.monthrange(year, month)[1]
    return d.replace(year=year, month=month, day=min(d.day, max_day))
def _date_minus_days(d: date, days: int) -> date:
    """Вычитает указанное количество дней из даты."""
    return d - timedelta(days=days)

# ─── main downloader ──────────────────────────────────────────────────────────

def download_wb_campaign_report(target_date: date = None, output_dir: Path = None) -> bool:
    print(f"[{time.strftime('%H:%M:%S')}] Starting WB Campaign Finances Downloader...")

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

    date_to   = target_date
    # date_from = _date_minus_2_months(target_date)
    date_from = _date_minus_days(target_date, 35)
    url = f"{BASE_URL}?from={date_from.strftime('%Y-%m-%d')}&to={date_to.strftime('%Y-%m-%d')}"

    print(f"[{time.strftime('%H:%M:%S')}] Date range: {date_from} → {date_to}")
    print(f"[{time.strftime('%H:%M:%S')}] URL: {url}")

    # ── Ensure Chrome is running ───────────────────────────────────────────────
    if not _ensure_chrome():
        return False
    print(f"[{time.strftime('%H:%M:%S')}] Chrome CDP ready on port {CDP_PORT}")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        context = browser.contexts[0]

        # ── Find or open tab ──────────────────────────────────────────────────
        page = None
        for pg in context.pages:
            if "cmp.wildberries.ru/campaigns/finances" in pg.url:
                page = pg
                break
        if not page:
            page = context.new_page()

        # ── Navigate to report URL ────────────────────────────────────────────
        print(f"[{time.strftime('%H:%M:%S')}] Navigating...")
        page.goto(url, wait_until="domcontentloaded", timeout=60000)

        # Wait for page to fully render (can be slow)
        print(f"[{time.strftime('%H:%M:%S')}] Waiting for page to render (up to 30s)...")
        # Try waiting for the Excel button to appear
        try:
            page.wait_for_selector(
                "button:has-text('Excel'), button:has-text('Скачать'), "
                "[class*='download'], [class*='excel']",
                timeout=30000
            )
            print(f"[{time.strftime('%H:%M:%S')}] Page loaded — found download element")
        except Exception:
            # Fallback: just wait 20 seconds
            print(f"[{time.strftime('%H:%M:%S')}] Selector not found quickly, waiting 20s flat...")
            page.wait_for_timeout(20000)

        # ── Find and click the Excel download button ───────────────────────────
        print(f"[{time.strftime('%H:%M:%S')}] Waiting for Excel button to become ENABLED (up to 120s)...")
        
        # Phase 1: Wait for button to be enabled (data may take time to load)
        enabled_selector = (
            "button:not([disabled]):has-text('Excel'), "
            "button:not([disabled]):has-text('Скачать в Excel'), "
            "button:not([disabled]):has-text('Скачать')"
        )
        try:
            page.wait_for_selector(enabled_selector, state="visible", timeout=120000)
            print(f"[{time.strftime('%H:%M:%S')}] Button is now enabled!")
        except Exception as e:
            print(f"[{time.strftime('%H:%M:%S')}] Timeout waiting for enabled button: {e}")
            # Dump button state for debug
            info = page.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll("button"))
                    .filter(b => b.innerText && (b.innerText.includes("Excel") || b.innerText.includes("Скачать")));
                return btns.map(b => ({text: b.innerText.trim(), disabled: b.disabled}));
            }''')
            print(f"[{time.strftime('%H:%M:%S')}] Current button state: {info}")
            return False

        # Phase 2: Click and capture download
        print(f"[{time.strftime('%H:%M:%S')}] Clicking Excel download button...")
        btn_strategies = [
            page.get_by_role("button", name="Скачать в Excel"),
            page.get_by_role("button", name="Excel"),
            page.locator("button:has-text('Скачать в Excel')"),
            page.locator("button:has-text('Excel')"),
            page.locator("button:has-text('Скачать')"),
        ]

        clicked = False
        for loc in btn_strategies:
            try:
                if loc.first.is_visible(timeout=2000):
                    with page.expect_download(timeout=120000) as dl_info:
                        loc.first.click(timeout=10000)
                    download = dl_info.value
                    suggested = download.suggested_filename
                    print(f"[{time.strftime('%H:%M:%S')}] Download started: {suggested}")
                    dest = DOWNLOADS_DIR / suggested
                    download.save_as(str(dest))
                    print(f"[{time.strftime('%H:%M:%S')}] SUCCESS: saved to {dest}")
                    clicked = True
                    break
            except Exception as e:
                print(f"[{time.strftime('%H:%M:%S')}] Strategy failed: {e}")
                continue

        if not clicked:
            info = page.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll("button, a"));
                return btns
                    .filter(b => b.innerText && (b.innerText.includes("Excel") || b.innerText.includes("Скачать")))
                    .map(b => ({tag: b.tagName, text: b.innerText.trim().substring(0, 60), top: Math.round(b.getBoundingClientRect().top)}));
            }''')
            print(f"[{time.strftime('%H:%M:%S')}] Buttons found by JS: {info}")
            print(f"[{time.strftime('%H:%M:%S')}] ERROR: Could not click Excel download button")
            return False


    return True


# ─── CLI entry ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Download WB Campaign Finances Excel Report")
    parser.add_argument("--date", type=str, default=None,
                        help="End date in YYYY-MM-DD format (default: yesterday)")
    parser.add_argument("--output-dir", type=str, default=None,
                        help="Directory to save the downloaded XLSX (default: ./downloads)")
    args = parser.parse_args()

    target = None
    if args.date:
        from datetime import datetime
        target = datetime.strptime(args.date, "%Y-%m-%d").date()

    success = download_wb_campaign_report(target_date=target, output_dir=args.output_dir)
    os._exit(0 if success else 1)
