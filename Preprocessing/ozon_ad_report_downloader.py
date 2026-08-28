"""
ozon_ad_report_downloader.py — Скачивает отчёт по рекламным кампаниям Ozon.

URL: https://seller.ozon.ru/app/advertisement/product/overview

Шаги:
  1. Навигация на URL
  2. Клик "Скачать отчёт" (ё!)
  3. В попапе: клик на поле "Выберите период" (label поверх input)
  4. Клик "Вчера" в datepicker
  5. Клик "Сформировать" — Ozon формирует файл на своей стороне
  6. Клик по иконке "менеджер загрузок" (слева от кнопки "Скачать отчёт")
  7. В окне "Готовые отчёты" скачиваем последний (верхний) готовый отчёт
     и сохраняем его в output_dir

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

# Сколько ждём "мгновенное" скачивание сразу после клика "Сформировать".
# Если Ozon отдаёт файл только через менеджер загрузок — таймаут просто истекает.
DIRECT_DOWNLOAD_TIMEOUT_MS  = 20000
# Ожидание готового отчёта в менеджере загрузок: 12 x 5s ~= 60s
MANAGER_ATTEMPTS            = 12
MANAGER_ATTEMPT_DELAY_MS    = 5000
MANAGER_DOWNLOAD_TIMEOUT_MS = 30000

# Кнопка "менеджер загрузок" — иконочная кнопка (без текста) слева от "Скачать отчёт".
JS_FIND_MANAGER_BUTTON = r"""
() => {
    const norm    = s  => (s || '').replace(/\s+/g, ' ').trim();
    const visible = el => { const r = el.getBoundingClientRect();
                            return r.width > 0 && r.height > 0; };
    const btns = Array.from(document.querySelectorAll("button, [role='button']"));

    // 1) по aria-label / title
    const byLabel = btns.find(b => visible(b) && /менеджер|загруз/i.test(
        (b.getAttribute('aria-label') || '') + ' ' + (b.getAttribute('title') || '')
    ));
    if (byLabel) return byLabel;

    // 2) иконочная кнопка в одном контейнере с "Скачать отчёт"
    const main = btns.find(b => /^Скачать отч[её]т$/i.test(norm(b.innerText)));
    if (main) {
        let box = main.parentElement;
        for (let d = 0; d < 4 && box; d++, box = box.parentElement) {
            const icon = Array.from(box.querySelectorAll("button, [role='button']"))
                .find(b => b !== main && !norm(b.innerText)
                        && b.querySelector('svg') && visible(b));
            if (icon) return icon;
        }
    }
    return null;
}
"""

# Кнопки скачивания строк в модалке "Готовые отчёты" (в порядке отображения —
# Ozon показывает самые свежие отчёты сверху). Текст строки кладём в el.__rowText.
JS_COLLECT_MANAGER_ROWS = r"""
() => {
    const norm    = s  => (s || '').replace(/\s+/g, ' ').trim();
    const visible = el => { const r = el.getBoundingClientRect();
                            return r.width > 0 && r.height > 0; };

    // Контейнер модалки "Готовые отчёты"
    let modal = Array.from(document.querySelectorAll("[role='dialog'], dialog"))
        .find(d => visible(d) && /Готовые отч[её]ты/.test(d.innerText || '')) || null;
    if (!modal) {
        const hits = Array.from(document.querySelectorAll('div, section'))
            .filter(n => visible(n)
                      && /Готовые отч[её]ты/.test(n.innerText || '')
                      && /\d{1,2}:\d{2}/.test(n.innerText || ''));
        hits.sort((a, b) => (a.innerText || '').length - (b.innerText || '').length);
        modal = hits[0] || null;
    }
    if (!modal) return [];

    const isHeader = t => /Готовые отч[её]ты|Последние за 24/i.test(t);

    const out = [];
    for (const b of Array.from(modal.querySelectorAll("button, [role='button'], a"))) {
        if (norm(b.innerText)) continue;         // текстовые кнопки ("Закрыть")
        if (!b.querySelector('svg')) continue;   // нужна иконка скачивания
        if (!visible(b)) continue;
        // крестик закрытия модалки
        if (/закрыть|close/i.test((b.getAttribute('aria-label') || '') + ' ' +
                                 (b.getAttribute('title') || ''))) continue;

        // Строка отчёта = ближайший предок с временем HH:MM, коротким текстом
        // и без заголовка модалки (иначе это сама модалка, а не строка).
        let row = b.parentElement, rowText = null;
        for (let d = 0; d < 5 && row; d++, row = row.parentElement) {
            const t = norm(row.innerText);
            if (t.length > 200 || isHeader(t)) break;   // вышли за пределы строки
            if (/\d{1,2}:\d{2}/.test(t)) { rowText = t; break; }
        }
        if (!rowText) continue;                  // крестик закрытия модалки и т.п.

        b.__rowText = rowText;
        out.push(b);
    }
    return out;
}
"""


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


def _save_download(download, downloads_dir: Path) -> bool:
    try:
        suggested = download.suggested_filename
        dest = Path(downloads_dir) / suggested
        download.save_as(str(dest))
        log.info(f"SUCCESS: saved to {dest}")
        return True
    except Exception as e:
        log.error(f"ERROR saving download: {e}")
        return False


def _dump_debug(page, tag: str) -> None:
    """Дамп страницы для разбора, если UI Ozon снова поменялся. Не критично."""
    try:
        (PROJECT_DIR / f"page_debug_{tag}.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(PROJECT_DIR / f"page_debug_{tag}.png"), full_page=True)
        log.error(f"Debug dump: page_debug_{tag}.html / .png")
    except Exception as e:
        log.warning(f"Debug dump failed: {e}")


def _wait_manager_modal(page, timeout_ms: int = 8000) -> bool:
    """Ждём заголовок модалки менеджера загрузок ("Готовые отчёты")."""
    try:
        page.get_by_text(re.compile(r"Готовые отч[её]ты")).first.wait_for(
            state="visible", timeout=timeout_ms
        )
        return True
    except Exception:
        return False


def _open_download_manager(page) -> bool:
    """Клик по иконке "менеджер загрузок" слева от кнопки "Скачать отчёт"."""
    if _wait_manager_modal(page, timeout_ms=1000):
        log.info("Download manager already open")
        return True

    # Стратегия 1: точечный поиск кнопки в DOM.
    el = None
    try:
        el = page.evaluate_handle(JS_FIND_MANAGER_BUTTON).as_element()
        if el:
            el.scroll_into_view_if_needed(timeout=5000)
            page.wait_for_timeout(200)
            el.click(timeout=5000)
            if _wait_manager_modal(page):
                log.info("Download manager opened (DOM lookup)")
                return True
    except Exception as e:
        log.warning(f"Manager strategy 'DOM lookup' failed: {e}")

    # Стратегия 1b: тот же элемент, но клик через JS (если что-то перекрывает кнопку).
    if el:
        try:
            el.evaluate("e => e.click()")
            if _wait_manager_modal(page):
                log.info("Download manager opened (JS click)")
                return True
        except Exception as e:
            log.warning(f"Manager strategy 'JS click' failed: {e}")

    # Стратегия 2: по доступному имени / по позиции относительно "Скачать отчёт".
    for loc in [
        page.get_by_role("button", name=re.compile(r"менеджер|загруз", re.I)),
        page.locator('button:left-of(:text("Скачать отчёт"))'),
        page.locator('button:left-of(:text("Скачать отчет"))'),
    ]:
        try:
            btn = loc.first
            if btn.is_visible(timeout=2000):
                btn.click(timeout=5000)
                if _wait_manager_modal(page):
                    log.info("Download manager opened (locator fallback)")
                    return True
        except Exception as e:
            log.warning(f"Manager fallback failed: {e}")

    return False


def _download_latest_from_manager(page, downloads_dir: Path, period_date: date) -> bool:
    """Скачивает последний (верхний) готовый отчёт из модалки "Готовые отчёты"."""
    period_str = period_date.strftime("%d.%m.%Y")
    log.info(f"Manager: looking for report with period {period_str}")

    for attempt in range(1, MANAGER_ATTEMPTS + 1):
        rows_handle, texts = None, []
        try:
            rows_handle = page.evaluate_handle(JS_COLLECT_MANAGER_ROWS)
            texts = rows_handle.evaluate("arr => arr.map(el => el.__rowText || '')")
        except Exception as e:
            log.warning(f"Attempt {attempt}: cannot read manager rows: {e}")

        if texts:
            log.info(f"Attempt {attempt}/{MANAGER_ATTEMPTS}: manager rows: {texts}")
            # "Последний файл" = верхняя строка списка; если есть строка с нужным
            # периодом — берём первую такую (она же самая свежая).
            idx = next((i for i, t in enumerate(texts) if period_str in t), 0)
            log.info(f"Selected row #{idx}: {texts[idx]}")
            try:
                el = rows_handle.get_property(str(idx)).as_element()
                el.scroll_into_view_if_needed(timeout=5000)
                with page.expect_download(timeout=MANAGER_DOWNLOAD_TIMEOUT_MS) as dl_info:
                    el.click(timeout=5000)
                download = dl_info.value
                log.info(f"Download from manager: {download.suggested_filename}")
                return _save_download(download, downloads_dir)
            except Exception as e:
                log.info(f"Attempt {attempt}: download did not start ({e})")
        else:
            # Fallback: у кнопки строки может быть доступное имя "Скачать"
            # (кнопка на самой странице называется "Скачать отчёт" — потому exact=True).
            try:
                btn = page.get_by_role("button", name="Скачать", exact=True).first
                if btn.is_visible(timeout=1000):
                    log.info(f"Attempt {attempt}: using role-based fallback button")
                    with page.expect_download(timeout=MANAGER_DOWNLOAD_TIMEOUT_MS) as dl_info:
                        btn.click(timeout=5000)
                    download = dl_info.value
                    log.info(f"Download from manager: {download.suggested_filename}")
                    return _save_download(download, downloads_dir)
            except Exception as e:
                log.info(f"Attempt {attempt}: no ready reports in manager yet ({e})")

        if attempt < MANAGER_ATTEMPTS:
            page.wait_for_timeout(MANAGER_ATTEMPT_DELAY_MS)
            # Модалка могла закрыться после клика — переоткрываем.
            if not _wait_manager_modal(page, timeout_ms=1500):
                log.info("Manager modal closed — reopening...")
                _open_download_manager(page)

    log.error("ERROR: could not download report from download manager")
    _dump_debug(page, "step6_manager_no_download")
    return False


def _close_download_manager(page) -> None:
    """Закрываем модалку, чтобы не мешала следующим скриптам в том же профиле."""
    try:
        loc = page.locator("button:has-text('Закрыть')").first
        if loc.is_visible(timeout=1000):
            loc.click(timeout=3000)
            log.info("Download manager closed")
            return
    except Exception:
        pass
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass


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

        # ── Step 4: Click "Сформировать" (Ozon формирует файл) ────────────────
        log.info("Step 4: Clicking 'Сформировать'...")
        click_time = datetime.now()
        clicked_generate = False
        download = None
        try:
            with page.expect_download(timeout=DIRECT_DOWNLOAD_TIMEOUT_MS) as dl_info:
                page.locator("button:has-text('Сформировать')").first.click(timeout=15000)
                clicked_generate = True
            download = dl_info.value
        except Exception as e:
            if not clicked_generate:
                log.error(f"ERROR: could not click 'Сформировать': {e}")
                return False
            log.info(
                f"Прямого скачивания не было за {DIRECT_DOWNLOAD_TIMEOUT_MS // 1000}s — "
                f"отчёт формируется на стороне Ozon, идём в менеджер загрузок"
            )

        # Старое поведение: файл отдался сразу — менеджер загрузок не нужен.
        if download is not None:
            log.info(f"Direct download: {download.suggested_filename}")
            return _save_download(download, DOWNLOADS_DIR)

        # ── Step 5: Открыть "менеджер загрузок" (иконка у "Скачать отчёт") ────
        log.info("Step 5: Opening download manager...")
        page.wait_for_timeout(3000)  # даём Ozon зарегистрировать отчёт в списке

        # Попап формирования отчёта мог остаться открытым и перехватывать клики.
        try:
            if page.locator("button:has-text('Сформировать')").first.is_visible(timeout=1500):
                log.info("Generate popup still open — closing it (Escape)")
                page.keyboard.press("Escape")
                page.wait_for_timeout(700)
        except Exception:
            pass

        if not _open_download_manager(page):
            _dump_debug(page, "step5_no_manager_btn")
            log.error("ERROR: Could not open download manager")
            return False

        # ── Step 6: Скачать последний готовый отчёт из менеджера ──────────────
        log.info("Step 6: Downloading latest report from download manager...")
        # На шаге 3 всегда выбирается "Вчера", значит период строки — вчерашняя дата.
        period_date = date.today() - timedelta(days=1)
        ok = _download_latest_from_manager(page, DOWNLOADS_DIR, period_date)
        _close_download_manager(page)
        if not ok:
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
