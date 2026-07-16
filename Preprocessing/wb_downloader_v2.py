import os
import sys
import time
import shutil
import glob
import logging
from datetime import date, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr is not None:
    sys.stderr.reconfigure(encoding="utf-8")

# Setup paths
PROJECT_DIR = Path(__file__).parent
downloads_folder = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"
DEFAULT_DOWNLOADS_DIR = Path(downloads_folder)
# DEFAULT_DOWNLOADS_DIR(exist_ok=True)

LOG_FILE = os.path.join(PROJECT_DIR, "wb_downloader.log")
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

def dump_page_debug(page, tag: str):
    """Сохраняет HTML страницы и снапшот скриншота для последующего анализа.
    Вызывается при подозрительных ошибках UI-флоу."""
    try:
        debug_dir = PROJECT_DIR
        html_path = debug_dir / f"page_debug_{tag}.html"
        png_path  = debug_dir / f"page_debug_{tag}.png"
        try:
            html = page.content()
            html_path.write_text(html, encoding="utf-8")
            log.info(f"  [debug] DOM saved to {html_path}")
        except Exception as e:
            log.warning(f"  [debug] DOM dump failed: {e}")
        try:
            page.screenshot(path=str(png_path), full_page=True)
            log.info(f"  [debug] Screenshot saved to {png_path}")
        except Exception as e:
            log.warning(f"  [debug] Screenshot failed: {e}")
    except Exception as e:
        log.warning(f"  [debug] dump_page_debug error: {e}")

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
        
        # 1. Open the MAIN date popup.
        # У основного поля даты есть id="dateRange". Popup открывает icon-button
        # ".Date-input__icon-button__*", которая лежит ПОВЕРХ input (intercepts pointer events).
        # Поэтому кликаем именно по icon-button, а не по input.
        # Также input для prevDateRange (Сравнение с периодом) имеет свою icon-button —
        # их надо различать. Берём ту, что в одном контейнере с #dateRange.
        log.info(f"Step 1: Opening MAIN date popup via #dateRange icon-button...")
        click_result = page.evaluate('''() => {
            const isVisible = (el) => {
                if (!el) return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0 && el.offsetParent !== null;
            };
            const dateInput = document.getElementById('dateRange');
            if (!dateInput) return { ok: false, reason: 'no #dateRange input' };
            dateInput.setAttribute('data-wb-role', 'main-date-input');

            // Поднимаемся вверх по DOM от #dateRange и ищем родителя,
            // в потомках которого есть icon-button — это и есть наш контейнер даты.
            let container = dateInput.parentElement;
            let iconBtn = null;
            for (let i = 0; i < 6 && container; i++) {
                const btn = container.querySelector("[class*='Date-input__icon-button']");
                if (btn && isVisible(btn)) {
                    // Проверяем, что эта icon-button относится именно к нашему input
                    // (а не к prevDateRange). Простая эвристика: они в одном parent.
                    if (container.contains(dateInput) && container.contains(btn)) {
                        iconBtn = btn;
                        break;
                    }
                }
                container = container.parentElement;
            }
            if (!iconBtn) {
                // Запасной вариант: все icon-button, фильтр по близости к dateInput
                const inpRect = dateInput.getBoundingClientRect();
                const allBtns = Array.from(document.querySelectorAll("[class*='Date-input__icon-button']")).filter(isVisible);
                allBtns.sort((a, b) => {
                    const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
                    const da = Math.abs(ra.top - inpRect.top) + Math.abs(ra.left - inpRect.right);
                    const db = Math.abs(rb.top - inpRect.top) + Math.abs(rb.left - inpRect.right);
                    return da - db;
                });
                iconBtn = allBtns[0] || null;
            }
            if (!iconBtn) return { ok: false, reason: 'no Date-input__icon-button near #dateRange' };

            iconBtn.setAttribute('data-wb-role', 'main-date-icon-btn');
            iconBtn.click();
            const r = iconBtn.getBoundingClientRect();
            return {
                ok: true,
                rect: { x: r.left, y: r.top, w: r.width, h: r.height },
                input_value: dateInput.value,
                input_rect: dateInput.getBoundingClientRect()
            };
        }''')
        log.info(f"  Icon-button click: {click_result}")
        if not click_result.get('ok'):
            log.error(f"❌ Не удалось открыть календарь: {click_result.get('reason')}")
            dump_page_debug(page, "step1_no_icon_btn")
            return False

        # Ждём появления popup'а (DatePickerView)
        try:
            page.wait_for_selector("[class*='DatePickerView']", state="visible", timeout=5000)
            log.info(f"  ✓ DatePickerView popup появился.")
        except Exception:
            log.warning(f"  ⚠ DatePickerView не появился за 5с. Продолжаю — fallback Strategy C/D попробуют дальше.")

        page.wait_for_timeout(800)

        log.info(f"Step 2: Setting period to '{date_str}'...")
        # WB периодически переделывает popup даты. Пробуем НЕСКОЛЬКО стратегий:
        #   (A) Подписанные поля "Начало периода" / "Конец периода" (старый popup)
        #   (B) Любые два input в popup'е, фильтруем по близости к области клика
        #   (C) Прямое заполнение основного date-input маской "DD.MM.YYYY - DD.MM.YYYY"
        #   (D) Клик по дню в календаре (по числу) дважды — стандартный способ выбора 1 дня
        # При неудаче — диагностический dump (HTML+screenshot+перечень всех visible элементов).
        period_set = False
        try:
            # === Strategy A: label-based search ===
            handles = page.evaluate('''() => {
                const allEls = Array.from(document.querySelectorAll('*'));
                const findInputAfterLabel = (labelText) => {
                    const el = allEls.find(e =>
                        e.children.length === 0 &&
                        e.textContent && e.textContent.trim() === labelText
                    );
                    if (!el) return null;
                    let parent = el.parentElement;
                    for (let i = 0; i < 5 && parent; i++) {
                        const inp = parent.querySelector('input:not([readonly])');
                        if (inp) return inp;
                        parent = parent.parentElement;
                    }
                    return null;
                };
                const startInp = findInputAfterLabel('Начало периода');
                const endInp = findInputAfterLabel('Конец периода');
                if (startInp) startInp.setAttribute('data-wb-role', 'period-start');
                if (endInp)   endInp.setAttribute('data-wb-role', 'period-end');
                return {
                    found_start: !!startInp,
                    found_end: !!endInp,
                };
            }''')
            log.info(f"  [A] Popup label search: {handles}")
            if handles.get('found_start') and handles.get('found_end'):
                start_el = page.locator('input[data-wb-role="period-start"]')
                end_el   = page.locator('input[data-wb-role="period-end"]')
                start_el.fill(""); start_el.fill(date_str); page.wait_for_timeout(300)
                end_el.fill("");   end_el.fill(date_str);   page.wait_for_timeout(300)
                period_set = True
                log.info(f"  ✓ [A] Заполнены 'Начало периода' и 'Конец периода' = {date_str}")
        except Exception as e:
            log.warning(f"  [A] failed: {e}")

        # === Strategy C: React native value setter на #dateRange ===
        # React контролирует значение input через свой state. Простой .value = "..."
        # НЕ запускает React onChange. Используем nativeInputValueSetter +
        # dispatch input event с правильным sender — это «обманывает» React.
        if not period_set:
            log.info(f"  [C] Trying React native setter on #dateRange...")
            try:
                range_str = f"{date_str} - {date_str}"
                fill_result = page.evaluate('''(args) => {
                    const value = args[0];
                    const inp = document.getElementById('dateRange');
                    if (!inp) return { ok: false, reason: 'no #dateRange' };
                    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    setter.call(inp, value);
                    inp.dispatchEvent(new Event('input', { bubbles: true }));
                    inp.dispatchEvent(new Event('change', { bubbles: true }));
                    // Также эмулируем blur — некоторые input'ы коммитят значение только по blur
                    inp.dispatchEvent(new Event('blur', { bubbles: true }));
                    return { ok: true, value_after: inp.value };
                }''', [range_str])
                log.info(f"  [C] React-setter result: {fill_result}")
                if fill_result.get('ok'):
                    page.wait_for_timeout(500)
                    # Проверяем что значение установилось и видим его в input
                    val = page.evaluate("() => document.getElementById('dateRange').value")
                    log.info(f"  [C] #dateRange.value after = {val!r}")
                    if val and date_str in val:
                        period_set = True
                        log.info(f"  ✓ [C] Direct-fill через React setter сработал.")
                    else:
                        # Иногда нужно дополнительно нажать Enter, чтобы попап закрылся
                        try:
                            page.locator('#dateRange').press("Enter")
                            page.wait_for_timeout(400)
                        except Exception:
                            pass
                        val2 = page.evaluate("() => document.getElementById('dateRange').value")
                        log.info(f"  [C] After Enter, #dateRange.value = {val2!r}")
                        if val2 and date_str in val2:
                            period_set = True
            except Exception as e:
                log.warning(f"  [C] failed: {e}")

        # === Strategy D: клик по числу в календаре дважды ===
        if not period_set:
            log.info(f"  [D] Trying calendar day-cell double-click...")
            day_num = int(target_date.strftime("%d"))
            try:
                clicked = page.evaluate(f'''(args) => {{
                    const day = args[0];
                    // Ищем кликабельные элементы (button/div/td/span) с текстом == day, видимые
                    const isVisible = (el) => {{
                        const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0 && el.offsetParent !== null;
                    }};
                    const all = Array.from(document.querySelectorAll('button, td, div, span'));
                    const matches = all.filter(el => {{
                        if (!isVisible(el)) return false;
                        if (el.children.length > 0) return false;
                        const t = (el.textContent || '').trim();
                        return t === String(day);
                    }});
                    // Фильтруем по размеру — ячейки календаря обычно 30-50px
                    const cells = matches.filter(el => {{
                        const r = el.getBoundingClientRect();
                        return r.width >= 20 && r.width <= 80 && r.height >= 20 && r.height <= 80;
                    }});
                    if (cells.length === 0) return {{ ok: false, reason: 'no cells', total_matches: matches.length }};
                    // Берём первую такую ячейку (обычно текущий месяц рендерится первым)
                    const cell = cells[0];
                    cell.click();
                    return {{
                        ok: true,
                        count: cells.length,
                        rect: cell.getBoundingClientRect(),
                        tag: cell.tagName
                    }};
                }}''', [day_num])
                log.info(f"  [D] day {day_num} click 1: {clicked}")
                if clicked.get('ok'):
                    page.wait_for_timeout(400)
                    # Второй клик по тому же числу — выбираем как конечную дату диапазона
                    clicked2 = page.evaluate(f'''(args) => {{
                        const day = args[0];
                        const isVisible = (el) => {{
                            const r = el.getBoundingClientRect();
                            return r.width > 0 && r.height > 0 && el.offsetParent !== null;
                        }};
                        const all = Array.from(document.querySelectorAll('button, td, div, span'));
                        const matches = all.filter(el => {{
                            if (!isVisible(el)) return false;
                            if (el.children.length > 0) return false;
                            const t = (el.textContent || '').trim();
                            return t === String(day);
                        }});
                        const cells = matches.filter(el => {{
                            const r = el.getBoundingClientRect();
                            return r.width >= 20 && r.width <= 80 && r.height >= 20 && r.height <= 80;
                        }});
                        if (cells.length === 0) return {{ ok: false }};
                        cells[0].click();
                        return {{ ok: true }};
                    }}''', [day_num])
                    log.info(f"  [D] day {day_num} click 2: {clicked2}")
                    if clicked2.get('ok'):
                        period_set = True
                        log.info(f"  ✓ [D] Календарь: 2 клика по {day_num}")
            except Exception as e:
                log.warning(f"  [D] failed: {e}")

        if not period_set:
            log.error(f"❌ Не удалось установить период. Сохраняю дамп для диагностики.")
            dump_page_debug(page, "step2_no_period")
            return False

            
        page.wait_for_timeout(1000)

        # Step 2.5: Verify MAIN date field contains the target date.
        # WB переехал — поле теперь ниже (Y≈400-500), и это div-widget, не input.
        # Расширяем Y-зону и допускаем поиск по textContent.
        log.info(f"Step 2.5: Verifying main date field value...")
        try:
            main_value = page.evaluate('''(args) => {
                const targetDate = args[0];
                const dateRegex = /\\d{2}\\.\\d{2}\\.\\d{4}/;
                // 1) input.value
                const inputs = Array.from(document.querySelectorAll("input"))
                    .filter(i => {
                        const r = i.getBoundingClientRect();
                        return r.top > 100 && r.top < 600 && r.left < 600 && i.value && dateRegex.test(i.value);
                    })
                    .sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
                if (inputs.length > 0) return inputs[0].value;
                // 2) текстовый widget
                const els = Array.from(document.querySelectorAll('div, span, button'))
                    .filter(el => {
                        const r = el.getBoundingClientRect();
                        if (!(r.top > 100 && r.top < 600 && r.left < 600)) return false;
                        if (el.children.length > 3) return false;
                        const t = (el.textContent || '').trim();
                        return dateRegex.test(t) && t.length < 80;
                    })
                    .sort((a,b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
                return els.length > 0 ? els[0].textContent.trim().substring(0,80) : null;
            }''', [date_str])
            log.info(f"  Main date field value: {main_value!r}")
            if main_value and date_str not in main_value:
                log.warning(f"  ⚠️ Main date field does NOT contain {date_str}. Скрипт мог открыть НЕ ТОТ календарь!")
        except Exception as e:
            log.info(f"  Verification skipped: {e}")

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

        # Step 3.5: Финальная проверка — основное поле даты должно содержать наш date_str.
        # Ищем и в input.value, и в textContent дочерних div/span (read-only widget).
        # Y-зона расширена до 600, т.к. в новой вёрстке поле смещено вниз.
        log.info(f"Step 3.5: Final check of main date field after save...")
        try:
            final_value = page.evaluate('''(args) => {
                const targetDate = args[0];
                const dateRegex = /\\d{2}\\.\\d{2}\\.\\d{4}/;
                // 1) input.value
                const inputs = Array.from(document.querySelectorAll("input"))
                    .filter(i => {
                        const r = i.getBoundingClientRect();
                        return r.top > 100 && r.top < 600 && r.left < 600 && i.value && dateRegex.test(i.value);
                    });
                if (inputs.length > 0) return inputs.sort((a,b)=>a.getBoundingClientRect().top-b.getBoundingClientRect().top)[0].value;
                // 2) div/span textContent
                const elems = Array.from(document.querySelectorAll('div, span, button'))
                    .filter(el => {
                        const r = el.getBoundingClientRect();
                        if (!(r.top > 100 && r.top < 600 && r.left < 600)) return false;
                        if (el.children.length > 3) return false;
                        const t = (el.textContent || '').trim();
                        return dateRegex.test(t) && t.length < 80;
                    });
                if (elems.length > 0) return elems.sort((a,b)=>a.getBoundingClientRect().top-b.getBoundingClientRect().top)[0].textContent.trim().substring(0,80);
                return null;
            }''', [date_str])
            log.info(f"  Final main date value: {final_value!r} (ожидалось: содержит '{date_str}')")
            if not final_value:
                log.warning(f"⚠️ Не удалось найти основное поле даты для проверки. Продолжаю (Step 4 покажет результат).")
            elif date_str not in final_value:
                # Не прерываем — Step 4/5 покажут реальный результат
                log.warning(f"⚠️ Основное поле даты содержит {final_value!r}, а не {date_str}. Продолжаю, но возможна выгрузка за другой период.")
            else:
                log.info(f"  ✓ Основное поле даты содержит целевую дату.")
        except Exception as e:
            log.warning(f"  Final check error: {e}")

        # 4. Click "Создать Excel" — icon-only button НАД таблицей товаров.
        # Таблица находится В НИЗУ страницы → сначала скроллим к ней (к тексту "Без детализации"
        # или "Итого по товарам"), иначе кнопка может быть вне viewport, а в onlyIcon-фолбэк
        # попадут иконки сайдбара слева (x≈28, y≈270).
        log.info(f"Step 4: Scrolling to table area...")
        try:
            scrolled = page.evaluate('''() => {
                const allEls = Array.from(document.querySelectorAll('div, span, button'));
                // Ищем якорь рядом с Excel-кнопкой
                const anchorTexts = ['Без детализации', 'Итого по товарам', 'Доля в выручке'];
                for (const text of anchorTexts) {
                    const el = allEls.find(e =>
                        e.children.length < 3 && (e.textContent || '').trim() === text
                    );
                    if (el) {
                        el.scrollIntoView({ block: 'center', behavior: 'instant' });
                        return { ok: true, anchor: text };
                    }
                }
                window.scrollTo(0, document.body.scrollHeight);
                return { ok: false, reason: 'no anchor — scrolled to bottom' };
            }''')
            log.info(f"  Scroll: {scrolled}")
            page.wait_for_timeout(800)
        except Exception as e:
            log.warning(f"  Scroll failed: {e}")

        log.info(f"Step 4: Clicking 'Создать Excel' (multi-strategy)...")
        excel_result = page.evaluate('''() => {
            const isVisible = (el) => {
                if (!el) return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0 && el.offsetParent !== null;
            };
            const tryClick = (el, via) => {
                if (!isVisible(el)) return null;
                el.click();
                const r = el.getBoundingClientRect();
                return { ok: true, via, rect: { x: r.left, y: r.top, w: r.width, h: r.height } };
            };

            // 0a) ПРИОРИТЕТ: класс "Download-manager__*" — это и есть Excel-кнопка
            //     (а "Download-manager-wrapper__*" — История загрузок).
            //     Уникальный селектор, который НЕ путается с dropdown "Без детализации".
            const dmContainers = Array.from(
                document.querySelectorAll("[class*='Download-manager__']:not([class*='Download-manager-wrapper'])")
            ).filter(isVisible);
            for (const dm of dmContainers) {
                const btn = dm.querySelector('button');
                if (btn && isVisible(btn)) {
                    const res = tryClick(btn, 'Download-manager-direct');
                    if (res) return res;
                }
            }

            // 0b) Контейнер 'Sales-funnel-table-header__right-buttons' (новый класс 2026-05).
            //     ВНИМАНИЕ: их на странице ДВА с одинаковым классом:
            //       [0] содержит dropdown "Без детализации" (Detalization-select-view)
            //       [1] содержит Excel/Download/Settings (Download-manager*)
            //     Игнорируем тот, что содержит Detalization-select / Select__/ input[id*=select].
            const rightBtnSelectors = [
                "[class*='Sales-funnel-table-header__right-buttons']",
                "[class*='Sales-funnel-stages-table-header-view__right-buttons']",
                "[class*='funnel'][class*='right-buttons']",
                "[class*='right-buttons']"
            ];
            for (const sel of rightBtnSelectors) {
                const containers = Array.from(document.querySelectorAll(sel)).filter(isVisible);
                if (containers.length === 0) continue;
                // Фильтруем — пропускаем контейнер с Detalization/Select dropdown'ом
                const actionContainers = containers.filter(c =>
                    !c.querySelector("[class*='Detalization-select']") &&
                    !c.querySelector("[class*='Select__']") &&
                    !c.querySelector("input[id*='select' i]")
                );
                const pool = actionContainers.length > 0 ? actionContainers : containers;
                for (const c of pool) {
                    const containerBtns = Array.from(c.querySelectorAll('button')).filter(isVisible);
                    if (containerBtns.length === 0) continue;
                    // Сортируем по X (слева направо) — Excel самый левый
                    containerBtns.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
                    const res = tryClick(containerBtns[0], 'right-buttons[' + sel + ',n=' + containerBtns.length + ',filtered=' + actionContainers.length + ']');
                    if (res) return res;
                }
            }

            // 1) aria-label / title — самый надёжный способ
            const ariaSelectors = [
                'button[aria-label="Создать Excel"]',
                'button[aria-label*="Создать Excel"]',
                'button[aria-label*="Excel"]',
                'button[title="Создать Excel"]',
                'button[title*="Создать Excel"]',
                'button[title*="Excel"]',
                'a[aria-label="Создать Excel"]',
                'a[title="Создать Excel"]',
                '[data-testid*="excel" i]',
                '[data-name*="excel" i]'
            ];
            for (const sel of ariaSelectors) {
                const el = document.querySelector(sel);
                const res = tryClick(el, 'aria:' + sel);
                if (res) return res;
            }

            // 2) Кнопка, у которой ВНУТРИ (или у соседнего tooltip) есть текст "Создать Excel"
            const allBtns = Array.from(document.querySelectorAll('button, a[role="button"]'))
                .filter(isVisible);
            const byText = allBtns.find(b => {
                const t = (b.textContent || '').trim();
                return t === 'Создать Excel' || t.includes('Создать Excel');
            });
            if (byText) {
                const res = tryClick(byText, 'text-content');
                if (res) return res;
            }

            // 3) Поиск по обёртке tooltip: ближайший <button> к элементу с текстом "Создать Excel"
            const allEls = Array.from(document.querySelectorAll('*'));
            const tooltipEl = allEls.find(e =>
                e.children.length === 0 &&
                e.textContent && e.textContent.trim() === 'Создать Excel'
            );
            if (tooltipEl) {
                // tooltip часто рендерится отдельно — ищем кнопку, на которую он указывает.
                // Самый частый паттерн: button с aria-describedby == id tooltip'а
                const tid = tooltipEl.getAttribute('id');
                if (tid) {
                    const owner = document.querySelector(`button[aria-describedby="${tid}"], button[aria-labelledby="${tid}"]`);
                    const res = tryClick(owner, 'tooltip-describedby');
                    if (res) return res;
                }
                // Иначе ищем ближайшую кнопку рядом по координатам
                const tr = tooltipEl.getBoundingClientRect();
                const near = allBtns
                    .map(b => {
                        const r = b.getBoundingClientRect();
                        const dx = (r.left + r.width/2) - (tr.left + tr.width/2);
                        const dy = (r.top + r.height/2) - (tr.top + tr.height/2);
                        return { b, dist: Math.sqrt(dx*dx + dy*dy), r };
                    })
                    .filter(x => x.dist < 200 && x.r.width < 80 && x.r.height < 80) // icon-only button рядом
                    .sort((a, b) => a.dist - b.dist);
                if (near.length > 0) {
                    const res = tryClick(near[0].b, 'tooltip-nearest');
                    if (res) return res;
                }
            }

            // 4) Поиск кнопки в группе из 3 icon-кнопок рядом с "Без детализации"
            //    "Без детализации" — это селектор детализации СЛЕВА от трёх action-кнопок.
            //    Кнопки находятся примерно в том же Y-диапазоне и правее.
            const anchorEls = Array.from(document.querySelectorAll('div, span, button')).filter(isVisible);
            const anchor = anchorEls.find(e =>
                e.children.length < 3 && (e.textContent || '').trim() === 'Без детализации'
            );
            if (anchor) {
                const ar = anchor.getBoundingClientRect();
                // Кандидаты — visible icon-only кнопки (узкие, без текста или с очень коротким),
                // в том же Y (±60px) и ПРАВЕЕ якоря
                const candidates = allBtns
                    .map(b => ({ b, r: b.getBoundingClientRect(), txt: (b.textContent || '').trim() }))
                    .filter(({r, txt}) =>
                        r.width <= 60 && r.height <= 60 &&
                        Math.abs((r.top + r.height/2) - (ar.top + ar.height/2)) < 80 &&
                        r.left > ar.right - 10 &&
                        txt.length < 5  // icon-only — текста нет либо очень мало
                    )
                    .sort((a, b) => a.r.left - b.r.left);
                if (candidates.length > 0) {
                    // Самая ЛЕВАЯ из этой группы — Excel (по скриншотам пользователя)
                    const res = tryClick(candidates[0].b, 'anchor:Без_детализации[' + candidates.length + ']');
                    if (res) return res;
                }
            }

            // 5) Фолбэк: ищем icon-кнопки ВНУТРИ заголовка таблицы воронки.
            //    Чтобы случайно не кликнуть Help/Notification — ограничиваем зоной заголовка.
            const tableHeaders = Array.from(document.querySelectorAll(
                "[class*='Sales-funnel-table-header'], [class*='Sales-funnel-stages-table-header-view'], [class*='funnel'][class*='header']"
            )).filter(isVisible);
            for (const h of tableHeaders) {
                const btns = Array.from(h.querySelectorAll('button')).filter(b =>
                    isVisible(b) && /onlyIcon/i.test(b.className || '')
                );
                if (btns.length === 0) continue;
                btns.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
                // ВНИМАНИЕ: если кнопок ровно одна — это, скорее всего, Excel.
                //           Если 2 — Excel + Download, берём левую.
                //           Если 3+ — Excel самая левая.
                const res = tryClick(btns[0], 'header-onlyIcon[n=' + btns.length + ']');
                if (res) return res;
            }

            return { ok: false, reason: 'no Excel button found by any strategy' };
        }''')
        log.info(f"  Excel button click: {excel_result}")
        if not (excel_result and excel_result.get('ok')):
            log.error(f"❌ Не удалось кликнуть 'Создать Excel'. Сохраняю дамп.")
            dump_page_debug(page, "step4_no_excel")
            return False

        page.wait_for_timeout(2500)

        # 5. Modal: enter filename and click "Сформировать"
        # Модалка — это portal, часто <div role="dialog">. Жёстко скопируем поиск
        # ВНУТРИ диалога, чтобы не залезть в случайные input'ы на основной странице.
        log.info(f"Step 5: Looking for filename modal (role=dialog or portal)...")

        modal_info = page.evaluate('''() => {
            const isVisible = (el) => {
                if (!el) return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0 && el.offsetParent !== null;
            };

            // 1) Прямой поиск диалога
            const dialogCandidates = Array.from(document.querySelectorAll(
                '[role="dialog"], [class*="Modal"], [class*="modal"], [class*="Dialog"], [class*="dialog"], [class*="Popup"], [class*="popup"]'
            )).filter(isVisible);

            // Берём именно ТОТ диалог, где есть кнопка "Сформировать" или "Отмена"
            let dialog = dialogCandidates.find(d => {
                const btns = d.querySelectorAll('button');
                return Array.from(btns).some(b =>
                    /Сформировать|Создать|Отмена/.test((b.textContent || '').trim())
                );
            });

            // 2) Если не нашли — ищем кнопку 'Сформировать' и поднимаемся до её диалога
            if (!dialog) {
                const allBtns = Array.from(document.querySelectorAll('button')).filter(isVisible);
                const submitBtn = allBtns.find(b => (b.textContent || '').trim() === 'Сформировать')
                                  || allBtns.find(b => (b.textContent || '').includes('Сформировать'));
                if (submitBtn) {
                    let p = submitBtn.parentElement;
                    for (let i = 0; i < 10 && p; i++) {
                        // Любой блок, в котором есть и input, и наша кнопка — считаем диалогом
                        if (p.querySelector('input') && p.querySelector('button')) {
                            dialog = p;
                            break;
                        }
                        p = p.parentElement;
                    }
                }
            }

            if (!dialog) {
                return { ok: false, reason: 'no modal/dialog with Сформировать button found' };
            }

            // Помечаем диалог data-атрибутом — Playwright потом сможет найти
            dialog.setAttribute('data-wb-role', 'excel-modal');

            // Помечаем поля внутри диалога
            const inputs = Array.from(dialog.querySelectorAll('input')).filter(isVisible);
            const writableInputs = inputs.filter(i => !i.disabled && !i.readOnly);
            if (writableInputs.length > 0) {
                writableInputs[0].setAttribute('data-wb-role', 'excel-modal-input');
            }
            const buttons = Array.from(dialog.querySelectorAll('button')).filter(isVisible);
            const submitBtn = buttons.find(b => /Сформировать/.test(b.textContent || ''))
                          || buttons.find(b => /Создать/.test(b.textContent || ''));
            if (submitBtn) submitBtn.setAttribute('data-wb-role', 'excel-modal-submit');

            return {
                ok: true,
                dialog_tag: dialog.tagName,
                dialog_class: (dialog.className || '').substring(0, 80),
                input_count: writableInputs.length,
                input_placeholder: writableInputs.length ? (writableInputs[0].placeholder || '') : null,
                input_value: writableInputs.length ? (writableInputs[0].value || '') : null,
                submit_text: submitBtn ? (submitBtn.textContent || '').trim() : null,
                button_count: buttons.length,
                button_texts: buttons.map(b => (b.textContent || '').trim()).filter(Boolean)
            };
        }''')
        log.info(f"  Modal search: {modal_info}")
        if not modal_info.get('ok'):
            log.error(f"❌ Не нашли модалку 'Создать Excel'. Сохраняю дамп.")
            dump_page_debug(page, "step5_no_modal")
            return False
        if not modal_info.get('submit_text'):
            log.error(f"❌ Не нашли кнопку 'Сформировать' внутри модалки. Кнопки: {modal_info.get('button_texts')}. Сохраняю дамп.")
            dump_page_debug(page, "step5_no_submit")
            return False

        try:
            modal_input = page.locator('input[data-wb-role="excel-modal-input"]')
            if modal_input.count() > 0:
                log.info(f"  Filling filename: '{short_date_name}' "
                         f"(текущее значение: {modal_info.get('input_value')!r})")
                modal_input.fill("")
                modal_input.fill(short_date_name)
                page.wait_for_timeout(400)
            else:
                log.warning(f"  ⚠ В модалке нет редактируемого input — продолжаю с автоназванием.")

            submit_btn = page.locator('button[data-wb-role="excel-modal-submit"]')
            if submit_btn.count() > 0:
                log.info(f"  Clicking 'Сформировать' button...")
                submit_btn.click()
                log.info(f"  ✓ 'Сформировать' нажата.")
            else:
                log.error(f"❌ Локатор кнопки 'Сформировать' не найден (после установки data-атрибута).")
                return False

            page.wait_for_timeout(1500)
        except Exception as e:
            log.error(f"Modal interaction failed: {e}")
            return False

        
        # 6 & 7. Polling downloads

        target_zip_name = f"{short_date_name} с {date_str} по {date_str}.zip"
        target_path = DOWNLOADS_DIR / target_zip_name
        
        log.info(f"Target Zip Name is expected to be: {target_zip_name}")
        log.info(f"Step 6 & 7: Polling downloads menu...")
        
        # --- INITIAL CHECK ---
        # Открываем менеджер загрузок — это 2-я (по X) кнопка в контейнере
        # 'Sales-funnel-table-header__right-buttons' (актуальный класс на 2026-05-20).
        # Раньше назывался 'Sales-funnel-stages-table-header-view__right-buttons' — оставлен fallback.
        # Сначала прокручиваем к контейнеру (модалка могла сместить viewport).
        log.info(f"Performing initial check for 'Формируется' status...")

        def click_downloads_btn():
            return page.evaluate('''() => {
                const isVisible = (el) => {
                    if (!el) return false;
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0 && el.offsetParent !== null;
                };
                // 1) ПРИОРИТЕТ: специфичный класс "Download-manager-wrapper__*" =
                //    кнопка "История загрузок" (рядом с Excel'ом в Download-manager__).
                const wrapper = Array.from(document.querySelectorAll("[class*='Download-manager-wrapper']")).filter(isVisible)[0];
                if (wrapper) {
                    const btn = wrapper.querySelector('button');
                    if (btn && isVisible(btn)) {
                        btn.scrollIntoView({ block: 'center', behavior: 'instant' });
                        btn.click();
                        const r = btn.getBoundingClientRect();
                        return { ok: true, via: 'Download-manager-wrapper', rect: { x: r.left, y: r.top, w: r.width, h: r.height } };
                    }
                }
                // 2) Фолбэк: контейнер right-buttons (отфильтрованный от Detalization-select)
                const selectors = [
                    "[class*='Sales-funnel-table-header__right-buttons']",
                    "[class*='Sales-funnel-stages-table-header-view__right-buttons']",
                    "[class*='funnel'][class*='right-buttons']",
                    "[class*='right-buttons']"
                ];
                for (const sel of selectors) {
                    const containers = Array.from(document.querySelectorAll(sel)).filter(isVisible);
                    const action = containers.filter(c =>
                        !c.querySelector("[class*='Detalization-select']") &&
                        !c.querySelector("[class*='Select__']") &&
                        !c.querySelector("input[id*='select' i]")
                    );
                    const pool = action.length > 0 ? action : containers;
                    for (const container of pool) {
                        container.scrollIntoView({ block: 'center', behavior: 'instant' });
                        const btns = Array.from(container.querySelectorAll('button')).filter(isVisible);
                        if (btns.length < 2) continue;
                        btns.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
                        // [0] = Excel, [1] = История загрузок, [2] = Settings
                        btns[1].click();
                        const r = btns[1].getBoundingClientRect();
                        return { ok: true, via: 'right-buttons:' + sel, count: btns.length, rect: { x: r.left, y: r.top, w: r.width, h: r.height } };
                    }
                }
                return { ok: false, reason: 'no downloads button found' };
            }''')

        # Подождём чтобы отчёт успел появиться в очереди формирования
        page.wait_for_timeout(2500)
        try:
            res = click_downloads_btn()
            log.info(f"  Downloads button click: {res}")
        except Exception as e:
            log.warning(f"  Downloads button click failed: {e}")

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
        not_found_streak = 0  # счётчик подряд отсутствующих результатов — даём ранний выход
        for attempt in range(30):  # max 30 attempts ~ 300 seconds (10s interval)
            log.info(f"Checking downloads (Attempt {attempt+1}/30)...")

            # Click the "Downloads/History" icon (2nd of 3 action buttons)
            try:
                click_downloads_btn()
            except Exception:
                pass

            page.wait_for_timeout(2000)

            # Проверяем готов ли файл (без клика!) И что вообще он есть в списке
            file_state = page.evaluate(f'''(args) => {{
                const fileName = args[0];
                const dateStr = args[1];
                const elements = Array.from(document.querySelectorAll('div, li'));
                const targetEl = elements.find(el => el.textContent && el.textContent.includes(fileName) && el.textContent.includes(dateStr));
                if (!targetEl) return "NOT_FOUND";
                if (targetEl.textContent.includes('Формируется')) return "GENERATING";
                if (targetEl.textContent.includes('МБ')) return "READY";
                return "UNKNOWN";
            }}''', [short_date_name, date_str])
            log.info(f"  state={file_state}")

            # Если файла нет в списке несколько раз подряд — значит скрипт жал НЕ ТУ кнопку,
            # запрос на формирование отчёта не отправился. Прерываем досрочно.
            if file_state == "NOT_FOUND":
                not_found_streak += 1
                if not_found_streak >= 3:
                    log.error("ERROR: Отчёт 3 проверки подряд не находится в списке загрузок. "
                              "Вероятно, на шаге 1/2 скрипт открыл НЕ ТОТ календарь "
                              "(например, поле 'Сравнение с периодом'). Прерываю.")
                    return False
                page.mouse.click(0, 0)
                time.sleep(10)
                continue
            not_found_streak = 0

            is_ready = (file_state == "READY")
            
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
