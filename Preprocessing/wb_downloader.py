import os
import sys
import time
from datetime import date, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# Setup paths
PROJECT_DIR = Path(__file__).parent
DOWNLOADS_DIR = PROJECT_DIR / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)

# Wildberries URL
WB_REPORT_URL = "https://seller.wildberries.ru/content-analytics/interactive-report/main"


def download_wb_report():
    print(f"[{time.strftime('%H:%M:%S')}] Starting WB Interactive Report download...")
    
    # Calculate yesterday's date
    yesterday = date.today() - timedelta(days=1)
    day = str(yesterday.day) # e.g., "16"
    
    # Russian month names needed for the Wildberries calendar popup
    months_ru = [
        "января", "февраля", "марта", "апреля", "мая", "июня",
        "июля", "августа", "сентября", "октября", "ноября", "декабря"
    ]
    month_ru = months_ru[yesterday.month - 1] 
    
    # The text we'll type into the popup (e.g. "16 марта")
    date_text = f"{day} {month_ru}"
    print(f"[{time.strftime('%H:%M:%S')}] Target date is: {yesterday.isoformat()} ({date_text})")

    with sync_playwright() as p:
        print(f"[{time.strftime('%H:%M:%S')}] Connecting to existing Chrome on port 9222...")
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
        except Exception as e:
            print(f"ERROR: Could not connect to Chrome. Is it running with --remote-debugging-port=9222? Error: {e}")
            return False

        # Find WB tab or create a new one
        page = None
        for pg in context.pages:
            if "seller.wildberries.ru/content-analytics/interactive-report" in pg.url:
                page = pg
                break
        
        if not page:
            print(f"[{time.strftime('%H:%M:%S')}] Opening new WB tab...")
            page = context.new_page()
            
        print(f"[{time.strftime('%H:%M:%S')}] Navigating to {WB_REPORT_URL}...")
        page.goto(WB_REPORT_URL, wait_until="domcontentloaded")
        
        # 1. Open the Calendar picker
        print(f"[{time.strftime('%H:%M:%S')}] Letting page render...")
        page.wait_for_timeout(5000)
        
        print(f"[{time.strftime('%H:%M:%S')}] Opening calendar...")
        
        # In the WB layout, the date picker has a class starting with Date-input or DatePicker
        # We can also use evaluate to click the element containing the current year
        current_year_short = str(date.today().year)[-2:] # '26'
        try:
            # Execute JS to find and click the date range button
            page.evaluate(f'''() => {{
                const btns = Array.from(document.querySelectorAll('button, div[role="button"], span[role="button"]'));
                const dateBtn = btns.find(b => b.textContent && b.textContent.includes("{current_year_short}"));
                if (dateBtn) dateBtn.click();
            }}''')
            print("Clicked calendar using JS")
        except Exception as e:
            print("JS click failed:", e)

        page.wait_for_timeout(2000)
        
        # 2. Select "Yesterday" (Вчера)
        print(f"[{time.strftime('%H:%M:%S')}] Selecting 'Yesterday'...")
        try:
            # The preset buttons are usually in a list or group. "Вчера" is almost always the first preset.
            # fallback to JS click
            page.evaluate('''() => {
                const els = Array.from(document.querySelectorAll('*'));
                const yesterdayEl = els.find(el => el.textContent === 'Вчера' && el.children.length === 0);
                if (yesterdayEl) {
                    yesterdayEl.click();
                } else {
                    // Try to find the radio button or list item
                    const presetEls = Array.from(document.querySelectorAll('label, li'));
                    const p = presetEls.find(el => el.textContent.includes('Вчера'));
                    if (p) p.click();
                }
            }''')
        except Exception as e:
            print(f"Could not click 'Yesterday': {e}")
            
        # Optional: Close popup by hitting Escape if it didn't auto-close
        page.keyboard.press("Escape")
        page.wait_for_timeout(1000)

        # 3. Click "Generate Excel"
        print(f"[{time.strftime('%H:%M:%S')}] Clicking 'Generate Excel'...")
        try:
            # We know it's a button with an SVG icon above the table.
            # Using JS to find the export buttons
            page.evaluate('''() => {
                // Find the header area usually containing "Без детализации" or similar dropdowns
                const els = Array.from(document.querySelectorAll('*'));
                const detailEl = els.find(el => el.textContent === 'Без детализации');
                if (detailEl) {
                    const container = detailEl.closest('div').parentElement.parentElement;
                    const buttons = container.querySelectorAll('button');
                    if (buttons.length >= 2) {
                        buttons[0].click(); // Generate Excel
                        return;
                    }
                }
                
                // Fallback: click the first button on the page that looks like an icon-only button
                // (classes often contain "onlyIcon")
                const iconBtns = document.querySelectorAll('button[class*="onlyIcon"]');
                if (iconBtns.length > 0) {
                    iconBtns[0].click();
                }
            }''')
        except Exception as e:
            print(f"[{time.strftime('%H:%M:%S')}] Could not click generate button: {e}")
            
        # 4. Wait for generation to finish. 
        # Wildberries shows a loading spinner or blocks the UI for 5-15 seconds.
        print(f"[{time.strftime('%H:%M:%S')}] Waiting for generation...")
        page.wait_for_timeout(10000)
        
        # 5. Type in the date in the popup if it appeared
        print(f"[{time.strftime('%H:%M:%S')}] Handling popup if any...")
        try:
            popup_input = page.locator("input[type='text']").locator("visible=true").last
            if popup_input.is_visible(timeout=3000):
                popup_input.fill(date_text)
                page.keyboard.press("Enter")
                page.wait_for_timeout(2000)
        except Exception as e:
            print("No text input popup found or error:", e)
            
        # 6. Sometimes expect_download misses hidden iframe downloads in Playwright when attached over CDP.
        
        print(f"[{time.strftime('%H:%M:%S')}] Listening for the Excel file download event...")
        # 6. Click the download button (physical fallback)
        print(f"[{time.strftime('%H:%M:%S')}] Clicking the download button via JS...")
        try:
            page.evaluate('''() => {
                const detailEls = Array.from(document.querySelectorAll('*')).filter(el => el.textContent === 'Без детализации');
                if (detailEls.length > 0) {
                    const box = detailEls[0].getBoundingClientRect();
                    const event = new MouseEvent('click', {
                        view: window,
                        bubbles: true,
                        cancelable: true,
                        clientX: box.x + 800, 
                        clientY: box.y + 10
                    });
                    document.elementFromPoint(box.x + 800, box.y + 10)?.dispatchEvent(event);
                }
                
                document.querySelectorAll('button').forEach(b => {
                    if(b.innerHTML.includes('Скачать') || b.title.includes('Скачать') || b.className.includes('download')) {
                        b.click();
                    }
                });
            }''')
            print(f"[{time.strftime('%H:%M:%S')}] Physical / SVG download button clicked via JS.")
            
            # Wait for Chrome to actually download the file to the default system Downloads folder
            print(f"[{time.strftime('%H:%M:%S')}] Waiting 10s for Chrome to download the file to the system Downloads folder...")
            time.sleep(10) # Using standard time.sleep to not tie up the Playwright event loop unnecessarily
            
            # 7. Find the downloaded file in the system Downloads folder
            import glob
            user_downloads = Path.home() / "Downloads"
            # WB files are usually named "Отчет_по_продажам..." or similar xlsx
            list_of_files = glob.glob(str(user_downloads / "*.xlsx")) 
            if not list_of_files:
                print(f"[{time.strftime('%H:%M:%S')}] ERROR: No Excel files found in {user_downloads}")
                return False
                
            # Get the most recently modified file
            latest_file = max(list_of_files, key=os.path.getctime)
            
            # Move it to our project folder
            import shutil
            file_name = f"WB_Interactive_{yesterday.strftime('%Y%m%d')}.xlsx"
            final_path = DOWNLOADS_DIR / file_name
            shutil.move(latest_file, str(final_path))
            
            print(f"[{time.strftime('%H:%M:%S')}] SUCCESS: Found and moved downloaded file to {final_path}")
            
        except Exception as e:
            print(f"[{time.strftime('%H:%M:%S')}] ERROR: Failed during download or file moving process: {e}")

        return True

if __name__ == "__main__":
    download_wb_report()
