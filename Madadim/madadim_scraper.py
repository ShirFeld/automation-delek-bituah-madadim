import os
import time
from datetime import datetime, time as dt_time

import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config

class MadadimScraper:
    def __init__(self):
        # קודי המדדים מאתר הלמ"ס
        self.cbs_indicators = {
            "מחירים לצרכן": "120010",
            "תשומה בבניה": "200010", 
            "תשומה בסלילה": "240010",
            "תשומה בחקלאות": "260010",
            "מחירים סיטונאיים(תפוקות בתעשיה)": "170010",
            "רכב פרטי ואחזקתו": "121360",
            "ביטוח רכב": "140720",
            "דלק ושמנים": "140690",
            "תיקונים וחלפים לרכב": "140725",
            "תשומה באוטובוסים": "440010",
            "תשומה בבניה למסחר ולמשרדים": "800010"
        }

        self.cbs_api_url = config.CBS_API_URL
        self.bls_url = config.BLS_URL
        self.cbs_api_headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/json",
        }
        
        # נתיב יעד לקבצים - משתמש בנתיב מקובץ הקונפיג
        self.target_path = config.MADADIM_OUTPUT_PATH
        
        # הגדרות selenium
        self.driver = None
        self.wait = None
        self._driver_error = None
    

    def get_previous_period(self):
        """מחזיר (year, month) של החודש הקודם לפי לוגיקת 15 בחודש / 18:30"""
        today = datetime.today()
        current_time = today.time()

        if today.month == 1:
            prev_month = 12
            prev_year = today.year - 1
        else:
            prev_month = today.month - 1
            prev_year = today.year

        if today.day <= 15 and current_time < dt_time(18, 30):
            prev_month -= 1
            if prev_month == 0:
                prev_month = 12
                prev_year -= 1

        return prev_year, prev_month

    def get_previous_month_filename(self):
        """חישוב שם הקובץ לפי החודש הקודם"""
        prev_year, prev_month = self.get_previous_period()
        month_str = f"{prev_month:02d}"
        year_str = f"{prev_year % 100:02d}"
        return f"madadim{month_str}{year_str}.txt"

    
    def get_file_path(self):
        """קבלת נתיב מלא לקובץ"""
        filename = self.get_previous_month_filename()
        return os.path.join(self.target_path, filename)
    
    def create_data_file(self):
        """יצירת קובץ הנתונים עם המבנה הבסיסי"""
        file_path = self.get_file_path()
        
        # וידוא שהתיקייה קיימת
        os.makedirs(self.target_path, exist_ok=True)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(f"מדדים לחודש קודם - {self.get_previous_month_filename()[7:9]}/{self.get_previous_month_filename()[9:11]}\n")
            f.write("=" * 50 + "\n\n")
            
            # כתיבת כותרות המדדים מ-CBS
            f.write("מדדים מאתר הלמ\"ס:\n")
            f.write("-" * 20 + "\n")
            for indicator_name, code in self.cbs_indicators.items():
                f.write(f"{indicator_name} ({code}): \n")
            
            f.write("\nמדד מאתר BLS:\n")
            f.write("-" * 15 + "\n")
            f.write("Consumer Price Index (CUUR0000SA0): \n")
        
        print(f"נוצר קובץ: {file_path}")
        return file_path
    
    def setup_driver(self):
        """הגדרת דפדפן Chrome עם הגנות מרביות. מחזיר True בהצלחה."""
        self._driver_error = None
        options = Options()
        
        # הגדרות אנטי-זיהוי מתקדמות
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # הגנות נוספות
        options.add_argument('--disable-extensions-file-access-check')
        options.add_argument('--disable-extensions-http-throttling')
        options.add_argument('--disable-background-networking')
        options.add_argument('--disable-sync')
        options.add_argument('--disable-default-apps')
        
        # user agent מציאותי
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        # הגדרות נוספות לייצוב
        options.add_argument('--disable-background-timer-throttling')
        options.add_argument('--disable-backgrounding-occluded-windows')
        options.add_argument('--disable-renderer-backgrounding')
        options.add_argument('--disable-features=TranslateUI')
        options.add_argument('--disable-ipc-flooding-protection')
        
        # השבתת crash reporting שיכול לגרום לסגירה
        options.add_argument('--disable-crash-reporter')
        options.add_argument('--disable-in-process-stack-traces')
        
        try:
            print("יוצר Chrome driver...")
            self.driver = webdriver.Chrome(options=options)
            print("OK Chrome driver נוצר")
            print("OK - Chrome driver נוצר")
            
            # הסרת כל המאפיינים שמסגירים אוטומציה - גרסה מתקדמת
            stealth_js = """
        // הסתרת webdriver
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        
        // הוספת plugins מזויפים מציאותיים
        Object.defineProperty(navigator, 'plugins', {
            get: () => [
                {name: 'Chrome PDF Plugin', description: 'Portable Document Format', filename: 'internal-pdf-viewer'},
                {name: 'Chromium PDF Plugin', description: 'Portable Document Format', filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai'},
                {name: 'Microsoft Edge PDF Plugin', description: 'Portable Document Format', filename: 'pdfium.dll'},
                {name: 'WebKit built-in PDF', description: 'Portable Document Format', filename: 'webkit-pdf'}
            ]
        });
        
        // שפות
        Object.defineProperty(navigator, 'languages', {get: () => ['he-IL', 'he', 'en-US', 'en']});
        Object.defineProperty(navigator, 'language', {get: () => 'he-IL'});
        
        // chrome object מתקדם
        window.chrome = {
            runtime: {
                onConnect: null,
                onMessage: null
            },
            loadTimes: function() {
                return {
                    commitLoadTime: Date.now() - Math.random() * 1000,
                    connectionInfo: 'h2',
                    finishDocumentLoadTime: Date.now() + Math.random() * 1000,
                    finishLoadTime: Date.now() + Math.random() * 1000,
                    firstPaintAfterLoadTime: 0,
                    firstPaintTime: Date.now() + Math.random() * 1000,
                    navigationType: 'Other',
                    npnNegotiatedProtocol: 'h2',
                    requestTime: Date.now() - Math.random() * 2000,
                    startLoadTime: Date.now() - Math.random() * 1000,
                    wasAlternateProtocolAvailable: false,
                    wasFetchedViaSpdy: true,
                    wasNpnNegotiated: true
                };
            }
        };
        
        // הרשאות
        Object.defineProperty(navigator, 'permissions', {
            get: () => ({
                query: () => Promise.resolve({state: 'granted'})
            })
        });
        
        // הסתרת מאפיינים של selenium
        delete navigator.__proto__.webdriver;
        
        // הוספת מאפיינים של דפדפן אמיתי
        Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
        Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});
        Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 0});
        
        // החלפת setTimeout להיראות יותר טבעי
        const originalSetTimeout = window.setTimeout;
        window.setTimeout = function(callback, delay) {
            return originalSetTimeout(callback, delay + Math.random() * 50);
        };
            """
            self.driver.execute_script(stealth_js)
            print("OK הגנות אנטי-זיהוי מתקדמות הופעלו")
            
            # הגדרת חלון
            self.driver.set_window_size(1936, 1048)  # גודל ספציפי שעבד
            print("OK גודל חלון הוגדר")
            
            self.wait = WebDriverWait(self.driver, 20)  # זמן המתנה ארוך יותר
            print("OK WebDriverWait הוגדר")
            
            print("Chrome driver מוכן לשימוש בהצלחה!")
            return True
            
        except Exception as e:
            self._driver_error = str(e)
            print(f"ERROR שגיאה בהגדרת דפדפן: {e}")
            self.driver = None
            self.wait = None
            return False
        
    def close_driver(self):
        """סגירת הדפדפן"""
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None

    def get_previous_month_number(self):
        """קבלת מספר החודש הקודם"""
        _, prev_month = self.get_previous_period()
        return prev_month

    def _format_cbs_value(self, value):
        """המרת ערך המדד למחרוזת בלי עיגול נוסף"""
        if value is None:
            return None
        if isinstance(value, float):
            return format(value, '.10f').rstrip('0').rstrip('.')
        return str(value)

    def scrape_cbs_indicator(self, indicator_name, indicator_code):
        """שליפת מדד בודד מ-API הלמ"ס"""
        target_year, target_month = self.get_previous_period()
        period = f"{target_month:02d}-{target_year}"
        params = {
            "id": indicator_code,
            "format": "json",
            "download": "false",
            "startPeriod": period,
            "endPeriod": period,
        }
        print(f"מתחיל לשלוף את המדד: {indicator_name} (קוד: {indicator_code}) לחודש {period}")
        try:
            response = requests.get(
                self.cbs_api_url,
                params=params,
                headers=self.cbs_api_headers,
                timeout=30,
            )
            response.raise_for_status()
            data = response.json()
        except Exception as e:
            print(f"ERROR שגיאה בקריאת API למדד {indicator_name}: {e}")
            return None

        month_blocks = data.get("month") or []
        dates = []
        for block in month_blocks:
            dates.extend(block.get("date") or [])

        matching = next(
            (
                item for item in dates
                if item.get("year") == target_year and item.get("month") == target_month
            ),
            None,
        )
        if not matching:
            print(f"ERROR לא נמצא ערך לחודש {period} במדד {indicator_name}")
            return None

        curr_base = matching.get("currBase") or {}
        value = curr_base.get("value")
        if value is None:
            print(f"ERROR אין currBase.value למדד {indicator_name} בחודש {period}")
            return None

        formatted = self._format_cbs_value(value)
        print(f"OK ערך המדד {indicator_name} לחודש {period}: '{formatted}'")
        return formatted

    def scrape_all_cbs_indicators(self):
        """שליפת כל המדדים: למ\"ס דרך API, BLS דרך דפדפן"""
        results = {}
        bls_value = None

        for indicator_name, indicator_code in self.cbs_indicators.items():
            value = self.scrape_cbs_indicator(indicator_name, indicator_code)
            if value:
                results[indicator_name] = value

        try:
            print("\nמתחיל לשלוף נתוני BLS...")
            if not self.setup_driver():
                err = self._driver_error or "לא ניתן להפעיל את Chrome"
                print(f"ERROR לא ניתן לשלוף BLS: {err}")
            else:
                bls_value = self.scrape_bls_cpi()
                if bls_value:
                    print(f"OK נתון BLS נשלף: {bls_value}")
                else:
                    print("ERROR לא הצליח לשלוף נתון BLS")
        finally:
            self.close_driver()

        return results, bls_value


    def scrape_bls_cpi(self):
        """שליפת נתוני CPI מאתר BLS. נכשל ויוצא אחרי BLS_FETCH_TIMEOUT_SECONDS."""
        timeout = getattr(config, "BLS_FETCH_TIMEOUT_SECONDS", 5)
        started = time.monotonic()

        def remaining():
            return timeout - (time.monotonic() - started)

        def timed_out():
            if remaining() <= 0:
                print(f"ERROR פג הזמן לשליפת מדד BLS ({timeout} שניות) — מדלגים וסוגרים")
                return True
            return False

        try:
            try:
                self.driver.current_url
            except Exception:
                print("ERROR - החלון נסגר, יוצר חלון חדש...")
                if not self.setup_driver():
                    return None

            print(f"מתחיל לשלוף נתוני CPI מאתר BLS (timeout {timeout} שניות)...")
            self.driver.set_page_load_timeout(timeout)
            self.driver.set_script_timeout(timeout)

            try:
                self.driver.get(self.bls_url)
            except TimeoutException:
                print(f"ERROR טעינת אתר BLS חרגה מ-{timeout} שניות")
                return None

            if timed_out():
                return None

            print("מחכה לטעינת הטבלה...")
            wait = WebDriverWait(self.driver, max(0.1, remaining()))
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable, table[id*='DataTables'], table")))
            except TimeoutException:
                print(f"ERROR הטבלה ב-BLS לא נטענה תוך {timeout} שניות")
                return None

            if timed_out():
                return None

            table = None
            for selector in (".dataTables_scrollBody", ".dataTable"):
                found = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if found:
                    table = found[0]
                    print(f"OK נמצאה טבלה ({selector})")
                    break
            if table is None:
                tables = self.driver.find_elements(By.TAG_NAME, "table")
                if tables:
                    table = tables[0]
                    print(f"OK נמצאה טבלה ראשונה מתוך {len(tables)} טבלאות")

            if not table:
                print("ERROR לא נמצאה טבלה מתאימה")
                return None

            prev_year, prev_month = self.get_previous_period()
            period = f"M{prev_month:02d}"
            print(f"מחפש נתונים לחודש: {period}, שנה: {prev_year}")

            if timed_out():
                return None

            try:
                rows = table.find_elements(By.TAG_NAME, "tr")
                print(f"נמצאו {len(rows)} שורות בטבלה")
            except Exception:
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table tr")
                print(f"נמצאו {len(rows)} שורות בדף")

            if not rows:
                print("ERROR לא נמצאו שורות בטבלה")
                return None

            target_value = None
            print("מחפש שורה עם הנתונים הנכונים...")

            for i, row in enumerate(rows):
                if timed_out():
                    return None
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4:
                        year_cell = cells[0].text.strip()
                        period_cell = cells[1].text.strip()
                        label_cell = cells[2].text.strip()
                        value_cell = cells[3].text.strip()
                        print(f"שורה {i}: שנה={year_cell}, period={period_cell}, label={label_cell}, value={value_cell}")
                        if year_cell == str(prev_year) and period_cell == period:
                            target_value = value_cell
                            print(f"OK נמצא נתון לחודש {period}/{prev_year}: {target_value}")
                            print(f"OK ערך BLS מקורי: '{target_value}'")
                            break
                except Exception as e:
                    print(f"שגיאה בעיבוד שורה {i}: {e}")
                    continue

            if target_value:
                print(f"OK נתון CPI נשלף בהצלחה: {target_value}")
                return target_value

            if timed_out():
                return None

            print(f"ERROR לא נמצא נתון לחודש {period}/{prev_year}")
            print("נסיון חיפוש חלופי...")
            try:
                date_element = self.driver.find_element(
                    By.XPATH,
                    f"//td[contains(text(), '{prev_year}') and following-sibling::td[contains(text(), '{period}')]]",
                )
                value_element = date_element.find_element(By.XPATH, "./following-sibling::td[2]")
                target_value = value_element.text.strip()
                print(f"OK נמצא נתון בחיפוש חלופי: {target_value}")
                print(f"OK ערך BLS חלופי מקורי: '{target_value}'")
                return target_value
            except Exception:
                print("ERROR גם החיפוש החלופי נכשל")
                return None

        except TimeoutException:
            print(f"ERROR פג הזמן לשליפת מדד BLS ({timeout} שניות) — מדלגים וסוגרים")
            return None
        except Exception as e:
            print(f"ERROR שגיאה בשליפת נתוני BLS: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def update_data_file_with_values(self, cbs_values, bls_value=None):
        """עדכון קובץ הנתונים עם הערכים שנשלפו"""
        file_path = self.get_file_path()
        
        print(f"מעדכן קובץ: {file_path}")
        print(f"מספר ערכי CBS לעדכון: {len(cbs_values) if cbs_values else 0}")
        
        # קריאת הקובץ הקיים
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # עדכון הערכים - עובדים ישירות על רשימת השורות
        for indicator_name, value in cbs_values.items():
            # חיפוש השורה עם השם והקוד
            pattern = f"{indicator_name} ({self.cbs_indicators[indicator_name]}): "
            print(f"מחפש דפוס: '{pattern}' עם ערך: {value}")
            
            for i, line in enumerate(lines):
                if line.startswith(pattern):
                    # החלפת השורה
                    lines[i] = f"{pattern}{value}\n"
                    print(f"OK עודכנה שורה {i}: {indicator_name} = {value}")
                    break
            else:
                print(f"⚠️ לא נמצא דפוס: '{pattern}'")
        
        if bls_value:
            bls_pattern = "Consumer Price Index (CUUR0000SA0): "
            print(f"מחפש דפוס BLS: '{bls_pattern}' עם ערך: {bls_value}")
            
            for i, line in enumerate(lines):
                if line.startswith(bls_pattern):
                    lines[i] = f"{bls_pattern}{bls_value}\n"
                    print(f"OK עודכנה שורה {i}: BLS = {bls_value}")
                    break
            else:
                print(f"⚠️ לא נמצא דפוס BLS: '{bls_pattern}'")
        
        # כתיבה חזרה לקובץ
        with open(file_path, 'w', encoding='utf-8') as f:
            f.writelines(lines)
        
        print(f"הקובץ עודכן בהצלחה: {file_path}")

if __name__ == "__main__":
    scraper = MadadimScraper()
    print(f"שם הקובץ לחודש קודם: {scraper.get_previous_month_filename()}")
    
    # יצירת קובץ בסיסי
    scraper.create_data_file()
    
    # שליפת כל המדדים
    print("\nמתחיל לשלוף את כל המדדים...")
    print(f"מספר מדדי CBS: {len(scraper.cbs_indicators)}")
    
    # שליפת כל המדדים
    cbs_results, bls_value = scraper.scrape_all_cbs_indicators()
    
    # עדכון הקובץ עם כל הערכים
    if cbs_results or bls_value:
        scraper.update_data_file_with_values(cbs_results, bls_value)
        print(f"\nOK סיכום התוצאות:")
        print(f"מדדי CBS שנשלפו: {len(cbs_results)}")
        for name, value in cbs_results.items():
            print(f"  - {name}: {value}")
        if bls_value:
            print(f"מדד BLS: {bls_value}")
    else:
        print("ERROR לא הצליח לשלוף שום נתון")
        
    print("שליפת כל המדדים הושלמה")
