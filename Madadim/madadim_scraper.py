import os
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime, time as dt_time
import time
from datetime import datetime, timedelta, date

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

        self.cbs_scenarios = [
            "מדד המחירים לצרכן, לפי קבוצות צריכה ראשיות",
            "מדד מחירי תשומה בבנייה למגורים",
            "מדד מחירי תשומה בסלילה וגישור, לפי קבוצות ראשיות וקבוצות משניות",
            "מדד מחירי תשומה בענף החקלאות, לפי קבוצות ראשיות וקבוצות משניות",
            "מדד המחירים הסיטוניים של תפוקת התעשייה וכרייה וחציבה ליעדים מקומיים",
            "מדד המחירים לצרכן, לפי קבוצות צריכה ראשיות ומשניות",
            "מדד המחירים לצרכן של מוצרים ושירותים נבחרים",
            "מדד המחירים לצרכן של מוצרים ושירותים נבחרים",
            "מדד המחירים לצרכן של מוצרים ושירותים נבחרים",
            "מדד מחירי תשומה באוטובוסים, כל האוכלוסייה",
            "מדד מחירי תשומה בבנייה למסחר ולמשרדים,אחוז שינוי חודשי ושנתי,מינואר 2012"

        ]
        self.current_scenario_index = 0 
        
        # אתרי המקור - משתמש בURLים מקובץ הקונפיג
        self.cbs_url = config.CBS_URL
        self.bls_url = config.BLS_URL
        
        # נתיב יעד לקבצים - משתמש בנתיב מקובץ הקונפיג
        self.target_path = config.MADADIM_OUTPUT_PATH
        
        # הגדרות selenium
        self.driver = None
        self.wait = None
        self._driver_error = None
    

    def get_previous_month_filename(self):
        """חישוב שם הקובץ לפי החודש הקודם"""
        today = datetime.today()
        current_time = today.time()

        # חישוב החודש הקודם
        if today.month == 1:
            prev_month = 12
            prev_year = today.year - 1
        else:
            prev_month = today.month - 1
            prev_year = today.year

        # אם היום <= 15 והשעה לפני 18:30, נחזור עוד חודש אחורה
        if today.day <= 15 and current_time < dt_time(18, 30):
            prev_month -= 1
            if prev_month == 0:
                prev_month = 12
                prev_year -= 1  # אם ירדנו מדצמבר לינואר, גם השנה קופצת אחורה

        # פורמט MMYY
        month_str = f"{prev_month:02d}"
        year_str = f"{prev_year % 100:02d}"

        filename = f"madadim{month_str}{year_str}.txt"
        return filename

    
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

    def _switch_to_cbs_tool_frame(self):
        """מעבר ל-iframe של מחולל המדדים (לא תמיד frame(0))."""
        self.driver.switch_to.default_content()
        time.sleep(1)

        code_field_selectors = [
            "input[ng-model='mainCtrl.codesearch']",
            "input[ng-model*='codesearch']",
            "input[placeholder*='קוד']",
        ]

        # ניסיון ישיר בדף הראשי
        for selector in code_field_selectors:
            if self.driver.find_elements(By.CSS_SELECTOR, selector):
                print(f"OK נמצא שדה קוד בדף הראשי ({selector})")
                return True

        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        print(f"נמצאו {len(iframes)} iframes, מחפש את מחולל המדדים...")
        for index in range(len(iframes)):
            try:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame(index)
                for selector in code_field_selectors:
                    if self.driver.find_elements(By.CSS_SELECTOR, selector):
                        print(f"OK נמצא מחולל מדדים ב-iframe {index} ({selector})")
                        return True
            except Exception:
                continue

        self.driver.switch_to.default_content()
        print("ERROR לא נמצא iframe של מחולל המדדים")
        return False

    def _select_search_by_code_mode(self):
        """בחירת מצב חיפוש לפי קוד."""
        from selenium.webdriver.common.action_chains import ActionChains

        radio_selectors = [
            (By.NAME, "7"),
            (By.CSS_SELECTOR, "input[type='radio'][value='7']"),
            (By.CSS_SELECTOR, "input[type='radio'][ng-value='7']"),
            (By.XPATH, "//label[contains(., 'קוד')]/preceding-sibling::input[@type='radio']"),
            (By.XPATH, "//label[contains(., 'קוד')]/input[@type='radio']"),
        ]

        for by, selector in radio_selectors:
            try:
                radio = self.wait.until(EC.presence_of_element_located((by, selector)))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radio)
                time.sleep(0.5)
                ActionChains(self.driver).move_to_element(radio).perform()
                try:
                    radio.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", radio)
                time.sleep(2)
                if radio.is_selected() or radio.get_attribute("checked"):
                    print(f"OK נבחר מצב חיפוש לפי קוד ({selector})")
                    return True
                # גם אם is_selected לא עובד, נמשיך אם לחצנו
                print(f"OK לחצנו על רדיו חיפוש לפי קוד ({selector})")
                return True
            except Exception:
                continue

        print("ERROR לא הצלחתי לבחור מצב חיפוש לפי קוד")
        return False

    def _find_code_input(self):
        """איתור שדה הזנת קוד המדד."""
        selectors = [
            "input[ng-model='mainCtrl.codesearch']",
            "input[ng-model*='codesearch']",
            "input[placeholder*='קוד']",
        ]
        for selector in selectors:
            elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
            if elements:
                return elements[0]
        return None

    def _enter_indicator_code(self, indicator_code):
        """הזנת קוד מדד ולחיצה על המשך."""
        from selenium.webdriver.common.action_chains import ActionChains

        code_field = self.wait.until(lambda d: self._find_code_input())
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", code_field)
        time.sleep(0.5)

        code_field.click()
        time.sleep(0.3)
        code_field.clear()
        time.sleep(0.3)

        # הזנה מהירה ויציבה (פחות סיכוי לאיבוד פוקוס)
        code_field.send_keys(indicator_code)
        entered = (code_field.get_attribute("value") or "").strip()
        if entered != indicator_code:
            self.driver.execute_script(
                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));",
                code_field,
                indicator_code,
            )
            entered = (code_field.get_attribute("value") or "").strip()

        if entered != indicator_code:
            print(f"ERROR הקוד לא הוזן. צפוי: {indicator_code}, בפועל: '{entered}'")
            return False

        print(f"OK קוד {indicator_code} הוזן בהצלחה")
        time.sleep(1)

        continue_selectors = [
            (By.CSS_SELECTOR, 'a.greenBigBtn[data-ng-click="mainCtrl.searchByCode();"]'),
            (By.XPATH, "//a[contains(@data-ng-click, 'searchByCode')]"),
            (By.XPATH, "//a[contains(@class, 'greenBigBtn') and contains(., 'המשך')]"),
        ]
        for by, selector in continue_selectors:
            try:
                continue_btn = self.wait.until(EC.element_to_be_clickable((by, selector)))
                continue_btn.click()
                print("OK נלחץ על המשך אחרי הזנת קוד")
                time.sleep(3)
                return True
            except Exception:
                continue

        print("ERROR לא נמצא כפתור המשך אחרי הזנת קוד")
        return False

    def _open_cbs_madadim_page(self):
        """פתיחת דף מחולל המדדים ומעבר ל-iframe הנכון."""
        if not self.driver:
            print("ERROR אין דפדפן פעיל")
            return False
        self.driver.switch_to.default_content()
        self.driver.get(self.cbs_url)
        self.driver.set_window_size(1936, 1048)
        time.sleep(4)
        return self._switch_to_cbs_tool_frame()
        
    def get_previous_month_number(self):
        """קבלת מספר החודש הקודם"""
        today = datetime.today()
        current_time = today.time()

        # חישוב החודש הקודם
        if today.month == 1:
            month = 12
        else:
            month = today.month - 1

        if today.day <= 15 and current_time < dt_time(18, 30):
            month -= 1
            if month == 0:
                month = 12  # במקרה שירדנו מדצמבר לינואר

        return month

        
    
    def scrape_cbs_indicator(self, indicator_name, indicator_code, scenario_text=None):
        """שליפת מדד בודד מאתר הלמ"ס"""
        try:
            # בדיקה אם החלון עדיין פתוח
            try:
                self.driver.current_url
            except Exception:
                print("ERROR - החלון נסגר, יוצר חלון חדש...")
                if not self.setup_driver():
                    return None
            
            if not self.driver:
                return None
            
            print(f"מתחיל לשלוף את המדד: {indicator_name} (קוד: {indicator_code})")
            
            if not self._open_cbs_madadim_page():
                return None

            if not self._select_search_by_code_mode():
                return None

            print(f"מכניס קוד {indicator_code}...")
            if not self._enter_indicator_code(indicator_code):
                return None

            
            # שלב 4: בחירת הנושא-variableBox - גנרי
            print("בוחר את הנושא השני...")
            try:
                # מחפש את כל הנושאים ב-variableBox
                print("מחפש את כל הנושאים...")
                topics = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.variableBox .variableBoxInner ul li a'))
                )
                print(f"נמצאו {len(topics)} נושאים")
                
                # בדיקה שיש לפחות 2 נושאים
                if len(topics) < 2:
                    print("ERROR לא נמצאו מספיק נושאים")
                    return None
                
                # בחירת הנושא השני (אינדקס 1)
                second_topic = topics[1]
                topic_text = second_topic.text
                print(f"בוחר את הנושא השני: {topic_text}")
                
                try:
                    second_topic.click()
                    print("OK נבחר הנושא השני בלחיצה רגילה")
                except:
                    # אם נכשל, ננסה JavaScript
                    self.driver.execute_script("arguments[0].click();", second_topic)
                    print("OK נבחר הנושא השני עם JavaScript")
                
                time.sleep(3)
            except Exception as e:
                print(f"ERROR שגיאה בבחירת הנושא השני: {e}")
                return None

            
            # לחיצה על כפתור המשך לבחירת הנושאים
            next_arrow = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'img[src*="nextArrow"]'))
            )
            next_arrow.click()
            time.sleep(2)
            
            # שלב 5: בחירת תת-נושא המתאים למדד הנוכחי בלבד
            print("בוחר תת נושא...")
            try:
                subtopics = self.wait.until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, 'div.jspPane ul.scroll-pane-inner li a.ellipsis.ng-binding')
                    )
                )
                print(f"נמצאו {len(subtopics)} תת נושאים")

                if not subtopics:
                    print("ERROR לא נמצאו תת נושאים")
                    return None

                if scenario_text:
                    target = next((t for t in subtopics if scenario_text in (t.text or "")), None)
                    if target:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
                        time.sleep(0.5)
                        try:
                            target.click()
                        except Exception:
                            self.driver.execute_script("arguments[0].click();", target)
                        print(f"OK נבחר תת נושא: {scenario_text}")
                        time.sleep(1)
                    else:
                        print(f"ERROR לא נמצא תת נושא: {scenario_text}")
                        return None
                else:
                    subtopics[0].click()
                    print("OK נבחר תת נושא ראשון (ברירת מחדל)")
                    time.sleep(1)

            except Exception as e:
                print(f"ERROR שגיאה בבחירת תת נושא: {e}")
                return None

            
            # שלב 6: בחירת הסדרה הראשונה (צ'ק בוקס) - גנרי
            try:
                # מחפש את כל ה-labels שמייצגים checkboxes
                labels = self.driver.find_elements(By.CSS_SELECTOR, 'label[for^="series_"]')
                
                if not labels:
                    print("ERROR לא נמצא צ'ק בוקס")
                else:
                    first_label = labels[0]
                    # סימון הצ'קבוקס דרך JavaScript
                    checkbox_id = first_label.get_attribute('for')
                    script = f"document.getElementById('{checkbox_id}').click();"
                    self.driver.execute_script(script)
                    print(f"OK צ'ק בוקס {checkbox_id} סומן בהצלחה")

            except Exception as e:
                print(f"ERROR שגיאה בסימון צ'ק בוקס: {e}")

            
            # לחיצה על המשך לבחירת תקופת זמן
            continue_time = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-ng-click="fltCtrl.continueNextStep();"]'))
            )
            continue_time.click()
            time.sleep(2)
            
            # שלב 7: # שלב 7: בחירת השנה של החודש הקודם
            print("בוחר שנה...")
            try:
                today = date.today()
                prev_month_date = today.replace(day=1) - timedelta(days=1)  # היום האחרון של החודש הקודם
                target_year = prev_month_date.year
                print("target year:", target_year)
                
                # מציאת האלמנט של השנה
                year_link = self.wait.until(
                    EC.presence_of_element_located((
                        By.CSS_SELECTOR,
                        f"div.variableBoxInner.scroll-pane.jspScrollable div.jspPane ul li a[title='{target_year}']"
                    ))
)
                
                # גלילה לאלמנט כדי שיהיה גלוי
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", year_link)
                time.sleep(0.5)
                
                # לחיצה על השנה
                year_link.click()
                print(f"OK נבחרה שנה {target_year}")
                time.sleep(2)
            except Exception as e:
                print(f"ERROR שגיאה בבחירת שנה: {e}")
                return None
            
            # שלב 8: בחירת החודש הקודם
            prev_month = self.get_previous_month_number()
            print(f"בוחר חודש {prev_month}...")
            month_containers = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.variableBoxInner.scroll-pane.jspScrollable div.jspContainer div.jspPane'))
            )
            month_link = month_containers[1].find_element(By.CSS_SELECTOR, f"ul li a[title='{prev_month}']")
            
            # גלילה לאלמנט כדי שיהיה גלוי
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", month_link)
            time.sleep(0.5)
            
            month_link.click()
            print(f"OK נבחר חודש {prev_month}")
            time.sleep(1)
            
            # שלב 9: בחירת עד שנה
            try:
                price_index = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.variableBoxInner.scroll-pane.jspScrollable div.jspContainer div.jspPane'))
                )
                price_link = price_index[2].find_element(By.CSS_SELECTOR, f"ul li a[title='{target_year}']")
                print(f"price_link:  {price_link}")
                price_link.click()
                time.sleep(1)
                
            except Exception as e:
                print(f"ERROR שגיאה בבחירת עד שנה: {e}")
                return None

            # שלב 10 :בחירת עד חודש
            try:
                prev_month = self.get_previous_month_number()
                ad_month = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.variableBox.Narrow div.variableBoxInner.scroll-pane.jspScrollable div.jspContainer div.jspPane'))
                )
                ad_month_link = ad_month[3].find_element(By.CSS_SELECTOR, f"ul li a[title='{prev_month}']")
                ad_month_link.click()
                time.sleep(1)
                
            except Exception as e:
                print(f"ERROR שגיאה בבחירת עד חודש: {e}")
                return None

            # שלב 11 :בחירת סוג מדד
            print("בוחר סוג מדד...")
            try:
                # מציאת העמודה "סוג מדד" לפי הכותרת
                index_type_link = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//p[@class="boxTitle ng-binding"][contains(text(), "סוג מדד")]/following-sibling::div//ul//li[1]/a'))
                )
                print(f"OK נמצא סוג מדד: {index_type_link.get_attribute('title')}")
                index_type_link.click()
                print("OK נבחר סוג מדד ראשון")
                time.sleep(1)
                
            except Exception as e:
                print(f"ERROR שגיאה בבחירת סוג מדד: {e}")
                return None

            # שלב 12 :בחירת סוג בסיס
            print("בוחר סוג בסיס...")
            try:
                # מציאת העמודה "סוג בסיס" לפי הכותרת
                index_type_link = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//p[@class="boxTitle ng-binding"][contains(text(), "סוג בסיס")]/following-sibling::div//ul//li[1]/a'))
                )
                print(f"OK נמצא סוג בסיס: {index_type_link.get_attribute('title')}")
                index_type_link.click()
                print("OK נבחר סוג בסיס ראשון")
                time.sleep(1)
                
            except Exception as e:
                print(f"ERROR שגיאה בבחירת סוג בסיס: {e}")
                return None

            # שלב 13 :בחירת תקופת בסיס
            print("בוחר תקופת בסיס...")
            try:
                # ניסיון ראשון (הישן)
                try:
                    # מציאת כל האפשרויות של תקופת בסיס (הגרסה הגנרית)
                    period_options = self.wait.until(
                        EC.presence_of_all_elements_located((
                            By.CSS_SELECTOR,
                            'div.variableBoxInner.scroll-pane.jspScrollable div.jspContainer div.jspPane a.ellipsis.ng-binding'
                        ))
                    )
                    first_period = period_options[0]
                    print(f"OK נמצאה תקופת בסיס: {first_period.get_attribute('title')}")
                    first_period.click()
                    print("OK נבחרה תקופת בסיס ראשונה (ניסיון ראשון)")
                    time.sleep(1)

                except Exception as e:
                    print(f"INFO ניסיון ראשון נכשל ({e}), מנסה גרסה חלופית...")

                    # ניסיון שני (החדש עם control-id)
                    try:
                        period_options = self.wait.until(
                            EC.presence_of_all_elements_located((
                                By.CSS_SELECTOR,
                                "li[control-id=\"'basePeriods'\"] div.variableBoxInner.scroll-pane ul.scroll-pane-inner a.ellipsis.ng-binding"
                            ))
                        )
                        first_period = period_options[0]
                        print(f"OK נמצאה תקופת בסיס: {first_period.get_attribute('title')}")
                        first_period.click()
                        print("OK נבחרה תקופת בסיס ראשונה (ניסיון שני)")
                        time.sleep(1)

                    except Exception as e2:
                        print(f"ERROR שני הניסיונות לבחירת תקופת בסיס נכשלו: {e2}")
                        return None

            except Exception as outer_e:
                print(f"שגיאה כללית בתהליך בחירת תקופת בסיס: {outer_e}")
                return None



            # שלב 12: המשך לטבלת נתונים
            print("עובר לטבלת נתונים...")
            try:
                continue_table_btn = self.wait.until(
                    EC.element_to_be_clickable((By.LINK_TEXT, "המשך לטבלת הנתונים"))
                )
                continue_table_btn.click()
                print("OK עבר לטבלת נתונים")
                time.sleep(5)  # המתנה לטעינת הטבלה
            except Exception as e:
                print(f"ERROR שגיאה במעבר לטבלה: {e}")
                return None

            # שלב 13: חילוץ הערך מהטבלה
            print(f"מחלץ ערך מהטבלה לחודש {prev_month}...")
            try:
                # חיפוש הטבלה
                table = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div#grid'))
                )
                print(f"OK נמצאה טבלה")
                
                # מציאת ה-tr עם data-uid (השורה עם הערכים)
                data_row = table.find_element(By.CSS_SELECTOR, 'tr[data-uid]')
                print(f"OK נמצאה שורת נתונים")
                
                # מציאת ה-td המתאים לפי מספר החודש
                # 3 ה-td הראשונים הם כותרות, אז צריך להוסיף 3
                # חודש 1 = td 4, חודש 8 = td 11 וכן הלאה
                td_index = prev_month + 3
                value_cell = data_row.find_element(By.CSS_SELECTOR, f'td:nth-child({td_index})')
                print(f"OK נמצא td במיקום {td_index} (חודש {prev_month})")
                                
                # גלילה לתא כדי שיהיה גלוי - חשוב לעשות זאת לפני קריאת הערך!
                print(f"גולל לתא...")
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'nearest', inline: 'center'});", value_cell)
                time.sleep(2)
                print(f"OK גלילה הושלמה")
                
                # עכשיו קוראים את הערך אחרי הגלילה
                indicator_value = value_cell.text.strip()
                print(f"OK ערך המדד לחודש {prev_month}: '{indicator_value}'")
                
                if not indicator_value:
                    print("WARNING הערך ריק, מנסה עם JavaScript...")
                    indicator_value = self.driver.execute_script("return arguments[0].innerText || arguments[0].textContent;", value_cell).strip()
                    print(f"OK ערך מ-JavaScript: '{indicator_value}'")
                
                return indicator_value
                
            except Exception as e:
                print(f"ERROR שגיאה בחילוץ ערך מהטבלה: {e}")
                import traceback
                traceback.print_exc()
                return None
                
            
        except Exception as e:
            print(f"שגיאה בשליפת המדד {indicator_name}: {str(e)}")
            return None
    
    def scrape_all_cbs_indicators(self):
        """שליפת כל המדדים מאתר הלמ"ס"""
        if not self.setup_driver():
            err = self._driver_error or "לא ניתן להפעיל את Chrome"
            raise RuntimeError(
                f"לא ניתן להפעיל את דפדפן Chrome: {err}\n"
                "ודאי ש-Chrome מותקן וש-Selenium מעודכן."
            )
        
        results = {}
        bls_value = None
        
        try:
            indicators = list(self.cbs_indicators.items())
            for i, (indicator_name, indicator_code) in enumerate(indicators):
                scenario_text = self.cbs_scenarios[i] if i < len(self.cbs_scenarios) else None
                value = self.scrape_cbs_indicator(indicator_name, indicator_code, scenario_text)
                if value:
                    results[indicator_name] = value
                    
                # הפסקה קצרה בין מדדים
                time.sleep(2)
            
            # שליפת נתוני BLS
            print("\nמתחיל לשלוף נתוני BLS...")
            bls_value = self.scrape_bls_cpi()
            if bls_value:
                print(f"OK נתון BLS נשלף: {bls_value}")
            else:
                print("ERROR לא הצליח לשלוף נתון BLS")
                
        finally:
            self.close_driver()
        
        return results, bls_value
    
    def scrape_bls_cpi(self):
        """שליפת נתוני CPI מאתר BLS"""
        try:
            # בדיקה אם החלון עדיין פתוח
            try:
                self.driver.current_url
            except:
                print(f"ERROR - החלון נסגר, יוצר חלון חדש...")
                self.setup_driver()
            
            print("מתחיל לשלוף נתוני CPI מאתר BLS...")
            
            # פתיחת האתר
            self.driver.get(self.bls_url)
            time.sleep(5)  # המתנה ארוכה יותר לטעינת JavaScript
            
            # המתנה לטעינת הטבלה - ננסה כמה אפשרויות
            print("מחכה לטעינת הטבלה...")
            try:
                # נסיון 1: המתנה לטבלה רגילה
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable")))
                print("OK נמצאה טבלת dataTable")
            except:
                try:
                    # נסיון 2: המתנה לטבלה עם id
                    self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table[id*='DataTables']")))
                    print("OK נמצאה טבלת DataTables")
                except:
                    # נסיון 3: המתנה לכל טבלה
                    self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
                    print("OK נמצאה טבלה כלשהי")
            
            # חיפוש הטבלה עם הנתונים - ננסה כמה אפשרויות
            table = None
            try:
                # נסיון 1: טבלה עם class='dataTables_scrollBody'
                table = self.driver.find_element(By.CSS_SELECTOR, ".dataTables_scrollBody")
                print("OK נמצאה טבלת dataTables_scrollBody")
            except:
                try:
                    # נסיון 2: טבלה עם class='dataTable'
                    table = self.driver.find_element(By.CSS_SELECTOR, ".dataTable")
                    print("OK נמצאה טבלת dataTable")
                except:
                    try:
                        # נסיון 3: כל טבלה
                        tables = self.driver.find_elements(By.TAG_NAME, "table")
                        if tables:
                            table = tables[0]  # ניקח את הטבלה הראשונה
                            print(f"OK נמצאה טבלה ראשונה מתוך {len(tables)} טבלאות")
                    except:
                        print("ERROR לא נמצאה טבלה")
                        return None
            
            if not table:
                print("ERROR לא נמצאה טבלה מתאימה")
                return None
            
            # קבלת החודש הקודם בפורמט הנכון
            today = datetime.today()
            current_time = today.time()
            
            # חישוב החודש והשנה הקודמים (זהה ללוגיקה של get_previous_month_filename)
            if today.month == 1:
                prev_month = 12
                prev_year = today.year - 1
            else:
                prev_month = today.month - 1
                prev_year = today.year
            
            # אם היום <= 15 והשעה לפני 18:30, נחזור עוד חודש אחורה
            if today.day <= 15 and current_time < dt_time(18, 30):
                prev_month -= 1
                if prev_month == 0:
                    prev_month = 12
                    prev_year -= 1
            
            # יצירת period בפורמט MXX
            period = f"M{prev_month:02d}"
            print(f"מחפש נתונים לחודש: {period}, שנה: {prev_year}")
            
            # המתנה נוספת לטעינת הנתונים
            time.sleep(3)
            
            # חיפוש השורות - ננסה כמה דרכים
            rows = []
            try:
                # נסיון 1: חיפוש בתוך הטבלה
                rows = table.find_elements(By.TAG_NAME, "tr")
                print(f"נמצאו {len(rows)} שורות בטבלה")
            except:
                # נסיון 2: חיפוש בכל הדף
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table tr")
                print(f"נמצאו {len(rows)} שורות בדף")
            
            if not rows:
                print("ERROR לא נמצאו שורות בטבלה")
                return None
            
            # חיפוש השורה עם השנה והחודש הנכונים
            target_value = None
            print("מחפש שורה עם הנתונים הנכונים...")
            
            for i, row in enumerate(rows):
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4:  # צריך לפחות 4 עמודות
                        year_cell = cells[0].text.strip()
                        period_cell = cells[1].text.strip()
                        label_cell = cells[2].text.strip()
                        value_cell = cells[3].text.strip()
                        
                        print(f"שורה {i}: שנה={year_cell}, period={period_cell}, label={label_cell}, value={value_cell}")
                        
                    # בדיקה אם זה השורה הנכונה
                    if year_cell == str(prev_year) and period_cell == period:
                        target_value = value_cell
                        print(f"OK נמצא נתון לחודש {period}/{prev_year}: {target_value}")
                        
                        # שמירה על הערך המדויק ללא עיגול
                        print(f"OK ערך BLS מקורי: '{target_value}'")
                        
                        break
                except Exception as e:
                    print(f"שגיאה בעיבוד שורה {i}: {e}")
                    continue
            
            if target_value:
                print(f"OK נתון CPI נשלף בהצלחה: {target_value}")
                return target_value
            else:
                print(f"ERROR לא נמצא נתון לחודש {period}/{prev_year}")
                print("נסיון חיפוש חלופי...")
                
                # נסיון חלופי - חיפוש לפי טקסט
                try:
                    # חיפוש אלמנט שמכיל את התאריך
                    date_element = self.driver.find_element(By.XPATH, f"//td[contains(text(), '{prev_year}') and following-sibling::td[contains(text(), '{period}')]]")
                    value_element = date_element.find_element(By.XPATH, "./following-sibling::td[2]")
                    target_value = value_element.text.strip()
                    print(f"OK נמצא נתון בחיפוש חלופי: {target_value}")
                    
                    # שמירה על הערך המדויק ללא עיגול
                    print(f"OK ערך BLS חלופי מקורי: '{target_value}'")
                    
                    return target_value
                except:
                    print("ERROR גם החיפוש החלופי נכשל")
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
