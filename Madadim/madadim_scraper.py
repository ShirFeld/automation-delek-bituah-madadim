import datetime
import os
import random
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException

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
            "מדד המחירים לצרכן - כללי",
            "מדד המחירים לצרכן, לפי קבוצות צריכה ראשיות",
            "סנריו 3",
            "סנריו 4",
            # ... עד 11
        ]
        self.current_scenario_index = 0 
        
        # אתרי המקור
        self.cbs_url = "https://www.cbs.gov.il/he/Statistics/Pages/%D7%9E%D7%97%D7%95%D7%9C%D7%9C%D7%99%D7%9D/%D7%9E%D7%97%D7%95%D7%9C%D7%9C-%D7%9E%D7%97%D7%99%D7%A8%D7%99%D7%9D.aspx"
        self.bls_url = "https://data.bls.gov/dataViewer/view/timeseries/CUUR0000SA0"
        
        # נתיב יעד לקבצים
        self.target_path = r"C:\Users\shir.feldman\Desktop\parametrsUpdate\Madadim"
        
        # הגדרות selenium
        self.driver = None
        self.wait = None
    
    def get_previous_month_filename(self):
        """חישוב שם הקובץ לפי החודש הקודם"""
        today = datetime.date.today()
        
        # חישוב החודש הקודם
        if today.month == 1:
            # אם אנחנו בינואר, החודש הקודם הוא דצמבר של השנה הקודמת
            prev_month = 12
            prev_year = today.year - 1
        else:
            prev_month = today.month - 1
            prev_year = today.year
        
        # פורמט MMYY (חודש דו ספרתי + שנה דו ספרתי)
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
        """הגדרת דפדפן Chrome עם הגנות מרביות"""
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
            print("✓ Chrome driver נוצר")
            
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
            print("✓ הגנות אנטי-זיהוי מתקדמות הופעלו")
            
            # הגדרת חלון
            self.driver.set_window_size(1936, 1048)  # גודל ספציפי שעבד
            print("✓ גודל חלון הוגדר")
            
            self.wait = WebDriverWait(self.driver, 20)  # זמן המתנה ארוך יותר
            print("✓ WebDriverWait הוגדר")
            
            print("Chrome driver מוכן לשימוש בהצלחה!")
            
        except Exception as e:
            print(f"❌ שגיאה בהגדרת דפדפן: {e}")
            self.driver = None
        
    def close_driver(self):
        """סגירת הדפדפן"""
        if self.driver:
            self.driver.quit()
    
    def get_previous_month_number(self):
        month = 0
        """קבלת מספר החודש הקודם"""
        today = datetime.date.today()
        if today.month == 1:
            month = 12
        else:
            month =  today.month - 1

        if today.day < 15:
            month = month - 1
        
        return month    
        
    
    def scrape_cbs_indicator(self, indicator_name, indicator_code):
        """שליפת מדד בודד מאתר הלמ"ס"""
        try:
            print(f"מתחיל לשלוף את המדד: {indicator_name} (קוד: {indicator_code})")
            
            # פתיחת האתר - בדיוק לפי הקוד שעבד
            self.driver.get("https://www.cbs.gov.il/he/Statistics/Pages/%D7%9E%D7%97%D7%95%D7%9C%D7%9C%D7%99%D7%9D/%D7%9E%D7%97%D7%95%D7%9C%D7%9C-%D7%9E%D7%97%D7%99%D7%A8%D7%99%D7%9D.aspx")
            self.driver.set_window_size(1936, 1048)  # חזרה לגודל הספציפי שעבד
            self.driver.switch_to.frame(0)
            time.sleep(3)  # המתנה ארוכה יותר
            
            # שלב 1: סימון הכפתור לחיפוש לפי קוד
            print("בוחר רדיו באטון...")
            try:
                # נסיון עם ActionChains כמו בקוד המקורי
                from selenium.webdriver.common.action_chains import ActionChains
                radio_element = self.driver.find_element(By.NAME, "7")
                actions = ActionChains(self.driver)
                actions.move_to_element(radio_element).perform()
                time.sleep(1)
                
                try:
                    radio_element.click()
                    print("✓ רדיו נלחץ בלחיצה רגילה")
                except:
                    self.driver.execute_script("arguments[0].click();", radio_element)
                    print("✓ רדיו נלחץ עם JavaScript")
                
                time.sleep(3)  # המתנה ארוכה אחרי הרדיו
            except Exception as e:
                print(f"❌ שגיאה ברדיו: {e}")
                return None
            
            # שלב 2: הכנסת הקוד - עם הגנות מרביות וההמתנות
            print(f"מכניס קוד {indicator_code}...")
            try:
                # המתנה ארוכה יותר לפני הכנסת קוד
                time.sleep(2)
                
                # בדיקה שהדפדפן עדיין פעיל
                try:
                    current_url = self.driver.current_url
                    print(f"✓ דפדפן פעיל: {current_url[:50]}...")
                except:
                    print("❌ דפדפן נסגר!")
                    return None
                
                # חיפוש השדה הנכון עם המתנה
                print("מחפש שדה הקוד...")
                code_field = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[ng-model='mainCtrl.codesearch']"))
                )
                print("✓ שדה קוד נמצא")
                
                # ניקוי השדה בעדינות
                code_field.clear()
                time.sleep(1)
                
                # לחיצה על השדה
                actions = ActionChains(self.driver)
                actions.move_to_element(code_field).perform()
                time.sleep(0.5)
                code_field.click()
                time.sleep(1)
                
                print(f"מתחיל להכניס קוד {indicator_code} אות אחרי אות...")
                
                # הכנסת הקוד אות אחרי אות (אנושי מאוד)
                for i, char in enumerate(indicator_code):
                    code_field.send_keys(char)
                    time.sleep(0.2 + random.random() * 0.3)  # המתנה אקראית בין אותיות
                    
                    # בדיקה כל כמה אותיות שהדפדפן עדיין פעיל
                    if i % 2 == 0:
                        try:
                            self.driver.current_url
                        except:
                            print("❌ דפדפן נסגר במהלך הכנסת קוד!")
                            return None
                
                print(f"✓ קוד {indicator_code} הוכנס בהצלחה")
                time.sleep(3)  # המתנה ארוכה אחרי הכנסת הקוד
                
                # בדיקה אחרונה שהדפדפן עדיין פעיל
                try:
                    self.driver.current_url
                    print("✓ דפדפן עדיין פעיל אחרי הכנסת קוד")
                except:
                    print("❌ דפדפן נסגר אחרי הכנסת קוד!")
                    return None
                    
            except Exception as e:
                print(f"❌ שגיאה בהכנסת קוד: {e}")
                return None
            
            # שלב 3: לחיצה על כפתור המשך
            print("מנסה ללחוץ על המשך...")
            try:
                continue_btn = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.greenBigBtn[data-ng-click="mainCtrl.searchByCode();"]'))
                )
                continue_btn.click()
                print("✓ נלחץ על המשך")
                time.sleep(3)
            except Exception as e:
                print(f"❌ שגיאה בלחיצת המשך: {e}")
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
                    print("❌ לא נמצאו מספיק נושאים")
                    return None
                
                # בחירת הנושא השני (אינדקס 1)
                second_topic = topics[1]
                topic_text = second_topic.text
                print(f"בוחר את הנושא השני: {topic_text}")
                
                try:
                    second_topic.click()
                    print("✓ נבחר הנושא השני בלחיצה רגילה")
                except:
                    # אם נכשל, ננסה JavaScript
                    self.driver.execute_script("arguments[0].click();", second_topic)
                    print("✓ נבחר הנושא השני עם JavaScript")
                
                time.sleep(3)
            except Exception as e:
                print(f"❌ שגיאה בבחירת הנושא השני: {e}")
                return None

            
            # לחיצה על כפתור המשך לבחירת הנושאים
            next_arrow = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'img[src*="nextArrow"]'))
            )
            next_arrow.click()
            time.sleep(2)
            
            # שלב 5: בחירת תת נושא ראשון - גנרי
            print("בוחר תת נושא ראשון...")
            
            try:
            # מחפש את כל התת נושאים
                subtopics = self.wait.until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, 'div.jspPane ul.scroll-pane-inner li a.ellipsis.ng-binding')
                    )
                )
                print(f"נמצאו {len(subtopics)} תת נושאים")

                if len(subtopics) == 0:
                    print(" לא נמצאו תת נושאים")
                else:
                    for topic_text in self.cbs_scenarios:
                        target = next((t for t in subtopics if topic_text in t.text), None)
                        if target:
                            print(f"בוחר תת נושא: {topic_text}")
                            target.click()
                            time.sleep(1)
                        else:
                            print(f" לא נמצא תת נושא עם טקסט: {topic_text}")

            except Exception as e:
                print(f" שגיאה בבחירת תת נושא: {e}")

            
            # שלב 6: בחירת הסדרה הראשונה (צ'ק בוקס) - גנרי
            try:
                # מחפש את כל ה-labels שמייצגים checkboxes
                labels = self.driver.find_elements(By.CSS_SELECTOR, 'label[for^="series_"]')
                
                if not labels:
                    print("❌ לא נמצא צ'ק בוקס")
                else:
                    first_label = labels[0]
                    # סימון הצ'קבוקס דרך JavaScript
                    checkbox_id = first_label.get_attribute('for')
                    script = f"document.getElementById('{checkbox_id}').click();"
                    self.driver.execute_script(script)
                    print(f"✓ צ'ק בוקס {checkbox_id} סומן בהצלחה")

            except Exception as e:
                print(f"❌ שגיאה בסימון צ'ק בוקס: {e}")

            
            # לחיצה על המשך לבחירת תקופת זמן
            continue_time = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-ng-click="fltCtrl.continueNextStep();"]'))
            )
            continue_time.click()
            time.sleep(2)
            
            # שלב 7: בחירת השנה הנוכחית
            print("בוחר שנה...")
            try:
                current_year = datetime.date.today().year
                print("curreny year: ", current_year)
                
                # מציאת האלמנט של השנה
                year_link = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, f"div.variableBoxInner.scroll-pane.jspScrollable div.jspPane ul li a[title='{current_year}']"))
                )
                
                # גלילה לאלמנט כדי שיהיה גלוי
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", year_link)
                time.sleep(0.5)
                
                # לחיצה על השנה
                year_link.click()
                print(f"✓ נבחרה שנה {current_year}")
                time.sleep(2)
            except Exception as e:
                print(f"❌ שגיאה בבחירת שנה: {e}")
                return None
            
            # שלב 8: בחירת החודש הקודם
            prev_month = self.get_previous_month_number()
            print(f"בוחר חודש {prev_month}...")
            month_containers = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.variableBoxInner.scroll-pane.jspScrollable div.jspContainer div.jspPane'))
            )
            month_link = month_containers[1].find_element(By.CSS_SELECTOR, f"ul li a[title='{prev_month}']")
            month_link.click()
            print(f"✓ נבחר חודש {prev_month}")
            time.sleep(1)
            
            # שלב 9: בחירת עד שנה
            try:
                price_index = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.variableBoxInner.scroll-pane.jspScrollable div.jspContainer div.jspPane'))
                )
                price_link = price_index[2].find_element(By.CSS_SELECTOR, f"ul li a[title='{current_year}']")
                print(f"price_link !!!!!!!!!!!!!!!!!!!!!!!!!:  {price_link}")
                price_link.click()
                time.sleep(1)
                
            except Exception as e:
                print(f"❌ שגיאה בבחירת עד שנה: {e}")
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
                print(f"❌ שגיאה בבחירת עד חודש: {e}")
                return None

            # שלב 11 :בחירת סוג מדד
            print("בוחר סוג מדד...")
            try:
                # מציאת העמודה "סוג מדד" לפי הכותרת
                index_type_link = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//p[@class="boxTitle ng-binding"][contains(text(), "סוג מדד")]/following-sibling::div//ul//li[1]/a'))
                )
                print(f"✓ נמצא סוג מדד: {index_type_link.get_attribute('title')}")
                index_type_link.click()
                print("✓ נבחר סוג מדד ראשון")
                time.sleep(1)
                
            except Exception as e:
                print(f"❌ שגיאה בבחירת סוג מדד: {e}")
                return None

            # שלב 12 :בחירת סוג בסיס
            print("בוחר סוג בסיס...")
            try:
                # מציאת העמודה "סוג בסיס" לפי הכותרת
                index_type_link = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//p[@class="boxTitle ng-binding"][contains(text(), "סוג בסיס")]/following-sibling::div//ul//li[1]/a'))
                )
                print(f"✓ נמצא סוג בסיס: {index_type_link.get_attribute('title')}")
                index_type_link.click()
                print("✓ נבחר סוג בסיס ראשון")
                time.sleep(1)
                
            except Exception as e:
                print(f"❌ שגיאה בבחירת סוג בסיס: {e}")
                return None

            # שלב 13 :בחירת תקופת בסיס
            print("בוחר תקופת בסיס...")
            try:
                # מציאת כל האפשרויות של תקופת בסיס
                period_options = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.variableBoxInner.scroll-pane.jspScrollable div.jspContainer div.jspPane a.ellipsis.ng-binding'))
                )
                # בחירת האפשרות הראשונה מהרשימה
                first_period = period_options[0]
                print(f"✓ נמצאה תקופת בסיס: {first_period.get_attribute('title')}")
                first_period.click()
                print("✓ נבחרה תקופת בסיס ראשונה")
                time.sleep(1)
                
            except Exception as e:
                print(f"❌ שגיאה בבחירת תקופת בסיס: {e}")
                return None


            # שלב 12: המשך לטבלת נתונים
            print("עובר לטבלת נתונים...")
            try:
                continue_table_btn = self.wait.until(
                    EC.element_to_be_clickable((By.LINK_TEXT, "המשך לטבלת הנתונים"))
                )
                continue_table_btn.click()
                print("✓ עבר לטבלת נתונים")
                time.sleep(5)  # המתנה לטעינת הטבלה
            except Exception as e:
                print(f"❌ שגיאה במעבר לטבלה: {e}")
                return None

            # שלב 13: חילוץ הערך מהטבלה
            print(f"מחלץ ערך מהטבלה לחודש {prev_month}...")
            try:
                # חיפוש הטבלה
                table = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div#grid'))
                )
                print(f"✓ נמצאה טבלה")
                
                # מציאת ה-tr עם data-uid (השורה עם הערכים)
                data_row = table.find_element(By.CSS_SELECTOR, 'tr[data-uid]')
                print(f"✓ נמצאה שורת נתונים")
                
                # מציאת ה-td המתאים לפי מספר החודש
                # 3 ה-td הראשונים הם כותרות, אז צריך להוסיף 3
                # חודש 1 = td 4, חודש 8 = td 11 וכן הלאה
                td_index = prev_month + 3
                value_cell = data_row.find_element(By.CSS_SELECTOR, f'td:nth-child({td_index})')
                print(f"✓ נמצא td במיקום {td_index} (חודש {prev_month})")
                                
                # גלילה לתא כדי שיהיה גלוי - חשוב לעשות זאת לפני קריאת הערך!
                print(f"גולל לתא...")
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'nearest', inline: 'center'});", value_cell)
                time.sleep(2)
                print(f"✓ גלילה הושלמה")
                
                # עכשיו קוראים את הערך אחרי הגלילה
                indicator_value = value_cell.text.strip()
                print(f"✓ ערך המדד לחודש {prev_month}: '{indicator_value}'")
                
                if not indicator_value:
                    print("⚠️ הערך ריק, מנסה עם JavaScript...")
                    indicator_value = self.driver.execute_script("return arguments[0].innerText || arguments[0].textContent;", value_cell).strip()
                    print(f"✓ ערך מ-JavaScript: '{indicator_value}'")
                
                return indicator_value
                
            except Exception as e:
                print(f"❌ שגיאה בחילוץ ערך מהטבלה: {e}")
                import traceback
                traceback.print_exc()
                return None
                
            
        except Exception as e:
            print(f"שגיאה בשליפת המדד {indicator_name}: {str(e)}")
            return None
    
    def scrape_all_cbs_indicators(self):
        """שליפת כל המדדים מאתר הלמ"ס"""
        self.setup_driver()
        
        results = {}
        
        try:
            for indicator_name, indicator_code in self.cbs_indicators.items():
                value = self.scrape_cbs_indicator(indicator_name, indicator_code)
                if value:
                    results[indicator_name] = value
                    
                # הפסקה קצרה בין מדדים
                time.sleep(2)
                
        finally:
            self.close_driver()
        
        return results
    
    def update_data_file_with_values(self, cbs_values, bls_value=None):
        """עדכון קובץ הנתונים עם הערכים שנשלפו"""
        file_path = self.get_file_path()
        
        # קריאת הקובץ הקיים
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # עדכון הערכים
        for indicator_name, value in cbs_values.items():
            old_line = f"{indicator_name} ({self.cbs_indicators[indicator_name]}): "
            new_line = f"{indicator_name} ({self.cbs_indicators[indicator_name]}): {value}"
            content = content.replace(old_line, new_line)
        
        if bls_value:
            content = content.replace("Consumer Price Index (CUUR0000SA0): ", 
                                    f"Consumer Price Index (CUUR0000SA0): {bls_value}")
        
        # כתיבה חזרה לקובץ
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"הקובץ עודכן בהצלחה: {file_path}")

if __name__ == "__main__":
    scraper = MadadimScraper()
    print(f"שם הקובץ לחודש קודם: {scraper.get_previous_month_filename()}")
    
    # יצירת קובץ בסיסי
    scraper.create_data_file()
    
    # בדיקה עם מדד אחד ראשון - מחירים לצרכן
    print("\nמתחיל בדיקה עם מדד אחד...")
    scraper.setup_driver()
    
    try:
        # בדיקה עם המדד הראשון
        first_indicator = list(scraper.cbs_indicators.items())[0]  # מחירים לצרכן
        indicator_name, indicator_code = first_indicator
        
        value = scraper.scrape_cbs_indicator(indicator_name, indicator_code)
        
        if value:
            print(f"הבדיקה הצליחה! ערך המדד: {value}")
            # עדכון הקובץ עם הערך הזה
            scraper.update_data_file_with_values({indicator_name: value})
        else:
            print("הבדיקה נכשלה")
            
    finally:
        scraper.close_driver()
        
    print("בדיקה הושלמה")
