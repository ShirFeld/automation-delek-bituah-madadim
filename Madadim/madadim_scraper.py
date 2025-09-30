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
        """קבלת מספר החודש הקודם"""
        today = datetime.date.today()
        if today.month == 1:
            return 12
        else:
            return today.month - 1
    
    def scrape_cbs_indicator(self, indicator_name, indicator_code):
        """שליפת מדד בודד מאתר הלמ"ס"""
        try:
            print(f"מתחיל לשלוף את המדד: {indicator_name} (קוד: {indicator_code})")
            
            # פתיחת האתר - בדיוק לפי הקוד שעבד
            self.driver.get("https://www.cbs.gov.il/he/Statistics/Pages/%D7%9E%D7%97%D7%95%D7%9C%D7%9C%D7%99%D7%9D/%D7%9E%D7%97%D7%95%D7%9C%D7%9C-%D7%9E%D7%97%D7%99%D7%A8%D7%99%D7%9D.aspx")
            self.driver.set_window_size(1936, 1048)  # חזרה לגודל הספציפי שעבד
            self.driver.switch_to.frame(0)
            time.sleep(3)  # המתנה ארוכה יותר
            
            # שלב 1: בחירת רדיו באטון - עם הגנות
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

            
            # שלב 4: בחירת הנושא השני ב-variableBox - גנרי
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

            
            # לחיצה על כפתור המשך הבא
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
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.variableBox ul li a'))
                )
                print(f"נמצאו {len(subtopics)} תת נושאים")
                
                if len(subtopics) == 0:
                    print("❌ לא נמצאו תת נושאים")
                    return None
                
                # בחירת התת נושא הראשון
                first_subtopic = subtopics[0]
                subtopic_text = first_subtopic.text
                print(f"בוחר תת נושא ראשון: {subtopic_text}")
                
                # גלילה ולחיצה
                self.driver.execute_script("arguments[0].scrollIntoView(true);", first_subtopic)
                time.sleep(1)
                
                try:
                    first_subtopic.click()
                    print("✓ נבחר תת נושא ראשון")
                except:
                    self.driver.execute_script("arguments[0].click();", first_subtopic)
                    print("✓ נבחר תת נושא ראשון עם JavaScript")
                
                time.sleep(3)
            except Exception as e:
                print(f"❌ שגיאה בבחירת תת נושא: {e}")
                return None
            
            # שלב 6: בחירת הסדרה הראשונה (צ'ק בוקס) - גנרי
            print("בוחר צ'ק בוקס ראשון...")
            try:
                # מחפש את כל הצ'ק בוקסים
                checkbox_selectors = [
                    'label[for^="series_"]',
                    '.checkbox > label',
                    'input[type="checkbox"]',
                    '.checkbox input'
                ]
                
                checkbox_element = None
                for selector in checkbox_selectors:
                    try:
                        checkboxes = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        if checkboxes:
                            checkbox_element = checkboxes[0]  # הראשון ברשימה
                            print(f"✓ נמצא צ'ק בוקס עם: {selector}")
                            break
                    except:
                        continue
                
                if not checkbox_element:
                    print("❌ לא נמצא צ'ק בוקס")
                    return None
                
                # גלילה ולחיצה על הצ'ק בוקס
                self.driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_element)
                time.sleep(1)
                
                try:
                    checkbox_element.click()
                    print("✓ צ'ק בוקס נבחר")
                except:
                    self.driver.execute_script("arguments[0].click();", checkbox_element)
                    print("✓ צ'ק בוקס נבחר עם JavaScript")
                
                time.sleep(2)
            except Exception as e:
                print(f"❌ שגיאה בבחירת צ'ק בוקס: {e}")
                return None
            
            # לחיצה על המשך לבחירת תקופת זמן
            continue_time = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-ng-click="fltCtrl.continueNextStep();"]'))
            )
            continue_time.click()
            time.sleep(2)
            
            # שלב 9: בחירת השנה הנוכחית
            print("בוחר שנה...")
            try:
                current_year = datetime.date.today().year
                year_link = self.wait.until(
                    EC.element_to_be_clickable((By.LINK_TEXT, str(current_year)))
                )
                year_link.click()
                print(f"✓ נבחרה שנה {current_year}")
                time.sleep(2)
            except Exception as e:
                print(f"❌ שגיאה בבחירת שנה: {e}")
                return None
            
            # שלב 6: בחירת החודש הקודם
            prev_month = self.get_previous_month_number()
            month_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, f'//a[@title="{prev_month}"]'))
            )
            month_link.click()
            time.sleep(1)
            
            # שלב 11: בחירת מדדי מחירים ואפשרויות
            print("בוחר מדדי מחירים...")
            try:
                # בחירת מדדי מחירים
                price_index = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//a[@title="מדדי מחירים"]'))
                )
                price_index.click()
                time.sleep(1)
                
                # בחירת סוג בסיס ראשון
                first_basis = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//div[@class="variableBox"]//ul/li[1]/a'))
                )
                first_basis.click()
                time.sleep(1)
                
                # בחירת תקופת בסיס ראשונה
                first_period = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//div[@class="variableBox"]//ul/li[1]/a'))
                )
                first_period.click()
                print("✓ נבחרו מדדי מחירים")
                time.sleep(2)
            except Exception as e:
                print(f"❌ שגיאה בבחירת מדדי מחירים: {e}")
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
            
            # שלב 13: חילוץ הערך מהטבלה - לפי הקוד שעבד
            # לפי הקוד שעבד - נבחר תא מסוים בטבלה
            try:
                # נבחר את התא הנכון בטבלה לפי הקוד שעבד
                selected_cell = self.driver.find_element(By.CSS_SELECTOR, ".k-state-selected:nth-child(11)")
                selected_cell.click()
                time.sleep(1)
                
                # חילוץ הערך מהתא הנבחר
                indicator_value = selected_cell.text.strip()
                print(f"ערך המדד {indicator_name}: {indicator_value}")
                return indicator_value
                
            except:
                # אם לא הצלחנו, ננסה שיטה אחרת
                print("מנסה שיטת חילוץ חלופית...")
                try:
                    # חיפוש הערך בטבלה לפי שם החודש
                    month_names = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]
                    prev_month_name = month_names[prev_month - 1]
                    
                    value_cell = self.driver.find_element(By.XPATH, f'//th[contains(text(), "{prev_month_name}")]/following-sibling::td[1]')
                    indicator_value = value_cell.text.strip()
                    print(f"ערך המדד {indicator_name} (שיטה חלופית): {indicator_value}")
                    return indicator_value
                except:
                    print("לא הצלחנו לחלץ ערך מהטבלה")
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
