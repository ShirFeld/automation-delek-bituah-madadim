#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'BituahRechev'))

from insurance_scraper import InsuranceScraper
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import sqlite3
from datetime import datetime, timedelta
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

def save_to_kne(results):
    """שמירת הנתונים לקובץ KNE"""
    try:
        # יצירת שם קובץ עם תאריך
        effective_date = (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")
        filename = f"insurance_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.mdb"
        
        print(f"💾 שומר לקובץ: {filename}")
        print(f"📅 תאריך יעיל: {effective_date}")
        
        if not HAS_WIN32COM:
            print("❌ win32com לא זמין - לא ניתן ליצור Access 2000")
            return None
        
        # יצירת Access 2000 אמיתי
        print("🔧 יוצר Access 2000 database...")
        pythoncom.CoInitialize()
        
        try:
                    # יצירת Access application
        access_app = win32com.client.Dispatch("Access.Application")
        access_app.NewCurrentDatabase(filename, 9)  # 9 = Access 2000
            print("✅ יצר Access 2000 database")
            
            # יצירת טבלה tblBituachHova_edit
            create_table_sql = """
            CREATE TABLE tblBituachHova_edit (
                EffectiveDate TEXT(10),
                Nigrar LONG,
                Handasi LONG,
                Agricalture LONG
            )
            """
            access_app.DoCmd.RunSQL(create_table_sql)
            print("✅ יצר טבלה tblBituachHova_edit")
            
            # הכנסת הנתונים (ללא נקודה עשרונית)
            nigrar_value = int(results.get('Nigrar', 0)) if results.get('Nigrar') else 0
            handasi_value = int(results.get('Handasi', 0)) if results.get('Handasi') else 0
            agricalture_value = int(results.get('Agricalture', 0)) if results.get('Agricalture') else 0
            
            insert_sql = f"""
            INSERT INTO tblBituachHova_edit (EffectiveDate, Nigrar, Handasi, Agricalture)
            VALUES ('{effective_date}', {nigrar_value}, {handasi_value}, {agricalture_value})
            """
            access_app.DoCmd.RunSQL(insert_sql)
            print("✅ הכניס נתונים לטבלה")
            
            # שמירה וסגירה
            try:
                access_app.CloseCurrentDatabase()
                access_app.Quit()
                print("✅ Access נסגר")
            except:
                print("⚠️ לא הצלחתי לסגור Access")
            
        finally:
            pythoncom.CoUninitialize()
        
        print(f"✅ נשמר בהצלחה לקובץ: {filename}")
        print(f"📊 נתונים שנשמרו:")
        print(f"   • Nigrar: {nigrar_value}")
        print(f"   • Handasi: {handasi_value}")
        print(f"   • Agricalture: {agricalture_value}")
        
        return filename
        
    except Exception as e:
        print(f"❌ שגיאה בשמירה לקובץ KNE: {str(e)}")
        return None

def test_special_vehicle():
    """בדיקה של רכב מיוחד בלבד"""
    scraper = InsuranceScraper()
    
    try:
        # הגדרת דפדפן
        if not scraper.setup_driver(visible=True):
            print("❌ לא הצלחתי להגדיר דפדפן")
            return False
        
        # מעבר לאתר
        print("🌐 עובר לאתר...")
        scraper.driver.get("https://car.cma.gov.il/Parameters/Get?next_page=2&curr_page=1&playAnimation=true&fontSize=12")
        time.sleep(5)
        
        # בחירת רכב מיוחד
        print("🚗 בוחר רכב מיוחד...")
        try:
            dropdown = scraper.driver.find_element(By.ID, "ddlSheets")
            select = Select(dropdown)
            select.select_by_value("7")  # רכב מיוחד
            print("✅ נבחר רכב מיוחד")
            time.sleep(5)
        except Exception as e:
            print(f"❌ שגיאה בבחירת רכב מיוחד: {str(e)}")
            return False
        
        # מילוי פרטים בסיסיים
        print("📝 ממלא פרטים בסיסיים...")
        
        # גיל הנהג - 19
        try:
            age_element = scraper.driver.find_element(By.XPATH, "//input[@id='D2']")
            age_element.clear()
            age_element.send_keys("19")
            print("✅ גיל הנהג: 19")
        except:
            print("⚠️ לא הצלחתי למלא גיל")
        
        # שנות רישוי - 2
        try:
            license_element = scraper.driver.find_element(By.XPATH, "//input[@id='E']")
            license_element.clear()
            license_element.send_keys("2")
            print("✅ שנות רישוי: 2")
        except:
            print("⚠️ לא הצלחתי למלא שנות רישוי")
        
        # בחירת מין הנהג - נהג
        print("👤 בוחר מין הנהג: נהג...")
        try:
            # ניסיון עם radio button
            male_radio = scraper.driver.find_element(By.XPATH, "//input[@type='radio'][@value='1']")
            male_radio.click()
            print("✅ נבחר מין הנהג: נהג (radio button)")
        except:
            try:
                # ניסיון עם checkbox
                gender_checkbox = scraper.driver.find_element(By.XPATH, "//input[@type='checkbox'][contains(@name, 'gender')]")
                if not gender_checkbox.is_selected():
                    gender_checkbox.click()
                print("✅ נבחר מין הנהג: נהג (checkbox)")
            except:
                try:
                    # ניסיון עם ID ספציפי
                    gender_element = scraper.driver.find_element(By.XPATH, "//input[@id='H']")
                    gender_element.click()
                    print("✅ נבחר מין הנהג: נהג (ID H)")
                except:
                    print("⚠️ לא הצלחתי לבחור מין הנהג")
        
        time.sleep(3)
        
        # תוצאות
        results = {}
        
        # בדיקת שלושת התרחישים
        scenarios = [
            {
                'name': 'נגררים אחרים רכינים עד 4 טון',
                'value': '35',
                'key': 'Nigrar'
            },
            {
                'name': 'ציוד הנדסי (מנועי או זחלי)',
                'value': '4',
                'key': 'Handasi'
            },
            {
                'name': 'רכב, כולל טרקטור המיועד לעבודות חקלאיות וייעור',
                'value': '11',
                'key': 'Agricalture'
            }
        ]
        
        for i, scenario in enumerate(scenarios, 1):
            print(f"\n🎯 תרחיש {i}: {scenario['name']}")
            print(f"📋 ערך: {scenario['value']}")
            print(f"🔑 מפתח: {scenario['key']}")
            
            try:
                # חיפוש הדרופדאון
                vehicle_type_selectors = [
                    "//select[@id='A']",
                    "//select[@name='parameters[5].value']"
                ]
                
                vehicle_dropdown = None
                for selector in vehicle_type_selectors:
                    try:
                        vehicle_dropdown = scraper.driver.find_element(By.XPATH, selector)
                        print(f"✅ נמצא דרופדאון עם selector: {selector}")
                        break
                    except:
                        continue
                
                if not vehicle_dropdown:
                    print("❌ לא נמצא דרופדאון")
                    continue
                
                # הדפסת כל האופציות
                select = Select(vehicle_dropdown)
                print("🔍 אופציות זמינות:")
                for j, option in enumerate(select.options):
                    option_value = option.get_attribute("value")
                    option_text = option.text
                    print(f"  {j+1}. value='{option_value}' - '{option_text}'")
                
                # בחירת הערך
                print(f"🎯 בוחר ערך: {scenario['value']}")
                
                # ניסיון עם JavaScript
                try:
                    scraper.driver.execute_script(f"arguments[0].value = '{scenario['value']}';", vehicle_dropdown)
                    scraper.driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", vehicle_dropdown)
                    print("✅ נבחר עם JavaScript")
                    time.sleep(3)
                except Exception as js_e:
                    print(f"❌ JavaScript נכשל: {str(js_e)}")
                
                # בדיקה מה נבחר
                try:
                    selected_value = select.first_selected_option.get_attribute("value")
                    selected_text = select.first_selected_option.text
                    print(f"🔍 נבחר: value='{selected_value}' - '{selected_text}'")
                    
                    if selected_value == scenario['value']:
                        print(f"✅ הצלחה! הערך הנכון נבחר")
                    else:
                        print(f"❌ הערך הלא נכון נבחר")
                        
                        # ניסיון נוסף עם select_by_value
                        try:
                            select.select_by_value(scenario['value'])
                            print("✅ ניסיון נוסף עם select_by_value")
                            time.sleep(3)
                            
                            selected_value = select.first_selected_option.get_attribute("value")
                            if selected_value == scenario['value']:
                                print(f"✅ הצלחה! הערך הנכון נבחר")
                            else:
                                print(f"❌ עדיין הערך הלא נכון")
                        except Exception as select_e:
                            print(f"❌ select_by_value נכשל: {str(select_e)}")
                            
                except Exception as check_e:
                    print(f"❌ בדיקה נכשלה: {str(check_e)}")
                
                # לחיצה על כפתור חישוב - חשוב מאוד!
                print("🔢 לוחץ על כפתור חישוב...")
                try:
                    # הדפסת כל הכפתורים בדף כדי לראות איך הכפתור נראה
                    print("🔍 מחפש כפתורים בדף...")
                    all_buttons = scraper.driver.find_elements(By.XPATH, "//input[@type='submit'] | //input[@type='button'] | //button")
                    for i, button in enumerate(all_buttons):
                        try:
                            button_text = button.get_attribute("value") or button.text
                            button_id = button.get_attribute("id") or "no-id"
                            print(f"  כפתור {i+1}: value='{button_text}' id='{button_id}'")
                        except:
                            print(f"  כפתור {i+1}: לא ניתן לקרוא")
                    
                    # חיפוש כפתור החישוב לפי ה-ID הנכון
                    calculate_selectors = [
                        "//input[@id='press_to_compare']",
                        "//input[@type='button'][@id='press_to_compare']",
                        "//input[@type='button'][contains(@title, 'השוואת תעריפים')]",
                        "//input[@type='submit'][@value='לחץ להשוואת תעריפים>>']",
                        "//input[@type='submit'][contains(@value, 'השוואת תעריפים')]",
                        "//input[@type='button'][@value='לחץ להשוואת תעריפים>>']",
                        "//button[contains(text(), 'השוואת תעריפים')]",
                        "//input[@type='submit'][contains(@value, 'תעריפים')]",
                        "//input[@type='submit'][@value='חישוב']",
                        "//input[@type='submit'][contains(@value, 'חישוב')]"
                    ]
                    
                    calculate_button = None
                    for calc_selector in calculate_selectors:
                        try:
                            calculate_button = scraper.driver.find_element(By.XPATH, calc_selector)
                            print(f"✅ נמצא כפתור חישוב עם selector: {calc_selector}")
                            break
                        except:
                            continue
                    
                    if calculate_button:
                        calculate_button.click()
                        print("✅ לחיצה על כפתור חישוב")
                        time.sleep(15)  # המתנה ארוכה יותר לטעינת התוצאות
                    else:
                        print("❌ לא נמצא כפתור חישוב")
                        
                except Exception as calc_e:
                    print(f"❌ שגיאה בלחיצה על כפתור חישוב: {str(calc_e)}")
                    # ניסיון נוסף
                    try:
                        print("🔄 ניסיון נוסף לחיצה על כפתור...")
                        time.sleep(3)
                        calculate_button = scraper.driver.find_element(By.XPATH, "//input[@id='press_to_compare']")
                        calculate_button.click()
                        print("✅ לחיצה נוספת על כפתור חישוב")
                        time.sleep(15)
                    except:
                        print("❌ גם הניסיון הנוסף נכשל")
                
                # לקיחת הערך של הראל
                print("💰 מחפש ערך של הראל...")
                try:
                    # חיפוש טבלת התוצאות
                    result_selectors = [
                        "//table[contains(@class, 'result')]",
                        "//table[contains(@id, 'result')]",
                        "//table[contains(@class, 'table')]",
                        "//div[contains(@class, 'result')]//table",
                        "//table"
                    ]
                    
                    result_table = None
                    for result_selector in result_selectors:
                        try:
                            result_table = scraper.driver.find_element(By.XPATH, result_selector)
                            print(f"✅ נמצאה טבלת תוצאות עם selector: {result_selector}")
                            break
                        except:
                            continue
                    
                    if result_table:
                        # הדפסת כל השורות בטבלה כדי לראות מה יש
                        print("🔍 הדפסת כל השורות בטבלת התוצאות:")
                        rows = result_table.find_elements(By.XPATH, ".//tr")
                        print(f"📊 נמצאו {len(rows)} שורות בטבלה")
                        
                        for row_idx, row in enumerate(rows):
                            cells = row.find_elements(By.XPATH, ".//td")
                            print(f"  שורה {row_idx+1}: {len(cells)} תאים")
                            for cell_idx, cell in enumerate(cells):
                                cell_text = cell.text.strip()
                                print(f"    תא {cell_idx+1}: '{cell_text}'")
                        
                        # חיפוש הראל כמו בקוד של רכב מסחרי
                        print("🔍 מחפש תאים עם הראל...")
                        harel_cells = scraper.driver.find_elements(By.XPATH, "//td[contains(text(), 'הראל')]")
                        
                        if not harel_cells:
                            print("❌ לא נמצא תא עם הראל")
                            # נדפיס את כל הטבלאות שנמצאות בדף
                            tables = scraper.driver.find_elements(By.TAG_NAME, "table")
                            print(f"נמצאו {len(tables)} טבלאות בדף")
                        else:
                            print(f"✅ נמצאו {len(harel_cells)} תאים עם הראל")
                            
                            # עבור כל תא הראל, נחפש את השורה ואת המחיר
                            for harel_cell in harel_cells:
                                try:
                                    # מוצא את השורה של התא הזה
                                    harel_row = harel_cell.find_element(By.XPATH, "./ancestor::tr")
                                    
                                    # מוצא את כל התאים בשורה
                                    cells = harel_row.find_elements(By.TAG_NAME, "td")
                                    print(f"🔍 בשורת הראל נמצאו {len(cells)} תאים")
                                    
                                    # מדפיס את כל התאים בשורה
                                    for i, cell in enumerate(cells):
                                        cell_text = cell.text.strip()
                                        cell_class = cell.get_attribute("class") or ""
                                        print(f"  תא {i+1}: '{cell_text}' (class: '{cell_class}')")
                                        
                                        # מחפש תא עם class="alignCenter" שמכיל מספר
                                        if "alignCenter" in cell_class and cell_text:
                                            # בודק אם זה נראה כמו מחיר
                                            if cell_text.replace(',', '').replace('.', '').isdigit():
                                                print(f"📊 מחיר מועמד: '{cell_text}'")
                                                
                                                # ניקוי המחיר
                                                price_clean = cell_text.replace('₪', '').replace(',', '').replace(' ', '').strip()
                                                
                                                try:
                                                    harel_value = float(price_clean)
                                                    print(f"✅ נמצא מחיר הראל: {harel_value} ₪")
                                                    results[scenario['key']] = harel_value
                                                    
                                                    # שמירה ב-KNE (כאן נוסיף את הקוד לשמירה)
                                                    print(f"💾 שומר ערך {scenario['key']}: {harel_value} ב-KNE")
                                                    break
                                                except ValueError:
                                                    continue
                                        
                                        # אם לא מצאנו עם alignCenter, נחפש תא שנראה כמו מחיר
                                        if scenario['key'] not in results:
                                            for i, cell in enumerate(cells):
                                                cell_text = cell.text.strip()
                                                # מחפש מספר עם פסיק (מחיר)
                                                if ',' in cell_text and len(cell_text) >= 3:
                                                    numbers_only = cell_text.replace(',', '').replace(' ', '')
                                                    if numbers_only.isdigit():
                                                        try:
                                                            harel_value = float(numbers_only)
                                                            print(f"✅ נמצא מחיר הראל: {harel_value} ₪ (בתא {i+1})")
                                                            results[scenario['key']] = harel_value
                                                            
                                                            # שמירה ב-KNE (כאן נוסיף את הקוד לשמירה)
                                                            print(f"💾 שומר ערך {scenario['key']}: {harel_value} ב-KNE")
                                                            break
                                                        except ValueError:
                                                            continue
                                                         
                                except Exception as row_e:
                                    print(f"שגיאה בעיבוד שורת הראל: {str(row_e)}")
                                    continue
                    else:
                        print("❌ לא נמצאה טבלת תוצאות")
                        
                except Exception as harel_e:
                    print(f"❌ שגיאה בחיפוש ערך הראל: {str(harel_e)}")
                
                # חזרה ל-URL המקורי (חשוב מאוד!)
                print("🔄 חוזר לעמוד החישוב...")
                try:
                    # בדיקה אם הדפדפן עדיין פעיל
                    try:
                        scraper.driver.current_url
                    except:
                        print("❌ הדפדפן נסגר, מנסה לפתוח מחדש...")
                        if not scraper.setup_driver(visible=True):
                            print("❌ לא הצלחתי לפתוח דפדפן מחדש")
                            break
                    
                    scraper.driver.get("https://car.cma.gov.il/Parameters/Get?next_page=2&curr_page=1&playAnimation=true&fontSize=12")
                    time.sleep(5)
                    
                    # בחירת רכב מיוחד שוב
                    dropdown = scraper.driver.find_element(By.ID, "ddlSheets")
                    select = Select(dropdown)
                    select.select_by_value("7")  # רכב מיוחד
                    print("✅ נבחר רכב מיוחד שוב")
                    time.sleep(5)
                    
                    # מילוי פרטים בסיסיים שוב
                    # גיל הנהג - 19
                    age_element = scraper.driver.find_element(By.XPATH, "//input[@id='D2']")
                    age_element.clear()
                    age_element.send_keys("19")
                    
                    # שנות רישוי - 2
                    license_element = scraper.driver.find_element(By.XPATH, "//input[@id='E']")
                    license_element.clear()
                    license_element.send_keys("2")
                    
                    # בחירת מין הנהג - נהג
                    male_radio = scraper.driver.find_element(By.XPATH, "//input[@type='radio'][@value='1']")
                    male_radio.click()
                    
                    print("✅ פרטים בסיסיים מולאו שוב")
                    time.sleep(3)
                    
                except Exception as back_e:
                    print(f"❌ שגיאה בחזרה לעמוד: {str(back_e)}")
                    # ניסיון לפתוח דפדפן מחדש
                    try:
                        scraper.driver.quit()
                    except:
                        pass
                    if not scraper.setup_driver(visible=True):
                        print("❌ לא הצלחתי לפתוח דפדפן מחדש")
                        break
                
                # המתנה בין תרחישים
                print("⏳ ממתין לפני התרחיש הבא...")
                time.sleep(3)
                
            except Exception as e:
                print(f"❌ שגיאה בתרחיש {i}: {str(e)}")
        
        # הדפסת התוצאות
        print(f"\n📊 תוצאות:")
        for key, value in results.items():
            print(f"  {key}: {value}")
        
        # שמירה לקובץ KNE
        if results:
            print(f"\n💾 שומר לקובץ KNE...")
            kne_file = save_to_kne(results)
            if kne_file:
                print(f"✅ הקובץ נשמר: {kne_file}")
            else:
                print("❌ נכשל בשמירה לקובץ KNE")
        else:
            print("❌ אין נתונים לשמירה")
        
        print("\n✅ בדיקה הושלמה!")
        
    except Exception as e:
        print(f"❌ שגיאה כללית: {str(e)}")
    finally:
        if scraper.driver:
            scraper.driver.quit()

if __name__ == "__main__":
    test_special_vehicle()
