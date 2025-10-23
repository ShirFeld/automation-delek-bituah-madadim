#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
הקלאס הראשי לשליפת נתוני ביטוח חובה
שולף נתונים עבור: רכב פרטי,מסחרי ומיוחד
יוצר תמונה עם טבלאות ונתונים ויוצר MDB
"""

import os
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config

class InsuranceScraper:
    def __init__(self):
        self.driver = None
        self._setup_driver()

    def _setup_driver(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1600,1000')
            self.driver = webdriver.Chrome(options=options)
            self.driver.implicitly_wait(10)
        except Exception:
            self.driver = None

    def cleanup(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass

    # public API expected by main_app
    def scrape_all_age_groups_complete(self):
        return self._scrape_private()

    def scrape_commercial_vehicle_complete(self):
        return self._scrape_commercial()

    def scrape_special_vehicle_data(self):
        return self._scrape_special()
    
    def scrape_all_insurance_data(self, update_callback=None, display_callback=None):
        """
        פונקציה מקיפה שמבצעת את כל תהליך השליפה והשמירה
        דומה ל-scrape_fuel_prices ב-fuel_scraper
        """
        results = {
            'private_success': 0,
            'commercial_success': 0,
            'special_success': 0,
            'total_success': 0,
            'image_path': None,
            'mdb_path': None
        }
        
        try:
            if update_callback:
                update_callback("מתחיל שליפה מלאה...")
            
            if display_callback:
                display_callback("🚀 שליפה מלאה - כל התרחישים!")
                display_callback("🚗 רכב פרטי: 24 תרחישים")
                display_callback("🚛 רכב מסחרי: 10 תרחישים")
                display_callback("🚗 רכב מיוחד: 3 תרחישים")
                display_callback("🎯 סך הכל: 37 תרחישים\n")
            
            if not self.driver:
                if display_callback:
                    display_callback("❌ שגיאה בדפדפן")
                return results
            
            if display_callback:
                display_callback("✅ דפדפן מוכן")
            
            # רכב פרטי
            if display_callback:
                display_callback("\n🚗 מתחיל רכב פרטי...")
            if update_callback:
                update_callback("שליפת רכב פרטי...")
            
            private_results = self.scrape_all_age_groups_complete()
            if private_results:
                results['private_success'] = sum(len([p for p in group.values() if p]) for group in private_results.values() if group)
            if display_callback:
                display_callback(f"✅ רכב פרטי: {results['private_success']}/24")
            
            # רכב מסחרי
            if display_callback:
                display_callback("\n🚛 מתחיל רכב מסחרי...")
            if update_callback:
                update_callback("שליפת רכב מסחרי...")
            
            commercial_results = self.scrape_commercial_vehicle_complete()
            if commercial_results:
                results['commercial_success'] = sum(sum(1 for price in group.values() if price) for group in commercial_results.values() if group)
            if display_callback:
                display_callback(f"✅ רכב מסחרי: {results['commercial_success']}/10")
            
            # רכב מיוחד
            if display_callback:
                display_callback("\n🚗 מתחיל רכב מיוחד...")
            if update_callback:
                update_callback("שליפת רכב מיוחד...")
            
            special_results = self.scrape_special_vehicle_data()
            if special_results:
                results['special_success'] = sum(1 for price in special_results.values() if price)
            if display_callback:
                display_callback(f"✅ רכב מיוחד: {results['special_success']}/3")
            
            results['total_success'] = results['private_success'] + results['commercial_success'] + results['special_success']
            
            if display_callback:
                display_callback(f"\n🏆 סיכום: {results['total_success']}/37 תרחישים")
            
            # איחוד נתונים
            insurance_data = {
                'private_car': private_results,
                'commercial_car': commercial_results,
                'special_vehicle': special_results
            }
            
            # שמירת תמונה
            if display_callback:
                display_callback("📊 יוצר טבלאות...")
            results['image_path'] = self.save_tables_as_image(insurance_data)
            if results['image_path'] and display_callback:
                display_callback(f"📷 טבלאות נשמרו: {results['image_path']}")
            
            # שמירת MDB
            if display_callback:
                display_callback("\n📊 יוצר קובץ MDB...")
            if update_callback:
                update_callback("יוצר קובץ MDB...")
            
            results['mdb_path'] = self.create_mdb_database(insurance_data)
            if results['mdb_path'] and display_callback:
                display_callback(f"✅ קובץ MDB נוצר: {results['mdb_path']}")
                display_callback("📋 הקובץ כולל 3 טבלאות:")
                display_callback("• tblBituachHova_edit (1 שורה)")
                display_callback("• tblBituachHovaMishari_edit (5 שורות)")
                display_callback("• tblBituachHovaPrati_edit (6 שורות)")
            elif display_callback:
                display_callback("⚠️ יצירת MDB נכשלה")
            
            if update_callback:
                update_callback(f"הושלם: {results['total_success']}/37 + MDB")
            
        except Exception as e:
            print(f"שגיאה בשליפה מקיפה: {e}")
            if display_callback:
                display_callback(f"❌ שגיאה: {str(e)}")
        
        return results

    # navigation helpers
    def _goto(self):
            # משתמש ב-URL מקובץ הקונפיג
            url = config.MINISTRY_OF_TRANSPORT_URL
            self.driver.get(url)
            time.sleep(2)

    def _press_compare(self):
        self.driver.find_element(By.ID, 'press_to_compare').click()
        time.sleep(12)  # Increased wait time for commercial vehicles

    def _extract_harel_price(self):
        # חיפוש הראל עם מספר אפשרויות
        harel_selectors = [
            "//td[contains(normalize-space(.), 'הראל')]",
            "//td[contains(text(), 'הראל חברה לביטוח')]",
            "//td[contains(text(), 'הראל')]"
        ]
        
        for selector in harel_selectors:
            cells = self.driver.find_elements(By.XPATH, selector)
            for cell in cells:
                try:
                    row = cell.find_element(By.XPATH, './ancestor::tr')
                    print(f"🔍 נמצאה שורת הראל: {row.text}")
                    
                    # חיפוש מחיר בשורה
                    for td in row.find_elements(By.TAG_NAME, 'td'):
                        txt = td.text.strip().replace('₪', '').replace(',', '').replace(' ', '')
                        if txt and txt.replace('.', '').isdigit():
                            price = float(txt)
                            print(f"💰 מצא מחיר הראל: {price} ₪")
                            return price
                except Exception as e:
                    print(f"⚠️ שגיאה בעיבוד שורת הראל: {e}")
                    continue
            
            print("❌ לא נמצא מחיר הראל")
            return None
                
    def _fill_common(self, age, lic):
        try:
            # מילוי גיל - שימוש בכמה שיטות
            age_field = self.driver.find_element(By.ID,'D2')
            age_field.clear()
            age_field.click()
            age_field.send_keys(str(age))
            
            # מילוי רישיון - שימוש בכמה שיטות  
            lic_field = self.driver.find_element(By.ID,'E')
            lic_field.clear()
            lic_field.click()
            lic_field.send_keys(str(lic))
            
            # לחיצה על כפתור רדיו
            radio_button = self.driver.find_element(By.XPATH, "//input[@type='radio' and @value='1']")
            self.driver.execute_script("arguments[0].click();", radio_button)
            
            time.sleep(1)  # המתנה קצרה לאחר מילוי
            print(f"✅ מילא שדות: גיל={age}, רישיון={lic}")
        except Exception as e:
            print(f"⚠️ שגיאה במילוי שדות: {e}")

    # private cars
    def _scrape_private(self):
        self._goto(); Select(self.driver.find_element(By.ID, 'ddlSheets')).select_by_value('1'); time.sleep(1)
        scenarios = [
            (19, 900, 2), (19, 1200, 2), (19, 1800, 2), (19, 2200, 2),
            (22, 900, 4), (22, 1200, 4), (22, 1800, 4), (22, 2200, 4),
            (25, 900, 7), (25, 1200, 7), (25, 1800, 7), (25, 2200, 7),
            (31, 900, 13), (31, 1200, 13), (31, 1800, 13), (31, 2200, 13),
            (41, 900, 23), (41, 1200, 23), (41, 1800, 23), (41, 2200, 23),
            (51, 900, 33), (51, 1200, 33), (51, 1800, 33), (51, 2200, 33),
        ]
        res = {}
        for age, vol, lic in scenarios:
            try:
                self._fill_common(age, lic)
                vol_el = self.driver.find_element(By.ID, 'A')
                if vol_el.tag_name == 'select':
                    Select(vol_el).select_by_value(self._private_val(vol))
                else:
                    vol_el.clear()
                    vol_el.click()
                    vol_el.send_keys(str(vol))
                
                time.sleep(1)  # המתנה אחרי מילוי נפח
                self._press_compare()
                price = self._extract_harel_price()
                res.setdefault(self._private_age_group(age), {})[self._private_col(vol)] = price
                self._goto(); Select(self.driver.find_element(By.ID,'ddlSheets')).select_by_value('1'); time.sleep(1)
            except Exception:
                        continue
        return res

    def _private_val(self, vol):
        if vol <= 1050: return '1'
        if vol <= 1550: return '2'
        if vol <= 2050: return '3'
        return '4'

    def _private_col(self, vol):
        if vol <= 1050: return 'עד 1050'
        if vol <= 1550: return 'מ-1051 עד 1550'
        if vol <= 2050: return 'מ-1551 עד 2050'
        return 'מ-2051 ומעלה'

    def _private_age_group(self, age):
        if age <= 20: return '17-20'
        if age <= 23: return '21-23'
        if age <= 29: return '24-29'
        if age <= 39: return '30-39'
        if age <= 49: return '40-49'
        return '50- ומעלה'

    # commercial
    def _scrape_commercial(self):
        self._goto(); Select(self.driver.find_element(By.ID,'ddlSheets')).select_by_value('5'); time.sleep(1)
        scenarios = [
            (19, 4, 1), (19, 4.5, 1), (23, 4, 5), (23, 4.5, 5),
            (25, 4, 7), (25, 4.5, 7), (43, 4, 17), (43, 4.5, 17), (52, 4, 26), (52, 4.5, 26)
        ]
        res = {}
        for age, weight, lic in scenarios:
            try:
                self._fill_common(age, lic)
                w = self.driver.find_element(By.ID,'A')
                if w.tag_name == 'select':
                    Select(w).select_by_value('1' if weight == 4 else '2')
                else:
                    w.clear(); w.send_keys(str(weight))
                self._press_compare()
                price = self._extract_harel_price()
                group = self._commercial_group(age)
                col = 'עד 4000 (כולל)' if weight == 4 else 'מעל 4000'
                res.setdefault(group, {})[col] = price
                self._goto(); Select(self.driver.find_element(By.ID,'ddlSheets')).select_by_value('5'); time.sleep(1)
            except Exception:
                continue
        return res

    def _commercial_group(self, age):
        if age <= 20: return '17-20'
        if age <= 23: return '21-23'
        if age <= 39: return '24-39'
        if age <= 49: return '40-49'
        return '50- ומעלה'

    # special vehicle
    def _scrape_special(self):
        self._goto(); Select(self.driver.find_element(By.ID,'ddlSheets')).select_by_value('7'); time.sleep(1)
        self._fill_common(19, 2)
        scenarios = [('Nigrar','35'), ('Handasi','4'), ('Agricalture','11')]
        res = {}
        for key, val in scenarios:
            try:
                print(f"🎯 תרחיש {key}: בוחר ערך {val}")
                Select(self.driver.find_element(By.ID,'A')).select_by_value(val)
                self._press_compare()
                price = self._extract_harel_price()
                res[key] = price
                print(f"📊 תוצאה {key}: {price}")
                # reset
                self._goto(); Select(self.driver.find_element(By.ID,'ddlSheets')).select_by_value('7'); time.sleep(1)
                self._fill_common(19, 2)
            except Exception as e:
                print(f"❌ שגיאה בתרחיש {key}: {e}")
                continue
        print(f"📋 תוצאות רכב מיוחד: {res}")
        return res

    # outputs
    def save_tables_as_image(self, insurance_data=None, save_path=None):
        try:
            from simple_mdb_creator import prepare_all_tables_data
            import matplotlib.pyplot as plt
            
            # אם לא סופק נתיב, נשתמש בנתיב מקובץ הקונפיג
            if save_path is None:
                save_path = config.BITUAH_RECHEV_OUTPUT_PATH
            
            # יצירת התיקיות אם לא קיימות
            try:
                os.makedirs(save_path, exist_ok=True)
                print(f"תיקייה מוכנה: {save_path}")
            except Exception as e:
                print(f"שגיאה ביצירת תיקייה: {e}")
                return None
            next_month = (datetime.now().replace(day=1) + timedelta(days=32)).replace(day=1)
            image_path = os.path.join(save_path, f"{next_month.strftime('%m%y')}.jpg")
            
            # מחיקת קובץ קיים אם קיים
            if os.path.exists(image_path):
                os.remove(image_path)
                print(f"🗑️ מחק קובץ תמונה קיים: {image_path}")
            
            tables = prepare_all_tables_data(next_month.strftime('%d/%m/%Y'), insurance_data or {})
            
            # בדיקה שיש נתונים לפחות בטבלה אחת
            has_data = any(len(data['rows']) > 0 for data in tables.values())
            if not has_data:
                print("⚠️ אין נתונים ליצירת תמונה - יוצר תמונה עם הודעת 'אין נתונים'")
            
            fig, axes = plt.subplots(1, 3, figsize=(20, 8))
            for ax, name in zip(axes, ['tblBituachHovaPrati_edit','tblBituachHovaMishari_edit','tblBituachHova_edit']):
                data = tables[name]
                if len(data['rows']) > 0:
                    table = ax.table(cellText=data['rows'], colLabels=data['headers'], cellLoc='center', loc='center')
                    table.auto_set_font_size(False); table.set_fontsize(9); table.scale(1, 2)
                else:
                    # אם אין נתונים, נציג הודעה
                    ax.text(0.5, 0.5, 'אין נתונים', ha='center', va='center', fontsize=14)
                ax.set_xlim(ax.get_xlim()[::-1]); ax.axis('off')
            plt.tight_layout(); plt.savefig(image_path, dpi=300, bbox_inches='tight', format='jpg'); plt.close()
            print(f"✅ תמונה נשמרה: {image_path}")
            return image_path
        except Exception as e:
            print(f"❌ שגיאה ביצירת תמונה: {e}")
            return None

    def create_mdb_database(self, insurance_data=None, save_path=None):
        try:
            from simple_mdb_creator import create_insurance_files
            
            # אם לא סופק נתיב, נשתמש בנתיב מקובץ הקונפיג
            if save_path is None:
                save_path = config.BITUAH_RECHEV_OUTPUT_PATH
            
            return create_insurance_files(save_path, insurance_data)
        except Exception as e:
            print(f"❌ שגיאה ביצירת MDB: {e}")
            return None
