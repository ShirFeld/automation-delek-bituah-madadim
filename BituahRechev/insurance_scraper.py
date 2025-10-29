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
            
            # עדכון קובץ par_rech.dat
            if display_callback:
                display_callback("\n📋 מעדכן קובץ par_rech.dat...")
            if update_callback:
                update_callback("מעדכן par_rech.dat...")
            
            self.update_par_rech_file(insurance_data)
            if display_callback:
                display_callback("✅ קובץ par_rech.dat עודכן בהצלחה")
            
            if update_callback:
                update_callback(f"הושלם: {results['total_success']}/37 + MDB + par_rech.dat")
            
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
            
            # תאריך עתידי - לשימוש בתוך הטבלה
            next_month = (datetime.now().replace(day=1) + timedelta(days=32)).replace(day=1)
            
            # תאריך נוכחי - לשם הקובץ
            current_month = datetime.now()
            image_path = os.path.join(save_path, f"{current_month.strftime('%m%y')}.jpg")
            
            # מחיקת קובץ קיים אם קיים
            if os.path.exists(image_path):
                os.remove(image_path)
                print(f"🗑️ מחק קובץ תמונה קיים: {image_path}")
            
            print(f"📁 שם קובץ תמונה: {current_month.strftime('%m%y')}.jpg (חודש נוכחי)")
            print(f"📅 תאריך בטבלה: {next_month.strftime('%d/%m/%Y')} (חודש עתידי)")
            
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
    
    def update_par_rech_file(self, insurance_data):
        """עדכון קובץ par_rech.dat עם נתוני ביטוח חדשים"""
        print("\n" + "="*60)
        print("🔄 מתחיל עדכון par_rech.dat")
        print("="*60)
        try:
            # נתיב קריאה - מהשרת
            par_rech_source_path = config.BITUAH_RECHEV_PARAM_SOURCE_FILE
            print(f"📂 נתיב מקור (קריאה): {par_rech_source_path}")
            
            # נתיב כתיבה - לתיקייה המקומית
            par_rech_output_path = os.path.join(config.BITUAH_RECHEV_OUTPUT_PATH, "par_rech.dat")
            print(f"📂 נתיב יעד (כתיבה): {par_rech_output_path}")
            
            if not os.path.exists(par_rech_source_path):
                print(f"❌ שגיאה: קובץ par_rech.dat לא נמצא ב-{par_rech_source_path}")
                print(f"💡 וודא שהקובץ קיים בשרת")
                return
            
            print("✅ קובץ par_rech.dat נמצא במקור")
            
            # קריאת הקובץ מהמקור
            print("📖 קורא את הקובץ מהשרת...")
            with open(par_rech_source_path, 'r', encoding='cp862') as f:
                lines = f.readlines()
            
            if not lines:
                print("❌ קובץ par_rech.dat ריק")
                return
            
            print(f"✅ נמצאו {len(lines)} שורות בקובץ")
            
            # מציאת השורה האחרונה שמתחילה ב-00012:
            last_00012_line = None
            last_00012_index = -1
            for i, line in enumerate(lines):
                if line.startswith('00012:'):
                    last_00012_line = line.rstrip('\n\r')
                    last_00012_index = i
            
            if not last_00012_line:
                print("❌ לא נמצאה שורה שמתחילה ב-00012:")
                return
            
            print(f"✅ נמצאה שורה אחרונה ב-00012 (שורה {last_00012_index + 1})")
            print(f"📄 שורה: {last_00012_line[:80]}...")
            
            # קביעת תאריך חדש (חודש הבא)
            next_month = (datetime.now().replace(day=1) + timedelta(days=32)).replace(day=1)
            new_date = f"{next_month.strftime('%y/%m')}"
            print(f"📅 תאריך חדש: {new_date}")
            
            # פיצול השורה לפי :
            parts = last_00012_line.split(':')
            
            if len(parts) < 45:
                print(f"❌ שגיאה: מבנה שורה לא תקין, יש רק {len(parts)} חלקים")
                return
            
            # פונקציה עזר לשמירת פורמט הרווחים
            def format_value_preserve_spaces(original_part, new_value):
                """שומר על הרווחים של החלק המקורי ומחליף רק את הערך"""
                # ספירת רווחים בהתחלה ובסוף
                leading_spaces = len(original_part) - len(original_part.lstrip())
                trailing_spaces = len(original_part) - len(original_part.rstrip())
                # בניית הערך החדש עם אותם רווחים
                return ' ' * leading_spaces + str(new_value) + ' ' * trailing_spaces
            
            # עדכון התאריך - שומר על הפורמט
            parts[1] = format_value_preserve_spaces(parts[1], new_date)
            
            # עדכון נתוני רכב פרטי (6 קבוצות גיל, 4 נפחי מנוע לכל אחת)
            private_data = insurance_data.get('private_car', {})
            age_groups_order = ['17-20', '21-23', '24-29', '30-39', '40-49', '50- ומעלה']
            engine_sizes_order = ['עד 1050', 'מ-1051 עד 1550', 'מ-1551 עד 2050', 'מ-2051 ומעלה']
            
            part_index = 2  # מתחיל אחרי התאריך
            print("\n💰 מעדכן נתוני רכב פרטי:")
            for age_group in age_groups_order:
                age_value = age_group.split('-')[0]  # 17, 21, 24, 30, 40, 50
                parts[part_index] = format_value_preserve_spaces(parts[part_index], age_value)
                part_index += 1
                
                print(f"   🚗 קבוצת גיל {age_group}:")
                for engine_size in engine_sizes_order:
                    price = private_data.get(age_group, {}).get(engine_size)
                    if price:
                        parts[part_index] = format_value_preserve_spaces(parts[part_index], int(price))
                        print(f"      • {engine_size}: {int(price)}")
                    part_index += 1
            
            # עדכון נתוני רכב מסחרי (5 קבוצות גיל, 2 משקלים לכל אחת)
            commercial_data = insurance_data.get('commercial_car', {})
            commercial_age_groups_order = ['17-20', '21-23', '24-39', '40-49', '50- ומעלה']
            weight_categories_order = ['עד 4000 (כולל)', 'מעל 4000']
            
            print("\n🚛 מעדכן נתוני רכב מסחרי:")
            for age_group in commercial_age_groups_order:
                age_value = age_group.split('-')[0]  # 17, 21, 24, 40, 50
                parts[part_index] = format_value_preserve_spaces(parts[part_index], age_value)
                part_index += 1
                
                print(f"   🚚 קבוצת גיל {age_group}:")
                for weight in weight_categories_order:
                    price = commercial_data.get(age_group, {}).get(weight)
                    if price:
                        parts[part_index] = format_value_preserve_spaces(parts[part_index], int(price))
                        print(f"      • {weight}: {int(price)}")
                    part_index += 1
            
            # עדכון נתוני רכב מיוחד (3 סוגים)
            special_data = insurance_data.get('special_vehicle', {})
            special_types_order = ['Nigrar', 'Handasi', 'Agricalture']
            
            print("\n🚜 מעדכן נתוני רכב מיוחד:")
            for special_type in special_types_order:
                price = special_data.get(special_type)
                if price:
                    parts[part_index] = format_value_preserve_spaces(parts[part_index], int(price))
                    print(f"   • {special_type}: {int(price)}")
                part_index += 1
            
            # בניית השורה החדשה
            new_line = ':'.join(parts)
            
            print(f"\n📝 שורה חדשה שתתווסף:")
            print(f"   {new_line[:100]}...")
            
            # וידוא שהשורה שלפני המיקום החדש מסתיימת ב-newline
            if last_00012_index >= 0 and last_00012_index < len(lines):
                if not lines[last_00012_index].endswith('\n'):
                    lines[last_00012_index] = lines[last_00012_index] + '\n'
            
            # הכנסת השורה החדשה מיד אחרי השורה האחרונה של 00012
            print("\n💾 כותב קובץ מעודכן לתיקייה המקומית...")
            print(f"   מכניס שורה חדשה במיקום {last_00012_index + 2} (אחרי השורה האחרונה של 00012)")
            print(f"   סה\"כ שורות בקובץ החדש: {len(lines) + 1}")
            
            # הכנסת השורה החדשה במיקום הנכון
            lines.insert(last_00012_index + 1, new_line + '\n')
            
            # כתיבת הקובץ המעודכן לתיקייה היעד
            with open(par_rech_output_path, 'w', encoding='cp862') as f:
                f.writelines(lines)
            
            print(f"\n✅✅✅ קובץ par_rech.dat עודכן בהצלחה! ✅✅✅")
            print(f"📂 נקרא מ: {par_rech_source_path}")
            print(f"📁 נשמר ב: {par_rech_output_path}")
            print("="*60 + "\n")
            
        except Exception as e:
            print(f"\n❌❌❌ שגיאה בעדכון par_rech.dat: {str(e)} ❌❌❌")
            import traceback
            traceback.print_exc()
            print("="*60 + "\n")
