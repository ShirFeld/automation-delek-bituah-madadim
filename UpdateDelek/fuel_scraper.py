#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
מתחבר לאתר פז, 
שולף מחירי דלק (95, 98, נפט, סולר), 
יוצר קובץ KNE בגרסת Access 2000
"""

from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from datetime import datetime
import os
import config
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

class ModernFuelScraper:
    def __init__(self):
        self.root = tk.Tk()
        self.setup_modern_ui()
        self.target_fuels = [
            "בנ\"ע 95",
            "בנ\"ע סופר 98", 
            "נפט",
            "סולר-תחבורה"
        ]
        self.driver = None
        
    def setup_modern_ui(self):
        """הגדרת ממשק משתמש מודרני בסגנון Windows 11"""
        self.root.title("שליפת מחירי דלק - אתר פז")
        self.root.geometry("700x500")  # הגדלת החלון
        self.root.configure(bg='#f0f0f0')
        
        # הגדרת חלון במרכז המסך
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    
        self.colors = {
            'primary': '#FFB900',  # צהוב יפה
            'primary_hover': '#E6A500',  # צהוב כהה יותר
            'background': '#f0f0f0',
            'surface': '#ffffff',
            'text': '#323130',
            'text_secondary': '#605e5c'
        }
        
        # הגדרת פונט מודרני
        self.fonts = {
            'title': ('Segoe UI', 18, 'bold'),
            'subtitle': ('Segoe UI', 12),
            'button': ('Segoe UI', 10),
            'text': ('Segoe UI', 9)
        }
        
        self.create_header()
        self.create_main_content()
        self.create_footer()
        
    def create_header(self):
        """יצירת כותרת העליונה"""
        header_frame = tk.Frame(self.root, bg=self.colors['primary'], height=100)  # הגדלה ל-100
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # תיאור ואייקון (מימין לשמאל)
        title_frame = tk.Frame(header_frame, bg=self.colors['primary'])
        title_frame.pack(side='right', fill='both', expand=True, pady=15)
        
        title_label = tk.Label(
            title_frame,
            text="שליפת מחירי דלק",
            font=self.fonts['title'],
            bg=self.colors['primary'],
            fg='black'  # שחור על רקע צהוב
        )
        title_label.pack(anchor='e')  # מיישר לימין
        
        subtitle_label = tk.Label(
            title_frame,
            text="מאתר פז",
            font=self.fonts['subtitle'],
            bg=self.colors['primary'],
            fg='#2d2d2d'  # אפור כהה על רקע צהוב
        )
        subtitle_label.pack(anchor='e')  # מיישר לימין
        
        # אייקון
        icon_label = tk.Label(
            header_frame, 
            text="⛽", 
            font=('Segoe UI Emoji', 32),
            bg=self.colors['primary'],
            fg='black'  # שחור על רקע צהוב
        )
        icon_label.pack(side='right', padx=20, pady=15)
        
    def create_main_content(self):
        """יצירת תוכן מרכזי"""
        main_frame = tk.Frame(self.root, bg=self.colors['background'])
        main_frame.pack(fill='both', expand=True, padx=(15, 0), pady=10)  # ללא רווח מימין כלל
        
        # כרטיס מידע (ללא שטח ריק)
        info_card = tk.Frame(main_frame, bg=self.colors['surface'], relief='flat', bd=0)
        info_card.pack(fill='x', pady=(0, 10))  # פחות רווח
        
        # הוספת צל עדין (סימולציה)
        shadow_frame = tk.Frame(main_frame, bg='#d0d0d0', height=2)
        shadow_frame.place(in_=info_card, x=2, y=2, relwidth=1, relheight=1)
        info_card.lift()
        
        # נקבל את התאריך הנוכחי לתצוגה
        current_date = datetime.now().strftime("%d/%m/%Y")
        
        info_label = tk.Label(
            info_card,
            text=f"התוכנה תחלץ מחירים לתאריך ה-{current_date} עבור המוצרים הבאים\n• בנ\"ע 95\n• בנ\"ע סופר 98\n• נפט\n• סולר-תחבורה",
            font=self.fonts['text'],
            bg=self.colors['surface'],
            fg=self.colors['text'],
            justify='right',
            padx=10,  # עוד פחות רווח
            pady=8    # עוד פחות רווח
        )
        info_label.pack(fill='x')
        
        # כפתור ראשי
        self.start_button = tk.Button(
            main_frame,
            text="התחל שליפת נתונים",
            font=self.fonts['button'],
            bg=self.colors['primary'],
            fg='black',  # שחור על רקע צהוב
            relief='flat',
            bd=0,
            padx=30,
            pady=10,
            cursor='hand2',
            command=self.start_scraping
        )
        self.start_button.pack(pady=3)  # עוד פחות רווח
        
        # הוספת אפקט hover
        self.start_button.bind('<Enter>', lambda e: self.start_button.config(bg=self.colors['primary_hover']))
        self.start_button.bind('<Leave>', lambda e: self.start_button.config(bg=self.colors['primary']))
        
        # אזור טקסט לתוצאות
        result_frame = tk.Frame(main_frame, bg=self.colors['surface'])
        result_frame.pack(fill='both', expand=True, pady=(3, 0))  # עוד פחות רווח
        
        # כותרת לתוצאות
        result_title = tk.Label(
            result_frame,
            text=":תוצאות",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['surface'],
            fg=self.colors['text']
        )
        result_title.pack(anchor='e', padx=(10, 0), pady=(10, 5))  # מיישר לימין, ללא רווח מימין
        
        # טבלת תוצאות יפה
        table_frame = tk.Frame(result_frame)
        table_frame.pack(fill='both', expand=True, padx=(10, 0), pady=(0, 10))  # ללא רווח מימין
        
        # יצירת טבלה עם עמודות (סדר מימין לשמאל: מוצר, מחיר, תאריך)
        columns = ('תאריך', 'מחיר', 'מוצר')  # הפוך הסדר
        self.result_table = ttk.Treeview(table_frame, columns=columns, show='headings', height=8)
        
        # הגדרת style לטבלה (מרווח אחיד)
        style = ttk.Style()
        style.configure("Treeview", font=self.fonts['text'])
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'))
        
        # הגדרת כותרות עמודות
        self.result_table.heading('מוצר', text='מוצר', anchor='e')
        self.result_table.heading('מחיר', text='(₪) מחיר', anchor='center')  # הסוגריים משמאל
        self.result_table.heading('תאריך', text='תאריך', anchor='center')
        
        # הגדרת רוחב עמודות (3 ס"מ בין עמודות = 114 פיקסל)
        self.result_table.column('מוצר', width=170, anchor='e')
        self.result_table.column('מחיר', width=170, anchor='center')  
        self.result_table.column('תאריך', width=170, anchor='center')
        
        # סקרולבר לטבלה
        table_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.result_table.yview)
        self.result_table.config(yscrollcommand=table_scrollbar.set)
        
        self.result_table.pack(side='right', fill='both', expand=True)
        table_scrollbar.pack(side='right', fill='y')
        
    def create_footer(self):
        """יצירת כותרת תחתונה"""
        footer_frame = tk.Frame(self.root, bg=self.colors['background'], height=30)
        footer_frame.pack(fill='x', side='bottom')
        footer_frame.pack_propagate(False)
        
        self.status_label = tk.Label(
            footer_frame,
            text="מוכן לעבודה",
            font=self.fonts['text'],
            bg=self.colors['background'],
            fg=self.colors['text_secondary']
        )
        self.status_label.pack(side='right', padx=20, pady=5)  # מיישר לימין
        
    def update_status(self, message):
        """עדכון הודעת סטטוס"""
        try:
            if hasattr(self, 'status_label') and self.status_label.winfo_exists():
                self.status_label.config(text=message)
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.update()
        except:
            pass  # האלמנטים כבר לא קיימים
    
    def setup_driver(self):
        """הגדרת דפדפן Selenium"""
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')  # דפדפן בלתי נראה
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # User agent אמיתי
            chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # הסרת זיהוי webdriver
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            return True
        except Exception as e:
            print(f"שגיאה בהגדרת דפדפן: {str(e)}")
            return False
    
    def close_driver(self):
        """סגירת הדפדפן"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
    
    def scrape_self_service_price(self, month, year):
        """שליפת מחיר שירות עצמי מאתר delekulator"""
        try:
            # הגדרת דפדפן אם עדיין לא הוגדר
            driver_was_none = self.driver is None
            if driver_was_none:
                if not self.setup_driver():
                    print("שגיאה בהגדרת דפדפן לשליפת שירות עצמי")
                    return None
            
            print(f"שולף מחיר שירות עצמי לחודש {month}/{year}...")
            
            # גלישה לאתר delekulator - משתמש בURL מקובץ הקונפיג
            url = config.DELEKULATOR_URL
            self.driver.get(url)
            
            # המתנה לטעינת העמוד
            time.sleep(3)
            
            # קבלת תוכן העמוד
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # חיפוש השורה "שירות עצמי: YYYY/MM - X.XX ₪"
            # נחפש את הפורמט הספציפי
            search_pattern = f"שירות עצמי: {year}/{month:02d}"
            print(f"מחפש את הדפוס: '{search_pattern}'")
            
            # חיפוש בכל הטקסט
            text_content = soup.get_text()
            
            # חיפוש השורה המתאימה
            for line in text_content.split('\n'):
                line = line.strip()
                if search_pattern in line:
                    print(f"נמצאה שורה: {line}")
                    # פרסור: "שירות עצמי: 2025/10 – 7.29 ₪"
                    # נחפש את המספר אחרי המקף
                    import re
                    # חיפוש מספר עשרוני אחרי המקף (כולל סוגים שונים של מקפים)
                    match = re.search(r'[-–—]\s*(\d+\.\d+)', line)
                    if match:
                        price = float(match.group(1))
                        print(f"נמצא מחיר שירות עצמי: {price}")
                        return price
            
            print(f"לא נמצא מחיר שירות עצמי עבור {month}/{year}")
            return None
            
        except Exception as e:
            print(f"שגיאה בשליפת מחיר שירות עצמי: {str(e)}")
            return None
        finally:
            # אם יצרנו דפדפן חדש, נסגור אותו
            if driver_was_none and self.driver:
                self.close_driver()
        
    def start_scraping(self):
        """התחלת תהליך השליפה בחוט נפרד"""
        self.start_button.config(state='disabled', text="מעבד...")
        threading.Thread(target=self.scrape_fuel_prices, daemon=True).start()
        
    def scrape_fuel_prices(self):
        """שליפת מחירי דלק מאתר פז באמצעות Selenium"""
        try:
            self.update_status("מכין דפדפן...")
            
            # הגדרת דפדפן
            if not self.setup_driver():
                raise Exception("לא הצלחתי להגדיר דפדפן")
            
            self.update_status("מתחבר לאתר פז...")
            
            # גלישה לאתר - משתמש בURL מקובץ הקונפיג
            url = config.PAZ_URL
            self.driver.get(url)
            
            # המתנה לטעינת העמוד
            print("ממתין לטעינת העמוד...")
            time.sleep(5)  # המתנה לטעינה מלאה
            
            # בדיקה אם יש CAPTCHA
            page_source = self.driver.page_source
            if "Radware" in page_source or "captcha" in page_source.lower():
                print("זוהה CAPTCHA - ממתין עוד קצת...")
                time.sleep(10)  # המתנה נוספת
                page_source = self.driver.page_source
            
            self.update_status("מנתח נתונים...")
            
            # ניתוח HTML
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # חיפוש טבלת "דלקים בתחנות"
            fuel_data = self.extract_fuel_data(soup)
            
            print(f"נתונים שנחלצו: {len(fuel_data) if fuel_data else 0}")
            if fuel_data and len(fuel_data) > 0:
                print("נמצאו נתונים אמיתיים מהאתר - משתמש בהם")
                
                # שליפת מחיר שירות עצמי מאתר delekulator
                date_from_data = fuel_data[0]['date']
                date_parts = date_from_data.split('/')
                month = int(date_parts[1])
                year = int(date_parts[2])
                
                print("\nשולף מחיר שירות עצמי מאתר delekulator...")
                self_service_price = self.scrape_self_service_price(month, year)
                if self_service_price:
                    print(f"נמצא מחיר שירות עצמי: {self_service_price}")
                else:
                    print("לא נמצא מחיר שירות עצמי")
                
                self.save_to_text_file(fuel_data, self_service_price)  # שמירה לקובץ טקסט
                self.save_to_database(fuel_data, self_service_price)   # שמירה לבסיס נתונים
                self.display_results(fuel_data)
                self.update_status("התהליך הושלם בהצלחה")
                try:
                    messagebox.showinfo("הצלחה", f"נתונים אמיתיים נשמרו בהצלחה!\nנמצאו {len(fuel_data)} מוצרים\nנשמרו קבצים: טקסט ובסיס נתונים")
                except:
                    print("הצלחה: נתונים אמיתיים נשמרו בהצלחה!")
            else:
                print("לא נמצאו נתונים אמיתיים ")
                self.update_status("לא נמצאו נתונים אמיתיים")
                try:
                    messagebox.showwarning("אזהרה", "לא נמצאו נתונים באתר.\nהוצגו נתונים לדוגמה.\nנשמרו קבצים: טקסט ובסיס נתונים")
                except:
                    print("אזהרה: לא נמצאו נתונים באתר")
                
        except Exception as e:
            self.update_status("אירעה שגיאה")
            print(f"שגיאה: {str(e)}")
            try:
                messagebox.showerror("שגיאה", f"אירעה שגיאה:\n{str(e)}")
            except:
                print(f"שגיאה: {str(e)}")
            
        finally:
            self.close_driver()
            # בדיקה אם הכפתור עדיין קיים (לא נהרס)
            try:
                if self.start_button.winfo_exists():
                    self.start_button.config(state='normal', text="התחל שליפת נתונים")
            except:
                pass  # הכפתור כבר לא קיים
            
    def extract_fuel_data(self, soup):
        """חילוץ נתוני דלק מה-HTML"""
        fuel_data = []
        
        try:
            # חיפוש הכותרת "דלקים בתחנות"
            headers = soup.find_all(string=lambda text: text and "דלקים בתחנות" in text)
            
            if not headers:
                print("לא נמצאה כותרת 'דלקים בתחנות'")
                return fuel_data
                
            print("נמצאה כותרת 'דלקים בתחנות'")
            
            # חיפוש טבלה אחרי הכותרת
            for header in headers:
                parent = header.parent
                while parent and parent.name != 'body':
                    # חיפוש טבלה
                    table = parent.find_next('table')
                    if table:
                        print("נמצאה טבלה")
                        fuel_data = self.parse_table(table)
                        if fuel_data:
                            break
                    parent = parent.parent
                    
                if fuel_data:
                    break
                    
        except Exception as e:
            print(f"שגיאה בחילוץ נתונים: {str(e)}")
            
        return fuel_data
        
    def parse_table(self, table):
        """ניתוח טבלת מחירים"""
        fuel_data = []
        
        try:
            rows = table.find_all('tr')
            
            # חיפוש כותרות לזיהוי עמודות
            header_row = None
            date_col_index = -1
            price_col_index = -1
            fuel_col_index = -1
            
            for i, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                if len(cells) >= 3:
                    # הדפסת כותרות לדיבוג
                    print(f"שורה {i}: {[self.clean_text(cell.get_text()) for cell in cells]}")
                    
                    # בדיקה אם זה שורת כותרות
                    for j, cell in enumerate(cells):
                        text = self.clean_text(cell.get_text()).lower()
                        if 'תקף מ' in text or 'תאריך' in text:
                            date_col_index = j
                            header_row = i
                        elif 'כולל מע' in text or 'כולל מעמ' in text:
                            # רק אם זה באמת עמודה "כולל" ולא "לא כולל"
                            if 'לא כולל' not in text and 'לא' not in text:
                                price_col_index = j
                                print(f"נמצאה עמודת מחיר כולל מע\"מ באינדקס: {j}, טקסט: '{text}'")
                        elif 'מוצר' in text or j == 0:
                            fuel_col_index = j
                    
                    if header_row is not None:
                        break
            
            # אם לא מצאנו כותרות, נשתמש בהנחות ברירת מחדל
            if price_col_index == -1:  # לא מצאנו עמודת מחיר
                fuel_col_index = 0      # עמודה 1: מוצר
                price_col_index = 1     # עמודה 2: מחיר כולל מע"מ  
                date_col_index = 3      # עמודה 4: תאריך
                header_row = 0
                print("משתמש בהנחות ברירת מחדל - עמודת מחיר: 1 (כולל מע\"מ)")
            
            print(f"עמודות שנמצאו: מוצר={fuel_col_index}, מחיר={price_col_index}, תאריך={date_col_index}")
            
            # עיבוד שורות הנתונים
            for i, row in enumerate(rows[header_row + 1:], start=header_row + 1):
                cells = row.find_all(['td', 'th'])
                
                if len(cells) > max(fuel_col_index, price_col_index, date_col_index):
                    fuel_type = self.clean_text(cells[fuel_col_index].get_text())
                    price_text = self.clean_text(cells[price_col_index].get_text())
                    date_text = self.clean_text(cells[date_col_index].get_text()) if date_col_index < len(cells) else ""
                    
                    print(f"שורה {i}: מוצר='{fuel_type}', מחיר='{price_text}', תאריך='{date_text}'")
                    
                    # בדיקה אם זה מוצר רצוי
                    if self.is_target_fuel(fuel_type):
                        try:
                            # ניקוי מחיר מסמלים
                            price_clean = price_text.replace('₪', '').replace(',', '').strip()
                            price = float(price_clean)
                            
                            # אם יש תאריך תקף, נשתמש בו, אחרת התאריך הנוכחי
                            if date_text and self.is_valid_date(date_text):
                                valid_date = date_text
                            else:
                                valid_date = datetime.now().strftime("%d/%m/%Y")
                            
                            fuel_data.append({
                                'fuel_type': fuel_type,
                                'price': price,
                                'date': valid_date
                            })
                            
                            print(f"נוסף: {fuel_type} - {price} - {valid_date}")
                            
                        except ValueError as e:
                            print(f"שגיאה בפרסור מחיר '{price_text}': {e}")
                            pass
                            
        except Exception as e:
            print(f"שגיאה כללית בפרסור טבלה: {e}")
            
        return fuel_data
        
    def clean_text(self, text):
        """ניקוי טקסט מתווים מיותרים"""
        if not text:
            return ""
        # הסרת כל סוגי הגרשיים והגרשיים הכפולים
        cleaned = text.strip()
        cleaned = cleaned.replace('"', '').replace("'", "")
        cleaned = cleaned.replace('\u05F4', '').replace('\u05F3', '')  # גרשיים עבריים
        cleaned = cleaned.replace('\u201C', '').replace('\u201D', '')  # גרשיים כפולים
        cleaned = cleaned.replace('\u2018', '').replace('\u2019', '')  # גרשיים בודדים
        cleaned = cleaned.replace('״', '').replace('׳', '')  # עוד גרשיים עבריים
        return cleaned
        
    def is_target_fuel(self, fuel_type):
        """בדיקה אם זה סוג דלק רצוי"""
        if not fuel_type:
            return False
            
        fuel_normalized = self.clean_text(fuel_type).replace(" ", "").replace("-", "").lower()
        
        # בדיקות ספציפיות לכל סוג דלק
        # נבדוק אם המילים המרכזיות מופיעות בטקסט
        
        # בנזין 95
        if ('95' in fuel_normalized and 
            ('בנ' in fuel_normalized or 'בנזין' in fuel_normalized) and
            'סופר' not in fuel_normalized):
            print(f"זוהה דלק: '{fuel_type}' -> בנזין 95")
            return True
        
        # בנזין סופר 98
        if ('98' in fuel_normalized or 'סופר' in fuel_normalized) and \
        ('בנ' in fuel_normalized or 'בנזין' in fuel_normalized or 'סופר' in fuel_normalized):
            print(f"זוהה דלק: '{fuel_type}' -> בנזין סופר 98")
            return True
        
        # נפט
        if 'נפט' in fuel_normalized:
            print(f"זוהה דלק: '{fuel_type}' -> נפט")
            return True
        
        # סולר תחבורה
        if 'סולר' in fuel_normalized and ('תחבורה' in fuel_normalized or fuel_normalized == 'סולר'):
            print(f"זוהה דלק: '{fuel_type}' -> סולר-תחבורה")
            return True
        
        print(f"לא זוהה: '{fuel_type}' (נורמליזציה: '{fuel_normalized}')")
        return False
        
    def is_valid_date(self, date_text):
        """בדיקה אם התאריך תקין"""
        if not date_text or len(date_text) != 10:
            return False
        return '/' in date_text and date_text.count('/') == 2
    
    def save_to_text_file(self, fuel_data, self_service_price=None):
        """שמירת נתונים לקובץ טקסט"""
        try:
            if not fuel_data:
                return
                
            # קבלת התאריך מהנתונים
            date_from_data = fuel_data[0]['date']  # התאריך מהנתונים
            # המרת התאריך לפורמט שם קובץ (dd-mm-yyyy)
            date_for_filename = date_from_data.replace('/', '-')
            
            # יצירת הנתיב המלא מקובץ הקונפיג
            base_path = config.DELEK_OUTPUT_PATH
            
            # יצירת התיקייה אם היא לא קיימת
            if not os.path.exists(base_path):
                os.makedirs(base_path)
            
            # שם הקובץ עם התאריך
            filename = f"{date_for_filename}.txt"
            full_path = os.path.join(base_path, filename)
            
            # מחיקת קובץ קיים אם יש
            if os.path.exists(full_path):
                os.remove(full_path)
                print(f"מחק קובץ טקסט קיים: {full_path}")
            
            # כתיבת הנתונים לקובץ
            with open(full_path, 'w', encoding='utf-8') as f:
                f.write(f"מחירי דלק לתאריך: {date_from_data}\n")
                f.write("=" * 40 + "\n\n")
                
                for item in fuel_data:
                    f.write(f"מוצר: {item['fuel_type']}\n")
                    f.write(f"מחיר: {item['price']:.2f} ₪\n")
                    f.write(f"תאריך: {item['date']}\n")
                    f.write("-" * 30 + "\n")
                
                # הוספת מחיר שירות עצמי אם זמין
                if self_service_price:
                    f.write(f"\nמוצר: בנזין 95 - שירות עצמי\n")
                    f.write(f"מחיר: {self_service_price:.2f} ₪\n")
                    f.write(f"תאריך: {date_from_data}\n")
                    f.write(f"מקור: delekulator.co.il\n")
                    f.write("-" * 30 + "\n")
                
                f.write(f"\nקובץ נוצר: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            
            print(f"קובץ טקסט נשמר: {full_path}")
            
        except Exception as e:
            print(f"שגיאה בשמירת קובץ טקסט: {str(e)}")
    
    def save_to_database(self, fuel_data, self_service_price=None):
        """יצירת בסיס נתונים SQLite ושמירת נתונים"""
        try:
            if not fuel_data:
                return
                
            # נתיב ושם קובץ בסיס הנתונים מקובץ הקונפיג
            base_path = config.DELEK_OUTPUT_PATH
            
            # יצירת שם קובץ עם התאריך (חודש ושנה)
            date_from_data = fuel_data[0]['date']  # פורמט: dd/mm/yyyy
            date_parts = date_from_data.split('/')
            month_year = date_parts[1] + date_parts[2][2:]  # mmyy
            db_filename = f"kne{month_year}.mdb"
            db_file = os.path.join(base_path, db_filename)
            
            # מחיקת קובץ MDB קיים אם יש
            if os.path.exists(db_file):
                os.remove(db_file)
                print(f"מחק קובץ MDB קיים: {db_file}")
            
            # יצירת התיקייה אם היא לא קיימת
            if not os.path.exists(base_path):
                os.makedirs(base_path)
            
            # יצירת מיפוי הנתונים לפי הסוגים שלנו
            data_mapping = {
                'EffectiveDate': fuel_data[0]['date'],
                'Benzin91': 0,
                'Benzin96': 0,
                'Benzin98': 0,  # יעודכן אם נמצא
                'Benzin95': 0,  # יעודכן אם נמצא
                'Soler': 0,     # יעודכן אם נמצא
                'Neft': 0,      # יעודכן אם נמצא
                'SAtzmi95': self_service_price if self_service_price else 0,  # מחיר שירות עצמי
                'SAtzmi96': 0
            }
            
            # מילוי הנתונים מהמידע שחילצנו
            for item in fuel_data:
                fuel_type = item['fuel_type']
                price = float(item['price'])
                
                print(f"מיפוי דלק: '{fuel_type}' -> מחיר: {price}")
                
                # מיפוי מדויק לפי שמות הדלקים מהאתר
                fuel_type_clean = fuel_type.strip()
                
                print(f"בודק מיפוי עבור: '{fuel_type_clean}'")
                
                if 'בנע סופר 98' in fuel_type_clean or 'בנ"ע סופר 98' in fuel_type_clean:
                    data_mapping['Benzin98'] = price
                    print(f"הוכנס ל-Benzin98: {price}")
                elif 'בנע 95' in fuel_type_clean or 'בנ"ע 95' in fuel_type_clean:
                    data_mapping['Benzin95'] = price
                    print(f"הוכנס ל-Benzin95: {price}")
                elif 'סולר-תחבורה' in fuel_type_clean or 'סולר תחבורה' in fuel_type_clean:
                    data_mapping['Soler'] = price
                    print(f"הוכנס ל-Soler: {price}")
                elif 'נפט' in fuel_type_clean:
                    data_mapping['Neft'] = price
                    print(f"הוכנס ל-Neft: {price}")
                else:
                    print(f"לא נמצא מיפוי עבור: '{fuel_type_clean}'")
            
            # הדפסת סיכום המיפוי
            print("\n=== סיכום נתונים לשמירה ===")
            print(f"תאריך: {data_mapping['EffectiveDate']}")
            print(f"Benzin91: {data_mapping['Benzin91']}")
            print(f"Benzin95: {data_mapping['Benzin95']}")
            print(f"Benzin96: {data_mapping['Benzin96']}")
            print(f"Benzin98: {data_mapping['Benzin98']}")
            print(f"Soler: {data_mapping['Soler']}")
            print(f"Neft: {data_mapping['Neft']}")
            print(f"SAtzmi95: {data_mapping['SAtzmi95']}")
            print(f"SAtzmi96: {data_mapping['SAtzmi96']}")
            print("=" * 40)
            
            # יצירת קובץ Access 2000 אמיתי
            if HAS_WIN32COM:
                try:
                    self.create_real_access_db(data_mapping, db_file)
                    print(f"נוצר קובץ Access 2000 אמיתי: {db_file}")
                except Exception as e:
                    print(f"לא הצלחתי ליצור Access 2000: {str(e)}")
                    return
            else:
                print("win32com לא זמין - לא ניתן ליצור Access 2000")
                return
            
            print(f"נתונים נשמרו בבסיס נתונים Access 2000: {db_file}")
            
        except Exception as e:
            print(f"שגיאה בשמירת בסיס נתונים: {str(e)}")
    

    
    def create_real_access_db(self, data_mapping, db_file):
        """יצירת קובץ Access 2000 באמצעות COM"""
        try:
            import pythoncom
            
            # מחיקת קובץ קיים
            if os.path.exists(db_file):
                os.remove(db_file)
            
            # אתחול COM
            pythoncom.CoInitialize()
            
            try:
                # יצירת Access application
                access_app = win32com.client.Dispatch("Access.Application")
                access_app.NewCurrentDatabase(db_file, 9)  # 9 = Access 2000
                
                # יצירת הטבלה
                create_table_sql = """
                CREATE TABLE tblMehirDelek_edit (
                    EffectiveDate TEXT(10),
                    Benzin91 DOUBLE,
                    Benzin96 DOUBLE,
                    Benzin98 DOUBLE,
                    Benzin95 DOUBLE,
                    Soler DOUBLE,
                    Neft DOUBLE,
                    SAtzmi95 DOUBLE,
                    SAtzmi96 DOUBLE
                )
                """
                
                access_app.DoCmd.RunSQL(create_table_sql)
                
                # הכנסת הנתונים
                insert_sql = f"""
                INSERT INTO tblMehirDelek_edit 
                (EffectiveDate, Benzin91, Benzin96, Benzin98, Benzin95, Soler, Neft, SAtzmi95, SAtzmi96)
                VALUES ('{data_mapping['EffectiveDate']}', {data_mapping['Benzin91']}, {data_mapping['Benzin96']}, 
                        {data_mapping['Benzin98']}, {data_mapping['Benzin95']}, {data_mapping['Soler']}, 
                        {data_mapping['Neft']}, {data_mapping['SAtzmi95']}, {data_mapping['SAtzmi96']})
                """
                
                access_app.DoCmd.RunSQL(insert_sql)
                
                # שמירה וסגירה
                access_app.DoCmd.Save()
                access_app.CloseCurrentDatabase()
                access_app.Quit()
                
            finally:
                # ניקוי COM
                pythoncom.CoUninitialize()
            
        except Exception as e:
            raise Exception(f"שגיאה ביצירת Access: {str(e)}")
            
    def display_results(self, fuel_data):
        """הצגת תוצאות בטבלה יפה"""
        try:
            # בדיקה אם הטבלה עדיין קיימת
            if hasattr(self, 'result_table') and self.result_table.winfo_exists():
                # נקה את הטבלה
                for item in self.result_table.get_children():
                    self.result_table.delete(item)
                
                # הוסף את התוצאות לטבלה (בסדר הנכון: תאריך, מחיר, מוצר)
                for item in fuel_data:
                    self.result_table.insert('', 'end', values=(
                        item['date'],
                        f"{item['price']:.2f}",
                        item['fuel_type']
                    ))
        except:
            pass  # הטבלה כבר לא קיימת
            
    def run(self):
        """הפעלת האפליקציה"""
        self.root.mainloop()

def main():
    """פונקציה ראשית"""
    app = ModernFuelScraper()
    app.run()

if __name__ == "__main__":
    main()
