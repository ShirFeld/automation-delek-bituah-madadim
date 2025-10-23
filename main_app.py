#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk
import sys
import os
import config

# ייבוא התוכנה הקיימת לדלק
from UpdateDelek.fuel_scraper import ModernFuelScraper

# ייבוא תוכנת המדדים
from Madadim.madadim_scraper import MadadimScraper

class MainApplication:
    def __init__(self):
        self.root = tk.Tk()
        self.setup_main_window()
        self.create_tabs()
        
    def setup_main_window(self):
        """הגדרת החלון הראשי"""
        self.root.title("עדכון דלק, ביטוח חובה ומדדים")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # מרכז החלון במסך
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        # צבעים ופונטים
        self.colors = {
            'primary': '#FFB900',
            'primary_hover': '#E6A500',
            'background': '#f0f0f0',
            'surface': '#ffffff',
            'text': '#323130',
            'text_secondary': '#605e5c'
        }
        
        self.fonts = {
            'title': ('Segoe UI', 20, 'bold'),
            'subtitle': ('Segoe UI', 12),
            'button': ('Segoe UI', 10),
            'text': ('Segoe UI', 9)
        }
        
        self.create_header()
        
    def create_header(self):
        """יצירת כותרת העליונה"""
        header_frame = tk.Frame(self.root, bg=self.colors['primary'], height=120)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # כותרת ראשית
        title_frame = tk.Frame(header_frame, bg=self.colors['primary'])
        title_frame.pack(side='right', fill='both', expand=True, pady=20)
        
        title_label = tk.Label(
            title_frame,
            text="עדכון דלק, ביטוח חובה ומדדים",
            font=self.fonts['title'],
            bg=self.colors['primary'],
            fg='black'
        )
        title_label.pack(anchor='e')
        
        subtitle_label = tk.Label(
            title_frame,
            text="מערכת משולבת לעדכון מחירים",
            font=self.fonts['subtitle'],
            bg=self.colors['primary'],
            fg='#2d2d2d'
        )
        subtitle_label.pack(anchor='e')
        
        # אייקון
        icon_label = tk.Label(
            header_frame, 
            text="🚗⛽", 
            font=('Segoe UI Emoji', 28),
            bg=self.colors['primary'],
            fg='black'
        )
        icon_label.pack(side='right', padx=20, pady=20)
        
    def create_tabs(self):
        """יצירת מערכת הטאבים"""
        # מסגרת לטאבים
        tab_frame = tk.Frame(self.root, bg=self.colors['background'])
        tab_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # יצירת notebook לטאבים
        self.notebook = ttk.Notebook(tab_frame)
        self.notebook.pack(fill='both', expand=True)
        
        # הגדרת סגנון לטאבים
        style = ttk.Style()
        style.configure('TNotebook.Tab', font=('Segoe UI', 11))
        
        # טאב ראשון - דלק
        self.create_fuel_tab()
        
        # טאב שני - ביטוח חובה
        self.create_insurance_tab()
        
        # טאב שלישי - מדדים
        self.create_madadim_tab()
        
    def create_fuel_tab(self):
        """יצירת טאב הדלק"""
        fuel_frame = ttk.Frame(self.notebook)
        self.notebook.add(fuel_frame, text="מחירי דלק")
        
        # יצירת instance של תוכנת הדלק בתוך הטאב
        self.fuel_app_frame = tk.Frame(fuel_frame, bg='#f0f0f0')
        self.fuel_app_frame.pack(fill='both', expand=True)
        
        # הודעה שהטאב יטען
        loading_label = tk.Label(
            self.fuel_app_frame,
            text="לחץ על 'טען תוכנת דלק' להפעיל את תוכנת שליפת המחירים",
            font=self.fonts['text'],
            bg='#f0f0f0',
            fg=self.colors['text']
        )
        loading_label.pack(pady=50)
        
        # כפתור להפעלת תוכנת הדלק
        load_fuel_button = tk.Button(
            self.fuel_app_frame,
            text="טען תוכנת דלק",
            font=self.fonts['button'],
            bg=self.colors['primary'],
            fg='black',
            relief='flat',
            bd=0,
            padx=30,
            pady=10,
            cursor='hand2',
            command=self.load_fuel_app
        )
        load_fuel_button.pack(pady=10)
        
        # הוספת אפקט hover
        load_fuel_button.bind('<Enter>', lambda e: load_fuel_button.config(bg=self.colors['primary_hover']))
        load_fuel_button.bind('<Leave>', lambda e: load_fuel_button.config(bg=self.colors['primary']))
        
    def load_fuel_app(self):
        """טעינת תוכנת הדלק בתוך הטאב"""
        # ניקוי הפריים
        for widget in self.fuel_app_frame.winfo_children():
            widget.destroy()
            
        # יצירה ישירה של ממשק הדלק בתוך הפריים
        self.create_embedded_fuel_interface()
        
    def create_embedded_fuel_interface(self):
        """יצירת ממשק הדלק המוטמע בטאב"""
        import threading
        from datetime import datetime
        
        # יצירת instance של תוכנת הדלק
        fuel_scraper = ModernFuelScraper()
        fuel_scraper.root.destroy()  # סוגר את החלון המקורי
        
        # יצירת הממשק בתוך הפריים
        # כותרת
        header_frame = tk.Frame(self.fuel_app_frame, bg=self.colors['primary'], height=80)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        title_frame = tk.Frame(header_frame, bg=self.colors['primary'])
        title_frame.pack(side='right', fill='both', expand=True, pady=15)
        
        title_label = tk.Label(
            title_frame,
            text="שליפת מחירי דלק",
            font=('Segoe UI', 16, 'bold'),
            bg=self.colors['primary'],
            fg='black'
        )
        title_label.pack(anchor='e')
        
        subtitle_label = tk.Label(
            title_frame,
            text="מאתר פז",
            font=('Segoe UI', 11),
            bg=self.colors['primary'],
            fg='#2d2d2d'
        )
        subtitle_label.pack(anchor='e')
        
        # אייקון
        icon_label = tk.Label(
            header_frame, 
            text="⛽", 
            font=('Segoe UI Emoji', 24),
            bg=self.colors['primary'],
            fg='black'
        )
        icon_label.pack(side='right', padx=15, pady=15)
        
        # תוכן מרכזי
        main_frame = tk.Frame(self.fuel_app_frame, bg=self.colors['background'])
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # כרטיס מידע
        info_card = tk.Frame(main_frame, bg=self.colors['surface'], relief='flat', bd=0)
        info_card.pack(fill='x', pady=(0, 10))
        
        current_date = datetime.now().strftime("%d/%m/%Y")
        
        info_label = tk.Label(
            info_card,
            text=f"התוכנה תחלץ מחירים לתאריך ה-{current_date} עבור המוצרים הבאים\n• בנ\"ע 95\n• בנ\"ע סופר 98\n• נפט\n• סולר-תחבורה",
            font=self.fonts['text'],
            bg=self.colors['surface'],
            fg=self.colors['text'],
            justify='right',
            padx=10,
            pady=8
        )
        info_label.pack(fill='x')
        
        # כפתור התחלה
        start_button = tk.Button(
            main_frame,
            text="התחל שליפת נתונים",
            font=self.fonts['button'],
            bg=self.colors['primary'],
            fg='black',
            relief='flat',
            bd=0,
            padx=30,
            pady=10,
            cursor='hand2'
        )
        start_button.pack(pady=5)
        
        # אפקט hover
        start_button.bind('<Enter>', lambda e: start_button.config(bg=self.colors['primary_hover']))
        start_button.bind('<Leave>', lambda e: start_button.config(bg=self.colors['primary']))
        
        # אזור תוצאות
        result_frame = tk.Frame(main_frame, bg=self.colors['surface'])
        result_frame.pack(fill='both', expand=True, pady=(5, 0))
        
        result_title = tk.Label(
            result_frame,
            text=":תוצאות",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['surface'],
            fg=self.colors['text']
        )
        result_title.pack(anchor='e', padx=(10, 0), pady=(10, 5))
        
        # טבלת תוצאות
        table_frame = tk.Frame(result_frame)
        table_frame.pack(fill='both', expand=True, padx=(10, 0), pady=(0, 10))
        
        columns = ('תאריך', 'מחיר', 'מוצר')
        result_table = ttk.Treeview(table_frame, columns=columns, show='headings', height=6)
        
        style = ttk.Style()
        style.configure("Treeview", font=self.fonts['text'])
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'))
        
        result_table.heading('מוצר', text='מוצר', anchor='e')
        result_table.heading('מחיר', text='(₪) מחיר', anchor='center')
        result_table.heading('תאריך', text='תאריך', anchor='center')
        
        result_table.column('מוצר', width=150, anchor='e')
        result_table.column('מחיר', width=120, anchor='center')  
        result_table.column('תאריך', width=120, anchor='center')
        
        table_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=result_table.yview)
        result_table.config(yscrollcommand=table_scrollbar.set)
        
        result_table.pack(side='right', fill='both', expand=True)
        table_scrollbar.pack(side='right', fill='y')
        
        # סטטוס בר
        footer_frame = tk.Frame(self.fuel_app_frame, bg=self.colors['background'], height=25)
        footer_frame.pack(fill='x', side='bottom')
        footer_frame.pack_propagate(False)
        
        status_label = tk.Label(
            footer_frame,
            text="מוכן לעבודה",
            font=self.fonts['text'],
            bg=self.colors['background'],
            fg=self.colors['text_secondary']
        )
        status_label.pack(side='right', padx=15, pady=2)
        
        # פונקציה לעדכון סטטוס
        def update_status(message):
            status_label.config(text=message)
            self.root.update()
        
        # פונקציה להצגת תוצאות
        def display_results(fuel_data):
            # ניקוי הטבלה
            for item in result_table.get_children():
                result_table.delete(item)
            
            # הוספת התוצאות
            for item in fuel_data:
                result_table.insert('', 'end', values=(
                    item['date'],
                    f"{item['price']:.2f}",
                    item['fuel_type']
                ))
        
        # פונקציה לשליפת הנתונים
        def start_scraping():
            start_button.config(state='disabled', text="מעבד...")
            
            def scrape_task():
                temp_scraper = None
                try:
                    # יצירת instance חדש של המחלץ
                    temp_scraper = ModernFuelScraper()
                    temp_scraper.root.destroy()
                    
                    # הגדרת פונקציות עדכון סטטוס מותאמות
                    temp_scraper.update_status = update_status
                    
                    # שינוי פונקציית ההצגה להציג בטבלה שלנו במקום של fuel_scraper
                    original_display = temp_scraper.display_results
                    temp_scraper.display_results = display_results
                    
                    # קריאה לפונקציה המקורית שמבצעת את כל השליפה
                    # זה מריץ את כל הלוגיקה מ-fuel_scraper.py
                    temp_scraper.scrape_fuel_prices()
                    
                    # הצגת הודעת הצלחה (רק אם לא הייתה שגיאה)
                    update_status("התהליך הושלם בהצלחה")
                    from tkinter import messagebox
                    messagebox.showinfo("הצלחה", "נתוני דלק נשמרו בהצלחה!")
                    
                except Exception as e:
                    update_status(f"שגיאה: {str(e)}")
                    print(f"שגיאה בשליפת נתונים: {str(e)}")
                    from tkinter import messagebox
                    messagebox.showerror("שגיאה", f"אירעה שגיאה:\n{str(e)}")
                    
                finally:
                    start_button.config(state='normal', text="התחל שליפת נתונים")
            
            # הרצה בחוט נפרד
            threading.Thread(target=scrape_task, daemon=True).start()
        
        # חיבור הפונקציה לכפתור
        start_button.config(command=start_scraping)
        
    def create_insurance_tab(self):
        """יצירת טאב ביטוח חובה"""
        insurance_frame = ttk.Frame(self.notebook)
        self.notebook.add(insurance_frame, text="ביטוח חובה לרכב")
        
        # יצירת instance של תוכנת הביטוח בתוך הטאב
        self.insurance_app_frame = tk.Frame(insurance_frame, bg='#f0f0f0')
        self.insurance_app_frame.pack(fill='both', expand=True)
        
        # הודעה שהטאב יטען
        loading_label = tk.Label(
            self.insurance_app_frame,
            text="לחץ על 'טען תוכנת ביטוח רכב' להפעיל את תוכנת שליפת המחירים",
            font=self.fonts['text'],
            bg='#f0f0f0',
            fg=self.colors['text']
        )
        loading_label.pack(pady=50)
        
        # כפתור להפעלת תוכנת הביטוח
        load_insurance_button = tk.Button(
            self.insurance_app_frame,
            text="טען תוכנת ביטוח רכב",
            font=self.fonts['button'],
            bg=self.colors['primary'],
            fg='black',
            relief='flat',
            bd=0,
            padx=30,
            pady=10,
            cursor='hand2',
            command=self.load_insurance_app
        )
        load_insurance_button.pack(pady=10)
        
        # הוספת אפקט hover
        load_insurance_button.bind('<Enter>', lambda e: load_insurance_button.config(bg=self.colors['primary_hover']))
        load_insurance_button.bind('<Leave>', lambda e: load_insurance_button.config(bg=self.colors['primary']))
        
    def load_insurance_app(self):
        """טעינת תוכנת הביטוח בתוך הטאב"""
        # ניקוי הפריים
        for widget in self.insurance_app_frame.winfo_children():
            widget.destroy()
            
        # יצירה ישירה של ממשק הביטוח בתוך הפריים
        self.create_embedded_insurance_interface()
        
    def create_embedded_insurance_interface(self):
        """יצירת ממשק הביטוח המוטמע בטאב"""
        import sys
        import os
        sys.path.append(os.path.join(os.path.dirname(__file__), 'BituahRechev'))
        
        from datetime import datetime
        import threading
        
        # כותרת
        header_frame = tk.Frame(self.insurance_app_frame, bg=self.colors['primary'], height=80)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        title_frame = tk.Frame(header_frame, bg=self.colors['primary'])
        title_frame.pack(side='right', fill='both', expand=True, pady=15)
        
        title_label = tk.Label(
            title_frame,
            text="שליפת מחירי ביטוח רכב",
            font=('Segoe UI', 16, 'bold'),
            bg=self.colors['primary'],
            fg='black'
        )
        title_label.pack(anchor='e')
        
        subtitle_label = tk.Label(
            title_frame,
            text="מאתר משרד התחבורה",
            font=('Segoe UI', 11),
            bg=self.colors['primary'],
            fg='#2d2d2d'
        )
        subtitle_label.pack(anchor='e')
        
        # אייקון
        icon_label = tk.Label(
            header_frame, 
            text="🚗🛡️", 
            font=('Segoe UI Emoji', 20),
            bg=self.colors['primary'],
            fg='black'
        )
        icon_label.pack(side='right', padx=15, pady=15)
        
        # תוכן מרכזי
        main_frame = tk.Frame(self.insurance_app_frame, bg=self.colors['background'])
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # כרטיס מידע
        info_card = tk.Frame(main_frame, bg=self.colors['surface'], relief='flat', bd=0)
        info_card.pack(fill='x', pady=(0, 10))
        
        info_text = """התוכנה תחלץ מחירי ביטוח חובה מאתר משרד התחבורה

        🚗 רכב פרטי (24 תרחישים):
        • כל קבוצות הגיל: 17-20, 21-23, 24-29, 30-39, 40-49, 50+
        • 4 נפחי מנוע לכל קבוצה: 900,1200,1800,2200
        
        🚛 רכב מסחרי (10 תרחישים):
        • כל קבוצות הגיל: 17-20, 21-23, 24-39, 40-49, 50+
        • 2 משקלים לכל קבוצה: עד 4000 ק"ג, מעל 4000 ק"ג
        
        🚀 שליפה מלאה (34 תרחישים):
        • כל התרחישים ברצף - דפדפן יציב אחד
        • פרטי + מסחרי יחד בתהליך אחד
        
        📊 מקור הנתונים: חברת הראל ביטוח"""
        
        info_label = tk.Label(
            info_card,
            text=info_text,
            font=self.fonts['text'],
            bg=self.colors['surface'],
            fg=self.colors['text'],
            justify='right',
            padx=10,
            pady=8
        )
        info_label.pack(fill='x')
        
        # מסגרת כפתורים
        button_frame = tk.Frame(main_frame, bg=self.colors['background'])
        button_frame.pack(pady=15)
        
        # כפתור שליפה מלאה בלבד
        combined_button = tk.Button(
            button_frame,
            text="🚀 שליפה מלאה - כל התרחישים (37 תרחישים)",
            font=('Segoe UI', 14, 'bold'),
            bg='#9C27B0',  # סגול
            fg='white',
            relief='flat',
            bd=0,
            padx=40,
            pady=15,
            cursor='hand2'
        )
        combined_button.pack(pady=20)
        
        # אפקטי hover
        combined_button.bind('<Enter>', lambda e: combined_button.config(bg='#7B1FA2'))
        combined_button.bind('<Leave>', lambda e: combined_button.config(bg='#9C27B0'))
        
        # אזור תוצאות
        result_frame = tk.Frame(main_frame, bg=self.colors['surface'])
        result_frame.pack(fill='both', expand=True, pady=(5, 0))
        
        result_title = tk.Label(
            result_frame,
            text=":תוצאות",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['surface'],
            fg=self.colors['text']
        )
        result_title.pack(anchor='e', padx=(10, 0), pady=(10, 5))
        
        # אזור טקסט לתוצאות
        text_frame = tk.Frame(result_frame)
        text_frame.pack(fill='both', expand=True, padx=(10, 0), pady=(0, 10))
        
        result_text = tk.Text(text_frame, height=8, font=self.fonts['text'], wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=result_text.yview)
        result_text.config(yscrollcommand=scrollbar.set)
        
        result_text.pack(side='right', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # סטטוס בר
        footer_frame = tk.Frame(self.insurance_app_frame, bg=self.colors['background'], height=25)
        footer_frame.pack(fill='x', side='bottom')
        footer_frame.pack_propagate(False)
        
        status_label = tk.Label(
            footer_frame,
            text="מוכן לעבודה",
            font=self.fonts['text'],
            bg=self.colors['background'],
            fg=self.colors['text_secondary']
        )
        status_label.pack(side='right', padx=15, pady=2)
        
        # פונקציה לעדכון סטטוס
        def update_status(message):
            status_label.config(text=message)
            self.root.update()
        
        # פונקציה להצגת תוצאות
        def display_results(message):
            result_text.insert(tk.END, message + "\n")
            result_text.see(tk.END)
            self.root.update()
        


        # פונקציה לשליפה משולבת עם יצירת MDB
        def start_combined_scraping():
            combined_button.config(state='disabled', text="מעבד כל התרחישים...")
            
            def scrape_task():
                scraper = None
                try:
                    import sys
                    import os
                    sys.path.append(os.path.join(os.path.dirname(__file__), 'BituahRechev'))
                    from BituahRechev.insurance_scraper import InsuranceScraper
                    
                    # יצירת scraper
                    scraper = InsuranceScraper()
                    
                    # קריאה לפונקציה המקיפה שמבצעת את כל התהליך
                    results = scraper.scrape_all_insurance_data(
                        update_callback=update_status,
                        display_callback=display_results
                    )
                    
                    # הצגת הודעת סיכום
                    if results['total_success'] > 0:
                        from tkinter import messagebox
                        msg = f"שליפה מלאה הושלמה!\n"
                        msg += f"רכב פרטי: {results['private_success']}/24\n"
                        msg += f"רכב מסחרי: {results['commercial_success']}/10\n"
                        msg += f"רכב מיוחד: {results['special_success']}/3\n"
                        msg += f"סך הכל: {results['total_success']}/37 תרחישים"
                        if results['image_path']:
                            msg += f"\n\n📷 טבלאות: {results['image_path']}"
                        if results['mdb_path']:
                            msg += f"\n📊 MDB: {results['mdb_path']}"
                        messagebox.showinfo("הצלחה", msg)
                    else:
                        from tkinter import messagebox
                        messagebox.showerror("שגיאה", "לא ניתן להגדיר דפדפן או לא נמצאו נתונים")
                
                except Exception as e:
                    display_results(f"❌ שגיאה: {str(e)}")
                    from tkinter import messagebox
                    messagebox.showerror("שגיאה", f"שגיאה: {str(e)}")
                    
                finally:
                    if scraper:
                        scraper.cleanup()
                    combined_button.config(state='normal', text="🚀 שליפה מלאה - כל התרחישים (37 תרחישים)")
            
            threading.Thread(target=scrape_task, daemon=True).start()

        # חיבור הפונקציה לכפתור
        combined_button.config(command=start_combined_scraping)
        
    def create_madadim_tab(self):
        """יצירת טאב המדדים"""
        madadim_frame = ttk.Frame(self.notebook)
        self.notebook.add(madadim_frame, text="מדדים")
        
        # יצירת instance של תוכנת המדדים בתוך הטאב
        self.madadim_app_frame = tk.Frame(madadim_frame, bg='#f0f0f0')
        self.madadim_app_frame.pack(fill='both', expand=True)
        
        # כותרת
        title_label = tk.Label(
            self.madadim_app_frame,
            text="שליפת מדדים מאתר: הלשכה המרכזית לסטטיסטיקה",
            font=self.fonts['title'],
            bg='#f0f0f0',
            fg=self.colors['text']
        )
        title_label.pack(pady=20)
        
        # תיאור
        desc_label = tk.Label(
            self.madadim_app_frame,
            text="המערכת שולפת 12 מדדים, 11 מאתר הלשכה המרכזית לסטטיסטיקה והמדד ה12 מהלשכה המרכזית לסטטיסטיקה של ארצות הברית",
            font=self.fonts['text'],
            bg='#f0f0f0',
            fg=self.colors['text_secondary'],
            justify='center'
        )
        desc_label.pack(pady=10)
        
        # מסגרת לכפתורים
        buttons_frame = tk.Frame(self.madadim_app_frame, bg='#f0f0f0')
        buttons_frame.pack(pady=30)
        
        # כפתור לשליפת כל המדדים
        fetch_all_button = tk.Button(
            buttons_frame,
            text="שלוף את כל המדדים",
            font=self.fonts['button'],
            bg=self.colors['primary'],
            fg='black',
            relief='flat',
            bd=0,
            padx=30,
            pady=10,
            cursor='hand2',
            command=self.fetch_all_madadim
        )
        fetch_all_button.pack(side='left', padx=10)
        
        # כפתור לבדיקה עם מדד אחד
        test_button = tk.Button(
            buttons_frame,
            text="בדיקה עם מדד אחד",
            font=self.fonts['button'],
            bg='#4CAF50',
            fg='white',
            relief='flat',
            bd=0,
            padx=30,
            pady=10,
            cursor='hand2',
            command=self.test_single_madad
        )
        test_button.pack(side='left', padx=10)
        
        # מסגרת לתוצאות
        self.results_frame = tk.Frame(self.madadim_app_frame, bg='#f0f0f0')
        self.results_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # הוספת אפקטי hover
        fetch_all_button.bind('<Enter>', lambda e: fetch_all_button.config(bg=self.colors['primary_hover']))
        fetch_all_button.bind('<Leave>', lambda e: fetch_all_button.config(bg=self.colors['primary']))
        
        test_button.bind('<Enter>', lambda e: test_button.config(bg='#45a049'))
        test_button.bind('<Leave>', lambda e: test_button.config(bg='#4CAF50'))
        
    def test_single_madad(self):
        """בדיקה עם מדד אחד"""
        # ניקוי תוצאות קודמות
        for widget in self.results_frame.winfo_children():
            widget.destroy()
            
        # הצגת סטטוס פשוט במקום טרמינל עמוס
        status_label = tk.Label(
            self.results_frame,
            text="מכין בדיקת מדד יחיד...",
            font=('Segoe UI', 12, 'bold'),
            bg='#f0f0f0',
            fg='#323130'
        )
        status_label.pack(pady=20)
        
        # אזור תוצאות פשוט
        results_text = tk.Text(
            self.results_frame,
            height=10,
            width=60,
            font=('Segoe UI', 10),
            wrap=tk.WORD
        )
        results_text.pack(fill='both', expand=True, pady=10)
        
        def add_log(message, color='#ffffff'):
            """הוספת הודעה פשוטה לתוצאות"""
            results_text.insert('end', f"{message}\n")
            results_text.see('end')
            self.root.update_idletasks()  # עדכון קל יותר
        
        add_log("=== מתחיל בדיקה עם מדד אחד ===", 'cyan')
        
        try:
            # שלב 1: יצירת scraper
            add_log("שלב 1: יוצר את ה-MadadimScraper...")
            try:
                scraper = MadadimScraper()
                add_log("✓ MadadimScraper נוצר בהצלחה", 'green')
            except Exception as e:
                add_log(f"❌ שגיאה ביצירת MadadimScraper: {str(e)}", 'red')
                return
            
            # שלב 2: הצגת המדד שנבדק
            first_indicator = list(scraper.cbs_indicators.items())[0]
            indicator_name, indicator_code = first_indicator
            add_log(f"שלב 2: המדד לבדיקה - {indicator_name} (קוד: {indicator_code})", 'yellow')
            
            # שלב 3: יצירת קובץ
            add_log("שלב 3: יוצר קובץ נתונים בסיסי...")
            try:
                file_path = scraper.create_data_file()
                add_log(f"✓ קובץ נוצר בהצלחה: {file_path}", 'green')
            except Exception as e:
                add_log(f"❌ שגיאה ביצירת קובץ: {str(e)}", 'red')
                return
            
            # שלב 4: הגדרת דפדפן
            add_log("שלב 4: מגדיר את הדפדפן (Chrome)...")
            try:
                scraper.setup_driver()
                if scraper.driver is None:
                    add_log("❌ שגיאה: לא הצלחתי ליצור דפדפן", 'red')
                    return
                add_log("✓ דפדפן הוגדר בהצלחה", 'green')
            except Exception as e:
                add_log(f"❌ שגיאה בהגדרת דפדפן: {str(e)}", 'red')
                return
            
            # שלב 5: שליפת המדד באמצעות הפונקציה המעודכנת שלך
            add_log("שלב 5: מריץ שליפת מדד עם הפונקציה המעודכנת...")
            try:
                # משתמש בפונקציה scrape_cbs_indicator שתיקנת
                result = scraper.scrape_cbs_indicator(indicator_name, indicator_code)
                if result:
                    add_log("✅ שליפת המדד הושלמה בהצלחה!", 'green')
                    add_log(f"תוצאה: {result}", 'green')
                    
                    # עדכון קובץ הנתונים
                    add_log("שלב 6: מעדכן קובץ נתונים...")
                    scraper.update_data_file_with_values({indicator_name: result})
                    add_log("✓ קובץ עודכן בהצלחה", 'green')
                else:
                    add_log("⚠️ שליפת המדד הושלמה אך ללא תוצאה", 'yellow')
            except Exception as e:
                add_log(f"❌ שגיאה בשליפת המדד: {str(e)}", 'red')
                return
            
        except Exception as main_e:
            add_log(f"❌ שגיאה כללית: {str(main_e)}", 'red')
        finally:
            # ניקוי
            try:
                if 'scraper' in locals() and scraper.driver:
                    scraper.close_driver()
                    add_log("✓ דפדפן נסגר", 'green')
            except:
                pass
            
            add_log("=== סיום בדיקת מדד ===", 'cyan')
    
    def fetch_all_madadim(self):
        """שליפת כל המדדים"""
        # ניקוי תוצאות קודמות
        for widget in self.results_frame.winfo_children():
            widget.destroy()
            
        # הודעת התחלה
        status_label = tk.Label(
            self.results_frame,
            text="מתחיל שליפת כל המדדים...",
            font=self.fonts['text'],
            bg='#f0f0f0',
            fg=self.colors['text']
        )
        status_label.pack(pady=10)
        
        self.root.update()
        
        try:
            # יצירת scraper
            scraper = MadadimScraper()
            
            # יצירת קובץ בסיסי
            scraper.create_data_file()
            
            status_label.config(text="שולף מדדים מאתר הלמ\"ס...")
            self.root.update()
            
            # שליפת כל המדדים
            cbs_results, bls_value = scraper.scrape_all_cbs_indicators()
            
            if cbs_results or bls_value:
                # עדכון הקובץ
                scraper.update_data_file_with_values(cbs_results, bls_value)
                
                success_label = tk.Label(
                    self.results_frame,
                    text=f"הושלמה שליפת {len(cbs_results)} מדדים מהלמ\"ס!",
                    font=self.fonts['text'],
                    bg='#f0f0f0',
                    fg='green'
                )
                success_label.pack(pady=10)
                
                # הצגת המדדים ששלפנו
                results_text = "מדדים ששלפנו:\n"
                for name, value in cbs_results.items():
                    results_text += f"• {name}: {value}\n"
                if bls_value:
                    results_text += f"• Consumer Price Index (BLS): {bls_value}\n"
                
                results_label = tk.Label(
                    self.results_frame,
                    text=results_text,
                    font=self.fonts['text'],
                    bg='#f0f0f0',
                    fg=self.colors['text'],
                    justify='right'
                )
                results_label.pack(pady=10)
            else:
                error_label = tk.Label(
                    self.results_frame,
                    text="לא הצלחנו לשלוף מדדים",
                    font=self.fonts['text'],
                    bg='#f0f0f0',
                    fg='red'
                )
                error_label.pack(pady=10)
                
        except Exception as e:
            error_label = tk.Label(
                self.results_frame,
                text=f"שגיאה: {str(e)}",
                font=self.fonts['text'],
                bg='#f0f0f0',
                fg='red'
            )
            error_label.pack(pady=10)

    def run(self):
        """הפעלת האפליקציה הראשית"""
        self.root.mainloop()

def main():
    """פונקציה ראשית"""
    app = MainApplication()
    app.run()

if __name__ == "__main__":
    main()
