#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sys
import os

# הוספת נתיב לייבוא
sys.path.append(os.path.dirname(__file__))

class CommercialInsuranceGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.setup_window()
        self.create_interface()
        
    def setup_window(self):
        """הגדרת החלון"""
        self.root.title("שליפת ביטוח רכב מסחרי")
        self.root.geometry("700x500")
        self.root.configure(bg='#f0f0f0')
        
        # מרכז החלון במסך
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        # צבעים
        self.colors = {
            'primary': '#FF9800',
            'primary_hover': '#F57C00',
            'background': '#f0f0f0',
            'surface': '#ffffff',
            'text': '#323130',
            'text_secondary': '#605e5c'
        }
        
    def create_interface(self):
        """יצירת הממשק"""
        # כותרת
        header_frame = tk.Frame(self.root, bg=self.colors['primary'], height=100)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="🚛 שליפת ביטוח רכב מסחרי",
            font=('Segoe UI', 18, 'bold'),
            bg=self.colors['primary'],
            fg='white'
        )
        title_label.pack(pady=20)
        
        subtitle_label = tk.Label(
            header_frame,
            text="מאתר משרד התחבורה - חברת הראל",
            font=('Segoe UI', 12),
            bg=self.colors['primary'],
            fg='white'
        )
        subtitle_label.pack()
        
        # תוכן מרכזי
        main_frame = tk.Frame(self.root, bg=self.colors['background'])
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # מידע על התרחישים
        info_frame = tk.Frame(main_frame, bg=self.colors['surface'], relief='ridge', bd=2)
        info_frame.pack(fill='x', pady=(0, 20))
        
        info_text = """🎯 התוכנה תריץ 10 תרחישים שונים:

📋 קבוצות גיל:
• 17-20: גיל 19, שנות רישוי 2
• 21-23: גיל 22, שנות רישוי 5  
• 24-39: גיל 30, שנות רישוי 13
• 40-49: גיל 42, שנות רישוי 17
• 50+: גיל 51, שנות רישוי 26

⚖️ משקלי רכב לכל קבוצה:
• עד 4000 ק"ג (כולל)
• מעל 4000 ק"ג

🏢 מקור הנתונים: חברת הראל ביטוח"""

        info_label = tk.Label(
            info_frame,
            text=info_text,
            font=('Segoe UI', 10),
            bg=self.colors['surface'],
            fg=self.colors['text'],
            justify='right',
            padx=15,
            pady=15
        )
        info_label.pack(fill='x')
        
        # כפתור התחלה
        self.start_button = tk.Button(
            main_frame,
            text="🚀 התחל שליפת תרחישי רכב מסחרי",
            font=('Segoe UI', 14, 'bold'),
            bg=self.colors['primary'],
            fg='white',
            relief='flat',
            bd=0,
            padx=40,
            pady=15,
            cursor='hand2',
            command=self.start_scraping
        )
        self.start_button.pack(pady=20)
        
        # אפקט hover
        self.start_button.bind('<Enter>', lambda e: self.start_button.config(bg=self.colors['primary_hover']))
        self.start_button.bind('<Leave>', lambda e: self.start_button.config(bg=self.colors['primary']))
        
        # אזור תוצאות
        result_frame = tk.Frame(main_frame, bg=self.colors['surface'], relief='ridge', bd=2)
        result_frame.pack(fill='both', expand=True)
        
        result_title = tk.Label(
            result_frame,
            text="תוצאות השליפה:",
            font=('Segoe UI', 12, 'bold'),
            bg=self.colors['surface'],
            fg=self.colors['text']
        )
        result_title.pack(anchor='e', padx=10, pady=(10, 5))
        
        # טקסט תוצאות
        text_frame = tk.Frame(result_frame)
        text_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        
        self.result_text = tk.Text(text_frame, height=8, font=('Consolas', 9), wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.result_text.yview)
        self.result_text.config(yscrollcommand=scrollbar.set)
        
        self.result_text.pack(side='right', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # סטטוס בר
        self.status_label = tk.Label(
            self.root,
            text="מוכן לעבודה",
            font=('Segoe UI', 9),
            bg=self.colors['background'],
            fg=self.colors['text_secondary'],
            relief='sunken',
            anchor='w'
        )
        self.status_label.pack(fill='x', side='bottom')
        
    def update_status(self, message):
        """עדכון הסטטוס"""
        self.status_label.config(text=message)
        self.root.update()
        
    def display_results(self, message):
        """הצגת תוצאות"""
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)
        self.root.update()
        
    def start_scraping(self):
        """התחלת השליפה"""
        self.start_button.config(state='disabled', text="מעבד 10 תרחישים...")
        self.update_status("מתחיל שליפה...")
        self.result_text.delete(1.0, tk.END)
        
        def scrape_task():
            try:
                from insurance_scraper import InsuranceScraper
                
                self.display_results("🚛 מתחיל תהליך שליפה מלאה - רכב מסחרי!")
                self.display_results("📋 כל קבוצות הגיל: 17-20, 21-23, 24-39, 40-49, 50+")
                self.display_results("🎯 סך הכל: 10 תרחישים שונים\n")
                
                scraper = InsuranceScraper()
                self.update_status("מגדיר דפדפן...")
                
                if scraper.setup_driver(visible=False):
                    self.display_results("✅ דפדפן הוגדר בהצלחה!")
                    self.update_status("מתחיל שליפת כל התרחישים...")
                    self.display_results("⚡ רץ במצב מהירות מקסימלית\n")
                    
                    # שליפה מלאה של רכב מסחרי
                    results = scraper.scrape_commercial_vehicle_complete()
                    
                    if results and any(any(group.values()) if group else False for group in results.values()):
                        self.display_results("🎉 שליפה מלאה הושלמה בהצלחה!")
                        
                        total_successful = 0
                        total_scenarios = 0
                        
                        # הצגת תוצאות לכל קבוצת גיל
                        for age_group, group_results in results.items():
                            if group_results:
                                successful_in_group = sum(1 for price in group_results.values() if price)
                                total_in_group = len(group_results)
                                total_successful += successful_in_group
                                total_scenarios += total_in_group
                                
                                self.display_results(f"\n📊 גיל {age_group}:")
                                for weight_group, price in group_results.items():
                                    if price:
                                        self.display_results(f"   ✅ {weight_group}: {price:,.0f} ₪")
                                    else:
                                        self.display_results(f"   ❌ {weight_group}: לא נמצא מחיר")
                        
                        self.display_results(f"\n🏆 סיכום סופי: {total_successful}/{total_scenarios} תרחישים הצליחו")
                        
                        # שמירת הטבלאות כתמונה
                        image_path = scraper.save_tables_as_image()
                        
                        if image_path:
                            self.display_results(f"\n📷 תמונת הטבלאות נשמרה ב:\n{image_path}")
                        
                        self.update_status(f"הושלם - {total_successful}/{total_scenarios} הצליחו")
                        
                        # הודעת הצלחה
                        msg = f"שליפה מלאה רכב מסחרי הושלמה!\n"
                        msg += f"הצליח: {total_successful}/{total_scenarios} תרחישים\n"
                        msg += f"מקור: חברת הראל ביטוח"
                        
                        if image_path:
                            msg += f"\n\nתמונה נשמרה ב:\n{image_path}"
                        
                        messagebox.showinfo("הצלחה", msg)
                    else:
                        self.display_results("❌ שליפה נכשלה - לא נמצאו נתונים")
                        self.update_status("שליפה נכשלה")
                        messagebox.showerror("שגיאה", "השליפה נכשלה.\nיתכן שהאתר לא זמין.")
                    
                    scraper.cleanup()
                else:
                    self.display_results("❌ שגיאה בהגדרת הדפדפן")
                    self.update_status("שגיאה בהגדרת דפדפן")
                    messagebox.showerror("שגיאה", "לא ניתן להגדיר את הדפדפן.\nוודא שכרום מותקן.")
                
            except Exception as e:
                self.display_results(f"❌ שגיאה: {str(e)}")
                self.update_status(f"שגיאה: {str(e)}")
                messagebox.showerror("שגיאה", f"אירעה שגיאה:\n{str(e)}")
                
            finally:
                self.start_button.config(state='normal', text="🚀 התחל שליפת תרחישי רכב מסחרי")
        
        # הרצה בחוט נפרד
        threading.Thread(target=scrape_task, daemon=True).start()
        
    def run(self):
        """הפעלת הממשק"""
        self.root.mainloop()

def main():
    """פונקציה ראשית"""
    app = CommercialInsuranceGUI()
    app.run()

if __name__ == "__main__":
    main()

