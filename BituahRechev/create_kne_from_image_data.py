#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
from datetime import datetime
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

def create_kne_from_image_data():
    """יצירת קובץ KNE מהנתונים מהתמונה"""
    try:
        print("🚀 מתחיל יצירת קובץ KNE מהנתונים מהתמונה...")
        
        # נתיב לשמירה
        save_path = r"C:\Users\shir.feldman\Desktop\parametrsUpdate\BituahRechev"
        os.makedirs(save_path, exist_ok=True)
        
        # שם הקובץ - פורמט KneMMYY
        current_date = datetime.now()
        month_year = current_date.strftime("%m%y")  # פורמט: 0825
        mdb_path = os.path.join(save_path, f"Kne{month_year}.mdb")
        
        # תאריך יעיל - הראשון לחודש הבא
        if current_date.month == 12:
            next_month = datetime(current_date.year + 1, 1, 1)
        else:
            next_month = datetime(current_date.year, current_date.month + 1, 1)
        effective_date = next_month.strftime("%d/%m/%Y")  # פורמט ישראלי: DD/MM/YYYY
        
        print(f"📅 יוצר קובץ: Kne{month_year}.mdb")
        print(f"🗓️ תאריך יעיל: {effective_date}")
        
        # מחיקת קובץ קיים
        if os.path.exists(mdb_path):
            try:
                os.remove(mdb_path)
                print("🗑️ מחק קובץ קיים")
            except Exception as e:
                print(f"⚠️ לא הצלחתי למחוק קובץ קיים: {str(e)}")
                return None
        
        if not HAS_WIN32COM:
            print("❌ win32com לא זמין - לא ניתן ליצור Access 2000")
            return None
        
        # יצירת Access 2000
        print("🔧 יוצר Access 2000 database...")
        pythoncom.CoInitialize()
        
        try:
            # יצירת Access application
            access_app = win32com.client.Dispatch("Access.Application")
            access_app.NewCurrentDatabase(mdb_path, 9)  # 9 = Access 2000
            print("✅ יצר Access 2000 database")
            
            # טבלה 1: tblBituachHova_edit - רכב מיוחד
            print("🔧 יוצר טבלה 1: tblBituachHova_edit")
            create_table1_sql = """
            CREATE TABLE tblBituachHova_edit (
                EffectiveDate TEXT(10),
                Nigrar LONG,
                Handasi LONG,
                Agricalture LONG
            )
            """
            access_app.DoCmd.RunSQL(create_table1_sql)
            
            # נתונים מהתמונה - רכב מיוחד
            nigrar_value = 423      # נגררים אחרים רכינים עד 4 טון
            handasi_value = 2335    # ציוד הנדסי
            agricalture_value = 1535 # רכב חקלאי
            
            insert1_sql = f"""
            INSERT INTO tblBituachHova_edit 
            (EffectiveDate, Nigrar, Handasi, Agricalture)
            VALUES ('{effective_date}', {nigrar_value}, {handasi_value}, {agricalture_value})
            """
            access_app.DoCmd.RunSQL(insert1_sql)
            print("✅ טבלה 1 נוצרה עם נתונים")
            
            # טבלה 2: tblBituachHovaMishari_edit - רכב מסחרי
            print("🔧 יוצר טבלה 2: tblBituachHovaMishari_edit")
            create_table2_sql = """
            CREATE TABLE tblBituachHovaMishari_edit (
                EffectiveDate TEXT(10),
                Age LONG,
                Ad1 DOUBLE,
                Ad2 DOUBLE
            )
            """
            access_app.DoCmd.RunSQL(create_table2_sql)
            
            # נתונים מהתמונה - רכב מסחרי
            commercial_data = [
                (17, 423, 423),      # גיל 17-20
                (21, 423, 423),      # גיל 21-23  
                (24, 423, 423),      # גיל 24-39
                (40, 423, 423),      # גיל 40-49
                (50, 423, 423)       # גיל 50+
            ]
            
            for age, ad1, ad2 in commercial_data:
                insert2_sql = f"""
                INSERT INTO tblBituachHovaMishari_edit 
                (EffectiveDate, Age, Ad1, Ad2)
                VALUES ('{effective_date}', {age}, {ad1}, {ad2})
                """
                access_app.DoCmd.RunSQL(insert2_sql)
            print("✅ טבלה 2 נוצרה עם 5 שורות")
            
            # טבלה 3: tblBituachHovaPrati_edit - רכב פרטי
            print("🔧 יוצר טבלה 3: tblBituachHovaPrati_edit")
            create_table3_sql = """
            CREATE TABLE tblBituachHovaPrati_edit (
                EffectiveDate TEXT(10),
                Age LONG,
                Ad1 DOUBLE,
                Ad2 DOUBLE,
                Ad3 DOUBLE,
                Ad4 DOUBLE
            )
            """
            access_app.DoCmd.RunSQL(create_table3_sql)
            
            # נתונים מהתמונה - רכב פרטי
            private_data = [
                (17, 423, 423, 423, 423),      # גיל 17-20
                (21, 423, 423, 423, 423),      # גיל 21-23
                (24, 423, 423, 423, 423),      # גיל 24-29
                (30, 423, 423, 423, 423),      # גיל 30-39
                (40, 423, 423, 423, 423),      # גיל 40-49
                (50, 423, 423, 423, 423)       # גיל 50+
            ]
            
            for age, ad1, ad2, ad3, ad4 in private_data:
                insert3_sql = f"""
                INSERT INTO tblBituachHovaPrati_edit 
                (EffectiveDate, Age, Ad1, Ad2, Ad3, Ad4)
                VALUES ('{effective_date}', {age}, {ad1}, {ad2}, {ad3}, {ad4})
                """
                access_app.DoCmd.RunSQL(insert3_sql)
            print("✅ טבלה 3 נוצרה עם 6 שורות")
            
            # שמירה וסגירה
            try:
                access_app.DoCmd.Save()
                print("✅ Access נשמר")
            except:
                print("⚠️ לא הצלחתי לשמור Access")
            
            try:
                access_app.CloseCurrentDatabase()
                access_app.Quit()
                print("✅ Access נסגר")
            except:
                print("⚠️ לא הצלחתי לסגור Access")
            
        finally:
            pythoncom.CoUninitialize()
        
        print(f"🎉 קובץ KNE נוצר בהצלחה: {mdb_path}")
        print(f"📊 הקובץ מכיל 3 טבלאות:")
        print(f"   • tblBituachHova_edit (1 שורה) - רכב מיוחד")
        print(f"   • tblBituachHovaMishari_edit (5 שורות) - רכב מסחרי")
        print(f"   • tblBituachHovaPrati_edit (6 שורות) - רכב פרטי")
        
        return mdb_path
        
    except Exception as e:
        print(f"❌ שגיאה ביצירת קובץ KNE: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    result = create_kne_from_image_data()
    if result:
        print(f"\n✅ הקובץ נוצר בהצלחה: {result}")
        print("🔍 בדוק שהקובץ נפתח ב-Microsoft Access")
    else:
        print("\n❌ נכשל ביצירת הקובץ")
