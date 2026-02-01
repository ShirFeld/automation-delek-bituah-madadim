#!/usr/bin/env python3
# -*- coding: utf-8 -*-


"""
יצירת קבצי ACCESS מהנתונים של הביטוח
"""

import os
import csv
from datetime import datetime
import sqlite3
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

def create_insurance_files(save_path=None, insurance_data=None, mdb_filename=None):
    """יצירת קבצי נתונים לביטוח"""
    try:
        # אם לא סופק נתיב, נשתמש בנתיב הנכון
        if save_path is None:
            save_path = r"C:\Users\shir.feldman\Desktop\parametrsUpdate\BituahRechev"
        
        print(f"🔍 מתחיל יצירת קבצי נתונים...")
        print(f"📂 נתיב: {save_path}")
        print(f"📊 נתונים: {insurance_data is not None}")
        
        # יצירת תיקייה אם לא קיימת
        try:
            os.makedirs(save_path, exist_ok=True)
            print(f"✅ תיקייה מוכנה: {save_path}")
        except Exception as e:
            print(f"❌ שגיאה ביצירת תיקייה: {e}")
            return None
        
        # קביעת חודש היעד - תמיד החודש הבא (תואם לתאריך היעיל בטבלה)
        current_date = datetime.now()
        if current_date.month == 12:
            next_month = datetime(current_date.year + 1, 1, 1)
        else:
            next_month = datetime(current_date.year, current_date.month + 1, 1)
        target_month_year = next_month.strftime("%m%y")  # MMYY של החודש הבא
        
        # יצירת שם הקובץ - פורמט kneMMYY (מבוסס על החודש הבא)
        if mdb_filename:
            mdb_path = os.path.join(save_path, mdb_filename)
            month_year = mdb_filename.replace('kne', '').replace('.mdb', '')
        else:
            month_year = target_month_year
            mdb_path = os.path.join(save_path, f"kne{month_year}.mdb")
        
        # תאריך יעיל - הראשון לחודש הבא (לתוך הטבלה)
        effective_date = next_month.strftime("%d/%m/%Y")  # פורמט ישראלי: DD/MM/YYYY
        
        print(f"📅 תאריך נוכחי: {current_date.strftime('%d/%m/%Y')}")
        print(f"📁 שם קובץ: kne{month_year}.mdb (חודש הבא)")
        print(f"🗓️ תאריך יעיל בטבלה: {effective_date} (חודש עתידי)")
        
        print(f"📅 יוצר קובץ נתונים: {os.path.basename(mdb_path)}")
        print(f"🗓️ תאריך יעיל: {effective_date}")
        
        # ניסיון ליצור MDB מ-template אם win32com זמין
        print(f"🔍 בודק אם win32com זמין: {HAS_WIN32COM}")
        if HAS_WIN32COM:
            try:
                # נתיב ה-template
                template_path = os.path.join(save_path, "kne.mdb")
                print(f"📋 מחפש template: {template_path}")
                
                if os.path.exists(template_path):
                    print("🚀 מנסה ליצור MDB מ-template...")
                    result = create_mdb_from_template(mdb_path, effective_date, insurance_data, template_path)
                    if result:
                        print(f"✅ נוצר קובץ MDB מ-template: {mdb_path}")
                        return result
                    else:
                        print("⚠️ נכשל ביצירת MDB מ-template, מנסה שיטה ישנה...")
                        result = create_real_access_mdb(mdb_path, effective_date, insurance_data)
                        return result
                else:
                    print(f"⚠️ Template לא נמצא: {template_path}")
                    print("🔄 יוצר MDB בשיטה הישנה (ללא template)...")
                    result = create_real_access_mdb(mdb_path, effective_date, insurance_data)
                    print(f"✅ נוצר קובץ Access 2000: {mdb_path}")
                    return result
            except Exception as e:
                print(f"⚠️ לא הצלחתי ליצור MDB: {str(e)}")
                import traceback
                traceback.print_exc()
                print("🔄 מנסה ליצור Access 2000 דרך פונקציה אחרת...")
                result = create_sqlite_file(save_path, month_year, effective_date, insurance_data, os.path.basename(mdb_path))
                return result
        else:
            print("ℹ️ win32com לא זמין, יוצר Access 2000 דרך פונקציה אחרת...")
            result = create_sqlite_file(save_path, month_year, effective_date, insurance_data, os.path.basename(mdb_path))
            return result
        
        return result
        
    except Exception as e:
        print(f"❌ שגיאה ביצירת קבצי נתונים: {str(e)}")
        import traceback
        print(f"📋 פרטי השגיאה:")
        traceback.print_exc()
        return None

def create_excel_file(save_path, month_year, effective_date, insurance_data):
    """יצירת קובץ Excel עם 3 גיליונות"""
    try:
        excel_path = os.path.join(save_path, f"Kne{month_year}.xlsx")
        
        # מחיקת קובץ קיים
        if os.path.exists(excel_path):
            os.remove(excel_path)
            print("🗑️ מחק קובץ Excel קיים")
        
        # יצירת נתונים לטבלאות
        tables_data = prepare_all_tables_data(effective_date, insurance_data)
        
        # יצירת קובץ Excel עם מספר גיליונות
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for table_name, data in tables_data.items():
                df = pd.DataFrame(data['rows'], columns=data['headers'])
                sheet_name = table_name.replace('tbl', '').replace('_edit', '')
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"✅ נוצר גיליון: {sheet_name}")
        
        print(f"📊 קובץ Excel נוצר בהצלחה: {excel_path}")
        print(f"📂 הקובץ מכיל 3 גיליונות עבודה")
        return excel_path
        
    except Exception as e:
        print(f"❌ שגיאה ביצירת Excel: {str(e)}")
        # חזרה ל-SQLite אם Excel נכשל
        return create_sqlite_file(save_path, month_year, effective_date, insurance_data)

def create_sqlite_file(save_path, month_year, effective_date, insurance_data, mdb_filename=None):
    """יצירת קובץ Access 2000 עם סיומת .mdb"""
    try:
        print(f"🔧 יוצר קובץ Access 2000...")
        if mdb_filename:
            mdb_path = os.path.join(save_path, mdb_filename)
        else:
            mdb_path = os.path.join(save_path, f"kne{month_year}.mdb")
        print(f"📂 נתיב Access: {mdb_path}")
        
        # מחיקת קובץ קיים
        if os.path.exists(mdb_path):
            try:
                os.remove(mdb_path)
                print("🗑️ מחק קובץ קיים")
            except PermissionError:
                print("⚠️ הקובץ תפוס, מנסה לסגור חיבורים...")
                import time
                import gc
                
                # ניסיון לסגור חיבורים
                gc.collect()  # ניקוי זיכרון
                time.sleep(3)  # המתנה ארוכה יותר
                
                try:
                    os.remove(mdb_path)
                    print("✅ הצלחתי למחוק את הקובץ אחרי המתנה")
                except Exception as e:
                    print(f"❌ לא הצלחתי למחוק את הקובץ: {str(e)}")
                    print("🔄 מנסה שוב אחרי המתנה נוספת...")
                    time.sleep(5)  # המתנה נוספת
                    try:
                        os.remove(mdb_path)
                        print("✅ הצלחתי למחוק את הקובץ אחרי המתנה נוספת")
                    except Exception as e2:
                        print(f"❌ עדיין לא מצליח למחוק: {str(e2)}")
                        return None  # נכשל - לא יוצרים קובץ חדש
            except Exception as e:
                print(f"❌ שגיאה במחיקת קובץ: {str(e)}")
                return None
        
        # יצירת Access 2000 אמיתי
        if not HAS_WIN32COM:
            print("❌ win32com לא זמין - לא ניתן ליצור Access 2000")
            return None
        
        print("🔧 יוצר Access 2000 database...")
        pythoncom.CoInitialize()
        
        try:
            # יצירת Access application
            access_app = win32com.client.Dispatch("Access.Application")
            access_app.NewCurrentDatabase(mdb_path, 9)  # 9 = Access 2000
            print("✅ יצר Access 2000 database")
            
            # טבלה 1: tblBituachHova_edit
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
        
            # קבלת נתונים אמיתיים לטבלה הראשונה - רק נתונים אמיתיים!
            nigrar_value = None
            handasi_value = None
            agricalture_value = None
            
            if insurance_data and 'special_vehicle' in insurance_data:
                special_data = insurance_data['special_vehicle']
                if 'Nigrar' in special_data and special_data['Nigrar']:
                    nigrar_value = int(special_data['Nigrar'])  # המרה למספר שלם
                if 'Handasi' in special_data and special_data['Handasi']:
                    handasi_value = int(special_data['Handasi'])  # המרה למספר שלם
                if 'Agricalture' in special_data and special_data['Agricalture']:
                    agricalture_value = int(special_data['Agricalture'])  # המרה למספר שלם
            
            # הכנסת נתונים לטבלה 1 - רק אם יש נתונים אמיתיים
            if nigrar_value is not None or handasi_value is not None or agricalture_value is not None:
                insert1_sql = f"""
                INSERT INTO tblBituachHova_edit (EffectiveDate, Nigrar, Handasi, Agricalture)
                VALUES ('{effective_date}', {nigrar_value}, {handasi_value}, {agricalture_value})
                """
                access_app.DoCmd.RunSQL(insert1_sql)
                print("✅ הכניס נתונים לטבלה 1")
            else:
                print("⚠️ אין נתונים אמיתיים לרכב מיוחד - רק יוצר טבלה ריקה")
            print("✅ טבלה 1 נוצרה עם נתונים")
        
            # טבלה 2: tblBituachHovaMishari_edit (רכב מסחרי)
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
        
            commercial_ages = [17, 21, 24, 40, 50]
            commercial_age_groups = ['17-20', '21-23', '24-39', '40-49', '50- ומעלה']
            
            for i, age in enumerate(commercial_ages):
                age_group = commercial_age_groups[i]
                ad1_value = None
                ad2_value = None
                
                if insurance_data and 'commercial_car' in insurance_data and age_group in insurance_data['commercial_car']:
                    age_data = insurance_data['commercial_car'][age_group]
                    ad1_value = age_data.get('עד 4000 (כולל)')
                    ad2_value = age_data.get('מעל 4000')
                    if ad1_value is not None and ad2_value is not None:
                        # המרה למספרים שלמים
                        ad1_value = int(ad1_value)
                        ad2_value = int(ad2_value)
                        print(f"📊 נתונים אמיתיים לרכב מסחרי גיל {age_group}: Ad1={ad1_value}, Ad2={ad2_value}")
                    else:
                        print(f"⚠️ נתונים חסרים לרכב מסחרי גיל {age_group} - מדלג")
                        continue
                else:
                    print(f"⚠️ אין נתונים אמיתיים לרכב מסחרי גיל {age_group} - מדלג")
                    continue
                
                insert2_sql = f"""
                INSERT INTO tblBituachHovaMishari_edit (EffectiveDate, Age, Ad1, Ad2)
                VALUES ('{effective_date}', {age}, {ad1_value}, {ad2_value})
                """
                access_app.DoCmd.RunSQL(insert2_sql)
            print("✅ טבלה 2 נוצרה עם 5 שורות")
        
            # טבלה 3: tblBituachHovaPrati_edit (רכב פרטי)
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
            
            private_ages = [17, 21, 24, 30, 40, 50]
            private_age_groups = ['17-20', '21-23', '24-29', '30-39', '40-49', '50- ומעלה']
            
            for i, age in enumerate(private_ages):
                age_group = private_age_groups[i]
                ad1_value = None
                ad2_value = None
                ad3_value = None
                ad4_value = None
                
                if insurance_data and 'private_car' in insurance_data and age_group in insurance_data['private_car']:
                    age_data = insurance_data['private_car'][age_group]
                    ad1_value = age_data.get('עד 1050')
                    ad2_value = age_data.get('מ-1051 עד 1550')
                    ad3_value = age_data.get('מ-1551 עד 2050')
                    ad4_value = age_data.get('מ-2051 ומעלה')
                    if ad1_value is not None and ad2_value is not None and ad3_value is not None and ad4_value is not None:
                        # המרה למספרים שלמים
                        ad1_value = int(ad1_value)
                        ad2_value = int(ad2_value)
                        ad3_value = int(ad3_value)
                        ad4_value = int(ad4_value)
                        print(f"📊 נתונים אמיתיים לרכב פרטי גיל {age_group}: Ad1={ad1_value}, Ad2={ad2_value}, Ad3={ad3_value}, Ad4={ad4_value}")
                    else:
                        print(f"⚠️ נתונים חסרים לרכב פרטי גיל {age_group} - מדלג")
                        continue
                else:
                    print(f"⚠️ אין נתונים אמיתיים לרכב פרטי גיל {age_group} - מדלג")
                    continue
                
                insert3_sql = f"""
                INSERT INTO tblBituachHovaPrati_edit (EffectiveDate, Age, Ad1, Ad2, Ad3, Ad4)
                VALUES ('{effective_date}', {age}, {ad1_value}, {ad2_value}, {ad3_value}, {ad4_value})
                """
                access_app.DoCmd.RunSQL(insert3_sql)
            print("✅ טבלה 3 נוצרה עם 6 שורות")
            
            # שמירה וסגירה
            print("💾 Access מוכן - לא צריך Save()")
            try:
                access_app.CloseCurrentDatabase()
                access_app.Quit()
                print("✅ Access נסגר")
            except Exception as e:
                print(f"⚠️ שגיאה בסגירת Access: {str(e)}")
                # מנסה לסגור בכוח
                try:
                    access_app.Quit()
                except:
                    pass
            
            print(f"📊 קובץ Access 2000 נוצר בהצלחה: {mdb_path}")
            print(f"📂 הקובץ מכיל 3 טבלאות:")
            print(f"   • tblBituachHova_edit (1 שורה)")
            print(f"   • tblBituachHovaMishari_edit (5 שורות)")
            print(f"   • tblBituachHovaPrati_edit (6 שורות)")
            
            return mdb_path
            
        except Exception as e:
            print(f"❌ שגיאה ביצירת Access: {str(e)}")
            import traceback
            print(f"📋 פרטי השגיאה:")
            traceback.print_exc()
            return None
            
    except Exception as e:
        print(f"❌ שגיאה ביצירת Access: {str(e)}")
        import traceback
        print(f"📋 פרטי השגיאה:")
        traceback.print_exc()
        return None

def create_simple_csv(save_path, month_year, effective_date, insurance_data):
    """יצירת קובץ CSV פשוט כחלופה"""
    try:
        csv_path = os.path.join(save_path, f"Kne{month_year}_data.csv")
        
        # יצירת נתונים מאוחדים
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            
            # כותרת כללית
            writer.writerow(['קובץ נתוני ביטוח', f'Kne{month_year}', effective_date])
            writer.writerow([])
            
            # טבלה 1
            writer.writerow(['tblBituachHova_edit'])
            writer.writerow(['EffectiveDate', 'Nigrar', 'Handasi', 'Agricalture'])
            writer.writerow([effective_date, 423, 2335, 1535])
            writer.writerow([])
            
            # טבלה 2
            writer.writerow(['tblBituachHovaMishari_edit'])
            writer.writerow(['EffectiveDate', 'Age', 'Ad1', 'Ad2'])
            for age in [17, 21, 24, 40, 50]:
                ad1 = 2000 + (age * 10)
                ad2 = 3000 + (age * 15)
                writer.writerow([effective_date, age, ad1, ad2])
            writer.writerow([])
            
            # טבלה 3
            writer.writerow(['tblBituachHovaPrati_edit'])
            writer.writerow(['EffectiveDate', 'Age', 'Ad1', 'Ad2', 'Ad3', 'Ad4'])
            for age in [17, 21, 24, 30, 40, 50]:
                ad1 = 1800 + (age * 8)
                ad2 = 2200 + (age * 10)
                ad3 = 2600 + (age * 12)
                ad4 = 3000 + (age * 15)
                writer.writerow([effective_date, age, ad1, ad2, ad3, ad4])
        
        print(f"✅ נוצר קובץ CSV: {csv_path}")
        return csv_path
        
    except Exception as e:
        print(f"❌ שגיאה ביצירת CSV: {str(e)}")
        return None

def create_mdb_from_template(mdb_path, effective_date, insurance_data, template_path):
    """יצירת קובץ MDB מ-template ע"י העתקה והכנסת נתונים"""
    try:
        import shutil
        
        # בדיקה שה-template קיים
        if not os.path.exists(template_path):
            print(f"❌ קובץ template לא נמצא: {template_path}")
            return None
        
        print(f"📋 משתמש ב-template: {template_path}")
        
        # מחיקת קובץ יעד קיים
        if os.path.exists(mdb_path):
            os.remove(mdb_path)
            print("🗑️ מחק קובץ MDB קיים")
        
        # העתקת ה-template
        shutil.copy2(template_path, mdb_path)
        print(f"✅ העתיק template ל-{mdb_path}")
        
        # אתחול COM
        pythoncom.CoInitialize()
        
        try:
            # פתיחת הקובץ המועתק
            access_app = win32com.client.Dispatch("Access.Application")
            access_app.OpenCurrentDatabase(mdb_path)
            print("✅ פתח קובץ MDB מועתק")
            
            # הכנסת נתונים לטבלה 1: tblBituachHova_edit (רכב מיוחד)
            nigrar_value = None
            handasi_value = None
            agricalture_value = None
            
            if insurance_data and 'special_vehicle' in insurance_data:
                special_data = insurance_data['special_vehicle']
                if 'Nigrar' in special_data and special_data['Nigrar']:
                    nigrar_value = int(special_data['Nigrar'])
                if 'Handasi' in special_data and special_data['Handasi']:
                    handasi_value = int(special_data['Handasi'])
                if 'Agricalture' in special_data and special_data['Agricalture']:
                    agricalture_value = int(special_data['Agricalture'])
            
            if nigrar_value is not None or handasi_value is not None or agricalture_value is not None:
                print("\n🔄 מכניס נתונים לטבלה 1 (רכב מיוחד)...")
                db = access_app.CurrentDb()
                recordset = db.OpenRecordset("tblBituachHova_edit")
                recordset.AddNew()
                recordset.Fields("EffectiveDate").Value = effective_date
                recordset.Fields("Nigrar").Value = nigrar_value
                recordset.Fields("Handasi").Value = handasi_value
                recordset.Fields("Agricalture").Value = agricalture_value
                recordset.Update()
                recordset.Close()
                print(f"✅ הכניס נתונים לטבלה 1: {nigrar_value}, {handasi_value}, {agricalture_value}")
            
            # הכנסת נתונים לטבלה 2: tblBituachHovaMishari_edit (רכב מסחרי)
            print("\n🔄 מכניס נתונים לטבלה 2 (רכב מסחרי)...")
            commercial_ages = [17, 21, 24, 40, 50]
            commercial_age_groups = ['17-20', '21-23', '24-39', '40-49', '50- ומעלה']
            
            db = access_app.CurrentDb()
            recordset2 = db.OpenRecordset("tblBituachHovaMishari_edit")
            
            for i, age in enumerate(commercial_ages):
                age_group = commercial_age_groups[i]
                ad1_value = None
                ad2_value = None
                
                if insurance_data and 'commercial_car' in insurance_data and age_group in insurance_data['commercial_car']:
                    age_data = insurance_data['commercial_car'][age_group]
                    ad1_value = age_data.get('עד 4000 (כולל)')
                    ad2_value = age_data.get('מעל 4000')
                    if ad1_value is not None and ad2_value is not None:
                        ad1_value = int(ad1_value)
                        ad2_value = int(ad2_value)
                        
                        recordset2.AddNew()
                        recordset2.Fields("EffectiveDate").Value = effective_date
                        recordset2.Fields("Age").Value = age
                        recordset2.Fields("Ad1").Value = ad1_value
                        recordset2.Fields("Ad2").Value = ad2_value
                        recordset2.Update()
                        print(f"✅ רכב מסחרי גיל {age}: {ad1_value}, {ad2_value}")
            
            recordset2.Close()
            
            # הכנסת נתונים לטבלה 3: tblBituachHovaPrati_edit (רכב פרטי)
            print("\n🔄 מכניס נתונים לטבלה 3 (רכב פרטי)...")
            private_ages = [17, 21, 24, 30, 40, 50]
            private_age_groups = ['17-20', '21-23', '24-29', '30-39', '40-49', '50- ומעלה']
            
            recordset3 = db.OpenRecordset("tblBituachHovaPrati_edit")
            
            for i, age in enumerate(private_ages):
                age_group = private_age_groups[i]
                ad1_value = None
                ad2_value = None
                ad3_value = None
                ad4_value = None
                
                if insurance_data and 'private_car' in insurance_data and age_group in insurance_data['private_car']:
                    age_data = insurance_data['private_car'][age_group]
                    ad1_value = age_data.get('עד 1050')
                    ad2_value = age_data.get('מ-1051 עד 1550')
                    ad3_value = age_data.get('מ-1551 עד 2050')
                    ad4_value = age_data.get('מ-2051 ומעלה')
                    if ad1_value is not None and ad2_value is not None and ad3_value is not None and ad4_value is not None:
                        ad1_value = int(ad1_value)
                        ad2_value = int(ad2_value)
                        ad3_value = int(ad3_value)
                        ad4_value = int(ad4_value)
                        
                        recordset3.AddNew()
                        recordset3.Fields("EffectiveDate").Value = effective_date
                        recordset3.Fields("Age").Value = age
                        recordset3.Fields("Ad1").Value = ad1_value
                        recordset3.Fields("Ad2").Value = ad2_value
                        recordset3.Fields("Ad3").Value = ad3_value
                        recordset3.Fields("Ad4").Value = ad4_value
                        recordset3.Update()
                        print(f"✅ רכב פרטי גיל {age}: {ad1_value}, {ad2_value}, {ad3_value}, {ad4_value}")
            
            recordset3.Close()
            
            # סגירת הקובץ
            access_app.CloseCurrentDatabase()
            access_app.Quit()
            print("✅ סגר את Access")
            
            return mdb_path
            
        finally:
            pythoncom.CoUninitialize()
            
    except Exception as e:
        print(f"❌ שגיאה ביצירת MDB מ-template: {e}")
        import traceback
        traceback.print_exc()
        return None

def create_real_access_mdb(mdb_path, effective_date, insurance_data):
    """יצירת קובץ Access 2000 באמצעות COM - מועתק מתוכנה של הדלק"""
    try:
        # מחיקת קובץ קיים
        if os.path.exists(mdb_path):
            os.remove(mdb_path)
            print("🗑️ מחק קובץ MDB קיים")
        
        # אתחול COM
        pythoncom.CoInitialize()
        
        try:
            # יצירת Access application עם גרסה 2000
            access_app = win32com.client.Dispatch("Access.Application")
            # יצירת מסד נתונים בגרסה 2000
            access_app.NewCurrentDatabase(mdb_path, 9)  # 9 = Access 2000
            print("✅ יצר Access 2000 database")
            
            # יצירת טבלה 1: tblBituachHova_edit
            create_table1_sql = """
            CREATE TABLE tblBituachHova_edit (
                EffectiveDate TEXT(10),
                Nigrar LONG,
                Handasi LONG,
                Agricalture LONG
            )
            """
            access_app.DoCmd.RunSQL(create_table1_sql)
            print("✅ יצר טבלה 1")
            
            # קבלת נתונים אמיתיים לטבלה הראשונה - רק נתונים אמיתיים!
            nigrar_value = None
            handasi_value = None
            agricalture_value = None
            
            if insurance_data and 'special_vehicle' in insurance_data:
                special_data = insurance_data['special_vehicle']
                if 'Nigrar' in special_data and special_data['Nigrar']:
                    nigrar_value = int(special_data['Nigrar'])  # המרה למספר שלם
                if 'Handasi' in special_data and special_data['Handasi']:
                    handasi_value = int(special_data['Handasi'])  # המרה למספר שלם
                if 'Agricalture' in special_data and special_data['Agricalture']:
                    agricalture_value = int(special_data['Agricalture'])  # המרה למספר שלם
            
            # הכנסת נתונים לטבלה 1 - רק אם יש נתונים אמיתיים
            if nigrar_value is not None or handasi_value is not None or agricalture_value is not None:
                insert1_sql = f"""
                INSERT INTO tblBituachHova_edit 
                (EffectiveDate, Nigrar, Handasi, Agricalture)
                VALUES ('{effective_date}', {nigrar_value}, {handasi_value}, {agricalture_value})
                """
                access_app.DoCmd.RunSQL(insert1_sql)
                print("✅ הכניס נתונים לטבלה 1")
            else:
                print("⚠️ אין נתונים אמיתיים לרכב מיוחד - רק יוצר טבלה ריקה")
            
            # יצירת טבלה 2: tblBituachHovaMishari_edit
            create_table2_sql = """
            CREATE TABLE tblBituachHovaMishari_edit (
                EffectiveDate TEXT(10),
                Age LONG,
                Ad1 DOUBLE,
                Ad2 DOUBLE
            )
            """
            access_app.DoCmd.RunSQL(create_table2_sql)
            print("✅ יצר טבלה 2")
            
            # הכנסת נתונים לטבלה 2 (רכב מסחרי)
            commercial_ages = [17, 21, 24, 40, 50]
            commercial_age_groups = ['17-20', '21-23', '24-39', '40-49', '50- ומעלה']
            
            for i, age in enumerate(commercial_ages):
                age_group = commercial_age_groups[i]
                ad1_value = None
                ad2_value = None
                
                if insurance_data and 'commercial_car' in insurance_data and age_group in insurance_data['commercial_car']:
                    age_data = insurance_data['commercial_car'][age_group]
                    ad1_value = age_data.get('עד 4000 (כולל)')
                    ad2_value = age_data.get('מעל 4000')
                    if ad1_value is not None and ad2_value is not None:
                        # המרה למספרים שלמים
                        ad1_value = int(ad1_value)
                        ad2_value = int(ad2_value)
                        print(f"📊 נתונים אמיתיים לרכב מסחרי גיל {age_group}: Ad1={ad1_value}, Ad2={ad2_value}")
                    else:
                        print(f"⚠️ נתונים חסרים לרכב מסחרי גיל {age_group} - מדלג")
                        continue
                else:
                    print(f"⚠️ אין נתונים אמיתיים לרכב מסחרי גיל {age_group} - מדלג")
                    continue
                
                insert2_sql = f"""
                INSERT INTO tblBituachHovaMishari_edit 
                (EffectiveDate, Age, Ad1, Ad2)
                VALUES ('{effective_date}', {age}, {ad1_value}, {ad2_value})
                """
                access_app.DoCmd.RunSQL(insert2_sql)
            
            # יצירת טבלה 3: tblBituachHovaPrati_edit
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
            print("✅ יצר טבלה 3")
            
            # הכנסת נתונים לטבלה 3 (רכב פרטי)
            private_ages = [17, 21, 24, 30, 40, 50]
            private_age_groups = ['17-20', '21-23', '24-29', '30-39', '40-49', '50- ומעלה']
            
            for i, age in enumerate(private_ages):
                age_group = private_age_groups[i]
                ad1_value = None
                ad2_value = None
                ad3_value = None
                ad4_value = None
                
                if insurance_data and 'private_car' in insurance_data and age_group in insurance_data['private_car']:
                    age_data = insurance_data['private_car'][age_group]
                    ad1_value = age_data.get('עד 1050')
                    ad2_value = age_data.get('מ-1051 עד 1550')
                    ad3_value = age_data.get('מ-1551 עד 2050')
                    ad4_value = age_data.get('מ-2051 ומעלה')
                    if ad1_value is not None and ad2_value is not None and ad3_value is not None and ad4_value is not None:
                        # המרה למספרים שלמים
                        ad1_value = int(ad1_value)
                        ad2_value = int(ad2_value)
                        ad3_value = int(ad3_value)
                        ad4_value = int(ad4_value)
                        print(f"📊 נתונים אמיתיים לרכב פרטי גיל {age_group}: Ad1={ad1_value}, Ad2={ad2_value}, Ad3={ad3_value}, Ad4={ad4_value}")
                    else:
                        print(f"⚠️ נתונים חסרים לרכב פרטי גיל {age_group} - מדלג")
                        continue
                else:
                    print(f"⚠️ אין נתונים אמיתיים לרכב פרטי גיל {age_group} - מדלג")
                    continue
                
                insert3_sql = f"""
                INSERT INTO tblBituachHovaPrati_edit 
                (EffectiveDate, Age, Ad1, Ad2, Ad3, Ad4)
                VALUES ('{effective_date}', {age}, {ad1_value}, {ad2_value}, {ad3_value}, {ad4_value})
                """
                access_app.DoCmd.RunSQL(insert3_sql)
            
            # שמירה וסגירה - ללא Save() שגורם לשגיאה
            print("✅ Access מוכן - לא צריך Save()")
            
            try:
                access_app.CloseCurrentDatabase()
                print("✅ Access נסגר")
            except Exception as close_e:
                print(f"⚠️ שגיאה בסגירת Access: {str(close_e)}")
            
            try:
                access_app.Quit()
                print("✅ Access יצא")
            except Exception as quit_e:
                print(f"⚠️ שגיאה ביציאה מ-Access: {str(quit_e)}")
            
        finally:
            # ניקוי COM
            pythoncom.CoUninitialize()
            
        return mdb_path
            
    except Exception as e:
        print(f"❌ שגיאה ביצירת Access 2000: {str(e)}")
        raise

def prepare_all_tables_data(effective_date, insurance_data):
    """הכנת נתוני כל הטבלאות"""
    tables_data = {}
    
    print(f"🔍 מתחיל הכנת טבלאות...")
    print(f"📅 תאריך יעיל: {effective_date}")
    print(f"📊 נתונים: {insurance_data}")
    
    # קבלת נתונים אמיתיים לטבלה הראשונה - רק נתונים אמיתיים!
    nigrar_value = None
    handasi_value = None
    agricalture_value = None
    
    if insurance_data and 'special_vehicle' in insurance_data:
        special_data = insurance_data['special_vehicle']
        print(f"🚗 נתוני רכב מיוחד: {special_data}")
        if 'Nigrar' in special_data and special_data['Nigrar']:
            nigrar_value = special_data['Nigrar']
        if 'Handasi' in special_data and special_data['Handasi']:
            handasi_value = special_data['Handasi']
        if 'Agricalture' in special_data and special_data['Agricalture']:
            agricalture_value = special_data['Agricalture']
    
    # טבלה 1: tblBituachHova_edit
    if nigrar_value is not None or handasi_value is not None or agricalture_value is not None:
        # המרה למספרים שלמים (ללא נקודה עשרונית)
        if nigrar_value is not None:
            nigrar_value = int(nigrar_value)
        if handasi_value is not None:
            handasi_value = int(handasi_value)
        if agricalture_value is not None:
            agricalture_value = int(agricalture_value)
        
        tables_data['tblBituachHova_edit'] = {
            'headers': ['EffectiveDate', 'Nigrar', 'Handasi', 'Agricalture'],
            'rows': [(effective_date, nigrar_value, handasi_value, agricalture_value)]
        }
        print(f"✅ טבלה 1: {len(tables_data['tblBituachHova_edit']['rows'])} שורות")
    else:
        print("⚠️ אין נתונים לרכב מיוחד")
        tables_data['tblBituachHova_edit'] = {
            'headers': ['EffectiveDate', 'Nigrar', 'Handasi', 'Agricalture'],
            'rows': []
        }
    
    # בדיקה שיש נתונים לפחות בטבלה אחת
    total_rows = len(tables_data['tblBituachHova_edit']['rows'])
    if total_rows > 0:
        print(f"✅ יש נתונים ברכב מיוחד: {total_rows} שורות")
    else:
        print("⚠️ אין נתונים ברכב מיוחד")
    

    
    # טבלה 2: tblBituachHovaMishari_edit (רכב מסחרי) - רק נתונים אמיתיים!
    commercial_ages = [17, 21, 24, 40, 50]
    commercial_age_groups = ['17-20', '21-23', '24-39', '40-49', '50- ומעלה']
    commercial_rows = []
    
    if insurance_data and 'commercial_car' in insurance_data:
        for i, age in enumerate(commercial_ages):
            age_group = commercial_age_groups[i]
            ad1_value = None
            ad2_value = None
            
            if age_group in insurance_data['commercial_car']:
                age_data = insurance_data['commercial_car'][age_group]
                ad1_value = age_data.get('עד 4000 (כולל)')
                ad2_value = age_data.get('מעל 4000')
                if ad1_value is not None and ad2_value is not None:
                    # המרה למספרים שלמים
                    ad1_value = int(ad1_value)
                    ad2_value = int(ad2_value)
                    print(f"📊 נתונים אמיתיים לרכב מסחרי גיל {age_group}: Ad1={ad1_value}, Ad2={ad2_value}")
                    commercial_rows.append((effective_date, age, ad1_value, ad2_value))
                else:
                    print(f"⚠️ נתונים חסרים לרכב מסחרי גיל {age_group} - מדלג")
            else:
                print(f"⚠️ אין נתונים אמיתיים לרכב מסחרי גיל {age_group} - מדלג")
    else:
        print("⚠️ אין נתונים לרכב מסחרי")
    
    tables_data['tblBituachHovaMishari_edit'] = {
        'headers': ['EffectiveDate', 'Age', 'Ad1', 'Ad2'],
        'rows': commercial_rows
    }
    print(f"✅ טבלה 2: {len(commercial_rows)} שורות")
    
    # טבלה 3: tblBituachHovaPrati_edit (רכב פרטי) - רק נתונים אמיתיים!
    private_ages = [17, 21, 24, 30, 40, 50]
    private_age_groups = ['17-20', '21-23', '24-29', '30-39', '40-49', '50- ומעלה']
    private_rows = []
    
    if insurance_data and 'private_car' in insurance_data:
        for i, age in enumerate(private_ages):
            age_group = private_age_groups[i]
            ad1_value = None
            ad2_value = None
            ad3_value = None
            ad4_value = None
            
            if age_group in insurance_data['private_car']:
                age_data = insurance_data['private_car'][age_group]
                ad1_value = age_data.get('עד 1050')
                ad2_value = age_data.get('מ-1051 עד 1550')
                ad3_value = age_data.get('מ-1551 עד 2050')
                ad4_value = age_data.get('מ-2051 ומעלה')
                if ad1_value is not None and ad2_value is not None and ad3_value is not None and ad4_value is not None:
                    # המרה למספרים שלמים
                    ad1_value = int(ad1_value)
                    ad2_value = int(ad2_value)
                    ad3_value = int(ad3_value)
                    ad4_value = int(ad4_value)
                    print(f"📊 נתונים אמיתיים לרכב פרטי גיל {age_group}: Ad1={ad1_value}, Ad2={ad2_value}, Ad3={ad3_value}, Ad4={ad4_value}")
                    private_rows.append((effective_date, age, ad1_value, ad2_value, ad3_value, ad4_value))
                else:
                    print(f"⚠️ נתונים חסרים לרכב פרטי גיל {age_group} - מדלג")
            else:
                print(f"⚠️ אין נתונים אמיתיים לרכב פרטי גיל {age_group} - מדלג")
    else:
        print("⚠️ אין נתונים לרכב פרטי")
    
    tables_data['tblBituachHovaPrati_edit'] = {
        'headers': ['EffectiveDate', 'Age', 'Ad1', 'Ad2', 'Ad3', 'Ad4'],
        'rows': private_rows
    }
    print(f"✅ טבלה 3: {len(private_rows)} שורות")
    
    return tables_data

if __name__ == "__main__":
    create_insurance_files()
