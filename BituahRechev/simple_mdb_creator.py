#!/usr/bin/env python3
# -*- coding: utf-8 -*-


"""
יצירת קבצי ACCESS מהנתונים של הביטוח
"""

import os
import csv
import time
import builtins
import sys
from datetime import datetime, timedelta
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


def print(*args, **kwargs):
    """
    הדפסה בטוחה גם בסביבות שבהן קידוד הטרמינל לא תומך באימוג'י/יוניקוד.
    מונע קריסה של תהליך יצירת ה-MDB בגלל UnicodeEncodeError.
    """
    try:
        builtins.print(*args, **kwargs)
    except UnicodeEncodeError:
        sep = kwargs.get("sep", " ")
        end = kwargs.get("end", "\n")
        file_obj = kwargs.get("file", sys.stdout)
        flush = kwargs.get("flush", False)
        text = sep.join(str(a) for a in args)
        encoding = getattr(file_obj, "encoding", None) or "utf-8"
        safe_text = text.encode(encoding, errors="ignore").decode(encoding, errors="ignore")
        builtins.print(safe_text, end=end, file=file_obj, flush=flush)


def _safe_remove_file(path, retries=3, delay_sec=2):
    """מוחק קובץ עם ניסיונות חוזרים במקרה של נעילה זמנית."""
    if not os.path.exists(path):
        return True
    for attempt in range(1, retries + 1):
        try:
            os.remove(path)
            return True
        except PermissionError as e:
            print(f"⚠️ הקובץ נעול ({attempt}/{retries}): {path} -> {e}")
            if attempt < retries:
                time.sleep(delay_sec)
        except Exception as e:
            print(f"❌ שגיאה במחיקת קובץ: {path} -> {e}")
            return False
    return False


def count_scraped_insurance_prices(insurance_data):
    """סופר כמה מחירים אמיתיים (לא None) נשלפו מכל הקטגוריות."""
    if not insurance_data:
        return 0
    count = 0
    for group in (insurance_data.get('private_car') or {}).values():
        if group:
            count += sum(1 for p in group.values() if p is not None)
    for group in (insurance_data.get('commercial_car') or {}).values():
        if group:
            count += sum(1 for p in group.values() if p is not None)
    special = insurance_data.get('special_vehicle') or {}
    count += sum(1 for p in special.values() if p is not None)
    return count


def has_sufficient_insurance_data(insurance_data, minimum=1):
    """בודק שיש לפחות מחיר אחד לפני יצירת קבצים."""
    return count_scraped_insurance_prices(insurance_data) >= minimum


def get_bituah_effective_dates(reference_date=None):
    """
    תאריך יעיל לביטוח חובה: ה-1 לחודש הבא (זהה ללוגיקה ב-par_rech.dat).
    מחזיר: (next_month_dt, effective_date_dd_mm_yyyy, month_year_mmyy)
    """
    if reference_date is None:
        reference_date = datetime.now()
    next_month = (reference_date.replace(day=1) + timedelta(days=32)).replace(day=1)
    effective_date = next_month.strftime("%d/%m/%Y")
    month_year = next_month.strftime("%m%y")
    return next_month, effective_date, month_year


def _access_effective_date_for_com(access_app, effective_date_str):
    """
    תאריך לשדה Date/Time ב-Access.
    Python datetime בחצות גורם להצגה 30/06 21:00 (הזזת timezone).
    DateSerial מחזיר 01/07/2026 ללא שעה שגויה.
    """
    dt = datetime.strptime(effective_date_str, "%d/%m/%Y")
    return access_app.Eval(f"DateSerial({dt.year}, {dt.month}, {dt.day})")


COMMERCIAL_AGES = [17, 21, 24, 40, 50]
COMMERCIAL_AGE_GROUPS = ['17-20', '21-23', '24-39', '40-49', '50- ומעלה']
PRIVATE_AGES = [17, 21, 24, 30, 40, 50]
PRIVATE_AGE_GROUPS = ['17-20', '21-23', '24-29', '30-39', '40-49', '50- ומעלה']
PRIVATE_ENGINE_KEYS = ['עד 1050', 'מ-1051 עד 1550', 'מ-1551 עד 2050', 'מ-2051 ומעלה']
COMMERCIAL_WEIGHT_KEYS = ['עד 4000 (כולל)', 'מעל 4000']


def _to_int_or_none(value):
    if value is None:
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _pick_merged_value(scraped, fallback):
    """מחיר חדש מנצח; אחרת נשאר ערך מה-template (כמו par_rech.dat)."""
    if scraped is not None:
        return _to_int_or_none(scraped)
    return _to_int_or_none(fallback)


def _empty_mdb_fallback():
    return {
        'special': {'Nigrar': None, 'Handasi': None, 'Agricalture': None},
        'commercial': {},
        'private': {},
    }


def _read_mdb_fallback_from_db(db):
    """קורא שורות קיימות מה-MDB לפני DELETE – לשימוש כ-fallback."""
    fallback = _empty_mdb_fallback()
    try:
        rs = db.OpenRecordset("tblBituachHova_edit")
        if not rs.EOF:
            fallback['special'] = {
                'Nigrar': rs.Fields("Nigrar").Value,
                'Handasi': rs.Fields("Handasi").Value,
                'Agricalture': rs.Fields("Agricalture").Value,
            }
        rs.Close()
    except Exception as e:
        print(f"WARNING: לא ניתן לקרוא fallback מ-tblBituachHova_edit: {e}")

    try:
        rs = db.OpenRecordset("tblBituachHovaMishari_edit")
        while not rs.EOF:
            age = int(rs.Fields("Age").Value)
            fallback['commercial'][age] = {
                'Ad1': rs.Fields("Ad1").Value,
                'Ad2': rs.Fields("Ad2").Value,
            }
            rs.MoveNext()
        rs.Close()
    except Exception as e:
        print(f"WARNING: לא ניתן לקרוא fallback מ-tblBituachHovaMishari_edit: {e}")

    try:
        rs = db.OpenRecordset("tblBituachHovaPrati_edit")
        while not rs.EOF:
            age = int(rs.Fields("Age").Value)
            fallback['private'][age] = {
                'Ad1': rs.Fields("Ad1").Value,
                'Ad2': rs.Fields("Ad2").Value,
                'Ad3': rs.Fields("Ad3").Value,
                'Ad4': rs.Fields("Ad4").Value,
            }
            rs.MoveNext()
        rs.Close()
    except Exception as e:
        print(f"WARNING: לא ניתן לקרוא fallback מ-tblBituachHovaPrati_edit: {e}")

    return fallback


def _read_mdb_fallback_from_file(mdb_path):
    if not HAS_WIN32COM or not os.path.exists(mdb_path):
        return _empty_mdb_fallback()
    pythoncom.CoInitialize()
    try:
        access_app = win32com.client.Dispatch("Access.Application")
        access_app.OpenCurrentDatabase(mdb_path)
        fallback = _read_mdb_fallback_from_db(access_app.CurrentDb())
        access_app.CloseCurrentDatabase()
        access_app.Quit()
        print(f"OK נקראו ערכי fallback מ-{os.path.basename(mdb_path)}")
        return fallback
    except Exception as e:
        print(f"WARNING: לא ניתן לקרוא fallback מ-{mdb_path}: {e}")
        return _empty_mdb_fallback()
    finally:
        pythoncom.CoUninitialize()


def build_merged_mdb_rows(effective_date, insurance_data, fallback=None):
    """
    בונה את כל השורות ל-MDB: 1 מיוחד + 5 מסחרי + 6 פרטי.
    ערכים חדשים מה-scrape מחליפים; חסרים נשארים מה-template (כמו par_rech).
    """
    if fallback is None:
        fallback = _empty_mdb_fallback()

    special_fb = fallback.get('special') or {}
    special_sc = (insurance_data or {}).get('special_vehicle') or {}
    special_row = (
        effective_date,
        _pick_merged_value(special_sc.get('Nigrar'), special_fb.get('Nigrar')),
        _pick_merged_value(special_sc.get('Handasi'), special_fb.get('Handasi')),
        _pick_merged_value(special_sc.get('Agricalture'), special_fb.get('Agricalture')),
    )

    commercial_rows = []
    commercial_sc = (insurance_data or {}).get('commercial_car') or {}
    for i, age in enumerate(COMMERCIAL_AGES):
        age_group = COMMERCIAL_AGE_GROUPS[i]
        age_fb = (fallback.get('commercial') or {}).get(age) or {}
        age_sc = commercial_sc.get(age_group) or {}
        commercial_rows.append((
            effective_date,
            age,
            _pick_merged_value(age_sc.get(COMMERCIAL_WEIGHT_KEYS[0]), age_fb.get('Ad1')),
            _pick_merged_value(age_sc.get(COMMERCIAL_WEIGHT_KEYS[1]), age_fb.get('Ad2')),
        ))

    private_rows = []
    private_sc = (insurance_data or {}).get('private_car') or {}
    ad_fields = ['Ad1', 'Ad2', 'Ad3', 'Ad4']
    for i, age in enumerate(PRIVATE_AGES):
        age_group = PRIVATE_AGE_GROUPS[i]
        age_fb = (fallback.get('private') or {}).get(age) or {}
        age_sc = private_sc.get(age_group) or {}
        ads = [
            _pick_merged_value(age_sc.get(PRIVATE_ENGINE_KEYS[j]), age_fb.get(ad_fields[j]))
            for j in range(4)
        ]
        private_rows.append((effective_date, age, *ads))

    return {
        'special': special_row,
        'commercial': commercial_rows,
        'private': private_rows,
    }


def _sql_value(value):
    return 'NULL' if value is None else str(value)


def _insert_merged_rows_sql(access_app, effective_date, merged):
    """מכניס את כל 12 השורות לטבלאות Access (לאחר CREATE TABLE)."""
    _, nigrar, handasi, agricalture = merged['special']
    access_app.DoCmd.RunSQL(f"""
        INSERT INTO tblBituachHova_edit (EffectiveDate, Nigrar, Handasi, Agricalture)
        VALUES ('{effective_date}', {_sql_value(nigrar)}, {_sql_value(handasi)}, {_sql_value(agricalture)})
    """)
    print(f"✅ הכניס נתונים לטבלה 1: {nigrar}, {handasi}, {agricalture}")

    for _, age, ad1, ad2 in merged['commercial']:
        access_app.DoCmd.RunSQL(f"""
            INSERT INTO tblBituachHovaMishari_edit (EffectiveDate, Age, Ad1, Ad2)
            VALUES ('{effective_date}', {age}, {_sql_value(ad1)}, {_sql_value(ad2)})
        """)
        print(f"✅ רכב מסחרי גיל {age}: {ad1}, {ad2}")

    for _, age, ad1, ad2, ad3, ad4 in merged['private']:
        access_app.DoCmd.RunSQL(f"""
            INSERT INTO tblBituachHovaPrati_edit (EffectiveDate, Age, Ad1, Ad2, Ad3, Ad4)
            VALUES ('{effective_date}', {age}, {_sql_value(ad1)}, {_sql_value(ad2)}, {_sql_value(ad3)}, {_sql_value(ad4)})
        """)
        print(f"✅ רכב פרטי גיל {age}: {ad1}, {ad2}, {ad3}, {ad4}")


def _clear_access_table(db, table_name):
    """מוחק את כל השורות מטבלה (כולל שורות template עם תאריך ישן)."""
    for sql in (f"DELETE FROM {table_name}", f"DELETE * FROM {table_name}"):
        try:
            db.Execute(sql)
        except Exception:
            pass

    try:
        rs = db.OpenRecordset(f"SELECT COUNT(*) AS c FROM {table_name}")
        remaining = int(rs.Fields("c").Value or 0)
        rs.Close()
    except Exception:
        remaining = -1

    if remaining != 0:
        try:
            rs = db.OpenRecordset(table_name)
            if not rs.EOF:
                rs.MoveFirst()
            while not rs.EOF:
                rs.Delete()
                rs.MoveNext()
            rs.Close()
            print(f"OK נוקתה הטבלה {table_name} (מחיקה שורה-שורה)")
        except Exception as e:
            print(f"WARNING שגיאה בניקוי {table_name}: {e}")
            return False
    else:
        print(f"OK נוקתה הטבלה: {table_name}")
    return True


def create_insurance_files(save_path=None, insurance_data=None, mdb_filename=None):
    """יצירת קבצי נתונים לביטוח"""
    try:
        # אם לא סופק נתיב, נשתמש בנתיב הנכון
        if save_path is None:
            save_path = r"C:\Users\shir.feldman\Desktop\parametrsUpdate\BituahRechev"

        scraped_count = count_scraped_insurance_prices(insurance_data)
        if not has_sufficient_insurance_data(insurance_data):
            print(f"WARNING: אין נתונים ליצירת MDB ({scraped_count}/37) - מדלג")
            return None
        if scraped_count < 37:
            print(f"WARNING: נתונים חלקיים ({scraped_count}/37) - יוצר MDB עם מה שנשלף")
        
        print(f"🔍 מתחיל יצירת קבצי נתונים...")
        print(f"📂 נתיב: {save_path}")
        print(f"📊 נתונים: {insurance_data is not None} ({scraped_count}/37)")
        
        # יצירת תיקייה אם לא קיימת
        try:
            os.makedirs(save_path, exist_ok=True)
            print(f"✅ תיקייה מוכנה: {save_path}")
        except Exception as e:
            print(f"❌ שגיאה ביצירת תיקייה: {e}")
            return None
        
        # קביעת חודש היעד - תמיד החודש הבא (זהה ל-par_rech.dat)
        current_date = datetime.now()
        next_month, effective_date, target_month_year = get_bituah_effective_dates(current_date)

        # יצירת שם הקובץ - פורמט kneMMYY (מבוסס על החודש הבא)
        if mdb_filename:
            mdb_path = os.path.join(save_path, mdb_filename)
            month_year = mdb_filename.replace('kne', '').replace('.mdb', '')
        else:
            month_year = target_month_year
            mdb_path = os.path.join(save_path, f"kne{month_year}.mdb")

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

            template_path = os.path.join(save_path, "kne.mdb")
            fallback = _read_mdb_fallback_from_file(template_path)
            merged = build_merged_mdb_rows(effective_date, insurance_data, fallback)
            _insert_merged_rows_sql(access_app, effective_date, merged)
            print("✅ כל 3 הטבלאות נוצרו עם 12 שורות (1+5+6)")
            
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

        if os.path.normcase(os.path.abspath(mdb_path)) == os.path.normcase(os.path.abspath(template_path)):
            print("ERROR: נתיב פלט MDB זהה ל-template - לא ניתן לדרוס את kne.mdb")
            return None
        
        # בדיקה שה-template קיים
        if not os.path.exists(template_path):
            print(f"❌ קובץ template לא נמצא: {template_path}")
            return None
        
        print(f"📋 משתמש ב-template: {template_path}")
        
        # מחיקת קובץ יעד קיים
        if os.path.exists(mdb_path):
            if not _safe_remove_file(mdb_path, retries=4, delay_sec=2):
                print(f"❌ לא ניתן למחוק קובץ MDB קיים: {mdb_path}")
                return None
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

            # ניקוי נתונים קיימים מה-template (לעיתים נשאר תאריך החודש הנוכחי)
            db = access_app.CurrentDb()
            access_date = _access_effective_date_for_com(access_app, effective_date)
            print(f"🗓️ תאריך שייכנס ל-MDB: {effective_date} (DateSerial)")

            # קריאת ערכי fallback מה-template לפני מחיקה (כמו par_rech שומר ערכים קודמים)
            fallback = _read_mdb_fallback_from_db(db)
            merged = build_merged_mdb_rows(effective_date, insurance_data, fallback)

            for table_name in ["tblBituachHova_edit", "tblBituachHovaMishari_edit", "tblBituachHovaPrati_edit"]:
                _clear_access_table(db, table_name)

            # טבלה 1: רכב מיוחד – תמיד שורה אחת
            print("\n🔄 מכניס נתונים לטבלה 1 (רכב מיוחד)...")
            _, nigrar_value, handasi_value, agricalture_value = merged['special']
            recordset = db.OpenRecordset("tblBituachHova_edit")
            recordset.AddNew()
            recordset.Fields("EffectiveDate").Value = access_date
            recordset.Fields("Nigrar").Value = nigrar_value
            recordset.Fields("Handasi").Value = handasi_value
            recordset.Fields("Agricalture").Value = agricalture_value
            recordset.Update()
            recordset.Close()
            print(f"✅ הכניס נתונים לטבלה 1: {nigrar_value}, {handasi_value}, {agricalture_value}")

            # טבלה 2: רכב מסחרי – תמיד 5 שורות
            print("\n🔄 מכניס נתונים לטבלה 2 (רכב מסחרי)...")
            recordset2 = db.OpenRecordset("tblBituachHovaMishari_edit")
            for _, age, ad1_value, ad2_value in merged['commercial']:
                recordset2.AddNew()
                recordset2.Fields("EffectiveDate").Value = access_date
                recordset2.Fields("Age").Value = age
                recordset2.Fields("Ad1").Value = ad1_value
                recordset2.Fields("Ad2").Value = ad2_value
                recordset2.Update()
                print(f"✅ רכב מסחרי גיל {age}: {ad1_value}, {ad2_value}")
            recordset2.Close()

            # טבלה 3: רכב פרטי – תמיד 6 שורות
            print("\n🔄 מכניס נתונים לטבלה 3 (רכב פרטי)...")
            recordset3 = db.OpenRecordset("tblBituachHovaPrati_edit")
            for _, age, ad1_value, ad2_value, ad3_value, ad4_value in merged['private']:
                recordset3.AddNew()
                recordset3.Fields("EffectiveDate").Value = access_date
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
            if not _safe_remove_file(mdb_path, retries=4, delay_sec=2):
                print(f"❌ לא ניתן למחוק קובץ MDB קיים: {mdb_path}")
                return None
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

            template_path = os.path.join(os.path.dirname(mdb_path), "kne.mdb")
            fallback = _read_mdb_fallback_from_file(template_path)
            merged = build_merged_mdb_rows(effective_date, insurance_data, fallback)
            _insert_merged_rows_sql(access_app, effective_date, merged)
            
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

def prepare_all_tables_data(effective_date, insurance_data, fallback=None):
    """הכנת נתוני כל הטבלאות – תמיד 1+5+6 שורות (עם fallback מה-template)."""
    print(f"🔍 מתחיל הכנת טבלאות...")
    print(f"📅 תאריך יעיל: {effective_date}")

    merged = build_merged_mdb_rows(effective_date, insurance_data, fallback)

    tables_data = {
        'tblBituachHova_edit': {
            'headers': ['EffectiveDate', 'Nigrar', 'Handasi', 'Agricalture'],
            'rows': [merged['special']],
        },
        'tblBituachHovaMishari_edit': {
            'headers': ['EffectiveDate', 'Age', 'Ad1', 'Ad2'],
            'rows': merged['commercial'],
        },
        'tblBituachHovaPrati_edit': {
            'headers': ['EffectiveDate', 'Age', 'Ad1', 'Ad2', 'Ad3', 'Ad4'],
            'rows': merged['private'],
        },
    }
    print(f"✅ טבלה 1: {len(tables_data['tblBituachHova_edit']['rows'])} שורות")
    print(f"✅ טבלה 2: {len(tables_data['tblBituachHovaMishari_edit']['rows'])} שורות")
    print(f"✅ טבלה 3: {len(tables_data['tblBituachHovaPrati_edit']['rows'])} שורות")
    return tables_data

if __name__ == "__main__":
    create_insurance_files()
