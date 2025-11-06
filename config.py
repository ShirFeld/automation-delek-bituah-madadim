#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
קובץ הגדרות מרכזי לתוכנה
מכיל את כל הנתיבים והגדרות של המערכת
"""

import os

# תיקיית הבסיס לכל הפרמטרים
BASE_PARAMETERS_PATH = r"C:\Users\shir.feldman\Desktop\parametrsUpdate"

# תיקיית יעד לקבצי מדדים
MADADIM_OUTPUT_PATH = os.path.join(BASE_PARAMETERS_PATH, "Madadim")

# פורמט שם קובץ מדדים: madadimMMYY.txt
# דוגמה: madadim0924.txt (ספטמבר 2024)
MADADIM_FILENAME_FORMAT = "madadim{month:02d}{year:02d}.txt"

# נתיבי ביטוח רכב (BituahRechev)

# תיקיית יעד לקבצי ביטוח רכב
BITUAH_RECHEV_OUTPUT_PATH = os.path.join(BASE_PARAMETERS_PATH, "BituahRechev")

# נתיב מקור לקובץ par_rech.dat (לקריאה)
BITUAH_RECHEV_PARAM_SOURCE_PATH = r"p:\kolnatun\updates\paramPro"
BITUAH_RECHEV_PARAM_SOURCE_FILE = os.path.join(BITUAH_RECHEV_PARAM_SOURCE_PATH, "par_rech.dat")

# פורמט שם קובץ תמונת טבלאות: insurance_tables_YYYYMMDD_HHMMSS.png
BITUAH_RECHEV_IMAGE_FILENAME_FORMAT = "insurance_tables_{timestamp}.png"

# פורמט שם קובץ MDB: insurance_data_YYYYMMDD_HHMMSS.mdb
BITUAH_RECHEV_MDB_FILENAME_FORMAT = "insurance_data_{timestamp}.mdb"


# תיקיית יעד לקבצי דלק
DELEK_OUTPUT_PATH = os.path.join(BASE_PARAMETERS_PATH, "DELEK")

# נתיב מקור לקובץ par_dlk.dat (לקריאה)
DELEK_PARAM_SOURCE_PATH = r"p:\kolnatun\updates\paramPro"
DELEK_PARAM_SOURCE_FILE = os.path.join(DELEK_PARAM_SOURCE_PATH, "par_dlk.dat")

# פורמט שם קובץ טקסט: DDMMYY.txt
# דוגמה: 200924.txt (20 ספטמבר 2024)
DELEK_TEXT_FILENAME_FORMAT = "{day:02d}{month:02d}{year:02d}.txt"

# פורמט שם קובץ MDB: DDMMYY.mdb
# דוגמה: 200924.mdb (20 ספטמבר 2024)
DELEK_MDB_FILENAME_FORMAT = "{day:02d}{month:02d}{year:02d}.mdb"

# הגדרות כלליות

# יצירת תיקיות באופן אוטומטי אם לא קיימות
AUTO_CREATE_DIRECTORIES = True

# קידוד קבצים
DEFAULT_ENCODING = 'utf-8'

# הגדרות זמן
# יום וזמן גבול לחישוב חודש קודם (למדדים)
# אם היום <= 15 והשעה לפני 18:30, נחזור עוד חודש אחורה
MADADIM_DAY_THRESHOLD = 15
MADADIM_TIME_THRESHOLD = (18, 30)  # (שעה, דקות)


# הגדרות אתרים חיצוניים

# אתר הלמ"ס (מדדים)
CBS_URL = "https://www.cbs.gov.il/he/Statistics/Pages/%D7%9E%D7%97%D7%95%D7%9C%D7%9C%D7%99%D7%9D/%D7%9E%D7%97%D7%95%D7%9C%D7%9C-%D7%9E%D7%97%D7%99%D7%A8%D7%99%D7%9D.aspx"

# אתר BLS (מדד אמריקאי)
BLS_URL = "https://data.bls.gov/dataViewer/view/timeseries/CUUR0000SA0"

# אתר משרד התחבורה (ביטוח רכב)
MINISTRY_OF_TRANSPORT_URL = "https://car.cma.gov.il/Parameters/Get?next_page=2&curr_page=1&playAnimation=true&fontSize=12"

# אתר פז (דלק)
PAZ_URL = "https://www.paz.co.il/price-lists"

# אתר delekulator (מחיר שירות עצמי)
DELEKULATOR_URL = "https://delekulator.co.il/היסטוריית-מחירי-הדלק/"


# פונקציות עזר

def ensure_directories_exist():
    """יוצר את כל התיקיות הנדרשות אם הן לא קיימות"""
    if AUTO_CREATE_DIRECTORIES:
        directories = [
            MADADIM_OUTPUT_PATH,
            BITUAH_RECHEV_OUTPUT_PATH,
            DELEK_OUTPUT_PATH
        ]
        
        for directory in directories:
            os.makedirs(directory, exist_ok=True)
            print(f"תיקייה מוכנה: {directory}")

def get_madadim_file_path(month, year):
    """מחזיר נתיב מלא לקובץ מדדים"""
    filename = MADADIM_FILENAME_FORMAT.format(month=month, year=year % 100)
    return os.path.join(MADADIM_OUTPUT_PATH, filename)

def get_bituah_rechev_image_path(timestamp):
    """מחזיר נתיב מלא לקובץ תמונת טבלאות ביטוח"""
    filename = BITUAH_RECHEV_IMAGE_FILENAME_FORMAT.format(timestamp=timestamp)
    return os.path.join(BITUAH_RECHEV_OUTPUT_PATH, filename)

def get_bituah_rechev_mdb_path(timestamp):
    """מחזיר נתיב מלא לקובץ MDB ביטוח"""
    filename = BITUAH_RECHEV_MDB_FILENAME_FORMAT.format(timestamp=timestamp)
    return os.path.join(BITUAH_RECHEV_OUTPUT_PATH, filename)

def get_delek_text_path(day, month, year):
    """מחזיר נתיב מלא לקובץ טקסט דלק"""
    filename = DELEK_TEXT_FILENAME_FORMAT.format(day=day, month=month, year=year % 100)
    return os.path.join(DELEK_OUTPUT_PATH, filename)

def get_delek_mdb_path(day, month, year):
    """מחזיר נתיב מלא לקובץ MDB דלק"""
    filename = DELEK_MDB_FILENAME_FORMAT.format(day=day, month=month, year=year % 100)
    return os.path.join(DELEK_OUTPUT_PATH, filename)


"""
אם תרצה להעביר את התוכנה למקום אחר, העתק את כל תיקיית UpdateDelek (לא רק את ה-EXE)
התיקייה _internal חייבת להישאר לצד ה-EXE
"""


# ===================================
# הרצת בדיקות בעת טעינת הקובץ
# ===================================

if __name__ == "__main__":
    ensure_directories_exist()
    print("\nכל התיקיות מוכנות!")

