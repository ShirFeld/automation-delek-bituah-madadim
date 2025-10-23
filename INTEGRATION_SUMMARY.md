# סיכום שילוב קובץ config.py במערכת

## תאריך: אוקטובר 2025

---

## מטרת השילוב
שילוב קובץ הגדרות מרכזי (config.py) בכל התוכנה כדי לנהל את כל הנתיבים וה-URLs במקום אחד.

---

## קבצים שעודכנו

### 1. **config.py** - קובץ ההגדרות המרכזי
**שינויים:**
- ✅ נוסף: `DELEKULATOR_URL` - אתר delekulator למחיר שירות עצמי
- ✅ תוקן: הסרת emoji מפונקציות print (בעיית encoding)

**משתנים זמינים:**
```python
# נתיבים
BASE_PARAMETERS_PATH
MADADIM_OUTPUT_PATH
BITUAH_RECHEV_OUTPUT_PATH
DELEK_OUTPUT_PATH

# URLs
CBS_URL                    # אתר הלמ"ס
BLS_URL                    # אתר BLS אמריקאי
MINISTRY_OF_TRANSPORT_URL  # משרד התחבורה
PAZ_URL                    # אתר פז
DELEKULATOR_URL           # אתר delekulator (חדש!)

# פונקציות עזר
ensure_directories_exist()
get_madadim_file_path(month, year)
get_bituah_rechev_image_path(timestamp)
get_bituah_rechev_mdb_path(timestamp)
get_delek_text_path(day, month, year)
get_delek_mdb_path(day, month, year)
```

---

### 2. **UpdateDelek/fuel_scraper.py** - סקריפט דלק
**שינויים:**
- ✅ נוסף: `import config`
- ✅ שונה: נתיב קבצים מקושח ל-`config.DELEK_OUTPUT_PATH`
- ✅ שונה: URL פז מקושח ל-`config.PAZ_URL`
- ✅ שונה: URL delekulator מקושח ל-`config.DELEKULATOR_URL`

**לפני:**
```python
base_path = r"C:\Users\shir.feldman\Desktop\parametrsUpdate\DELEK"
url = "https://www.paz.co.il/price-lists"
url = "https://delekulator.co.il/..."
```

**אחרי:**
```python
base_path = config.DELEK_OUTPUT_PATH
url = config.PAZ_URL
url = config.DELEKULATOR_URL
```

---

### 3. **Madadim/madadim_scraper.py** - סקריפט מדדים
**שינויים:**
- ✅ נוסף: `import config` (עם import path תיקון)
- ✅ שונה: נתיב קבצים מקושח ל-`config.MADADIM_OUTPUT_PATH`
- ✅ שונה: URLs מקושחים ל-`config.CBS_URL` ו-`config.BLS_URL`

**לפני:**
```python
self.cbs_url = "https://www.cbs.gov.il/..."
self.bls_url = "https://data.bls.gov/..."
self.target_path = r"C:\Users\shir.feldman\Desktop\parametrsUpdate\Madadim"
```

**אחרי:**
```python
self.cbs_url = config.CBS_URL
self.bls_url = config.BLS_URL
self.target_path = config.MADADIM_OUTPUT_PATH
```

---

### 4. **BituahRechev/insurance_scraper.py** - סקריפט ביטוח רכב
**שינויים:**
- ✅ נוסף: `import config` (עם import path תיקון)
- ✅ שונה: נתיב קבצים מקושח ל-`config.BITUAH_RECHEV_OUTPUT_PATH` (2 מקומות)
- ✅ שונה: URL משרד התחבורה מקושח ל-`config.MINISTRY_OF_TRANSPORT_URL`

**לפני:**
```python
url = "https://car.cma.gov.il/Parameters/Get?..."
save_path = r"C:\Users\shir.feldman\Desktop\parametrsUpdate\BituahRechev"
```

**אחרי:**
```python
url = config.MINISTRY_OF_TRANSPORT_URL
save_path = config.BITUAH_RECHEV_OUTPUT_PATH
```

---

### 5. **main_app.py** - אפליקציה ראשית
**שינויים:**
- ✅ נוסף: `import config`

---

### 6. **CONFIG_README.py** - מסמך תיעוד
**שינויים:**
- ✅ נוסף: סעיף "עדכונים אחרונים" עם פרטי השילוב

---

## יתרונות השילוב

### ✅ ניהול מרכזי
כל הנתיבים וה-URLs במקום אחד - קל לשנות ולעדכן

### ✅ תחזוקה קלה
שינוי נתיב אחד משפיע על כל התוכנה אוטומטית

### ✅ קוד נקי יותר
אין יותר נתיבים קשיחים (hard-coded paths) בקוד

### ✅ גמישות
קל להוסיף הגדרות חדשות או לשנות קיימות

### ✅ יצירת תיקיות אוטומטית
פונקציה `ensure_directories_exist()` יוצרת את כל התיקיות הנדרשות

---

## בדיקת התקינות

### הרצת בדיקה:
```bash
python config.py
```

**פלט צפוי:**
```
תיקייה מוכנה: C:\Users\shir.feldman\Desktop\parametrsUpdate\Madadim
תיקייה מוכנה: C:\Users\shir.feldman\Desktop\parametrsUpdate\BituahRechev
תיקייה מוכנה: C:\Users\shir.feldman\Desktop\parametrsUpdate\DELEK

כל התיקיות מוכנות!
```

---

## שימוש בתוכנה

### הרצת האפליקציה הראשית:
```bash
python main_app.py
```

כל הנתיבים וה-URLs ייקחו אוטומטית מקובץ ה-config!

---

## שינוי הגדרות

### דוגמה: שינוי תיקיית בסיס
ערוך את `config.py`:

```python
# לפני:
BASE_PARAMETERS_PATH = r"C:\Users\shir.feldman\Desktop\parametrsUpdate"

# אחרי (דוגמה):
BASE_PARAMETERS_PATH = r"D:\MyData\Parameters"
```

**כל שאר הנתיבים יתעדכנו אוטומטית!**

---

## קבצים שלא נגעו בהם
- ✅ `README.md` - ללא שינוי
- ✅ `requirements.txt` - ללא שינוי
- ✅ קבצי בדיקה (test_*.py) - נמחקו
- ✅ קבצי __pycache__ - נוקו

---

## סיכום טכני

**סה"כ קבצים ששונו:** 6
- config.py (עדכון)
- CONFIG_README.py (עדכון)
- fuel_scraper.py (שילוב)
- madadim_scraper.py (שילוב)
- insurance_scraper.py (שילוב)
- main_app.py (שילוב)

**נתיבים קשיחים שהוסרו:** 8
**URLs קשיחים שהוסרו:** 5

---

## בעיות שתוקנו
1. ✅ Emoji encoding errors בקובץ config.py
2. ✅ נתיבים קשיחים בכל הקבצים
3. ✅ URLs מפוזרים במספר קבצים

---

## המלצות לעתיד
1. 📌 תמיד עדכן את `config.py` במקום לשנות נתיבים בקוד
2. 📌 הוסף הגדרות חדשות ל-`config.py` כשצריך נתיבים/URLs חדשים
3. 📌 הרץ `python config.py` אחרי שינויים כדי לוודא שהתיקיות קיימות
4. 📌 עדכן את `CONFIG_README.py` כשמוסיפים הגדרות חדשות

---

**✅ השילוב הושלם בהצלחה!**

