# בדיקות יומיות (Daily Checks) – מדריך התקנה

הכלי מצלם את מסכי CNOM ואת הדשבורדים של Metabase, מריץ את בדיקות הטלפונים ושולח את התוצאות במייל דרך Brevo.

**כל ההתקנה נעשית ברמת המשתמש שלך – לא צריך הרשאות אדמין.**

---

## מה צריך לפני שמתחילים

- חיבור לרשת החברה / VPN, עם גישה ל-CNOM ול-Metabase
- קובץ `config.ini` מאחראי הכלי (פרטי ההתחברות המשותפים, מפתח Brevo, שולח ונמען)
- כ-500MB פנויים בדיסק

---

## שלב 1 – התקנת Python (ללא אדמין)

1. להוריד את Python מהאתר הרשמי: https://www.python.org/downloads/windows/
2. בחלון ההתקנה:
   - לסמן ✅ **Add python.exe to PATH**
   - לוודא ש-**Use admin privileges when installing py.exe** **לא** מסומן
   - ללחוץ **Install Now**

   ההתקנה נכנסת לתיקיית המשתמש שלך, ולכן לא תתבקש סיסמת אדמין.
3. לפתוח חלון **PowerShell חדש** ולבדוק:

```powershell
python --version
```

אם מוצגת גרסה (למשל `Python 3.14.x`), אפשר להמשיך.

---

## שלב 2 – העתקת קבצי הכלי

מה-branch צריך להעתיק את הקבצים הבאים בדיוק במבנה הזה:

```
C:\dailychecks\
├── allinone.exe        (או allinone.py)
├── config.ini          ← ההגדרות (מאחראי הכלי בלבד)
├── ms-playwright\      ← הדפדפן Chromium (רק כשמריצים את ה-exe)
├── phonechecks\        ← להעתיק את כל התיקייה כמו שהיא, עם כל הקבצים שבתוכה
└── README.md
```

כללים חשובים:

- שם התיקייה: `dailychecks`, באותיות קטנות ובלי רווחים.
- אם אין לך הרשאה ליצור תיקייה ישירות תחת `C:\`, אפשר לשים אותה בתיקיית המשתמש: `%USERPROFILE%\dailychecks`.
- לא לשים את התיקייה בתוך OneDrive או בשולחן עבודה מסונכרן, כי הסנכרון עלול לנעול קבצים בזמן ריצה.
- לא לשנות שמות של קבצים או תיקיות בתוך `phonechecks`.

---

## שלב 3 – התקנת הספריות

לפתוח PowerShell ולהריץ:

```powershell
python -m pip install --upgrade pip
python -m pip install playwright pillow requests python-docx
python -m playwright install chromium
```

הדפדפן (Chromium) מותקן לתיקייה `%LOCALAPPDATA%\ms-playwright` של המשתמש שלך, בלי צורך באדמין.

**אם יש פרוקסי בחברה** וההורדה נכשלת:

```powershell
$env:HTTPS_PROXY = 'http://PROXY_ADDRESS:PORT'
python -m pip install --proxy http://PROXY_ADDRESS:PORT playwright pillow requests python-docx
python -m playwright install chromium
```

---

## שלב 4 – קובץ ההגדרות config.ini (אחראי הכלי בלבד)

**לא צריך יותר משתני סביבה.** כל ההגדרות נמצאות בקובץ אחד, `config.ini`, שיושב ליד `allinone.exe`.
אחראי הכלי מגדיר אותו **פעם אחת בכל מחשב**, ומאותו רגע כל משתמש Windows במחשב יכול להריץ את הכלי בלי להגדיר כלום.

1. להעתיק את `config.example.ini` לאותה תיקייה, ולשנות את שמו ל-`config.ini`.
2. לפתוח אותו ב-Notepad ולמלא את הערכים:

```ini
[cnom]
username = shared_cnom_username
password = shared_cnom_password

[metabase]
username = shared_metabase_username
password = shared_metabase_password

[email]
brevo_api_key = your_brevo_api_key
from = verified_sender@company.com
to = receiver@company.com
```

שימו לב:

- כותבים את הערכים **בלי מירכאות**. תווים כמו `$`, `%` או `"` בסיסמה תקינים.
- `from` חייבת להיות הכתובת שאומתה ב-Brevo. כל המיילים נשלחים ממנה לנמען שב-`to`.
- כדי לדעת מי שלח כל דוח, בנושא המייל מופיעים תאריך הבדיקה ושם הבודק, ובגוף המייל מופיעים שם המחשב ומשתמש ה-Windows.
- אם ערך נשאר ריק, הכלי משתמש במשתנה הסביבה המתאים (למשל `CNOM_USERNAME`), אם הוגדר כזה.

### 📁 קבצים אישיים לכל משתמש

הכלי שומר באופן אוטומטי את הקבצים של כל משתמש Windows בתיקייה משלו:

```
%LOCALAPPDATA%\dailychecks\
├── session.json / metabase_session.json   ← החיבור ל-CNOM ול-Metabase
├── screenshots\                            ← צילומי המסך
└── reports\                                ← דוחות המשמרת שמולאו
```

ככה כמה אנשים יכולים לעבוד על אותו מחשב, כל אחד עם משתמש Windows משלו, בלי לדרוס קבצים זה של זה ובלי שגיאות הרשאה.
תבנית הדוח שב-`phonechecks\` רק נקראת, ואף פעם לא נדרסת.

### 🔒 אבטחה

- כל מי שמריץ את הכלי יכול לפתוח את `config.ini`. לכן כדאי להגביל ב-Brevo את המפתח לכתובת ה-IP של המשרד (Security ← Authorized IPs), ולהשתמש במפתח שנוצר רק בשביל הכלי הזה.
- לא מעלים את `config.ini` ל-Git (הוא כבר ב-`.gitignore`).

---

## שלב 5 – הרצה

```powershell
cd C:\dailychecks
.\allinone.exe
```

או, אם מריצים את הסקריפט ישירות:

```powershell
cd C:\dailychecks
python allinone.py
```

בסיום הריצה נשלח מייל עם התוצאות לנמען שהוגדר ב-`config.ini`.

> בהרצה הראשונה של ה-exe ייתכן ש-Windows SmartScreen יחסום אותו. במקרה כזה יש ללחוץ **More info** ואז **Run anyway**.

---

## הרצה אוטומטית כל בוקר (אופציונלי, ללא אדמין)

1. תפריט התחל ← **Task Scheduler** (מתזמן המשימות)
2. **Create Basic Task** ← לתת שם, למשל `Daily Checks`
3. Trigger: **Daily**, ולבחור שעה
4. Action: **Start a program**
   - **Program/script:** `C:\dailychecks\allinone.exe`
   - **Start in:** `C:\dailychecks` (מומלץ; הכלי מוצא את `config.ini` ואת `phonechecks` גם בלי זה)
5. לסיים. המשימה תרוץ כשהמשתמש מחובר (Run only when user is logged on), וזה לא דורש אדמין.

אם מריצים את `allinone.py` במקום ה-exe:

- **Program/script:** הנתיב המלא של Python. אפשר למצוא אותו עם הפקודה `(Get-Command python).Source`
- **Add arguments:** `allinone.py`
- **Start in:** `C:\dailychecks`

---

## פתרון תקלות

| הודעה / תסמין | סיבה | פתרון |
|---|---|---|
| `'python' is not recognized` | Python לא נוסף ל-PATH | לפתוח חלון PowerShell חדש. אם זה לא עוזר, להתקין מחדש ולסמן Add to PATH |
| הפקודה `python` פותחת את Microsoft Store | כינוי מערכת של Windows | הגדרות ← אפליקציות ← הגדרות אפליקציה מתקדמות ← App execution aliases ← לכבות את python.exe ו-python3.exe |
| `Executable doesn't exist ... ms-playwright` | הדפדפן לא הותקן | להריץ `python -m playwright install chromium` |
| `Timeout 30000ms exceeded` ב-Metabase | Metabase או מסד הנתונים שלו לא זמינים, או שאין VPN | לפתוח את Metabase בדפדפן רגיל ולבדוק. אם הוא למטה, הבעיה אינה בכלי |
| חלון "Daily Checks — setup needed" | `config.ini` חסר או שחסרים בו ערכים | לפנות לאחראי הכלי (שלב 4) |
| שגיאת `401` מ-Brevo | מפתח API שגוי, או שה-IP של המחשב לא מאושר ב-Brevo | לבדוק את `brevo_api_key` ואת רשימת ה-IP המאושרים |
| המייל לא נשלח / שגיאת sender | כתובת השולח לא מאומתת ב-Brevo | לבדוק את `from` ב-`config.ini` |
| שגיאת SSL או חיבור ב-pip | פרוקסי של החברה | להשתמש בפקודות הפרוקסי שבשלב 3 |

אם התקלה לא מופיעה כאן, יש לשלוח לאחראי הכלי צילום מסך של **כל** הודעת השגיאה.