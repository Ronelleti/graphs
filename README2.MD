# בדיקות יומיות (Daily Checks) – מדריך התקנה

הכלי מצלם את מסכי CNOM ואת הדשבורדים של Metabase, מריץ את בדיקות הטלפונים ושולח את התוצאות במייל דרך Brevo.

**כל ההתקנה נעשית ברמת המשתמש שלך – לא צריך הרשאות אדמין.**

---

## מה צריך לפני שמתחילים

- חיבור לרשת החברה / VPN, עם גישה ל-CNOM ול-Metabase
- שם משתמש וסיסמה ל-CNOM ול-Metabase
- מפתח API של Brevo וכתובת שולח מאומתת (לקבל מאחראי הכלי)
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

## שלב 4 – הגדרת משתני סביבה

> ⚠️ פקודות מהסוג `$env:NAME = "..."` תקפות **רק לחלון ה-PowerShell הנוכחי**, ונעלמות כשסוגרים אותו.
> כדי שההגדרות יישמרו לתמיד (למשתמש שלך בלבד וללא אדמין), צריך להשתמש בפקודות שבהמשך.

יש להחליף את הערכים בפרטים שלך ולהריץ ב-PowerShell:

```powershell
[Environment]::SetEnvironmentVariable('CNOM_USERNAME',     'your_cnom_username',     'User')
[Environment]::SetEnvironmentVariable('CNOM_PASSWORD',     'your_cnom_password',     'User')
[Environment]::SetEnvironmentVariable('METABASE_USERNAME', 'your_metabase_username', 'User')
[Environment]::SetEnvironmentVariable('METABASE_PASSWORD', 'your_metabase_password', 'User')
[Environment]::SetEnvironmentVariable('BREVO_API_KEY',     'your_brevo_api_key',     'User')
[Environment]::SetEnvironmentVariable('EMAIL_FROM',        'your_verified_sender@address.com', 'User')
[Environment]::SetEnvironmentVariable('EMAIL_TO',          'recipient1@company.com,recipient2@company.com', 'User')
```

שימו לב:

- יש להשתמש **בגרש בודד** `' '` ולא במירכאות `" "`. אם בסיסמה יש תו `$`, מירכאות כפולות ישבשו אותה.
- ב-`EMAIL_TO` מפרידים בין הנמענים בפסיק, **בלי רווחים**.
- `EMAIL_FROM` חייבת להיות כתובת שאומתה ב-Brevo.

**אחרי ההגדרה:** לסגור את כל חלונות ה-PowerShell ולפתוח חלון חדש, ואז לבדוק שהמשתנים קיימים (מוצגים רק השמות, בלי הסיסמאות):

```powershell
Get-ChildItem Env: | Where-Object { $_.Name -match 'CNOM|METABASE|BREVO|EMAIL' } | Select-Object Name
```

אמורות להופיע 7 שורות.

**דרך חלופית בלי פקודות:** תפריט התחל ← לחפש "ערוך משתני סביבה עבור החשבון שלך" (Edit environment variables for your account) ← New, ולהוסיף כל משתנה בנפרד. גם זה לא דורש אדמין.

### 🔒 אבטחה

- כל משתמש מגדיר את **פרטי ההתחברות שלו**. לא משתפים סיסמאות.
- לא כותבים סיסמאות או מפתח API בתוך הקוד, ולא מעלים אותם ל-Git.

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

בסיום הריצה נשלח מייל עם התוצאות לנמענים שהוגדרו ב-`EMAIL_TO`.

> בהרצה הראשונה של ה-exe ייתכן ש-Windows SmartScreen יחסום אותו. במקרה כזה יש ללחוץ **More info** ואז **Run anyway**.

---

## הרצה אוטומטית כל בוקר (אופציונלי, ללא אדמין)

1. תפריט התחל ← **Task Scheduler** (מתזמן המשימות)
2. **Create Basic Task** ← לתת שם, למשל `Daily Checks`
3. Trigger: **Daily**, ולבחור שעה
4. Action: **Start a program**
   - **Program/script:** `C:\dailychecks\allinone.exe`
   - **Start in:** `C:\dailychecks` ← **חובה למלא**, אחרת הכלי לא ימצא את תיקיית `phonechecks`
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
| משתנה סביבה ריק / `KeyError` | החלון נפתח לפני ההגדרה, או שהמשתנה לא הוגדר | לפתוח PowerShell חדש ולחזור על הבדיקה בשלב 4 |
| שגיאת `401` מ-Brevo | מפתח API שגוי | לבדוק את `BREVO_API_KEY` |
| המייל לא נשלח / שגיאת sender | כתובת השולח לא מאומתת ב-Brevo | לבדוק את `EMAIL_FROM` |
| שגיאת SSL או חיבור ב-pip | פרוקסי של החברה | להשתמש בפקודות הפרוקסי שבשלב 3 |

אם התקלה לא מופיעה כאן, יש לשלוח לאחראי הכלי צילום מסך של **כל** הודעת השגיאה.