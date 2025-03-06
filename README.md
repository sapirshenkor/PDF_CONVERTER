# Word & PowerPoint to PDF Converter
![לוגו](pdf-to-x.ico)

PDF לפורמט Word ו-PowerPoint  תוכנה פשוטה להמרת קבצי  .

## תכונות

-  PDF-ל DOCX המרת קבצי 
- PDF-ל PPTX המרת קבצי
- ממשק משתמש גרפי נוח וידידותי
- תמיכה בהמרת מספר קבצים בו-זמנית
- הצגת התקדמות ההמרה בזמן אמת

## דרישות מערכת

-  Windows 10 מערכת הפעלה: ומעלה
- מותקנים במחשב Microsoft Office (Word ו-PowerPoint) 


## שימוש
cmd-ב python main.py הרצת התוכנית, נריץ את הפקודה 
1. לחץ על כפתור "בחר קבצי Word / PowerPoint" לבחירת הקבצים להמרה.
2. בחר את התיקייה שבה ברצונך לשמור את קבצי ה-PDF.
3. המתן להשלמת תהליך ההמרה.
4. קבצי ה-PDF ימצאו בתיקייה שבחרת, עם אותם שמות כמו קבצי המקור.

## פיתוח

אם ברצונך לשנות או לפתח את הפרויקט:

1. שכפל את המאגר: `git clone https://github.com/sapirshenkor/PDF_CONVERTER.git`
2. התקן את תלויות הפיתוח: `pip install -r requirements.txt`
3. בצע את השינויים שלך
4. צור מחדש את קובץ ההרצה: `pyinstaller --onefile --windowed main.py`

## תלויות

- tkinter - לממשק המשתמש הגרפי
- pptxtopdf - להמרת קבצי PowerPoint
- docx2pdf - להמרת קבצי Word

