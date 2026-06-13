# דרייב פרו — השוואת אפשרויות WhatsApp לתזכורות שיעורים
**06 יוני 2026 | ברוך תור — בית ספר לנהיגה**

---

## מצב נוכחי

Green API **מחובר ופועל** (Instance 7107644824, מספר: 972544566181).  
הפלן הנוכחי: **Pro (בתשלום) — ללא הגבלת הודעות.**  
השדרוג בוצע — ניתן לשלוח תזכורות לכל התלמידים ללא הגבלה.

---

## השוואת אפשרויות

| אפשרות | עלות | שולח מהמספר שלך? | קל להגדרה? | מגבלות |
|---|---|---|---|---|
| **Green API Pro** ✅ (בשימוש) | בתשלום | ✅ כן | ✅ כבר מוכן | ללא הגבלה |
| **Meta WhatsApp Cloud API** | ❌ לא חינם | ❌ מספר ייעודי בלבד | ❌ מסובך מאוד | ראה פירוט למטה |
| **CallMeBot** | חינם | ❌ לא — שולח מהמספר שלהם | ✅ פשוט | כל תלמיד חייב לאשר בעצמו |

---

## מדוע Meta API **לא מתאים** לפרויקט זה

### 1. לא חינם לתזכורות
- מאז **יולי 2025** — Meta עברה לתמחור **לפי הודעה** (per-message)
- תזכורות שיעור = "Utility messages" = **עלות לכל הודעה**
- ה-1,000 השיחות החינמיות = שיחות שהתלמיד **פתח** (לא הודעות שאתה שולח)

### 2. לא ניתן להשתמש במספר האישי שלך
- Meta דורש **מספר ייעודי** שאינו משויך לוואטסאפ אישי/עסקי
- כלומר: צריך להחליף מספר או לקנות מספר נוסף

### 3. תהליך הגדרה מורכב
- Facebook Business Manager מאומת
- Meta App + WhatsApp Business Account
- אישור תבניות הודעה (Template approval) — תהליך של ימים
- Webhook לאישור מסירה

---

## המלצה

**הושלם** — Green API שודרג לתוכנית Pro (בתשלום), ללא הגבלת הודעות.

**למה זה טוב:**
- כבר מוכן ועובד
- שולח מהמספר האישי שלך (072/050/052...)
- התלמיד לא צריך לעשות כלום
- ללא הגבלת הודעות — מתאים לכל התלמידים

---

## פרטי הפרויקט לשיחה הבאה

| פריט | ערך |
|---|---|
| קובץ אפ | `driving-school-pro.html` |
| קובץ API | `driving-school-api.gs` |
| Spreadsheet ID | `1pN5Hs5t2PkwPKgq1gHFFFC3Zn0VZi-RQ8XO4NOkdaEA` |
| Script ID | `1WMEPlJRmQoGE1htS2WXWcYzb7232nxYCqpnJvTMqSL57gnIgPiNnzgWj` |
| API URL | `https://script.google.com/macros/s/AKfycby2HYv4njQ19wfXdrGR2tD3ThkHKdQNYbVAbRXXwrRF7G_cTNwwyjMaH9lgnVI5p696TQ/exec` |
| GitHub Repo | `bbtt181-ops/Driving-school-pro` |
| Green API Instance | `7107644824` |
| Green API Token | `5b3c5b0ac7d3476db19afbef1149d81baa291dd9ff0742bd81` |
| תיקייה | `C:\Users\PC\Desktop\cloude_workouts\driving-school-code2` |

---

*מסמך זה הופק אוטומטית — שמור ב-driving-school-code2\*
