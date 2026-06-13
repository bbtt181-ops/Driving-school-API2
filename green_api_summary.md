# דרייב פרו — סיכום שיחה | Green API + WhatsApp תזכורות
**06 יוני 2026 | ברוך תור — בית ספר לנהיגה**

---

## מה בוצע בשיחה זו

### Green API — מחובר ופועל ✅
- הוחלף CallMeBot ב-Green API
- שדה "WhatsApp Key" הוסר מכרטיס תלמיד
- תמונה שבועית מתחלפת מקולקציית Driving של Unsplash
- Trigger פעיל — כל 15 דקות

### פרטי Green API
| פריט | ערך |
|---|---|
| idInstance | `7107644824` |
| apiTokenInstance | `5b3c5b0ac7d3476db19afbef1149d81baa291dd9ff0742bd81` |
| מספר מחובר | 972544566181 |
| תוכנית | Pro (בתשלום) — ללא הגבלת הודעות |
| console | https://console.green-api.com |

### Script Properties שמוגדרות ב-GAS
| Key | Value |
|---|---|
| `GREEN_API_INSTANCE` | `7107644824` |
| `GREEN_API_TOKEN` | `5b3c5b0ac7d3476db19afbef1149d81baa291dd9ff0742bd81` |

---

## לוגיקת שליחת תזכורות (sendLessonReminders)
| פרמטר | ערך |
|---|---|
| תדירות | כל 15 דקות (GAS trigger) |
| חלון זמן | 25–35 דקות לפני שיעור |
| סינון שעה | רק שיעורים מ-07:40 ואילך |
| סטטוס | מתוכנן בלבד |
| כפילות | waSent=true אחרי שליחה |
| תמונה | מתחלפת שבועית מ-Unsplash collection 1461149 |

### נוסח ההודעה
```
היי [שם] מה קורה 😊
תזכורת - יש לנו שיעור נהיגה בשעה [HH:MM], אוקי?
```
(נשלחת כ-caption על תמונת נהיגה)

---

## פרטי פרויקט לשיחה הבאה
| פריט | ערך |
|---|---|
| קובץ אפ | `driving-school-pro.html` |
| קובץ API | `driving-school-api.gs` |
| Spreadsheet ID | `1pN5Hs5t2PkwPKgq1gHFFFC3Zn0VZi-RQ8XO4NOkdaEA` |
| Script ID | `1WMEPlJRmQoGE1htS2WXWcYzb7232nxYCqpnJvTMqSL57gnIgPiNnzgWj` |
| API URL | `https://script.google.com/macros/s/AKfycby2HYv4njQ19wfXdrGR2tD3ThkHKdQNYbVAbRXXwrRF7G_cTNwwyjMaH9lgnVI5p696TQ/exec` |
| GitHub Repo | `bbtt181-ops/Driving-school-pro` |
| תיקייה | `C:\Users\PC\Desktop\cloude_workouts\driving-school-code2` |
| Green API Deployment | `@4` |

---

## שלב הבא — בדיקת Meta WhatsApp API (שיחה עתידית)
ראה מסמך: `whatsapp_options_summary.md`

**סיכום: Meta API לא מומלץ** — דורש מספר ייעודי + עלות לכל הודעה עסקית.
המלצה: שדרג Green API ל-Business ($12/חודש) לשימוש מלא.

---
*מסמך זה הופק אוטומטית — שמור ב-driving-school-code2\*
