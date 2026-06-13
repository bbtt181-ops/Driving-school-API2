// ============================================================
// driving-school-api.gs — REST API for דרייב פרו
// גרסה 2.0 | מאי 2026
//
// הוראות פריסה:
// 1. פתח script.google.com → פרויקט חדש
// 2. הדבק קוד זה
// 3. Deploy → New deployment → Web App
//    Execute as: Me | Who has access: Anyone
// 4. העתק את ה-URL והדבק ב-⚙️ באפ
// ============================================================

const SPREADSHEET_ID = '1v-7JPq9huUjZYaqPFn05JRnlmlJt7I3886H_bZuAN7Q';
const SPREADSHEET_ID_OCCASIONAL = '13AQgLak3XyrwK_GksM49QiBYwuXVb0qhWurWjDdODIs';
const SHEET_STUDENTS = 'Students';
const SHEET_LESSONS  = 'Lessons';
const SHEET_PAYMENTS = 'Payments';

// ─── doGet — קריאות GET ────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'getAll';
  let result;
  try {
    if (action === 'getAll') {
      result = getAllData();
    } else if (action === 'getCalendarEvents') {
      result = getCalendarEvents(
        e.parameter.name || '',
        e.parameter.from || '',
        e.parameter.to   || ''
      );
    } else if (action === 'checkTriggers') {
      result = checkTriggers();
    } else {
      result = { ok: false, msg: 'פעולה לא מוכרת: ' + action };
    }
  } catch(err) {
    result = { ok: false, msg: err.toString() };
  }
  return _json(result);
}

// ─── doPost — כתיבה / מחיקה ────────────────────
// ─── שליחת מייל לתיבת דואר נכנס דרך Gmail API ──────
// משתמש ב-messages.insert כדי להכניס ישירות ל-INBOX (לא "לשלוח")
function sendEmailAction(to, subject, htmlBody) {
  if (!to || !subject || !htmlBody) return { ok: false, msg: 'חסרים שדות: to/subject/htmlBody' };
  try {
    const from = Session.getActiveUser().getEmail();
    const subjectEncoded = '=?UTF-8?B?' + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + '?=';
    const mimeMsg = [
      'From: ' + from,
      'To: ' + to,
      'Subject: ' + subjectEncoded,
      'MIME-Version: 1.0',
      'Content-Type: text/html; charset=UTF-8',
      '',
      htmlBody
    ].join('\r\n');

    const encoded = Utilities.base64EncodeWebSafe(mimeMsg);
    const token   = ScriptApp.getOAuthToken();

    const resp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/messages?internalDateSource=dateHeader',
      {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ raw: encoded, labelIds: ['INBOX'] }),
        muteHttpExceptions: true
      }
    );
    const result = JSON.parse(resp.getContentText());
    if (result.id) return { ok: true, sent: true, msgId: result.id };
    return { ok: false, msg: JSON.stringify(result) };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if      (action === 'saveAll')       result = saveAll(body.data);
    else if (action === 'save')          result = saveOne(body.type, body.data);
    else if (action === 'delete')        result = deleteOne(body.type, body.data);
    else if (action === 'deleteStudent') result = deleteStudent(body.data.id);
    else if (action === 'clearAll')      result = clearAll();
    else if (action === 'sendEmail')     result = sendEmailAction(body.to, body.subject, body.htmlBody);
    else if (action === 'setupReminders')   result = setupAllRemindersAction();
    else if (action === 'runReminders')     result = runRemindersNow();
    else result = { ok: false, msg: 'פעולה לא מוכרת: ' + action };
  } catch(err) {
    result = { ok: false, msg: err.toString() };
  }
  return _json(result);
}

// ─── קריאת כל הנתונים ──────────────────────────
function getAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    ok: true,
    data: {
      students: _readSheet(ss, SHEET_STUDENTS),
      lessons:  _readSheet(ss, SHEET_LESSONS),
      payments: _readSheet(ss, SHEET_PAYMENTS)
    }
  };
}

// ─── שמירת הכל (saveAll) ───────────────────────
function saveAll(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (data && data.students) _writeSheet(ss, SHEET_STUDENTS, data.students);
  if (data && data.lessons)  _writeSheet(ss, SHEET_LESSONS,  data.lessons);
  if (data && data.payments) _writeSheet(ss, SHEET_PAYMENTS, data.payments);
  return { ok: true };
}

// ─── שמירת רשומה בודדת ─────────────────────────
function saveOne(type, data) {
  if (!data || !data.id) return { ok: false, msg: 'חסר id' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shName = _sheetName(type);
  if (!shName) return { ok: false, msg: 'סוג לא מוכר: ' + type };

  const sh = _ensureSheet(ss, shName);
  const lastRow = sh.getLastRow();

  // חפש שורה קיימת ועדכן
  if (lastRow >= 2) {
    const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (!vals[i][0]) continue;
      try {
        const item = JSON.parse(vals[i][0]);
        if (item.id === data.id) {
          sh.getRange(i + 2, 1).setValue(JSON.stringify(data));
          return { ok: true };
        }
      } catch(e) {}
    }
  }
  // לא נמצא — הוסף שורה חדשה
  sh.appendRow([JSON.stringify(data)]);
  return { ok: true };
}

// ─── מחיקת רשומה בודדת ─────────────────────────
function deleteOne(type, data) {
  if (!data || !data.id) return { ok: true };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shName = _sheetName(type);
  if (!shName) return { ok: false, msg: 'סוג לא מוכר' };
  _deleteById(ss, shName, data.id, 'id');
  return { ok: true };
}

// ─── מחיקת תלמיד + כל הנתונים שלו ────────────
function deleteStudent(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  _deleteById(ss, SHEET_STUDENTS, id, 'id');
  _deleteById(ss, SHEET_LESSONS,  id, 'studentId');
  _deleteById(ss, SHEET_PAYMENTS, id, 'studentId');
  return { ok: true };
}

// ─── מחיקת כל הנתונים ──────────────────────────
function clearAll() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  [SHEET_STUDENTS, SHEET_LESSONS, SHEET_PAYMENTS].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
    }
  });
  return { ok: true };
}

// ─── שליפת אירועי יומן (מזדמן) ─────────────────
// צבעים:
//   כחול (9=Blueberry / 7=Peacock / ''=ברירת מחדל) → שיעור רגיל
//   צהוב (5=Banana) → טסט, חיוב ₪230 קבוע ללא חיוב לפי זמן
//   כל שאר הצבעים (אדום וכו') → מתעלם
function getCalendarEvents(name, from, to) {
  if (!name) return { ok: false, msg: 'שם חסר' };

  const fromDate = from ? new Date(from) : new Date(new Date().setFullYear(new Date().getFullYear() - 1));
  const toDate   = to   ? new Date(to)   : new Date();
  toDate.setHours(23, 59, 59);

  // טען שמות תלמידים לצורך ולידציה — מהגיליון המזדמנים
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_OCCASIONAL);
  const allStudents = _readSheet(ss, SHEET_STUDENTS);
  const studentNames = allStudents.map(s => (s.name || '').trim()).filter(n => n.length > 1);

  const lessons  = [];
  const payments = [];
  const tz = Session.getScriptTimeZone();
  const nameLower = name.toLowerCase();

  CalendarApp.getAllCalendars().forEach(cal => {
    cal.getEvents(fromDate, toDate).forEach(ev => {
      const title = ev.getTitle() || '';
      const t = title.toLowerCase();

      // חייב להכיל את שם התלמיד המבוקש
      if (!t.includes(nameLower)) return;

      // חייב להיות תלמיד שלי
      const isMyStudent = studentNames.some(sn => t.includes(sn.toLowerCase()));
      if (!isMyStudent) return;

      const color = String(ev.getColor() || '');
      const isBlue   = color === '' || color === '9' || color === '7' ||
                       color === 'BLUE' || color === 'CYAN';
      const isYellow = color === '5' || color === 'YELLOW' || color === 'BANANA';

      // מתעלם מכל צבע אחר (אדום, ירוק, סגול וכו')
      if (!isBlue && !isYellow) return;

      const start   = ev.getStartTime();
      const end     = ev.getEndTime();
      const dateStr = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
      const timeStr = Utilities.formatDate(start, tz, 'HH:mm');
      const desc    = (ev.getDescription() || '').toLowerCase();

      // ─── צהוב = טסט ───────────────────────────────
      if (isYellow) {
        // מצא את כל התלמידים בכותרת (תלמיד אחד או שניים)
        const matchedStudents = studentNames.filter(sn => t.includes(sn.toLowerCase()));

        matchedStudents.forEach(sn => {
          // רק אם השם המבוקש מופיע (לא מייבאים תלמידים שלא ביקשנו)
          if (sn.toLowerCase() !== nameLower && !nameLower.includes(sn.toLowerCase()) && !sn.toLowerCase().includes(nameLower)) return;
          lessons.push({
            date: dateStr, time: timeStr,
            type: 'test',
            duration: 1,        // לטסט אין חיוב לפי זמן — רק ₪230
            notes: '',
            freeInternal: false
          });
        });
        return; // לא בודקים תשלום/פנימי לאירוע צהוב
      }

      // ─── כחול = שיעור רגיל ────────────────────────
      const durationMins = (end - start) / 60000;
      const duration = Math.max(1, Math.round(durationMins / 40 * 10) / 10);

      let type = 'lesson';
      if      (t.includes('פנימי') || desc.includes('פנימי')) type = 'internal';
      else if (t.includes('חניות') || desc.includes('חניות')) type = 'parking';

      const freeInternal = type === 'internal' && (t.includes('חינם') || t.includes('נוסף'));

      lessons.push({ date: dateStr, time: timeStr, type, duration, notes: '', freeInternal });

      // פרסינג תשלום מהכותרת: "340/200201" = ₪340 קבלה 200201
      const payMatch = title.match(/(\d+)\/(\d{4,7})/);
      if (payMatch) {
        payments.push({
          date: dateStr,
          amount: parseInt(payMatch[1]),
          receiptNum: payMatch[2],
          method: 'cash',
          notes: 'ייבוא יומן'
        });
      }
    });
  });

  return { ok: true, data: { lessons, payments } };
}

// ─── עזרים פנימיים ─────────────────────────────

function _sheetName(type) {
  if (type === 'Students' || type === 'student') return SHEET_STUDENTS;
  if (type === 'Lessons'  || type === 'lesson')  return SHEET_LESSONS;
  if (type === 'Payments' || type === 'payment') return SHEET_PAYMENTS;
  return null;
}

function _ensureSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.setRightToLeft(true);
    sh.getRange(1, 1).setValue('data').setFontWeight('bold');
  }
  return sh;
}

function _readSheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const result = [];
  vals.forEach(row => {
    if (!row[0]) return;
    try { result.push(JSON.parse(row[0])); } catch(e) {}
  });
  return result;
}

function _writeSheet(ss, name, items) {
  const sh = _ensureSheet(ss, name);
  // מחק שורות נתונים ישנות
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
  }
  if (!items || items.length === 0) return;
  const rows = items.map(item => [JSON.stringify(item)]);
  sh.getRange(2, 1, rows.length, 1).setValues(rows);
}

// ============================================================
// WhatsApp Reminders — Green API
// הגדרה:
//   1. כנס ל-green-api.com → צור Instance → סרוק QR עם הוואטסאפ שלך
//   2. ב-Apps Script editor: Project Settings → Script Properties → הוסף:
//      GREEN_API_INSTANCE  = מספר ה-Instance שלך (למשל: 1234567890)
//      GREEN_API_TOKEN     = ה-apiTokenInstance שלך
//   3. הרץ setupWhatsAppTrigger() פעם אחת מה-editor
// ============================================================

/**
 * בדיקת Green API — הרץ פעם אחת מה-editor לבדיקה
 */
function testGreenApi() {
  const props  = PropertiesService.getScriptProperties();
  const instId = props.getProperty('GREEN_API_INSTANCE');
  const token  = props.getProperty('GREEN_API_TOKEN');
  Logger.log('Instance: ' + instId);
  Logger.log('Token: ' + token);

  const imgUrl  = _weeklyImage(Math.floor(Math.random() * 1000));
  Logger.log('תמונה שבועית: ' + imgUrl);

  const url = 'https://api.green-api.com/waInstance' + instId + '/sendFileByUrl/' + token;
  const payload = JSON.stringify({
    chatId:   '972587045405@c.us',
    urlFile:  imgUrl,
    fileName: 'reminder.jpg',
    caption:  'בדיקה מדרייב פרו 🚗\nתזכורת - יש לנו שיעור נהיגה בשעה 09:00, אוקי?'
  });
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  });
  Logger.log('Status: ' + resp.getResponseCode());
  Logger.log('Response: ' + resp.getContentText());
}

/**
 * הגדרת Trigger חד פעמית — הרץ רק פעם אחת מה-editor
 */
function setupWhatsAppTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendLessonReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendLessonReminders')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Trigger הוגדר — sendLessonReminders כל 15 דקות');
}

// ============================================================
// תזכורות שיעורים — sendLessonReminders
// גרסה 3.0 | יוני 2026
//
// שינויים מגרסה קודמת:
//   • חלון זמן: 50–70 דקות לפני שיעור (במקום 5–65)
//   • תיקון תאריך מחר: תמיד today+1, לא תלוי ב-winEnd
//   • waSent נשמר ב-Sheets לפני שליחה (מונע כפילות)
//   • לוג מפורט על כל שיעור שלא נשלח + סיבה
//   • סינון: שיעורים לפני 07:40 לא מקבלים תזכורת
// ============================================================

/**
 * שליחת תזכורות WhatsApp דרך Green API — רץ כל 15 דקות
 * שולח לשיעורים שמתחילים בעוד 50–70 דקות
 * רק שיעורים מ-07:40 ואילך
 */
function sendLessonReminders() {
  const props  = PropertiesService.getScriptProperties();
  const instId = props.getProperty('GREEN_API_INSTANCE');
  const token  = props.getProperty('GREEN_API_TOKEN');
  if (!instId || !token) {
    Logger.log('❌ חסרים GREEN_API_INSTANCE / GREEN_API_TOKEN ב-Script Properties');
    return;
  }

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz  = 'Asia/Jerusalem';
  const now = new Date();

  // תאריכים — today ו-tomorrow מחושבים בצורה מדויקת (לא תלוי ב-winEnd)
  const todayStr    = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const tomorrowD   = new Date(now.getTime() + 24 * 60 * 60000);
  const tomorrowStr = Utilities.formatDate(tomorrowD, tz, 'yyyy-MM-dd');

  // חלון: 50–70 דקות קדימה
  const winStart = new Date(now.getTime() + 50 * 60000);
  const winEnd   = new Date(now.getTime() + 70 * 60000);

  const students = _readSheet(ss, SHEET_STUDENTS);
  const lessons  = _readSheet(ss, SHEET_LESSONS);
  const baseUrl  = 'https://api.green-api.com/waInstance' + instId;

  let sentCount  = 0;
  let skipCount  = 0;

  lessons.forEach(lesson => {

    // ── סינון בסיסי ──────────────────────────────────────
    if (lesson.date !== todayStr && lesson.date !== tomorrowStr) return;
    if (lesson.status !== 'scheduled') {
      Logger.log('⏭️ דילוג [' + (lesson.date||'') + ' ' + (lesson.time||'') + '] — סטטוס: ' + lesson.status);
      skipCount++; return;
    }
    if (lesson.type !== 'lesson' && lesson.type !== 'internal') {
      // טסט מטופל בנפרד ב-sendTestReminders
      return;
    }
    if (lesson.waSent) {
      Logger.log('⏭️ כבר נשלח [' + lesson.date + ' ' + lesson.time + ']');
      skipCount++; return;
    }

    // ── בדיקת שעה ────────────────────────────────────────
    const timeParts = (lesson.time || '').split(':');
    if (timeParts.length < 2) {
      Logger.log('⚠️ שעה חסרה/שגויה [id=' + lesson.id + ']');
      skipCount++; return;
    }
    const [lh, lm] = timeParts.map(Number);
    if (isNaN(lh) || isNaN(lm)) {
      Logger.log('⚠️ שעה לא תקינה [id=' + lesson.id + '] time=' + lesson.time);
      skipCount++; return;
    }

    // סינון שיעורים לפני 07:40
    const lessonMin = lh * 60 + lm;
    if (lessonMin < 0) {
      Logger.log('⏭️ שיעור לפני 07:40 [' + lesson.date + ' ' + lesson.time + '] — לא שולח');
      skipCount++; return;
    }

    // ── בדיקת חלון זמן ───────────────────────────────────
    const [ly, lmo, ld] = lesson.date.split('-').map(Number);
    const lessonTime = new Date(ly, lmo - 1, ld, lh, lm, 0);
    if (lessonTime < winStart || lessonTime > winEnd) return; // לא בחלון — לא לוגים (יתפוס בריצה אחרת)

    // ── מציאת תלמיד ──────────────────────────────────────
    const student = students.find(s => s.id === lesson.studentId);
    if (!student) {
      Logger.log('❌ תלמיד לא נמצא [lessonId=' + lesson.id + ' studentId=' + lesson.studentId + ']');
      skipCount++; return;
    }
    if (!student.phone) {
      Logger.log('❌ אין טלפון לתלמיד ' + student.name + ' [id=' + student.id + ']');
      skipCount++; return;
    }

    const phone = normalizePhone(String(student.phone));
    if (!phone) {
      Logger.log('❌ מספר טלפון לא תקין: ' + student.phone + ' (תלמיד: ' + student.name + ')');
      skipCount++; return;
    }

    // ── סימון waSent לפני שליחה (מונע כפילות גם אם ה-fetch איטי) ──
    lesson.waSent = true;
    try {
      _upsertById(ss, SHEET_LESSONS, lesson);
    } catch(saveErr) {
      Logger.log('⚠️ לא הצלחתי לשמור waSent לפני שליחה: ' + saveErr.toString());
      // ממשיכים — עדיף לשלוח פעמיים מאשר לא לשלוח
    }

    // ── נוסח הודעה ───────────────────────────────────────
    const firstName = student.name.split(' ')[0];
    const caption   = 'היי ' + firstName + ' מה קורה 😊\n' +
                      'תזכורת - יש לנו שיעור נהיגה בשעה ' + lesson.time + ', אוקי? 🚗';

    const imgUrl  = _weeklyImage(Math.floor(Math.random() * 1000));
    const chatId  = phone + '@c.us';
    let   sent    = false;
    let   failReason = '';

    // ── ניסיון 1: תמונה + כיתוב ──────────────────────────
    try {
      const imgResp = UrlFetchApp.fetch(baseUrl + '/sendFileByUrl/' + token, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chatId: chatId, urlFile: imgUrl, fileName: 'reminder.jpg', caption: caption }),
        muteHttpExceptions: true
      });
      const imgBody = JSON.parse(imgResp.getContentText());
      if (imgBody && imgBody.idMessage) {
        Logger.log('✅ תמונה נשלחה ל-' + student.name + ' (' + phone + ') | שיעור: ' + lesson.date + ' ' + lesson.time);
        sent = true;
      } else {
        failReason = 'תמונה נכשלה: ' + imgResp.getContentText();
        Logger.log('⚠️ ' + failReason + ' — עובר לטקסט');
      }
    } catch(e) {
      failReason = 'שגיאת תמונה: ' + e.toString();
      Logger.log('⚠️ ' + failReason);
    }

    // ── ניסיון 2 (fallback): טקסט בלבד ──────────────────
    if (!sent) {
      try {
        const txtResp = UrlFetchApp.fetch(baseUrl + '/sendMessage/' + token, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ chatId: chatId, message: caption }),
          muteHttpExceptions: true
        });
        const txtBody = JSON.parse(txtResp.getContentText());
        if (txtBody && txtBody.idMessage) {
          Logger.log('✅ טקסט נשלח ל-' + student.name + ' (' + phone + ') | שיעור: ' + lesson.date + ' ' + lesson.time);
          sent = true;
        } else {
          failReason += ' | טקסט נכשל: ' + txtResp.getContentText();
          Logger.log('❌ ' + failReason);
        }
      } catch(e2) {
        failReason += ' | שגיאת טקסט: ' + e2.toString();
        Logger.log('❌ גם טקסט נכשל ל-' + student.name + ': ' + e2.toString());
      }
    }

    // ── אם נכשל לחלוטין — בטל את waSent ──────────────────
    if (!sent) {
      lesson.waSent = false;
      try { _upsertById(ss, SHEET_LESSONS, lesson); } catch(e) {}
      Logger.log('❌ שליחה נכשלה לחלוטין ל-' + student.name + ' | סיבה: ' + failReason);
      skipCount++;
      return;
    }

    sentCount++;

    // ── עותק לבעל בית הספר כל 15 הודעות ─────────────────
    const sendProps = PropertiesService.getScriptProperties();
    const count = parseInt(sendProps.getProperty('WA_SEND_COUNT') || '0') + 1;
    sendProps.setProperty('WA_SEND_COUNT', String(count));

    if (count % 15 === 0) {
      const ownerPhone   = '972544566181';
      const ownerCaption = '📋 עותק הודעה #' + count + ' ששלחנו:\n' + caption;
      try {
        UrlFetchApp.fetch(baseUrl + '/sendFileByUrl/' + token, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ chatId: ownerPhone + '@c.us', urlFile: imgUrl, fileName: 'reminder.jpg', caption: ownerCaption }),
          muteHttpExceptions: true
        });
        Logger.log('📋 עותק #' + count + ' נשלח לבעל בית הספר');
      } catch(e) {
        Logger.log('⚠️ עותק לבעל נכשל: ' + e.toString());
      }
    }
  });

  Logger.log('📊 סיכום ריצה: נשלחו=' + sentCount + ' | דולגו/נכשלו=' + skipCount);
}

// ============================================================
// תזכורות טסט — sendTestReminders
// יוצא 3 ימים לפני הטסט ושוב יום לפני
// ============================================================

/**
 * שליחת תזכורות טסט — רץ כל 15 דקות (אותו trigger כמו sendLessonReminders)
 * תזכורת 1: 3 ימים לפני הטסט (נשלחת פעם אחת בין 08:00–09:00)
 * תזכורת 2: יום לפני הטסט (נשלחת פעם אחת בין 08:00–09:00)
 */
function sendTestReminders() {
  const props  = PropertiesService.getScriptProperties();
  const instId = props.getProperty('GREEN_API_INSTANCE');
  const token  = props.getProperty('GREEN_API_TOKEN');
  if (!instId || !token) {
    Logger.log('❌ חסרים GREEN_API_INSTANCE / GREEN_API_TOKEN');
    return;
  }

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz  = 'Asia/Jerusalem';
  const now = new Date();

  // שולחים תזכורות טסט רק בין 08:00–09:00
  const nowHour = parseInt(Utilities.formatDate(now, tz, 'HH'));
  if (nowHour < 8 || nowHour >= 9) return;

  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  // חישוב תאריכים: בעוד 3 ימים ובעוד יום אחד
  const in3Days = new Date(now.getTime() + 3 * 24 * 60 * 60000);
  const in1Day  = new Date(now.getTime() + 1 * 24 * 60 * 60000);
  const in3Str  = Utilities.formatDate(in3Days, tz, 'yyyy-MM-dd');
  const in1Str  = Utilities.formatDate(in1Day,  tz, 'yyyy-MM-dd');

  const students = _readSheet(ss, SHEET_STUDENTS);
  const lessons  = _readSheet(ss, SHEET_LESSONS);
  const baseUrl  = 'https://api.green-api.com/waInstance' + instId;

  // פורמט תאריך נחמד לעברית: DD/MM/YYYY
  function fmtDate(dateStr) {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    return parts[2] + '/' + parts[1] + '/' + parts[0];
  }

  lessons.forEach(lesson => {
    if (lesson.type !== 'test') return;
    if (lesson.status !== 'scheduled') return;

    const isIn3Days = lesson.date === in3Str;
    const isIn1Day  = lesson.date === in1Str;
    if (!isIn3Days && !isIn1Day) return;

    // בדיקת דגל שליחה לפי סוג תזכורת
    const sentFlag = isIn3Days ? 'waSent3' : 'waSent1';
    if (lesson[sentFlag]) {
      Logger.log('⏭️ תזכורת טסט כבר נשלחה [' + sentFlag + '] לשיעור ' + lesson.date);
      return;
    }

    const student = students.find(s => s.id === lesson.studentId);
    if (!student) {
      Logger.log('❌ תלמיד לא נמצא לטסט [lessonId=' + lesson.id + ']');
      return;
    }
    if (!student.phone) {
      Logger.log('❌ אין טלפון לתלמיד ' + student.name + ' לטסט');
      return;
    }

    const phone = normalizePhone(String(student.phone));
    if (!phone) {
      Logger.log('❌ טלפון לא תקין: ' + student.phone + ' (תלמיד: ' + student.name + ')');
      return;
    }

    const firstName  = student.name.split(' ')[0];
    const testDate   = fmtDate(lesson.date);
    const testTime   = lesson.time || '';

    // ── נוסח הודעה לפי סוג תזכורת ───────────────────────
    let caption;
    if (isIn3Days) {
      caption =
        'היי ' + firstName + ' 😊\n' +
        'עוד 3 ימים, ב-' + testDate + ' בשעה ' + testTime + ' יש לנו טסט נהיגה! 🚗\n\n' +
        'כדאי להתחיל להתכונן 💪\n' +
        'שיהיה לנו בהצלחה! 🍀';
    } else {
      caption =
        'היי ' + firstName + ' מה קורה 😊\n' +
        'מחר ' + testDate + ' בשעה ' + testTime + ' יש לנו טסט נהיגה! 🎯\n\n' +
        'יש להצטייד ב:\n' +
        '✅ תעודת זהות\n' +
        '👓 משקפיים (אם צריך)\n' +
        '💪 הרבה ביטחון!\n\n' +
        'אנחנו מוכנים — שיהיה לנו בהצלחה! 🏆🍀';
    }

    // סימון לפני שליחה
    lesson[sentFlag] = true;
    try {
      _upsertById(ss, SHEET_LESSONS, lesson);
    } catch(saveErr) {
      Logger.log('⚠️ לא הצלחתי לשמור ' + sentFlag + ' לפני שליחה: ' + saveErr.toString());
    }

    const imgUrl = _weeklyImage(Math.floor(Math.random() * 1000));
    const chatId = phone + '@c.us';
    let sent     = false;
    let failReason = '';

    // ניסיון 1: תמונה
    try {
      const imgResp = UrlFetchApp.fetch(baseUrl + '/sendFileByUrl/' + token, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chatId: chatId, urlFile: imgUrl, fileName: 'test-reminder.jpg', caption: caption }),
        muteHttpExceptions: true
      });
      const imgBody = JSON.parse(imgResp.getContentText());
      if (imgBody && imgBody.idMessage) {
        Logger.log('✅ תזכורת טסט (' + (isIn3Days ? '3 ימים' : 'יום לפני') + ') נשלחה ל-' + student.name);
        sent = true;
      } else {
        failReason = 'תמונה נכשלה: ' + imgResp.getContentText();
      }
    } catch(e) {
      failReason = 'שגיאת תמונה: ' + e.toString();
    }

    // ניסיון 2: טקסט בלבד
    if (!sent) {
      try {
        const txtResp = UrlFetchApp.fetch(baseUrl + '/sendMessage/' + token, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ chatId: chatId, message: caption }),
          muteHttpExceptions: true
        });
        const txtBody = JSON.parse(txtResp.getContentText());
        if (txtBody && txtBody.idMessage) {
          Logger.log('✅ תזכורת טסט טקסט (' + (isIn3Days ? '3 ימים' : 'יום לפני') + ') נשלחה ל-' + student.name);
          sent = true;
        } else {
          failReason += ' | טקסט נכשל: ' + txtResp.getContentText();
        }
      } catch(e2) {
        failReason += ' | שגיאת טקסט: ' + e2.toString();
      }
    }

    if (!sent) {
      lesson[sentFlag] = false;
      try { _upsertById(ss, SHEET_LESSONS, lesson); } catch(e) {}
      Logger.log('❌ תזכורת טסט נכשלה ל-' + student.name + ' | סיבה: ' + failReason);
    }
  });
}

/**
 * הגדרת Trigger — מפעיל sendLessonReminders + sendTestReminders כל 15 דקות
 * הרץ פעם אחת מה-editor
 */
function setupAllReminderTriggers() {
  // מחק triggers קיימים
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'sendLessonReminders' || fn === 'sendTestReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // צור מחדש
  ScriptApp.newTrigger('sendLessonReminders').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('sendTestReminders').timeBased().everyMinutes(15).create();
  Logger.log('✅ Triggers הוגדרו — sendLessonReminders + sendTestReminders כל 15 דקות');
}

// ─── דיבאג: בדיקת ה-Triggers הקיימים (action=checkTriggers) ────
function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers().map(t => ({
    handler: t.getHandlerFunction(),
    type: String(t.getEventType()),
    source: String(t.getTriggerSource())
  }));
  const props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    triggers: triggers,
    count: triggers.length,
    hasGreenApiInstance: !!props.getProperty('GREEN_API_INSTANCE'),
    hasGreenApiToken: !!props.getProperty('GREEN_API_TOKEN')
  };
}

// ─── דיבאג: הגדרת ה-Triggers דרך ה-API (action=setupReminders, POST) ────
function setupAllRemindersAction() {
  setupAllReminderTriggers();
  return checkTriggers();
}

// ─── דיבאג: הרצת התזכורות עכשיו ידנית (action=runReminders, POST) ────
// שימוש לבדיקה בלבד — מריץ את שתי הפונקציות ומחזיר את הלוג
function runRemindersNow() {
  Logger.clear ? Logger.clear() : null;
  let lessonErr = '', testErr = '';
  try { sendLessonReminders(); } catch(e) { lessonErr = e.toString(); }
  try { sendTestReminders(); } catch(e) { testErr = e.toString(); }
  return {
    ok: true,
    log: Logger.getLog(),
    lessonErr: lessonErr,
    testErr: testErr
  };
}

/**
 * בחירת תמונה אקראית לכל התראה — פול 62 תמונות: נהיגה / מכוניות / כבישים / הגה / עיר
 */
function _weeklyImage(weekNum) {
  const images = [
    // — נהיגה ושיעורי נהיגה —
    'https://images.unsplash.com/photo-1756914593851-08ff6adce58a?w=800&q=80',
    'https://images.unsplash.com/photo-1594380660451-376177a138ed?w=800&q=80',
    'https://images.unsplash.com/photo-1701159636551-6c245ea1a7b9?w=800&q=80',
    'https://images.unsplash.com/photo-1568437234382-a8e1e70daa38?w=800&q=80',
    'https://images.unsplash.com/photo-1695230981824-8edd894a6c2c?w=800&q=80',
    'https://images.unsplash.com/photo-1557598628-bd2cb3c9fe11?w=800&q=80',
    'https://images.unsplash.com/photo-1632656269435-77b10f3fcbc6?w=800&q=80',
    'https://images.unsplash.com/photo-1636526423925-4e986022c9e5?w=800&q=80',
    'https://images.unsplash.com/photo-1759339645839-ce0b695c00b5?w=800&q=80',
    'https://images.unsplash.com/photo-1595306360139-abaa9dfb629b?w=800&q=80',
    'https://images.unsplash.com/photo-1727893467393-24bc37e8a117?w=800&q=80',
    'https://images.unsplash.com/photo-1703182571074-f1e2a0925c85?w=800&q=80',
    'https://images.unsplash.com/photo-1700227666565-9d06c7554d70?w=800&q=80',
    'https://images.unsplash.com/photo-1617180705595-58e1bde06b48?w=800&q=80',
    'https://images.unsplash.com/photo-1758525589731-72041b8cb413?w=800&q=80',
    'https://images.unsplash.com/photo-1626895862766-a6c93b6ea807?w=800&q=80',
    'https://images.unsplash.com/photo-1597986228841-4e10852def4e?w=800&q=80',
    'https://images.unsplash.com/photo-1568605117036-5fe5e7bab0b7?w=800&q=80',
    // — מכוניות —
    'https://images.unsplash.com/photo-1494976388531-d1058494cdd8?w=800&q=80',
    'https://images.unsplash.com/photo-1502877338535-766e1452684a?w=800&q=80',
    'https://images.unsplash.com/photo-1511919884226-fd3cad34687c?w=800&q=80',
    'https://images.unsplash.com/photo-1533473359331-0135ef1b58bf?w=800&q=80',
    'https://images.unsplash.com/photo-1520031441872-265e4ff70366?w=800&q=80',
    'https://images.unsplash.com/photo-1542362567-b07e54358753?w=800&q=80',
    'https://images.unsplash.com/photo-1503376780353-7e6692767b70?w=800&q=80',
    'https://images.unsplash.com/photo-1555215695-3004980ad54e?w=800&q=80',
    'https://images.unsplash.com/photo-1580274455191-1c62238fa333?w=800&q=80',
    'https://images.unsplash.com/photo-1544636331-e26879cd4d9b?w=800&q=80',
    'https://images.unsplash.com/photo-1621007947382-bb3c3994e3fb?w=800&q=80',
    'https://images.unsplash.com/photo-1553440569-bcc63803a83d?w=800&q=80',
    'https://images.unsplash.com/photo-1549317661-bd32c8ce0db2?w=800&q=80',
    // — כבישים ונסיעה —
    'https://images.unsplash.com/photo-1500674425229-f692875b0ab7?w=800&q=80',
    'https://images.unsplash.com/photo-1489824904134-891ab64532f1?w=800&q=80',
    'https://images.unsplash.com/photo-1504280390367-361c6d9f38f4?w=800&q=80',
    'https://images.unsplash.com/photo-1519003722824-194d4455a60c?w=800&q=80',
    'https://images.unsplash.com/photo-1473445730015-841f29a9490b?w=800&q=80',
    'https://images.unsplash.com/photo-1464219789935-c2d9d9aba644?w=800&q=80',
    'https://images.unsplash.com/photo-1446776811953-b23d57bd21aa?w=800&q=80',
    'https://images.unsplash.com/photo-1491921125492-f0b9c835b699?w=800&q=80',
    'https://images.unsplash.com/photo-1543465077-db45d34b88a5?w=800&q=80',
    // — הגה ופנים רכב —
    'https://images.unsplash.com/photo-1449965408869-eaa3f722e40d?w=800&q=80',
    'https://images.unsplash.com/photo-1580273916550-e323be2ae537?w=800&q=80',
    'https://images.unsplash.com/photo-1558618666-fcd25c85cd64?w=800&q=80',
    'https://images.unsplash.com/photo-1494905998402-395d579af36f?w=800&q=80',
    'https://images.unsplash.com/photo-1601362840469-51e4d8d58785?w=800&q=80',
    // — עיר ותנועה —
    'https://images.unsplash.com/photo-1477959858617-67f85cf4f1df?w=800&q=80',
    'https://images.unsplash.com/photo-1486325212027-8081e485255e?w=800&q=80',
    'https://images.unsplash.com/photo-1534430480872-3498386e7856?w=800&q=80',
    'https://images.unsplash.com/photo-1529655683826-aba9b3e77383?w=800&q=80',
    'https://images.unsplash.com/photo-1596394516093-501ba68a0ba6?w=800&q=80',
    'https://images.unsplash.com/photo-1590674899484-d5640e854abe?w=800&q=80',
    'https://images.unsplash.com/photo-1560958089-b8a1929cea89?w=800&q=80',
    'https://images.unsplash.com/photo-1485291571150-772bcfc10da5?w=800&q=80',
    'https://images.unsplash.com/photo-1507136566006-cfc505b114fc?w=800&q=80',
    'https://images.unsplash.com/photo-1473042904451-00171c69419d?w=800&q=80',
    'https://images.unsplash.com/photo-1604357209793-fca5dca89f97?w=800&q=80',
    'https://images.unsplash.com/photo-1647606648893-d9d3b6ec83ca?w=800&q=80',
    // — תוספת מכוניות —
    'https://images.unsplash.com/photo-1552519507-da3b142c6e3d?w=800&q=80',
    'https://images.unsplash.com/photo-1541899481282-d53bffe3c35d?w=800&q=80',
    'https://images.unsplash.com/photo-1583121274602-3e2820c69888?w=800&q=80',
    'https://images.unsplash.com/photo-1607082348824-0a96f2a4b9da?w=800&q=80',
    'https://images.unsplash.com/photo-1618843479313-40f8afb4b4d8?w=800&q=80',
    'https://images.unsplash.com/photo-1606016159991-dfe4f2746ad5?w=800&q=80',
    'https://images.unsplash.com/photo-1617814076367-b759c7d7e738?w=800&q=80',
    'https://images.unsplash.com/photo-1590955559496-763c8a36b2ac?w=800&q=80',
    'https://images.unsplash.com/photo-1541888946425-d81bb19240f5?w=800&q=80',
    'https://images.unsplash.com/photo-1609259266649-9e1e27aeafb4?w=800&q=80',
    'https://images.unsplash.com/photo-1612544444416-9be28c7cbe97?w=800&q=80',
    'https://images.unsplash.com/photo-1625231338027-ac14c0571082?w=800&q=80',
    'https://images.unsplash.com/photo-1622543925917-763c34d1a86e?w=800&q=80',
    'https://images.unsplash.com/photo-1571607388263-1044f9ea01dd?w=800&q=80',
    // — תוספת כבישים ונסיעה —
    'https://images.unsplash.com/photo-1514316454349-750a7fd3da3a?w=800&q=80',
    'https://images.unsplash.com/photo-1476357471311-43c0db9fb2b4?w=800&q=80',
    'https://images.unsplash.com/photo-1507919909716-c8262e491cde?w=800&q=80',
    'https://images.unsplash.com/photo-1519566335946-e6f65f0f4fdf?w=800&q=80',
    'https://images.unsplash.com/photo-1548345680-f5475ea5df84?w=800&q=80',
    'https://images.unsplash.com/photo-1553531384-397c80973a0b?w=800&q=80',
    'https://images.unsplash.com/photo-1469854523086-cc02fe5d8800?w=800&q=80',
    'https://images.unsplash.com/photo-1508739773434-c26b3d09e071?w=800&q=80',
    'https://images.unsplash.com/photo-1471445664267-a0a66869d0db?w=800&q=80',
    'https://images.unsplash.com/photo-1582719471384-894fbb16e074?w=800&q=80',
    // — תוספת הגה ופנים רכב —
    'https://images.unsplash.com/photo-1547245324-d777c6f05e80?w=800&q=80',
    'https://images.unsplash.com/photo-1504215680853-026ed2a45def?w=800&q=80',
    'https://images.unsplash.com/photo-1502161254066-6c74afbf07aa?w=800&q=80',
    'https://images.unsplash.com/photo-1621929747188-0b4dc28498d2?w=800&q=80',
    'https://images.unsplash.com/photo-1616455579100-2ceaa4eb2d37?w=800&q=80',
    // — תוספת לימוד נהיגה —
    'https://images.unsplash.com/photo-1558618666-fcd25c85cd64?w=800&q=80',
    'https://images.unsplash.com/photo-1517026575980-3e1e2dedeab4?w=800&q=80',
    'https://images.unsplash.com/photo-1600320254374-ce2d293c324e?w=800&q=80',
    'https://images.unsplash.com/photo-1623479322729-28b25c16b011?w=800&q=80',
    'https://images.unsplash.com/photo-1593642632559-0c6d3fc62b89?w=800&q=80',
    'https://images.unsplash.com/photo-1489824904134-891ab64532f1?w=800&q=80',
    'https://images.unsplash.com/photo-1585771724684-38269d6639fd?w=800&q=80',
    'https://images.unsplash.com/photo-1619604169985-b99cf4e071fc?w=800&q=80',
    'https://images.unsplash.com/photo-1612345710010-87d9c7c7f52a?w=800&q=80',
    'https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?w=800&q=80',
    'https://images.unsplash.com/photo-1493238792000-8113da705763?w=800&q=80',
    'https://images.unsplash.com/photo-1540912434892-7d83f5de5ba7?w=800&q=80',
    'https://images.unsplash.com/photo-1527786356703-4b100091cd2c?w=800&q=80',
    'https://images.unsplash.com/photo-1499828329641-13a8498a7e3c?w=800&q=80',
    // — הומוריסטיות (רכבים ונהיגה) —
    'https://images.unsplash.com/photo-1649513361889-c30692c1b024?w=800&q=80',  // ניידת משטרה מצחיקה
    'https://images.unsplash.com/photo-1750623982067-00c28d289357?w=800&q=80',  // לוחית רישוי "bejaysus"
    'https://images.unsplash.com/photo-1726573685140-cf114c7253eb?w=800&q=80',  // מכונית זעירה
    'https://images.unsplash.com/photo-1768194771550-9f7f6c4cd0ad?w=800&q=80',  // מדבקה "No Baby on Board"
    'https://images.unsplash.com/photo-1779263730551-e171cbc148d1?w=800&q=80',  // אווז ממולא על מנוע
    'https://images.unsplash.com/photo-1777669936862-1aede0da1d57?w=800&q=80',  // Buzz Lightyear על ספוילר
    'https://images.unsplash.com/photo-1713996702515-d91c412f6b4f?w=800&q=80',  // משאית מצחיקה
    'https://images.unsplash.com/photo-1588577289240-a6fc27f9c3f9?w=800&q=80',  // מכונית ורודה על שלט חנייה
  ];
  // אקראי אמיתי — תמונה שונה לכל תלמיד בכל התראה
  return images[Math.floor(Math.random() * images.length)];
}

/**
 * נרמול מספר טלפון ל-972XXXXXXXXX
 */
function normalizePhone(phone) {
  phone = phone.replace(/[\s\-\(\)]/g, '');
  if (phone.startsWith('0'))  phone = '972' + phone.slice(1);
  if (phone.startsWith('+'))  phone = phone.slice(1);
  if (!/^972\d{9}$/.test(phone)) return null;
  return phone;
}

/**
 * עדכון רשומה קיימת ב-sheet לפי id
 */
function _upsertById(ss, sheetName, item) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    try {
      const existing = JSON.parse(vals[i][0]);
      if (existing.id === item.id) {
        sh.getRange(i + 2, 1).setValue(JSON.stringify(item));
        return;
      }
    } catch(e) {}
  }
}

function _deleteById(ss, sheetName, id, field) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (!vals[i][0]) continue;
    try {
      const item = JSON.parse(vals[i][0]);
      if (String(item[field]) === String(id)) sh.deleteRow(i + 2);
    } catch(e) {}
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// מיגרציה חד-פעמית: טבלאות עבריות → JSON
// הרץ פעם אחת מה-editor לאחר clasp push
// ============================================================
function debugLessonsSheet() {
  const ss = SpreadsheetApp.openById('1v-7JPq9huUjZYaqPFn05JRnlmlJt7I3886H_bZuAN7Q');
  const sh = ss.getSheetByName('שיעורים');
  const rows = sh.getDataRange().getValues();
  Logger.log('שורות: ' + rows.length + ' | עמודות: ' + rows[0].length);
  for (let i = 0; i <= Math.min(3, rows.length - 1); i++) {
    Logger.log('שורה ' + i + ': col0=' + JSON.stringify(rows[i][0]) + ' | col1=' + JSON.stringify(rows[i][1]) + ' | col6=' + JSON.stringify(rows[i][6]));
  }
}

function migrateHebToJson2() {
  const ss = SpreadsheetApp.openById('1v-7JPq9huUjZYaqPFn05JRnlmlJt7I3886H_bZuAN7Q');

  const STATUS_S = {'פעיל':'active','ארכיון':'archive','לא פעיל':'inactive'};
  const STATUS_L = {'מתוכנן':'scheduled','בוצע':'done','בוטל':'cancelled','לא הגיע':'noshow','נכשל':'failed'};
  const TYPE_L   = {'שיעור':'lesson','טסט':'test','פנימי':'internal','הערה':'note','אירוע':'event','פנימי חינם':'internal_free'};

  function toDate(v){ if(!v) return ''; try{ return Utilities.formatDate(new Date(v),'Asia/Jerusalem','yyyy-MM-dd'); }catch(e){ return String(v).slice(0,10); } }
  function toTime(v){ if(!v) return ''; try{ return Utilities.formatDate(new Date(v),'UTC','HH:mm'); }catch(e){ return String(v); } }
  function toStr(v){ return String(v||'').trim(); }
  function toNum(v){ const n=parseFloat(String(v).replace(/[^\d.]/g,'')); return isNaN(n)?0:n; }
  function toBool(v){ return v==='✓'||v===true||v==='TRUE'; }

  const studRows = ss.getSheetByName('תלמידים').getDataRange().getValues().slice(1);
  const students = studRows.map(r => {
    const id = toStr(r[0]); if(!id) return null;
    return { id, name:toStr(r[1]), phone:toStr(r[2]).replace(/[-\s]/g,''),
      birthdate:toDate(r[3]), idNum:toStr(r[4]), email:toStr(r[5]),
      intakeType:toStr(r[6]), regFee:toNum(r[7]), regPaid:toStr(r[8]),
      package:toStr(r[9]), price:toNum(r[10])||170,
      startDate:toDate(r[11]), status:STATUS_S[toStr(r[12])]||'active',
      archiveDate:toDate(r[13]), health:toBool(r[14]), eyeTest:toBool(r[15]),
      theory:toBool(r[16]), internalExam:toBool(r[17]), parking:toBool(r[18]),
      notes:toStr(r[19]), periodStart:toDate(r[20]) };
  }).filter(Boolean);

  const lesRows = ss.getSheetByName('שיעורים').getDataRange().getValues().slice(1);
  const lessons = lesRows.map(r => {
    const id = toStr(r[0]); if(!id) return null;
    return { id, studentId:toStr(r[1]), date:toDate(r[2]), time:toTime(r[3]),
      type:TYPE_L[toStr(r[4])]||'lesson', duration:toNum(r[5])||1,
      status:STATUS_L[toStr(r[6])]||'scheduled', notes:toStr(r[7]), waSent:false };
  }).filter(Boolean);

  const payRows = ss.getSheetByName('תשלומים').getDataRange().getValues().slice(1);
  const payments = payRows.map(r => {
    const id = toStr(r[0]); if(!id) return null;
    return { id, studentId:toStr(r[1]), amount:toNum(r[2]),
      date:toDate(r[3]), method:toStr(r[4]), notes:toStr(r[5]) };
  }).filter(Boolean);

  _writeSheet(ss, 'Students', students);
  _writeSheet(ss, 'Lessons',  lessons);
  _writeSheet(ss, 'Payments', payments);

  Logger.log('✅ מיגרציה הושלמה');
  Logger.log('תלמידים: ' + students.length);
  Logger.log('שיעורים: ' + lessons.length);
  Logger.log('תשלומים: ' + payments.length);
  Logger.log('שיעורים ללא studentId: ' + lessons.filter(l=>!l.studentId).length);
  Logger.log('תשלומים ללא studentId: ' + payments.filter(p=>!p.studentId).length);
}

function migrateHebToJson() {
  const ss = SpreadsheetApp.openById('1v-7JPq9huUjZYaqPFn05JRnlmlJt7I3886H_bZuAN7Q');

  // ─── עזרים ───
  function toDate(v) {
    if (!v) return '';
    try { return Utilities.formatDate(new Date(v), 'Asia/Jerusalem', 'yyyy-MM-dd'); } catch(e) { return String(v); }
  }
  function toNum(v) { const n = parseFloat(String(v).replace(/[^\d.]/,'')); return isNaN(n) ? 0 : n; }
  function toStr(v) { return String(v || '').trim(); }
  function toBool(v) { return v === '✓' || v === true || v === 'TRUE'; }
  function uid() { return 'mig' + Date.now().toString(36) + Math.random().toString(36).slice(2,6); }

  const STATUS_S  = {'פעיל':'active','ארכיון':'archive','לא פעיל':'inactive'};
  const STATUS_L  = {'מתוכנן':'scheduled','בוצע':'done','בוטל':'cancelled','לא הגיע':'noshow','נכשל':'failed'};
  const TYPE_L    = {'שיעור':'lesson','טסט':'test','פנימי':'internal','הערה':'note','אירוע':'event','פנימי חינם':'internal_free'};

  // ─── תלמידים ───
  // עמודות: שם מלא(0) טלפון(1) תאריך לידה(2) ת"ז(3) מייל(4) סוג קליטה(5)
  //         דמי רישום(6) תשלום דמי רישום(7) חבילה(8) מחיר לשיעור(9)
  //         תחילת לימודים(10) סטטוס(11) תאריך ארכיון(12) הצהרת בריאות(13)
  //         בדיקת ראייה(14) תיאוריה(15) מבחן פנימי(16) חניות(17) הערות(18) תחילת תקופה נוכחית(19)
  const studSh   = ss.getSheetByName('תלמידים');
  const studRows = studSh.getDataRange().getValues().slice(1); // דלג כותרת

  const idNumMap  = {}; // ת"ז  → studentId
  const nameMap   = {}; // שם   → studentId

  const students = studRows.map(r => {
    const name = toStr(r[0]); if (!name) return null;
    const id   = uid();
    const idNum = toStr(r[3]);
    const phone = toStr(r[1]).replace(/[-\s]/g,'');
    if (idNum) idNumMap[idNum] = id;
    nameMap[name] = id;
    return {
      id, name, phone,
      birthdate:    toDate(r[2]),
      idNum,
      email:        toStr(r[4]),
      intakeType:   toStr(r[5]),
      regFee:       toNum(r[6]),
      regPaid:      toStr(r[7]),
      package:      toStr(r[8]),
      price:        toNum(r[9]) || 170,
      startDate:    toDate(r[10]),
      status:       STATUS_S[toStr(r[11])] || 'active',
      archiveDate:  toDate(r[12]),
      health:       toBool(r[13]),
      eyeTest:      toBool(r[14]),
      theory:       toBool(r[15]),
      internalExam: toBool(r[16]),
      parking:      toBool(r[17]),
      notes:        toStr(r[18]),
      periodStart:  toDate(r[19]),
    };
  }).filter(Boolean);

  // ─── שיעורים ───
  // עמודות: תאריך(0) שעה(1) סוג(2) יחידות(3) סטטוס(4) הערות(5) שם תלמיד(6) ת"ז תלמיד(7)
  const lesSh   = ss.getSheetByName('שיעורים');
  const lesRows = lesSh.getDataRange().getValues().slice(1);

  const lessons = lesRows.map(r => {
    const date = toDate(r[0]); if (!date) return null;
    const idNum = toStr(r[7]);
    const sName = toStr(r[6]);
    const studentId = idNumMap[idNum] || nameMap[sName] || null;
    return {
      id:        uid(),
      studentId,
      date,
      time:      toStr(r[1]),
      type:      TYPE_L[toStr(r[2])] || 'lesson',
      duration:  toNum(r[3]) || 1,
      status:    STATUS_L[toStr(r[4])] || 'scheduled',
      notes:     toStr(r[5]),
      waSent:    false,
    };
  }).filter(Boolean);

  // ─── תשלומים ───
  // עמודות: סכום(0) תאריך(1) אמצעי תשלום(2) הערות(3) שם תלמיד(4) ת"ז תלמיד(5)
  const paySh   = ss.getSheetByName('תשלומים');
  const payRows = paySh.getDataRange().getValues().slice(1);

  const payments = payRows.map(r => {
    const amount = toNum(r[0]); if (!amount) return null;
    const idNum  = toStr(r[5]);
    const sName  = toStr(r[4]);
    const studentId = idNumMap[idNum] || nameMap[sName] || null;
    return {
      id:       uid(),
      studentId,
      amount,
      date:     toDate(r[1]),
      method:   toStr(r[2]),
      notes:    toStr(r[3]),
    };
  }).filter(Boolean);

  // ─── כתיבה ───
  _writeSheet(ss, SHEET_STUDENTS, students);
  _writeSheet(ss, SHEET_LESSONS,  lessons);
  _writeSheet(ss, SHEET_PAYMENTS, payments);

  Logger.log('✅ מיגרציה הושלמה');
  Logger.log('תלמידים: '  + students.length);
  Logger.log('שיעורים: '  + lessons.length);
  Logger.log('תשלומים: '  + payments.length);
  Logger.log('שיעורים ללא תלמיד: ' + lessons.filter(l => !l.studentId).length);
  Logger.log('תשלומים ללא תלמיד: ' + payments.filter(p => !p.studentId).length);
}