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
const SHEET_STUDENTS = 'תלמידים';
const SHEET_LESSONS  = 'שיעורים';
const SHEET_PAYMENTS = 'תשלומים';

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
    } else {
      result = { ok: false, msg: 'פעולה לא מוכרת: ' + action };
    }
  } catch(err) {
    result = { ok: false, msg: err.toString() };
  }
  return _json(result);
}

// ─── doPost — כתיבה / מחיקה ────────────────────
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

  // טען שמות תלמידים לצורך ולידציה
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
