// ============================================================
// Code.gs — נקודת כניסה ראשית
// driving-school-code2 | גרסה 1.0 | 18/05/2026
// ============================================================

// ── קונפיגורציה ──
const SS_ID = '1v-7JPq9huUjZYaqPFn05JRnlmlJt7I3886H_bZuAN7Q';

// מחירון שיעורים
const PRICE_OPTIONS = [160, 170, 180, 190];
const PACKAGE_A = { name: 'חבילה א׳', total: 5200, lessons: 28, tests: 1, internal: 1 };
const PACKAGE_B = { name: 'חבילה ב׳', total: 5800, lessons: 28, tests: 2, internal: 1 };

// מחירים קבועים
const PRICE_TEST     = 230; // טסט חיצוני
const PRICE_INTERNAL_REGULAR = 50;  // פנימי לתלמיד רגיל
const PRICE_INTERNAL_PACKAGE = 0;   // פנימי בחבילה
const MAX_TESTS      = 20;  // מקסימום טסטים חיצוניים

// תאריך נקודת אפס
const ZERO_DATE = '18/05/2026';

// שמות גיליונות
const SHEET_STUDENTS  = 'תלמידים';
const SHEET_ARCHIVE   = 'ארכיון';
const SHEET_LESSONS   = 'שיעורים';
const SHEET_PAYMENTS  = 'תשלומים';
const SHEET_BALANCES  = 'יתרות';

// סטטוסים
const STATUS_PLANNED   = 'מתוכנן';
const STATUS_DONE      = 'בוצע';
const STATUS_CANCELED  = 'בוטל';
const STATUS_ABSENT    = 'לא הגיע';

// תוצאות טסט
const TEST_PENDING = 'ממתין';
const TEST_PASSED  = 'עבר';
const TEST_FAILED  = 'נכשל';

// ── doGet ──
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('מערכת ניהול שיעורי נהיגה')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── אתחול גיליונות (הרץ פעם אחת) ──
function initSpreadsheet() {
  const ss = SpreadsheetApp.openById(SS_ID);

  _ensureSheet(ss, SHEET_STUDENTS, [
    'מזהה', 'שם פרטי', 'שם משפחה', 'טלפון', 'ת"ז', 'תאריך לידה',
    'מייל', 'מחיר לשיעור', 'סוג מחיר', 'סטטוס', 'תאריך הצטרפות',
    'יתרת פתיחה', 'הערות'
  ]);

  _ensureSheet(ss, SHEET_ARCHIVE, [
    'מזהה', 'שם פרטי', 'שם משפחה', 'טלפון', 'ת"ז', 'תאריך לידה',
    'מייל', 'מחיר לשיעור', 'סוג מחיר', 'סטטוס', 'תאריך הצטרפות',
    'יתרת פתיחה', 'הערות', 'תאריך סיום'
  ]);

  _ensureSheet(ss, SHEET_LESSONS, [
    'EventID', 'מזהה תלמיד', 'שם תלמיד', 'תאריך', 'שעת התחלה',
    'סוג', 'מחיר', 'סטטוס', 'הערה', 'מזהה תלמיד 2'
  ]);

  _ensureSheet(ss, SHEET_PAYMENTS, [
    'מזהה תלמיד', 'שם תלמיד', 'תאריך', 'סכום', 'שיטה', 'קבלה'
  ]);

  _ensureSheet(ss, SHEET_BALANCES, [
    'מזהה תלמיד', 'שם', 'יתרת פתיחה', 'חוב שיעורים', 'סה"כ שולם', 'יתרה סופית'
  ]);

  Logger.log('✅ גיליונות אותחלו בהצלחה');
}

function _ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1B3A5C').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

// ── עזרים כלליים ──
function _ss() {
  return SpreadsheetApp.openById(SS_ID);
}

function _sheet(name) {
  return _ss().getSheetByName(name);
}

function _nextId(sheetName) {
  const sh = _sheet(sheetName);
  const vals = sh.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < vals.length; i++) {
    const v = Number(vals[i][0]);
    if (v > max) max = v;
  }
  return max + 1;
}

function _today() {
  const d = new Date();
  return _fmtDate(d);
}

function _fmtDate(d) {
  if (!d) return '';
  const day   = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year  = d.getFullYear();
  return `${day}/${month}/${year}`;
}

function _parseDate(str) {
  if (!str) return null;
  const parts = str.split('/');
  if (parts.length !== 3) return null;
  return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
}

function _dateVal(str) {
  const d = _parseDate(str);
  return d ? d.getTime() : 0;
}

function _tsNow() {
  return Date.now();
}
