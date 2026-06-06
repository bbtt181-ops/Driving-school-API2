// ============================================================
// lessons.gs — ניהול שיעורים, טסטים, אוטומציה לילית
// ============================================================

// ── קריאת שיעורים ──

function getLessonsData() {
  const vals = _sheet(SHEET_LESSONS).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0]) continue;
    out.push(_rowToLesson(r));
  }
  return out;
}

function getLessonsByStudent(internalId) {
  const sid  = String(internalId);
  const all  = getLessonsData();
  return all.filter(l => l.studentId === sid || l.studentId2 === sid);
}

function _rowToLesson(r) {
  return {
    eventId    : String(r[0]),
    studentId  : String(r[1]),
    studentName: String(r[2]),
    date       : String(r[3]),
    time       : String(r[4]),
    type       : String(r[5]),
    price      : Number(r[6]) || 0,
    status     : String(r[7]),
    note       : String(r[8]),
    studentId2 : String(r[9] || '')
  };
}

// ── הוספת שיעור ──

function addLessonFromUI(data) {
  const sh = _sheet(SHEET_LESSONS);

  const lessonType = String(data.type || 'בודד');
  const price      = _calcLessonPrice(data);
  const eventId    = 'L_' + _tsNow();

  sh.appendRow([
    eventId,
    String(data.studentId),
    String(data.studentName),
    String(data.date),
    String(data.time),
    lessonType,
    price,
    STATUS_PLANNED,
    String(data.note || ''),
    String(data.studentId2 || '')
  ]);
  SpreadsheetApp.flush();

  return { ok: true, eventId, price };
}

function _calcLessonPrice(data) {
  const type = String(data.type || '');

  if (type === 'טסט') {
    // בחבילה: הטסטים הראשונים חינם
    if (data.priceType === 'packageA' || data.priceType === 'packageB') {
      const freeTests = data.priceType === 'packageA' ? PACKAGE_A.tests : PACKAGE_B.tests;
      const usedTests = _countTestsByStudent(String(data.studentId));
      return usedTests < freeTests ? 0 : PRICE_TEST;
    }
    return PRICE_TEST;
  }

  if (type === 'פנימי') {
    if (data.priceType === 'packageA' || data.priceType === 'packageB') return PRICE_INTERNAL_PACKAGE;
    return PRICE_INTERNAL_REGULAR;
  }

  // שיעור רגיל בחבילה — אחרי 28 לפי מחיר נוסף
  if (data.priceType === 'packageA' || data.priceType === 'packageB') {
    const limit = data.priceType === 'packageA' ? PACKAGE_A.lessons : PACKAGE_B.lessons;
    const done  = _countRegularLessonsByStudent(String(data.studentId));
    if (done < limit) return Number(data.price) || 0; // מחיר חבילה
    return Number(data.extraPrice) || 160; // שיעור נוסף — סולם רגיל
  }

  return Number(data.price) || 0;
}

function _countTestsByStudent(sid) {
  const vals = _sheet(SHEET_LESSONS).getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][1]) === sid && String(vals[i][5]) === 'טסט') count++;
  }
  return count;
}

function _countRegularLessonsByStudent(sid) {
  const vals = _sheet(SHEET_LESSONS).getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][1]) !== sid) continue;
    const type   = String(vals[i][5]);
    const status = String(vals[i][7]);
    if (type !== 'טסט' && type !== 'פנימי' && status === STATUS_DONE) count++;
  }
  return count;
}

// ── עדכון שיעור ──

function updateLessonById(id, data) {
  const sh   = _sheet(SHEET_LESSONS);
  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) { row = i + 1; break; }
  }
  if (row === -1) return { ok: false, msg: 'שיעור לא נמצא' };

  const setCell = (col, val) => { try { sh.getRange(row, col).setValue(val); } catch(e) {} };

  if (data.date)        setCell(4, data.date);
  if (data.time)        setCell(5, data.time);
  if (data.type)        setCell(6, data.type);
  if (data.price !== undefined) setCell(7, Number(data.price) || 0);
  if (data.status)      setCell(8, data.status);
  if (data.note !== undefined)  setCell(9, data.note);
  if (data.studentId2 !== undefined) setCell(10, data.studentId2);

  SpreadsheetApp.flush();
  return { ok: true };
}

// ── מחיקת שיעור ──

function deleteLessonById(id) {
  const sh   = _sheet(SHEET_LESSONS);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, msg: 'שיעור לא נמצא' };
}

// ── טסטים ──

function getTestsByStudent(internalId) {
  const sid  = String(internalId);
  const vals = _sheet(SHEET_LESSONS).getDataRange().getValues();
  const out  = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[1]) !== sid) continue;
    if (String(r[5]) !== 'טסט') continue;
    out.push({
      eventId : String(r[0]),
      date    : String(r[3]),
      price   : Number(r[6]) || 0,
      status  : String(r[7]),  // ממתין / עבר / נכשל
      note    : String(r[8])
    });
  }
  return out;
}

function updateTestResult(eventId, result) {
  // result: 'עבר' / 'נכשל'
  const sh   = _sheet(SHEET_LESSONS);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(eventId)) {
      sh.getRange(i + 1, 8).setValue(result);
      SpreadsheetApp.flush();

      if (result === TEST_PASSED) {
        const sid = String(vals[i][1]);
        return { ok: true, passed: true, studentId: sid };
      }
      return { ok: true, passed: false };
    }
  }
  return { ok: false, msg: 'טסט לא נמצא' };
}

function canAddTest(internalId) {
  const count = _countTestsByStudent(String(internalId));
  return { canAdd: count < MAX_TESTS, count, max: MAX_TESTS };
}

// ── שיעורים ממתינים לאישור ──

function getPendingLessons() {
  const today = _dateVal(_today());
  const vals  = _sheet(SHEET_LESSONS).getDataRange().getValues();
  const out   = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[7]) !== STATUS_PLANNED) continue;
    if (_dateVal(String(r[3])) >= today) continue; // רק עבר
    out.push(_rowToLesson(r));
  }
  return out;
}

// ── אוטומציה לילית ──

function autoConfirmLessons() {
  const today = _dateVal(_today());
  const sh    = _sheet(SHEET_LESSONS);
  const vals  = sh.getDataRange().getValues();
  let count   = 0;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[7]) !== STATUS_PLANNED) continue;
    if (_dateVal(String(r[3])) >= today) continue;
    sh.getRange(i + 1, 8).setValue(STATUS_DONE);
    count++;
  }
  if (count > 0) SpreadsheetApp.flush();
  Logger.log(`✅ אושרו אוטומטית ${count} שיעורים`);
  return count;
}

function setupDailyTrigger() {
  // הרץ פעם אחת בלבד
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoConfirmLessons') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoConfirmLessons')
    .timeBased().atHour(0).everyDays(1).create();
  Logger.log('✅ טריגר לילי הוגדר');
}

// ── אישור ידני של שיעורים ──

function confirmLessonsBatch(eventIds) {
  const sh   = _sheet(SHEET_LESSONS);
  const vals = sh.getDataRange().getValues();
  const idSet = new Set(eventIds.map(String));
  let count = 0;
  for (let i = 1; i < vals.length; i++) {
    if (idSet.has(String(vals[i][0]))) {
      sh.getRange(i + 1, 8).setValue(STATUS_DONE);
      count++;
    }
  }
  if (count > 0) SpreadsheetApp.flush();
  return { ok: true, count };
}
