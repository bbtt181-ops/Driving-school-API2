// ============================================================
// students.gs — ניהול תלמידים, ארכיון, ולידציות
// ============================================================

// ── ולידציות ──

function _validateTZ(tz) {
  const s = String(tz).replace(/\D/g, '');
  if (s.length !== 9) return false;
  let sum = 0;
  for (let i = 0; i < 9; i++) {
    let val = Number(s[i]) * (i % 2 === 0 ? 1 : 2);
    if (val > 9) val -= 9;
    sum += val;
  }
  return sum % 10 === 0;
}

function _tzExists(tz, excludeId) {
  const vals = _sheet(SHEET_STUDENTS).getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (excludeId && String(vals[i][0]) === String(excludeId)) continue;
    if (String(vals[i][4]).replace(/\D/g, '') === String(tz).replace(/\D/g, '')) return true;
  }
  // בדוק גם בארכיון
  const arch = _sheet(SHEET_ARCHIVE).getDataRange().getValues();
  for (let i = 1; i < arch.length; i++) {
    if (String(arch[i][4]).replace(/\D/g, '') === String(tz).replace(/\D/g, '')) return true;
  }
  return false;
}

function _calcAge(dobStr) {
  const dob = _parseDate(dobStr);
  if (!dob) return null;
  const now = new Date();
  let years  = now.getFullYear() - dob.getFullYear();
  let months = now.getMonth() - dob.getMonth();
  if (months < 0) { years--; months += 12; }
  if (now.getDate() < dob.getDate()) months--;
  if (months < 0) { years--; months += 12; }
  return { years, months, totalMonths: years * 12 + months };
}

function _validateAge(dobStr) {
  const age = _calcAge(dobStr);
  if (!age) return { ok: false, msg: 'תאריך לידה לא תקין' };
  if (age.totalMonths < 16 * 12 + 6)
    return { ok: false, msg: `התלמיד צעיר מדי — גיל מינימלי 16.5 (גיל נוכחי: ${age.years} שנה ו-${age.months} חודשים)` };
  return { ok: true, age };
}

function _validateEmail(email) {
  if (!email) return true; // אופציונלי
  return /^[^\s@]+@gmail\.com$/i.test(email.trim());
}

// ── קריאת תלמידים ──

function getStudentsData() {
  const vals = _sheet(SHEET_STUDENTS).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0]) continue;
    out.push({
      internalId : String(r[0]),
      firstName  : String(r[1]),
      lastName   : String(r[2]),
      fullName   : `${r[1]} ${r[2]}`,
      phone      : String(r[3]),
      tz         : String(r[4]),
      dob        : String(r[5]),
      email      : String(r[6]),
      price      : Number(r[7]) || 0,
      priceType  : String(r[8]),  // 'regular' / 'packageA' / 'packageB'
      status     : String(r[9]),
      joinDate   : String(r[10]),
      openBalance: Number(r[11]) || 0,
      notes      : String(r[12])
    });
  }
  return out;
}

function getStudentsFullData() {
  const students = getStudentsData();
  return students.map(s => {
    const age = _calcAge(s.dob);
    return { ...s, ageStr: age ? `${age.years} שנה ו-${age.months} חודשים` : '' };
  });
}

// ── הוספת תלמיד ──

function addNewStudent(data) {
  // ולידציה
  if (!data.firstName || !data.lastName) return { ok: false, msg: 'שם חסר' };
  if (!_validateTZ(data.tz))             return { ok: false, msg: 'ת"ז לא תקינה' };
  if (_tzExists(data.tz))                return { ok: false, msg: 'ת"ז כבר קיימת במערכת' };
  const ageCheck = _validateAge(data.dob);
  if (!ageCheck.ok)                      return { ok: false, msg: ageCheck.msg };
  if (!_validateEmail(data.email))       return { ok: false, msg: 'מייל חייב להסתיים ב-@gmail.com' };

  const id      = _nextId(SHEET_STUDENTS);
  const price   = _resolvePrice(data.priceType, data.price);
  const sh      = _sheet(SHEET_STUDENTS);

  sh.appendRow([
    id,
    data.firstName.trim(),
    data.lastName.trim(),
    data.phone   || '',
    data.tz.replace(/\D/g, ''),
    data.dob     || '',
    (data.email  || '').trim().toLowerCase(),
    price,
    data.priceType || 'regular',
    'פעיל',
    _today(),
    Number(data.openBalance) || 0,
    data.notes   || ''
  ]);

  ensureBalanceRow(id, `${data.firstName} ${data.lastName}`);
  return { ok: true, id };
}

function _resolvePrice(priceType, customPrice) {
  if (priceType === 'packageA') return Math.round((PACKAGE_A.total / PACKAGE_A.lessons) * 100) / 100;
  if (priceType === 'packageB') return Math.round((PACKAGE_B.total / PACKAGE_B.lessons) * 100) / 100;
  return Number(customPrice) || 160;
}

// ── עדכון תלמיד ──

function updateStudentDetails(data) {
  const sh   = _sheet(SHEET_STUDENTS);
  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(data.internalId)) { row = i + 1; break; }
  }
  if (row === -1) return { ok: false, msg: 'תלמיד לא נמצא' };

  if (data.tz && !_validateTZ(data.tz))              return { ok: false, msg: 'ת"ז לא תקינה' };
  if (data.tz && _tzExists(data.tz, data.internalId)) return { ok: false, msg: 'ת"ז כבר קיימת' };
  if (data.dob) {
    const ageCheck = _validateAge(data.dob);
    if (!ageCheck.ok) return { ok: false, msg: ageCheck.msg };
  }
  if (!_validateEmail(data.email)) return { ok: false, msg: 'מייל חייב להסתיים ב-@gmail.com' };

  const setCell = (col, val) => {
    try { sh.getRange(row, col).setValue(val); } catch(e) {}
  };

  if (data.firstName)   setCell(2,  data.firstName.trim());
  if (data.lastName)    setCell(3,  data.lastName.trim());
  if (data.phone)       setCell(4,  data.phone);
  if (data.tz)          setCell(5,  data.tz.replace(/\D/g, ''));
  if (data.dob)         setCell(6,  data.dob);
  if (data.email !== undefined) setCell(7, (data.email || '').trim().toLowerCase());
  if (data.priceType) {
    const price = _resolvePrice(data.priceType, data.price);
    setCell(8,  price);
    setCell(9,  data.priceType);
  }
  if (data.status)      setCell(10, data.status);
  if (data.openBalance !== undefined) setCell(12, Number(data.openBalance) || 0);
  if (data.notes !== undefined)       setCell(13, data.notes);

  return { ok: true };
}

// ── ארכיון ──

function archiveStudent(internalId) {
  const sh   = _sheet(SHEET_STUDENTS);
  const arch = _sheet(SHEET_ARCHIVE);
  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(internalId)) { row = i + 1; break; }
  }
  if (row === -1) return { ok: false, msg: 'תלמיד לא נמצא' };

  const studentRow = vals[row - 1];
  arch.appendRow([...studentRow, _today()]);
  sh.deleteRow(row);

  return { ok: true };
}

function getArchiveStudents() {
  const vals = _sheet(SHEET_ARCHIVE).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0]) continue;
    out.push({
      internalId : String(r[0]),
      fullName   : `${r[1]} ${r[2]}`,
      phone      : String(r[3]),
      tz         : String(r[4]),
      status     : String(r[9]),
      joinDate   : String(r[10]),
      endDate    : String(r[13])
    });
  }
  return out;
}

// ── יתרות ──

function ensureBalanceRow(id, name) {
  const sh   = _sheet(SHEET_BALANCES);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return;
  }
  const row = vals.length + 1;
  sh.getRange(row, 1).setValue(id);
  sh.getRange(row, 2).setValue(name);

  // יתרת פתיחה מהתלמיד
  sh.getRange(row, 3).setFormula(
    `=IFERROR(VLOOKUP(A${row},תלמידים!A:L,12,0),0)`
  );
  // חוב שיעורים (רק בוצע, מ-18/05/2026)
  sh.getRange(row, 4).setFormula(
    `=SUMPRODUCT((שיעורים!B$2:B$5000=A${row})*(שיעורים!H$2:H$5000="${STATUS_DONE}")*(IFERROR(DATEVALUE(שיעורים!D$2:D$5000),0)>=DATEVALUE("${ZERO_DATE}"))*(IFERROR(DATEVALUE(שיעורים!D$2:D$5000),0)<=TODAY())*שיעורים!G$2:G$5000)`
  );
  // סה"כ שולם
  sh.getRange(row, 5).setFormula(
    `=SUMPRODUCT((תשלומים!A$2:A$5000=A${row})*תשלומים!D$2:D$5000)`
  );
  // יתרה סופית = פתיחה + חוב - שולם
  sh.getRange(row, 6).setFormula(
    `=C${row}+D${row}-E${row}`
  );
}

function ensureAllBalanceRows() {
  const students = getStudentsData();
  students.forEach(s => ensureBalanceRow(s.internalId, s.fullName));
  return { ok: true, count: students.length };
}

// ── היסטוריה לתלמיד ──

function getStudentHistory(internalId) {
  const sid = String(internalId);
  const today = _dateVal(_today());

  // שיעורים
  const lessonVals = _sheet(SHEET_LESSONS).getDataRange().getValues();
  const lessons = [];
  for (let i = 1; i < lessonVals.length; i++) {
    const r = lessonVals[i];
    if (String(r[1]) !== sid) continue;
    const dv = _dateVal(String(r[3]));
    if (dv > today) continue; // לא מציג עתידיים בהיסטוריה
    lessons.push({
      eventId : String(r[0]),
      date    : String(r[3]),
      time    : String(r[4]),
      type    : String(r[5]),
      price   : Number(r[6]) || 0,
      status  : String(r[7]),
      note    : String(r[8])
    });
  }

  // תשלומים
  const payVals = _sheet(SHEET_PAYMENTS).getDataRange().getValues();
  const payments = [];
  for (let i = 1; i < payVals.length; i++) {
    const r = payVals[i];
    if (String(r[0]) !== sid) continue;
    payments.push({
      date    : String(r[2]),
      amount  : Number(r[3]) || 0,
      method  : String(r[4]),
      receipt : String(r[5])
    });
  }

  // יתרה מגיליון יתרות
  const balVals = _sheet(SHEET_BALANCES).getDataRange().getValues();
  let balance = { openBalance: 0, lessonDebt: 0, totalPaid: 0, net: 0 };
  for (let i = 1; i < balVals.length; i++) {
    if (String(balVals[i][0]) === sid) {
      balance = {
        openBalance : Number(balVals[i][2]) || 0,
        lessonDebt  : Number(balVals[i][3]) || 0,
        totalPaid   : Number(balVals[i][4]) || 0,
        net         : Number(balVals[i][5]) || 0
      };
      break;
    }
  }

  return { lessons, payments, balance };
}

// ── חייבים ──

function getDebtors(threshold) {
  threshold = Number(threshold) || 0;
  const balVals = _sheet(SHEET_BALANCES).getDataRange().getValues();
  const students = getStudentsData();
  const statusMap = {};
  students.forEach(s => { statusMap[s.internalId] = s.status; });

  const out = [];
  for (let i = 1; i < balVals.length; i++) {
    const r   = balVals[i];
    const sid = String(r[0]);
    if (statusMap[sid] && statusMap[sid] !== 'פעיל' && statusMap[sid] !== 'מושהה') continue;
    const net = Number(r[5]) || 0;
    if (net > threshold) {
      out.push({ internalId: sid, name: String(r[1]), balance: net });
    }
  }
  out.sort((a, b) => b.balance - a.balance);
  return out;
}

// ── מחיקה מלאה ──

function deleteStudentFull(internalId) {
  const sid = String(internalId);

  // מחק שורות שיעורים
  const lsh  = _sheet(SHEET_LESSONS);
  const lvals = lsh.getDataRange().getValues();
  for (let i = lvals.length - 1; i >= 1; i--) {
    if (String(lvals[i][1]) === sid) lsh.deleteRow(i + 1);
  }

  // מחק שורות תשלומים
  const psh  = _sheet(SHEET_PAYMENTS);
  const pvals = psh.getDataRange().getValues();
  for (let i = pvals.length - 1; i >= 1; i--) {
    if (String(pvals[i][0]) === sid) psh.deleteRow(i + 1);
  }

  // מחק שורת יתרה
  const bsh  = _sheet(SHEET_BALANCES);
  const bvals = bsh.getDataRange().getValues();
  for (let i = bvals.length - 1; i >= 1; i--) {
    if (String(bvals[i][0]) === sid) bsh.deleteRow(i + 1);
  }

  // מחק תלמיד
  const ssh  = _sheet(SHEET_STUDENTS);
  const svals = ssh.getDataRange().getValues();
  for (let i = svals.length - 1; i >= 1; i--) {
    if (String(svals[i][0]) === sid) { ssh.deleteRow(i + 1); break; }
  }

  return { ok: true };
}
