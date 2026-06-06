// ============================================================
// reports.gs — דוחות, יתרות, חייבים
// ============================================================

function getBalancesData() {
  const students = getStudentsData();
  const statusMap = {};
  students.forEach(s => { statusMap[s.internalId] = s; });

  const vals = _sheet(SHEET_BALANCES).getDataRange().getValues();
  const out  = [];
  for (let i = 1; i < vals.length; i++) {
    const r   = vals[i];
    const sid = String(r[0]);
    const s   = statusMap[sid];
    if (!s) continue; // לא בתלמידים פעילים (ארכיון)
    out.push({
      internalId  : sid,
      name        : String(r[1]),
      priceType   : s.priceType,
      openBalance : Number(r[2]) || 0,
      lessonDebt  : Number(r[3]) || 0,
      totalPaid   : Number(r[4]) || 0,
      net         : Number(r[5]) || 0,
      status      : s.status
    });
  }
  return out;
}

function getMonthlyIncome(year, month) {
  // שיעורים שבוצעו בחודש נתון
  const vals = _sheet(SHEET_LESSONS).getDataRange().getValues();
  let total = 0;
  const lessons = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[7]) !== STATUS_DONE) continue;
    const d = _parseDate(String(r[3]));
    if (!d) continue;
    if (d.getFullYear() !== year || d.getMonth() + 1 !== month) continue;
    total += Number(r[6]) || 0;
    lessons.push(_rowToLesson(r));
  }
  return { year, month, total, lessons };
}

function getCancellations(year, month) {
  const vals = _sheet(SHEET_LESSONS).getDataRange().getValues();
  const out  = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    const status = String(r[7]);
    if (status !== STATUS_CANCELED && status !== STATUS_ABSENT) continue;
    const d = _parseDate(String(r[3]));
    if (!d) continue;
    if (year && d.getFullYear() !== year) continue;
    if (month && d.getMonth() + 1 !== month) continue;
    out.push(_rowToLesson(r));
  }
  return out;
}
