// ============================================================
// payments.gs — ניהול תשלומים וחבילות
// ============================================================

function getPaymentsData() {
  const vals = _sheet(SHEET_PAYMENTS).getDataRange().getValues();
  const out  = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0]) continue;
    out.push({
      studentId  : String(r[0]),
      studentName: String(r[1]),
      date       : String(r[2]),
      amount     : Number(r[3]) || 0,
      method     : String(r[4]),
      receipt    : String(r[5])
    });
  }
  return out;
}

function getPaymentsByStudent(internalId) {
  return getPaymentsData().filter(p => p.studentId === String(internalId));
}

function addPaymentFromUI(data) {
  if (!data.studentId)  return { ok: false, msg: 'חסר מזהה תלמיד' };
  if (!data.amount || Number(data.amount) <= 0) return { ok: false, msg: 'סכום לא תקין' };
  if (!data.receipt)    return { ok: false, msg: 'חסר מספר קבלה' };

  if (_receiptExists(String(data.receipt)))
    return { ok: false, msg: `קבלה ${data.receipt} כבר קיימת` };

  _sheet(SHEET_PAYMENTS).appendRow([
    String(data.studentId),
    String(data.studentName || ''),
    String(data.date || _today()),
    Number(data.amount),
    String(data.method || 'מזומן'),
    String(data.receipt)
  ]);
  SpreadsheetApp.flush();
  return { ok: true };
}

function deletePayment(receipt) {
  const sh   = _sheet(SHEET_PAYMENTS);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][5]) === String(receipt)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, msg: 'קבלה לא נמצאה' };
}

function _receiptExists(receipt) {
  const vals = _sheet(SHEET_PAYMENTS).getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][5]) === receipt) return true;
  }
  return false;
}

// ── מחירון לUI ──

function getPriceOptions() {
  return {
    regular : PRICE_OPTIONS,
    packageA: PACKAGE_A,
    packageB: PACKAGE_B,
    test    : PRICE_TEST,
    internalRegular: PRICE_INTERNAL_REGULAR,
    internalPackage: PRICE_INTERNAL_PACKAGE
  };
}
