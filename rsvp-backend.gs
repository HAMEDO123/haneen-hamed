/* =========================================================
   استقبال تأكيدات الحضور وحفظها في Google Sheet
   الصق هذا الكود في Extensions ▸ Apps Script داخل الجدول
   ========================================================= */
const SHEET_NAME = 'RSVP';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet();
    const d  = JSON.parse(e.postData.contents);
    const name = String(d.name || '').trim().slice(0, 60);
    if (!name) return json({ ok: false, error: 'no-name' });

    const going = !!d.going;
    const count = going ? Math.min(2, Math.max(1, Number(d.count) || 1)) : 0;

    // منع التكرار: لو نفس الاسم سجّل قبل، حدّث سطره بدل إضافة سطر جديد
    const rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][1]).trim() === name) {
        sh.getRange(i + 1, 1, 1, 4).setValues([[new Date(), name, going ? 'حاضر' : 'معتذر', count]]);
        return json({ ok: true, updated: true });
      }
    }
    sh.appendRow([new Date(), name, going ? 'حاضر' : 'معتذر', count]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const sh = getSheet();
  const rows = sh.getDataRange().getValues().slice(1);
  const guests = rows
    .filter(r => String(r[2]) === 'حاضر')
    .map(r => ({ name: String(r[1]), count: Number(r[3]) || 1 }));
  const total = guests.reduce((n, g) => n + g.count, 0);
  const payload = { ok: true, total: total, guests: guests.slice(-60) };

  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(payload);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['الوقت', 'الاسم', 'الحالة', 'عدد الأشخاص']);
  }
  return sh;
}

function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
