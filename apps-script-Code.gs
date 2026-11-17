/**
 * พร้อมดูแล — ตัวรับข้อมูลจากแอป (Google Apps Script + Google Sheets)
 * วิธีติดตั้ง:
 * 1) สร้าง Google Sheet ใหม่ (บัญชีของหน่วยงาน/ผู้วิจัย) > ส่วนขยาย > Apps Script
 * 2) วางโค้ดนี้แทนของเดิม > บันทึก
 * 3) ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > ประเภท "เว็บแอป"
 *    เรียกใช้ในฐานะ: ฉัน | ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 * 4) คัดลอก URL ที่ลงท้ายด้วย /exec ไปวางในแอป (โหมดผู้วิจัย > Google Apps Script Web App URL)
 * ข้อมูลที่รับมีเพียงรหัสผู้เข้าร่วม ไม่มีชื่อหรือข้อมูลระบุตัวตน
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    // 1) Summary — 1 แถวต่อผู้เข้าร่วม (อัปเดตทับ)
    const s = d.summary || {};
    const keys = Object.keys(s);
    const sh = sheet_(ss, 'Summary', ['received'].concat(keys));
    upsert_(sh, 2, s.pid, [now].concat(keys.map(k => s[k])));

    // 2) SE — คะแนนรับรู้สมรรถนะแห่งตน 12 ข้อ รายช่วงเวลา
    const se = sheet_(ss, 'SE', ['pid', 'group', 'phase', 'date'].concat([...Array(12)].map((_, i) => 'se' + (i + 1))).concat(['mean']));
    (d.se || []).forEach(r => {
      const m = r.scores.reduce((a, b) => a + b, 0) / r.scores.length;
      upsertKey_(se, d.pid + '|' + r.phase, [d.pid, d.group, r.phase, r.d].concat(r.scores).concat([m]));
    });

    // 3) Fidelity — กิจกรรม 5 กิจกรรม และการติดตาม
    const fi = sheet_(ss, 'Fidelity', ['pid', 'item', 'date', 'minutes', 'channel', 'note']);
    Object.keys(d.acts || {}).forEach(id => {
      const a = d.acts[id];
      upsertKey_(fi, d.pid + '|' + id, [d.pid, id, a.date, a.min, '', a.note]);
    });
    (d.followups || []).forEach((f, i) => upsertKey_(fi, d.pid + '|FU' + (i + 1), [d.pid, 'FU' + (i + 1), f.date, f.min, f.ch, 'สัปดาห์ ' + f.week + ': ' + f.note]));

    // 4) Raw JSON สำรอง
    const raw = sheet_(ss, 'Raw', ['received', 'pid', 'json']);
    raw.appendRow([now, d.pid, JSON.stringify(d).slice(0, 49000)]);

    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function upsert_(sh, keyCol, key, row) {
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][keyCol - 1] === key) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  sh.appendRow(row);
}
function upsertKey_(sh, key, row) {
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][0] + '|' + (vals[i][2] && sh.getName() === 'SE' ? vals[i][2] : vals[i][1]) === key) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  sh.appendRow(row);
}
