/**
 * ============================================================================
 *  NOVA-HR · Payload_Recon.gs — Reconcile Bplus ↔ JOBTRACK
 *  เทียบเวลา/ยอดเงินรายคนรายเดือน แล้วเดาต้นตอให้ HR ตามกฎ 6 ข้อ
 * ============================================================================
 */

/* ============================================================================
   เฟส 4 — Reconcile: Bplus (สแกนนิ้ว) ↔ JOBTRACK (ลงจ๊อบ)
   ไฟล์จริง: "N.Time Bplus 26/MM/2569 To 25/MM/2569" (แปลงเป็น Google Sheets แล้ว 7 เดือน)
   คอลัมน์: สาขา · แผนก · รหัส · ชื่อพนักงาน · เวลารูดบัตร · วันที่ · ชม.งาน · มาสาย · กลับก่อน ·
            OTx1 · OTx1.5 · OTx2 · OTx3 · ขาดงาน · ลาป่วยมี · ป่วยไม่มี · พักร้อน · กิจหัก · กิจไม่หัก · ลาอื่นๆ · ลืมรูดบัตร
   ============================================================================ */
function bplusFiles_() {
  var it = DriveApp.getFolderById(FILES.BPLUS_DIR).getFiles(), out = {};
  while (it.hasNext()) {
    var f = it.next(), n = String(f.getName());
    if (n.indexOf('Bplus') < 0) continue;
    if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;   // ยังไม่แปลงไฟล์ → อ่านไม่ได้ (ตรงกับที่ Mockup บอก)
    var mm = n.match(/^\s*(\d{1,2})\s*\./);
    if (mm) out[+mm[1]] = f.getId();
  }
  return out;
}

function readBplusMonth_(fileId) {
  var ss = SpreadsheetApp.openById(fileId), sh = ss.getSheets()[0];
  var t = nvReadSheet_(sh), by = {};
  t.rows.forEach(function (r) {
    var id = String(pick_(r, ['รหัส', 'รหัสพนักงาน'], '')).trim();
    if (!id) return;
    var o = by[id] || (by[id] = { h: 0, ot: 0, late: 0, absentD: 0, leaveD: 0, forget: 0, noOut: 0, days: 0 });
    o.h += nvNum_(pick_(r, ['ชม.งาน'], 0));
    o.ot += nvNum_(pick_(r, ['OTx1'], 0)) + nvNum_(pick_(r, ['OTx1.5'], 0)) + nvNum_(pick_(r, ['OTx2'], 0)) + nvNum_(pick_(r, ['OTx3'], 0));
    o.late += nvNum_(pick_(r, ['มาสาย'], 0));
    o.absentD += nvNum_(pick_(r, ['ขาดงาน'], 0));
    o.leaveD += nvNum_(pick_(r, ['ลาป่วยมี'], 0)) + nvNum_(pick_(r, ['ป่วยไม่มี'], 0)) + nvNum_(pick_(r, ['พักร้อน'], 0))
              + nvNum_(pick_(r, ['กิจหัก'], 0)) + nvNum_(pick_(r, ['กิจไม่หัก'], 0)) + nvNum_(pick_(r, ['ลาอื่นๆ'], 0));
    o.forget += nvNum_(pick_(r, ['ลืมรูดบัตร'], 0));
    var stamps = String(pick_(r, ['เวลารูดบัตร'], '')).trim();
    if (stamps && stamps.split(/[\s,]+/).filter(function (x) { return /\d/.test(x); }).length === 1) o.noOut++;
    o.days++;
  });
  return by;
}

/**
 * recon[i] = { m, id, bplus, jt, diff, c }
 *   bplus = ชั่วโมงจากเครื่องสแกนนิ้ว (ชม.งาน + OT)
 *   jt    = ชั่วโมงที่ลงจ๊อบใน JOBTRACK (= th ของ jobRows เดือนนั้น)
 *   c     = ดัชนีใน D.causes (-1 = ตรงกัน ไม่ต้องหาต้นตอ)
 *           0 ลืม Check Out · 1 ไม่ได้สแกนนิ้ว · 2 ลงจ๊อบซ้อนเวลา · 3 ประเภทชั่วโมงผิด · 4 ปฏิทินไม่ตรง
 */
function buildRecon_(cal, empMap) {
  var bf = bplusFiles_(), out = [];
  var jtBy = {};
  cal.jobRows.forEach(function (x) {
    var k = x.m + '|' + x.id;
    jtBy[k] = q2_((jtBy[k] || 0) + x.th);
  });
  var bpCache = {};
  cal.rows.forEach(function (r) {
    var m = r.m, id = r.id, e = empMap[id] || {};
    var jt = q2_(jtBy[m + '|' + id] || 0);
    if (!e.direct && !jt) return;                       // Indirect ที่ไม่ได้ลงจ๊อบ ไม่ต้องเทียบ
    var by = bpCache[m];
    if (by === undefined) by = bpCache[m] = (bf[m] ? readBplusMonth_(bf[m]) : null);
    if (!by) { out.push({ m: m, id: id, bplus: 0, jt: jt, diff: q2_(-jt), c: 1 }); return; }
    var b = by[id];
    if (!b) { out.push({ m: m, id: id, bplus: 0, jt: jt, diff: q2_(-jt), c: 1 }); return; }
    var bh = q2_(b.h + b.ot), diff = q2_(bh - jt);
    var c = -1;
    if (Math.abs(diff) > 0.05) {
      if (b.forget > 0 || b.noOut > 0) c = 0;            // ลืม Check Out
      else if (b.absentD > 0 || b.leaveD > 0) c = 4;     // ปฏิทิน/วันลาไม่ตรง
      else if (jt > bh) c = 2;                           // ลงจ๊อบซ้อนเวลา
      else c = 3;                                        // ประเภทชั่วโมงผิด
    }
    out.push({ m: m, id: id, bplus: bh, jt: jt, diff: diff, c: c });
  });
  return out;
}
