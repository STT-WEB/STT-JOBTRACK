/***********************************************************
 * STT NOVA-HR Hub — Import (โมดูลนำเข้าเดือนเก่า)
 * แยกไฟล์ต่างหากจาก Cost engine
 * ▶ rebuildJobcostFromMonthly() = ★ ตัวหลัก ★ ดึงต้นทุนต่อจ๊อบจาก
 *     ไฟล์ "DATA JOB COST" รายเดือน (verify แล้ว) → JOBCOST 2026 MASTER
 *     ใช้ทุกครั้งที่ตรวจเดือนใหม่เสร็จ (เพิ่ม ID ใน MONTH_FILES แล้วรันซ้ำ)
 * ▶ importOldMaster()  = ของเดิม (ดึงจาก ACTUAL_WIDE — เก็บไว้อ้างอิง)
 *   (อ้างอิงค่าคงที่ JOBCOST_FILE_ID + ตัวช่วย findCol_ / findSheetByHeader_ จาก Code.gs)
 ***********************************************************/

// ★ ไฟล์รายเดือนที่ "ตรวจแล้ว ยอดตรง" — เพิ่ม/แก้ที่นี่ที่เดียว
//   (เดือน 6 มิ.ย. ยังไม่ใส่ เพราะพนักงานกำลังกรอก — ตรวจเสร็จค่อยเพิ่มบรรทัด 6:'...')
var MONTH_FILES = {
  1: '1uUMKjX-xNfRJsq8kVxDp7jaulVP6B9OF0Dku4QmnOrQ',   // ม.ค.
  2: '1zjAyEBOz6Li2Tjt9rnPGDkFisJbE1OroO4jw1TlJZbk',   // ก.พ.
  3: '1ipNkMCzTXPh-UW4o6pOJu5wLr-JjlXSArNeZYfQDkEY',   // มี.ค. (NEW)
  4: '1oIhjmU7SwtfYuVmca8_5cJ44c8tUmdEHS1vFzOYXuSg',   // เม.ย. (NEW)
  5: '1bLA21F0fVbBKDAb_bNlB8BJnbqRG7KLUA5S0fiWN4a8',   // พ.ค. (NEW)
  6: '1M50fdr43mIzHwgPzJaVP-EIn91wsovqvuaml1SIPMhc',   // มิ.ย. (NEW)
  7: '17iUJAa0foHoPHXEJnkXaqhB1YHaR4d9zu02RxrnWK7k'    // ก.ค. (NEW)
};

/**
 * ★ ตัวหลัก: rebuild MASTER ใหม่ทั้งหมดจากไฟล์รายเดือนที่ verify แล้ว
 *   - อ่านแท็บ JOB_COST_SUMMARY (ต่อจ๊อบ) คอลัมน์ "ต้นทุน Direct รวม"
 *   - รวมต่อจ๊อบต่อเดือน → เขียน JOBCOST 2026 MASTER (ล้างของเดิมก่อน)
 */
function rebuildJobcostFromMonthly() {
  var jobs = {};                 // code -> {type, name, m:{month->cost}}
  var log = [];
  var months = Object.keys(MONTH_FILES).map(Number).sort(function(a, b){ return a - b; });

  months.forEach(function(m) {
    var rows = readMonthlySummary_(MONTH_FILES[m]);
    if (!rows) { log.push('   ⚠️ เดือน ' + m + ': หาแท็บ JOB_COST_SUMMARY ไม่เจอ (ข้าม)'); return; }
    var tot = 0, n = 0;
    rows.forEach(function(j) {
      if (!jobs[j.code]) jobs[j.code] = { type: j.type, name: j.name, m: {} };
      jobs[j.code].m[m] = (jobs[j.code].m[m] || 0) + j.cost;
      if (!jobs[j.code].name && j.name) jobs[j.code].name = j.name;
      if (!jobs[j.code].type && j.type) jobs[j.code].type = j.type;
      tot += j.cost; n++;
    });
    log.push('   เดือน ' + m + ': ' + n + ' จ๊อบ = ' + Math.round(tot).toLocaleString() + ' บาท');
  });

  // สร้างแถว MASTER
  var out = [];
  for (var code in jobs) {
    var j = jobs[code];
    var row = new Array(24).fill('');
    row[0] = j.type; row[1] = code; row[2] = j.name;
    var sum = 0;
    for (var mm = 1; mm <= 12; mm++) { var v = j.m[mm] || 0; row[2 + mm] = Math.round(v); sum += v; }
    row[15] = 0;                       // ยอดยกมา (ยังไม่ใช้)
    row[16] = Math.round(sum);         // รวมค่าแรง STT
    out.push(row);
  }
  out.sort(function(a, b){ return String(a[1]).localeCompare(String(b[1])); });

  var dst = SpreadsheetApp.openById(JOBCOST_FILE_ID).getSheetByName('MASTER');
  var last = dst.getLastRow();
  if (last > 1) dst.getRange(2, 1, last - 1, 24).clearContent();
  if (out.length) dst.getRange(2, 1, out.length, 24).setValues(out);

  var g = 0; for (var k = 0; k < out.length; k++) g += Number(out[k][16]) || 0;
  Logger.log('✅ Rebuild MASTER จากไฟล์รายเดือน (verify แล้ว)');
  Logger.log(log.join('\n'));
  Logger.log('   รวมทั้งหมด: ' + out.length + ' จ๊อบ | ยอดรวม ' + Math.round(g).toLocaleString() + ' บาท');
  Logger.log('👉 เปิดเว็บ Dashboard แล้วกดรีเฟรช จะเห็นข้อมูลครบ');
}

/** อ่านต้นทุนต่อจ๊อบจากไฟล์รายเดือน → [{code,name,type,cost}] (null ถ้าหาไม่เจอ)
 *  รองรับ 2 รูปแบบ:
 *   - ไฟล์ใหม่ (มี.ค.เป็นต้นไป): แท็บ JOB_COST_SUMMARY / คอลัมน์ "ต้นทุน Direct รวม" (ต่อจ๊อบ)
 *   - ไฟล์เก่า (ม.ค./ก.พ.): แท็บ DIRECT_COST_ALLOC / คอลัมน์ "Direct Cost" (ต่อบรรทัด — rebuild รวมเองต่อจ๊อบ)
 */
function readMonthlySummary_(fileId) {
  var ss;
  try { ss = SpreadsheetApp.openById(fileId); } catch (e) { return null; }
  var costHead = 'ต้นทุน Direct รวม';
  var hit = findSheetByHeader_(ss, ['JOB CODE', costHead]);      // รูปแบบใหม่
  if (!hit) { costHead = 'Direct Cost'; hit = findSheetByHeader_(ss, ['JOB CODE', costHead]); }  // รูปแบบเก่า
  if (!hit) return null;
  var cJob  = findCol_(hit.head, ['JOB CODE']);
  var cName = findCol_(hit.head, ['JOB NAME']);
  var cType = findCol_(hit.head, ['ประเภทงาน']);
  var cCost = findCol_(hit.head, [costHead]);
  if (cJob < 0 || cCost < 0) return null;

  var data = hit.sheet.getDataRange().getValues();
  var out = [];
  for (var i = hit.row; i < data.length; i++) {
    var code = String(data[i][cJob]).trim();
    if (!code || code.indexOf('-') < 0) continue;     // ข้ามหัว/แถวรวม/ST
    var cost = Number(data[i][cCost]) || 0;
    var type = (cType >= 0) ? String(data[i][cType]).trim() : '';
    if (!type) type = (code.indexOf('-') > 0) ? code.split('-')[0] : '';
    out.push({ code: code, name: (cName >= 0) ? String(data[i][cName]).trim() : '', type: type, cost: cost });
  }
  return out;
}

/** นำเข้าเดือนเดียว (เผื่อไม่อยาก rebuild ทั้งหมด) — เช่น importOneMonth(5) */
function importOneMonth(month) {
  if (!MONTH_FILES[month]) { Logger.log('❌ ไม่มี fileId ของเดือน ' + month + ' ใน MONTH_FILES'); return; }
  var rows = readMonthlySummary_(MONTH_FILES[month]);
  if (!rows) { Logger.log('⚠️ หาแท็บ JOB_COST_SUMMARY เดือน ' + month + ' ไม่เจอ'); return; }
  var jobTotal = {};
  rows.forEach(function(j) {
    if (!jobTotal[j.code]) jobTotal[j.code] = { name: j.name, total: 0 };
    jobTotal[j.code].total += j.cost;
  });
  updateMaster_(month, jobTotal);   // ใช้ตัวเขียนเดิมจาก Code.gs (ลงคอลัมน์เดือน + รวม STT)
  var g = 0, n = 0; for (var c in jobTotal) { g += jobTotal[c].total; n++; }
  Logger.log('✅ นำเข้าเดือน ' + month + ': ' + n + ' จ๊อบ = ' + Math.round(g).toLocaleString() + ' บาท');
}

/* ---------- ของเดิม: ดึงจาก ACTUAL_WIDE (เก็บไว้อ้างอิง ไม่ได้ใช้แล้ว) ---------- */
function importOldMaster() {
  var src = SpreadsheetApp.openById(OLD_MASTER_ID);
  var hit = findSheetByHeader_(src, ['JOB CODE', '2026-01']);   // แท็บ ACTUAL_WIDE
  if (!hit) { Logger.log('❌ หาแท็บ ACTUAL_WIDE (JOB CODE + 2026-01) ไม่เจอ'); return; }
  var head = hit.head;
  var cJob  = findCol_(head, ['JOB CODE']);
  var cName = findCol_(head, ['JOB NAME']);
  var cBf   = findCol_(head, ['ยอดยกมา']);
  var mcol = [];
  for (var m = 1; m <= 12; m++) mcol[m] = findCol_(head, ['2026-' + ('0' + m).slice(-2)]);

  var data = hit.sheet.getDataRange().getValues();
  var out = [], count = 0, monthsSeen = {};
  for (var i = hit.row; i < data.length; i++) {
    var r = data[i], job = String(r[cJob]).trim();
    if (!job || job.indexOf('-') < 0) continue;
    var row = new Array(24).fill('');
    row[0] = job.split('-')[0];
    row[1] = job;
    row[2] = String(r[cName]).trim();
    var bf = (cBf >= 0) ? (Number(r[cBf]) || 0) : 0;
    row[15] = Math.round(bf);
    var sum = bf, any = false;
    for (var mm = 1; mm <= 12; mm++) {
      var v = (mcol[mm] >= 0) ? (Number(r[mcol[mm]]) || 0) : 0;
      row[2 + mm] = Math.round(v);
      sum += v;
      if (v > 0) { any = true; monthsSeen[mm] = true; }
    }
    row[16] = Math.round(sum);
    if (any || bf > 0) { out.push(row); count++; }
  }

  var dst = SpreadsheetApp.openById(JOBCOST_FILE_ID).getSheetByName('MASTER');
  var last = dst.getLastRow();
  if (last > 1) dst.getRange(2, 1, last - 1, 24).clearContent();
  if (out.length) dst.getRange(2, 1, out.length, 24).setValues(out);

  var months = Object.keys(monthsSeen).map(Number).sort(function(a,b){return a-b;});
  Logger.log('✅ นำเข้า MASTER จากไฟล์เก่า: ' + count + ' จ๊อบ');
  Logger.log('   เดือนที่มีข้อมูล: ' + months.join(', '));
  var g = 0; for (var k = 0; k < out.length; k++) g += Number(out[k][16]) || 0;
  Logger.log('   รวมค่าแรง STT ทั้งหมด: ' + Math.round(g).toLocaleString() + ' บาท');
}
