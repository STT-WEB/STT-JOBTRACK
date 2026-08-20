/**
 * ============================================================================
 *  NOVA-HR · Payload_Util.gs — เครื่องมือกลาง
 *  อ่านชีตด้วย "ชื่อหัวคอลัมน์" ไม่ใช่ตำแหน่ง · หาแท็บแบบชื่อใกล้เคียง · หาไฟล์ Cal รายเดือน
 *  ⚠ ชื่อขึ้นต้น nv... เพราะ Apps Script ใช้ global ร่วมกันทุกไฟล์ (Cal.gs/Code.gs เดิมมี num_/readSheet_ อยู่แล้ว)
 * ============================================================================
 */

/* ============================================================================
   ตัวช่วยพื้นฐาน — อ่านชีตด้วย "ชื่อหัวคอลัมน์" ไม่ใช่ตำแหน่ง
   (จำเป็นมาก เพราะ JOBTRACK เปลี่ยนคอลัมน์มา 4 รุ่นในปีเดียว)
   ============================================================================ */
function readTab_(fileId, tabName, headerRow) {
  var sh = SpreadsheetApp.openById(fileId).getSheetByName(tabName);
  if (!sh) throw new Error('ไม่พบแท็บ "' + tabName + '" ในไฟล์ ' + fileId);
  var vals = sh.getDataRange().getValues();
  var hr = (headerRow || findHeaderRow_(vals)) - 1;
  var head = vals[hr].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = hr + 1; i < vals.length; i++) {
    var o = {}, empty = true;
    for (var c = 0; c < head.length; c++) {
      if (!head[c]) continue;
      o[head[c]] = vals[i][c];
      if (vals[i][c] !== '' && vals[i][c] !== null) empty = false;
    }
    if (!empty) rows.push(o);
  }
  return { head: head, rows: rows };
}

/** เดาแถวหัวตาราง = แถวแรกที่มีข้อความ >= 3 ช่อง */
function findHeaderRow_(vals) {
  for (var i = 0; i < Math.min(vals.length, 20); i++) {
    var n = vals[i].filter(function (v) { return typeof v === 'string' && v.trim() !== ''; }).length;
    if (n >= 3) return i + 1;
  }
  return 1;
}

/** หาแท็บแบบ "ชื่อใกล้เคียง" — ทนกับจุด/เว้นวรรค/ตัวพิมพ์ที่ไม่เหมือนกันในแต่ละเดือน */
function findTab_(ss, keys) {
  var norm = function (x) { return String(x).toUpperCase().replace(/[\s\.\_\-]/g, ''); };
  var shs = ss.getSheets();
  for (var k = 0; k < keys.length; k++) {
    var want = norm(keys[k]);
    for (var i = 0; i < shs.length; i++) if (norm(shs[i].getName()) === want) return shs[i];
    for (var j = 0; j < shs.length; j++) if (norm(shs[j].getName()).indexOf(want) >= 0) return shs[j];
  }
  return null;
}

/** อ่านแท็บเป็น array ของ object โดยใช้ "ชื่อหัวคอลัมน์" */
function nvReadSheet_(sh) {
  if (!sh) return { head: [], rows: [] };
  var vals = sh.getDataRange().getValues();
  if (!vals.length) return { head: [], rows: [] };
  var hr = findHeaderRow_(vals) - 1;
  var head = vals[hr].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = hr + 1; i < vals.length; i++) {
    var o = {}, empty = true;
    for (var c = 0; c < head.length; c++) {
      if (!head[c] || o.hasOwnProperty(head[c])) continue;   // คอลัมน์ชื่อซ้ำ → เอาตัวแรก
      o[head[c]] = vals[i][c];
      if (vals[i][c] !== '' && vals[i][c] !== null) empty = false;
    }
    if (!empty) rows.push(o);
  }
  return { head: head, rows: rows };
}

/** หยิบค่าจากชื่อคอลัมน์ โดยลองหลายชื่อ (รองรับไฟล์ที่ชื่อคอลัมน์ไม่เหมือนกันทุกรุ่น) */
function pick_(row, names, dflt) {
  for (var i = 0; i < names.length; i++) {
    if (row.hasOwnProperty(names[i]) && row[names[i]] !== '' && row[names[i]] !== null) return row[names[i]];
  }
  return (dflt === undefined) ? '' : dflt;
}

function nvNum_(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function q2_(v) { return Math.round(nvNum_(v) * 100) / 100; }

function monthOf_(d) {
  if (d instanceof Date) return d.getMonth() + 1;
  var s = String(d);
  var m = s.match(/(\d{1,2})\/(\d{4})/);           // "1/2026"
  if (m) return +m[1];
  m = s.match(/^\s*(\d{1,2})\s*\./);               // "1. มกราคม"
  if (m) return +m[1];
  return 0;
}

/** '1. ชั่วโมงวันทำงานปกติ' → 1 · '2A. ...' → 2 · '3. ...' → 3 · '4. ...' → 4 */
function htCode_(s) {
  var m = String(s).match(/^\s*(\d)/);
  return m ? +m[1] : 1;
}

function gradeOf_(pct) {
  if (pct >= 95) return 'A+';
  if (pct >= 90) return 'A';
  if (pct >= 85) return 'B';
  if (pct >= 80) return 'C';
  if (pct >= 75) return 'D';
  return 'F';
}

function fmtDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
  return String(d || '');
}

function leadDays_(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return 0;
  return Math.round((b - a) / 86400000);
}

/** ไฟล์ Cal รายเดือน — ชื่อไฟล์จริงมี 2 แบบ ("Cal JOB COST" / "DATA JOB COST") และมีทั้งรุ่นเก่า/(NEW) */
function calFiles_() {
  var it = DriveApp.getFolderById(FILES.CAL_DIR).getFiles(), best = {};
  while (it.hasNext()) {
    var f = it.next(), n = String(f.getName());
    if (n.toUpperCase().indexOf('JOB COST') < 0) continue;
    var mm = n.match(/^\s*(\d{1,2})\s*\./);
    if (!mm) continue;
    var m = +mm[1];
    if (m < 1 || m > 12) continue;
    var isNew = /\(NEW\)/i.test(n) ? 1 : 0;
    var cur = best[m];
    if (!cur || isNew > cur.isNew || (isNew === cur.isNew && f.getLastUpdated() > cur.f.getLastUpdated()))
      best[m] = { f: f, isNew: isNew, name: n };
  }
  var out = [];
  for (var m = 1; m <= 12; m++) if (best[m]) out.push({ m: m, id: best[m].f.getId(), name: best[m].name });
  return out;
}

function latestCalFile_() {
  var fs = calFiles_();
  if (!fs.length) return null;
  return DriveApp.getFileById(fs[fs.length - 1].id);
}
