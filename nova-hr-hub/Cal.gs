/***********************************************************
 * STT NOVA-HR Hub — Cal / Payroll / Bplus reader (ตัวอ่านกลาง)
 * ▶ getCalTab(month, tabName)   = อ่านแท็บใดก็ได้จากไฟล์ Cal เดือนนั้น
 * ▶ getPayrollActual(month)     = อ่านเงินเดือนจ่ายจริง (กรองเฉพาะเดือนที่เลือก)
 * ▶ getBplus(month)             = อ่านเวลาสแกนนิ้ว Bplus ของเดือนนั้น จากโฟลเดอร์ Time Bplus
 * คืน {ok, headers[], rows[][], total, tab, file} — ไม่ต้องเปิด Google Sheets
 ***********************************************************/

/** แถวนี้เป็นหัวตารางไหม? (>=3 ช่องไม่ว่าง และมีตัวหนังสือ(ไม่ใช่เลขล้วน) >=2) — กันจับแถวเลข 1,2,3 ผิด */
function isHeaderRow_(row) {
  var nonEmpty = 0, text = 0;
  for (var k = 0; k < Math.min(row.length, 60); k++) {
    var s = String(row[k]).trim();
    if (s) { nonEmpty++; if (isNaN(Number(s))) text++; }
  }
  return nonEmpty >= 3 && text >= 2;
}

/** เทียบค่าตัวกรองแบบยืดหยุ่น (เลข/ข้อความ) — '7' == 7 == '7.0' */
function eqVal_(a, b) {
  var sa = String(a).trim(), sb = String(b).trim();
  if (sa === sb) return true;
  var na = Number(sa), nb = Number(sb);
  return !isNaN(na) && !isNaN(nb) && na === nb;
}

/** อ่านช่วงข้อมูลจาก sheet → {headers, rows, total}
 *  - หาแถวหัวตารางเองใน 12 แถวแรก
 *  - ดึงคอลัมน์ให้ "ครบ" (นับจากหัวตาราง + ตัวอย่างข้อมูล ไม่ตัดคอลัมน์ที่มีข้อมูลทิ้ง)
 *  - filterCol/filterVal = กรองเฉพาะแถวที่คอลัมน์นั้นตรงค่า (เช่น เลขเดือน = 7) */
function readSheet_(sh, filterCol, filterVal, maxRows) {
  var v = sh.getDataRange().getValues();
  if (!v.length) return { headers: [], rows: [], total: 0 };
  var hr = 0, bestTxt = -1;
  for (var r = 0; r < Math.min(15, v.length); r++) {
    var txt = 0; for (var cc = 0; cc < v[r].length; cc++) { var sv = String(v[r][cc]).trim(); if (sv && isNaN(Number(sv))) txt++; }
    if (txt > bestTxt) { bestTxt = txt; hr = r; }
  }

  var width = 0; for (var i = 0; i < v.length; i++) if (v[i].length > width) width = v[i].length;
  var sampleEnd = Math.min(v.length, hr + 80);
  var lastCol = 0;
  for (var c = 0; c < width; c++) {
    var has = String(v[hr][c] || '').trim() !== '';
    if (!has) for (var rr = hr + 1; rr < sampleEnd; rr++) { if (String(v[rr][c] || '').trim() !== '') { has = true; break; } }
    if (has) lastCol = c;
  }

  var headers = [];
  for (var c2 = 0; c2 <= lastCol; c2++) { var h = String(v[hr][c2] || '').trim(); headers.push(h || ('คอลัมน์ ' + (c2 + 1))); }

  var mc = -1;
  if (filterCol) for (var c3 = 0; c3 < headers.length; c3++) if (headers[c3].indexOf(filterCol) >= 0) { mc = c3; break; }

  var rows = [], cap = maxRows || 1500;
  for (var r2 = hr + 1; r2 < v.length; r2++) {
    var row = v[r2].slice(0, lastCol + 1);
    if (row.every(function (x) { return String(x).trim() === ''; })) continue;
    if (mc >= 0 && !eqVal_(row[mc], filterVal)) continue;
    rows.push(row.map(function (x) { return (x instanceof Date) ? Utilities.formatDate(x, 'Asia/Bangkok', 'dd/MM/yyyy') : x; }));
    if (rows.length >= cap) break;
  }
  return { headers: headers, rows: rows, total: rows.length };
}

/** หา fileId ไฟล์ Cal เดือนที่ระบุ จากโฟลเดอร์ Cal ใน Registry (เลือก (NEW) ก่อน) */
function getMonthFileId_(month) {
  var reg = getRegistry_(2026);
  if (!reg.calFolder) throw new Error('ไม่พบ Cal Folder ใน Registry');
  var it = DriveApp.getFolderById(reg.calFolder).getFiles(), pfx = String(month) + '.', found = null, foundNew = null;
  while (it.hasNext()) { var f = it.next(), n = f.getName(); if (n.indexOf(pfx) === 0) { if (n.toUpperCase().indexOf('NEW') >= 0) foundNew = f.getId(); else if (!found) found = f.getId(); } }
  return foundNew || found;
}

/** อ่านแท็บ tabName จากไฟล์ Cal เดือน month */
function getCalTab(month, tabName) {
  try {
    var fid = getMonthFileId_(month);
    if (!fid) return { ok: false, message: 'ไม่พบไฟล์ Cal เดือน ' + month };
    var ss = SpreadsheetApp.openById(fid);
    var sh = ss.getSheetByName(tabName);
    if (!sh) { var shs = ss.getSheets(); for (var i = 0; i < shs.length; i++) if (shs[i].getName().toUpperCase().indexOf(String(tabName).toUpperCase()) >= 0) { sh = shs[i]; break; } }
    if (!sh) return { ok: false, message: 'ไม่พบแท็บ ' + tabName + ' ในไฟล์เดือน ' + month + ' (' + ss.getName() + ')' };
    var d = readSheet_(sh);
    return { ok: true, headers: d.headers, rows: d.rows, total: d.total, tab: sh.getName(), file: ss.getName() };
  } catch (e) { return { ok: false, message: String(e) }; }
}

/** อ่านเงินเดือนจ่ายจริง (Payroll Actual) ของเดือน month — กรองเฉพาะเดือนที่เลือก */
function getPayrollActual(month) {
  try {
    var reg = getRegistry_(2026);
    if (!reg.payrollActual) return { ok: false, message: 'ไม่พบไฟล์ Payroll Actual ใน Registry' };
    var ss = SpreadsheetApp.openById(reg.payrollActual), shs = ss.getSheets(), sh = null;
    for (var i = 0; i < shs.length; i++) if (shs[i].getName().indexOf('DATA BASE') >= 0) { sh = shs[i]; break; }
    if (!sh) sh = shs[0];
    var d = readSheet_(sh, 'เลขเดือน', month, 3000);
    return { ok: true, headers: d.headers, rows: d.rows, total: d.total, tab: sh.getName(), file: ss.getName() };
  } catch (e) { return { ok: false, message: String(e) }; }
}

/** อ่านเวลา Bplus (สแกนนิ้ว) ของเดือน month จากโฟลเดอร์ Time Bplus (เฉพาะไฟล์ Google Sheets) */
function getBplus(month) {
  try {
    var reg = getRegistry_(2026);
    if (!reg.timeBplusFolder) return { ok: false, message: 'ไม่พบโฟลเดอร์ Time Bplus ใน Registry' };
    var THMON = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
    var it = DriveApp.getFolderById(reg.timeBplusFolder).getFiles();
    var pfx = String(month) + '.', thm = THMON[month] || '###', foundSheet = null, foundOther = null, seen = [];
    while (it.hasNext()) {
      var f = it.next(), n = f.getName(); seen.push(n);
      var match = (n.indexOf(pfx) === 0) || (n.indexOf(thm) >= 0) || (n.indexOf('เดือน ' + month) >= 0) || (n.indexOf('เดือน' + month) >= 0);
      if (!match) continue;
      var mt = ''; try { mt = f.getMimeType(); } catch (e) {}
      if (mt === MimeType.GOOGLE_SHEETS) { if (n.toUpperCase().indexOf('NEW') >= 0) { foundSheet = f.getId(); break; } else if (!foundSheet) foundSheet = f.getId(); }
      else if (!foundOther) foundOther = { id: f.getId(), name: n, mime: mt };
    }
    if (!foundSheet) {
      if (foundOther) return { ok: false, message: 'ไฟล์ Bplus เดือน ' + month + ' คือ "' + foundOther.name + '" แต่เป็นไฟล์ Excel (ไม่ใช่ Google Sheets)\nวิธีแก้: เปิดไฟล์นั้นใน Google Sheets แล้ว File → Save as Google Sheets (หรืออัปโหลดแบบแปลงเป็น Sheets) ระบบถึงจะอ่านได้' };
      return { ok: false, message: 'ไม่พบไฟล์ Bplus เดือน ' + month + '\nไฟล์ที่เจอ: ' + (seen.slice(0, 12).join(', ') || '(ว่าง)') };
    }
    var ss = SpreadsheetApp.openById(foundSheet), sh = ss.getSheets()[0];
    var d = readSheet_(sh, null, null, 4000);
    return { ok: true, headers: d.headers, rows: d.rows, total: d.total, tab: sh.getName(), file: ss.getName() };
  } catch (e) { return { ok: false, message: String(e) }; }
}

/* ============================================================
 * สรุปค่าแรงพนักงาน (Sub 1 STT / Sub 3 STT+KEMREX) — อ่านจากไฟล์เงินเดือน 1U6Zt
 *  scope = 'STT' | 'KEMREX' | 'ALL'
 *  รวมยอดรายเดือน: เงินเดือน / Total OT / สวัสดิการ / ค่าแรงรวม(ต้นทุนแรงงาน) + ด่านตรวจยอด(ต้องเป็น 0)
 * ============================================================ */
function r2_(x){return Math.round((Number(x)||0)*100)/100;}
function num_(v) { if (v == null || v === '') return 0; var n = Number(String(v).replace(/[,\s]/g, '')); return isNaN(n) ? 0 : n; }
function exactCol_(H, name) { for (var c = 0; c < H.length; c++) if (H[c] === name) return c; return -1; }
function findCol_(H, cands) { for (var c = 0; c < H.length; c++) { for (var k = 0; k < cands.length; k++) if (H[c].indexOf(cands[k]) >= 0) return c; } return -1; }
function cacheGet_(k){try{var c=CacheService.getScriptCache().get(k);return c?JSON.parse(c):null;}catch(e){return null;}}
function cachePut_(k,o){try{var s=JSON.stringify(o);if(s.length<95000)CacheService.getScriptCache().put(k,s,120);}catch(e){}}
var MAJOR_ORDER = ['HR','Account','Design','After Sale Service','Production','Purchase','QC','Sale','Safety','Estimate','Dcc & Admin','(อื่นๆ)'];
function majorFromCode_(code){var s=String(code||'').trim();var m2={'11':'Estimate','12':'Dcc & Admin'};if(m2[s.substring(0,2)])return m2[s.substring(0,2)];var m1={'1':'HR','2':'Account','3':'Design','4':'After Sale Service','5':'Production','6':'Purchase','7':'QC','8':'Sale','9':'Safety'};return m1[s.charAt(0)]||'(อื่นๆ)';}

function getSalarySummary(scope, force) {
  try {
    scope = scope || 'STT';
    var CK='salsum_v4_2026_'+scope; if(!force){var cx=cacheGet_(CK); if(cx)return cx;}
    var reg = getRegistry_(2026);
    if (!reg.payrollActual) return { ok: false, message: 'ไม่พบไฟล์ Payroll Actual ใน Registry' };
    var ss = SpreadsheetApp.openById(reg.payrollActual), shs = ss.getSheets();
    function pick(kemrex) { for (var i = 0; i < shs.length; i++) { var n = shs[i].getName(); if (n.indexOf('DATA BASE') >= 0) { var isKem = n.toUpperCase().indexOf('KEMREX') >= 0; if (kemrex === isKem) return shs[i]; } } return null; }
    var sttSh = pick(false), kemSh = pick(true), tabs = [];
    if (scope === 'STT') tabs = [sttSh]; else if (scope === 'KEMREX') tabs = [kemSh]; else tabs = [sttSh, kemSh];
    tabs = tabs.filter(function (x) { return x; });
    if (!tabs.length) return { ok: false, message: 'ไม่พบแท็บ DATA BASE (scope=' + scope + ')' };

    var M = {}; for (var m = 1; m <= 12; m++) M[m] = { salary: 0, ot: 0, welfare: 0, labor: 0, chk1: 0, chk2: 0, rows: 0 };
    var di = { Direct: { salary:0,ot:0,welfare:0,other:0,labor:0,n:0 }, Indirect: { salary:0,ot:0,welfare:0,other:0,labor:0,n:0 } };
    var majAgg = {};
    var srcNames = [];
    for (var t = 0; t < tabs.length; t++) {
      var sh = tabs[t]; srcNames.push(sh.getName());
      var v = sh.getDataRange().getValues();
      // หาแถวหัวตารางจริง = แถวที่มีคำว่า "เลขเดือน" (กันแถวหัวกลุ่มด้านบน เช่น Oth.Income/Benefit)
      var hr = -1;
      for (var r = 0; r < Math.min(15, v.length); r++) {
        var rj = v[r].map(function (x) { return String(x).trim(); });
        if (rj.indexOf('เลขเดือน') >= 0) { hr = r; break; }
      }
      if (hr < 0) return { ok: false, message: 'หาแถวหัวตาราง (เลขเดือน) ไม่เจอในแท็บ ' + sh.getName() };
      var H = v[hr].map(function (x) { return String(x).trim(); });
      var iMon = findCol_(H, ['เลขเดือน']);
      var iSal = exactCol_(H, 'เงินเดือน'); if (iSal < 0) iSal = findCol_(H, ['เงินเดือน']);
      var iOt = findCol_(H, ['Total OT']);
      var iWf = findCol_(H, ['สวัสดิการพนักงาน']);
      var iLb = findCol_(H, ['ต้นทุนแรงงานบริษัท (เงินเดือน+OT']);
      var iC1 = findCol_(H, ['กระทบ Oth.Income']);
      var iC2 = findCol_(H, ['กระทบ รวมรายได้']);
      var iDI = findCol_(H, ['Direct/Indirect (จำนวนคน)']);
      if (iDI < 0) iDI = findCol_(H, ['Direct/Indirect (%)']);
      if (iDI < 0) iDI = findCol_(H, ['ประเภท พนักงาน Direct / Indirect', 'Direct/Indirect']);
      var iDepC = findCol_(H, ['รหัสแผนก']);
      for (var r2 = hr + 1; r2 < v.length; r2++) {
        var row = v[r2]; var mo = Math.round(num_(row[iMon]));
        if (!(mo >= 1 && mo <= 12)) continue;
        var _s = num_(row[iSal]), _o = num_(row[iOt]), _w = num_(row[iWf]), _l = num_(row[iLb]);
        M[mo].salary += _s; M[mo].ot += _o; M[mo].welfare += _w; M[mo].labor += _l;
        if (iC1 >= 0) M[mo].chk1 += num_(row[iC1]); if (iC2 >= 0) M[mo].chk2 += num_(row[iC2]);
        M[mo].rows++;
        var _dt = String(iDI >= 0 ? row[iDI] : '').toLowerCase().indexOf('indirect') >= 0 ? 'Indirect' : 'Direct';
        var _db = di[_dt]; _db.salary += _s; _db.ot += _o; _db.welfare += _w; _db.other += (_l - _s - _o - _w); _db.labor += _l; _db.n++;
        var _mk = majorFromCode_(String(iDepC >= 0 ? row[iDepC] : '').trim());
        if (!majAgg[_mk]) majAgg[_mk] = { name: _mk, salary: 0, ot: 0, welfare: 0, labor: 0, n: 0 };
        var _mb = majAgg[_mk]; _mb.salary += _s; _mb.ot += _o; _mb.welfare += _w; _mb.labor += _l; _mb.n++;
      }
    }
    var TH = ['', 'ม.ค', 'ก.พ', 'มี.ค', 'เม.ย', 'พ.ค', 'มิ.ย', 'ก.ค', 'ส.ค', 'ก.ย', 'ต.ค', 'พ.ย', 'ธ.ค'];
    var months = [], totals = { salary: 0, ot: 0, welfare: 0, other: 0, labor: 0, chk1: 0, chk2: 0 };
    for (var mm = 1; mm <= 12; mm++) {
      var b = M[mm]; if (b.rows === 0) continue;
      var other = Math.round((b.labor - b.salary - b.ot - b.welfare) * 100) / 100;
      months.push({ m: mm, name: TH[mm], salary: b.salary, ot: b.ot, welfare: b.welfare, other: other, labor: b.labor,
        otPct: b.labor ? Math.round(b.ot / b.labor * 1000) / 10 : 0, chk1: b.chk1, chk2: b.chk2, rows: b.rows });
      totals.salary += b.salary; totals.ot += b.ot; totals.welfare += b.welfare; totals.other += other; totals.labor += b.labor; totals.chk1 += b.chk1; totals.chk2 += b.chk2;
    }
    totals.otPct = totals.labor ? Math.round(totals.ot / totals.labor * 1000) / 10 : 0;
    var chkPass = Math.abs(totals.chk1) <= 1 && Math.abs(totals.chk2) <= 1;
    var OUT = { ok: true, scope: scope, source: srcNames.join(' + '), file: ss.getName(),
      months: months, totals: totals,
      majors: (function(){var a=Object.keys(majAgg).sort(function(x,y){var ix=MAJOR_ORDER.indexOf(x),iy=MAJOR_ORDER.indexOf(y);return (ix<0?99:ix)-(iy<0?99:iy);}).map(function(k){var b=majAgg[k];return {name:b.name,n:b.n,salary:r2_(b.salary),ot:r2_(b.ot),welfare:r2_(b.welfare),labor:r2_(b.labor),pct:0};});var tot=a.reduce(function(s,x){return s+x.labor;},0);a.forEach(function(x){x.pct=tot?Math.round(x.labor/tot*1000)/10:0;});return a;})(),
      di: (function(){var o={};['Direct','Indirect'].forEach(function(k){var b=di[k];o[k]={salary:r2_(b.salary),ot:r2_(b.ot),welfare:r2_(b.welfare),other:r2_(b.other),labor:r2_(b.labor),n:b.n,pct:totals.labor?Math.round(b.labor/totals.labor*1000)/10:0};});return o;})(),
      verify: { chk1: Math.round(totals.chk1 * 100) / 100, chk2: Math.round(totals.chk2 * 100) / 100, chkPass: chkPass, otherTotal: Math.round(totals.other * 100) / 100 } };
    cachePut_(CK, OUT); return OUT;
  } catch (e) { return { ok: false, message: String(e) }; }
}

/** ทดสอบใน Apps Script: รันแล้วดู Log ว่ายอดรวม/ด่านตรวจผ่านไหม */
function testSalarySummary() {
  var d = getSalarySummary('STT');
  if (!d.ok) { Logger.log('❌ ' + d.message); return; }
  Logger.log('📊 STT ' + d.file + ' (' + d.source + ')');
  d.months.forEach(function (r) { Logger.log(r.name + ': เงินเดือน ' + Math.round(r.salary).toLocaleString() + ' | OT ' + Math.round(r.ot).toLocaleString() + ' | สวัสดิการ ' + Math.round(r.welfare).toLocaleString() + ' | ค่าแรงรวม ' + Math.round(r.labor).toLocaleString() + ' | %OT ' + r.otPct + ' | diff ' + r.diff); });
  Logger.log('รวมทั้งปี: ค่าแรงรวม ' + Math.round(d.totals.labor).toLocaleString() + ' | ตรวจ: ' + (d.verify.chkPass ? '✔ เช็คยอดไฟล์=0' : '⚠ เช็คยอดไฟล์ ' + d.verify.chk1 + ',' + d.verify.chk2));
  return d;
}

/** เจาะดูรายเดือน (รายคน) + สรุปตามแผนกใหญ่(หน่วยงาน)→แผนกย่อย — scope STT/KEMREX/ALL, month 1..12 */
function getSalaryMonth(scope, month, force) {
  try {
    scope = scope || 'STT'; month = Math.round(Number(month)) || 0;
    if (!(month >= 1 && month <= 12)) return { ok: false, message: 'เดือนไม่ถูกต้อง' };
    var CK='salmon_v4_2026_'+scope+'_'+month; if(!force){var cx=cacheGet_(CK); if(cx)return cx;}
    var reg = getRegistry_(2026);
    if (!reg.payrollActual) return { ok: false, message: 'ไม่พบไฟล์ Payroll Actual' };
    var ss = SpreadsheetApp.openById(reg.payrollActual), shs = ss.getSheets();
    function pick(k) { for (var i = 0; i < shs.length; i++) { var n = shs[i].getName(); if (n.indexOf('DATA BASE') >= 0) { var isK = n.toUpperCase().indexOf('KEMREX') >= 0; if (k === isK) return shs[i]; } } return null; }
    var tabs = scope === 'STT' ? [pick(false)] : scope === 'KEMREX' ? [pick(true)] : [pick(false), pick(true)];
    tabs = tabs.filter(function (x) { return x; });
    if (!tabs.length) return { ok: false, message: 'ไม่พบแท็บ DATA BASE' };
    var HEAD = ['รหัส', 'ชื่อพนักงาน', 'แผนก', 'แผนกใหญ่', 'ประเภท', 'เงินเดือน', 'OT', 'สวัสดิการ', 'อื่นๆ', 'ค่าแรงรวม'];
    var rows = [], majors = {}, tot = { salary: 0, ot: 0, welfare: 0, other: 0, labor: 0 };
    for (var t = 0; t < tabs.length; t++) {
      var sh = tabs[t], v = sh.getDataRange().getValues(), hr = -1;
      for (var r = 0; r < Math.min(15, v.length); r++) { if (v[r].map(function (x) { return String(x).trim(); }).indexOf('เลขเดือน') >= 0) { hr = r; break; } }
      if (hr < 0) continue;
      var H = v[hr].map(function (x) { return String(x).trim(); });
      var iMon = findCol_(H, ['เลขเดือน']), iCode = findCol_(H, ['รหัสพนักงาน']), iName = findCol_(H, ['ชื่อพนักงาน']),
        iDepC = findCol_(H, ['รหัสแผนก']), iDepN = exactCol_(H, 'แผนก'),
        iUnitC = findCol_(H, ['รหัสหน่วยงาน']), iUnitN = exactCol_(H, 'ชื่อหน่วยงาน'),
        iDI = (findCol_(H,['Direct/Indirect (จำนวนคน)'])>=0?findCol_(H,['Direct/Indirect (จำนวนคน)']):(findCol_(H,['Direct/Indirect (%)'])>=0?findCol_(H,['Direct/Indirect (%)']):findCol_(H,['ประเภท พนักงาน Direct / Indirect']))),
        iSal = exactCol_(H, 'เงินเดือน'), iOt = findCol_(H, ['Total OT']), iWf = findCol_(H, ['สวัสดิการพนักงาน']), iLb = findCol_(H, ['ต้นทุนแรงงานบริษัท (เงินเดือน+OT']);
      if (iSal < 0) iSal = findCol_(H, ['เงินเดือน']);
      if (iDepN < 0) iDepN = findCol_(H, ['แผนก']);
      if (iUnitN < 0) iUnitN = findCol_(H, ['ชื่อหน่วยงาน']);
      for (var r2 = hr + 1; r2 < v.length; r2++) {
        var row = v[r2]; if (Math.round(num_(row[iMon])) !== month) continue;
        var s = num_(row[iSal]), o = num_(row[iOt]), w = num_(row[iWf]), l = num_(row[iLb]), oth = Math.round((l - s - o - w) * 100) / 100;
        var depC = String(iDepC >= 0 ? row[iDepC] : '').trim(), depN = String(iDepN >= 0 ? row[iDepN] : '').trim();
        var unitC = String(iUnitC >= 0 ? row[iUnitC] : '').trim(), unitN = String(iUnitN >= 0 ? row[iUnitN] : '').trim();
        var majKey = majorFromCode_(depC);
        rows.push([String(row[iCode] || '').trim(), String(row[iName] || '').trim(), (depN || depC || '(ไม่ระบุ)'), majKey, String(row[iDI] || '').trim(), r2_(s), r2_(o), r2_(w), r2_(oth), r2_(l), String(row[iDI]||'').trim()]);
        tot.salary += s; tot.ot += o; tot.welfare += w; tot.other += oth; tot.labor += l;
        if (!majors[majKey]) majors[majKey] = { name: majKey, code: unitC, n: 0, salary: 0, ot: 0, welfare: 0, other: 0, labor: 0, subs: {} };
        var mj = majors[majKey]; mj.n++; mj.salary += s; mj.ot += o; mj.welfare += w; mj.other += oth; mj.labor += l;
        var subKey = depC + '|' + depN;
        if (!mj.subs[subKey]) mj.subs[subKey] = { code: depC, name: depN || '(ไม่ระบุ)', n: 0, salary: 0, ot: 0, welfare: 0, other: 0, labor: 0 };
        var sb = mj.subs[subKey]; sb.n++; sb.salary += s; sb.ot += o; sb.welfare += w; sb.other += oth; sb.labor += l;
      }
    }
    var mArr = Object.keys(majors).sort(function(a,b){var ia=MAJOR_ORDER.indexOf(a),ib=MAJOR_ORDER.indexOf(b);return (ia<0?99:ia)-(ib<0?99:ib);}).map(function (k) {
      var m = majors[k];
      var subs = Object.keys(m.subs).sort().map(function (sk) { var s = m.subs[sk]; return { code: s.code, name: s.name, n: s.n, salary: r2_(s.salary), ot: r2_(s.ot), welfare: r2_(s.welfare), other: r2_(s.other), labor: r2_(s.labor) }; });
      return { name: m.name, code: m.code, n: m.n, salary: r2_(m.salary), ot: r2_(m.ot), welfare: r2_(m.welfare), other: r2_(m.other), labor: r2_(m.labor), subs: subs };
    });
    var OUTM = { ok: true, scope: scope, month: month, headers: HEAD, rows: rows, total: rows.length, majors: mArr,
      totals: { salary: r2_(tot.salary), ot: r2_(tot.ot), welfare: r2_(tot.welfare), other: r2_(tot.other), labor: r2_(tot.labor) } };
    cachePut_(CK, OUTM); return OUTM;
  } catch (e) { return { ok: false, message: String(e) }; }
}

/** เทียบรายปี — อ่านแท็บ HR&ACC จากไฟล์เงินเดือน (HR ใส่ยอดบัญชีรายเดือน/รายปี ไว้ที่นี่) */
function getYearCompare() {
  try {
    var reg = getRegistry_(2026);
    if (!reg.payrollActual) return { ok: false, message: 'ไม่พบไฟล์ Payroll Actual' };
    var ss = SpreadsheetApp.openById(reg.payrollActual), shs = ss.getSheets(), sh = null;
    for (var i = 0; i < shs.length; i++) { var n = shs[i].getName().toUpperCase().replace(/\s/g, ''); if (n.indexOf('HR&ACC') >= 0 || n.indexOf('HR&AC') >= 0) { sh = shs[i]; break; } }
    if (!sh) for (var j = 0; j < shs.length; j++) { var u = shs[j].getName().toUpperCase(); if (u.indexOf('HR') >= 0 && u.indexOf('ACC') >= 0) { sh = shs[j]; break; } }
    if (!sh) { var names = shs.map(function (s) { return s.getName(); }); return { ok: false, message: 'ไม่พบแท็บ HR&ACC ในไฟล์เงินเดือน\nแท็บที่มี: ' + names.join(', ') }; }
    var d = readSheet_(sh, null, null, 3000);
    return { ok: true, headers: d.headers, rows: d.rows, total: d.total, tab: sh.getName(), file: ss.getName() };
  } catch (e) { return { ok: false, message: String(e) }; }
}
