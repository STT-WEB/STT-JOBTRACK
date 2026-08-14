/***********************************************************
 * STT NOVA-HR Hub — Cal / Payroll reader (ตัวอ่านกลาง)
 * ▶ getCalTab(month, tabName)   = อ่านแท็บใดก็ได้จากไฟล์ Cal เดือนนั้น
 * ▶ getPayrollActual(month)     = อ่านเงินเดือนจ่ายจริง จากไฟล์ Payroll Actual (สรุปตารางเงินเดือน)
 * คืน {ok, headers[], rows[][], total, tab, file} — ไม่ต้องเปิด Google Sheets
 ***********************************************************/

/** แถวนี้เป็นหัวตารางไหม? (>=3 ช่องไม่ว่าง และมีตัวหนังสือ(ไม่ใช่เลขล้วน) >=2) — กันจับแถวเลข 1,2,3 ผิด */
function isHeaderRow_(row) {
  var nonEmpty = 0, text = 0;
  for (var k = 0; k < Math.min(row.length, 45); k++) {
    var s = String(row[k]).trim();
    if (s) { nonEmpty++; if (isNaN(Number(s))) text++; }
  }
  return nonEmpty >= 3 && text >= 2;
}

/** อ่านช่วงข้อมูลจาก sheet ที่ให้มา → {headers, rows, total} (หาหัวตารางเอง + แปลงวันที่) */
function readSheet_(sh, filterCol, filterVal, maxRows) {
  var v = sh.getDataRange().getValues();
  if (!v.length) return { headers: [], rows: [], total: 0 };
  var hr = 0;
  for (var r = 0; r < Math.min(10, v.length); r++) { if (isHeaderRow_(v[r])) { hr = r; break; } }
  var headers = v[hr].map(function (x) { return String(x).trim(); });
  var lastCol = 0; for (var k = 0; k < headers.length; k++) if (headers[k]) lastCol = k;
  headers = headers.slice(0, lastCol + 1);
  var mc = -1;
  if (filterCol) for (var c = 0; c < headers.length; c++) if (headers[c].indexOf(filterCol) >= 0) { mc = c; break; }
  var rows = [], cap = maxRows || 1500;
  for (var r2 = hr + 1; r2 < v.length; r2++) {
    var row = v[r2].slice(0, lastCol + 1);
    if (row.every(function (x) { return String(x).trim() === ''; })) continue;
    if (mc >= 0 && String(row[mc]).trim() !== String(filterVal)) continue;
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

/** อ่านเงินเดือนจ่ายจริง (Payroll Actual) ของเดือน month จากไฟล์ในทะเบียน (payrollActual = สรุปตารางเงินเดือน) */
function getPayrollActual(month) {
  try {
    var reg = getRegistry_(2026);
    if (!reg.payrollActual) return { ok: false, message: 'ไม่พบไฟล์ Payroll Actual ใน Registry' };
    var ss = SpreadsheetApp.openById(reg.payrollActual), shs = ss.getSheets(), sh = null;
    for (var i = 0; i < shs.length; i++) if (shs[i].getName().indexOf('DATA BASE') >= 0) { sh = shs[i]; break; }
    if (!sh) sh = shs[0];
    var d = readSheet_(sh, 'เลขเดือน', month, 2000);
    return { ok: true, headers: d.headers, rows: d.rows, total: d.total, tab: sh.getName(), file: ss.getName() };
  } catch (e) { return { ok: false, message: String(e) }; }
}
