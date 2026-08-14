/***********************************************************
 * STT NOVA-HR Hub — Cal reader (ตัวอ่านกลาง)
 * ▶ getCalTab(month, tabName) = อ่านแท็บใดก็ได้จากไฟล์ Cal เดือนนั้น → {headers, rows}
 *   ใช้ร่วมกันทุกหน้า: Cal / Data Payroll / Performance / Time JOBTRACK
 *   (ไม่ต้องเปิด Google Sheets — โปรแกรมดึงมาแสดงเอง)
 ***********************************************************/

/** หา fileId ของไฟล์ Cal เดือนที่ระบุ จากโฟลเดอร์ Cal ใน Registry (เลือกตัว (NEW) ก่อน) */
function getMonthFileId_(month) {
  var reg = getRegistry_(2026);
  if (!reg.calFolder) throw new Error('ไม่พบ Cal Folder ใน Registry');
  var folder = DriveApp.getFolderById(reg.calFolder);
  var it = folder.getFiles(), pfx = String(month) + '.', found = null, foundNew = null;
  while (it.hasNext()) {
    var f = it.next(), n = f.getName();
    if (n.indexOf(pfx) === 0) { if (n.toUpperCase().indexOf('NEW') >= 0) foundNew = f.getId(); else if (!found) found = f.getId(); }
  }
  return foundNew || found;
}

/** อ่านแท็บ tabName จากไฟล์ Cal เดือน month → {ok, headers[], rows[][], total, tab, file} */
function getCalTab(month, tabName) {
  try {
    var fid = getMonthFileId_(month);
    if (!fid) return { ok: false, message: 'ไม่พบไฟล์ Cal เดือน ' + month };
    var ss = SpreadsheetApp.openById(fid);
    var sh = ss.getSheetByName(tabName);
    if (!sh) { var shs = ss.getSheets(); for (var i = 0; i < shs.length; i++) if (shs[i].getName().toUpperCase().indexOf(String(tabName).toUpperCase()) >= 0) { sh = shs[i]; break; } }
    if (!sh) return { ok: false, message: 'ไม่พบแท็บ ' + tabName + ' ในไฟล์เดือน ' + month + ' (' + ss.getName() + ')' };
    var v = sh.getDataRange().getValues();
    if (!v.length) return { ok: true, headers: [], rows: [], total: 0, tab: sh.getName(), file: ss.getName() };
    // หาแถวหัวตาราง (แถวแรกที่มี >=3 ช่องไม่ว่าง ใน 6 แถวแรก)
    var hr = 0;
    for (var r = 0; r < Math.min(6, v.length); r++) { var c = 0; for (var k = 0; k < Math.min(v[r].length, 40); k++) if (String(v[r][k]).trim()) c++; if (c >= 3) { hr = r; break; } }
    var headers = v[hr].map(function (x) { return String(x).trim(); });
    var lastCol = 0; for (var k2 = 0; k2 < headers.length; k2++) if (headers[k2]) lastCol = k2;
    headers = headers.slice(0, lastCol + 1);
    var rows = [], maxR = Math.min(v.length, hr + 1 + 1500);
    for (var r2 = hr + 1; r2 < maxR; r2++) {
      var row = v[r2].slice(0, lastCol + 1);
      if (row.every(function (x) { return String(x).trim() === ''; })) continue;
      rows.push(row.map(function (x) { return (x instanceof Date) ? Utilities.formatDate(x, 'Asia/Bangkok', 'dd/MM/yyyy') : x; }));
    }
    return { ok: true, headers: headers, rows: rows, total: Math.max(0, v.length - hr - 1), tab: sh.getName(), file: ss.getName() };
  } catch (e) { return { ok: false, message: String(e) }; }
}
