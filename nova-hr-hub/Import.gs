/***********************************************************
 * STT NOVA-HR Hub — Import (โมดูลนำเข้าเดือนเก่า)
 * แยกไฟล์ต่างหากจาก Cost engine
 * ▶ importOldMaster()  = ดึงต้นทุน Direct ม.ค.–เม.ย. จากไฟล์ MASTER เก่า → JOBCOST 2026
 *   (อ้างอิงค่าคงที่ OLD_MASTER_ID / JOBCOST_FILE_ID + ตัวช่วย findCol_ ฯลฯ จากไฟล์ Code.gs)
 ***********************************************************/

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
    if (!job || job.indexOf('-') < 0) continue;                 // ข้ามหัว/แถวว่าง/ST
    var row = new Array(24).fill('');
    row[0] = job.split('-')[0];                                 // ประเภทงาน จาก prefix (JMC/JMI/JMT แยก)
    row[1] = job;
    row[2] = String(r[cName]).trim();
    var bf = (cBf >= 0) ? (Number(r[cBf]) || 0) : 0;
    row[15] = Math.round(bf);
    var sum = bf, any = false;
    for (var mm = 1; mm <= 12; mm++) {
      var v = (mcol[mm] >= 0) ? (Number(r[mcol[mm]]) || 0) : 0;
      row[2 + mm] = Math.round(v);                              // ม.ค.=index3 .. ธ.ค.=index14
      sum += v;
      if (v > 0) { any = true; monthsSeen[mm] = true; }
    }
    row[16] = Math.round(sum);                                  // รวมค่าแรง STT
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
