/**
 * ============================================================================
 *  Apply.gs — ลง template ใหม่ให้แท็บงวด  (ฉบับย่อ ตามที่เบียร์สั่ง)
 *
 *  คอลัมน์ที่เพิ่ม มีแค่ 9 ช่อง
 *    Y   เวลาเข้า (ปัดแล้ว)      ← ตัวที่เอาไปคิด Cost
 *    Z   เวลาออก (ปัดแล้ว)
 *    AA  ชม. OT ×1              วันหยุด · รายเดือน
 *    AB  ชม. OT ×1.5            OT วันทำงาน
 *    AC  ชม. OT ×2              วันหยุด · รายวัน
 *    AD  ชม. OT ×3              OT ในวันหยุด
 *    AE  ☑ OT ผ่าเที่ยง          ช่องติ๊กจริง คลิกได้เลย
 *    AF  สถานะ                  ธงบอกแถวที่ต้องแก้
 *    AG  หมายเหตุ                HR พิมพ์เอง
 *
 *  คอลัมน์เดิมที่กลายเป็นสูตร
 *    V   ชม.ทำงานปกติ
 *    W   ชม.OT รวม   = AA+AB+AC+AD
 *    E   ชม.รวม      = V+W        ← ดูว่าวันนั้นทำงานกี่ชั่วโมง
 *
 *  ช่องที่คนแก้ได้ : C · D (เวลาเข้า–ออก) · AE (ติ๊ก) · AG (หมายเหตุ)
 *
 *  ✗ ไม่แตะโค้ด JOBTRACK   ✗ ไม่ลบคอลัมน์เดิม   ✗ ไม่เรียงแถวใหม่
 *
 *  ลำดับ : runApply() → normalizeDates() → runApply() อีกรอบ
 *          พอใจแล้วค่อย lockColumns() + installRecalcTrigger()
 *          ไม่พอใจ → undoApply()
 * ============================================================================
 */

var AP_VER = 'apply v2.0 (2569-08-25)';

/* ตำแหน่งคอลัมน์ (1-based) */
var AC_ = {
  IN:25, OUT:26,                    /* Y  Z   เวลาที่ปัดแล้ว */
  OT1:27, OT15:28, OT2:29, OT3:30,  /* AA–AD ชั่วโมงแยกตามตัวคูณ */
  TICK:31, STAT:32, NOTE:33         /* AE AF AG */
};
var AP_HEAD = [
  'เวลาเข้า (ปัดแล้ว)', 'เวลาออก (ปัดแล้ว)',
  'ชม. OT ×1', 'ชม. OT ×1.5', 'ชม. OT ×2', 'ชม. OT ×3',
  '☑ OT ผ่าเที่ยง', 'สถานะ', 'หมายเหตุ'
];
var AP_N = AP_HEAD.length;          /* 9 */

function apSheet_() {
  var sh = SpreadsheetApp.openById(RC.LOG_ID).getSheetByName(RC.TAB);
  if (!sh) throw new Error('ไม่พบแท็บ ' + RC.TAB);
  return sh;
}
function apTicked_(v) {
  return v === true || String(v).trim() === '✓' || String(v).toUpperCase() === 'TRUE';
}

/* ============================================================ ① ลง template */
function runApply() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var sh = apSheet_();

  /* 1. สำรองก่อนเสมอ */
  var bk = RC.TAB + '_backup', k = 1;
  while (ss.getSheetByName(bk) && k < 50) { k++; bk = RC.TAB + '_backup' + k; }
  if (k >= 50) throw new Error('แท็บสำรองเกิน 50 อัน — ลบของเก่าก่อน');
  sh.copyTo(ss).setName(bk);

  /* 2. หัวตาราง */
  if (sh.getMaxColumns() < AC_.NOTE) sh.insertColumnsAfter(sh.getMaxColumns(), AC_.NOTE - sh.getMaxColumns());
  sh.getRange(1, AC_.IN, 1, AP_N).setValues([AP_HEAD]).setFontWeight('bold').setWrap(true);
  sh.getRange(1, 22).setValue('ชม.ทำงานปกติ');
  sh.getRange(1, 23).setValue('ชม.OT รวม');
  sh.getRange(1,  5).setValue('ชม.รวม (ปกติ+OT)');
  sh.getRange(1, 24).setValue('ชม.คิดค่าแรง (เลิกใช้)');

  /* 3. คำนวณใหม่ทั้งงวด */
  var v = sh.getDataRange().getValues();
  var calMap = rcLoadCal_(), C = RC.C;
  var last = v.length, nR = last - 1;
  var block = [], vNormal = [];
  var S = { rows:0, flag:0, cross:0, err:0 };

  for (var i = 1; i < last; i++) {
    var row = v[i];
    var tick = apTicked_(row[AC_.TICK - 1]);
    var line = ['', '', '', '', '', '', tick, '', row[AC_.NOTE - 1] || ''];

    if (String(row[C.STATUS]).trim() !== 'Check Out') {
      block.push(line); vNormal.push(['']); continue;
    }
    S.rows++;

    var n = Number(row[C.PCOUNT]) || 1; if (n < 1) n = 1;
    var isMonthly = String(row[C.TYPE] || '').indexOf('รายเดือน') >= 0;
    var r = rcCalc_(row[C.IN], row[C.OUT], row[C.DATE], isMonthly, calMap, tick, row[C.DAYTYPE]);
    var H = function (m) { return rcDec_(m / n); };

    var stat;
    if (r.err)                       stat = '⛔ ' + r.err;
    else if (r.crossLunch && !tick)  stat = '⚠ คร่อมเที่ยง — รอติ๊ก';
    else if (H(r.ot15 + r.ot3) > 12) stat = '⚠ OT สูงผิดปกติ';
    else                             stat = '';
    if (stat) S.flag++;
    if (r.err) S.err++;
    if (r.crossLunch) S.cross++;

    line[0] = r.err ? '' : r.inUse;
    line[1] = r.err ? '' : r.outUse;
    line[2] = H(r.ot1);
    line[3] = H(r.ot15);
    line[4] = H(r.ot2);
    line[5] = H(r.ot3);
    line[7] = stat;
    block.push(line);
    vNormal.push([r.err ? '' : H(r.hNormal)]);
  }

  if (block.length) {
    sh.getRange(2, AC_.IN, block.length, AP_N).setValues(block);
    sh.getRange(2, 22, vNormal.length, 1).setValues(vNormal);       /* V ชม.ทำงานปกติ */
  }

  /* 4. สูตร W และ E — ตรวจด้วยตาได้ */
  if (nR > 0) {
    var fW = [], fE = [];
    for (var r2 = 2; r2 <= last; r2++) {
      fW.push(['=IF($Q' + r2 + '<>"Check Out","",AA' + r2 + '+AB' + r2 + '+AC' + r2 + '+AD' + r2 + ')']);
      fE.push(['=IF($Q' + r2 + '<>"Check Out","",V' + r2 + '+W' + r2 + ')']);
    }
    sh.getRange(2, 23, nR, 1).setFormulas(fW);
    sh.getRange(2,  5, nR, 1).setFormulas(fE);
    sh.getRange(2,  5, nR, 1).setNumberFormat('0.00');
    sh.getRange(2, 22, nR, 2).setNumberFormat('0.00');
    sh.getRange(2, AC_.OT1, nR, 4).setNumberFormat('0.00');

    /* ★ ช่องติ๊กจริง คลิกได้เลย ไม่ต้องพิมพ์ */
    var tickRange = sh.getRange(2, AC_.TICK, nR, 1);
    tickRange.insertCheckboxes();

    /* ย้อมช่องที่คนแก้ได้ */
    sh.getRange(2,  3, nR, 2).setBackground('#E6F4EA');   /* C · D เวลาเข้า–ออก */
    tickRange.setBackground('#E6F4EA');
    sh.getRange(2, AC_.NOTE, nR, 1).setBackground('#E6F4EA');
  }
  sh.setFrozenRows(1);

  var msg = '\n===== ' + AP_VER + ' =====' +
    '\nแท็บ           : ' + RC.TAB +
    '\nสำรองไว้ที่    : ' + bk +
    '\nแถว Check Out  : ' + S.rows +
    '\nแถวขึ้นธง      : ' + S.flag + '   (⛔ คิดไม่ได้ ' + S.err + ')' +
    '\nแถวคร่อมเที่ยง  : ' + S.cross + '  ← ติ๊กช่อง AE ได้เลย' +
    '\n\nคอลัมน์ใหม่ : Y–AG (9 ช่อง)' +
    '\nคนแก้ได้    : C · D เวลาเข้า–ออก | AE ติ๊ก | AG หมายเหตุ  (พื้นเขียว)' +
    '\n\nขั้นต่อไป : normalizeDates() → runApply() อีกรอบ' +
    '\nไม่พอใจ   : undoApply()' +
    '\n==============================\n';
  Logger.log(msg);
  return msg;
}

/* ====================================================== ② ล็อกคอลัมน์คำนวณ */
function lockColumns() {
  var sh = apSheet_();
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
    if (p.getDescription() === 'NOVA lock') p.remove();
  });
  var last = Math.max(2, sh.getLastRow());
  var p = sh.protect().setDescription('NOVA lock');
  p.setUnprotectedRanges([
    sh.getRange(2, 3,        last - 1, 2),   /* C · D  เวลาเข้า–ออก */
    sh.getRange(2, AC_.TICK, last - 1, 1),   /* AE ติ๊ก */
    sh.getRange(2, AC_.NOTE, last - 1, 1)    /* AG หมายเหตุ */
  ]);
  Logger.log('ล็อกแล้ว — แก้ได้เฉพาะ C · D · AE · AG\n' +
             'ปลดล็อก: ข้อมูล → ชีตและช่วงที่มีการป้องกัน → ลบ "NOVA lock"');
  return 'locked';
}

/* ============================================ ③ แก้แล้วคำนวณใหม่อัตโนมัติ */
function installRecalcTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditRecalc') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditRecalc')
    .forSpreadsheet(SpreadsheetApp.openById(RC.LOG_ID)).onEdit().create();
  Logger.log('ติดตั้งแล้ว — แก้เวลาเข้า–ออก หรือติ๊ก OT ผ่าเที่ยง แถวนั้นคำนวณใหม่ทันที');
  return 'installed';
}

function onEditRecalc(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== RC.TAB) return;
    var col = e.range.getColumn(), row = e.range.getRow();
    if (row < 2) return;
    if (col !== 3 && col !== 4 && col !== AC_.TICK) return;   /* C · D · AE */
    apRecalcRow_(sh, row);
  } catch (err) { Logger.log('onEditRecalc: ' + err); }
}

function apRecalcRow_(sh, row) {
  var C = RC.C;
  var v = sh.getRange(row, 1, 1, AC_.NOTE).getValues()[0];
  if (String(v[C.STATUS]).trim() !== 'Check Out') return;

  var n = Number(v[C.PCOUNT]) || 1; if (n < 1) n = 1;
  var isMonthly = String(v[C.TYPE] || '').indexOf('รายเดือน') >= 0;
  var tick = apTicked_(v[AC_.TICK - 1]);
  var r = rcCalc_(v[C.IN], v[C.OUT], v[C.DATE], isMonthly, rcLoadCal_(), tick, v[C.DAYTYPE]);
  var H = function (m) { return rcDec_(m / n); };

  var stat;
  if (r.err)                       stat = '⛔ ' + r.err;
  else if (r.crossLunch && !tick)  stat = '⚠ คร่อมเที่ยง — รอติ๊ก';
  else if (H(r.ot15 + r.ot3) > 12) stat = '⚠ OT สูงผิดปกติ';
  else                             stat = '';

  sh.getRange(row, AC_.IN, 1, 6).setValues([[
    r.err ? '' : r.inUse, r.err ? '' : r.outUse,
    H(r.ot1), H(r.ot15), H(r.ot2), H(r.ot3)]]);
  sh.getRange(row, AC_.STAT).setValue(stat);
  sh.getRange(row, 22).setValue(r.err ? '' : H(r.hNormal));
}

/* ==================================== ④ แก้วันที่เก่าให้เป็นมาตรฐานเดียวกัน */
function normalizeDates() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  if (!ss.getSheets().some(function (s) { return s.getName().indexOf(RC.TAB + '_backup') === 0; }))
    throw new Error('ยังไม่มีแท็บสำรอง — รัน runApply() ก่อน');

  var sh = apSheet_(), last = sh.getLastRow();
  if (last < 2) return 'ไม่มีข้อมูล';
  var rng = sh.getRange(2, RC.C.DATE + 1, last - 1, 1);
  var v = rng.getValues(), out = [], changed = 0, sample = [];

  for (var i = 0; i < v.length; i++) {
    var cur = v[i][0];
    if (cur === '' || cur === null) { out.push(['']); continue; }
    var key = rcKey_(cur);
    if (String(cur) !== key) {
      changed++;
      if (sample.length < 5) sample.push('  แถว ' + (i + 2) + ' : ' + String(cur).substring(0, 30) + '  →  ' + key);
    }
    out.push([key]);
  }
  rng.setNumberFormat('@STRING@').setValues(out);   /* ตั้งรูปแบบก่อน ไม่งั้น Sheets แปลงกลับ */

  Logger.log('\n===== normalizeDates =====' +
    '\nแถวทั้งหมด : ' + (last - 1) + '   แก้ไป : ' + changed +
    '\n' + (sample.join('\n') || '  ไม่มีอะไรต้องแก้') +
    '\n\nขั้นต่อไป : รัน runApply() อีกรอบ เพื่อคิดวันหยุดให้ถูก' +
    '\n==========================\n');
  return changed;
}

/* ================================================================ ย้อนกลับ */
function undoApply() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var names = ss.getSheets().map(function (s) { return s.getName(); })
    .filter(function (nm) { return nm.indexOf(RC.TAB + '_backup') === 0; }).sort();
  if (!names.length) throw new Error('ไม่พบแท็บสำรอง');
  Logger.log('แท็บสำรองที่มี: ' + names.join(', ') +
    '\n\nวิธีย้อนกลับ (ทำมือ 3 ขั้น — ผมตั้งใจไม่ให้สคริปต์ลบแท็บเอง):' +
    '\n  1) เปลี่ยนชื่อ ' + RC.TAB + ' → ' + RC.TAB + '_ทิ้ง' +
    '\n  2) เปลี่ยนชื่อ ' + names[0] + ' → ' + RC.TAB +
    '\n  3) ตรวจว่าข้อมูลครบ แล้วค่อยลบแท็บ _ทิ้ง');
  return names;
}
