/**
 * ============================================================================
 *  Apply.gs — ลง template ใหม่ให้แท็บงวด (ขั้นที่ 1–3 ของแผน)
 *
 *  ทำอะไรบ้าง
 *    1) สำเนาแท็บทั้งแท็บเก็บไว้เป็น _backup ก่อน  ← ทำเป็นอย่างแรกเสมอ
 *    2) เพิ่มคอลัมน์ Y–AT ต่อท้าย (ไม่ย้าย ไม่ลบ A–X)
 *    3) คำนวณใหม่ทั้งงวดด้วยกฎ v5 แล้วเขียนลงคอลัมน์ใหม่
 *    4) ใส่สูตรให้ V · W · E · ธงต่าง ๆ
 *
 *  ไม่ทำอะไร
 *    ✗ ไม่แตะโค้ด JOBTRACK   (นั่นคือขั้น 4–6 ส่งรอบหน้า)
 *    ✗ ไม่ลบคอลัมน์เดิม A–X  (Cal Job Cost / NOVA-HR Hub จึงไม่พัง)
 *    ✗ ไม่เรียงแถวใหม่        (JOBTRACK ใช้ลำดับแถวหาคนที่ยัง Check In ค้าง)
 *
 *  ใช้เครื่องคิดชั่วโมงตัวเดียวกับ Recheck.gs (rcCalc_) ที่ผ่านด่าน 0 มาแล้ว
 *  1368/1369 แถว — ต้องมี Recheck.gs อยู่ในโปรเจกต์ด้วย
 *
 *  ลำดับการใช้
 *    ① runApply()            ลง template + คำนวณใหม่
 *    ② lockColumns()         ล็อกคอลัมน์คำนวณ (รันแยก จะได้ย้อนง่าย)
 *    ③ installRecalcTrigger() ให้แก้เวลาแล้วคำนวณใหม่อัตโนมัติ
 *  ถ้าไม่ชอบผล → undoApply() เอาแท็บ _backup กลับมา
 * ============================================================================
 */

var AP_VER = 'apply v1.0 (2569-08-25)';

/* ตำแหน่งคอลัมน์ใหม่ (1-based) */
var AC_ = {
  Y:25,  Z:26,          /* เวลาที่ HR แก้เป็น */
  IN:27, OUT:28, ADJ:29,                    /* เวลาที่ระบบใช้ + ปัดไปกี่นาที */
  OT_AM:30, N_AM:31, LUNCH:32, N_PM:33, OT_PM:34, OT_LUNCH:35,   /* แตกตามช่วงวัน */
  OK:36, BY:37,                             /* อนุมัติ OT ผ่าเที่ยง */
  X1:38, OT1:39, OT15:40, OT2:41, OT3:42,   /* แตกตามตัวคูณ */
  STAT:43, DAYSUM:44, VS8:45, NOTE:46       /* ธง */
};
var AP_HEAD = [
  'เวลาเข้า (แก้เป็น)','เวลาออก (แก้เป็น)',
  'เวลาเข้า (ระบบใช้)','เวลาออก (ระบบใช้)','ปัดไป (นาที)',
  'OT เช้า ก่อน 08:00','ปกติ เช้า 08–12','พักเที่ยง (ตัดออก)','ปกติ บ่าย 13–17','OT เย็น 17:00+','OT ผ่าเที่ยง',
  '✓ อนุมัติ OT ผ่าเที่ยง','ผู้อนุมัติ · วันที่',
  'ชม. ทำงานปกติ','ชม. OT ×1 (วันหยุด รายเดือน)','ชม. OT ×1.5 (OT วันทำงาน)','ชม. OT ×2 (วันหยุด รายวัน)','ชม. OT ×3 (OT วันหยุด)',
  'สถานะเวลา','ชม.รวมทั้งวันของคนนี้','เทียบเกณฑ์ 8 ชม.','หมายเหตุแก้เวลา'
];

/* ---------------------------------------------------------------- ตัวช่วย */
function apSheet_() {
  var sh = SpreadsheetApp.openById(RC.LOG_ID).getSheetByName(RC.TAB);
  if (!sh) throw new Error('ไม่พบแท็บ ' + RC.TAB);
  return sh;
}
/* อ่านเวลาจากเซลล์ "แก้เป็น" ถ้ามี ไม่งั้นใช้เวลาสแกน */
function apPick_(fix, scan) {
  var s = (fix instanceof Date) ? fix : String(fix || '').trim();
  return (s === '' ) ? scan : fix;
}

/* ============================================================ ① ลง template */
function runApply() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var sh = apSheet_();

  /* --- 1. สำรองก่อน --- */
  var bkName = RC.TAB + '_backup';
  var k = 1;
  while (ss.getSheetByName(bkName) && k < 50) { k++; bkName = RC.TAB + '_backup' + k; }
  if (k >= 50) throw new Error('มีแท็บสำรองเกิน 50 อัน — ลบของเก่าทิ้งก่อน');
  sh.copyTo(ss).setName(bkName);
  Logger.log('สำรองแท็บเป็น "' + bkName + '" เรียบร้อย');

  /* --- 2. ขยายคอลัมน์ให้พอถึง AT (46) --- */
  if (sh.getMaxColumns() < AC_.NOTE) sh.insertColumnsAfter(sh.getMaxColumns(), AC_.NOTE - sh.getMaxColumns());
  sh.getRange(1, AC_.Y, 1, AP_HEAD.length).setValues([AP_HEAD]).setFontWeight('bold');
  sh.getRange(1, 24).setValue('ชม.คิดค่าแรง (เลิกใช้)');   /* X — ไม่มีใครอ่าน เก็บค่าเดิมไว้เฉย ๆ */

  /* --- 3. คำนวณใหม่ทั้งงวด --- */
  var v = sh.getDataRange().getValues();
  var calMap = rcLoadCal_(), C = RC.C;
  var last = v.length;
  var block = [];                       /* Y..AT ของทุกแถว */
  var S = { rows:0, flag:0, cross:0 };

  for (var i = 1; i < last; i++) {
    var row = v[i];
    /* ★ เริ่มจากค่าที่ HR กรอกไว้เดิมเสมอ — รัน runApply ซ้ำกี่รอบก็ไม่ลบของ HR */
    var line = new Array(AP_HEAD.length).fill('');
    line[AC_.Y    - AC_.Y] = row[AC_.Y    - 1] || '';   /* เวลาเข้า (แก้เป็น) */
    line[AC_.Z    - AC_.Y] = row[AC_.Z    - 1] || '';   /* เวลาออก (แก้เป็น) */
    line[AC_.OK   - AC_.Y] = row[AC_.OK   - 1] || '';   /* ✓ อนุมัติ */
    line[AC_.BY   - AC_.Y] = row[AC_.BY   - 1] || '';   /* ผู้อนุมัติ */
    line[AC_.NOTE - AC_.Y] = row[AC_.NOTE - 1] || '';   /* หมายเหตุ */

    if (String(row[C.STATUS]).trim() !== 'Check Out') { block.push(line); continue; }
    S.rows++;

    var n = Number(row[C.PCOUNT]) || 1; if (n < 1) n = 1;
    var isMonthly = String(row[C.TYPE] || '').indexOf('รายเดือน') >= 0;

    /* เวลาที่ใช้ = "แก้เป็น" ถ้ามี ไม่งั้นเวลาสแกน */
    var fixIn  = row[AC_.Y - 1],  fixOut = row[AC_.Z - 1];
    var useIn  = apPick_(fixIn,  row[C.IN]);
    var useOut = apPick_(fixOut, row[C.OUT]);

    var tick = String(row[AC_.OK - 1] || '').trim() !== '';
    var r = rcCalc_(useIn, useOut, row[C.DATE], isMonthly, calMap, tick, row[C.DAYTYPE]);

    var rawIn = rcMin_(useIn), rawOut = rcMin_(useOut);
    var adj = '';
    if (!r.err && rawIn >= 0 && rawOut >= 0) {
      var dIn  = rcMin_(r.inUse)  - rawIn;
      var dOut = rcMin_(r.outUse) - rawOut;
      adj = 'เข้า ' + (dIn > 0 ? '+' : '') + dIn + ' · ออก ' + (dOut > 0 ? '+' : '') + dOut;
    }

    var H = function (m) { return rcDec_(m / n); };

    line[AC_.IN  - AC_.Y] = r.err ? '' : r.inUse;
    line[AC_.OUT - AC_.Y] = r.err ? '' : r.outUse;
    line[AC_.ADJ - AC_.Y] = adj;

    line[AC_.OT_AM   - AC_.Y] = H(r.amOT   || 0);
    line[AC_.N_AM    - AC_.Y] = H(r.amWork || 0);
    line[AC_.LUNCH   - AC_.Y] = r.lunch ? -H(r.lunch) : 0;
    line[AC_.N_PM    - AC_.Y] = H(r.pmWork || 0);
    line[AC_.OT_PM   - AC_.Y] = H(r.pmOT   || 0);
    line[AC_.OT_LUNCH- AC_.Y] = H(r.lunchOT|| 0);

    line[AC_.X1   - AC_.Y] = H(r.hNormal);
    line[AC_.OT1  - AC_.Y] = H(r.ot1);
    line[AC_.OT15 - AC_.Y] = H(r.ot15);
    line[AC_.OT2  - AC_.Y] = H(r.ot2);
    line[AC_.OT3  - AC_.Y] = H(r.ot3);

    var stat;
    if (r.err)                stat = '⛔ ' + r.err;
    else if (rawOut < 0)      stat = '⛔ ไม่มีเวลาออก';
    else if (r.crossLunch && !tick) stat = '⚠ คร่อมเที่ยง — รออนุมัติ';
    else if (H(r.ot15 + r.ot3) > 12) stat = '⚠ OT สูงผิดปกติ';
    else                      stat = 'ปกติ';
    if (stat.charAt(0) === '⛔' || stat.charAt(0) === '⚠') S.flag++;
    if (r.crossLunch) S.cross++;
    line[AC_.STAT - AC_.Y] = stat;

    block.push(line);
  }

  if (block.length) sh.getRange(2, AC_.Y, block.length, AP_HEAD.length).setValues(block);

  /* --- 4. ใส่สูตรให้คอลัมน์ที่ตรวจด้วยตาได้ --- */
  var nR = last - 1;
  if (nR > 0) {
    var fV = [], fW = [], fE = [], fSum = [], fVs = [];
    for (var r2 = 2; r2 <= last; r2++) {
      fV.push(['=IF($Q' + r2 + '<>"Check Out","",AL' + r2 + '+AM' + r2 + '+AO' + r2 + ')']);
      fW.push(['=IF($Q' + r2 + '<>"Check Out","",AN' + r2 + '+AP' + r2 + ')']);
      fE.push(['=IF($Q' + r2 + '<>"Check Out","",V' + r2 + '+W' + r2 + ')']);
      fSum.push(['=IF($Q' + r2 + '<>"Check Out","",SUMIFS($E:$E,$H:$H,$H' + r2 + ',$B:$B,$B' + r2 + '))']);
      fVs.push(['=IF(AR' + r2 + '="","",IF(ABS(AR' + r2 + '-8)<=0.17,"✅ ครบ",' +
                'IF(AR' + r2 + '<8,"⚠ ขาด "&TEXT(8-AR' + r2 + ',"0.00"),"ℹ เกิน "&TEXT(AR' + r2 + '-8,"0.00"))))']);
    }
    sh.getRange(2, 22, nR, 1).setFormulas(fV);        /* V ชม.ทำงานปกติ */
    sh.getRange(2, 23, nR, 1).setFormulas(fW);        /* W ชม.OT */
    sh.getRange(2,  5, nR, 1).setFormulas(fE);        /* E ชม.รวม */
    sh.getRange(2, AC_.DAYSUM, nR, 1).setFormulas(fSum);
    sh.getRange(2, AC_.VS8,    nR, 1).setFormulas(fVs);
    sh.getRange(2,  5, nR, 1).setNumberFormat('0.00');
    sh.getRange(2, 22, nR, 2).setNumberFormat('0.00');
    sh.getRange(2, AC_.OT_AM, nR, 6).setNumberFormat('0.00');
    sh.getRange(2, AC_.X1,    nR, 5).setNumberFormat('0.00');
    sh.getRange(2, AC_.DAYSUM, nR, 1).setNumberFormat('0.00');
  }

  /* --- 5. ย้อมสีช่องที่ HR กรอกเอง --- */
  var nRows = Math.max(1, nR);
  sh.getRange(2, AC_.Y,    nRows, 2).setBackground('#E6F4EA');   /* Y · Z */
  sh.getRange(2, AC_.OK,   nRows, 2).setBackground('#E6F4EA');   /* AJ · AK */
  sh.getRange(2, AC_.NOTE, nRows, 1).setBackground('#E6F4EA');   /* AT */
  sh.getRange(1, AC_.Y, 1, AP_HEAD.length).setWrap(true);
  sh.setFrozenRows(1);

  var msg = '\n===== ' + AP_VER + ' =====' +
    '\nแท็บ            : ' + RC.TAB +
    '\nสำรองไว้ที่     : ' + bkName +
    '\nแถว Check Out   : ' + S.rows +
    '\nแถวที่ขึ้นธง     : ' + S.flag +
    '\nแถวคร่อมเที่ยง   : ' + S.cross +
    '\nคอลัมน์ใหม่      : Y ถึง AT (' + AP_HEAD.length + ' คอลัมน์)' +
    '\nช่องที่ HR กรอก  : Y · Z · AJ · AK · AT (พื้นเขียว)' +
    '\n\nขั้นต่อไป : รัน lockColumns() แล้วตามด้วย installRecalcTrigger()' +
    '\nถ้าไม่ชอบผล : รัน undoApply() เพื่อเอาแท็บ ' + bkName + ' กลับมา' +
    '\n=========================================\n';
  Logger.log(msg);
  return msg;
}

/* ====================================================== ② ล็อกคอลัมน์คำนวณ */
function lockColumns() {
  var sh = apSheet_();
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
    if (p.getDescription() === 'NOVA lock') p.remove();
  });
  var p = sh.protect().setDescription('NOVA lock');
  var last = Math.max(2, sh.getLastRow());
  p.setUnprotectedRanges([
    sh.getRange(2, AC_.Y,    last - 1, 2),   /* Y · Z  เวลาที่แก้เป็น */
    sh.getRange(2, AC_.OK,   last - 1, 2),   /* AJ · AK อนุมัติ */
    sh.getRange(2, AC_.NOTE, last - 1, 1)    /* AT หมายเหตุ */
  ]);
  Logger.log('ล็อกแล้ว — แก้ได้เฉพาะ Y · Z · AJ · AK · AT\n' +
             'ถ้าจะปลดล็อก: ข้อมูล → ชีตและช่วงที่มีการป้องกัน → ลบ "NOVA lock"');
  return 'locked';
}

/* ============================================ ③ แก้เวลาแล้วคำนวณใหม่อัตโนมัติ */
function installRecalcTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditRecalc') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditRecalc')
    .forSpreadsheet(SpreadsheetApp.openById(RC.LOG_ID)).onEdit().create();
  Logger.log('ติดตั้งแล้ว — แก้ช่อง Y · Z · AJ แถวไหน แถวนั้นคำนวณใหม่ทันที');
  return 'installed';
}

function onEditRecalc(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== RC.TAB) return;
    var col = e.range.getColumn(), row = e.range.getRow();
    if (row < 2) return;
    if (col !== AC_.Y && col !== AC_.Z && col !== AC_.OK) return;
    apRecalcRow_(sh, row);
  } catch (err) { Logger.log('onEditRecalc: ' + err); }
}

/** คำนวณใหม่ 1 แถว — ใช้ตอนแก้มือ */
function apRecalcRow_(sh, row) {
  var C = RC.C;
  var v = sh.getRange(row, 1, 1, AC_.NOTE).getValues()[0];
  if (String(v[C.STATUS]).trim() !== 'Check Out') return;

  var calMap = rcLoadCal_();
  var n = Number(v[C.PCOUNT]) || 1; if (n < 1) n = 1;
  var isMonthly = String(v[C.TYPE] || '').indexOf('รายเดือน') >= 0;
  var useIn  = apPick_(v[AC_.Y - 1], v[C.IN]);
  var useOut = apPick_(v[AC_.Z - 1], v[C.OUT]);
  var tick = String(v[AC_.OK - 1] || '').trim() !== '';

  var r = rcCalc_(useIn, useOut, v[C.DATE], isMonthly, calMap, tick, v[C.DAYTYPE]);
  var rawIn = rcMin_(useIn), rawOut = rcMin_(useOut);
  var adj = '';
  if (!r.err && rawIn >= 0 && rawOut >= 0)
    adj = 'เข้า ' + (rcMin_(r.inUse) - rawIn) + ' · ออก ' + (rcMin_(r.outUse) - rawOut);

  var H = function (m) { return rcDec_(m / n); };
  var stat;
  if (r.err) stat = '⛔ ' + r.err;
  else if (r.crossLunch && !tick) stat = '⚠ คร่อมเที่ยง — รออนุมัติ';
  else if (H(r.ot15 + r.ot3) > 12) stat = '⚠ OT สูงผิดปกติ';
  else stat = 'ปกติ';

  sh.getRange(row, AC_.IN, 1, 3).setValues([[r.err ? '' : r.inUse, r.err ? '' : r.outUse, adj]]);
  sh.getRange(row, AC_.OT_AM, 1, 6).setValues([[
    H(r.amOT || 0), H(r.amWork || 0), r.lunch ? -H(r.lunch) : 0,
    H(r.pmWork || 0), H(r.pmOT || 0), H(r.lunchOT || 0)]]);
  sh.getRange(row, AC_.X1, 1, 5).setValues([[
    H(r.hNormal), H(r.ot1), H(r.ot15), H(r.ot2), H(r.ot3)]]);
  sh.getRange(row, AC_.STAT).setValue(stat);
}

/* ================================================================ ย้อนกลับ */
function undoApply() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var names = ss.getSheets().map(function (s) { return s.getName(); })
                .filter(function (nm) { return nm.indexOf(RC.TAB + '_backup') === 0; }).sort();
  if (!names.length) throw new Error('ไม่พบแท็บสำรอง — ยังไม่เคยรัน runApply()');
  Logger.log('พบแท็บสำรอง: ' + names.join(', ') +
    '\n\nวิธีย้อนกลับ (ทำมือ 3 ขั้น เพื่อไม่ให้สคริปต์ลบข้อมูลเอง):' +
    '\n  1) เปลี่ยนชื่อแท็บ ' + RC.TAB + ' เป็น ' + RC.TAB + '_ทิ้ง' +
    '\n  2) เปลี่ยนชื่อ ' + names[0] + ' เป็น ' + RC.TAB +
    '\n  3) ตรวจว่าข้อมูลครบ แล้วค่อยลบแท็บ _ทิ้ง' +
    '\n\nผมตั้งใจไม่ให้สคริปต์ลบแท็บเอง — ลบผิดแล้วกู้ไม่ได้');
  return names;
}

/* ============================================================================
 *  ④ normalizeDates() — แก้ต้นตอ "กุญแจวันที่ไม่ตรง" ของข้อมูลเก่า
 *
 *  ต้นตอ : JOBTRACK เขียนวันที่เป็นข้อความ "27/7/2569" ผ่าน appendRow
 *          แล้ว Google Sheets แปลงให้เองเป็น Date ที่ปี ค.ศ. = 2569
 *          (ห่างจากปีจริง 543 ปี) → กุญแจเทียบกับปฏิทินไม่มีวันตรง
 *
 *  ผลข้างเคียงที่เจอเพิ่ม : updateDailyHourAlert ใน JOBTRACK เทียบวันที่แบบ
 *          String(row[COL.DATE]) === dateStr ซึ่งไม่มีวันตรงเหมือนกัน
 *          → หมายเหตุ "ครบ 8 ชม." จึงนับได้แค่แถวปัจจุบันแถวเดียวมาตลอด
 *
 *  ตัวนี้ทำ : แปลงคอลัมน์ B ทุกแถวให้เป็น "ข้อความ d/m/พ.ศ." รูปแบบเดียว
 *  ต้องรัน runApply() (ซึ่งสำรองแท็บให้แล้ว) ก่อนเสมอ
 * ========================================================================== */
function normalizeDates() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var hasBackup = ss.getSheets().some(function (s) {
    return s.getName().indexOf(RC.TAB + '_backup') === 0;
  });
  if (!hasBackup) throw new Error('ยังไม่มีแท็บสำรอง — รัน runApply() ก่อน แล้วค่อยรันตัวนี้');

  var sh = apSheet_();
  var last = sh.getLastRow();
  if (last < 2) return 'ไม่มีข้อมูล';

  var rng = sh.getRange(2, RC.C.DATE + 1, last - 1, 1);
  var v = rng.getValues();
  var out = [], changed = 0, sample = [], kinds = { date:0, text:0, blank:0 };

  for (var i = 0; i < v.length; i++) {
    var cur = v[i][0];
    if (cur === '' || cur === null) { kinds.blank++; out.push(['']); continue; }
    if (cur instanceof Date) kinds.date++; else kinds.text++;
    var key = rcKey_(cur);
    if (String(cur) !== key) {
      changed++;
      if (sample.length < 6) sample.push('  แถว ' + (i + 2) + ' : ' +
        String(cur).substring(0, 32) + '  →  ' + key);
    }
    out.push([key]);
  }
  /* ตั้งรูปแบบเป็นข้อความก่อน แล้วค่อยเขียน ไม่งั้น Sheets จะแปลงกลับเป็นวันที่อีก */
  rng.setNumberFormat('@STRING@').setValues(out);

  var msg = '\n===== normalizeDates =====' +
    '\nแถวทั้งหมด        : ' + (last - 1) +
    '\n  เดิมเป็น Date    : ' + kinds.date +
    '\n  เดิมเป็นข้อความ  : ' + kinds.text +
    '\n  ว่าง             : ' + kinds.blank +
    '\nแก้ไป             : ' + changed + ' แถว' +
    '\n\nตัวอย่างที่แก้:\n' + (sample.join('\n') || '  ไม่มี') +
    '\n\nตอนนี้คอลัมน์ B เป็นข้อความรูปแบบเดียวกันหมดแล้ว' +
    '\nกุญแจเทียบปฏิทิน · SUMIFS รายวัน · หมายเหตุครบ 8 ชม. จะทำงานถูกต้อง' +
    '\nขั้นต่อไป: รัน runApply() อีกรอบ เพื่อคำนวณใหม่ด้วยประเภทวันที่ถูกต้อง' +
    '\n==========================\n';
  Logger.log(msg);
  return msg;
}
