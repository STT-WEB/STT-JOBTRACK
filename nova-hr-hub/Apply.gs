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

var AP_VER = 'apply v4.4 (2569-08-26)';

/* ตำแหน่งคอลัมน์ (1-based) */
var AC_ = {
  IN:25, OUT:26,                    /* Y  Z   เวลาที่ปัดแล้ว */
  AM:27, PM:28, WORK:29,            /* AA AB AC  ช่วงเช้า · ช่วงบ่าย · รวม */
  OT1:30, OT15:31, OT2:32, OT3:33,  /* AD–AG ชั่วโมงแยกตามตัวคูณ */
  TICK:34, STAT:35, NOTE:36         /* AH AI AJ */
};
var AP_HEAD = [
  'เวลาเข้า (ปัดแล้ว)', 'เวลาออก (ปัดแล้ว)',
  'ชม. ช่วงเช้า', 'ชม. ช่วงบ่าย', 'รวม ชม.ทำงาน',
  'ชม. OT ×1', 'ชม. OT ×1.5', 'ชม. OT ×2', 'ชม. OT ×3',
  '☑ OT ผ่าเที่ยง (แก้ได้)', 'สถานะ', 'หมายเหตุ (แก้ได้)'
];
var AP_N = AP_HEAD.length;          /* 9 */

function apSheet_() {
  var sh = SpreadsheetApp.openById(RC.LOG_ID).getSheetByName(rcTab_());
  if (!sh) throw new Error('ไม่พบแท็บ ' + rcTab_());
  return sh;
}
function apTicked_(v) {
  return v === true || String(v).trim() === '✓' || String(v).toUpperCase() === 'TRUE';
}

/* ============================================================ ① ลง template */
function runApply(opts) {
  opts = opts || {};
  var C0 = RC.C, dFixMsg = 0;
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var sh = apSheet_();

  /* 1. ไม่สำรองอัตโนมัติแล้ว — ถ้าอยากได้ ให้รัน backupNow() เอง */
  var bk = '(ปิดการสำรองอัตโนมัติ)';

  /* 2. หัวตาราง */
  if (sh.getMaxColumns() < AC_.NOTE) sh.insertColumnsAfter(sh.getMaxColumns(), AC_.NOTE - sh.getMaxColumns());

  /* ★ ล้างของค้างจากเวอร์ชันก่อน (คอลัมน์ AH เป็นต้นไป) ให้เกลี้ยง */
  var maxC = sh.getMaxColumns();
  if (maxC > AC_.NOTE) sh.getRange(1, AC_.NOTE + 1, sh.getMaxRows(), maxC - AC_.NOTE).clear();

  sh.getRange(1, AC_.IN, 1, AP_N).setValues([AP_HEAD]).setFontWeight('bold').setWrap(true);
  sh.getRange(1, 22).setValue('ชม.ทำงานปกติ');
  sh.getRange(1, 23).setValue('ชม.OT รวม');
  sh.getRange(1,  5).setValue('ชม.รวม (ปกติ+OT)');
  sh.getRange(1,  3).setValue('เวลาเข้า (แก้ได้)');
  sh.getRange(1,  4).setValue('เวลาออก (แก้ได้)');
  sh.getRange(1, 24).setValue('ชม.คิดค่าแรง (เลิกใช้)');
  /* ซ่อน ไม่ลบ — JOBTRACK เขียนคอลัมน์ที่ 24 อยู่ ถ้าลบแล้วตำแหน่งจะเลื่อนทั้งแถว */
  sh.hideColumns(24);

  /* คำอธิบายติดหัวคอลัมน์ — เอาเมาส์ชี้แล้วเห็นเลยว่าช่องไหนแก้ได้ */
  var EDIT_NOTE = '✏️ ช่องนี้แก้ได้\nพิมพ์เวลาใหม่ลงไปได้เลย ระบบจะคำนวณชั่วโมงใหม่ให้ทันที\n(ต้องรัน installRecalcTrigger ครั้งเดียวก่อน)';
  sh.getRange(1, 3).setNote(EDIT_NOTE);
  sh.getRange(1, 4).setNote(EDIT_NOTE);
  sh.getRange(1, AC_.TICK).setNote('✏️ ช่องนี้แก้ได้\nติ๊กเมื่อพนักงานทำงานผ่านช่วงพักเที่ยงจริง\nติ๊กแล้วได้ OT เพิ่ม 1 ชั่วโมงเต็ม');
  sh.getRange(1, AC_.NOTE).setNote('✏️ ช่องนี้แก้ได้\nเขียนว่าทำไมถึงแก้เวลา เช่น "ลืมสแกนออก หัวหน้ายืนยันแล้ว"');
  sh.getRange(1, 25).setNote('🔒 ผลคำนวณ — ห้ามแก้\nเวลาที่ระบบใช้จริงหลังปัดตามกฎ\nถ้าไม่ถูก ให้ไปแก้ช่อง "เวลาเข้า (แก้ได้)" แทน');
  sh.getRange(1, 26).setNote('🔒 ผลคำนวณ — ห้ามแก้\nถ้าไม่ถูก ให้ไปแก้ช่อง "เวลาออก (แก้ได้)" แทน');

  /* 2.5 ปกติวันที่ทุกรอบ — JOBTRACK ยังเขียนแถวใหม่เป็น Date ปี 2569 อยู่
         ถ้าไม่ปกติ ชนิดข้อมูลจะปนกัน แล้ว QUERY ในแท็บสรุปจะแยกคนคนเดียวเป็น 2 แถว */
  var lastR = sh.getLastRow();
  if (lastR > 1) {
    var dR = sh.getRange(2, C0.DATE + 1, lastR - 1, 1);
    var dV = dR.getValues(), dOut = [], dFix = 0;
    for (var z = 0; z < dV.length; z++) {
      var cur = dV[z][0];
      if (cur === '' || cur === null) { dOut.push(['']); continue; }
      var kk = rcKey_(cur);
      if (String(cur) !== kk) dFix++;
      dOut.push([kk]);
    }
    if (dFix) dR.setNumberFormat('@STRING@').setValues(dOut);
    dFixMsg = dFix;
  }

  /* 3. คำนวณใหม่ทั้งงวด */
  var v = sh.getDataRange().getValues();
  var calMap = rcLoadCal_(), C = RC.C;
  var last = v.length, nR = last - 1;
  var block = [];
  var S = { rows:0, flag:0, cross:0, err:0, split:0 };

  /* ★ นับจำนวนแถวจริงของแต่ละรอบ Check In (คอลัมน์ L = รหัสรอบงาน)
     ห้ามเชื่อเลขในคอลัมน์ P เพราะถ้าแถวคู่หายไป ชั่วโมงจะถูกหารหายโดยไม่มีใครรู้
     เช่น P=2 แต่มีแถวเดียว → เดิมหาร 2 ชั่วโมงหายครึ่งนึง */
  var seen = {};
  for (var q = 1; q < last; q++) {
    if (String(v[q][C.STATUS]).trim() !== 'Check Out') continue;
    var sid = String(v[q][C.SESSION] || '').trim();
    if (!sid) continue;
    seen[sid] = (seen[sid] || 0) + 1;
  }

  for (var i = 1; i < last; i++) {
    var row = v[i];
    var tick = apTicked_(row[AC_.TICK - 1]);
    var line = ['', '', '', '', '', '', '', '', '', tick, '', row[AC_.NOTE - 1] || ''];

    if (String(row[C.STATUS]).trim() !== 'Check Out') {
      block.push(line); continue;
    }
    S.rows++;

    /* หารด้วย "จำนวนแถวจริง" ของรอบงานนี้ ถ้าหาไม่เจอค่อยใช้เลขในคอลัมน์ P */
    var sid2 = String(row[C.SESSION] || '').trim();
    var pCol = Number(row[C.PCOUNT]) || 1; if (pCol < 1) pCol = 1;
    var n = (sid2 && seen[sid2]) ? seen[sid2] : pCol;
    if (n < 1) n = 1;
    if (n !== pCol) S.split++;

    var isMonthly = String(row[C.TYPE] || '').indexOf('รายเดือน') >= 0;
    var r = rcCalc_(row[C.IN], row[C.OUT], row[C.DATE], isMonthly, calMap, tick, row[C.DAYTYPE]);
    var H = function (m) { return rcDec_(m / n); };

    var stat;
    if (r.err)                       stat = '⛔ ' + r.err;
    else if (n !== pCol)             stat = '⚠ แถวซ้ำ — รอบนี้มี ' + n + ' แถว แต่คอลัมน์ P บอก ' + pCol;
    else if (r.crossLunch && !tick)  stat = '⚠ คร่อมเที่ยง — รอติ๊ก';
    else if (H(r.ot15 + r.ot3) > 12) stat = '⚠ OT สูงผิดปกติ';
    /* ไม่ใช่ปัญหา แต่บอกไว้ให้หายสงสัยว่าทำไมชั่วโมงน้อยกว่านาฬิกา */
    else if (r.minApplied)           stat = 'งานสั้น · คิดขั้นต่ำ 0.5 ชม.';
    else if (n > 1)                  stat = '✓ หารให้แล้ว · งานนี้ลง ' + n + ' จ๊อบ';
    else                             stat = '';
    if (stat.charAt(0) === '⛔' || stat.charAt(0) === '⚠') S.flag++;
    if (r.err) S.err++;
    if (r.crossLunch) S.cross++;

    var amH = H(r.amOT + r.amWork);                    /* ทุกชั่วโมงก่อนเที่ยง */
    var pmH = H(r.pmWork + r.pmOT + r.lunchOT);        /* ตั้งแต่เที่ยงเป็นต้นไป */
    line[0]  = r.err ? '' : r.inUse;
    line[1]  = r.err ? '' : r.outUse;
    line[2]  = amH;
    line[3]  = pmH;
    line[4]  = Math.round((amH + pmH) * 100) / 100;
    line[5]  = H(r.ot1);
    line[6]  = H(r.ot15);
    line[7]  = H(r.ot2);
    line[8]  = H(r.ot3);
    line[10] = stat;
    block.push(line);
  }

  if (block.length) sh.getRange(2, AC_.IN, block.length, AP_N).setValues(block);

  /* 4. สูตร W และ E — ตรวจด้วยตาได้ */
  if (nR > 0) {
    var fW = [], fE = [], fV = [];
    for (var r2 = 2; r2 <= last; r2++) {
      fW.push(['=IF($Q' + r2 + '<>"Check Out","",AD' + r2 + '+AE' + r2 + '+AF' + r2 + '+AG' + r2 + ')']);
      fE.push(['=IF($Q' + r2 + '<>"Check Out","",AC' + r2 + ')']);
      fV.push(['=IF($Q' + r2 + '<>"Check Out","",ROUND(AC' + r2 + '-W' + r2 + ',2))']);
    }
    sh.getRange(2, 23, nR, 1).setFormulas(fW);
    sh.getRange(2,  5, nR, 1).setFormulas(fE);
    sh.getRange(2, 22, nR, 1).setFormulas(fV);   /* V = รวมงาน − OT */
    sh.getRange(2,  5, nR, 1).setNumberFormat('0.00');
    sh.getRange(2, 22, nR, 2).setNumberFormat('0.00');
    sh.getRange(2, AC_.AM, nR, 7).setNumberFormat('0.00');

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
    '\nแท็บ           : ' + rcTab_() +
    (opts.auto ? '\n(รันอัตโนมัติโดยตัวตั้งเวลา)' : '') +
    '\nแถว Check Out  : ' + S.rows +
    '\nแถวขึ้นธง      : ' + S.flag + '   (⛔ คิดไม่ได้ ' + S.err + ')' +
    '\nแถวคร่อมเที่ยง  : ' + S.cross + '  ← ติ๊กช่อง AH ได้เลย' +
    '\nแถวที่คอลัมน์ P ไม่ตรงจำนวนแถวจริง : ' + S.split +
    (S.split ? '  ← เดิมชั่วโมงหายไป รอบนี้แก้ให้แล้ว' : '') +
    '\n\nคอลัมน์ใหม่ : Y–AJ (12 ช่อง)' +
    '\nคนแก้ได้    : C · D เวลาเข้า–ออก | AH ติ๊ก | AJ หมายเหตุ  (พื้นเขียว)' +
    (dFixMsg ? '\nปกติวันที่ให้ : ' + dFixMsg + ' แถว' : '') +
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
    if (sh.getName() !== rcTab_()) return;
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

  var n = apRowsInSession_(sh, String(v[C.SESSION] || '').trim(), Number(v[C.PCOUNT]) || 1);
  var isMonthly = String(v[C.TYPE] || '').indexOf('รายเดือน') >= 0;
  var tick = apTicked_(v[AC_.TICK - 1]);
  var r = rcCalc_(v[C.IN], v[C.OUT], v[C.DATE], isMonthly, rcLoadCal_(), tick, v[C.DAYTYPE]);
  var H = function (m) { return rcDec_(m / n); };

  var stat;
  if (r.err)                       stat = '⛔ ' + r.err;
  else if (r.crossLunch && !tick)  stat = '⚠ คร่อมเที่ยง — รอติ๊ก';
  else if (H(r.ot15 + r.ot3) > 12) stat = '⚠ OT สูงผิดปกติ';
  else if (r.minApplied)           stat = 'งานสั้น · คิดขั้นต่ำ 0.5 ชม.';
  else if (n > 1)                  stat = '✓ หารให้แล้ว · งานนี้ลง ' + n + ' จ๊อบ';
  else                             stat = '';

  var amH = H(r.amOT + r.amWork);
  var pmH = H(r.pmWork + r.pmOT + r.lunchOT);
  sh.getRange(row, AC_.IN, 1, 9).setValues([[
    r.err ? '' : r.inUse, r.err ? '' : r.outUse,
    amH, pmH, Math.round((amH + pmH) * 100) / 100,
    H(r.ot1), H(r.ot15), H(r.ot2), H(r.ot3)]]);
  sh.getRange(row, AC_.STAT).setValue(stat);
}

/* ==================================== ④ แก้วันที่เก่าให้เป็นมาตรฐานเดียวกัน */
function normalizeDates() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  if (!ss.getSheets().some(function (s) { return s.getName().indexOf(rcTab_() + '_backup') === 0; }))
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
    .filter(function (nm) { return nm.indexOf(rcTab_() + '_backup') === 0; }).sort();
  if (!names.length) throw new Error('ไม่พบแท็บสำรอง');
  Logger.log('แท็บสำรองที่มี: ' + names.join(', ') +
    '\n\nวิธีย้อนกลับ (ทำมือ 3 ขั้น — ผมตั้งใจไม่ให้สคริปต์ลบแท็บเอง):' +
    '\n  1) เปลี่ยนชื่อ ' + rcTab_() + ' → ' + rcTab_() + '_ทิ้ง' +
    '\n  2) เปลี่ยนชื่อ ' + names[0] + ' → ' + rcTab_() +
    '\n  3) ตรวจว่าข้อมูลครบ แล้วค่อยลบแท็บ _ทิ้ง');
  return names;
}


/* ============================================================================
 *  ⑤ buildDailySummary() — ตารางเทียบ "1 คน 1 วัน" (แบบ Pivot)
 *
 *  Job_Log มอง 1 แถว = 1 จ๊อบ  แต่ HR คิดเงินเป็น 1 คน 1 วัน
 *  คนหนึ่งอาจมี 5 แถวในวันเดียว จะดูว่าครบ 8 ชม.ไหมต้องบวกมือ
 *  แท็บนี้บวกให้ ด้วยสูตร QUERY ล้วน ไม่มีสคริปต์ ไม่กระทบข้อมูลเดิม
 *  ลบแท็บทิ้งได้ตลอด ของเดิมไม่พัง
 * ========================================================================== */
var AP_SUM_TAB  = 'สรุปรายคน–รายวัน';
var AP_PER_TAB  = 'สรุปรายคน–ทั้งงวด';

/* หัวตารางสรุป — ใช้ร่วมกันทั้งสองแท็บ */
var QQ = String.fromCharCode(34);      /* เครื่องหมาย " ในสูตร */
var CO = String.fromCharCode(39) + 'Check Out' + String.fromCharCode(39);
var AP_SUM_LBL =
  "label B 'วันที่', H 'รหัส', I 'ชื่อ', J 'แผนก', count(F) 'จ๊อบ', " +
  "sum(V) 'ชม.ทำงาน', sum(AD) 'OT×1', sum(AE) 'OT×1.5', sum(AF) 'OT×2', " +
  "sum(AG) 'OT×3', sum(W) 'รวม OT', sum(E) 'รวมทั้งวัน'";

function buildDailySummary() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var T  = "'" + rcTab_() + "'";

  /* ---------- ① รายคน–รายวัน ---------- */
  var sh = ss.getSheetByName(AP_SUM_TAB);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(AP_SUM_TAB, 0);
  sh.getRange('A1').setFormula(
    '=QUERY(' + T + '!A2:AJ,' + QQ
    + 'select B, H, I, J, count(F), sum(V), sum(AD), sum(AE), sum(AF), sum(AG), sum(W), sum(E) '
    + 'where Q = ' + CO + ' group by B, H, I, J order by B desc, H ' + AP_SUM_LBL + QQ + ',0)');
  sh.getRange('M1').setValue('เทียบ 8 ชม.');
  sh.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(L2:L="","",IF(ABS(L2:L-8)<=0.17,"ครบ",' +
    'IF(L2:L<8,"ขาด "&TEXT(8-L2:L,"0.00"),"เกิน "&TEXT(L2:L-8,"0.00")))))');
  apDressSummary_(sh, 13);

  /* ---------- ② รายคน–ทั้งงวด (ตัวที่เอาไปทำเงินเดือน) ---------- */
  var sp = ss.getSheetByName(AP_PER_TAB);
  if (sp) ss.deleteSheet(sp);
  sp = ss.insertSheet(AP_PER_TAB, 1);
  sp.getRange('A1').setValue('งวด ' + rcTab_().replace('Job_Log_', ''));
  var LBL2 = "label H 'รหัส', I 'ชื่อ', J 'แผนก', K 'ประเภท', count(F) 'จ๊อบ', "
    + "sum(V) 'ชม.ทำงาน', sum(AD) 'OT×1', sum(AE) 'OT×1.5', sum(AF) 'OT×2', "
    + "sum(AG) 'OT×3', sum(W) 'รวม OT', sum(E) 'รวมทั้งงวด'";
  sp.getRange('A2').setFormula(
    '=QUERY(' + T + '!A2:AJ,' + QQ
    + 'select H, I, J, K, count(F), sum(V), sum(AD), sum(AE), sum(AF), sum(AG), sum(W), sum(E) '
    + 'where Q = ' + CO + ' group by H, I, J, K order by H ' + LBL2 + QQ + ',0)');
  sp.getRange('M2').setValue('วันที่ทำงาน');
  sp.getRange('M3').setFormula(
    '=ARRAYFORMULA(IF(A3:A="","",COUNTUNIQUEIFS(' + T + '!$B$2:$B,' + T + '!$H$2:$H,A3:A,' +
    T + '!$Q$2:$Q,"Check Out")))');
  sp.getRange('A1').setFontWeight('bold').setFontSize(12);
  apDressSummary_(sp, 13, 2);

  Logger.log('สร้างแท็บสรุปแล้ว 2 แท็บ' +
    '\n  · ' + AP_SUM_TAB + '   1 แถว = 1 คน 1 วัน' +
    '\n  · ' + AP_PER_TAB + '  1 แถว = 1 คน ทั้งงวด  ← ตัวที่เอาไปทำเงินเดือน' +
    '\nอัปเดตเองทุกครั้งที่ Job_Log เปลี่ยน ไม่ต้องกดอะไร');
  return [AP_SUM_TAB, AP_PER_TAB];
}

/** จัดหน้าตาแท็บสรุปให้อ่านง่าย */
function apDressSummary_(sh, nCol, headRow) {
  headRow = headRow || 1;
  sh.setFrozenRows(headRow);
  sh.getRange(headRow, 1, 1, nCol).setFontWeight('bold').setBackground('#F2EDED').setWrap(true);
  sh.getRange(headRow + 1, 6, sh.getMaxRows() - headRow, 7).setNumberFormat('0.00');
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 150);
}

function apRowsInSession_(sh, sid, fallback) {
  if (!sid) return Math.max(1, fallback || 1);
  var C = RC.C, last = sh.getLastRow();
  var col = sh.getRange(2, C.SESSION + 1, last - 1, 1).getValues();
  var st  = sh.getRange(2, C.STATUS + 1,  last - 1, 1).getValues();
  var cnt = 0;
  for (var i = 0; i < col.length; i++)
    if (String(col[i][0]).trim() === sid && String(st[i][0]).trim() === 'Check Out') cnt++;
  return Math.max(1, cnt || fallback || 1);
}

/* ============================================================================
 *  ⑥ checkProcessRows() — ตรวจว่ามีแถวหายไหม (อ่านอย่างเดียว)
 *
 *  JOBTRACK ตอน Check Out ที่ลงหลาย Process จะ "เพิ่มแถวใหม่" ให้ Process ที่ 2, 3, ...
 *  ถ้าแถวพวกนั้นหายไป (เพิ่มไม่สำเร็จ หรือมีคนลบ) ชั่วโมงจะถูกหารหายเงียบ ๆ
 *  ตัวนี้ไล่นับให้ดูว่ามีกี่รอบงานที่แถวไม่ครบ
 * ========================================================================== */
function checkProcessRows() {
  var sh = apSheet_(), v = sh.getDataRange().getValues(), C = RC.C;
  var g = {};
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][C.STATUS]).trim() !== 'Check Out') continue;
    var sid = String(v[i][C.SESSION] || '').trim();
    if (!sid) continue;
    if (!g[sid]) g[sid] = { rows:0, p:Number(v[i][C.PCOUNT]) || 1,
                            emp:String(v[i][C.EMP]), name:String(v[i][C.NAME] || ''),
                            d:rcKey_(v[i][C.DATE]), t:String(v[i][C.IN]), first:i + 1 };
    g[sid].rows++;
  }
  var bad = [], okCnt = 0;
  for (var k in g) {
    if (g[k].rows === g[k].p) { okCnt++; continue; }
    bad.push(g[k]);
  }
  bad.sort(function (a, b) { return a.first - b.first; });

  var lines = bad.slice(0, 25).map(function (x) {
    return '  แถว ' + x.first + ' | ' + x.d + ' | ' + x.emp + ' ' + x.name.substring(0, 16) +
           ' | เข้า ' + x.t + ' | คอลัมน์ P บอก ' + x.p + ' แต่มีจริง ' + x.rows + ' แถว';
  });
  Logger.log('\n===== ตรวจแถวหาย =====' +
    '\nรอบงานทั้งหมด : ' + (okCnt + bad.length) +
    '\n  ครบ         : ' + okCnt +
    '\n  ไม่ครบ      : ' + bad.length +
    '\n' + (lines.join('\n') || '  ไม่มีปัญหา') +
    (bad.length > 25 ? '\n  ... และอีก ' + (bad.length - 25) + ' รอบ' : '') +
    '\n\nรอบที่ไม่ครบ = ชั่วโมงเคยถูกหารหายไป — runApply() รอบใหม่แก้ให้แล้ว' +
    '\n(หารด้วยจำนวนแถวจริง ไม่ใช่เลขในคอลัมน์ P)' +
    '\n=======================\n');
  return bad.length;
}


/* ============================================================================
 *  ⑦ ตัวรันอัตโนมัติ — แก้ปัญหา "งวดใหม่แล้วฟอร์มไม่ตามมา"
 *
 *  ทำไมต้องมี : ระบบมีจุดที่ onEdit ตามไม่ทัน 2 จุด
 *    ① JOBTRACK สร้างแท็บงวดใหม่เองทุกวันที่ 26 — แท็บนั้นเกิดมาไม่มีคอลัมน์ใหม่
 *    ② แถวที่ JOBTRACK เขียนตอน Check Out — onEdit ไม่ยิง เพราะสคริปต์เขียน ไม่ใช่คนพิมพ์
 *  ผลคือแถวใหม่ทุกแถวจะว่างจนกว่าจะมีคนรัน runApply มือ
 *
 *  ตัวนี้รันเองทุกชั่วโมง ลง template ให้แท็บงวดปัจจุบันเสมอ
 *  งวดใหม่มาเมื่อไหร่ ฟอร์มตามไปเองภายใน 1 ชั่วโมง ไม่ต้องรอใครกด
 * ========================================================================== */
function installAutoApply() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoApply') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoApply').timeBased().everyHours(1).create();

  /* ติดตั้งตัวคำนวณสดตอนแก้มือไปพร้อมกันเลย จะได้ไม่ลืม */
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditRecalc') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditRecalc')
    .forSpreadsheet(SpreadsheetApp.openById(RC.LOG_ID)).onEdit().create();

  var msg = 'ติดตั้งเรียบร้อย — ต่อจากนี้ไม่ต้องรันมืออีกเลย\n' +
    '  · ทุกชั่วโมง : ลง template + คำนวณแถวใหม่ให้แท็บงวดปัจจุบัน\n' +
    '  · งวดใหม่    : ฟอร์มตามไปเองภายใน 1 ชั่วโมง\n' +
    '  · แก้เวลามือ : คำนวณใหม่ทันที\n' +
    '  · ขึ้นงวดใหม่ : แช่แข็งสรุปงวดเก่าเก็บเป็นแท็บ สรุป_YYYY_MM ให้เอง\n' +
    '  · ไม่สำรองอัตโนมัติ (อยากได้ให้รัน backupNow เอง)';
  Logger.log(msg);
  return msg;
}

/** ตัวที่ตัวตั้งเวลาเรียก — ห้ามให้ error หลุดออกไป ไม่งั้น Google จะปิดตัวตั้งเวลาทิ้ง */
function autoApply() {
  try {
    runApply({ auto: true });
    /* ขึ้นงวดใหม่ → แช่แข็งสรุปงวดเก่าเก็บไว้ก่อน แล้วค่อยสร้างของงวดใหม่ */
    var ss2 = SpreadsheetApp.openById(RC.LOG_ID);
    var per = ss2.getSheetByName(AP_PER_TAB);
    var tag = 'งวด ' + rcTab_().replace('Job_Log_', '');
    var old = per ? String(per.getRange('A1').getValue()) : '';
    if (!ss2.getSheetByName(AP_SUM_TAB) || !per || old !== tag) {
      if (per && old && old !== tag) apArchive_(ss2, old.replace('งวด ', ''));
      buildDailySummary();
    }
  } catch (e) {
    Logger.log('autoApply ข้ามรอบนี้: ' + e.message);
  }
}

/** ดูว่ามีตัวตั้งเวลาอะไรทำงานอยู่บ้าง */
function listTriggers() {
  var t = ScriptApp.getProjectTriggers().map(function (x) {
    return '  · ' + x.getHandlerFunction() + '  [' + x.getEventType() + ']';
  });
  Logger.log('ตัวตั้งเวลาที่ทำงานอยู่:\n' + (t.join('\n') || '  ไม่มีเลย'));
  return t.length;
}


/** สำรองแท็บงวดปัจจุบันด้วยมือ — เรียกเองเมื่อต้องการเท่านั้น */
function backupNow() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var nm = rcTab_() + '_backup_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd_HHmm');
  apSheet_().copyTo(ss).setName(nm);
  Logger.log('สำรองเป็นแท็บ "' + nm + '" แล้ว');
  return nm;
}

/** ลบแท็บสำรองเก่าทั้งหมด — ยืนยันด้วยการพิมพ์ true ตอนเรียก */
function deleteAllBackups() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID), gone = [];
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('_backup') > 0) { gone.push(s.getName()); ss.deleteSheet(s); }
  });
  Logger.log(gone.length ? 'ลบแท็บสำรองแล้ว ' + gone.length + ' แท็บ:\n  ' + gone.join('\n  ')
                         : 'ไม่มีแท็บสำรองให้ลบ');
  return gone.length;
}


/* ============================================================================
 *  ⑧ แช่แข็งสรุปงวด — พอขึ้นงวดใหม่ ตัวเลขงวดเก่าจะไม่หายไปไหน
 *
 *  แท็บ "สรุปรายคน–ทั้งงวด" เป็นสูตร QUERY ที่ชี้แท็บงวดปัจจุบัน
 *  พอถึงวันที่ 26 มันจะเปลี่ยนไปชี้งวดใหม่ ตัวเลขงวดเก่าก็หายจากหน้าจอ
 *  ตัวนี้ก๊อปตัวเลข (ไม่เอาสูตร) เก็บเป็นแท็บ "สรุป_2569_09" ให้อัตโนมัติ
 *  ใช้อ้างอิงตอนจ่ายเงินเดือนย้อนหลังได้ตลอด
 * ========================================================================== */
function apArchive_(ss, tagOld) {
  if (!tagOld) return null;
  var src = ss.getSheetByName(AP_PER_TAB);
  if (!src) return null;

  var name = 'สรุป_' + tagOld;
  if (ss.getSheetByName(name)) return name;          /* เก็บไว้แล้ว ไม่ทำซ้ำ */

  var v = src.getDataRange().getValues();
  if (v.length < 3) return null;                     /* ยังไม่มีข้อมูล ไม่ต้องเก็บ */

  var t = ss.insertSheet(name, ss.getNumSheets());
  t.getRange(1, 1, v.length, v[0].length).setValues(v);   /* setValues = ได้ตัวเลข ไม่ติดสูตรมา */
  t.getRange(1, 1).setValue('สรุปงวด ' + tagOld + '  (ปิดงวดแล้ว · ตัวเลขแช่แข็ง)')
                  .setFontWeight('bold').setFontSize(12);
  t.setFrozenRows(2);
  t.getRange(2, 1, 1, v[0].length).setFontWeight('bold').setBackground('#F2EDED');
  t.getRange(3, 6, Math.max(1, v.length - 2), 7).setNumberFormat('0.00');
  t.setColumnWidth(2, 200);
  Logger.log('แช่แข็งสรุปงวด ' + tagOld + ' เป็นแท็บ "' + name + '" แล้ว');
  return name;
}

/** แช่แข็งสรุปงวดปัจจุบันด้วยมือ — ใช้ตอนจะปิดงวดก่อนกำหนด */
function archiveThisPeriod() {
  var ss = SpreadsheetApp.openById(RC.LOG_ID);
  var per = ss.getSheetByName(AP_PER_TAB);
  if (!per) throw new Error('ยังไม่มีแท็บสรุป — รัน buildDailySummary() ก่อน');
  var tag = String(per.getRange('A1').getValue()).replace('งวด ', '').trim();
  var nm = apArchive_(ss, tag);
  Logger.log(nm ? 'เก็บเป็นแท็บ "' + nm + '" แล้ว' : 'ไม่มีข้อมูลให้เก็บ');
  return nm;
}

/** ดูว่าเก็บสรุปงวดไหนไว้แล้วบ้าง */
function listArchives() {
  var a = SpreadsheetApp.openById(RC.LOG_ID).getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (n) { return n.indexOf('สรุป_') === 0; }).sort();
  Logger.log('สรุปงวดที่เก็บไว้แล้ว:\n  ' + (a.join('\n  ') || 'ยังไม่มี'));
  return a;
}
