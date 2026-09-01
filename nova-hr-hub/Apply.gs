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

var AP_VER = 'apply v6.2 (2569-09-01)';

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

/* บอกว่าอีก Process ของรอบงานเดียวกันไปอยู่แถวไหน
   จำเป็นเพราะ JOBTRACK เอาแถวที่ 2 ไปต่อท้ายชีต ไม่ได้วางติดกับแถวแรก */
function apMateMsg_(rows, self) {
  if (!rows || rows.length < 2) return '';
  var other = [];
  for (var i = 0; i < rows.length; i++) if (rows[i] !== self) other.push(rows[i]);
  if (!other.length) return '';
  return ' · อีกจ๊อบอยู่แถว ' + other.join(', ');
}


/* ============================================================================
 *  ป้าย "ประเภทชั่วโมง" คอลัมน์ U — ใช้ถ้อยคำชุดเดียวกับ JOBTRACK เป๊ะ
 *  ของเดิม JOBTRACK เขียนตอนสแกนแล้วจบ ไม่มีใครคำนวณใหม่ให้
 *  พอ HR แก้เวลาช่อง C/D ป้ายนี้จึงค้างเป็นของเก่า ไม่ตรงกับชั่วโมงจริง
 *
 *  จับคู่จากผลของเครื่องคำนวณ :
 *    hNormal → 1   ในกรอบ 08–17 วันทำงานปกติ
 *    ot1/ot2 → 2A หรือ 2B   ในกรอบ 08–17 แต่เป็นวันหยุด (รายเดือน ×1 · รายวัน ×2)
 *    ot15    → 3   นอกกรอบเวลา วันทำงานปกติ
 *    ot3     → 4   นอกกรอบเวลา วันหยุด
 * ========================================================================== */
/* ป้ายประเภทชั่วโมง — ใช้ตัวกลางใน Recheck.gs ไม่เก็บสำเนาไว้ที่นี่ */
function apHourType_(r, dayTypeVal) { return rcHourType_(r, dayTypeVal); }

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
  var S = { rows:0, flag:0, cross:0, err:0, split:0, collide:0 };
  var uCol = [];   /* ป้ายประเภทชั่วโมง คอลัมน์ U — คำนวณใหม่ทุกแถว */

  /* ★ นับจำนวนแถวจริงของแต่ละรอบ Check In (คอลัมน์ L = รหัสรอบงาน)
     ห้ามเชื่อเลขในคอลัมน์ P เพราะถ้าแถวคู่หายไป ชั่วโมงจะถูกหารหายโดยไม่มีใครรู้
     เช่น P=2 แต่มีแถวเดียว → เดิมหาร 2 ชั่วโมงหายครึ่งนึง

     ★★ ต้องผูกรหัสพนักงานเข้าไปในกุญแจด้วย ห้ามใช้รหัสรอบงานเดี่ยว ๆ
     JOBTRACK มีบั๊กแย่งแถวกันตอนสแกนพร้อมกัน (doCheckIn ไม่มี LockService)
     ทำให้รหัสรอบงานของคน A ไปตกอยู่บนแถวของคน B ได้
     ถ้าหารด้วยรหัสรอบงานเฉย ๆ ชั่วโมงของสองคนที่ไม่เกี่ยวกันจะโดนหารครึ่งทั้งคู่ */
  var seen = {}, sidEmp = {}, rowsOf = {};
  for (var q = 1; q < last; q++) {
    if (String(v[q][C.STATUS]).trim() !== 'Check Out') continue;
    var sid = String(v[q][C.SESSION] || '').trim();
    if (!sid) continue;
    var emp = String(v[q][C.EMP] || '').trim();
    var key = sid + '|' + emp;
    seen[key] = (seen[key] || 0) + 1;
    /* จำเลขแถวไว้ด้วย — JOBTRACK เอาแถว Process ที่ 2 ไปต่อท้ายชีต ห่างจากแถวแรกเป็นสิบแถว
       HR เปิดมาเห็นแถวเดียวแล้วนึกว่าอีก Process หายไป ต้องบอกเลขแถวคู่ให้ชัด */
    if (!rowsOf[key]) rowsOf[key] = [];
    rowsOf[key].push(q + 1);
    if (!sidEmp[sid]) sidEmp[sid] = {};
    sidEmp[sid][emp] = true;
  }
  var apOwners_ = function (sid) {
    if (!sid || !sidEmp[sid]) return 0;
    var c = 0; for (var e in sidEmp[sid]) c++;
    return c;
  };

  for (var i = 1; i < last; i++) {
    var row = v[i];
    var tick = apTicked_(row[AC_.TICK - 1]);
    var line = ['', '', '', '', '', '', '', '', '', tick, '', row[AC_.NOTE - 1] || ''];

    if (String(row[C.STATUS]).trim() !== 'Check Out') {
      block.push(line); uCol.push([row[20] || '']);   /* แถวที่ยังไม่ Check Out — คงป้ายเดิมไว้ */ continue;
    }
    S.rows++;

    /* หารด้วย "จำนวนแถวจริง" ของรอบงานนี้ ถ้าหาไม่เจอค่อยใช้เลขในคอลัมน์ P */
    var sid2 = String(row[C.SESSION] || '').trim();
    var emp2 = String(row[C.EMP] || '').trim();
    var key2 = sid2 + '|' + emp2;
    var pCol = Number(row[C.PCOUNT]) || 1; if (pCol < 1) pCol = 1;
    var n = (sid2 && seen[key2]) ? seen[key2] : pCol;
    if (n < 1) n = 1;
    if (n !== pCol) S.split++;
    /* รหัสรอบงานเดียวกันแต่มีหลายคน = JOBTRACK เขียนทับแถวกัน เวลาเข้าอาจไม่ใช่ของคนนี้ */
    var collide = apOwners_(sid2) > 1;
    if (collide) S.collide++;

    var isMonthly = String(row[C.TYPE] || '').indexOf('รายเดือน') >= 0;
    var r = rcCalc_(row[C.IN], row[C.OUT], row[C.DATE], isMonthly, calMap, tick, row[C.DAYTYPE]);
    var H = function (m) { return rcDec_(m / n); };

    /* ★ สถานะต้องบอก "ทุกเรื่อง" ของแถวนั้น ไม่ใช่เรื่องเดียวแล้วจบ
       ของเดิมเป็น if-else ต่อกัน แถวที่คร่อมเที่ยงจะขึ้นแค่ "คร่อมเที่ยง"
       แล้วกลบข้อความ "หารให้แล้ว 2 จ๊อบ" หายไป HR เลยนึกว่าระบบไม่ได้หารชั่วโมงให้
       ตอนนี้เก็บทุกข้อความมาต่อกัน เรื่องร้ายแรงขึ้นก่อน */
    var bad = [], info = [];
    if (r.err)            bad.push('⛔ ' + r.err);
    if (collide)          bad.push('⛔ รหัสรอบงานชนกับพนักงานคนอื่น — เวลาเข้าอาจไม่ใช่ของคนนี้ ให้ตรวจจากรูปสแกน');
    if (n !== pCol)       bad.push('⚠ แถวซ้ำ — รอบนี้มี ' + n + ' แถว แต่คอลัมน์ P บอก ' + pCol);
    if (H(r.ot15 + r.ot3) > 12) bad.push('⚠ OT สูงผิดปกติ');

    /* ไม่ใช่ปัญหา แค่บอกให้หายสงสัย — ไม่นับเป็นแถวติดธง ไม่ต้องมีใครมาตาม */
    if (r.crossLunch) {
      info.push(tick ? '☑ อนุมัติ OT ผ่าเที่ยงแล้ว · บวกคืนให้ 1 ชม.'
                     : 'ตัดพักเที่ยงให้แล้ว 1 ชม. · ถ้าทำงานจริงให้ติ๊กช่อง AH');
    }
    if (r.minApplied) info.push('งานสั้น · คิดขั้นต่ำ 0.5 ชม.');
    if (n > 1)        info.push('✓ หารให้แล้ว · งานนี้ลง ' + n + ' จ๊อบ' +
                                apMateMsg_(rowsOf[key2], i + 1));

    var stat = bad.concat(info).join('  |  ');
    if (bad.length) S.flag++;
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
    uCol.push([apHourType_(r, row[C.DAYTYPE])]);
  }

  if (block.length) sh.getRange(2, AC_.IN, block.length, AP_N).setValues(block);
  /* ★ เขียนป้ายประเภทชั่วโมง คอลัมน์ U ใหม่ทั้งงวด ให้ตรงกับชั่วโมงที่คิดได้จริง */
  if (uCol.length) sh.getRange(2, 21, uCol.length, 1).setValues(uCol);

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
    /* ★ ล้างโน้ต "ขาด/เกิน 8 ชม." ของเก่าที่ค้างอยู่บนคอลัมน์ V ให้เกลี้ยง
       เป็นภาพนิ่งที่ไม่อัปเดต ทำให้ HR เข้าใจผิดว่าชั่วโมงขาด ทั้งที่ไม่ได้ขาด
       ยอดจริงดูที่แท็บ สรุปรายคน–รายวัน ซึ่งคำนวณสดเสมอ */
    sh.getRange(2, 22, nR, 1).clearNote();

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
    '\nแถวคร่อมเที่ยง  : ' + S.cross + '  (ตัดพักเที่ยงให้แล้ว ไม่นับเป็นแถวติดธง · ถ้าทำงานจริงค่อยติ๊ก AH)' +
    '\nแถวที่คอลัมน์ P ไม่ตรงจำนวนแถวจริง : ' + S.split +
    (S.split ? '  ← เดิมชั่วโมงหายไป รอบนี้แก้ให้แล้ว' : '') +
    '\nแถวรหัสรอบงานชนกัน : ' + S.collide +
    (S.collide ? '  ← JOBTRACK เขียนทับแถวกันตอนสแกนพร้อมกัน ต้องตรวจเวลาเข้าจากรูปสแกน' : '') +
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
  var sh = null, row = 0;
  try {
    if (!e || !e.range) return;
    sh = e.range.getSheet();
    /* ★ ของเดิมเช็ก getName() !== rcTab_() แล้ว return ทันที
       แปลว่าถ้า HR แก้แถวในแท็บงวดเก่า (เช่น Job_Log_2569_08) จะไม่มีอะไรเกิดขึ้นเลย
       และไม่มีข้อความบอกด้วย — คนแก้จึงนึกว่าสูตรพัง
       ตอนนี้ให้ทำงานกับทุกแท็บที่ขึ้นต้นด้วย Job_Log_ */
    if (sh.getName().indexOf('Job_Log_') !== 0) return;
    var col = e.range.getColumn();
    row = e.range.getRow();
    if (row < 2) return;
    if (col !== 3 && col !== 4 && col !== AC_.TICK) return;   /* C · D · AH */
    apRecalcRow_(sh, row);
  } catch (err) {
    /* ★ ห้ามกลืน error เงียบ ๆ อีก — เขียนลงช่องสถานะให้คนแก้เห็นทันที */
    Logger.log('onEditRecalc แถว ' + row + ': ' + err);
    try {
      if (sh && row > 1) sh.getRange(row, AC_.STAT).setValue('⛔ คำนวณใหม่ไม่สำเร็จ : ' + err);
    } catch (e2) {}
  }
}

function apRecalcRow_(sh, row) {
  var C = RC.C;
  var v = sh.getRange(row, 1, 1, AC_.NOTE).getValues()[0];
  if (String(v[C.STATUS]).trim() !== 'Check Out') return;

  var n = apRowsInSession_(sh, String(v[C.SESSION] || '').trim(), Number(v[C.PCOUNT]) || 1,
                           String(v[C.EMP] || '').trim());
  var isMonthly = String(v[C.TYPE] || '').indexOf('รายเดือน') >= 0;
  var tick = apTicked_(v[AC_.TICK - 1]);
  var r = rcCalc_(v[C.IN], v[C.OUT], v[C.DATE], isMonthly, rcLoadCal_(), tick, v[C.DAYTYPE]);
  var H = function (m) { return rcDec_(m / n); };

  /* ใช้กติกาเดียวกับ runApply — บอกทุกเรื่องของแถวนั้น ไม่ใช่เรื่องเดียวแล้วกลบที่เหลือ */
  var bad = [], info = [];
  if (r.err)                  bad.push('⛔ ' + r.err);
  if (H(r.ot15 + r.ot3) > 12) bad.push('⚠ OT สูงผิดปกติ');
  if (r.crossLunch) {
    info.push(tick ? '☑ อนุมัติ OT ผ่าเที่ยงแล้ว · บวกคืนให้ 1 ชม.'
                   : 'ตัดพักเที่ยงให้แล้ว 1 ชม. · ถ้าทำงานจริงให้ติ๊กช่อง AH');
  }
  if (r.minApplied) info.push('งานสั้น · คิดขั้นต่ำ 0.5 ชม.');
  if (n > 1)        info.push('✓ หารให้แล้ว · งานนี้ลง ' + n + ' จ๊อบ');
  var stat = bad.concat(info).join('  |  ');

  /* ★ เขียนผลผ่าน rcWriteRow_ ใน Recheck.gs — ทางเดียวกับที่ JOBTRACK ใช้
     มีทางเขียนผลลัพธ์ทางเดียวทั้งระบบ จึงไม่มีวันเพี้ยนกันได้อีก */
  rcWriteRow_(sh, row, r, n, v[C.DAYTYPE], stat);
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
  /* ---------- คอลัมน์ตรวจเวลา ----------
     กฎเดิมเอา "รวมทั้งวัน (L)" ไปลบ 8 ตรง ๆ  →  ใครทำ OT ก็ขึ้น "เกิน" หมด
     ทั้งที่ OT ในกรอบบริษัทคือเรื่องปกติ ไม่ใช่ความผิดปกติที่ต้องเตือน
     ของใหม่แยกเป็นสองก้อน แล้วตรวจคนละเกณฑ์
       ในกรอบงาน 08:00–17:00  =  F (ชม.ทำงาน) วันปกติ · G+I (OT×1/OT×2) วันหยุด
       OT นอกเวลา            =  H (OT×1.5) + J (OT×3)
     ขาด        → วันทำงานปกติ ได้ในกรอบไม่ถึง 8 ชม.
     OT เลยกรอบ → OT เกิน 6.5 ชม. (17:00–19:30 = 2.5 + 20:00–00:00 = 4.0)
     ที่เหลือ    → เว้นว่าง = ปกติ ไม่ต้องให้ HR มานั่งอ่าน */
  sh.getRange('M1').setValue('ตรวจเวลา');
  var OTX  = '(H2:H+J2:J)';                        /* OT นอกเวลา */
  var OVER = 'IF(' + OTX + '>' + RC.OT_MAX + ',' + QQ + '⚠ OT เลยกรอบ ' + QQ +
             '&TEXT(' + OTX + '-' + RC.OT_MAX + ',' + QQ + '0.00' + QQ + ')&' +
             QQ + ' ชม.' + QQ + ',' + QQ + QQ + ')';
  sh.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(L2:L="",' + QQ + QQ + ',' +
      'IF((G2:G+I2:I)>0,' +                        /* มี OT×1/OT×2 = วันหยุด ไม่ต้องเทียบ 8 ชม. */
        OVER + ',' +
        'IF(F2:F<7.99,' + QQ + '⚠ ขาด ' + QQ + '&TEXT(8-F2:F,' + QQ + '0.00' + QQ + ')&' +
          QQ + ' ชม.' + QQ + ',' + OVER + '))))');
  apDressSummary_(sh, 13);
  sh.getRange('M1').setNote(
    'ในกรอบงาน 08:00–17:00 ไม่ถึง 8 ชม. → "ขาด"\n' +
    'OT เกินกรอบ ' + RC.OT_MAX + ' ชม. → "OT เลยกรอบ"\n' +
    'กรอบ OT ของบริษัท : 17:00–19:30 (2.5) · พัก 19:30–20:00 · 20:00–00:00 (4.0)\n' +
    'OT ที่อยู่ในกรอบ = ปกติ ช่องนี้จะเว้นว่าง');

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

/* นับแถวในรอบงานเดียวกัน — ต้องนับเฉพาะของพนักงานคนนี้เท่านั้น
   รหัสรอบงานเดี่ยว ๆ เชื่อไม่ได้ ของเก่าที่ JOBTRACK เขียนชนกันไว้ยังอยู่ในชีต
   ถ้านับรวมคนอื่นไปด้วย พอ HR แก้เวลามือทีเดียว ชั่วโมงจะโดนหารครึ่งอีกรอบ */
function apRowsInSession_(sh, sid, fallback, empId) {
  if (!sid) return Math.max(1, fallback || 1);
  var C = RC.C, last = sh.getLastRow();
  var col = sh.getRange(2, C.SESSION + 1, last - 1, 1).getValues();
  var st  = sh.getRange(2, C.STATUS + 1,  last - 1, 1).getValues();
  var em  = sh.getRange(2, C.EMP + 1,     last - 1, 1).getValues();
  var want = String(empId || '').trim();
  var cnt = 0;
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() !== sid) continue;
    if (String(st[i][0]).trim() !== 'Check Out') continue;
    if (want && String(em[i][0]).trim() !== want) continue;
    cnt++;
  }
  return Math.max(1, cnt || fallback || 1);
}

/* ============================================================================
 *  ⑤½ fixSessionClash() — ซ่อมรหัสรอบงานที่ JOBTRACK เขียนชนกันไว้ (ของเก่า)
 *
 *  รหัสรอบงานมีรูปแบบ S-<รหัสพนักงาน>-<วันเวลา> ถ้ารหัสที่ฝังอยู่ข้างในไม่ตรง
 *  กับรหัสพนักงานของแถวนั้น แปลว่าแถวนี้โดนของคนอื่นเขียนทับ ไม่มีทางถูกต้องได้
 *  ตัวนี้เปลี่ยนเฉพาะรหัสพนักงานที่ฝังอยู่ ให้ตรงกับเจ้าของแถว ส่วนวันเวลาคงเดิม
 *
 *  ปลอดภัย : ไม่แตะเวลา ไม่แตะชั่วโมง ไม่แตะแถวที่รหัสตรงอยู่แล้ว
 *  ไม่แก้ให้ : เวลาเข้าที่ถูกทับไปแล้ว — อันนั้นต้องดูรูปสแกนแล้วแก้มือเท่านั้น
 *  สคริปต์เขียนผ่านการป้องกันชีตได้ ไม่ต้องปลด NOVA lock
 * ========================================================================== */
function fixSessionClash(dryRun) {
  var sh = apSheet_(), C = RC.C, last = sh.getLastRow();
  if (last < 2) return 'ไม่มีข้อมูล';
  var rng = sh.getRange(2, C.SESSION + 1, last - 1, 1);
  var sid = rng.getValues();
  var emp = sh.getRange(2, C.EMP + 1,  last - 1, 1).getValues();
  var nam = sh.getRange(2, C.NAME + 1, last - 1, 1).getValues();
  var fixed = [], changed = false;

  for (var i = 0; i < sid.length; i++) {
    var s = String(sid[i][0] || '').trim();
    var e = String(emp[i][0] || '').trim();
    var m = s.match(/^S-(\d+)-(.+)$/);
    if (!s || !e || !m || m[1] === e) continue;      /* ปกติ ไม่ต้องยุ่ง */
    var ns = 'S-' + e + '-' + m[2];
    fixed.push({ row: i + 2, name: String(nam[i][0] || '').substring(0, 18),
                 from: s, to: ns, wrongEmp: m[1] });
    sid[i][0] = ns;
    changed = true;
  }

  if (changed && !dryRun) {
    rng.setNumberFormat('@STRING@').setValues(sid);
    SpreadsheetApp.flush();
  }

  var lines = fixed.slice(0, 30).map(function (f) {
    return '  แถว ' + f.row + ' | ' + f.name +
           '\n      เดิม ' + f.from + '  (เป็นของ ' + f.wrongEmp + ')' +
           '\n      ใหม่ ' + f.to;
  });
  Logger.log('\n===== ซ่อมรหัสรอบงานที่ชนกัน =====' +
    '\nแท็บ    : ' + sh.getName() +
    '\nพบและ' + (dryRun ? 'จะแก้' : 'แก้แล้ว') + ' : ' + fixed.length + ' แถว\n' +
    (lines.join('\n') || '  ไม่มีแถวไหนต้องแก้ — สะอาดแล้ว') +
    (fixed.length > 30 ? '\n  ... และอีก ' + (fixed.length - 30) + ' แถว' : '') +
    (fixed.length ? '\n\nเหลืออีกอย่าง : เวลาเข้าของแถวพวกนี้อาจถูกทับไปแล้ว' +
                    '\nให้เปิดรูปสแกนคอลัมน์ S เทียบ ถ้าไม่ตรงให้แก้ที่ช่อง C ได้เลย' : '') +
    '\n==============================\n');
  return fixed.length;
}

/* ดูก่อนว่าจะแก้อะไรบ้าง โดยยังไม่เขียนลงชีต */
function previewSessionClash() { return fixSessionClash(true); }


/* ============================================================================
 *  ⑤⅘ listFlagged() — ไล่รายการแถวที่ติดธง พร้อมเหตุผล (อ่านอย่างเดียว)
 *
 *  runApply บอกแค่จำนวน เช่น "แถวขึ้นธง 8 (⛔ คิดไม่ได้ 3)" แต่ไม่บอกว่าแถวไหน
 *  HR ต้องไปกรองคอลัมน์ AI เอง ตัวนี้พ่นออกมาให้เลย จะได้ไล่แก้ทีละแถว
 *
 *  ไม่นับแถวคร่อมเที่ยงที่ตัดพักเที่ยงให้แล้ว เพราะไม่ใช่ปัญหา
 * ========================================================================== */
function listFlagged() {
  var sh = apSheet_(), C = RC.C, last = sh.getLastRow();
  if (last < 2) return 'ไม่มีข้อมูล';
  var v = sh.getRange(2, 1, last - 1, AC_.NOTE).getValues();
  var stop = [], warn = [];

  for (var i = 0; i < v.length; i++) {
    var st = String(v[i][AC_.STAT - 1] || '').trim();
    if (!st) continue;
    var isStop = st.indexOf('⛔') >= 0, isWarn = st.indexOf('⚠') >= 0;
    if (!isStop && !isWarn) continue;
    /* เก็บเฉพาะข้อความที่เป็นปัญหาจริง ตัดข้อความบอกเฉย ๆ ที่ต่อท้ายออก */
    var why = st.split('  |  ').filter(function (x) {
      return x.indexOf('⛔') >= 0 || x.indexOf('⚠') >= 0;
    }).join(' + ');
    var item = '  แถว ' + (i + 2) + ' | ' + rcKey_(v[i][C.DATE]) +
               ' | ' + String(v[i][C.EMP]) + ' ' + String(v[i][C.NAME] || '').substring(0, 18) +
               ' | ' + rcHHMM_(rcMin_(v[i][C.IN])) + '–' + rcHHMM_(rcMin_(v[i][C.OUT])) +
               ' | ' + String(v[i][C.JOB] || '') +
               '\n      ' + why;
    (isStop ? stop : warn).push(item);
  }

  Logger.log('\n===== แถวที่ต้องแก้ ' + AP_VER + ' =====' +
    '\nแท็บ : ' + sh.getName() +
    '\n\n⛔ คิดชั่วโมงไม่ได้ — ต้องแก้ก่อนปิดงวด : ' + stop.length + '\n' +
    (stop.join('\n\n') || '  ไม่มี') +
    '\n\n⚠ ต้องตรวจ : ' + warn.length + '\n' +
    (warn.join('\n\n') || '  ไม่มี') +
    '\n\nวิธีแก้ : พิมพ์เวลาจริงลงช่อง C หรือ D (พื้นเขียว) แล้วคำนวณใหม่ให้เองทันที' +
    '\nดูเวลาสแกนจริงได้จากคอลัมน์ A (Timestamp) และรูปสแกนคอลัมน์ S' +
    '\n==============================\n');
  return stop.length + warn.length;
}


/* ============================================================================
 *  ⑤¾ fixRowProcess() — พิมพ์ Process ที่ถูกต้องกลับเข้าแถวที่โดนเขียนทับ
 *
 *  ตอนสแกนพร้อมกัน โค้ดเก่าเขียนทับ 3 ช่องพร้อมกันเสมอ :
 *      C  เวลาเข้า · O ชื่องานย่อย · L รหัสรอบงาน
 *  แล้วตอน Check Out ค่าใน O ยังไปสร้าง M (หมวด) กับ N (รหัสงาน) ต่ออีก
 *  รวมเป็น 4 ช่องที่เพี้ยน : M · N · O และ C
 *
 *  Process ของจริงไม่มีเก็บไว้ที่ไหนในระบบ ต้องถามพนักงานหรือหัวหน้าแล้วพิมพ์กลับเข้าไป
 *  ตัวนี้พิมพ์ให้ครบทั้ง 3 ช่องในครั้งเดียว จะได้ไม่ต้องไล่กรอกเอง และไม่ต้องปลด NOVA lock
 *
 *  วิธีใช้ :  fixRowProcess(9, 'E.2 พ่นสีภายนอก')
 *            fixRowProcess(9, 'E.2 พ่นสีภายนอก', '08:15')   ← แก้เวลาเข้าไปด้วย
 * ========================================================================== */
var AP_PROC_MAIN = {
  A:'FABRICATION & ASSEMBLY', B:'WELDING', C:'PART', D:'KEMREX', E:'PAINTING',
  F:'SYSTEM', G:'PIPING & VALVES', H:'RE-INSTALLATION', I:'SUSPENSION',
  J:'STICKER', K:'QC', L:'CLEANING'
};
function apParseProc_(label) {
  label = String(label || '').trim();
  var sp = label.indexOf(' ');
  var code = sp > 0 ? label.substring(0, sp) : label;
  var sub  = sp > 0 ? label.substring(sp + 1).trim() : '';
  var letter = code.split('.')[0].toUpperCase();
  return { code: code, sub: sub,
           main: AP_PROC_MAIN[letter] ? (letter + ' · ' + AP_PROC_MAIN[letter]) : letter };
}

function fixRowProcess(rowNo, label, newTimeIn) {
  rowNo = Number(rowNo);
  if (!rowNo || rowNo < 2) throw new Error('ใส่เลขแถวด้วย เช่น fixRowProcess(9, "E.2 พ่นสีภายนอก")');
  if (!String(label || '').trim()) throw new Error('ใส่ Process ด้วย เช่น "E.2 พ่นสีภายนอก"');

  var sh = apSheet_(), C = RC.C;
  var PMAIN = 13, PCODE = 14, PSUB = 15;          /* M · N · O (1-based) */
  var before = sh.getRange(rowNo, 1, 1, AC_.NOTE).getValues()[0];
  var who = String(before[C.EMP]) + ' ' + String(before[C.NAME] || '');
  var job = String(before[C.JOB] || '');
  var p = apParseProc_(label);

  sh.getRange(rowNo, PMAIN).setNumberFormat('@STRING@').setValue(p.main);
  sh.getRange(rowNo, PCODE).setNumberFormat('@STRING@').setValue(p.code);
  sh.getRange(rowNo, PSUB ).setNumberFormat('@STRING@').setValue(p.sub);

  var timeMsg = '(ไม่แก้)';
  if (newTimeIn) {
    sh.getRange(rowNo, C.IN + 1).setNumberFormat('@STRING@').setValue(String(newTimeIn).trim());
    timeMsg = String(before[C.IN]) + '  →  ' + newTimeIn;
  }

  var note = String(before[AC_.NOTE - 1] || '').trim();
  var mark = 'แก้ Process กลับตามที่ยืนยันแล้ว (แถวนี้เคยโดนเขียนทับตอนสแกนพร้อมกัน)';
  if (note.indexOf(mark) < 0) note = note ? (note + ' · ' + mark) : mark;
  sh.getRange(rowNo, AC_.NOTE).setValue(note);
  SpreadsheetApp.flush();

  var msg = '\n===== แก้ Process แถว ' + rowNo + ' =====' +
    '\nพนักงาน : ' + who +
    '\nจ๊อบ    : ' + job +
    '\n  M หมวด Process : ' + before[PMAIN - 1] + '  →  ' + p.main +
    '\n  N รหัสงาน      : ' + before[PCODE - 1] + '  →  ' + p.code +
    '\n  O ชื่องานย่อย   : ' + before[PSUB  - 1] + '  →  ' + p.sub +
    '\n  C เวลาเข้า      : ' + timeMsg +
    '\n\nชั่วโมงไม่เปลี่ยน — Process มีผลกับต้นทุนต่อจ๊อบเท่านั้น' +
    '\n===============================\n';
  Logger.log(msg);
  return msg;
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
  var JOBNAME = 6, PMAIN = 12, PCODE = 13, PSUB = 14;   /* G · M · N · O */
  var g = {}, order = [], sidEmp = {}, clash = [];
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][C.STATUS]).trim() !== 'Check Out') continue;
    var sid = String(v[i][C.SESSION] || '').trim();
    if (!sid) continue;
    var empId = String(v[i][C.EMP] || '').trim();
    /* กุญแจต้องมีรหัสพนักงาน — รหัสรอบงานเดี่ยว ๆ เชื่อไม่ได้ JOBTRACK เขียนชนกันได้ */
    var key = sid + '|' + empId;
    if (!sidEmp[sid]) sidEmp[sid] = {};
    sidEmp[sid][empId] = (sidEmp[sid][empId] || 0) + 1;
    if (!g[key]) {
      g[key] = { sid:sid, p:Number(v[i][C.PCOUNT]) || 1,
                 emp:empId, name:String(v[i][C.NAME] || ''),
                 d:rcKey_(v[i][C.DATE]), first:i + 1, det:[] };
      order.push(key);
    }
    /* ลายนิ้วมือของแถว — ใช้เทียบว่าสองแถวเป็นงานเดียวกันเป๊ะหรือเปล่า */
    var sig = rcHHMM_(rcMin_(v[i][C.IN])) + '~' + rcHHMM_(rcMin_(v[i][C.OUT])) + '~' +
              String(v[i][C.JOB] || '') + '~' + String(v[i][PMAIN] || '') + '/' +
              String(v[i][PCODE] || '') + '/' + String(v[i][PSUB] || '');
    g[key].det.push({
      row : i + 1,
      inT : rcHHMM_(rcMin_(v[i][C.IN])),
      outT: rcHHMM_(rcMin_(v[i][C.OUT])),
      job : String(v[i][C.JOB] || '') + ' ' + String(v[i][JOBNAME] || '').substring(0, 18),
      proc: [v[i][PMAIN], v[i][PCODE], v[i][PSUB]].filter(String).join(' / ').substring(0, 30),
      sig : sig
    });
  }
  var bad = [], okCnt = 0;
  for (var oi = 0; oi < order.length; oi++) {
    var x = g[order[oi]];
    if (x.det.length === x.p) { okCnt++; continue; }
    bad.push(x);
  }
  bad.sort(function (a, b) { return a.first - b.first; });

  /* ★ รหัสรอบงานเดียวแต่มีหลายคน = JOBTRACK เขียนทับแถวกันตอนสแกนพร้อมกัน
     อันตรายกว่าแถวซ้ำ เพราะเวลาเข้าของอีกคนถูกทับไปแล้ว ตัวเลขเชื่อไม่ได้ */
  for (var s2 in sidEmp) {
    var owners = [];
    for (var e2 in sidEmp[s2]) owners.push(e2);
    if (owners.length > 1) clash.push({ sid:s2, owners:owners });
  }
  var clashLines = clash.slice(0, 15).map(function (c) {
    var who = c.owners.map(function (o) {
      var k2 = c.sid + '|' + o;
      return o + ' ' + (g[k2] ? g[k2].name.substring(0, 18) + ' (แถว ' + g[k2].first + ')' : '');
    }).join('  ·  ');
    return '  ● ' + c.sid + '\n      ' + who;
  });

  var lines = bad.slice(0, 25).map(function (x) {
    /* แถวเหมือนกันเป๊ะทุกแถว = สแกนซ้ำ · ต่างกัน = คนละ Process แต่ P ไม่อัปเดต */
    var same = true;
    for (var q = 1; q < x.det.length; q++) if (x.det[q].sig !== x.det[0].sig) { same = false; break; }
    var verdict = x.det.length < x.p
      ? '  ⛔ แถวหาย — เพิ่มไม่สำเร็จหรือมีคนลบ ชั่วโมงเคยหาย'
      : (same
          ? '  ⛔ ทุกแถวเหมือนกันเป๊ะ = สแกนออกซ้ำ · ให้ลบทิ้งเหลือแถวเดียว'
          : '  ✓ คนละงาน/คนละ Process จริง — แค่คอลัมน์ P ไม่อัปเดต ไม่ต้องแก้');
    var det = x.det.map(function (d) {
      return '      แถว ' + d.row + ' | ' + d.inT + '–' + d.outT +
             ' | จ๊อบ ' + d.job + (d.proc ? ' | ' + d.proc : '');
    }).join('\n');
    return '  ● ' + x.d + ' | ' + x.emp + ' ' + x.name.substring(0, 20) +
           ' | คอลัมน์ P บอก ' + x.p + ' แต่มีจริง ' + x.det.length + ' แถว\n' +
           det + '\n' + verdict;
  });
  Logger.log('\n===== ตรวจแถวหาย ' + RC.VER + ' =====' +
    '\nแท็บ          : ' + sh.getName() +
    '\nรอบงานทั้งหมด : ' + (okCnt + bad.length) +
    '\n  ครบ         : ' + okCnt +
    '\n  ไม่ครบ      : ' + bad.length + '\n' +
    (lines.join('\n\n') || '  ไม่มีปัญหา') +
    (bad.length > 25 ? '\n  ... และอีก ' + (bad.length - 25) + ' รอบ' : '') +
    '\n\n----- รหัสรอบงานชนกัน (คนละคนแต่รหัสเดียวกัน) : ' + clash.length + ' -----\n' +
    (clashLines.join('\n') || '  ไม่มี') +
    (clash.length > 15 ? '\n  ... และอีก ' + (clash.length - 15) + ' รหัส' : '') +
    (clash.length ? '\n  สาเหตุ : doCheckIn ของ JOBTRACK ไม่มี LockService — สแกนพร้อมกันแล้วแย่งแถวกัน' +
                    '\n  ผลกระทบ : เวลาเข้า/Process ของแถวหลังถูกทับ ต้องตรวจกับรูปสแกนแล้วแก้มือ' : '') +
    '\n\nเรื่องเงิน : สรุปรายคนถูกต้องแล้วทุกกรณี (หารด้วยรหัสรอบงาน + รหัสพนักงาน ไม่ใช่เลขคอลัมน์ P)' +
    '\nที่ต้องแก้ : เฉพาะรอบที่ขึ้น ⛔ เพราะต้นทุนต่อจ๊อบจะเพี้ยน' +
    '\n=======================\n');
  return bad.length + clash.length;
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
  /* ★ ไม่กวาดทั้งชีตระหว่างวันอีกแล้ว — JOBTRACK คำนวณให้ตั้งแต่ตอนสร้างแถว
     ตัวนี้เหลือเป็นแค่ "ตาข่ายกันพลาด" รันตี 3 วันละครั้ง
     เก็บแถวตกหล่น + ตั้งฟอร์มให้แท็บงวดใหม่ตอนขึ้นวันที่ 26 */
  ScriptApp.newTrigger('autoApply').timeBased().everyDays(1).atHour(3).create();

  /* ติดตั้งตัวคำนวณสดตอนแก้มือไปพร้อมกันเลย จะได้ไม่ลืม */
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditRecalc') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditRecalc')
    .forSpreadsheet(SpreadsheetApp.openById(RC.LOG_ID)).onEdit().create();

  var msg = 'ติดตั้งเรียบร้อย — ต่อจากนี้ไม่ต้องรันมืออีกเลย\n' +
    '  · ตอนสแกนออก : JOBTRACK คำนวณชั่วโมงลงแถวนั้นให้ทันที\n' +
    '  · ตี 3 ทุกวัน  : ตรวจทานเก็บแถวตกหล่น (ไม่กวนคนใช้งานระหว่างวัน)\n' +
    '  · งวดใหม่    : ฟอร์มตามไปเองในรอบตี 3\n' +
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
