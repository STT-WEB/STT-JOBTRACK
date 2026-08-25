/**
 * ============================================================================
 *  Recheck.gs — รายงานทดลอง "กฎคิดชั่วโมง v5" (โหมดอ่านอย่างเดียว)
 *
 *  ⛔ ไฟล์นี้ไม่แก้ข้อมูลเดิมแม้แต่ช่องเดียว
 *     - อ่าน  : JOBTRACK_Job_Log 2026 → แท็บ Job_Log_2569_08
 *     - อ่าน  : STT Jobcost Database  → แท็บ ประเภทวันทำงาน
 *     - เขียน : "ไฟล์ใหม่" ที่สร้างขึ้นเองใน Drive เท่านั้น
 *
 *  ⛔ ไฟล์นี้อยู่ในโปรเจกต์ nova-hr-hub — ไม่ได้แตะโปรเจกต์ JOBTRACK เลย
 *
 *  วิธีใช้ : เปิด Apps Script Editor → เลือกฟังก์ชัน runRecheck → กด Run
 *           เสร็จแล้วดูลิงก์ไฟล์รายงานใน Execution log
 *
 *  ชื่อฟังก์ชันขึ้นต้น rc ทุกตัว กันชนกับ Code.gs / Cal.gs ที่ใช้ scope เดียวกัน
 * ============================================================================
 */

var RC = {
  LOG_ID   : '1ZPl3uVRtM5r4sPA-yX1OwTyp34XCTIcKRC0Sx8qsr9s',  // JOBTRACK_Job_Log 2026
  DB_ID    : '1MYWORYN3sOjov3Gxv3UqCV1jRSxgxwGi1tRomFUGSr0',  // ฐานข้อมูลพนักงาน
  TAB      : 'Job_Log_2569_08',                                // งวด 26 ก.ค. – 25 ส.ค. 69
  CAL_TAB  : 'ประเภทวันทำงาน',
  /* ตำแหน่งคอลัมน์ (0-based) — ตรงกับ JOBTRACK */
  C: { DATE:1, IN:2, OUT:3, HOURS:4, JOB:5, EMP:7, NAME:8, DEPT:9, TYPE:10,
       PCOUNT:15, STATUS:16, DAYTYPE:19, HNORM:21, HOT:22, HPAY:23 }
};

/* ---------------------------------------------------------------- ตัวช่วยเวลา */
function rcMin_(t) {
  if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
  var s = String(t || '').trim();
  if (!s || s.length < 3) return -1;
  if (s.indexOf('T') > -1) s = s.split('T')[1];
  var p = s.split(':');
  var h = parseInt(p[0], 10), m = parseInt(p[1], 10);
  if (isNaN(h) || isNaN(m)) return -1;
  return h * 60 + m;
}
function rcHHMM_(m) { m = ((m % 1440) + 1440) % 1440;
  return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }
function rcDec_(m) { return m <= 0 ? 0 : Math.round(m / 60 * 100) / 100; }

/* ------------------------------------------------- กฎปัดเวลา v5 (ตามข้อ 3) */
function rcSnapIn_(m) {
  if (m <= 420) return m;                 // ≤07:00  OT เช้า ใช้จริง
  if (m <= 495) return 480;               // 07:01–08:15 → 08:00
  if (m <  715) return m;                 // 08:16–11:54 ใช้จริง
  if (m <= 795) return 780;               // 11:55–13:15 → 13:00   ★ กฎใหม่
  return m;                               // ≥13:16 ใช้จริง
}
function rcSnapOut_(m) {
  if (m <  715)  return m;                // ≤11:54 ใช้จริง
  if (m <= 795)  return 720;              // 11:55–13:15 → 12:00   ★ กฎใหม่
  if (m <  1015) return m;                // 13:16–16:54 ใช้จริง
  if (m <= 1049) return 1020;             // 16:55–17:29 → 17:00
  return m;                               // ≥17:30 ใช้จริง
}

/* ------------------------------------------------------ ปฏิทินประเภทวัน */
function rcLoadCal_() {
  var map = {};
  var sh = SpreadsheetApp.openById(RC.DB_ID).getSheetByName(RC.CAL_TAB);
  if (!sh) return map;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    var d = new Date(v[i][0]);
    if (isNaN(d.getTime())) continue;
    map[d.toLocaleDateString('th-TH')] = String(v[i][2] || '').trim();
  }
  return map;
}
function rcKind_(s) {
  s = String(s || '');
  if (s.indexOf('นักขัตฤกษ์') >= 0) return 'นักขัตฤกษ์';
  if (s.indexOf('หยุด') >= 0)       return 'หยุด';
  return 'ปกติ';
}
/* บวกวันจากสตริง d/m/พ.ศ. */
function rcAddDay_(dateStr, n) {
  var p = String(dateStr).split('/');
  if (p.length < 3) return dateStr;
  var d = new Date(parseInt(p[2], 10) - 543, parseInt(p[1], 10) - 1, parseInt(p[0], 10));
  d.setDate(d.getDate() + n);
  return d.toLocaleDateString('th-TH');
}

/* =========================================================== เครื่องคิดชั่วโมง v5
 *  คืนค่านาที แยกเป็น 5 ชั้นค่าแรง + นาทีพักเที่ยงที่ถูกตัด
 *    hNormal ชม.ทำงานปกติ            (วันทำงาน ในกรอบ 08–17)
 *    ot1     OT ×1  วันหยุด รายเดือน  (ในกรอบ 08–17)
 *    ot15    OT ×1.5 วันทำงาน         (ก่อน 08 · หลัง 17 · ผ่าเที่ยงที่อนุมัติ)
 *    ot2     OT ×2  วันหยุด รายวัน    (ในกรอบ 08–17)
 *    ot3     OT ×3  OT ในวันหยุด      (ก่อน 08 · หลัง 17)
 */
function rcCalc_(inRaw, outRaw, dateStr, isMonthly, calMap, lunchApproved) {
  var r = { hNormal:0, ot1:0, ot15:0, ot2:0, ot3:0, lunch:0,
            inUse:'', outUse:'', crossLunch:false, err:'' };

  var i0 = rcMin_(inRaw), o0 = rcMin_(outRaw);
  if (i0 < 0) { r.err = 'ไม่มีเวลาเข้า'; return r; }
  if (o0 < 0) { r.err = 'ไม่มีเวลาออก'; return r; }

  /* คร่อมเที่ยงจริงไหม — ดูจากเวลาดิบก่อนปัด */
  var o0abs0 = o0 < i0 ? o0 + 1440 : o0;
  if (i0 <= 720 && o0abs0 >= 780) r.crossLunch = true;   // ต้องอยู่คร่อมเที่ยงเต็มชั่วโมงจริง

  var iM = rcSnapIn_(i0), oM = rcSnapOut_(o0);
  r.inUse = rcHHMM_(iM); r.outUse = rcHHMM_(oM);
  /* ข้ามเที่ยงคืนจริงหรือไม่ ต้องดูจาก "เวลาดิบ" ไม่ใช่เวลาที่ปัดแล้ว
     ไม่งั้นเคสออก 12:01 กลับเข้า 12:15 (ปัดเป็น 12:00 กับ 13:00) จะถูกนับเป็น 23 ชม. */
  if (o0 < i0) oM += 1440;
  if (oM <= iM) { r.err = 'ปัดแล้วออกก่อนเข้า'; return r; }

  var cursor = iM, guard = 0;
  while (cursor < oM && guard < 8) {
    var off  = Math.floor(cursor / 1440);
    var base = off * 1440;
    var segEnd = Math.min(oM, base + 1440);
    var dStr = off === 0 ? dateStr : rcAddDay_(dateStr, off);
    var kind = rcKind_(calMap[dStr] || '');
    var holiday = (kind !== 'ปกติ');

    var lo = cursor - base, hi = segEnd - base;
    var bands = [
      { f:0,    t:480,  z:'OT'     },   // ก่อน 08:00
      { f:480,  t:720,  z:'NORMAL' },   // 08:00–12:00
      { f:720,  t:780,  z:'LUNCH'  },   // 12:00–13:00
      { f:780,  t:1020, z:'NORMAL' },   // 13:00–17:00
      { f:1020, t:1440, z:'OT'     }    // 17:00–24:00
    ];
    for (var b = 0; b < bands.length; b++) {
      var a = Math.max(lo, bands[b].f), z = Math.min(hi, bands[b].t);
      if (z <= a) continue;
      var mins = z - a;
      if (bands[b].z === 'LUNCH') { r.lunch += mins; continue; }   // ★ พักเที่ยงไม่นับ
      if (bands[b].z === 'NORMAL') {
        if (!holiday)        r.hNormal += mins;
        else if (isMonthly)  r.ot1     += mins;
        else                 r.ot2     += mins;
      } else {
        if (!holiday) r.ot15 += mins; else r.ot3 += mins;
      }
    }
    cursor = segEnd; guard++;
  }

  /* OT ผ่าเที่ยง — ให้เต็ม 1.00 ชม. เมื่ออนุมัติ และต้องคร่อมเที่ยงจริง */
  if (lunchApproved && r.crossLunch) {
    if (rcKind_(calMap[dateStr] || '') === 'ปกติ') r.ot15 += 60; else r.ot3 += 60;
  }
  return r;
}

/* ================================================================== รันรายงาน */
function runRecheck() {
  var t0 = new Date().getTime();
  var sh = SpreadsheetApp.openById(RC.LOG_ID).getSheetByName(RC.TAB);
  if (!sh) throw new Error('ไม่พบแท็บ ' + RC.TAB + ' — เช็กชื่อแท็บอีกที');
  var v = sh.getDataRange().getValues();
  var calMap = rcLoadCal_();
  var C = RC.C;

  var out = [[
    'แถว','วันที่','รหัส','ชื่อ','ประเภท','Job ID','Proc',
    'เข้า(สแกน)','ออก(สแกน)','เข้า(ใช้)','ออก(ใช้)',
    'เดิม ชม.ปกติ','เดิม ชม.OT','เดิม ชม.รวม',
    'ใหม่ ปกติ','ใหม่ OT×1','ใหม่ OT×1.5','ใหม่ OT×2','ใหม่ OT×3',
    'ใหม่ รวมปกติ','ใหม่ รวมOT','ใหม่ ชม.รวม',
    'ต่าง ชม.รวม','พักเที่ยงที่ตัด','คร่อมเที่ยง','สถานะ'
  ]];

  var S = { rows:0, done:0, changed:0, err:0, cross:0,
            oN:0, oO:0, oT:0, nN:0, nO:0, nT:0, lunchCut:0, crossHrs:0 };
  var byEmp = {};

  for (var i = 1; i < v.length; i++) {
    var row = v[i];
    if (!row[C.EMP] && !row[C.IN]) continue;
    S.rows++;
    if (String(row[C.STATUS]).trim() !== 'Check Out') continue;
    S.done++;

    var n = Number(row[C.PCOUNT]) || 1; if (n < 1) n = 1;
    var typeStr  = String(row[C.TYPE] || '');
    var isMonthly = typeStr.indexOf('รายเดือน') >= 0;
    var dateStr  = String(row[C.DATE] || '').trim();

    var r;
    try { r = rcCalc_(row[C.IN], row[C.OUT], dateStr, isMonthly, calMap, false); }
    catch (e) { r = { err:'คำนวณพลาด: ' + e.message, hNormal:0,ot1:0,ot15:0,ot2:0,ot3:0,lunch:0,crossLunch:false,inUse:'',outUse:'' }; }

    var oldN = Number(row[C.HNORM]) || 0, oldO = Number(row[C.HOT]) || 0;
    var oldT = oldN + oldO;

    var nN = rcDec_((r.hNormal + r.ot1 + r.ot2) / n);
    var nO = rcDec_((r.ot15 + r.ot3) / n);
    var nT = Math.round((nN + nO) * 100) / 100;
    var diff = Math.round((nT - oldT) * 100) / 100;

    if (r.err) S.err++;
    if (r.crossLunch) { S.cross++; S.crossHrs += 1 / n; }
    if (Math.abs(diff) >= 0.01) S.changed++;

    S.oN += oldN; S.oO += oldO; S.oT += oldT;
    S.nN += nN;   S.nO += nO;   S.nT += nT;
    S.lunchCut += r.lunch / 60 / n;

    var id = String(row[C.EMP]).trim();
    if (!byEmp[id]) byEmp[id] = { name:String(row[C.NAME]||''), type:typeStr, old:0, nw:0, rows:0, cross:0 };
    byEmp[id].old += oldT; byEmp[id].nw += nT; byEmp[id].rows++;
    if (r.crossLunch) byEmp[id].cross++;

    out.push([
      i + 1, dateStr, id, String(row[C.NAME]||''), isMonthly ? 'รายเดือน' : 'รายวัน',
      String(row[C.JOB]||''), n,
      String(row[C.IN]||''), String(row[C.OUT]||''), r.inUse, r.outUse,
      oldN, oldO, Math.round(oldT*100)/100,
      rcDec_(r.hNormal/n), rcDec_(r.ot1/n), rcDec_(r.ot15/n), rcDec_(r.ot2/n), rcDec_(r.ot3/n),
      nN, nO, nT,
      diff, rcDec_(r.lunch/n), r.crossLunch ? 'คร่อมเที่ยง' : '',
      r.err ? '⛔ ' + r.err : (Math.abs(diff) >= 0.01 ? 'เปลี่ยน' : 'เท่าเดิม')
    ]);
  }

  /* ---------- สร้างไฟล์รายงานใหม่ (ไม่แตะไฟล์เดิม) ---------- */
  var ss = SpreadsheetApp.create('รายงานทดลอง JOBTRACK v5 — งวด ' + RC.TAB.replace('Job_Log_',''));
  var d1 = ss.getSheets()[0].setName('รายแถว');
  d1.getRange(1, 1, out.length, out[0].length).setValues(out);
  d1.setFrozenRows(1); d1.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
  d1.getRange(2, 23, out.length - 1, 1).setNumberFormat('0.00;[RED]-0.00');

  var emps = [];
  for (var k in byEmp) {
    var e = byEmp[k];
    emps.push([k, e.name, e.type, e.rows, e.cross,
               Math.round(e.old*100)/100, Math.round(e.nw*100)/100,
               Math.round((e.nw - e.old)*100)/100]);
  }
  emps.sort(function (a, b) { return a[7] - b[7]; });
  var d2 = ss.insertSheet('รายคน');
  d2.getRange(1, 1, 1, 8).setValues([['รหัส','ชื่อ','ประเภท','จำนวนแถว','แถวคร่อมเที่ยง','เดิม ชม.รวม','ใหม่ ชม.รวม','ต่าง']]).setFontWeight('bold');
  if (emps.length) d2.getRange(2, 1, emps.length, 8).setValues(emps);
  d2.setFrozenRows(1);

  var sum = [
    ['รายงานทดลอง — กฎคิดชั่วโมง v5', ''],
    ['แท็บต้นทาง', RC.TAB],
    ['สร้างเมื่อ', new Date()],
    ['', ''],
    ['แถวทั้งหมด', S.rows],
    ['แถว Check Out', S.done],
    ['แถวที่ตัวเลขเปลี่ยน', S.changed],
    ['แถวที่มีปัญหา (⛔)', S.err],
    ['แถวคร่อมเที่ยง (รออนุมัติ)', S.cross],
    ['', ''],
    ['เดิม · ชม.ปกติ', Math.round(S.oN*100)/100],
    ['เดิม · ชม.OT', Math.round(S.oO*100)/100],
    ['เดิม · ชม.รวม', Math.round(S.oT*100)/100],
    ['', ''],
    ['ใหม่ · ชม.ปกติ', Math.round(S.nN*100)/100],
    ['ใหม่ · ชม.OT', Math.round(S.nO*100)/100],
    ['ใหม่ · ชม.รวม', Math.round(S.nT*100)/100],
    ['', ''],
    ['ต่าง · ชม.รวม', Math.round((S.nT-S.oT)*100)/100],
    ['ต่าง · ชม.OT', Math.round((S.nO-S.oO)*100)/100],
    ['พักเที่ยงที่ถูกตัดออกรวม (ชม.)', Math.round(S.lunchCut*100)/100],
    ['ถ้าอนุมัติ OT ผ่าเที่ยงทุกแถว จะเพิ่มอีก (ชม.)', Math.round(S.crossHrs*100)/100],
    ['', ''],
    ['หมายเหตุ', 'รายงานนี้สมมติว่ายังไม่อนุมัติ OT ผ่าเที่ยงเลยสักแถว'],
    ['', 'ไม่มีการแก้ไขไฟล์ต้นทางใด ๆ ทั้งสิ้น']
  ];
  var d3 = ss.insertSheet('สรุป', 0);
  d3.getRange(1, 1, sum.length, 2).setValues(sum);
  d3.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  d3.setColumnWidth(1, 300); d3.setColumnWidth(2, 220);

  var msg =
    '\n========== รายงานทดลอง กฎคิดชั่วโมง v5 ==========' +
    '\nแท็บต้นทาง       : ' + RC.TAB +
    '\nแถวทั้งหมด       : ' + S.rows + '   (Check Out ' + S.done + ')' +
    '\nแถวที่เปลี่ยน     : ' + S.changed +
    '\nแถวมีปัญหา ⛔    : ' + S.err +
    '\nแถวคร่อมเที่ยง    : ' + S.cross +
    '\n-------------------------------------------------' +
    '\nเดิม  ปกติ ' + Math.round(S.oN*100)/100 + '  OT ' + Math.round(S.oO*100)/100 + '  รวม ' + Math.round(S.oT*100)/100 +
    '\nใหม่  ปกติ ' + Math.round(S.nN*100)/100 + '  OT ' + Math.round(S.nO*100)/100 + '  รวม ' + Math.round(S.nT*100)/100 +
    '\nต่าง  รวม ' + Math.round((S.nT-S.oT)*100)/100 + '  ชม.  ·  OT ' + Math.round((S.nO-S.oO)*100)/100 + ' ชม.' +
    '\nพักเที่ยงที่ตัดออก : ' + Math.round(S.lunchCut*100)/100 + ' ชม.' +
    '\nถ้าอนุมัติผ่าเที่ยงหมด จะเพิ่ม : ' + Math.round(S.crossHrs*100)/100 + ' ชม.' +
    '\n-------------------------------------------------' +
    '\nไฟล์รายงาน : ' + ss.getUrl() +
    '\nใช้เวลา ' + Math.round((new Date().getTime()-t0)/1000) + ' วินาที' +
    '\n=================================================\n';
  Logger.log(msg);
  return msg;
}
