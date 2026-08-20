/**
 * ============================================================================
 *  NOVA-HR · Payload.gs — ตั้งค่า + ประกอบร่าง
 *  ไฟล์นี้คือ "หน้าปก" — ตั้งค่าทะเบียนไฟล์/ปี/แคช แล้วสั่งไฟล์ย่อยทำงานทีละส่วน
 *  แก้ตัวเลขไม่ตรง? ดูว่าเป็นเรื่องอะไรแล้วเปิดไฟล์นั้นไฟล์เดียว:
 *    พนักงาน/แผนก → Payload_Emp.gs   ชั่วโมง/ค่าแรงรายเดือน → Payload_Cal.gs
 *    จ๊อบ/งบ/ปัน Indirect → Payload_Job.gs   Reconcile → Payload_Recon.gs
 *    ด่านตรวจ → Payload_Verify.gs   ตัวช่วยอ่านชีต → Payload_Util.gs
 * ============================================================================
 */

/* ---------------- ทะเบียนไฟล์ (จากแท็บ File Registry ของจริง) ---------------- */
var FILES = {
  EMPLOYEE : '1MYWORYN3sOjov3Gxv3UqCV1jRSxgxwGi1tRomFUGSr0', // Employee & Account
  JOBLOG   : '1ZPl3uVRtM5r4sPA-yX1OwTyp34XCTIcKRC0Sx8qsr9s', // JOBTRACK_Job_Log 2026
  BPLUS_DIR: '1RSI7YcjDkGLzIkqldLghJ44PSPDN-7v6',            // โฟลเดอร์ Time Bplus
  CAL_DIR  : '1Ah_oiVr1qyf7h2BHpdllIYUb8XpFz6Kg',            // โฟลเดอร์ Cal รายเดือน
  MASTER   : '1gzryio6lwbhozn19-aOUGkzScNtk1bvJGk5tUw-gcbM', // Jobcost Master/Dashboard 2026
  PAYROLLDB: '1hw5cyLEZO3dEol5X1pX3_2oBYtuJmADTTMwrOsll1D8', // Salary Master + Budget labour
  PAYACTUAL: '1U6Zt_G6J3rXgvbOrWrWw9H_gaYeZIOOekql34fTvQgo'  // ตารางเงินเดือนจริง + HR&ACC
};

var CFG = {
  YEAR      : 2569,
  NMONTH    : 8,                 // เดือนที่มีข้อมูลแล้ว
  KEMREX    : 5018,              // ✅ รหัสแผนก KEMREX (เบียร์ยืนยันแล้ว) — อยู่ใต้หน่วยงาน 5000 PRODUCTION
  KEMREX_SPLIT: true,            // true = ยก KEMREX ขึ้นเป็น "แผนกใหญ่" แยกบนหน้าจอ (แม้ไฟล์จะอยู่ใต้ PRODUCTION)
  PERIOD_START: 26,              // งวดจ่าย = วันที่ 26 เดือนก่อน ถึง 25 เดือนนี้ (ตรงกับไฟล์ Bplus)
  SHOW_ALL_JOBS: false,          // true = โชว์จ๊อบเก่าที่ไม่มียอดอะไรเลยด้วย
  NPROC     : 7,                 // โชว์ Process 7 อันดับแรกแยกสี ที่เหลือรวมเป็น "อื่นๆ" (จานสีมี 8 ช่อง)
  CACHE_KEY : 'NOVA_PAYLOAD_v3_1_0',   // ⚠️ บั๊มเลขนี้ทุกครั้งที่แก้ logic (กันเว็บโชว์ของเก่า)
  CACHE_SEC : 300
};

var MONTHS_TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

/* ============================================================================
   ประกอบร่าง + แคช
   ============================================================================ */
function buildPayload() {
  var cf = calFiles_();
  if (!cf.length) throw new Error('ไม่พบไฟล์ Cal ในโฟลเดอร์ ' + FILES.CAL_DIR);
  CFG.NMONTH = cf[cf.length - 1].m;          // เดือนล่าสุดที่มีไฟล์จริง (ต้องรู้ก่อนอ่านทะเบียนพนักงาน)

  var emps = buildEmps_();
  var empMap = {}; emps.STT.concat(emps.S1 || []).forEach(function (e) { empMap[e.id] = e; });

  var cal = readAllCalMonths_(empMap);          // { rows, jobRows, perf, procRows, empInfo }
  CFG._added = mergeEmpInfo_(emps, empMap, cal.empInfo);   // ไฟล์ Cal คือความจริง — ทับทะเบียนให้ตรง
  var alloc = buildAlloc_(cal.rows, cal.jobRows, empMap);
  var jobs = buildJobs_(alloc);
  var recon = buildRecon_(cal, empMap);        // ⚠️ ต้องเห็นไฟล์ Bplus ก่อนถึงจะครบ

  var D = {
    meta: {
      note: '', year: CFG.YEAR, nmonth: CFG.NMONTH, monthsTH: MONTHS_TH,
      workdays: cal.workdays, holidays: cal.holidays,
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd MMM yyyy HH:mm')
    },
    depts: buildDepts_(emps),
    processes: cal.processes,
    hourTypes: cal.hourTypes,
    causes: [
      { t: 'ลืม Check Out', d: 'JOBTRACK บวมเกินจริง — ชั่วโมงออกงานไม่ถูกบันทึก' },
      { t: 'ไม่ได้สแกนนิ้ว', d: 'มีใน JOBTRACK แต่ไม่มีใน Bplus' },
      { t: 'ลงจ๊อบซ้อนเวลา', d: 'คนเดียวอยู่ 2 จ๊อบช่วงเวลาเดียวกัน' },
      { t: 'ประเภทชั่วโมงผิด', d: 'ลง OT เป็นชั่วโมงปกติ ตัวคูณเลยผิด' },
      { t: 'ปฏิทินไม่ตรง', d: 'วันหยุดนักขัตฤกษ์ในระบบไม่ตรงกับที่ Payroll ใช้' }
    ],
    emps: emps,
    companies: {
      STT: { rows: cal.rows, jobRows: cal.jobRows, recon: recon, perf: cal.perf, alloc: alloc },
      S1: { rows: [], jobRows: [], recon: [], perf: [], alloc: {} }
    },
    jobs: jobs, procRows: cal.procRows, hracc: buildHrAcc_()
  };
  D.verify = verify_(D);
  return D;
}

function getPayloadJson() {
  var c = CacheService.getScriptCache();
  var hit = c.get(CFG.CACHE_KEY);
  if (hit) return hit;
  var s = JSON.stringify(buildPayload()).replace(/<\//g, '<\\/');   // กัน '</script>' ที่อาจหลุดมาในชื่องาน
  try { c.put(CFG.CACHE_KEY, s, CFG.CACHE_SEC); } catch (e) { /* ใหญ่เกิน 100KB → ข้ามแคช */ }
  return s;
}

function forceRefresh() { CacheService.getScriptCache().remove(CFG.CACHE_KEY); }

/* ============================================================================
   probe() — รันใน Apps Script Editor เพื่อ "ส่องไฟล์จริง" ก่อน deploy
   ดูผลที่ View → Logs · ถ้าแท็บไหนขึ้น ✗ แปลว่าชื่อแท็บในไฟล์เปลี่ยน ต้องมาแก้ที่ findTab_
   ============================================================================ */
function probe() {
  var L = [];
  var cf = calFiles_();
  L.push('พบไฟล์ Cal ' + cf.length + ' เดือน: ' + cf.map(function (x) { return x.m + '=' + x.name; }).join(' | '));
  var bf = bplusFiles_();
  L.push('พบไฟล์ Bplus (แปลงเป็น Google Sheets แล้ว) เดือน: ' + Object.keys(bf).join(', '));
  if (cf.length) {
    var ss = SpreadsheetApp.openById(cf[cf.length - 1].id);
    L.push('แท็บทั้งหมดในไฟล์ ' + cf[cf.length - 1].name + ':');
    ss.getSheets().forEach(function (s) { L.push('   • ' + s.getName() + '  (' + s.getLastRow() + ' แถว)'); });
    ['JOB_COST_DIRECT', 'PAYROLL_SUMMARY', 'PERFORMANCE', 'สรุปตารางเงินเดือน', 'ปฏิทินวันทำงาน', 'WORK_HOUR_TYPE', 'ALL WIP']
      .forEach(function (k) {
        var sh = findTab_(ss, [k]);
        L.push((sh ? '✓ ' : '✗ ') + k + (sh ? ' → ' + sh.getName() : ' → หาไม่เจอ'));
      });
  }
  try {
    var D = buildPayload();
    L.push('— ผลประกอบร่าง —');
    L.push('พนักงาน ' + D.emps.STT.length + ' คน · แผนกใหญ่ ' + D.depts.length + ' · Process ' + D.processes.length);
    L.push('เงินเดือน ' + D.companies.STT.rows.length + ' แถว · ลงจ๊อบ ' + D.companies.STT.jobRows.length +
           ' แถว · Performance ' + D.companies.STT.perf.length + ' แถว · Reconcile ' + D.companies.STT.recon.length + ' แถว');
    L.push('จ๊อบ ' + D.jobs.length + ' · เดือนที่มีข้อมูล ' + D.meta.nmonth);
    L.push('เติมพนักงานที่ไม่มีในทะเบียนจากไฟล์ Cal: ' + (CFG._added || 0) + ' คน');
    L.push('แผนกใหญ่: ' + D.depts.map(function (d) { return d.code + '=' + d.name; }).join(' · '));
    L.push('ด่านตรวจ ผ่าน ' + D.verify.pass + '/' + D.verify.total);
    D.verify.fails.forEach(function (f) { L.push('   ✗ ' + f); });
    L.push(diagDirect_(D));
  } catch (e) {
    L.push('!! buildPayload ล้ม: ' + e.message);
  }
  Logger.log(L.join('\n'));
  return L.join('\n');
}

/**
 * วิเคราะห์ว่าทำไม "ค่าแรง Direct ≠ ต้นทุนที่ลงจ๊อบ" — ชี้ตัวคนที่เป็นต้นเหตุ
 * (ใช้ตอนรัน probe() เท่านั้น ไม่ได้ถูกเรียกตอนเปิดเว็บ)
 */
function diagDirect_(D) {
  var co = D.companies.STT, L = ['— หาต้นเหตุ Direct ≠ ต้นทุนลงจ๊อบ —'];
  for (var m = 1; m <= CFG.NMONTH; m++) {
    var bpD = 0, jc = 0, inJob = {}, payOf = {};
    co.jobRows.forEach(function (x) { if (x.m === m) { jc += x.cost; inJob[x.id] = (inJob[x.id] || 0) + x.cost; } });
    co.rows.forEach(function (r) { if (r.m === m) { payOf[r.id] = r; if (r.dir) bpD += r.bp; } });
    var gap = q2_(bpD - jc);
    if (Math.abs(gap) <= 0.05) continue;
    L.push(MONTHS_TH[m - 1] + ': Direct ' + q2_(bpD).toLocaleString() + ' − ลงจ๊อบ ' + q2_(jc).toLocaleString() + ' = ' + gap.toLocaleString());
    var a = [], b = [], c = [];
    Object.keys(payOf).forEach(function (id) {
      var r = payOf[id];
      if (r.dir && !inJob[id]) a.push(id + ' ' + (r.bp).toLocaleString());          // Direct แต่ไม่มีในตารางลงจ๊อบ
      if (!r.dir && inJob[id]) b.push(id + ' ' + q2_(inJob[id]).toLocaleString());  // Indirect แต่ไปโผล่ในจ๊อบ
    });
    Object.keys(inJob).forEach(function (id) { if (!payOf[id]) c.push(id + ' ' + q2_(inJob[id]).toLocaleString()); });
    if (a.length) L.push('   • เป็น Direct แต่ไม่มีในตารางลงจ๊อบ (' + a.length + ' คน): ' + a.slice(0, 6).join(' | '));
    if (b.length) L.push('   • เป็น Indirect แต่ไปโผล่ในตารางลงจ๊อบ (' + b.length + ' คน): ' + b.slice(0, 6).join(' | '));
    if (c.length) L.push('   • ลงจ๊อบแต่ไม่มีในตารางเงินเดือนเลย (' + c.length + ' คน): ' + c.slice(0, 6).join(' | '));
  }
  return L.length > 1 ? L.join('\n') : '— Direct = ต้นทุนลงจ๊อบ ครบทุกเดือน ✅ —';
}
