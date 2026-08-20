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
  /* ⚠ บั๊ม BUILD ทุกครั้งที่แก้โค้ด — เลขนี้จะไปโชว์บนหน้า Login และมุมขวาบนของแอป
     ถ้าเลขบนเว็บยังเป็นของเก่า = deploy ยังไม่ขึ้น (หรือยังไม่ได้กด Ctrl+Shift+R) */
  VERSION   : 'v3.1',
  BUILD     : 43,
  YEAR      : 2569,
  NMONTH    : 8,                 // เดือนที่มีข้อมูลแล้ว
  KEMREX    : 5018,              // ✅ รหัสแผนก KEMREX (เบียร์ยืนยันแล้ว) — อยู่ใต้หน่วยงาน 5000 PRODUCTION
  KEMREX_SPLIT: true,            // true = ยก KEMREX ขึ้นเป็น "แผนกใหญ่" แยกบนหน้าจอ (แม้ไฟล์จะอยู่ใต้ PRODUCTION)
  PERIOD_START: 26,              // งวดจ่าย = วันที่ 26 เดือนก่อน ถึง 25 เดือนนี้ (ตรงกับไฟล์ Bplus)
  /* เดือน 1–7/2569 ยกเวลามาจากระบบเก่า → เทียบ Bplus ↔ JOBTRACK ไม่ตรงเป็นเรื่องปกติ
     ด่าน "เช็คเวลา" ของเดือนเหล่านี้จึงเป็นข้อมูลอ้างอิงเท่านั้น ไม่บล็อกการปิดงวด
     แต่ด่าน "เช็คยอดเงิน" (ต้นทุนลงจ๊อบ = BP จ่ายจริง) ยังเข้มเหมือนเดิมทุกเดือน
     เดือนไหนเริ่มใช้ JOBTRACK จริงเต็มตัวแล้ว ให้ลดเลขนี้ลง */
  LEGACY_TIME_UNTIL: 7,
  SHOW_ALL_JOBS: false,          // true = โชว์จ๊อบเก่าที่ไม่มียอดอะไรเลยด้วย
  NPROC     : 7,                 // โชว์ Process 7 อันดับแรกแยกสี ที่เหลือรวมเป็น "อื่นๆ" (จานสีมี 8 ช่อง)
  CACHE_KEY : 'NOVA_PAYLOAD_v3_1_0',   // ⚠️ บั๊มเลขนี้ทุกครั้งที่แก้ logic (กันเว็บโชว์ของเก่า)
  CACHE_SEC : 21600,        // แคชในหน่วยความจำ 6 ชม. (สูงสุดที่ Google ให้)
  SNAP_MAX_SEC : 86400      // ไฟล์สแนปช็อตเก่าได้ไม่เกิน 24 ชม. (ตัวตั้งเวลาสร้างใหม่ทุก 1 ชม.)
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
      ver: CFG.VERSION, build: CFG.BUILD,
      /* เดือนที่ไฟล์ Bplus ถูกแปลงเป็น Google Sheets แล้ว → หน้า ① Time Bplus ใช้ตัวนี้ตัดสิน */
      bplusMonths: Object.keys(bplusFiles_()).map(Number).sort(function (a, b) { return a - b; }),
      legacyTimeUntil: CFG.LEGACY_TIME_UNTIL,
      kemrex: CFG.KEMREX,
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
  D.verify = verify_(D, cal.warn.concat(CFG._budgetWarn || []));
  return D;
}

/* ============================================================================
   แคชแบบ "หั่นเป็นชิ้น"
   CacheService เก็บได้ชิ้นละ 100KB แต่ข้อมูลจริงใหญ่หลาย MB
   ของเดิม put() ทั้งก้อนแล้วพัง → แคชไม่ติดสักครั้ง → ทุกครั้งที่เปิดเว็บต้อง
   ไล่อ่านไฟล์ Cal 7 เดือน + Bplus 7 เดือน + Master + Payroll DB ใหม่หมด (~1 นาที)
   ตอนนี้หั่นเป็นชิ้นละ 20,000 ตัวอักษร (ภาษาไทยตัวละ 3 ไบต์ → ยังไม่ถึง 60KB ปลอดภัย)
   ============================================================================ */
var NV_CHUNK = 20000;
var NV_MAXCHUNK = 250;

function nvCacheGet_(key) {
  var c = CacheService.getScriptCache();
  var n = c.get(key + '_N');
  if (!n) return null;
  n = Number(n);
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var all = c.getAll(keys), out = '';
  for (var j = 0; j < n; j++) {
    var part = all[key + '_' + j];
    if (part === undefined || part === null) return null;   // ชิ้นไหนหมดอายุ = ใช้ไม่ได้ทั้งก้อน
    out += part;
  }
  return out;
}
function nvCachePut_(key, s, sec) {
  var n = Math.ceil(s.length / NV_CHUNK);
  if (n > NV_MAXCHUNK) return;                              // ใหญ่เกินไปจริงๆ → ยอมไม่แคช
  var o = {};
  for (var i = 0; i < n; i++) o[key + '_' + i] = s.substr(i * NV_CHUNK, NV_CHUNK);
  o[key + '_N'] = String(n);
  try { CacheService.getScriptCache().putAll(o, sec); } catch (e) { /* แคชไม่ได้ก็ยังทำงานได้ แค่ช้า */ }
}
function nvCacheClear_(key) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get(key + '_N') || 0), keys = [key + '_N'];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  try { c.removeAll(keys); } catch (e) {}
}

/* ============================================================================
   ทำไมต้องมี "สแนปช็อต"
   สร้างข้อมูลรอบหนึ่งต้องเปิดไฟล์ Cal 7 เดือน + Bplus 7 เดือน + Master + Payroll DB
   ใช้เวลาราว 1 นาที — ถ้าปล่อยให้สร้างตอนคนกดเข้าเว็บ ทุกคนต้องนั่งรอ 1 นาที
   จึงเปลี่ยนเป็น: ตัวตั้งเวลาสร้างไว้ล่วงหน้าเก็บเป็นไฟล์ JSON ในไดรฟ์
   คนเข้าเว็บแค่ "หยิบไฟล์ที่ทำไว้แล้ว" → ขึ้นทันที 2-3 วินาที
   ลำดับการหา:  แคชในหน่วยความจำ → ไฟล์สแนปช็อต → สร้างใหม่ (ช้า ใช้เมื่อจำเป็นจริงๆ)
   ============================================================================ */
var NV_SNAP_NAME = 'NOVA_PAYLOAD_SNAPSHOT.json';

function nvSnapFolder_() {
  var it = DriveApp.getFileById(FILES.MASTER).getParents();
  return it.hasNext() ? it.next() : DriveApp.getRootFolder();
}
function nvSnapFile_() {
  var it = nvSnapFolder_().getFilesByName(NV_SNAP_NAME);
  return it.hasNext() ? it.next() : null;
}
function nvSnapRead_() {
  try {
    var f = nvSnapFile_(); if (!f) return null;
    var ageSec = (new Date().getTime() - f.getLastUpdated().getTime()) / 1000;
    if (ageSec > CFG.SNAP_MAX_SEC) return null;          // เก่าเกินไป → สร้างใหม่
    return f.getBlob().getDataAsString('UTF-8');
  } catch (e) { return null; }
}
function nvSnapWrite_(s) {
  try {
    var f = nvSnapFile_();
    if (f) f.setContent(s);
    else nvSnapFolder_().createFile(NV_SNAP_NAME, s, 'application/json');
  } catch (e) { /* เขียนไม่ได้ก็ยังทำงานได้ แค่ช้า */ }
}

/** ตัวที่หน้าเว็บใช้ — เร็วที่สุดเท่าที่จะเป็นไปได้ */
function getPayloadJson() {
  var hit = nvCacheGet_(CFG.CACHE_KEY);
  if (hit) return hit;                                    // ① แคช (เร็วสุด)
  var snap = nvSnapRead_();
  if (snap) { nvCachePut_(CFG.CACHE_KEY, snap, CFG.CACHE_SEC); return snap; }   // ② ไฟล์สแนปช็อต
  return rebuildSnapshot();                               // ③ สร้างใหม่ (~1 นาที)
}

/** สร้างข้อมูลใหม่ทั้งก้อน แล้วเก็บทั้งไฟล์และแคช — ตัวตั้งเวลาเรียกตัวนี้ */
function rebuildSnapshot() {
  var t0 = new Date().getTime();
  var s = JSON.stringify(buildPayload()).replace(/<\//g, '<\\/');   // กัน '</script>' ที่อาจหลุดมาในชื่องาน
  nvSnapWrite_(s);
  nvCachePut_(CFG.CACHE_KEY, s, CFG.CACHE_SEC);
  Logger.log('สร้างข้อมูลใหม่เสร็จ ' + Math.round(s.length / 1024) + ' KB · ใช้เวลา ' +
             Math.round((new Date().getTime() - t0) / 1000) + ' วินาที');
  return s;
}

/**
 * ⭐ รันครั้งเดียวหลัง deploy ครั้งแรก — ตั้งให้ระบบสร้างข้อมูลล่วงหน้าทุกชั่วโมง
 * ผลคือคนเข้าเว็บไม่ต้องรอสร้างข้อมูลอีกเลย
 */
function installSnapshotTrigger() {
  try { ScriptApp.getProjectTriggers(); }
  catch (e) {
    throw new Error('ยังไม่ได้ให้สิทธิ์ตั้งตัวตั้งเวลา — กด Review permissions → Allow แล้วรันฟังก์ชันนี้ใหม่อีกครั้ง\n' +
                    '(ถ้าไม่ขึ้นหน้าต่างขอสิทธิ์ ให้กด Run ซ้ำอีกรอบ)');
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rebuildSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rebuildSnapshot').timeBased().everyHours(1).create();
  var s = rebuildSnapshot();                              // สร้างรอบแรกให้เลย
  Logger.log('✅ ตั้งตัวตั้งเวลาแล้ว — ระบบจะสร้างข้อมูลใหม่ทุก 1 ชั่วโมงโดยอัตโนมัติ');
  return 'ตั้งเวลาเรียบร้อย · ข้อมูลชุดแรก ' + Math.round(s.length / 1024) + ' KB';
}

function forceRefresh() {
  nvCacheClear_(CFG.CACHE_KEY);
  var s = rebuildSnapshot();
  return 'ดึงข้อมูลสดจากไฟล์ต้นทางใหม่แล้ว (' + Math.round(s.length / 1024) + ' KB)';
}

/* ============================================================================
   probe() — รันใน Apps Script Editor เพื่อ "ส่องไฟล์จริง" ก่อน deploy
   ดูผลที่ View → Logs · ถ้าแท็บไหนขึ้น ✗ แปลว่าชื่อแท็บในไฟล์เปลี่ยน ต้องมาแก้ที่ findTab_
   ============================================================================ */
function probe() {
  var L = ['NOVA-HR ' + CFG.VERSION + ' · build ' + CFG.BUILD +
           '   (เวลาที่รัน ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd MMM yyyy HH:mm:ss') + ')',
           '────────────────────────────────────────────'];
  var cf = calFiles_();
  L.push('พบไฟล์ Cal ' + cf.length + ' เดือน: ' + cf.map(function (x) { return x.m + '=' + x.name; }).join(' | '));
  var bf = bplusFiles_();
  L.push('พบไฟล์ Bplus (แปลงเป็น Google Sheets แล้ว) เดือน: ' + Object.keys(bf).join(', '));
  if (cf.length) {
    var ss = SpreadsheetApp.openById(cf[cf.length - 1].id);
    L.push('แท็บทั้งหมดในไฟล์ ' + cf[cf.length - 1].name + ':');
    ss.getSheets().forEach(function (s) { L.push('   • ' + s.getName() + '  (' + s.getLastRow() + ' แถว)'); });
    /* ต้องใช้ "ชื่อที่โค้ดจริงค้นหา" ไม่ใช่ชื่อที่เราเรียกกันเอง ไม่งั้นจะขึ้น ✗ หลอก */
    [['JOB_COST_DIRECT', ['JOB_COST_DIRECT']],
     ['PAYROLL_SUMMARY', ['PAYROLL_SUMMARY']],
     ['PERFORMANCE',     ['PERFORMANCE']],
     ['ตารางเงินเดือน',   ['PAYROLL_ACTUAL', 'สรุปตารางเงินเดือน']],
     ['ปฏิทินวันทำงาน',   ['CALENDAR_MASTER', 'ปฏิทินวันทำงาน']],
     ['ประเภทชั่วโมง',    ['HOUR_TYPE_RULE', 'WORK_HOUR_TYPE']],
     ['ALL WIP',         ['ALL WIP']]
    ].forEach(function (t) {
      var sh = findTab_(ss, t[1]);
      L.push((sh ? '✓ ' : '✗ ') + t[0] + (sh ? ' → ' + sh.getName() : ' → หาไม่เจอ'));
    });
  }
  try {
    var D = buildPayload();
    L.push('— ผลประกอบร่าง —');
    L.push('พนักงาน ' + D.emps.STT.length + ' คน · แผนกใหญ่ ' + D.depts.length + ' · Process ' + D.processes.length);
    L.push('เงินเดือน ' + D.companies.STT.rows.length + ' แถว · ลงจ๊อบ ' + D.companies.STT.jobRows.length +
           ' แถว · Performance ' + D.companies.STT.perf.length + ' แถว · Reconcile ' + D.companies.STT.recon.length + ' แถว');
    L.push('จ๊อบ ' + D.jobs.length + ' · เดือนที่มีข้อมูล ' + D.meta.nmonth);
    L.push('เดือนที่อ่าน Bplus ได้: ' + D.meta.bplusMonths.join(', '));
    L.push(diagBplus_());
    L.push('เติมพนักงานที่ไม่มีในทะเบียนจากไฟล์ Cal: ' + (CFG._added || 0) + ' คน');
    L.push('แผนกใหญ่: ' + D.depts.map(function (d) { return d.code + '=' + d.name; }).join(' · '));
    var sf = nvSnapFile_();
    L.push('ไฟล์สแนปช็อต: ' + (sf ? 'มีแล้ว · อัปเดตล่าสุด ' +
      Utilities.formatDate(sf.getLastUpdated(), 'Asia/Bangkok', 'd MMM yyyy HH:mm') +
      ' · ' + Math.round(sf.getSize()/1024) + ' KB'
      : '❌ ยังไม่มี — รัน installSnapshotTrigger() หนึ่งครั้ง ไม่งั้นเข้าเว็บจะรอ ~1 นาทีทุกครั้ง'));
    try {
      L.push('ตัวตั้งเวลาสร้างข้อมูลอัตโนมัติ: ' +
        (ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='rebuildSnapshot';}).length
          ? '✓ ตั้งแล้ว (ทุก 1 ชั่วโมง)' : '❌ ยังไม่ได้ตั้ง — รัน installSnapshotTrigger()'));
    } catch (e) {
      L.push('ตัวตั้งเวลา: ยังไม่ได้ให้สิทธิ์ → รัน installSnapshotTrigger() แล้วกด Allow หนึ่งครั้ง');
    }
    L.push('ด่านตรวจ ผ่าน ' + D.verify.pass + '/' + D.verify.total);
    D.verify.fails.forEach(function (f) { L.push('   ✗ ' + f); });
    var nEst = 0, nSale = 0;
    D.jobs.forEach(function (j) { if (j.estBudget) nEst++; if (j.saleBudget) nSale++; });
    L.push('งบ: มี Sale Budget ' + nSale + ' จ๊อบ · Est. Budget ' + nEst + ' จ๊อบ (จากทั้งหมด ' + D.jobs.length + ')');
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

/** พิมพ์ชื่อคอลัมน์จริงของไฟล์ Bplus ทุกเดือน — ไว้เช็กว่าเดือนไหนสะกดไม่เหมือนกัน */
function diagBplus_() {
  var bf = bplusFiles_(), L = ['— คอลัมน์จริงในไฟล์ Bplus แต่ละเดือน —'];
  Object.keys(bf).sort(function (a, b) { return a - b; }).forEach(function (m) {
    try {
      var sh = SpreadsheetApp.openById(bf[m]).getSheets()[0];
      var t = nvReadSheet_(sh);
      var need = ['ชม.งาน', 'มาสาย', 'ขาดงาน', 'เวลารูดบัตร'];
      var miss = [];
      need.forEach(function (n) { if (t.head.indexOf(n) < 0) miss.push(n); });
      var forget = t.head.filter(function (h) { return /ืมรูดบัตร/.test(h); });
      L.push('  เดือน ' + m + ': ' + t.rows.length + ' แถว · คอลัมน์ลืมรูดบัตร = "' +
             (forget[0] || '❌ ไม่มี') + '"' + (miss.length ? ' · ❌ ขาด ' + miss.join(',') : ' ✓'));
    } catch (e) { L.push('  เดือน ' + m + ': อ่านไม่ได้ — ' + e.message); }
  });
  return L.join('\n');
}
