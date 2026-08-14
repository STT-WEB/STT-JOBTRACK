/***********************************************************
 * STT NOVA-HR Hub  —  Job Cost (Part A)
 * ---------------------------------------------------------
 * ไฟล์นี้เป็น "โปรเจกต์ใหม่" แยกจาก JOBTRACK เดิม
 * ไม่ยุ่งกับระบบที่พนักงานใช้อยู่
 *
 * ▶ ขั้นแรก: รันฟังก์ชัน  setupJobcost2026  หนึ่งครั้ง
 *   จะได้ไฟล์ผลลัพธ์ "JOBCOST 2026" (3 แท็บ พร้อมหัวคอลัมน์)
 *   แล้วดู Log เพื่อก๊อปลิงก์/ไฟล์ ID มาให้ Candy
 ***********************************************************/

var JOBCOST_YEAR = 2026;
var HUB_FOLDER_PATH = ['STT NOVA-HR Hub', 'STT', '02_JOBCOST'];

/**
 * สร้างไฟล์ JOBCOST 2026 (แท็บ MASTER / ผลรายเดือน / Reconcile)
 * รันครั้งเดียว — กันสร้างซ้ำอัตโนมัติ
 */
function setupJobcost2026() {
  var name = 'JOBCOST ' + JOBCOST_YEAR;
  var folder = getOrCreatePath_(HUB_FOLDER_PATH);

  // กันสร้างซ้ำ: ถ้ามีไฟล์ชื่อเดียวกันอยู่แล้ว ให้คืนของเดิม
  var existing = folder.getFilesByName(name);
  if (existing.hasNext()) {
    var f = existing.next();
    Logger.log('⚠️ มีไฟล์อยู่แล้ว ไม่สร้างซ้ำ');
    Logger.log('ลิงก์: ' + f.getUrl());
    Logger.log('File ID: ' + f.getId());
    return f.getUrl();
  }

  var ss = SpreadsheetApp.create(name);

  // ---- แท็บ 1: MASTER (ต่อจ๊อบทั้งปี) ----
  var master = ss.getSheets()[0];
  master.setName('MASTER');
  writeHeader_(master, [
    'ประเภทงาน', 'Job Code', 'ชื่องาน',
    'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
    'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.',
    'ยอดยกมา', 'รวมค่าแรง STT', 'ค่าแรงผู้รับเหมา', 'Total Cost',
    'Est. Budget', 'Sale Budget', 'ส่วนต่าง Budget', '% Cost/Budget', 'สถานะ'
  ], '#E2231A');

  // ---- แท็บ 2: ผลรายเดือน (engine เขียนผลลงที่นี่) ----
  var monthly = ss.insertSheet('ผลรายเดือน');
  writeHeader_(monthly, [
    'งวด', 'Job Code', 'ชื่องาน', 'Process',
    'รหัสพนักงาน', 'ชื่อพนักงาน', 'ประเภทพนักงาน',
    'ชม.คิดค่าแรง', 'Rate/ชม.', 'ต้นทุน (บาท)', 'อัปเดตเมื่อ'
  ], '#185FA5');

  // ---- แท็บ 3: Reconcile (เทียบเงินเดือน + Performance) ----
  var recon = ss.insertSheet('Reconcile');
  writeHeader_(recon, [
    'งวด', 'รหัสพนักงาน', 'ชื่อพนักงาน', 'ประเภทพนักงาน',
    'ชม.ควรทำ', 'ชม.ทำจริง', 'ชม.หาย',
    'ต้นทุนลงจ๊อบ', 'เงินเดือนจ่ายจริง (BP)', 'ส่วนต่าง (dif)', 'สถานะ'
  ], '#1B7A3D');

  // ย้ายไฟล์เข้าโฟลเดอร์ STT NOVA-HR Hub/STT/02_JOBCOST
  DriveApp.getFileById(ss.getId()).moveTo(folder);

  Logger.log('✅ สร้างไฟล์ ' + name + ' สำเร็จ');
  Logger.log('📁 อยู่ในโฟลเดอร์: ' + HUB_FOLDER_PATH.join(' / '));
  Logger.log('🔗 ลิงก์: ' + ss.getUrl());
  Logger.log('🆔 File ID: ' + ss.getId());
  Logger.log('👉 ก๊อป "File ID" บรรทัดบนส่งให้ Candy เพื่อทำ engine ต่อ');
  return ss.getUrl();
}

/**
 * ★ ตั้งค่าโครงไฟล์ Master (ไฟล์เปล่าที่เบียร์สร้างเอง = JOBCOST_FILE_ID)
 *   สร้าง 3 แท็บ + หัวคอลัมน์ · รันครั้งเดียวหลังสร้างไฟล์เปล่า
 *   → จากนั้นรัน rebuildJobcostFromMonthly เพื่อเติมข้อมูลเดือน 1–7
 */
function setupMasterFile() {
  var ss = SpreadsheetApp.openById(JOBCOST_FILE_ID);
  var master = ss.getSheetByName('MASTER') || ss.getSheets()[0];
  master.setName('MASTER');
  writeHeader_(master, [
    'ประเภทงาน','Job Code','ชื่องาน',
    'ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.',
    'ยอดยกมา','รวมค่าแรง STT','ค่าแรงผู้รับเหมา','Total Cost',
    'Est. Budget','Sale Budget','ส่วนต่าง Budget','% Cost/Budget','สถานะ'
  ], '#E2231A');
  var monthly = ss.getSheetByName('ผลรายเดือน') || ss.insertSheet('ผลรายเดือน');
  writeHeader_(monthly, [
    'งวด','Job Code','ชื่องาน','Process','รหัสพนักงาน','ชื่อพนักงาน','ประเภทพนักงาน',
    'ชม.คิดค่าแรง','Rate/ชม.','ต้นทุน (บาท)','อัปเดตเมื่อ'
  ], '#185FA5');
  var recon = ss.getSheetByName('Reconcile') || ss.insertSheet('Reconcile');
  writeHeader_(recon, [
    'งวด','รหัสพนักงาน','ชื่อพนักงาน','ประเภทพนักงาน',
    'ชม.ควรทำ','ชม.ทำจริง','ชม.หาย','ต้นทุนลงจ๊อบ','เงินเดือนจ่ายจริง (BP)','ส่วนต่าง (dif)','สถานะ'
  ], '#1B7A3D');
  Logger.log('✅ ตั้งค่าโครงไฟล์ Master เสร็จ (3 แท็บ): ' + ss.getUrl());
  Logger.log('👉 ต่อไปรัน rebuildJobcostFromMonthly เพื่อเติมข้อมูลเดือน 1–7');
}

/***********************************************************
 *  FILE REGISTRY — ทะเบียนไฟล์ (อยู่ในไฟล์ Payroll, แท็บ "File Registry")
 *  layout: ฟิลด์อยู่คอลัมน์ A · ปีเป็นคอลัมน์ B, C, ... (แถว 1 = ปี)
 *  ▶ testRegistry() = ทดสอบว่าอ่าน ID ครบไหม
 ***********************************************************/
function extractDriveId_(v) {
  var s = String(v || '').trim();
  var m = s.match(/[-\w]{25,}/);   // Drive ID = โทเคนยาว ≥25 ตัว (รองรับทั้ง ID ล้วนและ URL)
  return m ? m[0] : s;
}

/** อ่านทะเบียนไฟล์ของปีที่ระบุ → คืน object {employeeDB, jobLog, timeBplusFolder, calFolder, master, salaryPayroll, payrollActual} */
function getRegistry_(year) {
  var sh = SpreadsheetApp.openById(SALARY_DB_ID).getSheetByName('File Registry');
  if (!sh) throw new Error('ไม่พบแท็บ File Registry ในไฟล์ Payroll');
  var vals = sh.getDataRange().getValues();
  var yCol = -1;
  for (var c = 1; c < vals[0].length; c++) if (String(vals[0][c]).trim() === String(year)) { yCol = c; break; }
  if (yCol < 0) throw new Error('ไม่พบปี ' + year + ' ใน File Registry');
  var reg = {};
  for (var r = 1; r < vals.length; r++) {
    var label = String(vals[r][0] || '');
    var val = extractDriveId_(vals[r][yCol]);
    if (!label || !val) continue;
    if (label.indexOf('ฐานพนักงาน') >= 0)                 reg.employeeDB = val;
    else if (label.indexOf('Job_Log') >= 0)              reg.jobLog = val;
    else if (label.indexOf('Time Bplus') >= 0)           reg.timeBplusFolder = val;
    else if (label.indexOf('Cal') >= 0)                  reg.calFolder = val;
    else if (label.indexOf('Salary Master') >= 0)        reg.salaryPayroll = val;
    else if (label.indexOf('Master') >= 0 || label.indexOf('Dashboard') >= 0) reg.master = val;
    else if (label.indexOf('อัตราต่องวด') >= 0 || label.indexOf('ตารางเงินเดือนจริง') >= 0) reg.payrollActual = val;
  }
  return reg;
}

/** ทดสอบอ่าน Registry ปี 2026 — ดู Log ว่าครบไหม */
function testRegistry() {
  var reg = getRegistry_(2026);
  Logger.log('📋 File Registry ปี 2026:');
  ['employeeDB','jobLog','timeBplusFolder','calFolder','master','salaryPayroll','payrollActual']
    .forEach(function(k){ Logger.log('   ' + k + ' = ' + (reg[k] || '❌ ขาด')); });
  return reg;
}

/** เขียนหัวคอลัมน์ + จัดรูปแบบ (ตัวหนา พื้นสี ตรึงแถวแรก) */
function writeHeader_(sheet, head, color) {
  sheet.getRange(1, 1, 1, head.length)
    .setValues([head])
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground(color)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 30);
  for (var c = 1; c <= head.length; c++) {
    sheet.setColumnWidth(c, 110);
  }
}

/** สร้าง/หาโฟลเดอร์ตาม path (array) — คืนโฟลเดอร์ปลายทาง */
function getOrCreatePath_(parts) {
  var parent = DriveApp.getRootFolder();
  for (var i = 0; i < parts.length; i++) {
    var it = parent.getFoldersByName(parts[i]);
    parent = it.hasNext() ? it.next() : parent.createFolder(parts[i]);
  }
  return parent;
}


/***********************************************************
 *  ENGINE — ต้นทุนค่าแรง (ทำตามสูตรระบบเก่าเป๊ะ)
 *  ▶ recheckJanuary()          = ทดสอบ: รันเลข ม.ค. เทียบไฟล์เก่า
 *  ▶ computeJobcostMonth('2026-08') = คำนวณเดือนใหม่จาก JOBTRACK
 ***********************************************************/

// ---- Config (ไอดีไฟล์จริง) ----
var JOBCOST_FILE_ID = '1gzryio6lwbhozn19-aOUGkzScNtk1bvJGk5tUw-gcbM';   // Jobcost Master/Dashboard 2026 (ไฟล์ใหม่สะอาด)
var DB_ID           = '1MYWORYN3sOjov3Gxv3UqCV1jRSxgxwGi1tRomFUGSr0';   // STT Jobcost Database-Employee & Account
var SALARY_DB_ID    = '1hw5cyLEZO3dEol5X1pX3_2oBYtuJmADTTMwrOsll1D8';   // STT Jobcost Database-Payroll (ลับ) — Salary Master อยู่ที่นี่
var LOG_ID_BY_YEAR  = { '2026': '1ZPl3uVRtM5r4sPA-yX1OwTyp34XCTIcKRC0Sx8qsr9s' };
var SETUP_TAB       = 'Salary Master';   // เดิมชื่อ 'SETUP & MASTER Employee' ย้ายมาไฟล์ Payroll

// ไฟล์ระบบเก่า (ใช้ recheck)
var OLD_CAL_ID    = '1keVRkC7tTWvEJGbsGsGBzOY0PcIsg-dGDSeIPgQB58U';   // Cal JOB COST JAN (JOB_COST_DIRECT)
var OLD_MASTER_ID = '10u8sLbe6dtGtFtWwAMruvAgrawDOScqEBmJRA3avPcM';   // MASTER (ACTUAL_WIDE)

// ตำแหน่งคอลัมน์ Job_Log (0-based) — ตรงกับ JOBTRACK
var JL = {JOB_ID:5, JOB_NAME:6, EMP_ID:7, EMP_NAME:8, EMP_TYPE:10,
          PROC_MAIN:12, PROC_CODE:13, PROC_SUB:14, STATUS:16,
          DAY_TYPE:19, HOURS_NORMAL:21, HOURS_OT:22};

/**
 * ต้นทุน 1 บรรทัด ตามสูตรระบบเก่า
 *   code 1 ปกติ      : ชม × normalRate
 *   code 3 OT        : ชม × otRate × 1.5
 *   code 2 วันหยุด    : รายเดือน ชม×(normalRate+otRate) | รายวัน ชม×otRate×2
 *   code 4 OT วันหยุด : ชม × otRate × 3
 */
function costLine_(code, hours, normalRate, otRate, isMonthly) {
  hours = Number(hours) || 0;
  code = String(code).charAt(0);
  if (code === '1') return hours * normalRate;
  if (code === '3') return hours * otRate * 1.5;
  if (code === '2') return isMonthly ? hours * (normalRate + otRate) : hours * otRate * 2;
  if (code === '4') return hours * otRate * 3;
  return hours * normalRate;
}

/**
 * ทดสอบ engine กับข้อมูลมกราคมของระบบเก่า:
 *  1) อ่านตาราง JOB_COST_DIRECT ในไฟล์ Cal เก่า
 *  2) คำนวณต้นทุนแต่ละบรรทัดใหม่ด้วยสูตรเรา → เทียบกับ "ต้นทุนค่าแรง" ในไฟล์
 *  3) รวมต่อจ๊อบ → เทียบกับ MASTER ACTUAL_WIDE คอลัมน์ 2026-01
 */
function recheckJanuary() {
  // (1) อ่าน JOB_COST_DIRECT
  var cal = SpreadsheetApp.openById(OLD_CAL_ID);
  var hit = findSheetByHeader_(cal, ['ต้นทุนค่าแรง', 'รวมที่ใช้หาร']);
  if (!hit) { Logger.log('❌ หาตาราง JOB_COST_DIRECT ไม่เจอในไฟล์ Cal'); return; }
  var sheet = hit.sheet, head = hit.head, hrow = hit.row;
  var cEmp  = findCol_(head, ['รหัสพนักงาน']);
  var cJob  = findCol_(head, ['JOB CODE']);
  var cHT   = findCol_(head, ['ประเภทชั่วโมง']);
  var cHrs  = findCol_(head, ['จำนวนชั่วโมง']);
  var cType = findCol_(head, ['Employee Type', 'ประเภทพนักงาน', 'ปรเภท']);
  var cBase = findCol_(head, ['ยอดฐาน']);
  var cOT   = findCol_(head, ['OT Rate']);
  var cD    = findCol_(head, ['รวมที่ใช้หาร']);
  var cCost = findCol_(head, ['ต้นทุนค่าแรง']);

  var data = sheet.getDataRange().getValues();
  var jobOur = {};             // jobCode -> ต้นทุนรวม (สูตรเรา) เฉพาะงานลูกค้า
  var lines = 0, mism = 0, maxDiff = 0, sumOur = 0, sumFile = 0;
  var misLines = [];

  for (var i = hrow; i < data.length; i++) {
    var r = data[i];
    var emp = String(r[cEmp]).trim();
    if (!emp) continue;
    var hrs = Number(r[cHrs]) || 0;
    if (hrs <= 0) continue;
    var code = String(r[cHT]).charAt(0);
    var base = Number(r[cBase]) || 0;
    var otR  = Number(r[cOT]) || 0;
    var D    = Number(r[cD]) || 0;
    if (D <= 0) continue;
    var isMonthly = String(r[cType]).indexOf('เดือน') >= 0;
    var normalRate = base / D;
    var our = costLine_(code, hrs, normalRate, otR, isMonthly);
    var file = Number(r[cCost]) || 0;

    var d = Math.abs(our - file);
    if (d > 0.5) {
      mism++; if (d > maxDiff) maxDiff = d;
      if (misLines.length < 15) misLines.push('  ' + emp + ' | ' + String(r[cJob]).trim() +
        ' | code=' + String(r[cHT]).trim() + ' | ชม.' + hrs + ' | เรา ' + our.toFixed(2) + ' vs ไฟล์ ' + file.toFixed(2));
    }
    sumOur += our; sumFile += file; lines++;

    var job = String(r[cJob]).trim();
    // ตัดงานภายใน ST- / งานที่ไม่มีรหัสจ๊อบจริง ออกจากการเทียบต่อจ๊อบ
    if (job && !/^ST/i.test(job) && job.indexOf('-') > 0) jobOur[job] = (jobOur[job] || 0) + our;
  }

  Logger.log('===== RECHECK มกราคม =====');
  Logger.log('บรรทัดที่คำนวณ: ' + lines);
  Logger.log('ตรงกับไฟล์ (±0.5): ' + (lines - mism) + ' | ไม่ตรง: ' + mism + ' | ต่างสูงสุด: ' + maxDiff.toFixed(2));
  Logger.log('รวมสูตรเรา: ' + Math.round(sumOur).toLocaleString() + ' | รวมไฟล์เก่า: ' + Math.round(sumFile).toLocaleString());
  if (misLines.length) Logger.log('บรรทัดที่ไม่ตรง (สูงสุด 15):\n' + misLines.join('\n'));

  // (3) เทียบต่อจ๊อบกับ MASTER 2026-01
  var master = SpreadsheetApp.openById(OLD_MASTER_ID);
  var mh = findSheetByHeader_(master, ['JOB CODE', '2026-01']);
  if (!mh) { Logger.log('⚠️ หาแท็บ ACTUAL_WIDE ไม่เจอ'); return; }
  var mJob = findCol_(mh.head, ['JOB CODE']);
  var mJan = findCol_(mh.head, ['2026-01']);
  var mData = mh.sheet.getDataRange().getValues();
  var jobFile = {};
  for (var j = mh.row; j < mData.length; j++) {
    var jc = String(mData[j][mJob]).trim();
    if (jc) jobFile[jc] = Number(mData[j][mJan]) || 0;
  }

  var okJ = 0, badJ = 0, samples = [];
  for (var jc2 in jobOur) {
    var mine = jobOur[jc2], theirs = jobFile[jc2] || 0;
    if (Math.abs(mine - theirs) <= 1) okJ++;
    else { badJ++; if (samples.length < 8) samples.push(jc2 + ': เรา ' + Math.round(mine).toLocaleString() + ' vs เก่า ' + Math.round(theirs).toLocaleString()); }
  }
  Logger.log('----- ต่อจ๊อบ -----');
  Logger.log('จ๊อบตรง: ' + okJ + ' | ไม่ตรง: ' + badJ);
  if (samples.length) Logger.log('ตัวอย่างไม่ตรง:\n  ' + samples.join('\n  '));
  Logger.log('เทียบจ๊อบตัวอย่าง JT-68/0058: เรา ' + Math.round(jobOur['JT-68/0058']||0).toLocaleString() + ' (เก่า 84,554)');
}

/**
 * คำนวณต้นทุนเดือนใหม่จาก JOBTRACK (สูตรเดียวกับระบบเก่า) → เขียนลง JOBCOST 2026
 */
function computeJobcostMonth(period) {
  var year  = parseInt(period.split('-')[0], 10);
  var month = parseInt(period.split('-')[1], 10);
  var mm = ('0' + month).slice(-2);
  var logTab = 'Job_Log_' + (year + 543) + '_' + mm;

  var logId = LOG_ID_BY_YEAR[String(year)];
  if (!logId) throw new Error('ไม่พบไฟล์ Job_Log ของปี ' + year);
  var logSheet = SpreadsheetApp.openById(logId).getSheetByName(logTab);
  if (!logSheet) { Logger.log('⚠️ ไม่พบแท็บ ' + logTab + ' (ยังไม่มีข้อมูลเดือนนี้)'); return; }

  var emp = loadEmployeeMaster_();          // id -> {base, salary, isMonthly}
  var data = logSheet.getDataRange().getValues();

  // รอบ 1: หา D ต่อคน = Σ ชม.ปกติ (รวมทั้งวันปกติ+วันหยุด)
  var Dmap = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[JL.STATUS]).trim() !== 'Check Out') continue;
    var id = String(r[JL.EMP_ID]).trim();
    Dmap[id] = (Dmap[id] || 0) + (Number(r[JL.HOURS_NORMAL]) || 0);
  }

  // รอบ 2: คิดต้นทุนต่อบรรทัด
  var agg = {}, jobTotal = {}, used = 0, skipped = 0;
  for (var k = 1; k < data.length; k++) {
    var row = data[k];
    if (String(row[JL.STATUS]).trim() !== 'Check Out') continue;
    var empId = String(row[JL.EMP_ID]).trim();
    var e = emp[empId], D = Dmap[empId];
    if (!e || !D) { skipped++; continue; }

    var normalH = Number(row[JL.HOURS_NORMAL]) || 0;
    var otH     = Number(row[JL.HOURS_OT]) || 0;
    if (normalH + otH <= 0) continue;
    var isHoliday = /หยุด|นักขัต/.test(String(row[JL.DAY_TYPE]));
    var normalRate = e.base / D;
    var otRate = e.salary / (e.isMonthly ? 240 : 208);

    var cost = costLine_(isHoliday ? '2' : '1', normalH, normalRate, otRate, e.isMonthly)
             + costLine_(isHoliday ? '4' : '3', otH,     normalRate, otRate, e.isMonthly);

    var jobId = String(row[JL.JOB_ID]).trim();
    var jobName = String(row[JL.JOB_NAME]).trim();
    var proc = (String(row[JL.PROC_CODE]).trim() + ' ' + String(row[JL.PROC_SUB]).trim()).trim() || String(row[JL.PROC_MAIN]).trim();
    var key = jobId + '|' + proc + '|' + empId;
    if (!agg[key]) agg[key] = {period:period, jobId:jobId, jobName:jobName, proc:proc,
                    empId:empId, empName:String(row[JL.EMP_NAME]).trim(),
                    empType:(e.isMonthly?'รายเดือน':'รายวัน'), hrs:0, cost:0};
    agg[key].hrs  += normalH + otH;
    agg[key].cost += cost;
    if (!jobTotal[jobId]) jobTotal[jobId] = {name:jobName, total:0};
    jobTotal[jobId].total += cost;
    used++;
  }

  writeMonthlyRows_(period, agg);
  updateMaster_(month, jobTotal);

  Logger.log('✅ คำนวณงวด ' + period + ' เสร็จ | แถวใช้ ' + used + ' ข้าม ' + skipped + ' | จ๊อบ ' + Object.keys(jobTotal).length);
  var g = 0; for (var t in jobTotal) g += jobTotal[t].total;
  Logger.log('   รวมต้นทุนค่าแรง: ' + Math.round(g).toLocaleString() + ' บาท');
}

/** โหลด base/salary/type ต่อพนักงาน จาก SETUP & MASTER Employee */
function loadEmployeeMaster_() {
  var sheet = SpreadsheetApp.openById(SALARY_DB_ID).getSheetByName(SETUP_TAB);
  if (!sheet) throw new Error('ไม่พบแท็บ ' + SETUP_TAB + ' ในไฟล์ Payroll');
  var vals = sheet.getDataRange().getValues();
  var head = vals[0].map(function(h){ return String(h).trim(); });
  var cId = findCol_(head, ['รหัสพนักงาน']);
  var cType = findCol_(head, ['ปรเภทพนักงาน', 'ประเภทพนักงาน']);
  var cBase = findCol_(head, ['เงินเดือน+สวัสดิการ']);
  var cSalary = findExact_(head, 'เงินเดือน');   // salary ล้วน (ไม่เอา +สวัสดิการ)
  if (cId < 0 || cType < 0 || cBase < 0 || cSalary < 0)
    throw new Error('หาคอลัมน์ SETUP ไม่ครบ (id/type/base/salary)');
  var map = {};
  for (var i = 1; i < vals.length; i++) {
    var id = String(vals[i][cId]).trim();
    if (!id) continue;
    var base = Number(vals[i][cBase]) || 0;
    if (base <= 0) continue;
    map[id] = {base:base, salary:Number(vals[i][cSalary]) || 0,
               isMonthly:String(vals[i][cType]).indexOf('เดือน') >= 0};
  }
  return map;
}

/** เขียนผลรายเดือน (ลบงวดเดิมก่อน) */
function writeMonthlyRows_(period, agg) {
  var sh = SpreadsheetApp.openById(JOBCOST_FILE_ID).getSheetByName('ผลรายเดือน');
  var last = sh.getLastRow(), now = new Date(), keep = [];
  if (last > 1) {
    var old = sh.getRange(2, 1, last - 1, 11).getValues();
    for (var i = 0; i < old.length; i++) if (String(old[i][0]) !== period) keep.push(old[i]);
  }
  var rows = [];
  for (var key in agg) {
    var a = agg[key];
    rows.push([a.period, a.jobId, a.jobName, a.proc, a.empId, a.empName, a.empType,
               round2_(a.hrs), '', Math.round(a.cost), now]);
  }
  if (last > 1) sh.getRange(2, 1, last - 1, 11).clearContent();
  var all = keep.concat(rows);
  if (all.length) sh.getRange(2, 1, all.length, 11).setValues(all);
}

/** อัปเดต MASTER: ต้นทุนต่อจ๊อบ ลงคอลัมน์เดือน + รวมค่าแรง STT */
function updateMaster_(month, jobTotal) {
  var sh = SpreadsheetApp.openById(JOBCOST_FILE_ID).getSheetByName('MASTER');
  var last = sh.getLastRow();
  var vals = last > 1 ? sh.getRange(2, 1, last - 1, 24).getValues() : [];
  var idx = {};
  for (var i = 0; i < vals.length; i++) idx[String(vals[i][1]).trim()] = i;
  var monthCol0 = 2 + month;
  for (var jobId in jobTotal) {
    var total = jobTotal[jobId].total, name = jobTotal[jobId].name;
    var type = jobId.indexOf('-') > 0 ? jobId.split('-')[0] : '';
    if (idx[jobId] === undefined) {
      var row = new Array(24).fill('');
      row[0] = type; row[1] = jobId; row[2] = name;
      for (var c = 3; c <= 15; c++) row[c] = 0;
      row[monthCol0] = Math.round(total);
      vals.push(row); idx[jobId] = vals.length - 1;
    } else vals[idx[jobId]][monthCol0] = Math.round(total);
  }
  for (var j = 0; j < vals.length; j++) {
    var sum = Number(vals[j][15]) || 0;
    for (var c2 = 3; c2 <= 14; c2++) sum += Number(vals[j][c2]) || 0;
    vals[j][16] = sum;
  }
  if (vals.length) sh.getRange(2, 1, vals.length, 24).setValues(vals);
}

/** หาแท็บที่มีหัวคอลัมน์ครบตาม needles (สแกน 15 แถวแรกหาแถวหัว) */
function findSheetByHeader_(ss, needles) {
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var rng = sheets[s].getRange(1, 1, Math.min(15, sheets[s].getLastRow() || 1),
                                 Math.min(60, sheets[s].getLastColumn() || 1)).getValues();
    for (var rr = 0; rr < rng.length; rr++) {
      var head = rng[rr].map(function(h){ return String(h).trim(); });
      var ok = needles.every(function(n){ return findCol_(head, [n]) >= 0; });
      if (ok) return {sheet:sheets[s], head:head, row:rr + 1};
    }
  }
  return null;
}

/** index คอลัมน์แรกที่หัวมี candidate ตัวใดตัวหนึ่ง (substring) */
function findCol_(head, cands) {
  for (var i = 0; i < head.length; i++)
    for (var c = 0; c < cands.length; c++)
      if (head[i].indexOf(cands[c]) >= 0) return i;
  return -1;
}

/** index คอลัมน์ที่หัว = ข้อความเป๊ะ (ตรงตัว) */
function findExact_(head, text) {
  for (var i = 0; i < head.length; i++) if (head[i] === text) return i;
  return -1;
}

function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
