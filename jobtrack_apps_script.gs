// ============================================================
// JOBTRACK — Google Apps Script Backend v2
// Logic: 1 Check In = 1 Process = 1 Job
//        Check Out → Match กับ Check In → คำนวณชั่วโมง
//        บล็อก Check In ใหม่ถ้ายัง Check Out งานเก่าไม่ครบ
// ============================================================

const SHEET_ID   = '1MYWORYN3sOjov3Gxv3UqCV1jRSxgxwGi1tRomFUGSr0';   // JOBTRACK_Database
const SHEET_NAME = 'Job_Log';

// ============================================================
// รับ POST request จาก Web App
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;

    if (data.action === 'CHECK_STATUS') {
      // Web App ถามว่า พนักงานคนนี้มี Job ค้างอยู่ไหม?
      result = checkOpenJob(data.emp_id);

    } else if (data.action === 'CHECK_IN') {
      result = doCheckIn(data);

    } else if (data.action === 'CHECK_OUT') {
      result = doCheckOut(data);

    } else {
      result = { ok: false, message: 'Unknown action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// GET — รองรับ JSONP (มือถือ, QR Code, ทุกเครือข่าย)
// ============================================================
function doGet(e) {
  var params   = e.parameter;
  var callback = params.callback || '';
  var dataStr  = params.data    || '{}';

  var result;
  try {
    var data = JSON.parse(dataStr);
    if (data.action === 'CHECK_STATUS') {
      result = checkOpenJob(data.emp_id);
    } else if (data.action === 'CHECK_IN') {
      result = doCheckIn(data);
    } else if (data.action === 'CHECK_OUT') {
      result = doCheckOut(data);
    } else {
      result = { ok: true, status: 'JOBTRACK API v2 running ✅' };
    }
  } catch(err) {
    result = { ok: false, message: err.toString() };
  }

  var json = JSON.stringify(result);

  // ถ้ามี callback → ส่งกลับแบบ JSONP
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // ไม่มี callback → ส่ง JSON ปกติ
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// checkOpenJob — เช็คว่าพนักงานมี Job ค้างอยู่ไหม
// Return: { ok, hasOpen, openJob }
// ============================================================
function checkOpenJob(emp_id) {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();

  // วนหาแถวที่ emp_id ตรง และ checkout_time ว่าง (ยังไม่ Check Out)
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rowEmpId      = row[COL.EMP_ID];
    const rowCheckOut   = row[COL.TIME_OUT];
    const rowStatus     = row[COL.STATUS];

    if (rowEmpId === emp_id && rowStatus === 'Check In' && rowCheckOut === '') {
      return {
        ok: true,
        hasOpen: true,
        openJob: {
          row_index: i + 1,          // 1-based สำหรับ getRange
          job_id:    row[COL.JOB_ID],
          job_name:  row[COL.JOB_NAME],
          process:   row[COL.PROCESS],
          time_in:   row[COL.TIME_IN],
          date:      row[COL.DATE],
        }
      };
    }
  }

  return { ok: true, hasOpen: false };
}

// ============================================================
// doCheckIn — บันทึก Check In (สร้างแถวใหม่)
// ============================================================
function doCheckIn(data) {
  // เช็คก่อนว่ามี Job ค้างไหม
  const statusCheck = checkOpenJob(data.emp_id);
  if (statusCheck.hasOpen) {
    return {
      ok: false,
      blocked: true,
      message: 'OPEN_JOB_EXISTS',
      openJob: statusCheck.openJob
    };
  }

  const sheet = getSheet();
  ensureHeader(sheet);

  const now      = new Date();
  const timeStr  = formatTime(now);
  const dateStr  = formatDate(now);
  const tsStr    = now.toLocaleString('th-TH');

  const row = buildRow({
    timestamp: tsStr,
    date:      dateStr,
    time_in:   timeStr,
    time_out:  '',
    hours:     '',
    job_id:    data.job_id,
    job_name:  data.job_name,
    job_zone:  data.job_zone  || '',
    emp_id:    data.emp_id,
    emp_name:  data.emp_name,
    emp_dept:  data.emp_dept  || '',
    status:    'Check In',
    process:   data.process   || '',
    lang:      data.lang      || 'th',
    note:      data.note      || '',
  });

  sheet.appendRow(row);

  // สี Check In = เขียว
  const lastRow = sheet.getLastRow();
  colorStatusCell(sheet, lastRow, 'Check In');

  return { ok: true, action: 'CHECK_IN', time_in: timeStr, date: dateStr };
}

// ============================================================
// doCheckOut — หา Check In คู่ → คำนวณชั่วโมง → Update แถวเดิม
// ============================================================
function doCheckOut(data) {
  const sheet   = getSheet();
  const values  = sheet.getDataRange().getValues();
  const now     = new Date();
  const timeStr = formatTime(now);

  // หาแถว Check In ที่ยังไม่ Check Out ของ emp_id + job_id + process
  let targetRowIndex = -1;
  for (let i = values.length - 1; i >= 1; i--) {
    const row = values[i];
    if (
      row[COL.EMP_ID]  === data.emp_id  &&
      row[COL.JOB_ID]  === data.job_id  &&
      row[COL.PROCESS] === data.process  &&
      row[COL.STATUS]  === 'Check In'   &&
      row[COL.TIME_OUT] === ''
    ) {
      targetRowIndex = i + 1; // 1-based
      break;
    }
  }

  if (targetRowIndex === -1) {
    // ไม่เจอ Check In คู่ → หาแบบหลวมๆ (emp + job ไม่สนใจ process)
    for (let i = values.length - 1; i >= 1; i--) {
      const row = values[i];
      if (
        row[COL.EMP_ID]   === data.emp_id &&
        row[COL.JOB_ID]   === data.job_id &&
        row[COL.STATUS]   === 'Check In'  &&
        row[COL.TIME_OUT] === ''
      ) {
        targetRowIndex = i + 1;
        break;
      }
    }
  }

  if (targetRowIndex === -1) {
    return { ok: false, message: 'NO_CHECKIN_FOUND' };
  }

  // คำนวณชั่วโมง
  const timeInStr = values[targetRowIndex - 1][COL.TIME_IN];
  const hours     = calcHours(timeInStr, timeStr);

  // Update แถวเดิม: time_out, hours, status
  sheet.getRange(targetRowIndex, COL.TIME_OUT + 1).setValue(timeStr);
  sheet.getRange(targetRowIndex, COL.HOURS + 1).setValue(hours);
  sheet.getRange(targetRowIndex, COL.STATUS + 1).setValue('Check Out');

  // สี Check Out = แดง
  colorStatusCell(sheet, targetRowIndex, 'Check Out');

  // สี Hours cell = amber ถ้าเกิน 8 ชม.
  const hoursCell = sheet.getRange(targetRowIndex, COL.HOURS + 1);
  if (hours > 8) {
    hoursCell.setBackground('#fff3cd');
    hoursCell.setFontColor('#856404');
  } else {
    hoursCell.setBackground('#e8f5e9');
    hoursCell.setFontColor('#1b5e20');
  }

  return {
    ok: true,
    action: 'CHECK_OUT',
    time_in:  timeInStr,
    time_out: timeStr,
    hours:    hours,
  };
}

// ============================================================
// COLUMN INDEX MAP (0-based)
// ============================================================
const COL = {
  TIMESTAMP: 0,
  DATE:      1,
  TIME_IN:   2,
  TIME_OUT:  3,
  HOURS:     4,
  JOB_ID:    5,
  JOB_NAME:  6,
  JOB_ZONE:  7,
  EMP_ID:    8,
  EMP_NAME:  9,
  EMP_DEPT:  10,
  STATUS:    11,
  PROCESS:   12,
  LANG:      13,
  NOTE:      14,
};

// ============================================================
// buildRow — สร้าง array ตาม COL map
// ============================================================
function buildRow(d) {
  const row = new Array(15).fill('');
  row[COL.TIMESTAMP] = d.timestamp;
  row[COL.DATE]      = d.date;
  row[COL.TIME_IN]   = d.time_in;
  row[COL.TIME_OUT]  = d.time_out;
  row[COL.HOURS]     = d.hours;
  row[COL.JOB_ID]    = d.job_id;
  row[COL.JOB_NAME]  = d.job_name;
  row[COL.JOB_ZONE]  = d.job_zone;
  row[COL.EMP_ID]    = d.emp_id;
  row[COL.EMP_NAME]  = d.emp_name;
  row[COL.EMP_DEPT]  = d.emp_dept;
  row[COL.STATUS]    = d.status;
  row[COL.PROCESS]   = d.process;
  row[COL.LANG]      = d.lang;
  row[COL.NOTE]      = d.note;
  return row;
}

// ============================================================
// calcHours — รองรับทั้ง String "HH:MM" และ Date Object จาก Sheets
// ============================================================
function calcHours(timeIn, timeOut) {
  function toMin(t) {
    if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
    var s = String(t).trim();
    var parts = s.split(':');
    return parseInt(parts[0]) * 60 + parseInt(parts[1]);
  }
  var totalMin = toMin(timeOut) - toMin(timeIn);
  if (totalMin <= 0) return 0;
  return Math.round((totalMin / 60) * 100) / 100;
}

// ============================================================
// Helpers
// ============================================================
function formatTime(d) {
  return d.getHours().toString().padStart(2,'0') + ':' + d.getMinutes().toString().padStart(2,'0');
}
function formatDate(d) {
  return d.toLocaleDateString('th-TH');
}
function getSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    ensureHeader(sheet);
  }
  return sheet;
}
function ensureHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  const headers = [
    'Timestamp','วันที่','เวลาเข้า','เวลาออก','ชม.รวม',
    'Job ID','ชื่องาน','Zone',
    'รหัสพนักงาน','ชื่อพนักงาน','แผนก',
    'สถานะ','Process','ภาษา','หมายเหตุ'
  ];
  sheet.appendRow(headers);
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground('#0f1117');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setFontSize(11);
  sheet.setFrozenRows(1);
  // กำหนดความกว้าง
  const widths = [150,90,70,70,60,70,150,100,100,130,100,80,120,50,150];
  widths.forEach((w,i) => sheet.setColumnWidth(i+1, w));
}
function colorStatusCell(sheet, rowIndex, status) {
  const cell = sheet.getRange(rowIndex, COL.STATUS + 1);
  if (status === 'Check In') {
    cell.setBackground('#e8f5e9'); cell.setFontColor('#1b5e20');
  } else {
    cell.setBackground('#ffebee'); cell.setFontColor('#7f0000');
  }
  cell.setFontWeight('bold');
}

// ============================================================
// testCheckIn — ทดสอบ Check In
// ============================================================
function testCheckIn() {
  const fakeData = {
    action:   'CHECK_IN',
    emp_id:   'EMP001',
    emp_name: 'สมชาย รักงาน',
    emp_dept: 'ฝ่ายผลิต',
    job_id:   'T-0042',
    job_name: 'แท็งค์น้ำมัน #4',
    job_zone: 'ลาน B · Zone 3',
    process:  'Fabrication',
    lang:     'th',
  };
  const result = doCheckIn(fakeData);
  Logger.log('CHECK IN result: ' + JSON.stringify(result));
}

// ============================================================
// testCheckOut — ทดสอบ Check Out (รันหลัง testCheckIn)
// ============================================================
function testCheckOut() {
  const fakeData = {
    action:   'CHECK_OUT',
    emp_id:   'EMP001',
    job_id:   'T-0042',
    process:  'Fabrication',
  };
  const result = doCheckOut(fakeData);
  Logger.log('CHECK OUT result: ' + JSON.stringify(result));
}

// ============================================================
// testBlockedCheckIn — ทดสอบ Block ถ้ายังไม่ Check Out
// ============================================================
function testBlockedCheckIn() {
  // รัน testCheckIn ก่อน แล้วรันอันนี้ → ต้องได้ blocked: true
  const fakeData = {
    action:   'CHECK_IN',
    emp_id:   'EMP001',
    emp_name: 'สมชาย รักงาน',
    emp_dept: 'ฝ่ายผลิต',
    job_id:   'T-0043',
    job_name: 'แท็งค์น้ำมัน #5',
    job_zone: 'ลาน C',
    process:  'QC',
    lang:     'th',
  };
  const result = doCheckIn(fakeData);
  Logger.log('BLOCKED result: ' + JSON.stringify(result));
  // ต้องได้: { ok: false, blocked: true, message: "OPEN_JOB_EXISTS", openJob: {...} }
}


// ============================================================
// JOB LIST — Admin จัดการ Job ผ่าน Sheet Tab "Job_List"
// ============================================================
const JOB_LIST_SHEET = 'Job_List';

// สร้าง Tab Job_List พร้อม Header (รันครั้งแรกครั้งเดียว)
function setupJobListSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(JOB_LIST_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(JOB_LIST_SHEET);
  } else {
    sheet.clear();
  }

  const headers = ['ประเภทงาน','Job Code','Job Name / ลูกค้า','สถานะ','QR Link','วันที่เพิ่ม'];
  sheet.appendRow(headers);

  // Style Header
  const hRange = sheet.getRange(1,1,1,headers.length);
  hRange.setBackground('#0f1117');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setFontSize(11);
  sheet.setFrozenRows(1);

  // ความกว้าง
  [80,140,280,70,400,100].forEach((w,i)=>sheet.setColumnWidth(i+1,w));

  // ใส่ข้อมูลตัวอย่าง
  const sampleJobs = [
    ['JM','JM-68/0944','บริษัท ยูนิปาล์ม อินดัสทรี จำกัด','Active'],
    ['JM','JM-68/0945','บริษัท ยูนิปาล์ม อินดัสทรี จำกัด','Active'],
    ['JT','JT-68/0049','บริษัท ธาราลำพูนอีซูซุเซลล์ จำกัด','Active'],
    ['JT','JT-68/0052','คุณพิชญา สุทธิพรประสิทธิ์','Active'],
    ['JMS','JMS-69/0002','บจก.เอสวัน โลจิสติกส์','Active'],
    ['JP','JP-68/0003','บจก.เข็มเหล็ก','Active'],
    ['Other Job','ST-0001','Stand By','Active'],
    ['Other Job','ST-0003','ประชุม','Active'],
  ];

  const BASE_URL = 'https://stt-web.github.io/STT-JOBTRACK/job_checkin_app.html';
  const today = new Date().toLocaleDateString('th-TH');

  sampleJobs.forEach(j => {
    const jobCode = j[1];
    const jobName = j[2];
    const jobType = j[0];
    const qrLink  = BASE_URL +
      '?job='  + encodeURIComponent(jobCode) +
      '&name=' + encodeURIComponent(jobName) +
      '&type=' + encodeURIComponent(jobType) +
      '&cust=' + encodeURIComponent(jobName);
    sheet.appendRow([j[0], jobCode, jobName, j[3], qrLink, today]);
  });

  // สี Active = เขียว
  const dataRange = sheet.getRange(2, 4, sampleJobs.length, 1);
  dataRange.setBackground('#e8f5e9');
  dataRange.setFontColor('#1b5e20');
  dataRange.setFontWeight('bold');

  Logger.log('✅ Job_List Sheet พร้อมแล้ว — ' + sampleJobs.length + ' Jobs');
}

// เพิ่ม Job ใหม่ 1 รายการ (Admin เรียกจาก Apps Script หรือ Trigger)
function addNewJob(jobType, jobCode, jobName) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(JOB_LIST_SHEET);
  if (!sheet) { setupJobListSheet(); return; }

  const BASE_URL = 'https://stt-web.github.io/STT-JOBTRACK/job_checkin_app.html';
  const qrLink   = BASE_URL +
    '?job='  + encodeURIComponent(jobCode) +
    '&name=' + encodeURIComponent(jobName) +
    '&type=' + encodeURIComponent(jobType) +
    '&cust=' + encodeURIComponent(jobName);

  const today = new Date().toLocaleDateString('th-TH');
  sheet.appendRow([jobType, jobCode, jobName, 'Active', qrLink, today]);

  // สีแถวใหม่
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 4).setBackground('#e8f5e9').setFontColor('#1b5e20').setFontWeight('bold');

  Logger.log('✅ เพิ่ม Job: ' + jobCode + ' — ' + jobName);
  return qrLink;
}

// อัปเดตสถานะ Job (Active / Done)
function updateJobStatus(jobCode, status) {
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const sheet  = ss.getSheetByName(JOB_LIST_SHEET);
  const data   = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === jobCode) {
      const cell = sheet.getRange(i+1, 4);
      cell.setValue(status);
      if (status === 'Done') {
        cell.setBackground('#ffebee').setFontColor('#7f0000');
      } else {
        cell.setBackground('#e8f5e9').setFontColor('#1b5e20');
      }
      Logger.log('✅ อัปเดต ' + jobCode + ' → ' + status);
      return;
    }
  }
  Logger.log('❌ ไม่พบ Job: ' + jobCode);
}

// ── ทดสอบ ──
function testSetupJobList()  { setupJobListSheet(); }
function testAddJob()        { addNewJob('JM','JM-69/0999','บริษัท ทดสอบ จำกัด'); }
function testDoneJob()       { updateJobStatus('JM-69/0999','Done'); }
