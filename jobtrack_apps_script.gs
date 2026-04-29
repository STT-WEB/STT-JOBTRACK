// ============================================================
// JOBTRACK — Google Apps Script Backend v2
// Helper: pad number
function pad(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}
// Logic: 1 Check In = 1 Process = 1 Job
//        Check Out → Match กับ Check In → คำนวณชั่วโมง
//        บล็อก Check In ใหม่ถ้ายัง Check Out งานเก่าไม่ครบ
// ============================================================

const SHEET_ID   = '1MYWORYN3sOjov3Gxv3UqCV1jRSxgxwGi1tRomFUGSr0';   // JOBTRACK_Database
// ── รอบเงินเดือน 26-25 ──
function getPayrollSheetName() {
  var today = new Date();
  var day   = today.getDate();
  var month = today.getMonth() + 1;
  var year  = today.getFullYear() + 543;

  // วันที่ 26-31 → นับเป็นรอบเดือนถัดไป
  if (day >= 26) {
    month = month + 1;
    if (month > 12) { month = 1; year++; }
  }

  return 'Job_Log_' + year + '_' + pad(month, 2);
}

const SHEET_NAME = getPayrollSheetName(); // Dynamic ตามรอบเงินเดือน

// ============================================================
// รับ POST request จาก Web App
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;

    if (data.action === 'CHECK_STATUS') {
      result = checkOpenJob(data.emp_id);
    } else if (data.action === 'CHECK_IN') {
      result = doCheckIn(data);
    } else if (data.action === 'CHECK_OUT') {
      result = doCheckOut(data);
    } else if (data.action === 'UPLOAD_PHOTO') {
      result = uploadPhoto(data);
    } else if (data.action === 'REGISTER_EMPLOYEE') {
      result = registerEmployeeByCode(data.lineId, data.empId, data.lineName || '');
    } else {
      result = { ok: false, message: 'Unknown action' };
    }

    // เพิ่ม CORS headers
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
    } else if (data.action === 'REGISTER_EMPLOYEE') {
      result = registerEmployeeByCode(data.lineId, data.empId, data.lineName || '');
    } else if (data.action === 'UPLOAD_PHOTO') {
      result = uploadPhoto(data);
    } else if (data.action === 'SAVE_PHOTO_URLS') {
      result = savePhotoUrls(data);
    } else if (data.action === 'LINE_CALLBACK') {
      result = handleLineCallback(data);
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
  // ค้นหาใน Tab รอบเงินเดือนปัจจุบัน
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    if (row[COL.EMP_ID] === emp_id &&
        row[COL.STATUS]  === 'Check In' &&
        row[COL.TIME_OUT] === '') {
      return {
        ok: true,
        hasOpen: true,
        openJob: {
          row_index:  i + 1,
          job_id:     row[COL.JOB_ID],
          job_name:   row[COL.JOB_NAME],
          process:    row[COL.PROCESS],
          time_in:    row[COL.TIME_IN],
          date:       row[COL.DATE],
          sheet_name: sheet.getName(), // บันทึกชื่อ Tab ไว้ด้วย
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

  // ดึงประเภทวันทำงาน
  const dayType = getDayType(dateStr);

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
    day_type:  dayType,
    hour_type: '',
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
  // ต้องหาใน Tab ที่ Check In อยู่ (อาจเป็น Tab เดือนก่อนถ้า Check In ข้ามรอบ)
  var sheet  = getSheet();
  var values = sheet.getDataRange().getValues();

  // ถ้าไม่เจอใน Tab ปัจจุบัน → ลองหาใน Tab เดือนก่อน
  var found = false;
  for (var k = values.length - 1; k >= 1; k--) {
    if (values[k][COL.EMP_ID] === data.emp_id && values[k][COL.STATUS] === 'Check In') {
      found = true; break;
    }
  }

  if (!found) {
    // ลองหาใน Tab เดือนก่อน
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var allSheets = ss.getSheets();
    var currentName = sheet.getName();

    for (var s = 0; s < allSheets.length; s++) {
      var sName = allSheets[s].getName();
      if (sName.indexOf('Job_Log_') === 0 && sName !== currentName) {
        var testData = allSheets[s].getDataRange().getValues();
        for (var t = testData.length - 1; t >= 1; t--) {
          if (testData[t][COL.EMP_ID] === data.emp_id && testData[t][COL.STATUS] === 'Check In') {
            sheet  = allSheets[s];
            values = testData;
            Logger.log('พบ Check In ใน Tab: ' + sName);
            break;
          }
        }
      }
      if (sheet.getName() !== currentName) break;
    }
  }
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

  // คำนวณประเภทชั่วโมง
  const dayTypeVal  = values[targetRowIndex - 1][COL.DAY_TYPE] || getDayType(values[targetRowIndex - 1][COL.DATE]);
  const hourTypeVal = getHourType(timeInStr, timeStr, dayTypeVal);

  // Update แถวเดิม: time_out, hours, status, day_type, hour_type
  sheet.getRange(targetRowIndex, COL.TIME_OUT + 1).setValue(timeStr);
  sheet.getRange(targetRowIndex, COL.HOURS + 1).setValue(hours);
  sheet.getRange(targetRowIndex, COL.STATUS + 1).setValue('Check Out');
  sheet.getRange(targetRowIndex, COL.DAY_TYPE + 1).setValue(dayTypeVal);
  sheet.getRange(targetRowIndex, COL.HOUR_TYPE + 1).setValue(hourTypeVal);

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
  DAY_TYPE:  15,   // ประเภทวันทำงาน
  HOUR_TYPE: 16,   // ประเภทชั่วโมง
};

// ============================================================
// buildRow — สร้าง array ตาม COL map
// ============================================================
function buildRow(d) {
  const row = new Array(17).fill('');
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
  row[COL.DAY_TYPE]  = d.day_type  || '';
  row[COL.HOUR_TYPE] = d.hour_type || '';
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
  return pad(d.getHours(),2) + ':' + pad(d.getMinutes(),2);
}
function formatDate(d) {
  return d.toLocaleDateString('th-TH');
}
function getSheet() {
  var sheetName = getPayrollSheetName();
  var ss        = SpreadsheetApp.openById(SHEET_ID);
  var sheet     = ss.getSheetByName(sheetName);

  if (!sheet) {
    // สร้าง Tab ใหม่สำหรับรอบเงินเดือนนี้
    sheet = ss.insertSheet(sheetName);
    ensureHeader(sheet);

    // ย้าย Tab ให้อยู่ในลำดับที่ถูกต้อง (ด้านหน้าสุด)
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);

    Logger.log('สร้าง Tab ใหม่: ' + sheetName);
  }
  return sheet;
}
function ensureHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  const headers = [
    'Timestamp','วันที่','เวลาเข้า','เวลาออก','ชม.รวม',
    'Job ID','ชื่องาน','Zone',
    'รหัสพนักงาน','ชื่อพนักงาน','แผนก',
    'สถานะ','Process','ภาษา','หมายเหตุ',
    'ประเภทวันทำงาน','ประเภทชั่วโมง'
  ];
  sheet.appendRow(headers);
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground('#0f1117');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setFontSize(11);
  sheet.setFrozenRows(1);
  // กำหนดความกว้าง
  const widths = [150,90,70,70,60,70,150,100,100,130,100,80,120,50,150,150,180];
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

  // ติดตั้ง Trigger onEdit อัตโนมัติ
  installTrigger();

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

// ============================================================
// AUTO QR GENERATOR — เมื่อ Admin กรอก A B C ใน Job_List
// QR Link จะ Generate อัตโนมัติใน Column E ทันที
// ============================================================
function onJobListEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== JOB_LIST_SHEET) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return; // ข้าม Header

  // ถ้าแก้ A, B, หรือ C → Re-generate QR Link
  if (col >= 1 && col <= 3) {
    const rowData = sheet.getRange(row, 1, 1, 3).getValues()[0];
    const jobType = rowData[0];
    const jobCode = rowData[1];
    const jobName = rowData[2];

    if (!jobCode || !jobName) return; // ยังกรอกไม่ครบ

    const BASE_URL = 'https://stt-web.github.io/STT-JOBTRACK/job_checkin_app.html';
    const qrLink   = BASE_URL +
      '?job='  + encodeURIComponent(jobCode) +
      '&name=' + encodeURIComponent(jobName) +
      '&type=' + encodeURIComponent(jobType || 'JM') +
      '&cust=' + encodeURIComponent(jobName);

    // ใส่ QR Link ใน Column E
    sheet.getRange(row, 5).setValue(qrLink);

    // ใส่วันที่ใน Column F ถ้ายังว่าง
    if (!sheet.getRange(row, 6).getValue()) {
      sheet.getRange(row, 6).setValue(new Date().toLocaleDateString('th-TH'));
    }

    // ใส่ Active ใน Column D ถ้ายังว่าง
    const statusCell = sheet.getRange(row, 4);
    if (!statusCell.getValue()) {
      statusCell.setValue('Active');
      statusCell.setBackground('#e8f5e9');
      statusCell.setFontColor('#1b5e20');
      statusCell.setFontWeight('bold');
    }

    // Highlight แถวใหม่
    sheet.getRange(row, 5).setFontColor('#1565c0').setFontStyle('italic');
  }
}

// ติดตั้ง Trigger (รันแค่ครั้งเดียว)
function installTrigger() {
  // ลบ Trigger เก่าก่อน
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onJobListEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // สร้างใหม่
  ScriptApp.newTrigger('onJobListEdit')
    .forSpreadsheet(SHEET_ID)
    .onEdit()
    .create();
  Logger.log('✅ Trigger ติดตั้งแล้ว — Auto QR พร้อมใช้งาน');
}

// ── ทดสอบ ──
function testInstallTrigger() { installTrigger(); }

// ============================================================
// QR PRINT SHEET — สร้าง Sheet แสดง QR Code รูปภาพพร้อมพิมพ์
// ใช้ Google Charts API สร้าง QR ฟรี 100%
// ============================================================
const QR_SHEET_NAME = 'QR_Print';
const QR_SIZE = 150; // pixel

function generateQRPrintSheet() {
  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const jobSheet = ss.getSheetByName(JOB_LIST_SHEET);
  if (!jobSheet) { Logger.log('ไม่พบ Job_List'); return; }

  // ดึง Job ทั้งหมด (ไม่กรอง Status — เอาทุก Job ที่มี Link)
  const jobs = jobSheet.getDataRange().getValues();
  const allJobs = [];
  for (let i = 1; i < jobs.length; i++) {
    const row = jobs[i];
    if (row[1] && row[4]) { // มี Job Code และ QR Link
      allJobs.push({
        type: row[0] || 'JM',
        code: row[1],
        name: row[2],
        status: row[3] || 'Active',
        link: row[4],
      });
    }
  }

  if (allJobs.length === 0) {
    SpreadsheetApp.getUi().alert('ไม่พบ Job ที่มี QR Link\nกรุณาตรวจสอบ Tab Job_List');
    return;
  }

  // สร้าง HTML พร้อมพิมพ์
  const BASE_QR = 'https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=';
  const typeColors = {
    'JM':'#1565c0','JT':'#1b5e20','JMS':'#e65100',
    'JP':'#6a1b9a','Other Job':'#424242','JMC':'#00695c',
    'JMI':'#4527a0','JMT':'#558b2f','JMS':'#bf360c'
  };
  const typeBg = {
    'JM':'#e3f2fd','JT':'#e8f5e9','JMS':'#fff3e0',
    'JP':'#f3e5f5','Other Job':'#f5f5f5','JMC':'#e0f2f1',
    'JMI':'#ede7f6','JMT':'#f9fbe7','JMS':'#fbe9e7'
  };

  let cards = '';
  allJobs.forEach(j => {
    const qrUrl  = BASE_QR + encodeURIComponent(j.link);
    const color  = typeColors[j.type]  || '#333';
    const bg     = typeBg[j.type]      || '#f5f5f5';
    const short  = j.name.length > 35 ? j.name.substring(0,33)+'...' : j.name;
    const isDone = j.status === 'Done';
    cards += '<div class="card' + (isDone?' done':'') + '">' +
      '<div class="type-badge" style="color:' + color + ';background:' + bg + '">[' + j.type + ']</div>' +
      '<img src="' + qrUrl + '" width="160" height="160" alt="QR"/>' +
      '<div class="job-code">' + j.code + '</div>' +
      '<div class="job-name">' + short + '</div>' +
      (isDone ? '<div class="done-badge">✓ DONE</div>' : '') +
      '</div>';
  });

  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<title>JOBTRACK QR Codes</title>' +
    '<style>' +
    'body{font-family:sans-serif;background:#f0f0f0;padding:16px;margin:0}' +
    'h1{text-align:center;color:#0f1117;font-size:16px;margin-bottom:16px}' +
    '.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;max-width:900px;margin:0 auto}' +
    '.card{background:#fff;border-radius:12px;padding:14px;text-align:center;border:1px solid #ddd;break-inside:avoid}' +
    '.card.done{opacity:0.45;background:#f5f5f5}' +
    '.type-badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;margin-bottom:8px}' +
    '.card img{display:block;margin:0 auto 8px;border:1px solid #eee;border-radius:6px}' +
    '.job-code{font-size:13px;font-weight:700;color:#111;margin-bottom:4px}' +
    '.job-name{font-size:10px;color:#666;line-height:1.4}' +
    '.done-badge{margin-top:6px;font-size:10px;color:#999;font-weight:600}' +
    '@media print{body{background:#fff;padding:0}h1{margin:8px 0}.grid{gap:8px}}' +
    '</style></head><body>' +
    '<h1>JOBTRACK — QR Code พร้อมพิมพ์ | ' + allJobs.length + ' Jobs | สร้างเมื่อ: ' + new Date().toLocaleString('th-TH') + '</h1>' +
    '<div class="grid">' + cards + '</div>' +
    '<script>window.onload=function(){window.print();}<\/script>' +
    '</body></html>';

  // บันทึก HTML ลง Drive แล้วเปิด
  const folder = DriveApp.getRootFolder();
  const fileName = 'JOBTRACK_QR_' + Utilities.formatDate(new Date(),'Asia/Bangkok','yyyyMMdd_HHmm') + '.html';

  // ลบไฟล์เก่า
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const file = folder.createFile(fileName, html, MimeType.HTML);
  const url  = file.getUrl();

  SpreadsheetApp.getUi().alert(
    'QR Code พร้อมแล้ว! ' + allJobs.length + ' Jobs\n\n' +
    'กด OK แล้วเปิด Link นี้ใน Browser:\n' + url + '\n\n' +
    '(ไฟล์อยู่ใน Google Drive ของเบียร์)'
  );
  Logger.log('URL: ' + url);
}

// ── ทดสอบ ──
function testGenerateQR() {
  // เติม QR Link ให้ทุกแถวที่ยังไม่มีก่อน แล้วค่อย Generate
  fillMissingQRLinks();
  generateQRPrintSheet();
}

// เติม QR Link อัตโนมัติทุกแถวที่ยังว่าง
function fillMissingQRLinks() {
  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const sheet    = ss.getSheetByName(JOB_LIST_SHEET);
  if (!sheet) return;
  const BASE_URL = 'https://stt-web.github.io/STT-JOBTRACK/job_checkin_app.html';
  const data     = sheet.getDataRange().getValues();
  let   filled   = 0;

  for (let i = 1; i < data.length; i++) {
    const jobType = data[i][0];
    const jobCode = data[i][1];
    const jobName = data[i][2];
    const hasLink = data[i][4];

    if (!jobCode || !jobName) continue; // ข้ามแถวว่าง
    if (hasLink) continue;              // มี Link แล้ว ข้าม

    // สร้าง Link
    const qrLink = BASE_URL +
      '?job='  + encodeURIComponent(jobCode) +
      '&name=' + encodeURIComponent(jobName) +
      '&type=' + encodeURIComponent(jobType || 'JM') +
      '&cust=' + encodeURIComponent(jobName);

    sheet.getRange(i+1, 5).setValue(qrLink);

    // ใส่ Active + วันที่ถ้าว่าง
    if (!data[i][3]) {
      const statusCell = sheet.getRange(i+1, 4);
      statusCell.setValue('Active');
      statusCell.setBackground('#e8f5e9').setFontColor('#1b5e20').setFontWeight('bold');
    }
    if (!data[i][5]) {
      sheet.getRange(i+1, 6).setValue(new Date().toLocaleDateString('th-TH'));
    }
    filled++;
  }
  Logger.log('เติม QR Link: ' + filled + ' แถว');
}

// สร้างไฟล์ HTML QR โดยไม่ต้องใช้ Drive API
function generateQRHtml() {
  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const jobSheet = ss.getSheetByName(JOB_LIST_SHEET);
  if (!jobSheet) { Logger.log('ไม่พบ Job_List'); return; }

  fillMissingQRLinks();
  SpreadsheetApp.flush();

  const jobs    = jobSheet.getDataRange().getValues();
  const allJobs = [];
  for (let i = 1; i < jobs.length; i++) {
    const row = jobs[i];
    if (row[1] && row[4]) {
      allJobs.push({ type:row[0]||'JM', code:row[1], name:row[2], status:row[3]||'Active', link:row[4] });
    }
  }

  if (allJobs.length === 0) {
    SpreadsheetApp.getUi().alert('ไม่พบ Job ใน Job_List กรุณาเพิ่มข้อมูลก่อน');
    return;
  }

  const BASE_QR = 'https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=';
  const colors  = {JM:'#1565c0',JT:'#1b5e20',JMS:'#bf360c',JP:'#6a1b9a','Other Job':'#424242',JMC:'#00695c',JMI:'#4527a0',JMT:'#558b2f'};
  const bgs     = {JM:'#e3f2fd',JT:'#e8f5e9',JMS:'#fbe9e7',JP:'#f3e5f5','Other Job':'#f5f5f5',JMC:'#e0f2f1',JMI:'#ede7f6',JMT:'#f9fbe7'};

  let cards = '';
  allJobs.forEach(function(j) {
    var qrUrl = BASE_QR + encodeURIComponent(j.link);
    var col   = colors[j.type] || '#333';
    var bg    = bgs[j.type]    || '#f5f5f5';
    var short = j.name.length > 32 ? j.name.substring(0,30)+'...' : j.name;
    cards += '<div class="card"><div class="badge" style="color:' + col + ';background:' + bg + '">' + j.type + '</div>' +
             '<img src="' + qrUrl + '"><div class="code">' + j.code + '</div>' +
             '<div class="name">' + short + '</div></div>';
  });

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>JOBTRACK QR</title>' +
    '<style>*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:sans-serif;padding:12px;background:#fff}' +
    'h2{text-align:center;font-size:13px;color:#333;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid #ddd}' +
    '.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}' +
    '.card{border:1px solid #ddd;border-radius:10px;padding:10px;text-align:center;break-inside:avoid}' +
    '.badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:10px;font-weight:700;margin-bottom:6px}' +
    'img{width:150px;height:150px;display:block;margin:0 auto 6px;border-radius:4px}' +
    '.code{font-size:12px;font-weight:700;color:#111;margin-bottom:3px}' +
    '.name{font-size:9px;color:#666;line-height:1.4}' +
    '@media print{body{padding:4px}.grid{gap:6px}img{width:130px;height:130px}}' +
    '</style></head><body>' +
    '<h2>JOBTRACK QR Code | ' + allJobs.length + ' Jobs | ' + new Date().toLocaleDateString('th-TH') + '</h2>' +
    '<div class="grid">' + cards + '</div></body></html>';

  // บันทึกใน Google Drive
  var folder   = DriveApp.getRootFolder();
  var fname    = 'JOBTRACK_QR.html';
  var existing = folder.getFilesByName(fname);
  while (existing.hasNext()) existing.next().setTrashed(true);
  var file = folder.createFile(fname, html, MimeType.HTML);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
  Logger.log('URL: ' + url);
  SpreadsheetApp.getUi().alert('QR Code พร้อมแล้ว ' + allJobs.length + ' Jobs\n\nเปิดไฟล์ใน Google Drive:\n' + url);
}

// ============================================================
// LINE LOGIN — รับ code แล้วแลก Token + ดึง Profile
// ============================================================
const LINE_CHANNEL_ID     = '2009924671';
const LINE_CHANNEL_SECRET = '94135f1ad6adfcdb127db0bb89fa5b53'; // ← ใส่ Channel Secret จาก LINE Developers
const EMP_LIST_SHEET      = 'Employee_List';

function handleLineCallback(data) {
  try {
    // แลก code เป็น Access Token
    var tokenRes = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: {
        grant_type:    'authorization_code',
        code:           data.code,
        redirect_uri:   data.redirect,
        client_id:      LINE_CHANNEL_ID,
        client_secret:  LINE_CHANNEL_SECRET,
      }
    });
    var token = JSON.parse(tokenRes.getContentText());
    if (!token.access_token) return { ok: false, message: 'ไม่สามารถแลก Token ได้' };

    // ดึง Profile
    var profileRes = UrlFetchApp.fetch('https://api.line.me/v2/profile', {
      headers: { 'Authorization': 'Bearer ' + token.access_token }
    });
    var profile = JSON.parse(profileRes.getContentText());

    // หาข้อมูลพนักงานจาก Employee_List
    var empData = getEmployeeByLineId(profile.userId);

    // ถ้าไม่พบ LINE ID ใน Employee_List
    // ไม่สร้างแถวใหม่ — ให้ App บังคับกรอกรหัสพนักงานแทน
    if (!empData) {
      return {
        ok:          true,
        profile:     profile,
        empId:       '',
        empName:     '',
        empDept:     '',
        empRole:     '',
        needRegister: true, // แจ้ง App ว่าต้องลงทะเบียนก่อน
      };
    }

    return {
      ok:      true,
      profile: profile,
      empId:   empData.empId   || '',
      empName: empData.empName || profile.displayName,
      empDept: empData.empDept || '',
      empRole: empData.empRole || '',
      needRegister: false,
    };
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}

// หาพนักงานจาก LINE ID
function getEmployeeByLineId(lineId) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(EMP_LIST_SHEET);
  if (!sheet) return null;
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === lineId) { // คอลัมน์ D = LINE ID
      return { empId: data[i][0], empName: data[i][1], empDept: data[i][2], empRole: data[i][4] };
    }
  }
  return null;
}

// registerNewEmployee — ไม่ใช้แล้ว ระบบไม่สร้างแถวอัตโนมัติ
// พนักงานต้องกรอกรหัสพนักงานเองผ่านหน้า Register
function registerNewEmployee(profile) {
  Logger.log('registerNewEmployee ถูกเรียกแต่ไม่ทำงานแล้ว: ' + profile.userId);
  return { empId: '', empName: profile.displayName, empDept: '', empRole: '' };
}

// ── เพิ่ม handleLineCallback ใน doGet ──

// ============================================================
// ทดสอบ UrlFetch Permission — รันครั้งเดียวเพื่อให้สิทธิ์
// ============================================================
function testUrlFetch() {
  try {
    var res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: { grant_type: 'test' },
      muteHttpExceptions: true
    });
    Logger.log('UrlFetch OK: ' + res.getResponseCode());
  } catch(e) {
    Logger.log('Error: ' + e.toString());
  }
}

// ============================================================
// IMPORT พนักงานจาก DATA EMPLOYEE → Employee_List
// ============================================================
const DATA_EMP_SHEET = 'DATA EMPLOYEE';

function importEmployeeData() {
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var srcSheet = ss.getSheetByName(DATA_EMP_SHEET);
  var empSheet = ss.getSheetByName(EMP_LIST_SHEET);

  if (!srcSheet) { Logger.log('ไม่พบ Sheet: ' + DATA_EMP_SHEET); return; }

  // สร้าง Employee_List ถ้ายังไม่มี
  if (!empSheet) {
    empSheet = ss.insertSheet(EMP_LIST_SHEET);
  } else {
    // เคลียร์ข้อมูลเก่า แต่เก็บ LINE ID ไว้
    // บันทึก LINE ID เดิมก่อน
  }

  // อ่านข้อมูลจาก DATA EMPLOYEE
  var srcData = srcSheet.getDataRange().getValues();

  // หา Header row (แถวที่มี "รหัสพนักงาน")
  var headerRow = -1;
  for (var i = 0; i < srcData.length; i++) {
    if (String(srcData[i][3]).indexOf('รหัสพนักงาน') >= 0 ||
        String(srcData[i][2]).indexOf('รหัสพนักงาน') >= 0) {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) headerRow = 2; // แถว 3 (index 2) คือ Header

  // เก็บ LINE ID เดิมไว้ก่อน (key = รหัสพนักงาน)
  var existingLineIds = {};
  if (empSheet.getLastRow() > 1) {
    var empData = empSheet.getDataRange().getValues();
    for (var e = 1; e < empData.length; e++) {
      var empId   = String(empData[e][0]).trim();
      var lineId  = String(empData[e][3]).trim();
      if (empId && lineId) existingLineIds[empId] = lineId;
    }
  }

  // Clear Employee_List
  empSheet.clear();
  empSheet.clearFormats();

  // สร้าง Header ใหม่
  var headers = ['รหัสพนักงาน','ชื่อพนักงาน','แผนก','LINE ID','ตำแหน่ง','ประเภทพนักงาน','Direct/Indirect','LINE Name','วันที่ลงทะเบียน'];
  empSheet.appendRow(headers);
  var hRange = empSheet.getRange(1,1,1,headers.length);
  hRange.setBackground('#0f1117').setFontColor('#00e187').setFontWeight('bold').setFontSize(11);
  empSheet.setFrozenRows(1);
  [100,180,150,150,160,140,120,150,110].forEach(function(w,i){ empSheet.setColumnWidth(i+1,w); });

  // Import ข้อมูลพนักงาน
  var imported = 0;
  var rows = [];

  for (var r = headerRow + 1; r < srcData.length; r++) {
    var row      = srcData[r];
    var deptCode = String(row[1]).trim(); // B = รหัสแผนก
    var dept     = String(row[2]).trim(); // C = แผนก
    var empId    = String(row[3]).trim(); // D = รหัสพนักงาน
    var empName  = String(row[4]).trim(); // E = ชื่อพนักงาน
    var position = String(row[7]).trim(); // H = ตำแหน่ง
    var empType  = String(row[8]).trim(); // I = ประเภทพนักงาน
    var direct   = String(row[10]).trim(); // K = Direct/Indirect

    // ข้ามแถวว่างหรือไม่มีรหัสพนักงาน
    if (!empId || empId === '' || isNaN(empId)) continue;
    if (!empName || empName === '') continue;

    // ดึง LINE ID เดิมถ้ามี
    var lineId = existingLineIds[empId] || '';

    rows.push([empId, empName, dept, lineId, position, empType, direct, '', '']);
    imported++;
  }

  // เขียนข้อมูลทีเดียว
  if (rows.length > 0) {
    empSheet.getRange(2, 1, rows.length, 9).setValues(rows);

    // สี Alternate rows
    for (var i = 0; i < rows.length; i++) {
      if (i % 2 === 0) {
        empSheet.getRange(i+2, 1, 1, 9).setBackground('#f8f9fa');
      }
    }

    // Highlight แถวที่มี LINE ID แล้ว (สีเขียว)
    for (var j = 0; j < rows.length; j++) {
      if (rows[j][3]) { // มี LINE ID
        empSheet.getRange(j+2, 4).setBackground('#e8f5e9').setFontColor('#1b5e20').setFontWeight('bold');
      }
    }
  }

  Logger.log('Import สำเร็จ: ' + imported + ' คน | LINE ID ที่รักษาไว้: ' + Object.keys(existingLineIds).length + ' คน');
  SpreadsheetApp.getUi().alert('Import สำเร็จ! ' + imported + ' คน\nLINE ID ที่รักษาไว้: ' + Object.keys(existingLineIds).length + ' คน');
}

// ── ทดสอบ ──
function testImportEmployee() { importEmployeeData(); }

// ============================================================
// REGISTER — ผูก LINE ID กับรหัสพนักงาน (พนักงานกรอกเอง)
// ============================================================
function registerEmployeeByCode(lineId, empId, lineName) {
  Logger.log('Register: lineId=' + lineId + ' empId=' + empId + ' lineName=' + lineName);
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var empSheet = ss.getSheetByName(EMP_LIST_SHEET);
  if (!empSheet) return { ok: false, message: 'ไม่พบ Employee List กรุณาติดต่อ Admin' };

  var data  = empSheet.getDataRange().getValues();
  var today = new Date().toLocaleDateString('th-TH');

  // ลบแถวที่ระบบสร้างอัตโนมัติ (EMP prefix) ที่มี LINE ID นี้ก่อน
  for (var j = data.length - 1; j >= 1; j--) {
    var rowId  = String(data[j][0]).trim();
    var rowLid = String(data[j][3]).trim();
    if (rowLid === lineId && rowId.indexOf('EMP') === 0) {
      empSheet.deleteRow(j + 1);
      Logger.log('ลบแถว Auto: ' + rowId);
      // โหลดข้อมูลใหม่หลังลบ
      data = empSheet.getDataRange().getValues();
      break;
    }
  }

  // หารหัสพนักงานใน Employee_List
  for (var i = 1; i < data.length; i++) {
    var rowEmpId = String(data[i][0]).trim();
    if (rowEmpId === String(empId).trim()) {
      // เจอแล้ว — ผูก LINE ID + เติม LINE Name + วันที่
      // lineName มาจาก parameter ที่ส่งมา ไม่ใช่จาก Sheet
      // บันทึก LINE ID, LINE Name, วันที่ลงทะเบียน
      empSheet.getRange(i+1, 4).setValue(lineId);
      empSheet.getRange(i+1, 4).setBackground('#e8f5e9').setFontColor('#1b5e20').setFontWeight('bold');

      // LINE Name — Column H (8)
      var nameCell = empSheet.getRange(i+1, 8);
      nameCell.setValue(String(lineName || ''));
      nameCell.setFontColor('#1565c0');

      // วันที่ลงทะเบียน — Column I (9)
      empSheet.getRange(i+1, 9).setValue(today);

      Logger.log('บันทึก LINE Name: ' + lineName + ' ที่แถว ' + (i+1));

      Logger.log('ผูก LINE ID สำเร็จ: ' + rowEmpId + ' → ' + lineId);
      return {
        ok:      true,
        empId:   String(data[i][0]).trim(),
        empName: String(data[i][1]).trim(),
        empDept: String(data[i][2]).trim(),
        empRole: String(data[i][4]).trim(),
      };
    }
  }

  // ไม่พบรหัสพนักงาน
  return { ok: false, message: 'ไม่พบรหัสพนักงาน ' + empId + ' กรุณาตรวจสอบใหม่' };
}

// ============================================================
// WORK DAY CALCULATOR
// ดึงประเภทวันจาก Tab "ประเภทวันทำงาน" และคำนวณประเภทชั่วโมง
// ============================================================
const CALENDAR_SHEET = 'ประเภทวันทำงาน';
const OT_START_HOUR  = 17; // OT เริ่ม 17:00

// ดึงประเภทวันทำงานจาก Sheet
function getDayType(dateStr) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(CALENDAR_SHEET);
    if (!sheet) return 'วันทำงานปกติ'; // default

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0]; // Column A = วันที่
      if (!rowDate) continue;

      // แปลงเป็น Date object
      var d = new Date(rowDate);
      var dStr = d.toLocaleDateString('th-TH');

      if (dStr === dateStr) {
        return String(data[i][2]).trim(); // Column C = ประเภทวันทำงาน
      }
    }

    // ถ้าหาไม่เจอ → เช็คจากวันในสัปดาห์
    var checkDate = new Date(dateStr.split('/').reverse().join('-'));
    var dow = checkDate.getDay();
    if (dow === 0) return 'วันหยุด'; // อาทิตย์
    if (dow === 6) return 'วันหยุด'; // เสาร์
    return 'วันทำงานปกติ';

  } catch(e) {
    Logger.log('getDayType error: ' + e);
    return 'วันทำงานปกติ';
  }
}

// คำนวณประเภทชั่วโมงจากเวลาเข้า-ออก และประเภทวัน
function getHourType(timeIn, timeOut, dayType) {
  try {
    // แปลงเวลาเป็นนาที
    function toMin(t) {
      if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
      var s = String(t).trim();
      var parts = s.split(':');
      return parseInt(parts[0]) * 60 + parseInt(parts[1]);
    }

    var inMin  = toMin(timeIn);
    var outMin = toMin(timeOut);
    var otMin  = OT_START_HOUR * 60; // 17:00 = 1020 นาที

    var isHoliday = dayType.indexOf('หยุด') >= 0 ||
                    dayType.indexOf('นักขัตฤกษ์') >= 0;

    // กรณี Check Out ข้ามวัน (เช่น ออก 00:30 = 30 นาที < เข้า)
    if (outMin < inMin) outMin += 24 * 60;

    if (isHoliday) {
      // วันหยุด/นักขัตฤกษ์
      if (outMin <= otMin) {
        return 'ชั่วโมงวันหยุด/นักขัตฤกษ์';
      } else if (inMin >= otMin) {
        return 'โอทีวันหยุด/นักขัตฤกษ์ x3';
      } else {
        // ช่วงคาบเกี่ยว → ดูว่าเวลาส่วนใหญ่อยู่ที่ไหน
        return 'ชั่วโมงวันหยุด/นักขัตฤกษ์';
      }
    } else {
      // วันทำงานปกติ
      if (outMin <= otMin) {
        return 'ชั่วโมงวันทำงานปกติ';
      } else if (inMin >= otMin) {
        return 'โอทีวันทำงานปกติ x1.5';
      } else {
        // คาบเกี่ยว เช่น เข้า 13:00 ออก 19:30
        return 'ชั่วโมงวันทำงานปกติ';
      }
    }
  } catch(e) {
    Logger.log('getHourType error: ' + e);
    return '-';
  }
}

// ============================================================
// ทดสอบ
// ============================================================
function testDayType() {
  var today = new Date().toLocaleDateString('th-TH');
  Logger.log('วันนี้ (' + today + '): ' + getDayType(today));
  Logger.log('ประเภทชั่วโมง 08:00-12:00 วันปกติ: ' + getHourType('08:00','12:00','วันทำงานปกติ'));
  Logger.log('ประเภทชั่วโมง 17:00-19:30 วันปกติ: ' + getHourType('17:00','19:30','วันทำงานปกติ'));
  Logger.log('ประเภทชั่วโมง 08:00-17:00 วันหยุด: ' + getHourType('08:00','17:00','วันหยุด'));
  Logger.log('ประเภทชั่วโมง 17:00-22:00 วันหยุดนักขัตฤกษ์: ' + getHourType('17:00','22:00','วันหยุดนักขัตฤกษ์'));
}

// ============================================================
// GOOGLE DRIVE PHOTO UPLOAD
// โครงสร้าง: JOBTRACK_Photos / Job / วันที่ / Process / รูป
// ============================================================
const DRIVE_ROOT_FOLDER = 'JOBTRACK_Photos';

// หรือสร้าง Folder ถ้ายังไม่มี
function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

// Upload รูปไป Drive
function uploadPhoto(data) {
  try {
    var jobId    = String(data.job_id   || 'UNKNOWN').replace(/\//g, '_');
    var empId    = String(data.emp_id   || '');
    var empName  = String(data.emp_name || '');
    var process  = String(data.process  || 'General').replace(/\//g, '-');
    var dateStr  = String(data.date     || new Date().toLocaleDateString('th-TH')).replace(/\//g, '-');
    var timeStr  = String(data.time     || formatTime(new Date())).replace(/:/g, '-');
    var base64   = data.photo_base64;

    if (!base64) return { ok: false, message: 'ไม่มีรูปภาพ' };

    // แยก base64 data URL
    var parts     = base64.split(',');
    var mimeMatch = parts[0].match(/data:([^;]+)/);
    var mime      = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    var ext       = mime.split('/')[1] || 'jpg';
    var imgData   = parts[1] || parts[0];

    // สร้าง Folder structure
    var root        = DriveApp.getRootFolder();
    var rootFolder  = getOrCreateFolder(root, DRIVE_ROOT_FOLDER);
    var jobFolder   = getOrCreateFolder(rootFolder, jobId);
    var dateFolder  = getOrCreateFolder(jobFolder, dateStr);
    var procFolder  = getOrCreateFolder(dateFolder, process);

    // ชื่อไฟล์
    var fileName = empId + '_' + empName + '_' + timeStr + '.' + ext;
    fileName = fileName.replace(/[\/:*?"<>|]/g, '_');

    // สร้างไฟล์
    var blob = Utilities.newBlob(
      Utilities.base64Decode(imgData),
      mime,
      fileName
    );
    var file = procFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    Logger.log('Upload สำเร็จ: ' + fileName + ' → ' + fileUrl);

    return { ok: true, url: fileUrl, fileName: fileName };

  } catch(e) {
    Logger.log('Upload Error: ' + e.toString());
    return { ok: false, message: e.toString() };
  }
}

// ── ทดสอบ ──
function testUploadPhoto() {
  // สร้างรูปทดสอบ 1x1 pixel สีแดง
  var testBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg==';
  var result = uploadPhoto({
    job_id:       'JM-68/0944',
    emp_id:       'EMP001',
    emp_name:     'สมชาย รักงาน',
    process:      'Fabrication',
    date:         new Date().toLocaleDateString('th-TH'),
    time:         formatTime(new Date()),
    photo_base64: testBase64,
  });
  Logger.log('Test Upload: ' + JSON.stringify(result));
}

// ============================================================
// SAVE PHOTO URLS — บันทึก ImgBB URLs เป็น Hyperlink ใน Job_Log
// ============================================================
function savePhotoUrls(data) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { ok: false, message: 'ไม่พบ Job_Log' };

    var values = sheet.getDataRange().getValues();
    var empId  = String(data.emp_id  || '').trim();
    var jobId  = String(data.job_id  || '').trim();
    var urls   = String(data.photo_urls || '').split('|').map(function(u){ return u.trim(); }).filter(Boolean);

    if (urls.length === 0) return { ok: false, message: 'ไม่มี URL' };

    // หาแถว Check Out ล่าสุดของพนักงานและ Job นี้
    for (var i = values.length - 1; i >= 1; i--) {
      var row = values[i];
      if (String(row[COL.EMP_ID]).trim() === empId &&
          String(row[COL.JOB_ID]).trim() === jobId &&
          String(row[COL.STATUS]) === 'Check Out') {

        var noteCol  = COL.NOTE + 1; // คอลัมน์หมายเหตุ
        var noteCell = sheet.getRange(i+1, noteCol);

        // สร้าง Rich Text พร้อม Hyperlink
        var rtb = SpreadsheetApp.newRichTextValue();
        var fullText = '';
        var links    = [];

        urls.forEach(function(url, idx) {
          var label = '📷 รูปที่ ' + (idx + 1);
          var sep   = idx < urls.length - 1 ? '   ' : '';
          links.push({ text: label, url: url });
          fullText += label + sep;
        });

        rtb.setText(fullText);

        // ใส่ Hyperlink ทีละรูป
        var pos = 0;
        links.forEach(function(link) {
          var start = fullText.indexOf(link.text, pos);
          var end   = start + link.text.length;
          rtb.setLinkUrl(start, end, link.url);
          pos = end;
        });

        noteCell.setRichTextValue(rtb.build());
        noteCell.setFontColor('#1565c0');
        noteCell.setFontWeight('bold');

        Logger.log('บันทึก ' + urls.length + ' Photo Links สำเร็จ');
        return { ok: true, count: urls.length };
      }
    }
    return { ok: false, message: 'ไม่พบแถว Check Out' };
  } catch(e) {
    Logger.log('savePhotoUrls Error: ' + e);
    return { ok: false, message: e.toString() };
  }
}

// ============================================================
// PAYROLL PERIOD UTILITIES
// ============================================================

// ดึงชื่อ Tab ทั้งหมดของ Job_Log
function getAllJobLogSheets() {
  var ss     = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var result = [];
  sheets.forEach(function(s) {
    if (s.getName().indexOf('Job_Log_') === 0) {
      result.push(s.getName());
    }
  });
  return result.sort().reverse(); // เรียงจากใหม่ → เก่า
}

// ทดสอบ — ดู Tab ปัจจุบัน
function testPayrollPeriod() {
  var name = getPayrollSheetName();
  Logger.log('Tab ปัจจุบัน: ' + name);
  Logger.log('Tab ทั้งหมด: ' + getAllJobLogSheets().join(', '));
}
