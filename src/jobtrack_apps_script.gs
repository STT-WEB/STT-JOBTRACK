// ============================================================
// JOBTRACK Apps Script v4.0 — Band-based Hour Engine + Multiplier
// ============================================================
// v4.0 เปลี่ยนการคำนวณชั่วโมงจาก "ตัด 17:00 ครั้งเดียว" เป็น
// "overlay ช่วงเวลา (time-band)" แตกเป็นหลาย Work Hour Type Code
// ต่อ 1 record + คูณ Multiplier ตามประเภทพนักงาน (HOUR_TYPE_RULE)
// ============================================================

// ── ฐานข้อมูลหลัก (Employee_List, Job_List, ปฏิทิน, HOUR_TYPE_RULE, อัตราค่าจ้าง ฯลฯ) — คงที่ ──
var DATA_SHEET_ID = '1MYWORYN3sOjov3Gxv3UqCV1jRSxgxwGi1tRomFUGSr0';
// ── ไฟล์บันทึกเวลา Job_Log แยกรายปี — ปีใหม่เพิ่ม ID ไฟล์ใหม่ตรงนี้ (คีย์ = ปี ค.ศ.) ──
var LOG_FILE_BY_YEAR = {
  '2026': '1ZPl3uVRtM5r4sPA-yX1OwTyp34XCTIcKRC0Sx8qsr9s',   // JOBTRACK_Job_Log 2026
  // '2027': 'ใส่ Sheet ID ไฟล์ปี 2027 ที่นี่',
};
function getLogSpreadsheetId() {
  var y = String(new Date().getFullYear());
  if (LOG_FILE_BY_YEAR[y]) return LOG_FILE_BY_YEAR[y];
  var keys = Object.keys(LOG_FILE_BY_YEAR).sort();
  return keys.length ? LOG_FILE_BY_YEAR[keys[keys.length-1]] : DATA_SHEET_ID;
}
var LOG_SHEET_ID = getLogSpreadsheetId();
var JOB_LIST_SHEET    = 'Job_List';
var EMP_LIST_SHEET    = 'Employee_List';
var DATA_EMP_SHEET    = 'DATA EMPLOYEE';
var CALENDAR_SHEET    = 'ประเภทวันทำงาน';
var HOUR_RULE_SHEET   = 'HOUR_TYPE_RULE';   // ตารางหน่วยคูณค่าแรง
var EMP_TYPE_COL      = 5;                   // Employee_List คอลัมน์ F = ประเภทพนักงาน (0-indexed)
var LINE_CHANNEL_ID   = '2009924671';
var LINE_CHANNEL_SECRET = 'c06a208ad23109b13a050d6a3fc3d4b9';
var OT_START_HOUR     = 17;
var WORK_START_HOUR   = 8;
var WORK_HOURS        = 8;
var BASE_URL          = 'https://stt-web.github.io/STT-JOBTRACK/job_checkin_app.html';
var PHOTO_ROOT_FOLDER = 'JOBTRACK_Photos';   // โฟลเดอร์เก็บรูปใน Drive
/* ★★ กติกาสำคัญ : เวลารอคิวของเซิร์ฟเวอร์ (waitLock) ต้องน้อยกว่า
   เวลาที่หน้าเว็บรอคำตอบ (apiPost timeout = 20 วินาที) เสมอ
   ไม่งั้นหน้าเว็บจะขึ้น error ก่อนที่เซิร์ฟเวอร์จะทันตอบ
   พนักงานเห็นว่า "ระบบไม่ออกให้" ทั้งที่บางครั้งเซิร์ฟเวอร์เขียนสำเร็จทีหลัง
   ตอนนี้ : เซิร์ฟเวอร์ 12 วินาที < หน้าเว็บ 20 วินาที  ← ห้ามตั้งเกิน 15 */
var PHOTO_RETENTION_DAYS = 14;               // เก็บรูปสูงสุด 2 สัปดาห์ แล้วลบอัตโนมัติ (ลดจาก 31 — พื้นที่ Drive เต็มจนสแกนออกไม่ได้)

// ── ขอบเขตช่วงเวลา (นาทีจากเที่ยงคืน) สำหรับ time-band ──
var T_0800 = 8 * 60;    // 480
var T_1200 = 12 * 60;   // 720
var T_1300 = 13 * 60;   // 780
var T_1700 = 17 * 60;   // 1020
var T_1730 = 17 * 60 + 30; // 1050 — OT grace: OT เริ่มนับ 17:30

var COL = {
  TIMESTAMP:0, DATE:1, TIME_IN:2, TIME_OUT:3, HOURS:4,        // A-E
  JOB_ID:5, JOB_NAME:6,                                        // F-G
  EMP_ID:7, EMP_NAME:8, EMP_DEPT:9, EMP_TYPE:10,               // H-K ประเภทพนักงาน
  SESSION:11, PROC_MAIN:12, PROC_CODE:13, PROC_SUB:14, PROC_COUNT:15, // L-P
  STATUS:16, LANG:17, NOTE:18,                                 // Q-S
  DAY_TYPE:19, HOUR_TYPE:20,                                   // T-U
  HOURS_NORMAL_NUM:21, HOURS_OT_NUM:22, PAY_HOURS:23           // V-X (ตัด Code1-4 ออก)
};
// ชื่อหมวด Process หลัก (A-L) สำหรับเติมคอลัมน์ "หมวด Process"
var PROC_MAIN_NAME = {A:'FABRICATION & ASSEMBLY',B:'WELDING',C:'PART',D:'KEMREX',E:'PAINTING',F:'SYSTEM',G:'PIPING & VALVES',H:'RE-INSTALLATION',I:'SUSPENSION',J:'STICKER',K:'QC',L:'CLEANING'};
// แยก label "A.1 ประกอบเชลล์" → {code:'A.1', sub:'ประกอบเชลล์', main:'A · FABRICATION & ASSEMBLY'}
function parseProcLabel(label){
  label=String(label||'').trim();
  var sp=label.indexOf(' ');
  var code=sp>0?label.substring(0,sp):label;
  var sub=sp>0?label.substring(sp+1).trim():'';
  var letter=code.split('.')[0];
  return {code:code, sub:sub, main:PROC_MAIN_NAME[letter]?(letter+' · '+PROC_MAIN_NAME[letter]):letter};
}

// ============================================================
// UTILITIES
// ============================================================
function pad(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}
function formatTime(d) { return pad(d.getHours(),2)+':'+pad(d.getMinutes(),2); }
function formatDate(d) { return d.toLocaleDateString('th-TH'); }

function toMinutes(t) {
  if (t instanceof Date) return t.getHours()*60+t.getMinutes();
  var s = String(t||'').trim();
  if (!s||s.length<3) return 0;
  if (s.indexOf('T')>-1) { var tp=s.split('T')[1]; var p=tp.split(':'); return parseInt(p[0])*60+parseInt(p[1]); }
  var p=s.split(':'); return parseInt(p[0]||0)*60+parseInt(p[1]||0);
}

function timeToString(t) {
  if (!t&&t!==0) return '';
  if (t instanceof Date) return pad(t.getHours(),2)+':'+pad(t.getMinutes(),2);
  var s=String(t).trim();
  if (s.indexOf('T')>-1) return s.split('T')[1].substring(0,5);
  return s.substring(0,5);
}

function minutesToHHMM(totalMin) {
  if (totalMin<=0) return '0:00';
  var h=Math.floor(totalMin/60); var m=totalMin%60;
  return h+':'+pad(m,2);
}

function minutesToDecimal(totalMin) {
  if (totalMin<=0) return 0;
  return Math.round(totalMin/60*100)/100;
}

// ============================================================
// v4.0 HOUR ENGINE — time-band segmentation + multiplier
// ============================================================
// แปลงสตริงวันที่ (ไทย พ.ศ. หรือ Date object) → Date (ระดับวัน)
function parseAnyDate(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v || '').trim();
  if (!s) return null;
  if (s.indexOf('/') < 0) { var dd = new Date(s); return isNaN(dd) ? null : new Date(dd.getFullYear(), dd.getMonth(), dd.getDate()); }
  var p = s.split('/');
  if (p.length < 3) return null;
  var day = parseInt(p[0], 10), mon = parseInt(p[1], 10) - 1, yr = parseInt(p[2], 10);
  if (yr > 2400) yr -= 543;   // พ.ศ. → ค.ศ.
  return new Date(yr, mon, day);
}
function addDaysThai(dateStr, n) {
  var d = parseAnyDate(dateStr);
  if (!d) return dateStr;
  d.setDate(d.getDate() + n);
  return d.toLocaleDateString('th-TH');
}

// ลดประเภทวันให้เหลือ 3 กลุ่ม
function dayKindOf(dayTypeStr) {
  var s = String(dayTypeStr || '');
  if (s.indexOf('นักขัตฤกษ์') >= 0) return 'นักขัตฤกษ์';
  if (s.indexOf('หยุด') >= 0)       return 'หยุด';
  return 'ปกติ';
}

// แผนที่ time-band ของแต่ละประเภทวัน → [{from,to,code}] (นาที 0..1440)
function getDayTypeBands(dayKind) {
  var normalCode, otCode;
  if (dayKind === 'นักขัตฤกษ์') { normalCode = '2B'; otCode = '4'; }
  else if (dayKind === 'หยุด')  { normalCode = '2A'; otCode = '4'; }
  else                          { normalCode = '1';  otCode = '3'; }
  // ★ ไม่มี band ของ 12:00-13:00 อีกแล้ว = ตัดพักเที่ยงทิ้งอัตโนมัติเสมอ
  //   ของเดิมตั้งเป็น Code 3 (OT ×1.5) คนที่ลืมสแกนออกตอนเที่ยงเลยได้ OT ฟรีทุกคน
  //   ถ้าทำงานผ่าเที่ยงจริง ให้ HR ติ๊กช่อง AH เอง แล้วเครื่องคำนวณจะบวกคืนให้
  return [
    { from: 0,      to: T_0800, code: otCode },     // ก่อน 08:00 = OT
    { from: T_0800, to: T_1200, code: normalCode }, // 08:00-12:00 ทำงานปกติ
    /* 12:00-13:00 พักเที่ยง — ไม่นับให้ทุกกรณี ต้องติ๊กอนุมัติถึงจะได้ */
    { from: T_1300, to: T_1700, code: normalCode }, // 13:00-17:00 ทำงานปกติ
    { from: T_1700, to: 1440,   code: otCode },     // 17:00-24:00 OT (OT gate จัดการที่ snapCheckOut)
  ];
}

// แตกช่วงเวลาทำงาน (นาที absolute) เป็น segment ตาม band + ข้ามเที่ยงคืน
function segmentByBands(inMin, outMin, kindForOffset) {
  var segs = [], cursor = inMin, guard = 0;
  while (cursor < outMin && guard < 7) {
    var dayOffset   = Math.floor(cursor / 1440);
    var dayStartAbs = dayOffset * 1440;
    var segEnd      = Math.min(outMin, dayStartAbs + 1440);
    var bands       = getDayTypeBands(kindForOffset(dayOffset));
    var loLocal = cursor - dayStartAbs, hiLocal = segEnd - dayStartAbs;
    for (var i = 0; i < bands.length; i++) {
      var b = bands[i];
      var lo = Math.max(loLocal, b.from), hi = Math.min(hiLocal, b.to);
      if (hi > lo) segs.push({ code: b.code, dayOffset: dayOffset, mins: hi - lo });
    }
    cursor = segEnd; guard++;
  }
  return segs;
}

// ── ปัดเวลา (snap) ตามกติกาบริษัท — นาทีจากเที่ยงคืน ──
function minToStr(m){ m=((m%1440)+1440)%1440; return pad(Math.floor(m/60),2)+':'+pad(m%60,2); }
function snapCheckIn(m){
  if (m <= 420) return m;      // <=07:00 : OT เช้า ใช้เวลาจริง
  if (m <= 495) return 480;    // 07:01-08:15 -> 08:00
  if (m <  715) return m;      // 08:16-11:54 : ใช้จริง
  if (m <= 725) return 720;    // 11:55-12:05 -> 12:00 (OT ผ่าเที่ยง)
  if (m <  770) return m;      // 12:06-12:49 : ใช้จริง
  if (m <= 795) return 780;    // 12:50-13:15 -> 13:00
  return m;                    // >=13:16 : ใช้จริง
}
function snapCheckOut(m){
  if (m <= 720)  return m;     // <=12:00 : ใช้จริง
  if (m <= 774)  return 720;   // 12:01-12:54 -> 12:00 (จบรอบเช้า)
  if (m <= 785)  return 780;   // 12:55-13:05 -> 13:00 (OT ผ่าเที่ยง)
  if (m <  1015) return m;     // 13:06-16:54 : ใช้จริง
  if (m <= 1049) return 1020;  // 16:55-17:29 -> 17:00 (ยังไม่ถึง 17:30 = ไม่มี OT เย็น)
  return m;                    // >=17:30 : ใช้จริง (OT นับตั้งแต่ 17:00)
}

// คำนวณ breakdown เต็มรูปของ 1 record
function calcWorkBreakdown(timeInStr, timeOutStr, startDateStr, empType) {
  var inMin  = toMinutes(timeInStr);
  var outMin = toMinutes(timeOutStr);
  if (outMin < inMin) outMin += 1440;   // ข้ามเที่ยงคืน

  var segs = segmentByBands(inMin, outMin, function (off) {
    var dStr = off === 0 ? startDateStr : addDaysThai(startDateStr, off);
    return dayKindOf(getDayType(dStr));
  });

  var byCodeMin = { '1':0, '2A':0, '2B':0, '3':0, '4':0 };
  var totalMin = 0, payMin = 0;
  for (var s = 0; s < segs.length; s++) {
    var seg = segs[s];
    byCodeMin[seg.code] = (byCodeMin[seg.code] || 0) + seg.mins;
    totalMin += seg.mins;
    var dStr = seg.dayOffset === 0 ? startDateStr : addDaysThai(startDateStr, seg.dayOffset);
    payMin += seg.mins * getMultiplier(empType, seg.code, dStr);
  }
  return {
    byCodeMin: byCodeMin,
    byCodeHours: {
      '1':  minutesToDecimal(byCodeMin['1']),
      '2A': minutesToDecimal(byCodeMin['2A']),
      '2B': minutesToDecimal(byCodeMin['2B']),
      '3':  minutesToDecimal(byCodeMin['3']),
      '4':  minutesToDecimal(byCodeMin['4']),
    },
    totalHours: minutesToDecimal(totalMin),
    payHours: Math.round(payMin / 60 * 100) / 100,
    segments: segs
  };
}

// ป้ายกำกับ "ประเภทชั่วโมง" จาก code ที่มีชั่วโมง > 0
function buildHourTypeLabel(byCodeMin) {
  var names = {
    '1':  '1. ชั่วโมงวันทำงานปกติ',
    '3':  '3. ชั่วโมงทำงานโอทีวันทำงานปกติ',
    '2A': '2A. ชั่วโมงวันทำงานหยุด (วันหยุดทั่วไป)',
    '2B': '2B. ชั่วโมงวันทำงานหยุดนักขัตฤกษ์',
    '4':  '4. ชั่วโมงทำงานโอทีวันหยุด/วันหยุดนักขัตฤกษ์ *3'
  };
  var parts = [];
  ['1','2A','2B','3','4'].forEach(function (c) { if (byCodeMin[c] > 0) parts.push(names[c]); });
  return parts.join(' + ') || '-';
}

// ── HOUR_TYPE_RULE: โหลด + cache + lookup multiplier ──
var _ruleCache = null;
function loadHourRules() {
  if (_ruleCache) return _ruleCache;
  var rules = [];
  var sheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName(HOUR_RULE_SHEET);
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r[0]) continue;
      rules.push({
        empType: String(r[0]).trim(),
        code:    String(r[1]).trim(),
        dayType: String(r[2]).trim(),
        name:    String(r[3]).trim(),
        mult:    parseFloat(r[4]) || 0,
        from:    r[5] ? parseAnyDate(r[5]) : null,
        to:      r[6] ? parseAnyDate(r[6]) : null,
        status:  String(r[7] || '').trim()
      });
    }
  }
  _ruleCache = rules;
  return rules;
}
function getMultiplier(empType, code, dateStr) {
  var DEFAULT = { '1':1, '2A':1, '2B':1, '3':1.5, '4':3 };
  var rules = loadHourRules();
  var d = parseAnyDate(dateStr);
  var cands = [];
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    if (r.empType !== empType) continue;
    if (r.code !== String(code)) continue;
    if (r.status && r.status.toLowerCase() !== 'active') continue;
    if (d && r.from && d < r.from) continue;
    if (d && r.to && d > r.to) continue;
    cands.push(r);
  }
  if (cands.length) {
    cands.sort(function (a, b) {
      var fa = a.from ? a.from.getTime() : 0, fb = b.from ? b.from.getTime() : 0;
      return fb - fa;
    });
    return cands[0].mult;
  }
  Logger.log('getMultiplier: ไม่พบกฎ ' + empType + '/' + code + '/' + dateStr + ' → ใช้ค่า default');
  return DEFAULT[code] != null ? DEFAULT[code] : 1;
}

// ดึงประเภทพนักงานจาก Employee_List
function getEmployeeType(empId) {
  var sheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName(EMP_LIST_SHEET);
  if (!sheet) return '';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(empId).trim()) return String(data[i][EMP_TYPE_COL] || '').trim();
  }
  return '';
}

// ============================================================
// PAYROLL PERIOD
// ============================================================
function getPayrollSheetName() {
  var today=new Date(); var day=today.getDate();
  var month=today.getMonth()+1; var year=today.getFullYear()+543;
  if (day>=26) { month+=1; if (month>12) { month=1; year+=1; } }
  return 'Job_Log_'+year+'_'+pad(month,2);
}

// ============================================================
// SHEET MANAGEMENT
// ============================================================
function getJobLogSheet() {
  var name=getPayrollSheetName(); var ss=SpreadsheetApp.openById(LOG_SHEET_ID);
  var sheet=ss.getSheetByName(name);
  if (!sheet) { sheet=ss.insertSheet(name); setupJobLogHeader(sheet); ss.setActiveSheet(sheet); ss.moveActiveSheet(1); }
  return sheet;
}

// ── ขยายจำนวนคอลัมน์ของชีตให้พอ (กัน range out of bounds) ──
function ensureColumns(sheet, n) {
  var cur = sheet.getMaxColumns();
  if (cur < n) sheet.insertColumnsAfter(cur, n - cur);
}

function setupJobLogHeader(sheet) {
  if (!sheet) sheet = getJobLogSheet();
  var headers = [
    'Timestamp','วันที่','เวลาเข้า','เวลาออก','ชม.รวม',
    'Job ID','ชื่องาน/ลูกค้า',
    'รหัสพนักงาน','ชื่อพนักงาน','แผนก','ประเภทพนักงาน',
    'รหัสรอบงาน','หมวด Process','รหัสงาน','ชื่องานย่อย','จำนวน Process',
    'สถานะ','ภาษา','หมายเหตุ',
    'ประเภทวันทำงาน','ประเภทชั่วโมง',
    'ชม.ปกติ','ชม.OT','ชม.คิดค่าแรง'
  ];
  ensureColumns(sheet, headers.length);   // 29 คอลัมน์
  sheet.appendRow(headers);
  var hRange=sheet.getRange(1,1,1,headers.length);
  hRange.setBackground('#CC0000').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sheet.setFrozenRows(1);
  var widths=[150,90,70,70,70,90,180,90,140,110,110,170,170,70,150,70,80,50,160,150,180,80,80,100];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  sheet.getRange(2,COL.TIME_IN+1,3000,1).setNumberFormat('@STRING@');
  sheet.getRange(2,COL.TIME_OUT+1,3000,1).setNumberFormat('@STRING@');
  sheet.getRange(2,COL.SESSION+1,3000,1).setNumberFormat('@STRING@');
  sheet.getRange(2,COL.PROC_MAIN+1,3000,3).setNumberFormat('@STRING@'); // M,N,O
  sheet.getRange(2,COL.HOURS+1,3000,1).setNumberFormat('0.00');
  sheet.getRange(2,COL.HOURS_NORMAL_NUM+1,3000,3).setNumberFormat('0.00'); // V,W,X
}

// ล้างชีตเดือนปัจจุบัน แล้วสร้างหัวคอลัมน์ใหม่ 29 ช่อง (ใช้กับข้อมูลทดสอบ)
function rebuildCurrentSheet() {
  var sheet = getJobLogSheet();
  sheet.clear();
  sheet.clearNotes();
  setupJobLogHeader(sheet);
  Logger.log('rebuild ' + sheet.getName() + ' ด้วยโครงใหม่ 24 คอลัมน์');
}

function colorStatus(sheet, rowIndex, status) {
  var cell=sheet.getRange(rowIndex,COL.STATUS+1);
  if (status==='Check In') cell.setBackground('#E8F5E9').setFontColor('#1B5E20');
  else cell.setBackground('#FFEBEE').setFontColor('#7F0000');
  cell.setFontWeight('bold');
}

// ============================================================
// doGet, doPost, doOptions
// ============================================================
function doGet(e) {
  var params=e.parameter; var callback=params.callback||''; var dataStr=params.data||'{}'; var result;
  try { var data=JSON.parse(dataStr); result=dispatch(data); } catch(err) { result={ok:false,message:err.toString()}; }
  var json=JSON.stringify(result);
  if (callback) return ContentService.createTextOutput(callback+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function doOptions(e) { return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT); }
function doPost(e) {
  try { var data=JSON.parse(e.postData.contents); var result=dispatch(data); return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON); }
  catch(err) { return ContentService.createTextOutput(JSON.stringify({ok:false,message:err.toString()})).setMimeType(ContentService.MimeType.JSON); }
}

// ============================================================
// DISPATCHER
// ============================================================
function dispatch(data) {
  var action=String(data.action||'');
  Logger.log('dispatch: '+action);
  if (action==='PING')              return {ok:true,pong:true,ts:new Date().toISOString()};
  if (action==='GET_JOB_INFO')      return getJobInfo(data.job_id);
  if (action==='CHECK_STATUS')      return checkOpenJob(data.emp_id);
  if (action==='CHECK_IN')          return doCheckIn(data);
  if (action==='CHECK_OUT')         return doCheckOut(data);
  if (action==='UPDATE_PROCESS')    return doUpdateProcess(data);
  if (action==='SAVE_PHOTO_URLS')   return doSavePhotoUrls(data);
  if (action==='UPLOAD_PHOTOS')     return doUploadPhotos(data);
  if (action==='REGISTER_EMPLOYEE') return registerEmployeeByCode(data.lineId,data.empId,data.lineName||'');
  if (action==='LINE_CALLBACK')     return handleLineCallback(data);
  return {ok:true,status:'JOBTRACK v3.2 running ✅'};
}

// ============================================================
// CHECK OPEN JOB
// ============================================================
function checkOpenJob(empId) {
  if (!empId) return {ok:true,hasOpen:false};
  var ss=SpreadsheetApp.openById(LOG_SHEET_ID); var currentSheet=getJobLogSheet(); var sheetsToCheck=[currentSheet];
  if (currentSheet.getLastRow()<5) {
    var prev=ss.getSheets().filter(function(s){ return s.getName().indexOf('Job_Log_')===0&&s.getName()!==currentSheet.getName(); });
    prev.sort(function(a,b){ return b.getName().localeCompare(a.getName()); });
    if (prev.length>0) sheetsToCheck.push(prev[0]);
  }
  for (var s=0;s<sheetsToCheck.length;s++) {
    var sheet=sheetsToCheck[s]; var values=sheet.getDataRange().getValues();
    for (var i=values.length-1;i>=1;i--) {
      var row=values[i];
      if (String(row[COL.EMP_ID]).trim()===String(empId).trim()&&String(row[COL.STATUS])==='Check In'&&String(row[COL.TIME_OUT]).trim()==='') {
        return {ok:true,hasOpen:true,openJob:{row_index:i+1,job_id:String(row[COL.JOB_ID]),job_name:String(row[COL.JOB_NAME]),process:String(row[COL.PROC_SUB]),time_in:timeToString(row[COL.TIME_IN]),date:String(row[COL.DATE]),sheet_name:sheet.getName()}};
      }
    }
  }
  return {ok:true,hasOpen:false};
}

// ============================================================
// เพิ่มแถวใหม่แบบปลอดภัย — ห้ามใช้ appendRow + getLastRow อีกเด็ดขาด
//
// ทำไม : appendRow กับ getLastRow เป็นคนละคำสั่ง ระหว่างสองคำสั่งนี้
//        คนอื่นแทรกแถวเข้ามาได้ getLastRow จะคืนเลขแถวของ "คนที่มาทีหลัง"
//        แล้วเราจะเอาข้อมูลของเราไปเขียนทับแถวเขา
//        เคสจริง 26/8/2569 08:01:30 : รหัสรอบงานของ 1069018 ไปตกบนแถวของ 4900044
//
// ตัวนี้ต้องเรียกใต้ LockService เท่านั้น จองเลขแถว เขียนรวดเดียว แล้วคืนเลขแถวไปเลย
// จะได้ไม่มีใครต้องเรียก getLastRow ซ้ำอีก
// ============================================================
function appendRowSafe_(sheet, row) {
  var r = sheet.getLastRow() + 1;
  var w = row.slice();
  while (w.length < 29) w.push('');
  // ตั้งรูปแบบเป็นข้อความก่อนเขียน ค่าจะได้ไม่โดน Sheets แปลงเป็นเวลา/ตัวเลข
  // (นี่คือสาเหตุที่เวลาเข้าบางแถวขึ้น 8:00 ชิดขวา แทนที่จะเป็น 08:00 ชิดซ้าย)
  sheet.getRange(r, COL.TIME_IN + 1).setNumberFormat('@STRING@');
  sheet.getRange(r, COL.TIME_OUT + 1).setNumberFormat('@STRING@');
  sheet.getRange(r, COL.SESSION + 1).setNumberFormat('@STRING@');
  sheet.getRange(r, COL.PROC_SUB + 1).setNumberFormat('@STRING@');
  sheet.getRange(r, 1, 1, w.length).setValues([w]);
  return r;
}

// ============================================================
// CHECK IN
// ============================================================
function doCheckIn(data) {
  // ตรวจงานค้างก่อนเข้าล็อก — ส่วนนี้อ่านอย่างเดียวและช้า ไม่ควรถือล็อกไว้
  var status=checkOpenJob(data.emp_id);
  if (status.hasOpen) return {ok:false,blocked:true,message:'OPEN_JOB_EXISTS',openJob:status.openJob};
  var sheet=getJobLogSheet(); var now=new Date();
  // ★ เก็บ "เวลาจริง" ลงชีต ไม่ปัดตั้งแต่ตอนสแกน
  //   ของเดิมปัดแล้วเขียนทับเวลาจริงทิ้งเลย พอ HR เปิดคอลัมน์ C มาก็เห็นแต่เวลาที่ปัดแล้ว
  //   ตรวจสอบย้อนหลังไม่ได้ และขัดกับที่ตกลงกันว่า C·D = เวลาจริง / Y·Z = เวลาที่ปัด
  //   เวลาที่ปัดให้เครื่องคำนวณ (Recheck) ทำตอนคำนวณแทน จะแก้กฎเมื่อไหร่ก็ได้ไม่ต้องแก้ข้อมูลเก่า
  var nowMin=now.getHours()*60+now.getMinutes();
  var timeStr=formatTime(now);                      /* เวลาจริง → ลงชีต */
  var timeSnap=minToStr(snapCheckIn(nowMin));       /* เวลาที่ปัด → โชว์บนหน้าจอพนักงาน */
  var dateStr=formatDate(now); var dayType=getDayType(dateStr);
  var procRaw=String(data.process||'');
  var procCount=procRaw.split(' || ').map(function(x){return x.trim();}).filter(Boolean).length || 1;
  var sessionId='S-'+String(data.emp_id||'')+'-'+Utilities.formatDate(now,'Asia/Bangkok','yyMMddHHmmss');
  var row=new Array(29).fill('');
  row[COL.TIMESTAMP]=now.toLocaleString('th-TH'); row[COL.DATE]=dateStr; row[COL.TIME_IN]=timeStr;
  row[COL.JOB_ID]=String(data.job_id||''); row[COL.JOB_NAME]=String(data.job_name||'');
  row[COL.EMP_ID]=String(data.emp_id||''); row[COL.EMP_NAME]=String(data.emp_name||''); row[COL.EMP_DEPT]=String(data.emp_dept||'');
  row[COL.EMP_TYPE]=getEmployeeType(String(data.emp_id||''));
  row[COL.SESSION]=sessionId; row[COL.PROC_SUB]=procRaw; row[COL.PROC_COUNT]=procCount;
  row[COL.STATUS]='Check In'; row[COL.LANG]=String(data.lang||'th'); row[COL.DAY_TYPE]=dayType;

  // ★ ล็อกเฉพาะตอนเขียน — ใช้ล็อกตัวเดียวกับ doCheckOut จะได้ไม่แย่งแถวข้ามกัน
  //   ถือล็อกสั้น ๆ แค่ครึ่งวินาที คนสแกนพร้อมกันหลายคนจึงต่อคิวได้ทัน
  var _lock=LockService.getScriptLock();
  try{ _lock.waitLock(12000); }catch(e){ return {ok:false,code:'BUSY',message:'ระบบกำลังบันทึกของคนอื่นอยู่ กรุณากดสแกนอีกครั้งใน 5 วินาที'}; }
  var lastRow;
  try{
    lastRow=appendRowSafe_(sheet,row);
    colorStatus(sheet,lastRow,'Check In');
    SpreadsheetApp.flush();   // ต้อง commit ก่อนปล่อยล็อก ไม่งั้นคิวถัดไปมองไม่เห็นแถวนี้
  }finally{ try{_lock.releaseLock();}catch(e){} }

  /* หน้าจอพนักงานยังโชว์เวลาที่ปัดเหมือนเดิม จะได้ไม่งงว่าทำไมไม่ใช่ 08:00
     ส่วน time_in_raw คือเวลาจริงที่บันทึกลงชีต เผื่อหน้าเว็บอยากโชว์ทั้งคู่ */
  return {ok:true,action:'CHECK_IN',time_in:timeSnap,time_in_raw:timeStr,date:dateStr,session:sessionId,row:lastRow};
}
// ============================================================
// CHECK OUT
// ============================================================
function doCheckOut(data) {
  var _lock=LockService.getScriptLock();
  try{ _lock.waitLock(12000); }catch(e){ return {ok:false,code:'BUSY',message:'ระบบกำลังบันทึกของคนอื่นอยู่ กรุณากดสแกนอีกครั้งใน 5 วินาที'}; }
  try{
  var openResult=checkOpenJob(data.emp_id); var sheet,values;
  if (openResult.hasOpen&&openResult.openJob.sheet_name) sheet=SpreadsheetApp.openById(LOG_SHEET_ID).getSheetByName(openResult.openJob.sheet_name);
  else sheet=getJobLogSheet();
  values=sheet.getDataRange().getValues();
  // ★ เหมือนตอน Check In — เก็บเวลาจริงลงชีต ส่วนเวลาที่ปัดใช้เฉพาะคำนวณกับโชว์หน้าจอ
  var now=new Date(); var nowMinOut=now.getHours()*60+now.getMinutes();
  var timeOut=formatTime(now);                          /* เวลาจริง → ลงชีต */
  var timeOutSnap=minToStr(snapCheckOut(nowMinOut));    /* เวลาที่ปัด → ใช้คำนวณ */
  var targetRow=-1;
  for (var i=values.length-1;i>=1;i--) {
    var r=values[i];
    if (String(r[COL.EMP_ID]).trim()===String(data.emp_id).trim()&&String(r[COL.JOB_ID]).trim()===String(data.job_id).trim()&&String(r[COL.STATUS])==='Check In'&&String(r[COL.TIME_OUT]).trim()==='') {
      targetRow=i+1; break;
    }
  }
  if (targetRow===-1) {
    // อาจ Check Out ไปแล้วจากครั้งก่อน (retry หลัง timeout) — ถ้าเจอแถวที่ออกไปแล้ว ถือว่าสำเร็จ กัน phantom + กันหน้าเว็บขึ้นเขียวหลอก
    for (var j=values.length-1;j>=1;j--) {
      var rj=values[j];
      if (String(rj[COL.EMP_ID]).trim()===String(data.emp_id).trim()&&String(rj[COL.JOB_ID]).trim()===String(data.job_id).trim()&&String(rj[COL.STATUS])==='Check Out'&&String(rj[COL.TIME_OUT]).trim()!=='') {
        return {ok:true,action:'CHECK_OUT',already:true,time_in:timeToString(rj[COL.TIME_IN]),time_out:String(rj[COL.TIME_OUT]),proc_count:Number(rj[COL.PROC_COUNT])||1,message:'ALREADY_CHECKED_OUT'};
      }
    }
    return {ok:false,message:'NO_CHECKIN_FOUND'};
  }

  var timeInStr = timeToString(values[targetRow-1][COL.TIME_IN]);
  var startDate = String(values[targetRow-1][COL.DATE]);
  var dayTypeVal = String(values[targetRow-1][COL.DAY_TYPE] || getDayType(startDate));
  var empType = String(values[targetRow-1][COL.EMP_TYPE] || getEmployeeType(String(data.emp_id)));

  /* คอลัมน์ C เก็บเวลาจริงแล้ว ต้องปัดตอนนี้ก่อนเอาไปคำนวณ
     (แถวเก่าที่เก็บเวลาปัดไว้แล้ว ปัดซ้ำได้ค่าเดิม ไม่เพี้ยน) */
  var timeInSnap = minToStr(snapCheckIn(toMinutes(timeInStr)));
  var bd = calcWorkBreakdown(timeInSnap, timeOutSnap, startDate, empType);

  // ── แยก Process (เลือกได้หลายงาน คั่นด้วย ' || ') → หารเฉลี่ยชั่วโมง ÷ N ──
  var baseRow = values[targetRow-1].slice();
  var procStr = String(baseRow[COL.PROC_SUB] || '').trim();
  var procs = procStr.split(' || ').map(function(x){ return x.trim(); }).filter(Boolean);
  if (procs.length === 0) procs = [procStr];
  var n = procs.length, scale = 1 / n;

  // แถวแรก = process แรก (อัปเดตแถวเดิม)
  writeCheckoutRow(sheet, targetRow, timeOut, procs[0], n, dayTypeVal, bd, scale);
  // process ที่ 2..N = เพิ่มแถวใหม่ (คัดลอกข้อมูลฐานจากแถวเช็คอิน)
  // ★ แทรกแถวต่อจากแถวแรกทันที ไม่เอาไปต่อท้ายชีต
  //   ของเดิมใช้ appendRow แถว Process ที่ 2 เลยไปโผล่ห่างจากแถวแรกเป็นสิบแถว
  //   (ระหว่างที่คนนี้ทำงานอยู่ คนอื่นสแกนเข้ามาคั่นเรื่อย ๆ)
  //   HR เปิดมาเห็นแถวเดียวก็นึกว่าอีก Process หายไป แล้วไปเจอทีหลังก็นึกว่าลงซ้ำ
  //   แทรกติดกันแบบนี้ 1 รอบงานจะอ่านได้จบในสายตาเดียว
  //   ปลอดภัยเพราะอยู่ใต้ LockService และแทรก "ใต้" targetRow เลขแถวแรกไม่ขยับ
  //   ★ Timestamp ต้องบังคับเป็นข้อความ ไม่งั้นแถวที่ก๊อปมาจะเหลือแต่วันที่ เวลาจริงหายไป
  //     (คอลัมน์ A คือที่เดียวที่เก็บวินาทีของการสแกนไว้ ห้ามให้หาย)
  var stampTxt = String(values[targetRow-1][COL.TIMESTAMP] || '');
  for (var pi = 1; pi < n; pi++) {
    var at = targetRow + pi;
    sheet.insertRowAfter(at - 1);
    sheet.getRange(at, 1, 1, baseRow.length).setValues([baseRow]);
    sheet.getRange(at, COL.TIMESTAMP + 1).setNumberFormat('@STRING@').setValue(stampTxt);
    writeCheckoutRow(sheet, at, timeOut, procs[pi], n, dayTypeVal, bd, scale, baseRow[COL.TIME_IN]);
  }

  /* ★★★ คำนวณชั่วโมงด้วย "เครื่องคิดเลขตัวเดียวของระบบ" (rcCalc_ ใน Recheck.gs)
     แล้วเขียนผลลงแถวเลย ตั้งแต่วินาทีที่พนักงานกดออกงาน

     ทำไมต้องทำตรงนี้ : ของเดิม JOBTRACK คิดชั่วโมงเองด้วยกฎเก่า แล้วฝั่ง Hub
     ต้องมากวาดทั้งชีตทุก 5 นาทีเพื่อคำนวณทับให้ถูก — ทั้งเปลืองและทำชีตกระตุก
     พอ JOBTRACK ใช้เครื่องเดียวกันตั้งแต่แรก แถวก็ถูกตั้งแต่เกิด ไม่ต้องมีใครตามเก็บ

     Recheck.gs เป็นไฟล์เดียวกับที่ Hub ใช้ ก๊อปขึ้นมาตอน deploy ต้นฉบับมีที่เดียว */
  try {
    var rcCal  = rcLoadCal_();
    var rcIsM  = String(empType).indexOf('รายเดือน') >= 0;
    for (var rq = 0; rq < n; rq++) {
      var rcRow = targetRow + rq;                    /* แถวเรียงติดกันแล้ว */
      var rcR   = rcCalc_(timeInStr, timeOut, startDate, rcIsM, rcCal, false, dayTypeVal);
      var rcSt  = rcStatus_(rcR, n, n, false, false);
      if (n > 1) rcSt.info.push('✓ หารให้แล้ว · งานนี้ลง ' + n + ' จ๊อบ');
      rcWriteRow_(sheet, rcRow, rcR, n, dayTypeVal, rcSt.bad.concat(rcSt.info).join('  |  '));
    }
  } catch (eCalc) {
    /* คิดไม่ได้ก็ไม่เป็นไร เวลาเข้า-ออกบันทึกไว้แล้ว รอบตรวจทานตี 3 จะเก็บให้ */
    Logger.log('คำนวณชั่วโมงตอน Check Out ไม่สำเร็จ: ' + eCalc);
  }

  updateDailyHourAlert(sheet, values, String(data.emp_id), startDate, targetRow);
  SpreadsheetApp.flush();   // commit ก่อนปล่อยล็อก คิวถัดไปจะได้เห็นแถวที่เพิ่งเพิ่ม

  var totMin = bd.byCodeMin['1']+bd.byCodeMin['2A']+bd.byCodeMin['2B']+bd.byCodeMin['3']+bd.byCodeMin['4'];
  Logger.log('CHECK OUT v4.2: '+data.emp_id+' ['+empType+'] '+n+' proc ÷'+n+' ('+timeInStr+'-'+timeOut+') รวม='+minutesToHHMM(Math.round(totMin))+' payHrs='+bd.payHours);
  return {ok:true,action:'CHECK_OUT',time_in:timeInStr,time_out:timeOut,proc_count:n,hours:minutesToHHMM(Math.round(totMin)),pay_hours:bd.payHours,breakdown:bd.byCodeHours};
  }finally{ try{_lock.releaseLock();}catch(e){} }
}

// เขียนค่า Check Out ลง 1 แถว รองรับหารเฉลี่ย (scale = 1/N) — ใช้ทั้งแถวเดิมและแถวที่เพิ่มใหม่
function writeCheckoutRow(sheet, rowIndex, timeOut, procLabel, count, dayTypeVal, bd, scale, timeInForNew) {
  var nMin = (bd.byCodeMin['1'] + bd.byCodeMin['2A'] + bd.byCodeMin['2B']) * scale;
  var oMin = (bd.byCodeMin['3'] + bd.byCodeMin['4']) * scale;
  var hourTypeVal = buildHourTypeLabel(bd.byCodeMin);
  var pl = parseProcLabel(procLabel);
  if (timeInForNew !== undefined) sheet.getRange(rowIndex,COL.TIME_IN+1).setNumberFormat('@STRING@').setValue(String(timeInForNew));
  sheet.getRange(rowIndex,COL.TIME_OUT+1).setNumberFormat('@STRING@').setValue(timeOut);
  /* E · V · W · X และคอลัมน์ผลคำนวณทั้งหมด เขียนโดย rcWriteRow_ ที่เดียว (ดู doCheckOut) */
  sheet.getRange(rowIndex,COL.STATUS+1).setValue('Check Out');
  sheet.getRange(rowIndex,COL.PROC_MAIN+1).setNumberFormat('@STRING@').setValue(pl.main);
  sheet.getRange(rowIndex,COL.PROC_CODE+1).setNumberFormat('@STRING@').setValue(pl.code);
  sheet.getRange(rowIndex,COL.PROC_SUB+1).setNumberFormat('@STRING@').setValue(pl.sub);
  sheet.getRange(rowIndex,COL.PROC_COUNT+1).setValue(count);
  sheet.getRange(rowIndex,COL.DAY_TYPE+1).setValue(dayTypeVal);
  colorStatus(sheet, rowIndex, 'Check Out');
  var hoursCell = sheet.getRange(rowIndex,COL.HOURS+1);
  if ((nMin+oMin) > 480) hoursCell.setBackground('#FFF3CD').setFontColor('#856404');
  else hoursCell.setBackground('#E8F5E9').setFontColor('#1B5E20');
}
// ============================================================
// updateDailyHourAlert
// ============================================================
function updateDailyHourAlert(sheet, values, empId, dateStr, currentRow) {
  try {
    var totalNormal=0;
    for (var i=1;i<values.length;i++) {
      var row=values[i];
      if (String(row[COL.EMP_ID]).trim()!==empId) continue;
      if (String(row[COL.DATE]).trim()!==dateStr) continue;
      if (String(row[COL.STATUS])!=='Check Out') continue;
      totalNormal += parseFloat(row[COL.HOURS_NORMAL_NUM])||0;
    }
    totalNormal += parseFloat(sheet.getRange(currentRow,COL.HOURS_NORMAL_NUM+1).getValue())||0;
    var diff=totalNormal-WORK_HOURS;
    var noteCell=sheet.getRange(currentRow,COL.HOURS_NORMAL_NUM+1);
    if (Math.abs(diff)<=0.17) noteCell.setNote('✅ ครบ 8 ชม.');
    else if (diff<0) noteCell.setNote('⚠️ ขาด '+Math.abs(diff).toFixed(2)+' ชม.');
    else noteCell.setNote('ℹ️ เกิน '+diff.toFixed(2)+' ชม.');
  } catch(e) { Logger.log('updateDailyHourAlert error: '+e); }
}
// ============================================================
// UPDATE PROCESS
// ============================================================
function doUpdateProcess(data) {
  try {
    var empId=String(data.emp_id||'').trim(); var newProc=String(data.process||'').trim();
    if (!empId||!newProc) return {ok:false,message:'ข้อมูลไม่ครบ'};
    var open=checkOpenJob(empId); if (!open.hasOpen) return {ok:false,message:'ไม่พบงานค้าง'};
    var ss=SpreadsheetApp.openById(LOG_SHEET_ID); var sheet=ss.getSheetByName(open.openJob.sheet_name||getPayrollSheetName());
    if (!sheet) return {ok:false,message:'ไม่พบ Sheet'};
    sheet.getRange(open.openJob.row_index,COL.PROC_SUB+1).setNumberFormat('@STRING@').setValue(newProc);
    return {ok:true,process:newProc};
  } catch(e) { return {ok:false,message:e.toString()}; }
}

// ============================================================
// SAVE PHOTO URLS
// ============================================================
function doSavePhotoUrls(data) {
  try {
    var ss=SpreadsheetApp.openById(LOG_SHEET_ID); var empId=String(data.emp_id||'').trim(); var jobId=String(data.job_id||'').trim();
    var urls=String(data.photo_urls||'').split('|').map(function(u){return u.trim();}).filter(Boolean);
    if (urls.length===0) return {ok:false,message:'ไม่มี URL'};
    var allSheets=ss.getSheets().filter(function(s){return s.getName().indexOf('Job_Log_')===0;}).sort(function(a,b){return b.getName().localeCompare(a.getName());});
    for (var si=0;si<Math.min(allSheets.length,2);si++) {
      var sheet=allSheets[si]; var values=sheet.getDataRange().getValues();
      for (var i=values.length-1;i>=1;i--) {
        var row=values[i];
        if (String(row[COL.EMP_ID]).trim()===empId&&String(row[COL.JOB_ID]).trim()===jobId&&String(row[COL.STATUS])==='Check Out') {
          var noteCell=sheet.getRange(i+1,COL.NOTE+1); var rtb=SpreadsheetApp.newRichTextValue();
          var text=''; var links=[];
          urls.forEach(function(url,idx){ var label='📷 รูปที่'+(idx+1); var sep=idx<urls.length-1?'   ':''; links.push({text:label,url:url}); text+=label+sep; });
          rtb.setText(text); var pos=0;
          links.forEach(function(link){ var s=text.indexOf(link.text,pos); var e=s+link.text.length; rtb.setLinkUrl(s,e,link.url); pos=e; });
          noteCell.setRichTextValue(rtb.build()).setFontColor('#1565C0').setFontWeight('bold');
          return {ok:true,count:urls.length};
        }
      }
    }
    return {ok:false,message:'ไม่พบแถว Check Out'};
  } catch(e) { return {ok:false,message:e.toString()}; }
}

// ============================================================
// v4.1 PHOTO UPLOAD → GOOGLE DRIVE (แทน imgbb)
// ============================================================
function getRootPhotoFolder() {
  var it = DriveApp.getFoldersByName(PHOTO_ROOT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_ROOT_FOLDER);
}
function getOrCreateSubfolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function sanitizeName(s) {
  return String(s || '').replace(/[\\\/:*?"<>|]/g, '-');
}

// รับรูป base64 หลายรูป → เซฟลง Drive โฟลเดอร์ตามวันที่ → เขียนลิงก์ลง note
function doUploadPhotos(data) {
  try {
    var imgs = data.images || [];
    if (!imgs.length) return { ok:false, message:'ไม่มีรูป' };
    var now = new Date();
    var dateName = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
    var folder = getOrCreateSubfolder(getRootPhotoFolder(), dateName);
    var timeStamp = Utilities.formatDate(now, 'Asia/Bangkok', 'HHmmss');
    var urls = [];
    for (var i = 0; i < imgs.length; i++) {
      var b = String(imgs[i] || '');
      var c = b.indexOf(',');
      if (c > -1 && b.substring(0, c).indexOf('base64') > -1) b = b.substring(c + 1);
      if (!b) continue;
      var bytes = Utilities.base64Decode(b);
      var fname = sanitizeName((data.job_id||'JOB')+'_'+(data.emp_id||'')+'_'+timeStamp+'_'+(i+1)) + '.jpg';
      var file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', fname));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls.push('https://drive.google.com/uc?export=view&id=' + file.getId());
    }
    if (!urls.length) return { ok:false, message:'decode รูปไม่ได้' };
    writePhotoUrlsToRow(String(data.emp_id||'').trim(), String(data.job_id||'').trim(), urls, String(data.proc_code||'').trim());
    return { ok:true, count:urls.length, urls:urls };
  } catch (e) { return { ok:false, message:e.toString() }; }
}

// เขียนลิงก์รูป (rich text) ลงช่องหมายเหตุของแถว Check Out
function writePhotoUrlsToRow(empId, jobId, urls, procCode) {
  var ss = SpreadsheetApp.openById(LOG_SHEET_ID);
  var allSheets = ss.getSheets().filter(function(s){return s.getName().indexOf('Job_Log_')===0;})
                    .sort(function(a,b){return b.getName().localeCompare(a.getName());});
  for (var si = 0; si < Math.min(allSheets.length, 2); si++) {
    var sheet = allSheets[si]; var values = sheet.getDataRange().getValues();
    for (var i = values.length-1; i >= 1; i--) {
      var row = values[i];
      if (String(row[COL.EMP_ID]).trim()===empId && String(row[COL.JOB_ID]).trim()===jobId && String(row[COL.STATUS])==='Check Out' && (!procCode || String(row[COL.PROC_CODE]).trim()===String(procCode).trim()) && String(row[COL.NOTE]||'').trim()==='') {
        var noteCell = sheet.getRange(i+1, COL.NOTE+1);
        var rtb = SpreadsheetApp.newRichTextValue();
        var text = '', links = [];
        urls.forEach(function(url,idx){ var label='📷 รูปที่'+(idx+1); links.push({text:label,url:url}); text += label + (idx<urls.length-1?'   ':''); });
        rtb.setText(text); var pos=0;
        links.forEach(function(link){ var st=text.indexOf(link.text,pos); var en=st+link.text.length; rtb.setLinkUrl(st,en,link.url); pos=en; });
        noteCell.setRichTextValue(rtb.build()).setFontColor('#1565C0').setFontWeight('bold');
        return true;
      }
    }
  }
  return false;
}

// ลบรูปเก่ากว่า PHOTO_RETENTION_DAYS อัตโนมัติ (โฟลเดอร์ตามวันที่ yyyy-MM-dd)
function cleanupOldPhotos() {
  try {
    var root = getRootPhotoFolder();
    var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - PHOTO_RETENTION_DAYS);
    var folders = root.getFolders(); var removed = 0;
    while (folders.hasNext()) {
      var f = folders.next();
      var p = f.getName().split('-');
      if (p.length === 3) {
        var d = new Date(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10));
        if (!isNaN(d.getTime()) && d < cutoff) { f.setTrashed(true); removed++; }
      }
    }
    Logger.log('cleanupOldPhotos: ลบ ' + removed + ' โฟลเดอร์ (เก่ากว่า ' + PHOTO_RETENTION_DAYS + ' วัน)');
  } catch (e) { Logger.log('cleanupOldPhotos error: ' + e); }
}

// ติดตั้ง trigger ลบรูปอัตโนมัติทุกวัน 02:00
function installPhotoCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='cleanupOldPhotos') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('cleanupOldPhotos').timeBased().everyDays(1).atHour(2).create();
  Logger.log('ติดตั้ง trigger ลบรูปอัตโนมัติแล้ว (ทุกวัน 02:00)');
}

// ============================================================
// DAY TYPE & HOUR TYPE
// ============================================================
function getDayType(dateStr) {
  try {
    var sheet=SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName(CALENDAR_SHEET);
    if (!sheet) return 'วันทำงานปกติ';
    var data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      if (new Date(data[i][0]).toLocaleDateString('th-TH')===dateStr) return String(data[i][2]).trim();
    }
    var parts=dateStr.split('/'); var d=new Date(parseInt(parts[2])-543,parseInt(parts[1])-1,parseInt(parts[0]));
    return (d.getDay()===0)?'วันหยุด':'วันทำงานปกติ';
  } catch(e) { return 'วันทำงานปกติ'; }
}

// ============================================================
// LINE LOGIN
// ============================================================
function handleLineCallback(data) {
  try {
    if (!data.code) return {ok:false,message:'ไม่มี authorization code'};
    var tokenRes=UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/token',{method:'post',contentType:'application/x-www-form-urlencoded',payload:{grant_type:'authorization_code',code:data.code,redirect_uri:data.redirect,client_id:LINE_CHANNEL_ID,client_secret:LINE_CHANNEL_SECRET}});
    var token=JSON.parse(tokenRes.getContentText());
    if (!token.access_token) return {ok:false,message:'ไม่สามารถแลก Token ได้'};
    var profile=JSON.parse(UrlFetchApp.fetch('https://api.line.me/v2/profile',{headers:{'Authorization':'Bearer '+token.access_token}}).getContentText());
    var empData=getEmployeeByLineId(profile.userId);
    if (!empData) return {ok:true,profile:profile,empId:'',empName:'',empDept:'',empRole:'',empDeptMain:'',needRegister:true};
    return {ok:true,profile:profile,empId:empData.empId,empName:empData.empName,empDept:empData.empDept,empRole:empData.empRole,empDeptMain:empData.empDeptMain||'',needRegister:false};
  } catch(e) { return {ok:false,message:e.toString()}; }
}

function getEmployeeByLineId(lineId) {
  var sheet=SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName(EMP_LIST_SHEET);
  if (!sheet) return null;
  var data=sheet.getDataRange().getValues();
  for (var i=1;i<data.length;i++) {
    if (String(data[i][3]).trim()===String(lineId).trim()) {
      return {empId:String(data[i][0]).trim(),empName:String(data[i][1]).trim(),empDept:String(data[i][2]).trim(),empRole:String(data[i][4]).trim(),empDeptMain:String(data[i][9]||'').trim()};
    }
  }
  return null;
}

function registerEmployeeByCode(lineId, empId, lineName) {
  var ss=SpreadsheetApp.openById(DATA_SHEET_ID); var empSheet=ss.getSheetByName(EMP_LIST_SHEET);
  if (!empSheet) return {ok:false,message:'ไม่พบ Employee List'};
  var data=empSheet.getDataRange().getValues(); var today=new Date().toLocaleDateString('th-TH');
  for (var i=1;i<data.length;i++) {
    if (String(data[i][0]).trim()===String(empId).trim()) {
      empSheet.getRange(i+1,4).setValue(lineId).setBackground('#E8F5E9').setFontColor('#1B5E20').setFontWeight('bold');
      empSheet.getRange(i+1,8).setValue(String(lineName||'')).setFontColor('#1565C0');
      empSheet.getRange(i+1,9).setValue(today);
      return {ok:true,empId:String(data[i][0]).trim(),empName:String(data[i][1]).trim(),empDept:String(data[i][2]).trim(),empRole:String(data[i][4]).trim(),empDeptMain:String(data[i][9]||'').trim()};
    }
  }
  return {ok:false,message:'ไม่พบรหัสพนักงาน '+empId};
}

// ============================================================
// IMPORT EMPLOYEE
// ============================================================
function importEmployeeData() {
  var ss=SpreadsheetApp.openById(DATA_SHEET_ID); var srcSheet=ss.getSheetByName(DATA_EMP_SHEET);
  if (!srcSheet) { Logger.log('ไม่พบ: '+DATA_EMP_SHEET); return; }
  var empSheet=ss.getSheetByName(EMP_LIST_SHEET)||ss.insertSheet(EMP_LIST_SHEET);
  // เก็บค่ารันไทม์เดิม (LINE/Department/PIN/Role/บริษัท/Device/สถานะ) keyed by รหัส — กัน sync แล้วข้อมูล login หาย
  var keep={};
  if (empSheet.getLastRow()>1) { var ed=empSheet.getDataRange().getValues();
    for (var e=1;e<ed.length;e++) { var eid=String(ed[e][0]).trim(); if (!eid) continue;
      keep[eid]={line:ed[e][3]||'',lname:ed[e][7]||'',reg:ed[e][8]||'',dept:ed[e][9]||'',pin:ed[e][10]||'',role:ed[e][11]||'',co:ed[e][12]||'',dev:ed[e][13]||'',st:ed[e][14]||''};
    } }
  empSheet.clear(); empSheet.clearFormats();
  var headers=['รหัสพนักงาน','ชื่อพนักงาน','แผนก','LINE ID','ตำแหน่ง','ประเภทพนักงาน','Direct/Indirect','LINE Name','วันที่ลงทะเบียน','Department','PIN','Role','บริษัท','Device ID','สถานะ'];
  empSheet.appendRow(headers);
  empSheet.getRange(1,1,1,headers.length).setBackground('#0F1117').setFontColor('#00E187').setFontWeight('bold').setFontSize(11);
  empSheet.setFrozenRows(1);
  empSheet.getRange(1,11,empSheet.getMaxRows(),1).setNumberFormat('@STRING@'); // PIN เป็น text กัน 0 นำหน้าหาย
  empSheet.getRange(1,14,empSheet.getMaxRows(),1).setNumberFormat('@STRING@'); // Device ID เป็น text
  var srcData=srcSheet.getDataRange().getValues(); var headerRow=2;
  for (var i=0;i<srcData.length;i++) { if (String(srcData[i][3]).indexOf('รหัสพนักงาน')>=0||String(srcData[i][2]).indexOf('รหัสพนักงาน')>=0) { headerRow=i; break; } }
  var rows=[];
  for (var r=headerRow+1;r<srcData.length;r++) { var row=srcData[r]; var empId=String(row[3]).trim(); var empName=String(row[4]).trim(); if (!empId||isNaN(Number(empId))||!empName) continue;
    var k=keep[empId]||{}; var dept=String(row[2]).trim();
    rows.push([empId,empName,dept,k.line||'',String(row[7]).trim(),String(row[8]).trim(),String(row[10]).trim(),k.lname||'',k.reg||'',k.dept||dept,k.pin||'',k.role||'',k.co||'',k.dev||'',k.st||'']);
  }
  if (rows.length>0) { empSheet.getRange(2,1,rows.length,headers.length).setValues(rows);
    for (var j=0;j<rows.length;j++) { if (j%2===0) empSheet.getRange(j+2,1,1,headers.length).setBackground('#F8F9FA'); if (rows[j][3]) empSheet.getRange(j+2,4).setBackground('#E8F5E9').setFontColor('#1B5E20').setFontWeight('bold'); } }
  Logger.log('Sync: '+rows.length+' คน (เก็บ LINE/PIN/Role/Department ไว้)');
  try { SpreadsheetApp.getUi().alert('Sync พนักงานสำเร็จ! '+rows.length+' คน'); } catch(e) { Logger.log('Sync สำเร็จ '+rows.length+' คน (standalone)'); }
}

// ============================================================
// JOB LIST & QR
// ============================================================
function onJobListEdit(e) {
  var sheet=e.range.getSheet(); if (sheet.getName()!==JOB_LIST_SHEET) return;
  var row=e.range.getRow(); var col=e.range.getColumn(); if (row<2||col>3) return;
  var rowData=sheet.getRange(row,1,1,3).getValues()[0];
  var jobType=String(rowData[0]||'JM'); var jobCode=String(rowData[1]||''); var jobName=String(rowData[2]||'');
  if (!jobCode||!jobName) return;
  var qr=BASE_URL+'?job='+encodeURIComponent(jobCode)+'&name='+encodeURIComponent(jobName)+'&type='+encodeURIComponent(jobType)+'&cust='+encodeURIComponent(jobName);
  sheet.getRange(row,5).setValue(qr).setFontColor('#1565C0').setFontStyle('italic');
  if (!sheet.getRange(row,6).getValue()) sheet.getRange(row,6).setValue(new Date().toLocaleDateString('th-TH'));
  var sc=sheet.getRange(row,4); if (!sc.getValue()) sc.setValue('Active').setBackground('#E8F5E9').setFontColor('#1B5E20').setFontWeight('bold');
}

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='onJobListEdit') ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('onJobListEdit').forSpreadsheet(DATA_SHEET_ID).onEdit().create();
  Logger.log('Trigger ติดตั้งแล้ว');
}

// ============================================================
// GET JOB INFO
// ============================================================
function getJobInfo(jobId) {
  try {
    if (!jobId) return {ok:false,message:'ไม่มี job_id'};
    var sheet=SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName(JOB_LIST_SHEET);
    if (!sheet) return {ok:false,message:'ไม่พบ Job_List'};
    var data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) { if (String(data[i][1]).trim()===String(jobId).trim()) return {ok:true,job:{type:String(data[i][0]||'JM'),code:String(data[i][1]),name:String(data[i][2]),cust:String(data[i][2])}}; }
    return {ok:false,message:'ไม่พบ Job: '+jobId};
  } catch(e) { return {ok:false,message:e.toString()}; }
}

// ============================================================
// TEST FUNCTIONS
// ============================================================
function testPing() { Logger.log('PING: '+JSON.stringify(dispatch({action:'PING'}))); }
function testCheckIn() { Logger.log('checkIn: '+JSON.stringify(doCheckIn({emp_id:'EMP001',emp_name:'สมชาย',emp_dept:'ฝ่ายผลิต',job_id:'T-0042',job_name:'แท็งค์ทดสอบ',job_zone:'ลาน B',process:'Fabrication',lang:'th'}))); }
function testCheckOut() { Logger.log('checkOut: '+JSON.stringify(doCheckOut({emp_id:'EMP001',job_id:'T-0042'}))); }
function testImportEmployee() { importEmployeeData(); }
function testInstallTrigger() { installTrigger(); }
function testPayrollPeriod() { Logger.log('Tab: '+getPayrollSheetName()); }
// ── v4.0 TEST: ตรวจ time-band segmentation (ไม่แตะชีตจริง) ──
function _segTest(timeIn, timeOut, kindMap) {
  var inMin = toMinutes(timeIn), outMin = toMinutes(timeOut);
  if (outMin < inMin) outMin += 1440;
  var segs = segmentByBands(inMin, outMin, function (off) { return kindMap[off] || kindMap[kindMap.length-1]; });
  var by = { '1':0,'2A':0,'2B':0,'3':0,'4':0 };
  segs.forEach(function (s) { by[s.code] += s.mins; });
  var out = [];
  ['1','2A','2B','3','4'].forEach(function (c) { if (by[c] > 0) out.push(c + '=' + minutesToDecimal(by[c])); });
  return out.join('  ');
}
function testHourEngine() {
  Logger.log('— v4.0 Hour Engine Tests —');
  Logger.log('A 08:00-12:00 ปกติ   -> ' + _segTest('08:00','12:00',['ปกติ']));
  Logger.log('B 13:00-22:00 ปกติ   -> ' + _segTest('13:00','22:00',['ปกติ']));
  Logger.log('C 12:00-13:00 ปกติ   -> ' + _segTest('12:00','13:00',['ปกติ']));
  Logger.log('D 08:00-17:00 ปกติ   -> ' + _segTest('08:00','17:00',['ปกติ']));
  Logger.log('E 08:00-17:00 หยุด    -> ' + _segTest('08:00','17:00',['หยุด']));
  Logger.log('F 08:00-17:00 นักขัต  -> ' + _segTest('08:00','17:00',['นักขัตฤกษ์']));
  Logger.log('G 17:00-19:30 หยุด    -> ' + _segTest('17:00','19:30',['หยุด']));
  Logger.log('H เสาร์ 22:00->02:00   -> ' + _segTest('22:00','02:00',['ปกติ','หยุด']));
  Logger.log('I อาทิตย์ 22:00->06:00 -> ' + _segTest('22:00','06:00',['หยุด','ปกติ']));
}
function testMultiplier() {
  ['พนักงานรายวัน','พนักงานรายเดือน'].forEach(function (t) {
    ['1','2A','2B','3','4'].forEach(function (c) {
      Logger.log(t + ' Code ' + c + ' (1/7/2569) -> x' + getMultiplier(t, c, '1/7/2569'));
    });
  });
}
