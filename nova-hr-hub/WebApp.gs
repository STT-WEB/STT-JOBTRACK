/***********************************************************
 * STT NOVA-HR Hub — WebApp (โมดูลหน้าเว็บ)
 * doGet เสิร์ฟหน้า Dashboard · getMasterData ส่งข้อมูลจาก JOBCOST 2026
 * (ชั้น 5 อ่านจากชั้น 4 "ผลสำเร็จรูป" — ไม่คำนวณซ้ำ = เร็ว)
 ***********************************************************/

/* เวอร์ชันอยู่ที่เดียวคือ CFG.VERSION / CFG.BUILD ใน Payload.gs
   โชว์บนหน้า Login และมุมขวาบนของแอป → เช็กได้ทันทีว่า deploy ขึ้นหรือยัง */
function getVersion() { return 'STT NOVA-HR HUB · ' + CFG.VERSION + ' · build ' + CFG.BUILD; }
var HUB_VERSION = 'STT NOVA-HR HUB';   // เก็บไว้เผื่อโค้ดเก่าเรียกใช้

/* ---------------------------------------------------------------------------
   รวมไฟล์ย่อยเข้าหน้าเดียว — ใช้ใน Hub.html ด้วย  <?!= include('ชื่อไฟล์') ?>
   ตอนนี้ทุกไฟล์เป็น HTML/JS ธรรมดาล้วน ไม่มีตัวแปรฝังเลย → ต่อดิบๆ เร็วและปลอดภัยที่สุด
   ⚠ ห้ามเอาข้อมูลมาฝังในหน้าเว็บอีก — ข้อมูลต้องผ่าน nvData() หลังล็อกอินเท่านั้น
   --------------------------------------------------------------------------- */
function doGet() {
  /* เสิร์ฟแค่ "หน้าเปล่า" — ไม่มีข้อมูลเงินเดือนติดไปแม้แต่ตัวเดียว
     ใครเปิดลิงก์ก็เห็นแค่หน้า Login · ข้อมูลจะถูกดึงหลังใส่รหัสพนักงาน + PIN ถูก (ดู Payload_Auth.gs) */
  return HtmlService.createTemplateFromFile('Hub').evaluate()
    .setTitle('STT NOVA-HR Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** อ่านสรุปต่อจ๊อบจากแท็บ MASTER ของ JOBCOST 2026 (เร็ว: อ่านผลที่คำนวณไว้แล้ว) */
/** ทดสอบว่า backend เห็นข้อมูล Master ไหม — รันแล้วดู Log */
function testMaster() {
  var d = getMasterData();
  Logger.log('📊 MASTER: ' + d.jobs.length + ' จ๊อบ | รวม ' + Math.round(d.total).toLocaleString() + ' บาท');
  Logger.log('   อ่านจากไฟล์: ' + JOBCOST_FILE_ID);
  Logger.log('   รายเดือน: ' + d.months.map(function(x){return Math.round(x/1000)+'K';}).join(' '));
  return d;
}

function getMasterData() {
  var sh = SpreadsheetApp.openById(JOBCOST_FILE_ID).getSheetByName('MASTER');
  if (!sh) return { jobs: [], months: [0,0,0,0,0,0,0,0,0,0,0,0], total: 0, generated: new Date().toLocaleString('th-TH'), note: 'ยังไม่มีแท็บ MASTER — รัน setupMasterFile ก่อน' };
  var last = sh.getLastRow();
  var months = [0,0,0,0,0,0,0,0,0,0,0,0], total = 0, jobs = [];
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, 24).getValues();
    for (var i = 0; i < v.length; i++) {
      var mo = [];
      for (var m = 0; m < 12; m++) { var x = Number(v[i][3 + m]) || 0; mo.push(x); months[m] += x; }
      var stt = Number(v[i][16]) || 0; total += stt;
      if (String(v[i][1]).trim()) jobs.push({
        type: String(v[i][0]).trim(), code: String(v[i][1]).trim(),
        name: String(v[i][2]).trim(), months: mo, stt: stt
      });
    }
  }
  return { jobs: jobs, months: months, total: total, generated: new Date().toLocaleString('th-TH') };
}
