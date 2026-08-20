/***********************************************************
 * STT NOVA-HR Hub — WebApp (โมดูลหน้าเว็บ)
 * doGet เสิร์ฟหน้า Dashboard · getMasterData ส่งข้อมูลจาก JOBCOST 2026
 * (ชั้น 5 อ่านจากชั้น 4 "ผลสำเร็จรูป" — ไม่คำนวณซ้ำ = เร็ว)
 ***********************************************************/

var HUB_VERSION = 'STT NOVA-HR HUB · version 31';   // ★ บวก +1 ทุกครั้งที่แก้โค้ด — ดูบนหน้า Login ว่าตรงไหม = deploy ล่าสุดหรือยัง
function getVersion() { return HUB_VERSION; }

/* ---------------------------------------------------------------------------
   รวมไฟล์ย่อยเข้าหน้าเดียว — ใช้ใน Hub.html ด้วย  <?!= include('ชื่อไฟล์') ?>
   ไฟล์ส่วนใหญ่เป็น HTML/JS ธรรมดา → ต่อดิบๆ (เร็วและปลอดภัยที่สุด)
   มีแค่ 2 ไฟล์ที่มีตัวแปรฝังอยู่ ต้องผ่านตัวแทนค่าก่อน:
     Hub_Body  → <?!= YEAR ?>          (ป้ายปีงบมุมบน)
     Hub_Data  → <?!= PAYLOAD_JSON ?>  (ข้อมูลจริงทั้งก้อน)
   --------------------------------------------------------------------------- */
var NV_TPL = { 'Hub_Body': 1, 'Hub_Data': 1 };
var NV_CTX = {};                       // เติมค่าครั้งเดียวใน doGet (กันสร้าง payload ซ้ำ 12 รอบ)

function include(name) {
  if (!NV_TPL[name]) return HtmlService.createHtmlOutputFromFile(name).getContent();
  var t = HtmlService.createTemplateFromFile(name);
  t.YEAR = NV_CTX.YEAR;
  t.PAYLOAD_JSON = NV_CTX.PAYLOAD_JSON;
  return t.evaluate().getContent();
}

function doGet() {
  try {
    NV_CTX.YEAR = CFG.YEAR;
    NV_CTX.PAYLOAD_JSON = getPayloadJson();   // ← Payload.gs (ข้อมูลจริง) — สร้างครั้งเดียวต่อการเปิดหน้า
    var t = HtmlService.createTemplateFromFile('Hub');
    return t.evaluate()
      .setTitle('STT NOVA-HR Hub')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    /* ตาข่ายกันตก: ถ้าอ่านไฟล์ต้นทางพลาด อย่าปล่อยให้เว็บขึ้นหน้า error ดิบของ Google
       ให้บอกเป็นภาษาคนว่าติดตรงไหน แล้วบอกวิธีตรวจต่อ */
    return HtmlService.createHtmlOutput(
      '<div style="font-family:system-ui,sans-serif;background:#0B0D12;color:#E7EAF0;min-height:100vh;' +
      'display:grid;place-items:center;padding:40px;text-align:center">' +
      '<div style="max-width:620px">' +
      '<div style="font-size:44px">🔧</div>' +
      '<h2 style="margin:12px 0 6px">ระบบอ่านข้อมูลต้นทางไม่สำเร็จ</h2>' +
      '<p style="color:#9AA3B2;margin:0 0 18px">หน้าตาเว็บไม่ได้พัง — แค่ดึงตัวเลขจากไฟล์ต้นทางไม่ได้</p>' +
      '<pre style="text-align:left;background:#151922;border:1px solid #2A3140;border-radius:10px;' +
      'padding:14px;white-space:pre-wrap;color:#ff9d9d;font-size:13px">' +
      String(e && e.message ? e.message : e).replace(/</g, '&lt;') + '</pre>' +
      '<p style="color:#9AA3B2;font-size:13px;margin-top:18px">วิธีตรวจ: เปิด Apps Script Editor → เลือกฟังก์ชัน ' +
      '<b style="color:#E7EAF0">probe()</b> → กด Run → ดู Log ว่าแท็บไหนหาไม่เจอ<br>' +
      'แล้วส่ง Log ให้ Candy แก้ให้</p>' +
      '</div></div>')
      .setTitle('STT NOVA-HR Hub')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
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
