/***********************************************************
 * STT NOVA-HR Hub — Auth (Login / Register ด้วยรหัสพนักงาน + PIN)
 * ---------------------------------------------------------
 * เร็วด้วย CacheService: อ่าน Employee_List ครั้งเดียว cache 6 ชม.
 * ล็อกอินครั้งต่อไปดึงจาก cache = ไม่ต้องอ่านชีต = เร็วมาก
 * ใช้ DB_ID (ไฟล์ Employee & Account) จาก Code.gs
 ***********************************************************/

var EMP_SHEET = 'Employee_List';
// คอลัมน์ Employee_List (0-based): A รหัส, B ชื่อ, C แผนก, D LINE ID, ... J Department(9),
//   K PIN(10), L Role(11), M บริษัท(12), N Device(13), O สถานะ(14)

/** โหลด map พนักงาน (cache) → { code: {name,dept,line,deptMain,pin,role,company,status} } */
function getEmpMap_(force) {
  var cache = CacheService.getScriptCache();
  if (!force) { var c = cache.get('empmap'); if (c) return JSON.parse(c); }
  var sh = SpreadsheetApp.openById(DB_ID).getSheetByName(EMP_SHEET);
  if (!sh) throw new Error('ไม่พบแท็บ ' + EMP_SHEET);
  var v = sh.getDataRange().getValues(), map = {};
  for (var i = 1; i < v.length; i++) {
    var code = String(v[i][0]).trim(); if (!code) continue;
    map[code] = {
      name: String(v[i][1] || ''), dept: String(v[i][2] || ''), line: String(v[i][3] || ''),
      deptMain: String(v[i][9] || ''), pin: String(v[i][10] || '').trim(),
      role: String(v[i][11] || '').trim(), company: String(v[i][12] || 'STT').trim(),
      status: String(v[i][14] || '').trim()
    };
  }
  try { cache.put('empmap', JSON.stringify(map), 21600); } catch (e) {}   // 6 ชม. (ถ้าใหญ่เกิน 100KB จะข้าม — ยังทำงานได้แค่ช้าลงนิด)
  return map;
}

/** ล้าง cache (เรียกหลังแก้ข้อมูลพนักงาน / sync) */
function clearEmpCache() { try { CacheService.getScriptCache().remove('empmap'); } catch (e) {} }

/** ล็อกอิน: รหัสพนักงาน + PIN → คืนโปรไฟล์ + Role */
function hubLogin(code, pin) {
  code = String(code || '').trim(); pin = String(pin || '').trim();
  if (!code || !pin) return { ok: false, message: 'กรอกรหัสพนักงานและ PIN' };
  var e = getEmpMap_(false)[code];
  if (!e) return { ok: false, message: 'ไม่พบรหัสพนักงานนี้' };
  if (e.status.indexOf('ลาออก') >= 0) return { ok: false, message: 'บัญชีนี้ปิดใช้งานแล้ว' };
  if (!e.pin) return { ok: false, needRegister: true, message: 'ยังไม่ได้ลงทะเบียน — กดลงทะเบียนเพื่อรับ PIN' };
  if (e.pin !== pin) return { ok: false, message: 'PIN ไม่ถูกต้อง' };
  return { ok: true, code: code, name: e.name, role: e.role || 'พนักงาน', company: e.company || 'STT', dept: e.dept, deptMain: e.deptMain };
}

/** ลงทะเบียนครั้งแรกด้วยรหัสพนักงาน → ระบบออก PIN 6 หลัก (โชว์ให้จำ) */
function hubRegister(code) {
  code = String(code || '').trim();
  if (!code) return { ok: false, message: 'กรอกรหัสพนักงาน' };
  var sh = SpreadsheetApp.openById(DB_ID).getSheetByName(EMP_SHEET);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === code) {
      var pin = String(v[i][10] || '').trim();
      if (pin) return { ok: false, already: true, message: 'รหัสนี้ลงทะเบียนแล้ว — ถ้าลืม PIN ติดต่อ HR' };
      pin = ('00000' + Math.floor(Math.random() * 1000000)).slice(-6);
      sh.getRange(i + 1, 11).setNumberFormat('@STRING@').setValue(pin);          // K PIN
      if (!String(v[i][11] || '').trim()) sh.getRange(i + 1, 12).setValue('พนักงาน'); // L Role default
      if (!String(v[i][12] || '').trim()) sh.getRange(i + 1, 13).setValue('STT');     // M บริษัท
      if (!String(v[i][14] || '').trim()) sh.getRange(i + 1, 15).setValue('Active');  // O สถานะ
      clearEmpCache();
      return { ok: true, code: code, name: String(v[i][1] || ''), pin: pin };
    }
  }
  return { ok: false, message: 'ไม่พบรหัสพนักงาน ' + code };
}
