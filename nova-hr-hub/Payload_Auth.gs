/**
 * ============================================================================
 *  NOVA-HR · Payload_Auth.gs — ประตูเข้าออกของข้อมูล
 *  ทุกอย่างที่หน้าเว็บเรียกได้ ต้องผ่านไฟล์นี้ไฟล์เดียว
 *
 *  กฎเหล็ก 3 ข้อ (ห้ามแก้ให้หลวมกว่านี้):
 *    ① หน้าเว็บที่เสิร์ฟออกไป ไม่มีข้อมูลเงินเดือนติดไปเลยแม้แต่ตัวเดียว
 *    ② จะได้ข้อมูลต้องมี "โทเคน" ซึ่งออกให้เฉพาะคนที่ใส่รหัสพนักงาน + PIN ถูก
 *    ③ เซิร์ฟเวอร์ตัดข้อมูลตามสิทธิ์ก่อนส่ง — คนที่ไม่มีสิทธิ์ จะไม่มีข้อมูลนั้นอยู่ในเครื่องเลย
 * ============================================================================
 */

var NV_SESS_PREFIX = 'NVSESS_';
var NV_SESS_SEC    = 21600;      // อายุเซสชัน 6 ชั่วโมง
var NV_FAIL_MAX    = 5;          // ใส่ PIN ผิดเกินนี้ → ล็อก 15 นาที
var NV_FAIL_SEC    = 900;

/* ---------------------------------------------------------------- เซสชัน */
function nvPutSess_(p) {
  var t = Utilities.getUuid();
  CacheService.getScriptCache().put(NV_SESS_PREFIX + t, JSON.stringify(p), NV_SESS_SEC);
  return t;
}
function nvGetSess_(token) {
  if (!token) return null;
  var s = CacheService.getScriptCache().get(NV_SESS_PREFIX + String(token));
  if (!s) return null;
  try { return JSON.parse(s); } catch (e) { return null; }
}
function nvTouchSess_(token, p) {
  CacheService.getScriptCache().put(NV_SESS_PREFIX + String(token), JSON.stringify(p), NV_SESS_SEC);
}

/* ------------------------------------------------- กันเดา PIN (brute force) */
function nvFailKey_(code) { return 'NVFAIL_' + String(code); }
function nvFailCount_(code) { return Number(CacheService.getScriptCache().get(nvFailKey_(code)) || 0); }
function nvFailAdd_(code) {
  var n = nvFailCount_(code) + 1;
  CacheService.getScriptCache().put(nvFailKey_(code), String(n), NV_FAIL_SEC);
  return n;
}
function nvFailClear_(code) { CacheService.getScriptCache().remove(nvFailKey_(code)); }

/* ================================================================ ที่หน้าเว็บเรียก */

/** เวอร์ชัน — เรียกได้โดยไม่ต้องล็อกอิน (ใช้โชว์ป้าย build บนหน้า Login) */
function nvVersion() {
  return { ver: CFG.VERSION, build: CFG.BUILD, year: CFG.YEAR };
}

/** เข้าสู่ระบบ: รหัสพนักงาน + PIN 6 หลัก → คืนโปรไฟล์ + โทเคน */
function nvLogin(code, pin) {
  code = String(code || '').trim();
  pin = String(pin || '').trim();
  if (!code || !pin) return { ok: false, message: 'กรอกรหัสพนักงานและ PIN' };

  if (nvFailCount_(code) >= NV_FAIL_MAX)
    return { ok: false, message: 'ใส่ PIN ผิดเกิน ' + NV_FAIL_MAX + ' ครั้ง — ถูกล็อก 15 นาที' };

  var r = hubLogin(code, pin);            // ใช้ตัวตรวจของเดิมใน Auth.gs (ชีต Employee_List)
  if (!r.ok) {
    if (!r.needRegister) {
      var n = nvFailAdd_(code);
      r.message = (r.message || 'เข้าสู่ระบบไม่สำเร็จ') + ' (ผิดครั้งที่ ' + n + '/' + NV_FAIL_MAX + ')';
    }
    return r;
  }
  nvFailClear_(code);
  var prof = {
    code: r.code, name: r.name, role: r.role || 'พนักงาน',
    company: r.company || 'STT', dept: r.dept || '', deptMain: r.deptMain || '',
    pin: pin, confirmed: false
  };
  return {
    ok: true, token: nvPutSess_(prof),
    code: prof.code, name: prof.name, role: prof.role, company: prof.company, dept: prof.dept
  };
}

/** ยืนยัน PIN อีกครั้งที่หน้าคีย์แพด (ด่านที่ 2 กันคนมาแอบใช้เครื่องที่เปิดค้างไว้) */
function nvConfirmPin(token, pin) {
  var s = nvGetSess_(token);
  if (!s) return { ok: false, message: 'เซสชันหมดอายุ — กรุณาเข้าสู่ระบบใหม่' };
  if (String(pin || '') !== String(s.pin)) {
    var n = nvFailAdd_(s.code);
    if (n >= NV_FAIL_MAX) { nvLogout(token); return { ok: false, message: 'ผิดเกินกำหนด — ออกจากระบบแล้ว' }; }
    return { ok: false, message: 'PIN ไม่ถูกต้อง (ผิดครั้งที่ ' + n + '/' + NV_FAIL_MAX + ')' };
  }
  nvFailClear_(s.code);
  s.confirmed = true; nvTouchSess_(token, s);
  return { ok: true };
}

/** ลงทะเบียนครั้งแรก — ระบบออก PIN 6 หลักให้ */
function nvRegister(code) { return hubRegister(code); }

function nvLogout(token) {
  if (token) CacheService.getScriptCache().remove(NV_SESS_PREFIX + String(token));
  return { ok: true };
}

/** ขอข้อมูล — ต้องมีโทเคนที่ยืนยัน PIN แล้วเท่านั้น และได้เฉพาะส่วนที่สิทธิ์ตัวเองดูได้ */
function nvData(token) {
  var s = nvGetSess_(token);
  if (!s) return JSON.stringify({ error: 'เซสชันหมดอายุ — กรุณาเข้าสู่ระบบใหม่' });
  if (!s.confirmed) return JSON.stringify({ error: 'ยังไม่ได้ยืนยัน PIN' });
  nvTouchSess_(token, s);
  try {
    return nvScope_(JSON.parse(getPayloadJson()), s);
  } catch (e) {
    return JSON.stringify({ error: 'อ่านข้อมูลต้นทางไม่สำเร็จ — ' + e.message });
  }
}

/* ================================================================ ตัดข้อมูลตามสิทธิ์ */
/**
 *  ผู้บริหาร / HR   → เห็นทุกอย่าง
 *  หัวหน้าแผนก      → เห็นเฉพาะคนในหน่วยงานตัวเอง · ไม่เห็นยอดจ่ายรวมบริษัท (HR&ACC)
 *  พนักงาน          → เห็นเฉพาะข้อมูลของตัวเอง · ไม่เห็นต้นทุนงาน/งบ/ยอดบริษัทเลย
 */
function nvScope_(D, s) {
  var role = String(s.role || 'พนักงาน');
  if (/ผู้บริหาร|HR|Admin/i.test(role)) {
    D.me = { code: s.code, name: s.name, role: role };
    return JSON.stringify(D);
  }

  var boss = /หัวหน้า|Manager|Supervisor/i.test(role);
  var myDept = null;
  ['STT', 'S1'].forEach(function (c) {
    (D.emps[c] || []).forEach(function (e) { if (e.id === s.code) myDept = e.dept; });
  });

  var keep = {};                                   // รหัสพนักงานที่คนนี้ดูได้
  ['STT', 'S1'].forEach(function (c) {
    (D.emps[c] || []).forEach(function (e) {
      if (boss ? (myDept !== null && e.dept === myDept) : (e.id === s.code)) keep[e.id] = 1;
    });
  });
  keep[s.code] = 1;

  ['STT', 'S1'].forEach(function (c) {
    D.emps[c] = (D.emps[c] || []).filter(function (e) { return keep[e.id]; });
    var co = D.companies[c]; if (!co) return;
    co.rows  = (co.rows  || []).filter(function (r) { return keep[r.id]; });
    co.perf  = (co.perf  || []).filter(function (r) { return keep[r.id]; });
    co.recon = (co.recon || []).filter(function (r) { return keep[r.id]; });
    if (boss) {
      co.jobRows = (co.jobRows || []).filter(function (x) { return keep[x.id]; });
    } else {
      co.jobRows = []; co.alloc = {};
    }
  });

  D.hracc = { '2567': [], '2568': [], '2569': [] };   // ยอดจ่ายรวมบริษัท = ผู้บริหาร/HR เท่านั้น
  if (!boss) { D.jobs = []; D.procRows = []; }        // พนักงานไม่เห็นต้นทุนงาน
  D.verify = { total: 0, pass: 0, fails: [] };
  D.me = { code: s.code, name: s.name, role: role };
  return JSON.stringify(D);
}

/* ================================================================ เครื่องมือให้ HR */
/** ตั้ง PIN ใหม่ให้พนักงาน (รันจาก Apps Script Editor เท่านั้น — ไม่เปิดให้หน้าเว็บเรียก) */
function resetPinForEmployee(code, newPin) {
  var sh = SpreadsheetApp.openById(DB_ID).getSheetByName(EMP_SHEET);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === String(code).trim()) {
      var pin = String(newPin || ('00000' + Math.floor(Math.random() * 1000000)).slice(-6));
      sh.getRange(i + 1, 11).setNumberFormat('@STRING@').setValue(pin);
      nvFailClear_(code);
      Logger.log('ตั้ง PIN ใหม่ให้ ' + code + ' (' + v[i][1] + ') = ' + pin);
      return pin;
    }
  }
  Logger.log('ไม่พบรหัสพนักงาน ' + code);
  return null;
}

/** ดูว่าใครลงทะเบียนแล้วบ้าง / ใครยังไม่มี PIN */
function whoCanLogin() {
  var sh = SpreadsheetApp.openById(DB_ID).getSheetByName(EMP_SHEET);
  var v = sh.getDataRange().getValues(), yes = [], no = [];
  for (var i = 1; i < v.length; i++) {
    var code = String(v[i][0]).trim(); if (!code) continue;
    if (String(v[i][14] || '').indexOf('ลาออก') >= 0) continue;
    var line = code + ' ' + v[i][1] + ' [' + (v[i][11] || 'พนักงาน') + ']';
    if (String(v[i][10] || '').trim()) yes.push(line); else no.push(line);
  }
  Logger.log('เข้าระบบได้แล้ว ' + yes.length + ' คน:\n  ' + yes.join('\n  ') +
             '\n\nยังไม่ได้ลงทะเบียน ' + no.length + ' คน:\n  ' + no.slice(0, 40).join('\n  '));
  return { ready: yes.length, notYet: no.length };
}
