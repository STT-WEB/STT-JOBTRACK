/**
 * ============================================================================
 *  NOVA-HR · Payload_Emp.gs — ทะเบียนพนักงาน + แผนก + HR&ACC
 *  Salary Master → emps · หน่วยงาน/แผนก → depts · แท็บ HR&ACC → Dashboard ผู้บริหาร
 * ============================================================================
 */

/* ============================================================================
   1) emps — ทะเบียนพนักงาน (แท็บ Salary Master)
   ============================================================================ */
function buildEmps_() {
  var t = readTab_(FILES.PAYROLLDB, 'Salary Master');
  var out = { STT: [], S1: [] };
  t.rows.forEach(function (r) {
    var id = String(pick_(r, ['รหัสพนักงาน'], '')).trim();
    if (!id) return;
    var typeRaw = String(pick_(r, ['ปรเภทพนักงาน (มีสูตรแล้วลากลง)', 'ประเภทพนักงาน'], ''));
    var di = String(pick_(r, ['Direct/Indirect (จำนวนคน)'], ''));
    var e = {
      id: id,
      name: String(pick_(r, ['ชื่อพนักงาน'], '')),
      /* โครงจริง: หน่วยงาน (1000/3000/4000/5000...) = แผนกใหญ่ · รหัสแผนก (1001/5013/5018) = แผนกย่อย
         KEMREX 5018 อยู่ใต้ PRODUCTION ในไฟล์ แต่เบียร์ต้องการเห็นแยก → ยกขึ้นเป็นแผนกใหญ่ */
      /* บางแถวไม่ได้กรอก "รหัสหน่วยงาน" → เดาจากรหัสแผนก (5013 → 5000) */
      dept: (CFG.KEMREX_SPLIT && nvNum_(pick_(r, ['รหัสแผนก'], 0)) === CFG.KEMREX)
              ? CFG.KEMREX
              : (nvNum_(pick_(r, ['รหัสหน่วยงาน'], 0)) || Math.floor(nvNum_(pick_(r, ['รหัสแผนก'], 0)) / 1000) * 1000),
      deptName: (CFG.KEMREX_SPLIT && nvNum_(pick_(r, ['รหัสแผนก'], 0)) === CFG.KEMREX)
              ? 'KEMREX' : String(pick_(r, ['ชื่อหน่วยงาน'], '')),
      subCode: nvNum_(pick_(r, ['รหัสแผนก'], 0)),
      sub: String(pick_(r, ['แผนก'], '')),
      direct: di.indexOf('Direct') === 0,
      type: typeRaw.indexOf('รายวัน') >= 0 ? 'รายวัน' : 'รายเดือน',
      rate: q2_(pick_(r, ['อัตรา'], 0)),
      salary: q2_(pick_(r, ['เงินเดือน'], 0)),
      benefit: q2_(pick_(r, ['**รวมสวัสดิการ 2568', '**รวมสวัสดิการ'], 0)),
      hireM: monthOf_(pick_(r, ['วันเริ่มงาน'], '')) || 1,
      leaveM: String(pick_(r, ['สถานะพนักงาน'], '')).indexOf('In Active') >= 0
        ? (monthOf_(pick_(r, ['วันที่พ้นสภาพ'], '')) || CFG.NMONTH) : CFG.NMONTH
    };
    // TODO: แยกบริษัทด้วยกฎที่เบียร์กำหนด (ตอนนี้ทุกคนอยู่ไฟล์เดียว = STT)
    out.STT.push(e);
  });
  return out;
}

/* ============================================================================
   depts — หน่วยงาน (แผนกใหญ่) + แผนกย่อย
   ============================================================================ */
function buildDepts_(emps) {
  var by = {};
  emps.STT.concat(emps.S1 || []).forEach(function (e) {
    var d = by[e.dept] || (by[e.dept] = { code: e.dept, name: e.deptName || String(e.dept), dir: 0, ind: 0, subs: {} });
    if (e.direct) d.dir++; else d.ind++;
    if (e.sub) d.subs[e.sub] = 1;
  });
  return Object.keys(by).map(function (k) {
    var d = by[k];
    return { code: d.code, name: d.name, direct: d.dir >= d.ind, subs: Object.keys(d.subs).sort() };
  }).sort(function (a, b) { return a.code - b.code; });
}

/* ============================================================================
   2) hracc — Dashboard ผู้บริหาร (แท็บ HR&ACC · HR กรอกเอง)
   ============================================================================ */
function buildHrAcc_() {
  var sh = SpreadsheetApp.openById(FILES.PAYACTUAL).getSheetByName('HR&ACC');
  if (!sh) throw new Error('ไม่พบแท็บ HR&ACC');
  var v = sh.getDataRange().getValues();
  // โครงจริง: 2 แถวหัว (แถวบน = ปี · แถวล่าง = ชื่อค่า) แล้วตามด้วย 12 เดือน
  var hTop = -1;
  for (var i = 0; i < Math.min(v.length, 15); i++) {
    if (String(v[i]).indexOf('จำนวนพนงที่จ่าย') >= 0) { hTop = i - 1; break; }
  }
  if (hTop < 0) throw new Error('อ่านหัวตาราง HR&ACC ไม่ได้ — โครงไฟล์เปลี่ยน');
  var yearRow = v[hTop], nameRow = v[hTop + 1];
  var col = {};                                  // col['2569']['ยอดจ่ายสุทธิ'] = index
  var curYear = '';
  for (var c = 0; c < nameRow.length; c++) {
    var y = String(yearRow[c]).replace(/\[merged\]\s*/g, '').trim();
    if (/^25\d\d/.test(y)) curYear = y.replace(/\s*\((HR|ACC)\)/, '') + (/\(ACC\)/.test(y) ? '|ACC' : '');
    var nm = String(nameRow[c]).trim();
    if (!nm) continue;
    var key = curYear.split('|')[0];
    if (!key) continue;
    col[key] = col[key] || {};
    var suffix = (curYear.indexOf('|ACC') >= 0) ? ' (ACC)' : '';
    col[key][nm + suffix] = c;
  }
  var out = {};
  ['2567', '2568', '2569'].forEach(function (y) {
    var arr = [];
    for (var m = 1; m <= 12; m++) {
      var row = null;
      for (var i = hTop + 2; i < v.length; i++) {
        if (monthOf_(v[i][0]) === m) { row = v[i]; break; }
      }
      var C = col[y] || {};
      var get = function (n) { return C[n] === undefined ? 0 : q2_(row[C[n]]); };
      if (!row || (y === '2569' && m > CFG.NMONTH)) { arr.push(null); continue; }
      var bank = get('ยอดจ่ายธนาคาร'), acc = get('ยอดตัดธนาคาร');
      arr.push({
        m: m, heads: nvNum_(row[C['จำนวนพนงที่จ่าย']]),
        salary: get('เงินเดือน'), benefit: get('รวมสวัสดิการ'),
        comp: get('เงินชดเชย'), ot: get('OT'), net: get('ยอดจ่ายสุทธิ'),
        cash: get('ยอดจ่ายเงินสด'), bank: bank, accBank: acc,
        match: Math.abs(bank - acc) < 0.005
      });
    }
    out[y] = arr;
  });
  return out;
}

/**
 * รวมทะเบียนพนักงานกับสิ่งที่ "ไฟล์ Cal" บอก
 * ไฟล์ Cal คือความจริงของการเดินงวด — Salary Master อาจอัปเดตไม่ทัน/คนลาออกถูกลบทิ้ง
 * ทำ 2 อย่าง: ① ทับ หน่วยงาน/แผนก/Direct-Indirect ด้วยของจริง  ② เติมคนที่ไม่มีในทะเบียน
 * คืนค่า = จำนวนคนที่เติมเข้ามา + รายชื่อคนที่เปลี่ยน Direct↔Indirect ระหว่างปี
 */
function mergeEmpInfo_(emps, empMap, empInfo) {
  var added = 0;
  Object.keys(empInfo).forEach(function (id) {
    var c = empInfo[id], e = empMap[id];
    if (!e) {                                        // มีใน Cal แต่ไม่มีในทะเบียน (เช่น ลาออกแล้วถูกลบ)
      e = { id: id, name: c.name || id, rate: 0, salary: 0, benefit: 0,
            hireM: 1, leaveM: CFG.NMONTH, subCode: 0 };
      emps.STT.push(e); empMap[id] = e; added++;
    }
    var isKem = (c.sub === 'KEMREX') || (e.subCode === CFG.KEMREX);
    e.direct   = !!c.direct;
    e.type     = c.type || e.type || 'รายเดือน';
    e.name     = e.name || c.name || id;
    e.sub      = c.sub || e.sub || '';
    e.dept     = (CFG.KEMREX_SPLIT && isKem) ? CFG.KEMREX : (c.dept || e.dept || 0);
    e.deptName = (CFG.KEMREX_SPLIT && isKem) ? 'KEMREX'   : (c.deptName || e.deptName || String(e.dept));
  });
  return added;
}
