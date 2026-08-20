/**
 * ============================================================================
 *  NOVA-HR · Payload_Verify.gs — ด่านตรวจ 8 ข้อ
 *  ไม่ผ่าน = ขึ้นแถบแดงล่างจอ ไม่ให้ปิดงวดเงียบๆ
 * ============================================================================
 */

/* ============================================================================
   5) ด่านตรวจ 8 ข้อ — ไม่ผ่าน = ขึ้นแถบแดงล่างจอ
   ============================================================================ */
function verify_(D) {
  var fails = [], total = 0, co = D.companies.STT, EM = {};
  D.emps.STT.forEach(function (e) { EM[e.id] = e; });
  for (var m = 1; m <= CFG.NMONTH; m++) {
    (function (m) {
      total += 4;
      var bpD = 0, hD = 0, jc = 0, hJ = 0, bTot = 0, potTot = 0, bpTot = 0, ind = 0;
      co.rows.forEach(function (r) {
        if (r.m !== m) return;
        bTot += r.base; potTot += r.pot; bpTot += r.bp;
        /* ใช้ Direct/Indirect "ของเดือนนั้น" จากไฟล์ Cal ไม่ใช่สถานะปัจจุบันในทะเบียน
           (คนย้ายประเภทกลางปีได้ ถ้าใช้สถานะปัจจุบันย้อนหลังจะไม่มีวันตรง) */
        var dir = (r.dir !== undefined) ? r.dir : !!(EM[r.id] || {}).direct;
        if (dir) { bpD += r.bp; hD += r.th; } else ind += r.bp;
      });
      co.jobRows.forEach(function (x) { if (x.m === m) { jc += x.cost; hJ += x.th; } });
      if (Math.abs(q2_(bTot + potTot) - q2_(bpTot)) > 0.05) fails.push(MONTHS_TH[m - 1] + ': ค่าแรง+OT ≠ ยอดจ่าย');
      if (Math.abs(q2_(bpD) - q2_(jc)) > 0.05) fails.push(MONTHS_TH[m - 1] + ': Direct ≠ ต้นทุนลงจ๊อบ');
      if (Math.abs(hD - hJ) > 0.05) fails.push(MONTHS_TH[m - 1] + ': ชั่วโมง Direct ≠ ชั่วโมงลงจ๊อบ');
      var a = co.alloc[String(m)], s = 0;
      Object.keys(a.byJob).forEach(function (k) { s += a.byJob[k]; });
      if (Math.abs(q2_(s) - a.pool) > 0.05) fails.push(MONTHS_TH[m - 1] + ': Indirect กระจายไม่ครบ');
    })(m);
  }
  D.jobs.forEach(function (j) {
    total++;
    if (Math.abs(q2_(j.carry + j.sttLabor + j.indirect + j.sub) - j.total) > 0.05)
      fails.push(j.code + ': ยอดรวมจ๊อบไม่ลงตัว');
  });
  ['2567', '2568', '2569'].forEach(function (y) {
    (D.hracc[y] || []).forEach(function (r) {
      if (!r) return; total++;
      if (Math.abs(q2_(r.cash + r.bank) - r.net) > 0.05) fails.push(y + '/' + MONTHS_TH[r.m - 1] + ': เงินสด+ธนาคาร ≠ สุทธิ');
    });
  });
  return { total: total, pass: total - fails.length, fails: fails.slice(0, 6) };
}
