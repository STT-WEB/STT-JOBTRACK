/**
 * ============================================================================
 *  NOVA-HR · Payload_Cal.gs — อ่านไฟล์ Cal รายเดือน (หัวใจของตัวเลขทั้งหมด)
 *  สรุปตารางเงินเดือน · PAYROLL_SUMMARY · JOB_COST_DIRECT · PERFORMANCE · ปฏิทินวันทำงาน
 * ============================================================================
 */

function readAllCalMonths_(empMap) {
  var files = calFiles_();
  if (!files.length) throw new Error('ไม่พบไฟล์ Cal ในโฟลเดอร์ ' + FILES.CAL_DIR);
  CFG.NMONTH = files[files.length - 1].m;          // เดือนล่าสุดที่มีไฟล์จริง

  var rows = [], jobRows = [], perf = [], hourTypes = [], procCost = {};
  var empInfo = {};                       // ทะเบียนพนักงานตามที่ "ไฟล์ Cal" บอก (เดือนล่าสุดชนะ)
  var warn = [];                          // เดือนไหนอ่านอะไรไม่ได้ ต้องดังออกมา ไม่ใช่เงียบแล้วได้ 0
  var workdays = [], holidays = [];
  var seenHT = {};

  files.forEach(function (fi) {
    var m = fi.m, ss = SpreadsheetApp.openById(fi.id);

    /* ── ด่านที่ 0: แท็บที่ต้องใช้มีครบไหม ──
       ถ้าแท็บหาย โค้ดจะอ่านได้ 0 แถวแบบเงียบๆ แล้วเดือนนั้นจะหายไปจากรายงานโดยไม่มีใครรู้
       จึงต้องบันทึกไว้และเอาไปขึ้นธงแดงบนหน้าจอ */
    [['JOB_COST_DIRECT', ['JOB_COST_DIRECT']],
     ['PAYROLL_SUMMARY', ['PAYROLL_SUMMARY']],
     ['PERFORMANCE',     ['PERFORMANCE']],
     ['PAYROLL_ACTUAL',  ['PAYROLL_ACTUAL', 'สรุปตารางเงินเดือน']],
     ['CALENDAR_MASTER', ['CALENDAR_MASTER', 'ปฏิทินวันทำงาน']],
     ['HOUR_TYPE_RULE',  ['HOUR_TYPE_RULE', 'WORK_HOUR_TYPE']]
    ].forEach(function (t) {
      if (!findTab_(ss, t[1])) warn.push(MONTHS_TH[m - 1] + ': ไม่พบแท็บ ' + t[0] + ' ในไฟล์ "' + fi.name + '"');
    });

    /* --- 1) JOB_COST_DIRECT → รวมยอดต่อ คน × จ๊อบ × process --- */
    var jd = nvReadSheet_(findTab_(ss, ['JOB_COST_DIRECT']));
    var byKey = {}, byEmpH = {};
    jd.rows.forEach(function (r) {
      var id = String(pick_(r, ['รหัสพนักงาน'], '')).trim();
      var job = String(pick_(r, ['JOB CODE'], '')).trim();
      if (!id || !job) return;
      var proc = String(pick_(r, ['Process'], '')).trim() || '(ไม่ระบุ)';
      var t = htCode_(pick_(r, ['ประเภทชั่วโมง'], '1'));
      var h = nvNum_(pick_(r, ['จำนวนชั่วโมง'], 0));
      var c = nvNum_(pick_(r, ['ต้นทุนค่าแรง'], 0));
      var k = id + '|' + job + '|' + proc;
      var o = byKey[k] || (byKey[k] = { m: m, id: id, job: job, proc: proc, hn: 0, ot15: 0, ot30: 0, hhol: 0, th: 0, cost: 0 });
      var e = byEmpH[id] || (byEmpH[id] = { hn: 0, ot15: 0, ot30: 0, hhol: 0, th: 0, cost: 0 });
      var f = (t === 2) ? 'hhol' : (t === 3) ? 'ot15' : (t === 4) ? 'ot30' : 'hn';
      o[f] += h; o.th += h; o.cost = q2_(o.cost + c);
      e[f] += h; e.th += h; e.cost = q2_(e.cost + c);
      procCost[proc] = q2_((procCost[proc] || 0) + c);
    });
    Object.keys(byKey).forEach(function (k) { jobRows.push(byKey[k]); });

    /* --- 2) สรุปตารางเงินเดือน → วัน/OT รายก้อน --- */
    var pay = nvReadSheet_(findTab_(ss, ['PAYROLL_ACTUAL', 'สรุปตารางเงินเดือน']));
    var payBy = {};
    pay.rows.forEach(function (r) {
      var mo = nvNum_(pick_(r, ['เลขเดือน'], 0)) || monthOf_(pick_(r, ['เดือน'], ''));
      if (mo !== m) return;
      var id = String(pick_(r, ['รหัสพนักงาน'], '')).trim();
      if (!id) return;
      payBy[id] = {
        days: nvNum_(pick_(r, ['จำนวนวันทำงาน'], 0)),
        ot1: q2_(pick_(r, ['OTx1'], 0)), ot15: q2_(pick_(r, ['OTx1.5'], 0)),
        othol: q2_(pick_(r, ['OTx2'], 0)), ot3: q2_(pick_(r, ['OTx3'], 0))
      };
    });

    /* --- 3) PERFORMANCE → ชั่วโมงมาตรฐาน --- */
    var pf = nvReadSheet_(findTab_(ss, ['PERFORMANCE']));
    var pfBy = {};
    pf.rows.forEach(function (r) {
      var id = String(pick_(r, ['รหัสพนักงาน'], '')).trim();
      if (!id) return;
      empInfo[id] = empInfo[id] || {};
      empInfo[id].dept     = nvNum_(pick_(r, ['รหัสหน่วยงาน'], 0)) || empInfo[id].dept || 0;
      empInfo[id].deptName = String(pick_(r, ['หน่วยงาน'], '')) || empInfo[id].deptName || '';
      empInfo[id].sub      = String(pick_(r, ['แผนก'], '')) || empInfo[id].sub || '';
      empInfo[id].name     = String(pick_(r, ['ชื่อพนักงาน'], '')) || empInfo[id].name || '';
      pfBy[id] = {
        stdH: nvNum_(pick_(r, ['ชม. มาตรฐาน', 'ชม.มาตรฐาน'], 0)),
        holH: nvNum_(pick_(r, ['ชม. หยุดนักขัตฤกษ์', 'ชม.หยุดนักขัตฤกษ์'], 0)),
        netH: nvNum_(pick_(r, ['ชม. มาตรฐานสุทธิ', 'ชม.มาตรฐานสุทธิ'], 0)),
        hn: nvNum_(pick_(r, ['ชม. ปกติจาก Timesheet'], 0)),
        hhol: nvNum_(pick_(r, ['ชม. ทำงานในวันหยุด'], 0)),
        ot: nvNum_(pick_(r, ['ชม. OT รวม'], 0)),
        th: nvNum_(pick_(r, ['ชม. รวมทั้งหมด'], 0)),
        absentH: nvNum_(pick_(r, ['ชม. ขาด/ลาจริง (กระทบ Performance)', 'ชม. ขาด/ลาจริง'], 0)),
        pay: q2_(pick_(r, ['(เงินเดือน+OT+สวัสดิการ)'], 0))
      };
    });

    /* --- 4) PAYROLL_SUMMARY → แหล่งความจริงของยอดเงิน --- */
    var ps = nvReadSheet_(findTab_(ss, ['PAYROLL_SUMMARY']));
    ps.rows.forEach(function (r) {
      var id = String(pick_(r, ['รหัสพนักงาน'], '')).trim();
      if (!id) return;
      var dir = String(pick_(r, ['Direct/Indirect'], '')).indexOf('Direct') === 0;
      empInfo[id] = empInfo[id] || {};
      empInfo[id].direct = dir;                                        // เดือนหลังทับเดือนก่อน = ได้สถานะล่าสุด
      empInfo[id].type   = String(pick_(r, ['Employee Type (รายวัน/รายเดือน)', 'Employee Type'], '')).indexOf('รายวัน') >= 0 ? 'รายวัน' : 'รายเดือน';
      empInfo[id].name   = empInfo[id].name || String(pick_(r, ['ชื่อพนักงาน'], ''));
      var bp = q2_(pick_(r, ['(เงินเดือน+OT+สวัสดิการ)'], 0));
      var base = q2_(pick_(r, ['(เงินเดือน+สวัสดิการ ไม่รวม OT)'], 0));
      var pot = q2_(pick_(r, ['Total OT'], 0));
      if (Math.abs(q2_(base + pot) - bp) > 0.05) base = q2_(bp - pot);   // บังคับด่านตรวจข้อ 1
      var P = payBy[id] || { days: 0, ot1: 0, ot15: 0, othol: 0, ot3: 0 };
      var F = pfBy[id] || {};
      var H = byEmpH[id];
      var absentH = nvNum_(pick_(r, ['ชม. ขาด/ลาจริง (กระทบ Performance)', 'ชม. ขาด/ลาจริง'], F.absentH || 0));
      var phol = P.othol, p30 = P.ot3, p15 = q2_(pot - phol - p30);      // OTx1 ถูกกลืนเข้า p15 เพื่อให้รวมเท่า Total OT
      rows.push({
        m: m, id: id, dir: dir,
        days: P.days,
        absentD: Math.round(absentH / 8 * 100) / 100,
        hn:    H ? H.hn    : nvNum_(pick_(r, ['ชม. ปกติจาก Timesheet'], 0)),
        ot15:  H ? H.ot15  : 0,
        ot30:  H ? H.ot30  : 0,
        hhol:  H ? H.hhol  : nvNum_(pick_(r, ['ชม. ทำงานในวันหยุด'], 0)),
        th:    H ? H.th    : nvNum_(pick_(r, ['ชม. รวมที่ทำงานจริง'], 0)),
        otr: q2_(pick_(r, ['OT Rate'], 0)),
        base: base, p15: p15, p30: p30, phol: phol, pot: pot, bp: bp,
        stdH: F.stdH || nvNum_(pick_(r, ['ชม. ที่ควรทำงานจริง (รายวัน)'], 0)),
        holH: F.holH || nvNum_(pick_(r, ['ชม. หยุดนักขัตฤกษ์'], 0)),
        netH: F.netH || 0,
        absentH: absentH
      });
      /* perf = เฉพาะพนักงาน Direct (ตามที่เบียร์สั่ง) และ "คิดเกรดใหม่" */
      if (dir && F.netH) {
        var pct = q2_(F.hn / F.netH * 100);
        var th = F.th || (H ? H.th : 0);
        perf.push({
          m: m, id: id, stdH: F.stdH, holH: F.holH, netH: F.netH,
          hn: F.hn, hhol: F.hhol, ot: F.ot, th: th, absentH: F.absentH,
          pct: pct, grade: gradeOf_(pct),
          pay: F.pay || bp, cph: th ? q2_((F.pay || bp) / th) : 0
        });
      }
    });

    /* --- 5) ปฏิทินวันทำงาน --- */
    var cal = nvReadSheet_(findTab_(ss, ['CALENDAR_MASTER', 'ปฏิทินวันทำงาน']));
    var wd = 0, hd = 0;
    cal.rows.forEach(function (r) {
      var pk = String(pick_(r, ['PeriodKey_รายเดือน'], ''));
      var mm = pk.match(/-(\d{2})$/);
      if (!mm || +mm[1] !== m) return;
      var t = String(pick_(r, ['ประเภทวันทำงาน'], ''));
      if (t.indexOf('ปกติ') >= 0) wd++; else if (t.indexOf('หยุด') >= 0) hd++;
    });
    workdays.push(wd); holidays.push(hd);
    if (!wd) warn.push(MONTHS_TH[m - 1] + ': ปฏิทินไม่มีวันทำงานเลย — เช็ก PeriodKey_รายเดือน ในแท็บ CALENDAR_MASTER');

    /* --- 6) WORK_HOUR_TYPE (อ่านครั้งเดียวพอ) --- */
    if (!hourTypes.length) {
      var wh = nvReadSheet_(findTab_(ss, ['HOUR_TYPE_RULE', 'WORK_HOUR_TYPE']));
      wh.rows.forEach(function (r) {
        var code = htCode_(pick_(r, ['Work Hour Type Name'], pick_(r, ['Work Hour Type Code'], '')));
        if (seenHT[code]) return; seenHT[code] = 1;
        hourTypes.push({ code: code, name: String(pick_(r, ['Work Hour Type Name'], '')), mult: nvNum_(pick_(r, ['Multiplier'], 1)) });
      });
      hourTypes.sort(function (a, b) { return a.code - b.code; });
    }
  });

  /* --- 7) ยุบ Process ให้เหลือ 7 อันดับแรก + "อื่นๆ" (จานสีมี 8 ช่อง เท่ากับ Mockup) --- */
  var order = Object.keys(procCost).sort(function (a, b) { return procCost[b] - procCost[a]; });
  var map = {}, processes = [];
  order.forEach(function (nm, i) {
    if (i < CFG.NPROC) { var code = 'P' + (i + 1); map[nm] = code; processes.push({ code: code, name: nm }); }
    else map[nm] = 'PX';
  });
  if (order.length > CFG.NPROC) processes.push({ code: 'PX', name: 'อื่นๆ (' + (order.length - CFG.NPROC) + ' process)' });
  jobRows.forEach(function (x) { x.proc = map[x.proc] || 'PX'; });

  /* รวมซ้ำหลังยุบรหัส + สร้าง procRows */
  var jr = {}, pr = {};
  jobRows.forEach(function (x) {
    var k = x.m + '|' + x.id + '|' + x.job + '|' + x.proc;
    var o = jr[k] || (jr[k] = { m: x.m, id: x.id, job: x.job, proc: x.proc, hn: 0, ot15: 0, ot30: 0, hhol: 0, th: 0, cost: 0 });
    o.hn += x.hn; o.ot15 += x.ot15; o.ot30 += x.ot30; o.hhol += x.hhol; o.th += x.th; o.cost = q2_(o.cost + x.cost);
    var pk = x.job + '|' + x.proc;
    var p = pr[pk] || (pr[pk] = { job: x.job, proc: x.proc, hn: 0, ot: 0, hhol: 0, th: 0, cost: 0, ids: {} });
    p.hn += x.hn; p.ot += x.ot15 + x.ot30; p.hhol += x.hhol; p.th += x.th; p.cost = q2_(p.cost + x.cost); p.ids[x.id] = 1;
  });
  var jobRows2 = Object.keys(jr).map(function (k) { return jr[k]; });
  var procRows = Object.keys(pr).map(function (k) {
    var p = pr[k], heads = Object.keys(p.ids).length;
    return { job: p.job, proc: p.proc, hn: p.hn, ot: p.ot, hhol: p.hhol, th: p.th, cost: p.cost, heads: heads, cph: p.th ? q2_(p.cost / p.th) : 0 };
  });

  /* เดือนไหนไม่มีแถวเงินเดือนเลย = อ่านไม่ได้จริงๆ */
  for (var mm = 1; mm <= CFG.NMONTH; mm++) {
    var n = 0;
    rows.forEach(function (r) { if (r.m === mm) n++; });
    if (!n) warn.push(MONTHS_TH[mm - 1] + ': ไม่มีแถวเงินเดือนเลย — อ่านแท็บ PAYROLL_SUMMARY ไม่ได้');
  }

  return {
    empInfo: empInfo, warn: warn,
    rows: rows, jobRows: jobRows2, perf: perf, procRows: procRows,
    processes: processes, hourTypes: hourTypes,
    workdays: workdays, holidays: holidays
  };
}
