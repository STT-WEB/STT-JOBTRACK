/**
 * ============================================================================
 *  NOVA-HR · Payload_Job.gs — จ๊อบ + งบ + ปันค่าแรงทางอ้อม
 *  Master/Dashboard → jobs · Sale/Estimate Budget → งบ · Allocate Indirect ตามชั่วโมงแรงงานตรง
 * ============================================================================
 */

/* ============================================================================
   3) jobs — Master รายจ๊อบ + งบ (Master/Dashboard + Budget labour + ยอดขาย)
   ============================================================================ */
function buildJobs_(alloc) {
  var mt = readTab_(FILES.MASTER, SpreadsheetApp.openById(FILES.MASTER).getSheets()[0].getName());
  var budget = readBudget_();          // { jobCode: {est, sale, cust} }
  var wip = readAllWip_();             // { jobCode: {start, due, progress} }
  var jobs = [];
  mt.rows.forEach(function (r) {
    var code = String(pick_(r, ['Job Code', 'JOB CODE'], '')).trim();
    if (!code) return;
    var byMonth = {}, sum = 0;
    for (var m = 1; m <= 12; m++) {
      var v = q2_(pick_(r, [MONTHS_TH[m - 1]], 0));
      byMonth[String(m)] = v; sum = q2_(sum + v);
    }
    var carry = q2_(pick_(r, ['ยอดยกมา'], 0));
    var sub = q2_(pick_(r, ['ค่าแรงผู้รับเหมา'], 0));
    var ind = 0;
    for (var mm = 1; mm <= CFG.NMONTH; mm++) {
      ind = q2_(ind + ((alloc[String(mm)] || { byJob: {} }).byJob[code] || 0));
    }
    var b = budget[code] || {};
    var w = wip[code] || {};
    /* ไฟล์ Master เก็บจ๊อบสะสมทุกปี (800+ จ๊อบ) — จ๊อบที่ไม่มียอดอะไรเลยไม่ต้องเอาขึ้นจอ
       เปิดดูทั้งหมดได้ที่ CFG.SHOW_ALL_JOBS = true */
    if (!CFG.SHOW_ALL_JOBS && !carry && !sum && !ind && !sub) return;
    var total = q2_(carry + sum + ind + sub);
    /* Est. Budget = แท็บ Estimate Budget คอลัมน์ C · Sale Budget = แท็บ Sale Budget คอลัมน์ Budget labour
       จ๊อบไหนยังไม่ได้กรอก Estimate → ส่วนต่าง/% เทียบกับ Sale Budget แทน (ไม่โชว์ตัวเลขมั่ว) */
    var est = q2_(b.est || 0), sale = q2_(b.sale || 0);
    var ref = est || sale;
    jobs.push({
      code: code,
      name: String(pick_(r, ['ชื่องาน'], '')),
      cust: String(b.cust || pick_(r, ['ชื่องาน'], '')),
      type: String(pick_(r, ['ประเภทงาน'], '')),
      startM: monthOf_(w.start) || 1, dueM: monthOf_(w.due) || 12,
      start: fmtDate_(w.start), due: fmtDate_(w.due),
      carry: carry, byMonth: byMonth, sttLabor: sum, indirect: ind, sub: sub, total: total,
      estBudget: est, saleBudget: sale,
      varBudget: ref ? q2_(ref - total) : 0,
      pctBudget: ref ? q2_(total / ref * 100) : 0,
      done: String(w.progress || '').indexOf('ส่งมอบ') >= 0,
      progress: String(w.progress || 'อยู่ระหว่างผลิต'),
      leadDays: leadDays_(w.start, w.due)
    });
  });
  return jobs;
}

/**
 * งบต่อจ๊อบ — ไฟล์ STT Jobcost Database-Payroll (เบียร์ยืนยัน 20 ส.ค. 69)
 *   แท็บ "Sale Budget"      → Sale Budget = คอลัมน์ `Budget labour`
 *                             (แท็บนี้มีหลายบล็อกวางเรียงข้างกัน: JT ใช้ Budget labour ·
 *                              JM/JMI ใช้ `SUM ของ billaftrDiscAmnt` → อ่านทุกบล็อก)
 *   แท็บ "Estimate Budget"  → Est. Budget = คอลัมน์ C (ถัดจาก custName · หัวคอลัมน์ว่าง)
 * ผูกกันด้วย DocuNo = Job Code
 */
function readBudget_() {
  var ss = SpreadsheetApp.openById(FILES.PAYROLLDB), out = {};
  var put = function (code, key, val, cust) {
    code = String(code).trim(); if (!code || code === 'DocuNo') return;
    var o = out[code] || (out[code] = {});
    if (val !== '' && val !== null && val !== undefined) o[key] = q2_(val);
    if (cust) o.cust = String(cust);
  };
  /* หาแถวหัวตาราง = แถวสุดท้ายที่มีคำว่า DocuNo (บางแท็บมีแถวหัวซ้อน 2 ชั้น) */
  var headRowOf = function (v) {
    var hr = -1;
    for (var i = 0; i < Math.min(v.length, 15); i++)
      for (var c = 0; c < v[i].length; c++)
        if (String(v[i][c]).trim() === 'DocuNo') { hr = i; break; }
    return hr;
  };

  CFG._budgetWarn = [];
  /* ---- Sale Budget ---- */
  var sh = findTab_(ss, ['Sale Budget', 'SaleBudget', 'Budget labour']);
  if (!sh) CFG._budgetWarn.push('ไม่พบแท็บ "Sale Budget" ในไฟล์ Payroll DB → Sale Budget จะว่างทั้งหมด');
  if (sh) {
    var v = sh.getDataRange().getValues(), hr = headRowOf(v);
    if (hr >= 0) {
      var head = v[hr].map(function (x) { return String(x).trim(); });
      var blocks = [];
      for (var c = 0; c < head.length; c++) {
        if (head[c] !== 'DocuNo') continue;
        var cust = (head[c + 1] === 'custName') ? c + 1 : -1, val = -1;
        for (var k = c + 1; k <= c + 3 && k < head.length; k++)
          if (/Budget labour|billaftrDisc/i.test(head[k])) { val = k; break; }
        if (val >= 0) blocks.push({ doc: c, cust: cust, val: val });
      }
      for (var r = hr + 1; r < v.length; r++)
        blocks.forEach(function (b) { put(v[r][b.doc], 'sale', v[r][b.val], b.cust >= 0 ? v[r][b.cust] : ''); });
    }
  }

  /* ---- Estimate Budget (คอลัมน์ C · หัวคอลัมน์ว่าง) ---- */
  var sh2 = findTab_(ss, ['Estimate Budget', 'EstimateBudget', 'Est Budget']);
  if (!sh2) CFG._budgetWarn.push('ไม่พบแท็บ "Estimate Budget" ในไฟล์ Payroll DB → Est. Budget จะว่างทั้งหมด');
  if (sh2) {
    var v2 = sh2.getDataRange().getValues(), hr2 = headRowOf(v2);
    if (hr2 >= 0) {
      var head2 = v2[hr2].map(function (x) { return String(x).trim(); });
      var d2 = head2.indexOf('DocuNo');
      for (var r2 = hr2 + 1; r2 < v2.length; r2++)
        put(v2[r2][d2], 'est', v2[r2][d2 + 2], v2[r2][d2 + 1]);
    }
  }
  return out;
}

/** ALL WIP (วันเริ่มงาน / วันส่งมอบ / สถานะ) จากไฟล์ Cal เดือนล่าสุด */
function readAllWip_() {
  var f = latestCalFile_();
  if (!f) return {};
  var t = readTab_(f.getId(), 'ALL WIP');
  var out = {};
  t.rows.forEach(function (r) {
    var code = String(pick_(r, ['JOB CODE'], '')).trim();
    if (!code) return;
    out[code] = {
      start: pick_(r, ['START DATE ', 'START DATE'], ''),
      due: pick_(r, ['LATEST DATE DELIVERY'], ''),
      progress: pick_(r, ['TOTAL PROGRESS'], '')
    };
  });
  return out;
}

/* ============================================================================
   4) allocate Indirect เข้าจ๊อบ (ตามชั่วโมงแรงงานตรง — วิธีที่เบียร์เลือก)
   ============================================================================ */
function buildAlloc_(rows, jobRows, empMap) {
  var out = {};
  for (var m = 1; m <= CFG.NMONTH; m++) {
    var pool = 0, hours = 0, byJobH = {};
    rows.forEach(function (r) {
      if (r.m !== m) return;
      var dir = (r.dir !== undefined) ? r.dir : !!(empMap[r.id] || {}).direct;
      if (!dir) pool = q2_(pool + r.bp);
    });
    jobRows.forEach(function (x) {
      if (x.m !== m) return;
      hours += x.th; byJobH[x.job] = (byJobH[x.job] || 0) + x.th;
    });
    var keys = Object.keys(byJobH).sort(), alloc = {}, acc = 0;
    keys.forEach(function (k, i) {
      var v = (i === keys.length - 1) ? q2_(pool - acc) : q2_(byJobH[k] * pool / hours);
      alloc[k] = v; acc = q2_(acc + v);
    });
    out[String(m)] = { pool: pool, hours: hours, rate: hours ? q2_(pool / hours) : 0, byJob: alloc };
  }
  return out;
}
