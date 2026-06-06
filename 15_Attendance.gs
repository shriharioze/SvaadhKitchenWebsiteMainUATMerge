// ════════════════════════════════════════════════════════════════════
// STAFF ATTENDANCE & SALARY  (payroll module)
//  SK_Staff      : roster + pay config (auto-seeded). Paid_Leaves = number of
//                  free (non-deducted) leaves allowed per month before the
//                  per-day deduction kicks in.
//  SK_Attendance : EXCEPTIONS only — a row exists when a day is marked ABSENT,
//                  carries an INCENTIVE (+ bonus) or a DEDUCTION (− advance/
//                  fine), and/or a Note (leave reason / memo). No row = present.
//  SK_Salary_Log : records when a salary/wage was credited (clears the alert
//                  AND forms each staff's pay history).
//  Monthly: per-day = salary / days-in-month; non-Sunday leaves BEYOND the
//  Paid_Leaves allowance each deduct one per-day. Daily (Abhijeet): weekday/Sat
//  rates, Sun off, paid weekly.
// ════════════════════════════════════════════════════════════════════
var ATT_STAFF_TAB = "SK_Staff", ATT_REC_TAB = "SK_Attendance", ATT_SAL_TAB = "SK_Salary_Log";
var ATT_STAFF_HEADERS = ["Name","Type","Monthly_Salary","Pay_Day","Pay_Cycle","Weekday_Rate","Saturday_Rate","Sunday_Rate","Paid_Leaves","Active"];
var ATT_REC_HEADERS   = ["Date","Staff_Name","Status","Incentive","Deduction","Note","Updated"];
var ATT_SAL_HEADERS   = ["Staff_Name","Period","Amount","Credited_At"];

// Append any missing header columns to an existing tab (non-destructive migration
// so older SK_Staff / SK_Attendance tabs gain Paid_Leaves / Deduction).
function _attEnsureCols(ws, headers) {
  var lastCol = ws.getLastColumn();
  var existing = lastCol > 0 ? ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
  var missing = headers.filter(function (h) { return existing.indexOf(h) === -1; });
  if (missing.length) ws.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
}
function _attSeed(ss) {
  var ws = getOrCreateTab(ss, ATT_STAFF_TAB, ATT_STAFF_HEADERS);
  _attEnsureCols(ws, ATT_STAFF_HEADERS);
  if (ws.getLastRow() > 1) return ws;
  ws.getRange(2, 1, 6, ATT_STAFF_HEADERS.length).setValues([
    ["Rupa Tai","Monthly",24000,15,"Monthly","","","",0,"Yes"],
    ["Pooja Tai","Monthly",8000,18,"Monthly","","","",0,"Yes"],
    ["Meena Tai","Monthly",5000,5,"Monthly","","","",0,"Yes"],
    ["Anita Tai","Monthly",3000,1,"Monthly","","","",0,"Yes"],
    ["Manisha Tai","Monthly",2000,1,"Monthly","","","",0,"Yes"],
    ["Abhijeet Parekar","Daily",0,"","Weekly",700,600,0,0,"Yes"]
  ]);
  return ws;
}
function _attPad(n) { return (n < 10 ? "0" : "") + n; }
function _attDateStr(v) { return v instanceof Date ? Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd") : String(v || "").trim(); }
function _attDow(ds) { return new Date(ds + "T12:00:00+05:30").getDay(); } // 0 Sun .. 6 Sat
function _attMonthShift(monthStr, delta) {
  var d = new Date(+monthStr.slice(0, 4), +monthStr.slice(5, 7) - 1 + delta, 1);
  return d.getFullYear() + "-" + _attPad(d.getMonth() + 1);
}

function getAttendanceData(month, selDate) {
  try {
    var ss = getSpreadsheet();
    var staffWs = _attSeed(ss);
    var recWs = getOrCreateTab(ss, ATT_REC_TAB, ATT_REC_HEADERS); _attEnsureCols(recWs, ATT_REC_HEADERS);
    var salWs = getOrCreateTab(ss, ATT_SAL_TAB, ATT_SAL_HEADERS);
    var staff = getAllRows(staffWs).filter(function (r) { return String(r.Name || "").trim() && String(r.Active || "Yes").toLowerCase() !== "no"; });
    var recs = getAllRows(recWs);
    var sals = getAllRows(salWs);

    var now = getISTDate();
    var today = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
    var curMonth = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM");
    var rosterDate = (selDate && /^\d{4}-\d{2}-\d{2}$/.test(selDate)) ? selDate : today;
    if (selDate && /^\d{4}-\d{2}-\d{2}$/.test(selDate)) month = rosterDate.slice(0, 7);
    var y, m;
    if (month && /^\d{4}-\d{2}$/.test(month)) { y = +month.slice(0, 4); m = +month.slice(5, 7); }
    else { y = now.getFullYear(); m = now.getMonth() + 1; }
    var dim = new Date(y, m, 0).getDate();
    var monthStr = y + "-" + _attPad(m);
    var from = monthStr + "-01", to = monthStr + "-" + _attPad(dim);
    var endDay = (monthStr === curMonth) ? now.getDate() : dim;

    var byStaff = {};
    recs.forEach(function (r) {
      var ds = _attDateStr(r.Date); if (ds < from || ds > to) return;
      var nm = String(r.Staff_Name || "").trim();
      (byStaff[nm] = byStaff[nm] || []).push({ date: ds, status: String(r.Status || "").trim().toLowerCase(), incentive: Number(r.Incentive || 0), deduction: Number(r.Deduction || 0), note: String(r.Note || "") });
    });
    var absentAll = {}, incentiveAll = {}, leavesByStaffMonth = {};
    recs.forEach(function (r) {
      var ds = _attDateStr(r.Date), nm = String(r.Staff_Name || "").trim();
      var ab = String(r.Status || "").trim().toLowerCase() === "absent";
      if (ab) absentAll[nm + "|" + ds] = true;
      if (Number(r.Incentive || 0) > 0) incentiveAll[nm + "|" + ds] = Number(r.Incentive || 0);
      if (ab && ds && _attDow(ds) !== 0) {
        var ym = ds.slice(0, 7);
        leavesByStaffMonth[nm] = leavesByStaffMonth[nm] || {};
        leavesByStaffMonth[nm][ym] = (leavesByStaffMonth[nm][ym] || 0) + 1;
      }
    });
    var paid = {}, salByStaff = {};
    sals.forEach(function (s) {
      var nm = String(s.Staff_Name || "").trim();
      paid[nm + "|" + String(s.Period || "").trim()] = Number(s.Amount || 0);
      (salByStaff[nm] = salByStaff[nm] || []).push({ period: String(s.Period || "").trim(), amount: Number(s.Amount || 0), at: String(s.Credited_At || "") });
    });

    var dow = now.getDay();
    var monOff = (dow === 0 ? -6 : 1 - dow);
    var monD = new Date(now.getFullYear(), now.getMonth(), now.getDate() + monOff);
    var satStr = Utilities.formatDate(new Date(monD.getFullYear(), monD.getMonth(), monD.getDate() + 5), "Asia/Kolkata", "yyyy-MM-dd");
    var prevKeys = [_attMonthShift(monthStr, -1), _attMonthShift(monthStr, -2), _attMonthShift(monthStr, -3)];

    var alerts = [], totalPayroll = 0;
    var staffOut = staff.map(function (st) {
      var name = String(st.Name).trim();
      var type = String(st.Type || "Monthly").trim().toLowerCase();
      var mySal = Number(st.Monthly_Salary || 0);
      var payDay = Number(st.Pay_Day || 0);
      var paidLeaves = Number(st.Paid_Leaves || 0);
      var wk = Number(st.Weekday_Rate || 0), sat = Number(st.Saturday_Rate || 0), sun = Number(st.Sunday_Rate || 0);
      var myRecs = byStaff[name] || [];

      var selRec = null; myRecs.forEach(function (r) { if (r.date === rosterDate) selRec = r; });
      var presentSel = !(selRec && selRec.status === "absent");
      var selIncentive = selRec ? selRec.incentive : 0;
      var selDeduction = selRec ? selRec.deduction : 0;

      var incentiveTotal = 0, deductionTotal = 0, leaveDays = 0, leaveList = [], incentives = [], deductions = [];
      myRecs.forEach(function (r) {
        if (r.incentive > 0) { incentiveTotal += r.incentive; incentives.push({ date: r.date, amount: r.incentive, note: r.note }); }
        if (r.deduction > 0) { deductionTotal += r.deduction; deductions.push({ date: r.date, amount: r.deduction, note: r.note }); }
        if (r.status === "absent") {
          var sunday = _attDow(r.date) === 0;
          if (!sunday) leaveDays++;
          leaveList.push({ date: r.date, sunday: sunday, reason: r.note });
        }
      });
      leaveList.sort(function (a, b) { return a.date.localeCompare(b.date); });
      var freeLeft = paidLeaves;
      leaveList.forEach(function (lv) {
        if (lv.sunday) { lv.deductible = false; return; }
        if (freeLeft > 0) { lv.deductible = false; freeLeft--; } else { lv.deductible = true; }
      });
      var deductibleLeaves = Math.max(0, leaveDays - paidLeaves);
      var freeUsed = Math.min(leaveDays, paidLeaves);

      var workingDays = 0, leavesElapsed = 0;
      for (var wd = 1; wd <= endDay; wd++) {
        var wds = monthStr + "-" + _attPad(wd);
        if (_attDow(wds) === 0) continue;
        workingDays++;
        if (absentAll[name + "|" + wds]) leavesElapsed++;
      }
      var presentDays = Math.max(0, workingDays - leavesElapsed);

      var curLeaves = (leavesByStaffMonth[name] && leavesByStaffMonth[name][monthStr]) || 0;
      var prevVals = prevKeys.map(function (k) { return (leavesByStaffMonth[name] && leavesByStaffMonth[name][k]) || 0; });
      var avg3 = (prevVals[0] + prevVals[1] + prevVals[2]) / 3;

      var out = {
        name: name, type: type, monthly_salary: mySal, pay_day: payDay, pay_cycle: String(st.Pay_Cycle || "Monthly"),
        paid_leaves: paidLeaves, weekday_rate: wk, saturday_rate: sat, sunday_rate: sun,
        present_sel: presentSel, sel_incentive: selIncentive, sel_deduction: selDeduction,
        leave_list: leaveList, incentives: incentives, deductions: deductions,
        incentive_total: incentiveTotal, deduction_total: deductionTotal,
        leave_days: leaveDays, deductible_leaves: deductibleLeaves, free_used: freeUsed,
        working_days: workingDays, present_days: presentDays,
        trend: { current: curLeaves, prev: prevVals, prev_keys: prevKeys, avg3: Math.round(avg3 * 10) / 10, up: (curLeaves > avg3 + 0.0001 && curLeaves >= 2) },
        salary_history: (salByStaff[name] || []).slice(-6).reverse()
      };

      if (type === "daily") {
        var endDs = (monthStr === curMonth) ? today : to;
        var earned = 0;
        for (var dd = 1; dd <= dim; dd++) {
          var ds = monthStr + "-" + _attPad(dd); if (ds > endDs) break;
          var w = _attDow(ds); if (w === 0) continue;
          if (absentAll[name + "|" + ds]) continue;
          earned += (w === 6 ? sat : wk);
        }
        out.base_earned = earned; out.per_day = wk; out.deduction = 0;
        out.net_salary = earned + incentiveTotal - deductionTotal;
        var weekAmt = 0, weekInc = 0;
        for (var k = 0; k < 6; k++) {
          var dx = new Date(monD.getFullYear(), monD.getMonth(), monD.getDate() + k);
          var dsx = Utilities.formatDate(dx, "Asia/Kolkata", "yyyy-MM-dd");
          var wx = dx.getDay();
          if (wx !== 0 && !absentAll[name + "|" + dsx]) weekAmt += (wx === 6 ? sat : wk);
          if (incentiveAll[name + "|" + dsx]) weekInc += incentiveAll[name + "|" + dsx];
        }
        var weekPeriod = "W" + satStr;
        out.week_amount = weekAmt + weekInc; out.week_period = weekPeriod; out.week_credited = paid.hasOwnProperty(name + "|" + weekPeriod);
        if ((dow === 6 || dow === 0) && !out.week_credited) {
          alerts.push({ name: name, amount: weekAmt + weekInc, period: weekPeriod, label: "Weekly wage (week ending Sat " + satStr + ")", base: weekAmt, deduction: 0, incentive: weekInc, adj: 0, cycle: "Weekly" });
        }
      } else {
        var perDay = dim > 0 ? mySal / dim : 0;
        var leaveDeduction = Math.round(perDay * deductibleLeaves);
        out.per_day = Math.round(perDay * 100) / 100; out.deduction = leaveDeduction;
        out.net_salary = mySal - leaveDeduction + incentiveTotal - deductionTotal;
        var monthPeriod = "M" + monthStr;
        out.pay_period = monthPeriod; out.credited = paid.hasOwnProperty(name + "|" + monthPeriod);
        out.pay_due = (monthStr === curMonth) && (now.getDate() >= payDay) && !out.credited;
        if (out.pay_due) alerts.push({ name: name, amount: out.net_salary, period: monthPeriod, label: "Salary (pay day " + payDay + ")", base: mySal, deduction: leaveDeduction, incentive: incentiveTotal, adj: deductionTotal, cycle: "Monthly" });
      }
      totalPayroll += Number(out.net_salary || 0);
      return out;
    });

    return { success: true, month: monthStr, days_in_month: dim, today: today, roster_date: rosterDate, today_dow: dow, total_payroll: totalPayroll, staff: staffOut, alerts: alerts };
  } catch (e) { return { success: false, error: String(e) }; }
}

// Header-driven upsert so it's robust to column order / added columns.
function _attUpsert(ss, date, name, mutate) {
  var ws = getOrCreateTab(ss, ATT_REC_TAB, ATT_REC_HEADERS);
  _attEnsureCols(ws, ATT_REC_HEADERS);
  var hIdx = headerIndex(ws);
  var headerRow = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var data = ws.getDataRange().getValues();
  var dCol = hIdx.Date - 1, nCol = hIdx.Staff_Name - 1;
  var rowNum = -1, cur = { Status: "", Incentive: "", Deduction: "", Note: "" };
  for (var i = 1; i < data.length; i++) {
    if (_attDateStr(data[i][dCol]) === date && String(data[i][nCol] || "").trim() === name) {
      rowNum = i + 1;
      cur.Status = String(data[i][hIdx.Status - 1] || "");
      cur.Incentive = data[i][hIdx.Incentive - 1];
      cur.Deduction = data[i][hIdx.Deduction - 1];
      cur.Note = String(data[i][hIdx.Note - 1] || "");
      break;
    }
  }
  mutate(cur);
  var stamp = Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyy-MM-dd HH:mm");
  var vals = headerRow.map(function (h) {
    if (h === "Date") return date;
    if (h === "Staff_Name") return name;
    if (h === "Status") return cur.Status;
    if (h === "Incentive") return cur.Incentive;
    if (h === "Deduction") return cur.Deduction;
    if (h === "Note") return cur.Note;
    if (h === "Updated") return stamp;
    return "";
  });
  if (rowNum > 0) ws.getRange(rowNum, 1, 1, vals.length).setValues([vals]);
  else ws.appendRow(vals);
  return { success: true };
}

function markAttendance(name, date, status, reason) {
  if (!name || !date) return { success: false, error: "name and date required" };
  var absent = String(status || "").trim().toLowerCase() === "absent";
  return _attUpsert(getSpreadsheet(), date, String(name).trim(), function (o) {
    o.Status = absent ? "Absent" : "";
    if (absent && reason) o.Note = reason;
  });
}

function addIncentive(name, date, amount, note) {
  if (!name || !date) return { success: false, error: "name and date required" };
  var amt = Number(amount || 0);
  return _attUpsert(getSpreadsheet(), date, String(name).trim(), function (o) { o.Incentive = amt > 0 ? amt : ""; if (note) o.Note = note; });
}

// Ad-hoc deduction (salary advance taken, fine, breakage…) on a date.
function addDeduction(name, date, amount, note) {
  if (!name || !date) return { success: false, error: "name and date required" };
  var amt = Number(amount || 0);
  return _attUpsert(getSpreadsheet(), date, String(name).trim(), function (o) { o.Deduction = amt > 0 ? amt : ""; if (note) o.Note = note; });
}

function markSalaryCredited(name, period, amount) {
  if (!name || !period) return { success: false, error: "name and period required" };
  var ws = getOrCreateTab(getSpreadsheet(), ATT_SAL_TAB, ATT_SAL_HEADERS);
  ws.appendRow([String(name).trim(), String(period).trim(), Number(amount || 0), Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyy-MM-dd HH:mm")]);
  // Force Period to plain text so values are never auto-converted to dates/numbers.
  try { ws.getRange(ws.getLastRow(), 2).setNumberFormat("@").setValue(String(period).trim()); } catch (e) {}
  return { success: true };
}

// Edit a staff member's pay config in SK_Staff (salary, pay day, daily rates,
// paid-leave allowance, active). Name and Type are NOT editable — attendance
// records are keyed by Name, so renaming would orphan them.
function updateStaff(name, fields) {
  if (!name) return { success: false, error: "name required" };
  var EDITABLE = { Monthly_Salary: 1, Pay_Day: 1, Weekday_Rate: 1, Saturday_Rate: 1, Sunday_Rate: 1, Paid_Leaves: 1, Active: 1 };
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, ATT_STAFF_TAB, ATT_STAFF_HEADERS);
  _attEnsureCols(ws, ATT_STAFF_HEADERS);
  var hIdx = headerIndex(ws);
  var data = ws.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === String(name).trim()) {
      Object.keys(fields || {}).forEach(function (k) {
        if (EDITABLE[k] && hIdx[k]) ws.getRange(i + 1, hIdx[k]).setValue(fields[k]);
      });
      return { success: true };
    }
  }
  return { success: false, error: "Staff not found: " + name };
}
