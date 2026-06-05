// ════════════════════════════════════════════════════════════════════
// STAFF ATTENDANCE & SALARY
//  SK_Staff      : roster + pay config (auto-seeded on first read).
//  SK_Attendance : EXCEPTIONS only — a row exists when a day is marked ABSENT
//                  and/or carries an INCENTIVE. No row = present, no incentive
//                  (everyone auto-present unless explicitly marked absent).
//  SK_Salary_Log : records when a salary/wage was credited (clears the alert).
//  Monthly staff: per-day = salary / days-in-month; each non-Sunday leave
//  deducts one per-day. Daily staff (Abhijeet): weekday/Sat rates, Sun off,
//  paid weekly (alert on Sat/Sun).
// ════════════════════════════════════════════════════════════════════
var ATT_STAFF_TAB = "SK_Staff", ATT_REC_TAB = "SK_Attendance", ATT_SAL_TAB = "SK_Salary_Log";
var ATT_STAFF_HEADERS = ["Name","Type","Monthly_Salary","Pay_Day","Pay_Cycle","Weekday_Rate","Saturday_Rate","Sunday_Rate","Active"];
var ATT_REC_HEADERS   = ["Date","Staff_Name","Status","Incentive","Note","Updated"];
var ATT_SAL_HEADERS   = ["Staff_Name","Period","Amount","Credited_At"];

function _attSeed(ss) {
  var ws = getOrCreateTab(ss, ATT_STAFF_TAB, ATT_STAFF_HEADERS);
  if (ws.getLastRow() > 1) return ws;
  ws.getRange(2, 1, 6, ATT_STAFF_HEADERS.length).setValues([
    ["Rupa Tai","Monthly",24000,15,"Monthly","","","","Yes"],
    ["Pooja Tai","Monthly",8000,18,"Monthly","","","","Yes"],
    ["Meena Tai","Monthly",5000,5,"Monthly","","","","Yes"],
    ["Anita Tai","Monthly",3000,1,"Monthly","","","","Yes"],
    ["Manisha Tai","Monthly",2000,1,"Monthly","","","","Yes"],
    ["Abhijeet Parekar","Daily",0,"","Weekly",700,600,0,"Yes"]
  ]);
  return ws;
}
function _attPad(n) { return (n < 10 ? "0" : "") + n; }
function _attDateStr(v) { return v instanceof Date ? Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd") : String(v || "").trim(); }
function _attDow(ds) { return new Date(ds + "T12:00:00+05:30").getDay(); } // 0 Sun .. 6 Sat

function getAttendanceData(month) {
  try {
    var ss = getSpreadsheet();
    var staffWs = _attSeed(ss);
    var recWs = getOrCreateTab(ss, ATT_REC_TAB, ATT_REC_HEADERS);
    var salWs = getOrCreateTab(ss, ATT_SAL_TAB, ATT_SAL_HEADERS);
    var staff = getAllRows(staffWs).filter(function (r) { return String(r.Name || "").trim() && String(r.Active || "Yes").toLowerCase() !== "no"; });
    var recs = getAllRows(recWs);
    var sals = getAllRows(salWs);

    var now = getISTDate();
    var today = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
    var curMonth = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM");
    var y, m;
    if (month && /^\d{4}-\d{2}$/.test(month)) { y = +month.slice(0, 4); m = +month.slice(5, 7); }
    else { y = now.getFullYear(); m = now.getMonth() + 1; }
    var dim = new Date(y, m, 0).getDate();
    var monthStr = y + "-" + _attPad(m);
    var from = monthStr + "-01", to = monthStr + "-" + _attPad(dim);

    var byStaff = {};
    recs.forEach(function (r) {
      var ds = _attDateStr(r.Date); if (ds < from || ds > to) return;
      var nm = String(r.Staff_Name || "").trim();
      (byStaff[nm] = byStaff[nm] || []).push({ date: ds, status: String(r.Status || "").trim().toLowerCase(), incentive: Number(r.Incentive || 0), note: String(r.Note || "") });
    });
    var absentAll = {}, incentiveAll = {};
    recs.forEach(function (r) {
      var ds = _attDateStr(r.Date), nm = String(r.Staff_Name || "").trim();
      if (String(r.Status || "").trim().toLowerCase() === "absent") absentAll[nm + "|" + ds] = true;
      if (Number(r.Incentive || 0) > 0) incentiveAll[nm + "|" + ds] = Number(r.Incentive || 0);
    });
    var paid = {};
    sals.forEach(function (s) { paid[String(s.Staff_Name || "").trim() + "|" + String(s.Period || "").trim()] = Number(s.Amount || 0); });

    // Current week (Mon..Sat) for daily weekly pay
    var dow = now.getDay();
    var monOff = (dow === 0 ? -6 : 1 - dow);
    var monD = new Date(now.getFullYear(), now.getMonth(), now.getDate() + monOff);
    var satStr = Utilities.formatDate(new Date(monD.getFullYear(), monD.getMonth(), monD.getDate() + 5), "Asia/Kolkata", "yyyy-MM-dd");

    var alerts = [];
    var staffOut = staff.map(function (st) {
      var name = String(st.Name).trim();
      var type = String(st.Type || "Monthly").trim().toLowerCase();
      var mySal = Number(st.Monthly_Salary || 0);
      var payDay = Number(st.Pay_Day || 0);
      var wk = Number(st.Weekday_Rate || 0), sat = Number(st.Saturday_Rate || 0), sun = Number(st.Sunday_Rate || 0);
      var myRecs = byStaff[name] || [];

      var todayRec = null; myRecs.forEach(function (r) { if (r.date === today) todayRec = r; });
      var presentToday = !(todayRec && todayRec.status === "absent");
      var todayIncentive = todayRec ? todayRec.incentive : 0;

      var incentiveTotal = 0, leaveDays = 0, absentDates = [], incentives = [];
      myRecs.forEach(function (r) {
        if (r.incentive > 0) { incentiveTotal += r.incentive; incentives.push({ date: r.date, amount: r.incentive, note: r.note }); }
        if (r.status === "absent") { absentDates.push(r.date); if (_attDow(r.date) !== 0) leaveDays++; }
      });

      var out = { name: name, type: type, monthly_salary: mySal, pay_day: payDay, pay_cycle: String(st.Pay_Cycle || "Monthly"),
        weekday_rate: wk, saturday_rate: sat, sunday_rate: sun,
        present_today: presentToday, today_incentive: todayIncentive,
        absent_dates: absentDates, incentives: incentives, incentive_total: incentiveTotal, leave_days: leaveDays };

      if (type === "daily") {
        var endDs = (monthStr === curMonth) ? today : to;
        var earned = 0;
        for (var dd = 1; dd <= dim; dd++) {
          var ds = monthStr + "-" + _attPad(dd); if (ds > endDs) break;
          var w = _attDow(ds); if (w === 0) continue;          // Sunday — no pay
          if (absentAll[name + "|" + ds]) continue;            // absent — no pay
          earned += (w === 6 ? sat : wk);
        }
        out.base_earned = earned; out.per_day = wk; out.deduction = 0;
        out.net_salary = earned + incentiveTotal;
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
          alerts.push({ name: name, amount: weekAmt + weekInc, period: weekPeriod, label: "Weekly wage (week ending Sat " + satStr + ")", base: weekAmt, deduction: 0, incentive: weekInc, cycle: "Weekly" });
        }
      } else {
        var perDay = dim > 0 ? mySal / dim : 0;
        var deduction = Math.round(perDay * leaveDays);
        out.per_day = Math.round(perDay * 100) / 100; out.deduction = deduction;
        out.net_salary = mySal - deduction + incentiveTotal;
        out.pay_period = monthStr; out.credited = paid.hasOwnProperty(name + "|" + monthStr);
        out.pay_due = (monthStr === curMonth) && (now.getDate() >= payDay) && !out.credited;
        if (out.pay_due) alerts.push({ name: name, amount: out.net_salary, period: monthStr, label: "Salary (pay day " + payDay + ")", base: mySal, deduction: deduction, incentive: incentiveTotal, cycle: "Monthly" });
      }
      return out;
    });

    return { success: true, month: monthStr, days_in_month: dim, today: today, today_dow: dow, staff: staffOut, alerts: alerts };
  } catch (e) { return { success: false, error: String(e) }; }
}

function _attUpsert(ss, date, name, mutate) {
  var ws = getOrCreateTab(ss, ATT_REC_TAB, ATT_REC_HEADERS);
  var data = ws.getDataRange().getValues();
  var rowNum = -1, rowObj = { Status: "", Incentive: "", Note: "" };
  for (var i = 1; i < data.length; i++) {
    if (_attDateStr(data[i][0]) === date && String(data[i][1] || "").trim() === name) {
      rowNum = i + 1; rowObj.Status = String(data[i][2] || ""); rowObj.Incentive = data[i][3]; rowObj.Note = String(data[i][4] || ""); break;
    }
  }
  mutate(rowObj);
  var stamp = Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyy-MM-dd HH:mm");
  var vals = [date, name, rowObj.Status, rowObj.Incentive, rowObj.Note, stamp];
  if (rowNum > 0) ws.getRange(rowNum, 1, 1, ATT_REC_HEADERS.length).setValues([vals]);
  else ws.appendRow(vals);
  return { success: true };
}

function markAttendance(name, date, status) {
  if (!name || !date) return { success: false, error: "name and date required" };
  var absent = String(status || "").trim().toLowerCase() === "absent";
  return _attUpsert(getSpreadsheet(), date, String(name).trim(), function (o) { o.Status = absent ? "Absent" : ""; });
}

function addIncentive(name, date, amount, note) {
  if (!name || !date) return { success: false, error: "name and date required" };
  var amt = Number(amount || 0);
  return _attUpsert(getSpreadsheet(), date, String(name).trim(), function (o) { o.Incentive = amt > 0 ? amt : ""; if (note) o.Note = note; });
}

function markSalaryCredited(name, period, amount) {
  if (!name || !period) return { success: false, error: "name and period required" };
  var ws = getOrCreateTab(getSpreadsheet(), ATT_SAL_TAB, ATT_SAL_HEADERS);
  ws.appendRow([String(name).trim(), String(period).trim(), Number(amount || 0), Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyy-MM-dd HH:mm")]);
  return { success: true };
}
