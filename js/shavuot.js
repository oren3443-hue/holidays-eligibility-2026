/**
 * shavuot.js - Shavuot 2026 eligibility logic.
 *
 * Shavuot 2026 falls on Friday 22.5.2026 (a short day → 7-hour cap).
 *
 * Eligible for ONE holiday day only if ALL of the following hold (checked in order;
 * the first failing test sets the reason):
 *   1. Present in the employee master file.
 *   2. Defined as a 6-days/week employee (5-day → not eligible).
 *   3. "Present" in the BEFORE file (worked ≥1 day OR vacation/sick/military marked).
 *   4. "Present" in the AFTER file (same rule).
 *   5. Tenure: start date on/before 22.2.2026 (3 months before the holiday).
 *
 * Payment = part-time fraction over the 3 paid months before the holiday:
 *   avg = paidWorkHours / paidWorkDays  → holiday hours = round2(min(avg, 7)).
 *
 * Universe = anyone appearing in the BEFORE or AFTER file, so non-eligible employees
 * are still shown with a reason. IDs join directly (payroll-system number ≡ master
 * מספר עובד), no name fallback.
 */

const ShavuotEligibility = (function () {
  const TENURE_CUTOFF = new Date(2026, 1, 22); // 22.2.2026 (3 months before 22.5)
  const HOLIDAY_DAY_CAP = 7;                    // Shavuot 2026 is a Friday — short day
  const THREE_MONTH_WINDOW = [2, 3, 4];         // Feb/Mar/Apr — the 3 months before May

  function round2(n) {
    return Math.round(n * 100) / 100;
  }

  function dailyStatus(rec) {
    if (!rec) return "נעדר";
    if (rec.work > 0) return "משמרת";
    if (rec.vacation > 0) return "חופש";
    if (rec.sick > 0) return "מחלה";
    if (rec.military > 0) return "מילואים";
    return "נעדר";
  }

  function isQualifyingPresence(rec) {
    if (!rec) return false;
    return rec.work > 0 || rec.vacation > 0 || rec.sick > 0 || rec.military > 0;
  }

  function hasTenure(startDate) {
    if (!startDate) return false;
    return startDate.getTime() <= TENURE_CUTOFF.getTime();
  }

  /**
   * Calculate one employee row.
   * @param {*} empId
   * @param {*} sources - { employees, threeMonths, before, after }
   * @returns row object with raw + computed fields.
   */
  function calculateEmployee(empId, sources) {
    const emp = sources.employees.get(empId);
    const b = sources.before.get(empId);
    const a = sources.after.get(empId);
    const three = sources.threeMonths.get(empId);

    // Names/department come from the network-pay reports; dept number from master.
    const nameRec = b || a || {};
    const firstName = nameRec.firstName || "";
    const lastName = nameRec.lastName || "";
    const deptName = nameRec.deptName || "";
    const deptNumber = emp ? emp.deptNumber : null;
    const daysPerWeek = emp ? emp.daysPerWeek : null;
    const basisCode = emp ? emp.basisCode : null;
    const isGlobal = basisCode === "ח"; // ח = חודשי/גלובלי — paid the holiday inside the monthly salary, so no separate holiday pay
    const startDate = emp ? emp.startDate : null;
    const tenureOk = hasTenure(startDate);

    const totalPaidWorkDays = three ? three.paidWorkDays : 0;
    const totalPaidWorkHours = three ? three.paidWorkHours : 0;

    const avgHoursPerDay =
      totalPaidWorkHours > 0 && totalPaidWorkDays > 0
        ? totalPaidWorkHours / totalPaidWorkDays
        : 0;

    const statusBefore = dailyStatus(b);
    const statusAfter = dailyStatus(a);
    const presentBefore = isQualifyingPresence(b);
    const presentAfter = isQualifyingPresence(a);

    // Eligibility chain — first failing test wins.
    let eligible = false, reason = "";
    if (!emp) {
      reason = "חסר בקובץ פרטי עובדים";
    } else if (isGlobal) {
      reason = "עובד חודשי (גלובלי)";
    } else if (daysPerWeek !== 6) {
      reason = "לא במתכונת 6 ימים";
    } else if (!presentBefore) {
      reason = "לא נכח לפני החג";
    } else if (!presentAfter) {
      reason = "לא נכח אחרי החג";
    } else if (!tenureOk) {
      reason = "אין ותק (פחות מ-3 חודשים)";
    } else {
      eligible = true;
    }

    const holidayDaysCount = eligible ? 1 : 0;
    const holidayPayHours = eligible ? round2(Math.min(avgHoursPerDay, HOLIDAY_DAY_CAP)) : 0;

    return {
      empId,
      firstName,
      lastName,
      fullName: `${firstName} ${lastName}`.trim() || String(empId),
      deptName,
      deptNumber,
      daysPerWeek,
      basisCode,
      startDate,
      tenureOk,
      statusBefore,
      statusAfter,
      workBefore: b ? round2(b.work) : 0,
      workAfter: a ? round2(a.work) : 0,
      eligible,
      reason,
      totalPaidWorkHours: round2(totalPaidWorkHours),
      totalPaidWorkDays,
      avgHoursPerDay: round2(avgHoursPerDay),
      holidayDaysCount,
      holidayPayHours,
    };
  }

  /**
   * Run full computation. Universe = union of IDs in the before/after files.
   * @param {*} sources - { employees, threeMonths, before, after }
   * @returns Array of result rows, sorted by employee id.
   */
  function calculateAll(sources) {
    const ids = new Set();
    for (const id of sources.before.keys()) ids.add(id);
    for (const id of sources.after.keys()) ids.add(id);
    const sorted = Array.from(ids).sort((a, b) => {
      const na = Number(a), nb = Number(b);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return String(a).localeCompare(String(b));
    });
    return sorted.map((id) => calculateEmployee(id, sources));
  }

  return {
    calculateAll,
    calculateEmployee,
    TENURE_CUTOFF,
    HOLIDAY_DAY_CAP,
    THREE_MONTH_WINDOW,
  };
})();
