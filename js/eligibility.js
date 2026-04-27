/**
 * eligibility.js - Pesach 2026 eligibility logic.
 *
 * Holiday calendar (Israel):
 *  - 31.3 (Tue): Pesach A eligibility check day (regular work day).
 *  - 1.4 (Wed):  Erev Pesach. Company doesn't work. Counted as a Chol HaMoed vacation day.
 *  - 2.4 (Thu):  Pesach A (paid as holiday if eligible).
 *  - 3.4 (Fri):  Chol HaMoed. Vacation day for 6-day employees only.
 *  - 4.4 (Sat):  Shabbat - not counted.
 *  - 5.4 (Sun):  Chol HaMoed. Vacation day.
 *  - 6.4 (Mon):  Chol HaMoed. Vacation day.
 *  - 7.4 (Tue):  Chol HaMoed / Erev Pesach B. Vacation day (6h cap for 6-day).
 *  - 8.4 (Wed):  Pesach B (paid as holiday if eligible).
 *  - 9.4 (Thu):  Pesach B eligibility check day (regular work day).
 */

const Eligibility = (function () {
  const TENURE_CUTOFF = new Date(2026, 0, 1); // 1.1.2026

  // Vacation day list per work-week pattern.
  // Each entry: { date: 'DD.MM', cap5, cap6 } - hour cap for 5-day and 6-day employees.
  // null cap means the day is not used for that pattern.
  const CHOL_HAMOED_DAYS = [
    { date: "1.4",  cap5: 8.4, cap6: 6 }, // Erev Pesach (Wed)
    { date: "3.4",  cap5: null, cap6: 6 }, // Friday - 6-day only
    { date: "5.4",  cap5: 8.4, cap6: 8 }, // Sunday
    { date: "6.4",  cap5: 8.4, cap6: 8 }, // Monday
    { date: "7.4",  cap5: 8.4, cap6: 6 }, // Tue (Erev Pesach B for 6-day)
  ];

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

  function chosenDays(daysPerWeek) {
    if (daysPerWeek === 5) return CHOL_HAMOED_DAYS.filter((d) => d.cap5 !== null);
    if (daysPerWeek === 6) return CHOL_HAMOED_DAYS;
    return [];
  }

  function maxDaysFor(daysPerWeek) {
    if (daysPerWeek === 5) return 4;
    if (daysPerWeek === 6) return 5;
    return 0;
  }

  function capFor(day, daysPerWeek) {
    if (daysPerWeek === 5) return day.cap5;
    if (daysPerWeek === 6) return day.cap6;
    return null;
  }

  /**
   * Calculate one employee row.
   * @param {*} empId - employee id
   * @param {*} sources - { employees, threeMonths, march31, april9 }
   * @returns row object with raw + computed fields.
   */
  function calculateEmployee(empId, sources) {
    const emp = sources.employees.get(empId);
    const three = sources.threeMonths.get(empId);
    const m31 = sources.march31.get(empId);
    const a9 = sources.april9.get(empId);
    // Fallback names from any daily file (april preferred for relevance, then 31.3, then 9.4)
    const aprilRec = sources.april ? sources.april.get(empId) : null;
    const firstName = (aprilRec && aprilRec.firstName) || (m31 && m31.firstName) || (a9 && a9.firstName) || "";
    const lastName  = (aprilRec && aprilRec.lastName)  || (m31 && m31.lastName)  || (a9 && a9.lastName)  || "";

    const daysPerWeek = emp ? emp.daysPerWeek : null;
    const startDate = emp ? emp.startDate : null;
    const endDate = emp ? emp.endDate : null;
    const tenureOk = hasTenure(startDate);

    // Avg hours per day from 3-month file
    const totalWorkHours = three ? three.workHours : 0;
    const totalWorkDays = three ? three.workDays : 0;
    const avgHoursPerDay = totalWorkDays > 0 ? totalWorkHours / totalWorkDays : 0;

    // Vacation balance from March
    const vacationBalance = three ? three.marchBalance : 0;

    // Eligibility checks
    const status31 = dailyStatus(m31);
    const status9  = dailyStatus(a9);
    const eligibleA = tenureOk && isQualifyingPresence(m31);
    const eligibleB = tenureOk && isQualifyingPresence(a9);

    // Reason for ineligibility
    function reason(elig, status) {
      if (elig) return "";
      if (!tenureOk) return "אין ותק";
      if (status === "נעדר") return "לא נכח ביום הבדיקה";
      return "";
    }

    // Chol HaMoed vacation usage
    const dayList = chosenDays(daysPerWeek);
    const maxDays = maxDaysFor(daysPerWeek);
    const daysFromBalance = Math.floor(vacationBalance);
    const daysToUse = Math.max(0, Math.min(daysFromBalance, maxDays));

    // Allocate hours per day (chronological, first daysToUse days). Rounded to 2 decimals.
    const dayHoursByDate = {};
    for (let i = 0; i < dayList.length; i++) {
      const d = dayList[i];
      if (i < daysToUse && avgHoursPerDay > 0) {
        const cap = capFor(d, daysPerWeek);
        dayHoursByDate[d.date] = round2(Math.min(avgHoursPerDay, cap));
      } else {
        dayHoursByDate[d.date] = null; // not used
      }
    }
    const totalHolidayHours = Object.values(dayHoursByDate).reduce((s, v) => s + (v || 0), 0);

    return {
      empId,
      firstName,
      lastName,
      daysPerWeek,
      startDate,
      endDate,
      tenureOk,
      status31,
      status9,
      eligibleA,
      eligibleB,
      reasonA: reason(eligibleA, status31),
      reasonB: reason(eligibleB, status9),
      totalWorkHours: round2(totalWorkHours),
      totalWorkDays,
      avgHoursPerDay: round2(avgHoursPerDay),
      vacationBalance: round2(vacationBalance),
      daysToUse,
      hours_1_4: dayHoursByDate["1.4"],
      hours_3_4: dayHoursByDate["3.4"],
      hours_5_4: dayHoursByDate["5.4"],
      hours_6_4: dayHoursByDate["6.4"],
      hours_7_4: dayHoursByDate["7.4"],
      totalHolidayHours: round2(totalHolidayHours),
    };
  }

  function round2(n) {
    return Math.round(n * 100) / 100;
  }

  /**
   * Run full computation across all April-active employees.
   * @param {*} sources - { employees, threeMonths, march31, april9, april }
   * @returns Array of result rows.
   */
  function calculateAll(sources) {
    // Use April file as the relevance filter
    const activeIds = Array.from(sources.april.keys());
    // Sort numerically when possible, else lexicographically
    activeIds.sort((a, b) => {
      const na = Number(a), nb = Number(b);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return String(a).localeCompare(String(b));
    });
    return activeIds.map((id) => calculateEmployee(id, sources));
  }

  return {
    calculateAll,
    calculateEmployee,
    CHOL_HAMOED_DAYS,
    TENURE_CUTOFF,
  };
})();
