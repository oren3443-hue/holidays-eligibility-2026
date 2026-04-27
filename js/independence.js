/**
 * independence.js - Yom HaAtzmaut 2026 eligibility logic.
 *
 * Holiday window: 21.4 20:00 → 22.4 20:00 (24 hours, crosses midnight).
 *
 * Two mutually exclusive payment modes per employee:
 *   - "worked":      worked some hours inside the window → pay actual hours (תוספת שעות חג 100%).
 *   - "holiday_pay": didn't work inside the window, but had any shift on 21.4 OR 23.4
 *                   (full calendar days, any time) → pay 3-month avg capped at 8/8.4 (שעות חג).
 *   - "ineligible":  no qualifying day OR tenure < 22.1.2026.
 *
 * Universe: only employees who appear in BOTH the employees file AND in at least
 * one shift report (20.4-23.4). This naturally excludes employees on military
 * reserve for the entire window — they have no shifts in the report.
 */

const IndependenceEligibility = (function () {
  const TENURE_CUTOFF = new Date(2026, 0, 22); // 22.1.2026 (3 months before holiday)

  // 21.4.2026 20:00 → 22.4.2026 20:00
  const HOLIDAY_WINDOW = {
    start: new Date(2026, 3, 21, 20, 0, 0),
    end:   new Date(2026, 3, 22, 20, 0, 0),
  };

  // Day-before / day-after qualifying ranges (full calendar day each).
  const DAY_BEFORE = {
    start: new Date(2026, 3, 21, 0, 0, 0),
    end:   new Date(2026, 3, 21, 23, 59, 59, 999),
  };
  const DAY_AFTER = {
    start: new Date(2026, 3, 23, 0, 0, 0),
    end:   new Date(2026, 3, 23, 23, 59, 59, 999),
  };

  function overlapMs(aStart, aEnd, bStart, bEnd) {
    const s = Math.max(aStart.getTime(), bStart.getTime());
    const e = Math.min(aEnd.getTime(), bEnd.getTime());
    return Math.max(0, e - s);
  }

  function sumHoursInWindow(shifts, win) {
    let ms = 0;
    for (const sh of shifts) {
      ms += overlapMs(sh.entryAt, sh.exitAt, win.start, win.end);
    }
    return ms / 3600000;
  }

  function hasShiftOverlapping(shifts, win) {
    for (const sh of shifts) {
      if (overlapMs(sh.entryAt, sh.exitAt, win.start, win.end) > 0) return true;
    }
    return false;
  }

  function round2(n) {
    return Math.round(n * 100) / 100;
  }

  function modeLabelOf(mode) {
    if (mode === "worked") return "עבד בחג";
    if (mode === "holiday_pay") return "זכאי תשלום חג";
    return "לא זכאי";
  }

  function calculateEmployee(empId, sources) {
    const emp = sources.employees.get(empId);
    if (!emp) return null; // employee not in master file → skip per spec

    const shiftRec = sources.shifts.get(empId);
    const shifts = (shiftRec && shiftRec.shifts) || [];
    if (shifts.length === 0) return null; // no shifts in 20.4-23.4 → not relevant

    const three = sources.threeMonths.get(empId);
    const totalWorkHours = three ? three.workHours : 0;
    const totalWorkDays  = three ? three.workDays : 0;
    const avgHoursPerDay = totalWorkDays > 0 ? totalWorkHours / totalWorkDays : 0;

    const tenureOk = !!(emp.startDate && emp.startDate.getTime() <= TENURE_CUTOFF.getTime());
    const hoursWorkedInWindow = sumHoursInWindow(shifts, HOLIDAY_WINDOW);
    const workedDayBefore = hasShiftOverlapping(shifts, DAY_BEFORE);
    const workedDayAfter  = hasShiftOverlapping(shifts, DAY_AFTER);
    const qualifyingDay   = workedDayBefore || workedDayAfter;

    const cap = emp.daysPerWeek === 5 ? 8.4 : emp.daysPerWeek === 6 ? 8 : 0;
    const avgCapped = avgHoursPerDay > 0 && cap > 0 ? round2(Math.min(avgHoursPerDay, cap)) : 0;

    let mode, holidayPayHours = 0, extraHolidayHours = 0, holidayDaysCount = 0, reason = "";

    if (hoursWorkedInWindow > 0) {
      mode = "worked";
      extraHolidayHours = round2(hoursWorkedInWindow);
    } else if (qualifyingDay && tenureOk) {
      mode = "holiday_pay";
      holidayPayHours = avgCapped;
      holidayDaysCount = 1;
    } else {
      mode = "ineligible";
      if (!tenureOk && !qualifyingDay) reason = "אין ותק וגם לא עבד יום לפני/אחרי";
      else if (!qualifyingDay) reason = "לא עבד יום לפני/אחרי";
      else if (!tenureOk) reason = "אין ותק";
    }

    const branches = shiftRec.branches ? Array.from(shiftRec.branches).join(", ") : "";
    const firstName = emp.firstName || (shiftRec.firstName || "");
    const lastName  = emp.lastName  || (shiftRec.lastName  || "");
    const fullName  = `${firstName} ${lastName}`.trim() || String(empId);

    return {
      empId,
      firstName,
      lastName,
      fullName,
      deptName: shiftRec.deptName || "",
      deptNumber: emp.deptNumber || null,
      daysPerWeek: emp.daysPerWeek,
      startDate: emp.startDate,
      tenureOk,
      branches,
      workedDayBefore,
      workedDayAfter,
      hoursWorkedInWindow: round2(hoursWorkedInWindow),
      mode,
      modeLabel: modeLabelOf(mode),
      avgHoursPerDay: round2(avgHoursPerDay),
      totalWorkHours: round2(totalWorkHours),
      totalWorkDays,
      holidayDaysCount,
      holidayPayHours,
      extraHolidayHours,
      reason,
    };
  }

  function calculateAll(sources) {
    const ids = [];
    for (const [id] of sources.shifts.entries()) {
      if (sources.employees.has(id)) ids.push(id);
    }
    ids.sort((a, b) => {
      const na = Number(a), nb = Number(b);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return String(a).localeCompare(String(b));
    });
    const rows = [];
    for (const id of ids) {
      const r = calculateEmployee(id, sources);
      if (r) rows.push(r);
    }
    return rows;
  }

  return {
    calculateAll,
    calculateEmployee,
    HOLIDAY_WINDOW,
    TENURE_CUTOFF,
  };
})();
