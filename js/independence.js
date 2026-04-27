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
  // The Atzmaut calendar day itself (22.4 full day). Used to detect overlap
  // with sick leave — can't double-pay holiday pay AND sick pay for the same day.
  const HOLIDAY_DAY = {
    start: new Date(2026, 3, 22, 0, 0, 0),
    end:   new Date(2026, 3, 22, 23, 59, 59, 999),
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

  function sameCalendarDay(d, win) {
    return d.getTime() >= win.start.getTime() && d.getTime() <= win.end.getTime();
  }

  function absenceTypeLabel(t) {
    if (t === "vacation") return "חופש";
    if (t === "sick") return "מחלה";
    if (t === "military") return "מילואים";
    return "";
  }

  /** Sum of shift hours that fall on a calendar day (full 24h, not the holiday window). */
  function shiftHoursOnDay(shifts, win) {
    let ms = 0;
    for (const sh of shifts) {
      ms += overlapMs(sh.entryAt, sh.exitAt, win.start, win.end);
    }
    return ms / 3600000;
  }

  /** Build a human-readable presence string for a calendar day:
   *  e.g. "06:00→14:00; 18:00→02:00" or "חופש" or "—" */
  function presenceOnDay(shifts, absences, win) {
    const overlapping = shifts.filter((sh) => overlapMs(sh.entryAt, sh.exitAt, win.start, win.end) > 0);
    const fmt = (d) => `${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
    const parts = overlapping.map((sh) => {
      // Show portion of shift that touches this day (capped to day boundaries when crosses midnight)
      const s = sh.entryAt.getTime() < win.start.getTime() ? win.start : sh.entryAt;
      const e = sh.exitAt.getTime()  > win.end.getTime()   ? win.end   : sh.exitAt;
      const crossed = sh.exitAt.getTime() > win.end.getTime() || sh.entryAt.getTime() < win.start.getTime();
      return `${fmt(s)}→${fmt(e)}${crossed ? "*" : ""}`;
    });
    if (parts.length) return parts.join("; ");
    const a = absences.find((ab) => sameCalendarDay(ab.date, win));
    if (a) return absenceTypeLabel(a.type);
    return "—";
  }

  function hasAbsenceOnDay(absences, win) {
    return absences.some((ab) => sameCalendarDay(ab.date, win));
  }

  function calculateEmployee(empId, sources) {
    const emp = sources.employees.get(empId);
    if (!emp) return null; // employee not in master file → skip per spec

    const shiftRec = sources.shifts.get(empId);
    const shifts = (shiftRec && shiftRec.shifts) || [];
    const absences = (shiftRec && shiftRec.absences) || [];
    if (shifts.length === 0 && absences.length === 0) return null;

    // Military-only filter: no shifts AND every absence is military → exclude entirely
    const nonMilitaryAbsences = absences.filter((a) => a.type !== "military");
    if (shifts.length === 0 && nonMilitaryAbsences.length === 0) return null;

    const three = sources.threeMonths.get(empId);
    const totalWorkHours = three ? three.workHours : 0;
    const totalWorkDays  = three ? three.workDays : 0;
    const avgHoursPerDay = totalWorkDays > 0 ? totalWorkHours / totalWorkDays : 0;

    const tenureOk = !!(emp.startDate && emp.startDate.getTime() <= TENURE_CUTOFF.getTime());
    const hoursWorkedInWindow = sumHoursInWindow(shifts, HOLIDAY_WINDOW);
    const presentDayBefore = hasShiftOverlapping(shifts, DAY_BEFORE) || hasAbsenceOnDay(absences, DAY_BEFORE);
    const presentDayAfter  = hasShiftOverlapping(shifts, DAY_AFTER)  || hasAbsenceOnDay(absences, DAY_AFTER);
    const workedDayBefore  = hasShiftOverlapping(shifts, DAY_BEFORE);
    const workedDayAfter   = hasShiftOverlapping(shifts, DAY_AFTER);
    const qualifyingDay    = presentDayBefore || presentDayAfter;

    // Can't double-pay: if the employee was on sick leave OR military reserve
    // on the holiday day itself (22.4), that pay component covers the day —
    // no holiday pay on top. Vacation does NOT disqualify (in Israel a vacation
    // day that falls on a holiday is credited back to the balance, so the
    // employee still receives the holiday pay).
    const sickOnHoliday     = absences.some((ab) => ab.type === "sick"     && sameCalendarDay(ab.date, HOLIDAY_DAY));
    const militaryOnHoliday = absences.some((ab) => ab.type === "military" && sameCalendarDay(ab.date, HOLIDAY_DAY));

    const cap = emp.daysPerWeek === 5 ? 8.4 : emp.daysPerWeek === 6 ? 8 : 0;
    const avgCapped = avgHoursPerDay > 0 && cap > 0 ? round2(Math.min(avgHoursPerDay, cap)) : 0;

    let mode, holidayPayHours = 0, extraHolidayHours = 0, holidayDaysCount = 0, reason = "";

    if (hoursWorkedInWindow > 0) {
      mode = "worked";
      extraHolidayHours = round2(hoursWorkedInWindow);
    } else if (sickOnHoliday) {
      mode = "ineligible";
      reason = "במחלה ביום החג (22.4) — מקבל דמי מחלה, לא תשלום חג";
    } else if (militaryOnHoliday) {
      mode = "ineligible";
      reason = "במילואים ביום החג (22.4) — מקבל תשלום מילואים, לא תשלום חג";
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
    const fullName  = shiftRec.fullName || `${firstName} ${lastName}`.trim() || String(empId);

    // Raw per-day totals + presence text for verification
    const hours_21_4 = round2(shiftHoursOnDay(shifts, DAY_BEFORE));
    const hours_22_4 = round2(shiftHoursOnDay(shifts, { start: new Date(2026, 3, 22, 0, 0), end: new Date(2026, 3, 22, 23, 59, 59, 999) }));
    const hours_23_4 = round2(shiftHoursOnDay(shifts, DAY_AFTER));
    const presence_21_4 = presenceOnDay(shifts, absences, DAY_BEFORE);
    const presence_22_4 = presenceOnDay(shifts, absences, { start: new Date(2026, 3, 22, 0, 0), end: new Date(2026, 3, 22, 23, 59, 59, 999) });
    const presence_23_4 = presenceOnDay(shifts, absences, DAY_AFTER);

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
      presentDayBefore,
      presentDayAfter,
      workedDayBefore,
      workedDayAfter,
      hours_21_4,
      hours_22_4,
      hours_23_4,
      presence_21_4,
      presence_22_4,
      presence_23_4,
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
    // Universe = anyone in the shift map (whether shift or absence) AND in employees file.
    // Military-only filtering happens inside calculateEmployee.
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
