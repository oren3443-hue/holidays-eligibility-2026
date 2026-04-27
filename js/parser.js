/**
 * parser.js - Reads the 5 input Excel files and produces a unified data structure.
 *
 * Column mappings (1-indexed in the plan, converted to 0-indexed below):
 *
 * 1) Employee details (פרטי עובדים) - sheet "גיליון1", header row 1:
 *    col 1  (idx 0)  : מספר עובד            - master employee ID
 *    col 65 (idx 64) : תאריך תחילת עבודה   - start date
 *    col 66 (idx 65) : תאריך הפסקת עבודה  - termination date (active if empty)
 *    col 76 (idx 75) : ימי עבודה בשבוע     - 5 or 6
 *
 * 2) Daily attendance (31.3 / 9.4 / April) - sheet "Sheet1", header row 2:
 *    col 5  (idx 4)  : שם פרטי
 *    col 6  (idx 5)  : שם משפחה
 *    col 12 (idx 11) : מס' עובד במע' שכר   - join key  (drop if 0/null)
 *    col 15 (idx 14) : ימי עבודה בפועל
 *    col 24 (idx 23) : ימי חופש
 *    col 25 (idx 24) : ימי מחלה
 *    col 26 (idx 25) : ימי מילואים
 *
 * 3) 3-month attendance - sheet "גיליון1", header row 1:
 *    col 3  (idx 2)  : מספר עובד   - master ID
 *    col 6  (idx 5)  : מספר חודש   - 1/2/3
 *    col 14 (idx 13) : י"ע בפועל
 *    col 16 (idx 15) : ש"ע בפועל
 *    col 19 (idx 18) : חופשה - יתרה
 */

const Parser = (function () {
  /** Read a workbook from an ArrayBuffer and return rows of the first sheet (or named sheet) as arrays. */
  function readSheet(buffer, sheetName) {
    const wb = XLSX.read(buffer, { type: "array", cellDates: true });
    const name = sheetName && wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames[0];
    const ws = wb.Sheets[name];
    // header:1 => array-of-arrays. defval:null keeps empty cells as null.
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  function toNumber(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    const n = parseFloat(String(v).replace(/,/g, ""));
    return isNaN(n) ? 0 : n;
  }

  function toDate(v) {
    if (!v) return null;
    if (v instanceof Date) return v;
    if (typeof v === "number") {
      // Excel serial date
      const d = XLSX.SSF.parse_date_code(v);
      if (!d) return null;
      return new Date(d.y, d.m - 1, d.d);
    }
    if (typeof v === "string") {
      // Try DD/MM/YYYY first
      const m = v.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
      if (m) {
        let [_, d, mo, y] = m;
        y = parseInt(y, 10);
        if (y < 100) y += 2000;
        return new Date(y, parseInt(mo, 10) - 1, parseInt(d, 10));
      }
      const dt = new Date(v);
      return isNaN(dt.getTime()) ? null : dt;
    }
    return null;
  }

  /** Parse employee details file. Returns Map<empId, {empId, startDate, endDate, daysPerWeek, deptNumber}>. */
  function parseEmployees(rows) {
    const out = new Map();
    // Header row 1, data starts row 2 (index 1)
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const empId = r[0];
      if (empId === null || empId === undefined || empId === "") continue;
      out.set(empId, {
        empId,
        startDate: toDate(r[64]),
        endDate: toDate(r[65]),
        daysPerWeek: r[75] === 5 ? 5 : r[75] === 6 ? 6 : (toNumber(r[75]) === 5 ? 5 : toNumber(r[75]) === 6 ? 6 : null),
        deptNumber: r[71] !== null && r[71] !== undefined && r[71] !== "" ? r[71] : null, // col 72: מספר מחלקה
      });
    }
    return out;
  }

  /** Parse a daily attendance file (31.3 / 9.4 / April). Returns Map<empId, {empId, firstName, lastName, work, vacation, sick, military}>. */
  function parseDaily(rows) {
    const out = new Map();
    // Header row 2 (index 1), data starts row 3 (index 2)
    for (let i = 2; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const empId = r[11]; // col 12
      if (empId === null || empId === undefined || empId === "" || empId === 0) continue;
      out.set(empId, {
        empId,
        firstName: r[4] || "",
        lastName: r[5] || "",
        deptName: r[6] || "", // col 7: מחלקה (name)
        work: toNumber(r[14]),
        vacation: toNumber(r[23]),
        sick: toNumber(r[24]),
        military: toNumber(r[25]),
      });
    }
    return out;
  }

  /** Parse 3-month attendance. Returns Map<empId, {workDays, workHours, marchBalance}> (aggregated). */
  function parseThreeMonths(rows) {
    const map = new Map();
    // Header row 1, data starts row 2 (index 1)
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const empId = r[2]; // col 3
      if (empId === null || empId === undefined || empId === "") continue;
      const month = toNumber(r[5]); // col 6
      const workDays = toNumber(r[13]); // col 14: י"ע בפועל
      const workHours = toNumber(r[15]); // col 16: ש"ע בפועל
      const balance = r[18]; // col 19: חופשה - יתרה (raw, may be null)

      let entry = map.get(empId);
      if (!entry) {
        entry = { empId, workDays: 0, workHours: 0, marchBalance: null, _months: {} };
        map.set(empId, entry);
      }
      entry.workDays += workDays;
      entry.workHours += workHours;
      entry._months[month] = { workDays, workHours, balance };
      // Take the latest available balance, prefer March (3), else 2, else 1
      if (month === 3 && balance !== null && balance !== undefined && balance !== "") {
        entry.marchBalance = toNumber(balance);
      }
    }
    // For employees without explicit March balance, fall back to latest month available
    for (const e of map.values()) {
      if (e.marchBalance === null) {
        for (const m of [3, 2, 1]) {
          const md = e._months[m];
          if (md && md.balance !== null && md.balance !== undefined && md.balance !== "") {
            e.marchBalance = toNumber(md.balance);
            break;
          }
        }
      }
      if (e.marchBalance === null) e.marchBalance = 0;
      delete e._months;
    }
    return map;
  }

  // ===== Detailed shift report (יום העצמאות) =====
  // Per-branch files: same column NAMES but different positions, so we look up by header.

  // Header aliases for required & optional fields. First exact match wins,
  // then fallback to partial match. Order matters — payroll-system ID first
  // so we don't accidentally match the internal branch employee number ("מספר עובד").
  const SHIFT_HEADER_ALIASES = {
    empId:     ["מס' עובד במע' שכר", "מס' עובד", "מס עובד", "קוד עובד", "מספר עובד"],
    entryDate: ["תאריך כניסה", "תאריך משמרת", "תאריך"],
    entryTime: ["שעת התחלה", "שעת כניסה", "כניסה"],
    exitTime:  ["שעת סיום", "שעת יציאה", "יציאה"],
    fullName:  ["שם עובד", "שם העובד"],
    firstName: ["שם פרטי"],
    lastName:  ["שם משפחה"],
    deptName:  ["מחלקה", "סניף", "ענף"],
    vacation:  ["ימי חופש"],
    sick:      ["ימי מחלה"],
    military:  ["ימי מילואים"],
  };
  // Required for SHIFT rows (with times). Absence rows (vacation/sick/military)
  // only need empId+entryDate (no times).
  const SHIFT_REQUIRED = ["empId", "entryDate", "entryTime", "exitTime"];

  /** Pick the first row that has at least 3 of the alias keywords — that's the header row. */
  function findHeaderRow(rows) {
    let best = -1, bestScore = 0;
    const keywords = Object.values(SHIFT_HEADER_ALIASES).flat();
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const r = rows[i];
      if (!r) continue;
      let score = 0;
      const flat = r.map((c) => (c == null ? "" : String(c).trim()));
      for (const kw of keywords) {
        if (flat.some((v) => v === kw || v.includes(kw))) score++;
      }
      if (score > bestScore) {
        bestScore = score;
        best = i;
      }
    }
    return bestScore >= 3 ? best : -1;
  }

  /** Build {field: colIdx} map from a header row using aliases. */
  function mapShiftColumns(headerRow) {
    const result = {};
    const flat = headerRow.map((c) => (c == null ? "" : String(c).trim()));
    for (const [field, aliases] of Object.entries(SHIFT_HEADER_ALIASES)) {
      for (const alias of aliases) {
        let idx = flat.findIndex((v) => v === alias);
        if (idx < 0) idx = flat.findIndex((v) => v.includes(alias));
        if (idx >= 0) {
          result[field] = idx;
          break;
        }
      }
    }
    return result;
  }

  /** Parse Excel time-of-day. Returns { h, m } or null.
   *  Handles: number 0..1 (Excel time fraction), "HH:MM", "HH:MM:SS", Date. */
  function toTimeOfDay(v) {
    if (v === null || v === undefined || v === "") return null;
    if (v instanceof Date) {
      return { h: v.getHours(), m: v.getMinutes() };
    }
    if (typeof v === "number") {
      // Excel time fraction (0–1) OR full serial — we only want HH:MM portion
      const frac = v - Math.floor(v);
      const totalMin = Math.round(frac * 24 * 60);
      const h = Math.floor(totalMin / 60) % 24;
      const m = totalMin % 60;
      return { h, m };
    }
    if (typeof v === "string") {
      const m = v.trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
      if (!m) return null;
      return { h: parseInt(m[1], 10), m: parseInt(m[2], 10) };
    }
    return null;
  }

  /** Combine a date and a time-of-day into one Date. */
  function combineDateTime(date, tod) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), tod.h, tod.m, 0);
  }

  /**
   * Parse one detailed shift report file.
   * @param {Array<Array>} rows - sheet rows
   * @param {string} fileLabel - human-friendly label (filename without extension), used in errors and as branch fallback
   * @returns {Map<empId, {empId, branches:Set<string>, firstName?, lastName?, deptName?, shifts: Shift[]}>}
   * @throws Error with file label + missing column on malformed file.
   */
  function parseShiftReport(rows, fileLabel) {
    const headerIdx = findHeaderRow(rows);
    if (headerIdx < 0) {
      throw new Error(`קובץ "${fileLabel}": לא נמצאה שורת כותרת מזוהה (מצופה: מספר עובד, תאריך כניסה, שעת כניסה, שעת יציאה).`);
    }
    const colMap = mapShiftColumns(rows[headerIdx]);
    const missing = SHIFT_REQUIRED.filter((f) => colMap[f] === undefined);
    if (missing.length) {
      const labelMap = {
        empId: "מספר עובד",
        entryDate: "תאריך כניסה",
        entryTime: "שעת כניסה",
        exitTime: "שעת יציאה",
      };
      const headers = rows[headerIdx].filter((c) => c != null && String(c).trim() !== "").join(", ");
      throw new Error(`קובץ "${fileLabel}": חסרה עמודה "${labelMap[missing[0]]}" (כותרות בקובץ: ${headers}).`);
    }

    const out = new Map();
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const empId = r[colMap.empId];
      if (empId === null || empId === undefined || empId === "" || empId === 0) continue;
      // Skip footer summary rows (no full name AND no proper date)
      const fullNameRaw = colMap.fullName !== undefined ? (r[colMap.fullName] || "") : "";
      const hasName = colMap.fullName !== undefined ? !!fullNameRaw :
                      ((colMap.firstName !== undefined && r[colMap.firstName]) || (colMap.lastName !== undefined && r[colMap.lastName]));
      const date = toDate(r[colMap.entryDate]);
      if (!hasName && !date) continue;
      if (!date) continue;
      const tIn  = toTimeOfDay(r[colMap.entryTime]);
      const tOut = toTimeOfDay(r[colMap.exitTime]);
      const vacation = colMap.vacation !== undefined ? toNumber(r[colMap.vacation]) : 0;
      const sick     = colMap.sick     !== undefined ? toNumber(r[colMap.sick])     : 0;
      const military = colMap.military !== undefined ? toNumber(r[colMap.military]) : 0;

      // Get-or-create employee entry
      let entry = out.get(empId);
      if (!entry) {
        let fn = "", ln = "";
        if (fullNameRaw) {
          const parts = String(fullNameRaw).trim().split(/\s+/);
          fn = parts[0] || "";
          ln = parts.slice(1).join(" ");
        } else {
          fn = colMap.firstName !== undefined ? (r[colMap.firstName] || "") : "";
          ln = colMap.lastName  !== undefined ? (r[colMap.lastName]  || "") : "";
        }
        entry = {
          empId,
          branches: new Set(),
          firstName: fn,
          lastName: ln,
          fullName: fullNameRaw || `${fn} ${ln}`.trim(),
          deptName: colMap.deptName !== undefined ? (r[colMap.deptName] || "") : "",
          shifts: [],
          absences: [], // [{date: Date, type: "vacation"|"sick"|"military"}]
        };
        out.set(empId, entry);
      }
      entry.branches.add(fileLabel);

      if (tIn && tOut) {
        // Shift row
        const entryAt = combineDateTime(date, tIn);
        let exitAt   = combineDateTime(date, tOut);
        if (exitAt.getTime() <= entryAt.getTime()) {
          exitAt = new Date(exitAt.getTime() + 24 * 3600 * 1000);
        }
        entry.shifts.push({ entryAt, exitAt, fileLabel });
      } else if (vacation > 0) {
        entry.absences.push({ date, type: "vacation" });
      } else if (sick > 0) {
        entry.absences.push({ date, type: "sick" });
      } else if (military > 0) {
        entry.absences.push({ date, type: "military" });
      }
      // else: empty row with no times and no absence markers — skip silently
    }
    return out;
  }

  /** Merge multiple per-file shift maps into one combined map. */
  function combineShiftReports(perFileMaps) {
    const combined = new Map();
    for (const m of perFileMaps) {
      for (const [empId, rec] of m.entries()) {
        let cur = combined.get(empId);
        if (!cur) {
          cur = {
            empId,
            firstName: rec.firstName || "",
            lastName: rec.lastName || "",
            fullName: rec.fullName || "",
            deptName: rec.deptName || "",
            branches: new Set(),
            shifts: [],
            absences: [],
          };
          combined.set(empId, cur);
        }
        rec.branches.forEach((b) => cur.branches.add(b));
        cur.shifts.push(...rec.shifts);
        if (rec.absences && rec.absences.length) cur.absences.push(...rec.absences);
        if (!cur.firstName && rec.firstName) cur.firstName = rec.firstName;
        if (!cur.lastName  && rec.lastName)  cur.lastName  = rec.lastName;
        if (!cur.fullName  && rec.fullName)  cur.fullName  = rec.fullName;
        if (!cur.deptName  && rec.deptName)  cur.deptName  = rec.deptName;
      }
    }
    return combined;
  }

  return {
    readSheet,
    parseEmployees,
    parseDaily,
    parseThreeMonths,
    parseShiftReport,
    combineShiftReports,
    toDate,
    toNumber,
  };
})();
