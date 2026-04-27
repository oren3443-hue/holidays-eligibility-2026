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

  /** Parse employee details file. Returns Map<empId, {empId, startDate, endDate, daysPerWeek}>. */
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

  return {
    readSheet,
    parseEmployees,
    parseDaily,
    parseThreeMonths,
    toDate,
    toNumber,
  };
})();
