/**
 * detector.js - identifies which of the 5 input slots a given Excel file belongs to,
 * based on content signatures (not filename). Returns one of:
 *   "employees"    - פרטי עובדים (150-col file, header row 1, has "תאריך תחילת עבודה")
 *   "three_months" - נוכחות 3 חודשים (has "מספר חודש" + "חופשה - יתרה")
 *   "march31"      - daily attendance for 31.3
 *   "april9"       - daily attendance for 9.4
 *   "april"        - daily attendance for the whole April month (relevance filter)
 *   null           - unknown
 *
 * Daily files all share the same shape (header row 2, "מס' עובד במע' שכר" at col 12).
 * They are differentiated by the date range in the title row (row 1).
 */

const Detector = (function () {
  function read(buffer) {
    const wb = XLSX.read(buffer, { type: "array", cellDates: true });
    return wb;
  }

  function getRowsForSheet(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  function rowContains(row, needle) {
    if (!row) return false;
    return row.some((c) => c !== null && c !== undefined && String(c).includes(needle));
  }

  function detectFromRows(rows) {
    if (!rows || rows.length === 0) return null;

    // 1) Employee details: very wide (150 cols), header row 1 has "מספר עובד" at col 0
    //    and "תאריך תחילת עבודה" at col 64.
    if (rows[0] && rows[0].length >= 100) {
      const h1 = rows[0];
      if (h1[0] === "מספר עובד" && rowContains(h1, "תאריך תחילת עבודה")) {
        return { type: "employees" };
      }
    }

    // 2) Three-month attendance: header row 1 has both "מספר חודש" and "חופשה - יתרה"
    if (rows[0] && rowContains(rows[0], "מספר חודש") && rowContains(rows[0], "חופשה")) {
      return { type: "three_months" };
    }

    // 3) Daily attendance: header row 2 has "מס' עובד במע' שכר".
    //    Title row 1 contains "דוח שכר" + a date range like "31/03/26 - 31/03/26".
    if (rows.length >= 2 && rows[1]) {
      const h2 = rows[1];
      const isDaily = h2.some(
        (c) => c && (String(c).includes("עובד במע") || String(c).includes("מס' עובד במע"))
      );
      if (isDaily) {
        const titleCell = (rows[0] || []).find((c) => c && String(c).includes("דוח שכר"));
        const title = titleCell ? String(titleCell) : "";
        const m = title.match(
          /(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*[-–]\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})/
        );
        if (m) {
          const d1 = m[1].padStart(2, "0"),
            mo1 = m[2].padStart(2, "0"),
            d2 = m[4].padStart(2, "0"),
            mo2 = m[5].padStart(2, "0");
          const startEqEnd = d1 === d2 && mo1 === mo2;
          if (startEqEnd) {
            if (d1 === "31" && mo1 === "03") return { type: "march31", date: `${d1}/${mo1}` };
            if (d1 === "09" && mo1 === "04") return { type: "april9", date: `${d1}/${mo1}` };
            return { type: "daily_unknown_single", date: `${d1}/${mo1}` };
          }
          // multi-day range
          if (mo1 === "04" || mo2 === "04") return { type: "april", range: `${d1}/${mo1}-${d2}/${mo2}` };
          return { type: "daily_unknown_range", range: `${d1}/${mo1}-${d2}/${mo2}` };
        }
        return { type: "daily_unknown" };
      }
    }

    return null;
  }

  /** Detect across all sheets; returns the first match found. */
  function detect(buffer) {
    const wb = read(buffer);
    for (const sn of wb.SheetNames) {
      const rows = getRowsForSheet(wb, sn);
      const r = detectFromRows(rows);
      if (r) return Object.assign({ sheetName: sn }, r);
    }
    return null;
  }

  /** Friendly Hebrew label per slot key. */
  const LABEL = {
    employees: "פרטי עובדים",
    three_months: "נוכחות 3 חודשים",
    april: "נוכחות אפריל",
    march31: "נוכחות 31.3",
    april9: "נוכחות 9.4",
  };

  return { detect, detectFromRows, LABEL };
})();
