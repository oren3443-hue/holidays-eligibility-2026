/**
 * exporter.js - builds a styled, RTL Hebrew Excel from result rows.
 * Uses xlsx-js-style (a drop-in fork of SheetJS) for cell styling.
 */

const Exporter = (function () {
  // Each column: { key, header, fmt, width, type? }
  // type: "number" | "date" | "string" — drives Excel cell type & number format.
  const COLUMNS = [
    { key: "empId", header: "מספר עובד", type: "number", width: 11 },
    { key: "firstName", header: "שם פרטי", type: "string", width: 14 },
    { key: "lastName", header: "שם משפחה", type: "string", width: 14 },
    { key: "deptName", header: "מחלקה", type: "string", width: 16 },
    { key: "deptNumber", header: "מס' מחלקה (מיכפל)", type: "number", width: 14 },
    { key: "daysPerWeek", header: "מתכונת (5/6)", type: "number", width: 11 },
    { key: "startDate", header: "תאריך תחילת עבודה", type: "date", width: 16 },
    { key: "tenureOk", header: "ותק תקין?", fmt: "yesno", width: 10 },
    { key: "status31", header: "סטטוס 31.3", type: "string", width: 12 },
    { key: "eligibleA", header: "זכאות פסח א", fmt: "yesno", width: 12 },
    { key: "reasonA", header: "סיבה - פסח א", type: "string", width: 18 },
    { key: "status9", header: "סטטוס 9.4", type: "string", width: 12 },
    { key: "eligibleB", header: "זכאות פסח ב", fmt: "yesno", width: 12 },
    { key: "reasonB", header: "סיבה - פסח ב", type: "string", width: 18 },
    { key: "totalWorkHours", header: 'ש"ע בפועל (ינ-מרץ)', type: "number", numFmt: "0.00", width: 14 },
    { key: "totalWorkDays", header: 'י"ע בפועל (ינ-מרץ)', type: "number", width: 14 },
    { key: "avgHoursPerDay", header: "ממוצע שעות יומי", type: "number", numFmt: "0.00", width: 14 },
    { key: "vacationBalance", header: "יתרת חופש מרץ", type: "number", numFmt: "0.00", width: 14 },
    { key: "daysToUse", header: "ימי חופש לניצול", type: "number", width: 14 },
    { key: "hours_1_4", header: "שעות 1.4", type: "number", numFmt: "0.00", width: 9 },
    { key: "hours_3_4", header: "שעות 3.4", type: "number", numFmt: "0.00", width: 9 },
    { key: "hours_5_4", header: "שעות 5.4", type: "number", numFmt: "0.00", width: 9 },
    { key: "hours_6_4", header: "שעות 6.4", type: "number", numFmt: "0.00", width: 9 },
    { key: "hours_7_4", header: "שעות 7.4", type: "number", numFmt: "0.00", width: 9 },
    { key: "totalHolidayHours", header: 'סה"כ שעות חופש', type: "number", numFmt: "0.00", width: 14 },
    { key: "holidayDaysCount", header: "ימי חג", type: "number", width: 9 },
    { key: "holidayPayHours", header: "שעות רכיב חג", type: "number", numFmt: "0.00", width: 13 },
  ];

  // Import file for the payroll software (קובץ קליטה).
  // CRITICAL: zero values must be left empty (null), not 0 — that's how the software
  // knows to skip importing that field for that employee.
  const IMPORT_COLUMNS = [
    { key: "empId", header: "מספר עובד", type: "number", width: 11, blankIfZero: false },
    { key: "fullName", header: "שם העובד", type: "string", width: 22, blankIfZero: false },
    { key: "deptNumber", header: "מחלקה (מיכפל)", type: "number", width: 14, blankIfZero: false },
    { key: "daysToUse", header: "ימי חופש לניצול", type: "number", width: 14, blankIfZero: true },
    { key: "totalHolidayHours", header: "שעות רכיב חופשה", type: "number", numFmt: "0.00", width: 16, blankIfZero: true },
    { key: "holidayDaysCount", header: "ימי חג", type: "number", width: 9, blankIfZero: true },
    { key: "holidayPayHours", header: "שעות חג לתשלום", type: "number", numFmt: "0.00", width: 16, blankIfZero: true },
  ];

  // ===== Yom HaAtzmaut (Independence Day) =====
  const COLUMNS_INDEPENDENCE = [
    { key: "empId", header: "מספר עובד", type: "number", width: 11 },
    { key: "firstName", header: "שם פרטי", type: "string", width: 14 },
    { key: "lastName", header: "שם משפחה", type: "string", width: 14 },
    { key: "deptName", header: "מחלקה", type: "string", width: 16 },
    { key: "deptNumber", header: "מס' מחלקה (מיכפל)", type: "number", width: 14 },
    { key: "branches", header: "סניפים", type: "string", width: 18 },
    { key: "daysPerWeek", header: "מתכונת (5/6)", type: "number", width: 11 },
    { key: "startDate", header: "תאריך תחילת עבודה", type: "date", width: 16 },
    { key: "tenureOk", header: "ותק תקין?", fmt: "yesno", width: 10 },
    { key: "presence_21_4", header: "נוכחות 21.4", type: "string", width: 22 },
    { key: "hours_21_4", header: "שעות 21.4", type: "number", numFmt: "0.00", width: 11 },
    { key: "presence_22_4", header: "נוכחות 22.4", type: "string", width: 22 },
    { key: "hours_22_4", header: "שעות 22.4", type: "number", numFmt: "0.00", width: 11 },
    { key: "presence_23_4", header: "נוכחות 23.4", type: "string", width: 22 },
    { key: "hours_23_4", header: "שעות 23.4", type: "number", numFmt: "0.00", width: 11 },
    { key: "presentDayBefore", header: "נכח 21.4 (ש/ח/מ)", fmt: "yesno", width: 14 },
    { key: "presentDayAfter", header: "נכח 23.4 (ש/ח/מ)", fmt: "yesno", width: 14 },
    { key: "hoursWorkedInWindow", header: "שעות בחלון 20:00-20:00", type: "number", numFmt: "0.00", width: 18 },
    { key: "modeLabel", header: "מצב", type: "string", width: 16 },
    { key: "totalWorkHours", header: 'ש"ע בפועל (3 חודשים)', type: "number", numFmt: "0.00", width: 16 },
    { key: "totalWorkDays", header: 'י"ע בפועל (3 חודשים)', type: "number", width: 14 },
    { key: "avgHoursPerDay", header: "ממוצע שעות יומי", type: "number", numFmt: "0.00", width: 14 },
    { key: "holidayDaysCount", header: "ימי חג", type: "number", width: 9 },
    { key: "holidayPayHours", header: "שעות חג (תשלום)", type: "number", numFmt: "0.00", width: 14 },
    { key: "extraHolidayHours", header: "תוספת שעות חג 100 אחוז", type: "number", numFmt: "0.00", width: 22 },
    { key: "reason", header: "הערה", type: "string", width: 22 },
  ];

  // Import file for Independence Day. Format matches the Shavuot sample exactly:
  //   row 1: [companyNumber, year, month, null, null, null]
  //   row 2: 6 column headers
  //   row 3+: data, with zero values written as truly empty cells.
  const IMPORT_COLUMNS_INDEPENDENCE = [
    { key: "empId", header: "מספר עובד", type: "number", width: 11, blankIfZero: false },
    { key: "fullName", header: "שם העובד", type: "string", width: 22, blankIfZero: false },
    { key: "deptNumber", header: "מחלקה (מיכפל)", type: "number", width: 14, blankIfZero: false },
    { key: "holidayDaysCount", header: "ימי חג", type: "number", width: 9, blankIfZero: true },
    { key: "holidayPayHours", header: "שעות חג לתשלום", type: "number", numFmt: "0.00", width: 16, blankIfZero: true },
    { key: "extraHolidayHours", header: "תוספת שעות חג 100 אחוז", type: "number", numFmt: "0.00", width: 22, blankIfZero: true },
  ];

  // Style palette
  const COLOR = {
    headerBg: "4F46E5",   // indigo-600
    headerFg: "FFFFFF",
    rowAlt: "F3F4F6",     // gray-100 (zebra)
    border: "D1D5DB",     // gray-300
    yes: "D1FAE5",        // green-100
    yesText: "065F46",    // green-800
    no: "FEE2E2",         // red-100
    noText: "991B1B",     // red-800
    summaryBg: "EEF2FF",  // indigo-50
    summaryBold: "3730A3",// indigo-800
  };

  function thinBorder() {
    const side = { style: "thin", color: { rgb: COLOR.border } };
    return { top: side, bottom: side, left: side, right: side };
  }

  function headerStyle() {
    return {
      font: { bold: true, color: { rgb: COLOR.headerFg }, sz: 11 },
      fill: { fgColor: { rgb: COLOR.headerBg } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true, readingOrder: 2 },
      border: thinBorder(),
    };
  }

  function bodyStyle({ alt, key, value, columns }) {
    const base = {
      alignment: { horizontal: "right", vertical: "center", readingOrder: 2 },
      border: thinBorder(),
      font: { sz: 10 },
    };
    if (alt) base.fill = { fgColor: { rgb: COLOR.rowAlt } };

    const col = (columns || COLUMNS).find((c) => c.key === key);

    // Yes/No coloring (driven by fmt:"yesno" rather than a hardcoded key list)
    if (col && col.fmt === "yesno") {
      const yes = value === true || value === "כן";
      base.fill = { fgColor: { rgb: yes ? COLOR.yes : COLOR.no } };
      base.font = { sz: 10, bold: true, color: { rgb: yes ? COLOR.yesText : COLOR.noText } };
      base.alignment.horizontal = "center";
    }

    // Numeric cells: align center for compactness
    if (col && col.type === "number") {
      base.alignment.horizontal = "center";
      if (col.numFmt) base.numFmt = col.numFmt;
    }
    if (col && col.type === "date") {
      base.alignment.horizontal = "center";
      base.numFmt = "dd/mm/yyyy";
    }

    return base;
  }

  function buildCell(value, fmt, type) {
    if (value === null || value === undefined || value === "") {
      return null; // truly empty cell — caller skips writing it
    }
    if (fmt === "yesno") {
      return { v: value ? "כן" : "לא", t: "s" };
    }
    if (type === "date") {
      if (value instanceof Date) return { v: value, t: "d" };
      return { v: String(value), t: "s" };
    }
    if (type === "number") {
      const n = typeof value === "number" ? value : parseFloat(value);
      if (isNaN(n)) return null;
      return { v: n, t: "n" };
    }
    return { v: String(value), t: "s" };
  }

  function colLetter(idx /* 0-based */) {
    let n = idx + 1;
    let s = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /** Build a styled sheet from rows + a column spec.
   *  opts:
   *    - styled (bool, default true) - apply colors/borders/zebra
   *    - metaRow (array | null) - prepend an optional metadata row (row 1 in מיכפל imports)
   *    - autofilter (bool, default true)
   *    - freeze (bool, default true)
   *    - rtl (bool, default true)
   */
  function buildSheet(rows, columns, opts) {
    opts = opts || {};
    const styled = opts.styled !== false;
    const metaRow = opts.metaRow || null;
    const autofilter = opts.autofilter !== false;
    const freeze = opts.freeze !== false;
    const rtl = opts.rtl !== false;

    const ws = {};
    const lastColIdx = columns.length - 1;
    const headerRowIdx = metaRow ? 2 : 1; // 1-based row number for the header row
    const dataStartRow = headerRowIdx + 1;
    const lastRowIdx = rows.length + headerRowIdx;

    // Optional metadata row at row 1 (used by payroll software to identify the import)
    if (metaRow) {
      metaRow.forEach((val, ci) => {
        if (val === null || val === undefined || val === "") return; // truly empty
        const addr = colLetter(ci) + "1";
        ws[addr] = typeof val === "number"
          ? { v: val, t: "n" }
          : { v: String(val), t: "s" };
      });
    }

    // Header row
    columns.forEach((col, ci) => {
      const addr = colLetter(ci) + headerRowIdx;
      ws[addr] = { v: col.header, t: "s", s: styled ? headerStyle() : undefined };
    });

    // Data rows. NOTE: when buildCell returns null we DO NOT write the cell at all —
    // truly empty cells are essential for the payroll software's import logic
    // (a blank string "" would still count as a value).
    rows.forEach((r, ri) => {
      const alt = ri % 2 === 1;
      columns.forEach((col, ci) => {
        const addr = colLetter(ci) + (ri + dataStartRow);
        let value = r[col.key];

        // Per-column "blankIfZero" rule (used by the import file)
        if (col.blankIfZero && (value === 0 || value === "0")) value = null;

        const cell = buildCell(value, col.fmt, col.type);
        if (cell === null) return; // skip — leaves the cell truly empty
        if (col.numFmt && cell.t === "n") cell.z = col.numFmt;
        if (styled) cell.s = bodyStyle({ alt, key: col.key, value: cell.v, columns });
        ws[addr] = cell;
      });
    });

    ws["!ref"] = `A1:${colLetter(lastColIdx)}${lastRowIdx}`;
    ws["!cols"] = columns.map((c) => ({ wch: c.width || 12 }));
    if (styled) ws["!rows"] = metaRow ? [undefined, { hpt: 28 }] : [{ hpt: 28 }];
    if (autofilter) {
      ws["!autofilter"] = {
        ref: `A${headerRowIdx}:${colLetter(lastColIdx)}${lastRowIdx}`,
      };
    }
    if (freeze) ws["!freeze"] = { ySplit: headerRowIdx };
    ws["!views"] = [
      rtl
        ? { rightToLeft: true, state: freeze ? "frozen" : undefined, ySplit: freeze ? headerRowIdx : undefined }
        : { state: freeze ? "frozen" : undefined, ySplit: freeze ? headerRowIdx : undefined },
    ];
    return ws;
  }

  function buildWorkbook(ws, sheetName) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    if (!wb.Workbook) wb.Workbook = {};
    if (!wb.Workbook.Views) wb.Workbook.Views = [{}];
    wb.Workbook.Views[0].RTL = true;
    return wb;
  }

  function exportXlsx(rows, filename) {
    const ws = buildSheet(rows, COLUMNS, { styled: true });
    const wb = buildWorkbook(ws, "זכאות פסח");
    const fname = filename || `זכאות-פסח-${todayStr()}.xlsx`;
    XLSX.writeFile(wb, fname, { cellStyles: true });
  }

  /** Default מיכפל import metadata: [companyNumber, year, month].
   *  - companyNumber: 10 = מאפיית אורן משי (per sample file).
   *  - year/month: pay period for which the import applies. Pesach 2026 → 4/2026. */
  const IMPORT_META = { companyNumber: 10, year: 2026, month: 4 };

  /** Export the import file for the payroll software (קובץ קליטה).
   *  Layout matches the מיכפל sample file:
   *    Row 1: [companyNumber, year, month, blank, blank, blank, blank]
   *    Row 2: column headers
   *    Row 3+: data
   *  Cells with value 0 on importable columns are written as TRULY EMPTY cells
   *  so the payroll software skips that field for that employee. */
  function exportImportXlsx(rows, filename, meta) {
    const m = meta || IMPORT_META;
    const metaRow = [m.companyNumber, m.year, m.month, null, null, null, null];
    const ws = buildSheet(rows, IMPORT_COLUMNS, {
      styled: false,        // payroll software prefers a plain file
      metaRow,
      autofilter: false,
      freeze: false,
      rtl: true,
    });
    const wb = buildWorkbook(ws, "גיליון1"); // match sample's sheet name
    const fname = filename || `קליטה-פסח-${m.year}-${String(m.month).padStart(2,"0")}.xlsx`;
    XLSX.writeFile(wb, fname);
  }

  /** Independence Day: full report. */
  function exportXlsxIndependence(rows, filename) {
    const ws = buildSheet(rows, COLUMNS_INDEPENDENCE, { styled: true });
    const wb = buildWorkbook(ws, "זכאות יום העצמאות");
    const fname = filename || `זכאות-יום-העצמאות-${todayStr()}.xlsx`;
    XLSX.writeFile(wb, fname, { cellStyles: true });
  }

  /** Independence Day: payroll import file (מיכפל). */
  function exportImportXlsxIndependence(rows, filename, meta) {
    const m = meta || { companyNumber: 10, year: 2026, month: 4 };
    const metaRow = [m.companyNumber, m.year, m.month, null, null, null];
    const ws = buildSheet(rows, IMPORT_COLUMNS_INDEPENDENCE, {
      styled: false,
      metaRow,
      autofilter: false,
      freeze: false,
      rtl: true,
    });
    const wb = buildWorkbook(ws, "גיליון1");
    const fname = filename || `קליטה-יום-העצמאות-${m.year}-${String(m.month).padStart(2,"0")}.xlsx`;
    XLSX.writeFile(wb, fname);
  }

  function todayStr() {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yy = d.getFullYear();
    return `${yy}-${mm}-${dd}`;
  }

  return {
    exportXlsx,
    exportImportXlsx,
    exportXlsxIndependence,
    exportImportXlsxIndependence,
    COLUMNS,
    IMPORT_COLUMNS,
    COLUMNS_INDEPENDENCE,
    IMPORT_COLUMNS_INDEPENDENCE,
  };
})();
