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

  function bodyStyle({ alt, key, value }) {
    const base = {
      alignment: { horizontal: "right", vertical: "center", readingOrder: 2 },
      border: thinBorder(),
      font: { sz: 10 },
    };
    if (alt) base.fill = { fgColor: { rgb: COLOR.rowAlt } };

    // Yes/No coloring on eligibility + tenure cells
    if (key === "eligibleA" || key === "eligibleB" || key === "tenureOk") {
      const yes = value === true || value === "כן";
      base.fill = { fgColor: { rgb: yes ? COLOR.yes : COLOR.no } };
      base.font = { sz: 10, bold: true, color: { rgb: yes ? COLOR.yesText : COLOR.noText } };
      base.alignment.horizontal = "center";
    }

    // Numeric cells: align center for compactness
    const col = COLUMNS.find((c) => c.key === key);
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
      return { v: "", t: "s" };
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
      if (isNaN(n)) return { v: "", t: "s" };
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

  function exportXlsx(rows, filename) {
    const ws = {};
    const lastColIdx = COLUMNS.length - 1;
    const lastRowIdx = rows.length; // header + rows

    // Header row
    COLUMNS.forEach((col, ci) => {
      const addr = colLetter(ci) + "1";
      ws[addr] = { v: col.header, t: "s", s: headerStyle() };
    });

    // Data rows
    rows.forEach((r, ri) => {
      const alt = ri % 2 === 1;
      COLUMNS.forEach((col, ci) => {
        const addr = colLetter(ci) + (ri + 2);
        const cell = buildCell(r[col.key], col.fmt, col.type);
        cell.s = bodyStyle({ alt, key: col.key, value: cell.v });
        ws[addr] = cell;
      });
    });

    // Sheet metadata
    ws["!ref"] = `A1:${colLetter(lastColIdx)}${lastRowIdx + 1}`;
    ws["!cols"] = COLUMNS.map((c) => ({ wch: c.width || 12 }));
    ws["!rows"] = [{ hpt: 28 }]; // taller header row
    ws["!autofilter"] = { ref: ws["!ref"] };
    ws["!freeze"] = { ySplit: 1 };
    ws["!views"] = [{ rightToLeft: true, state: "frozen", ySplit: 1 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "זכאות פסח");

    // Workbook-level: open as RTL by default
    if (!wb.Workbook) wb.Workbook = {};
    if (!wb.Workbook.Views) wb.Workbook.Views = [{}];
    wb.Workbook.Views[0].RTL = true;

    const fname = filename || `זכאות-פסח-${todayStr()}.xlsx`;
    XLSX.writeFile(wb, fname, { cellStyles: true });
  }

  function todayStr() {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yy = d.getFullYear();
    return `${yy}-${mm}-${dd}`;
  }

  return { exportXlsx, COLUMNS };
})();
