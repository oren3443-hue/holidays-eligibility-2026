/**
 * exporter.js - builds the output Excel file from result rows.
 */

const Exporter = (function () {
  const COLUMNS = [
    { key: "empId", header: "מספר עובד" },
    { key: "firstName", header: "שם פרטי" },
    { key: "lastName", header: "שם משפחה" },
    { key: "daysPerWeek", header: "מתכונת (5/6)" },
    { key: "startDate", header: "תאריך תחילת עבודה", fmt: "date" },
    { key: "tenureOk", header: "ותק תקין?", fmt: "yesno" },
    { key: "status31", header: "סטטוס 31.3" },
    { key: "eligibleA", header: "זכאות פסח א", fmt: "yesno" },
    { key: "reasonA", header: "סיבה - פסח א" },
    { key: "status9", header: "סטטוס 9.4" },
    { key: "eligibleB", header: "זכאות פסח ב", fmt: "yesno" },
    { key: "reasonB", header: "סיבה - פסח ב" },
    { key: "totalWorkHours", header: 'ש"ע בפועל (ינ-מרץ)' },
    { key: "totalWorkDays", header: 'י"ע בפועל (ינ-מרץ)' },
    { key: "avgHoursPerDay", header: "ממוצע שעות יומי" },
    { key: "vacationBalance", header: "יתרת חופש מרץ" },
    { key: "daysToUse", header: "ימי חופש לניצול" },
    { key: "hours_1_4", header: "שעות 1.4" },
    { key: "hours_3_4", header: "שעות 3.4" },
    { key: "hours_5_4", header: "שעות 5.4" },
    { key: "hours_6_4", header: "שעות 6.4" },
    { key: "hours_7_4", header: "שעות 7.4" },
    { key: "totalHolidayHours", header: 'סה"כ שעות חופש' },
  ];

  function formatDate(d) {
    if (!d) return "";
    if (!(d instanceof Date)) return String(d);
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yy = d.getFullYear();
    return `${dd}/${mm}/${yy}`;
  }

  function formatCell(val, fmt) {
    if (val === null || val === undefined) return "";
    if (fmt === "date") return formatDate(val);
    if (fmt === "yesno") return val ? "כן" : "לא";
    return val;
  }

  function buildAOA(rows) {
    const aoa = [COLUMNS.map((c) => c.header)];
    for (const r of rows) {
      aoa.push(COLUMNS.map((c) => formatCell(r[c.key], c.fmt)));
    }
    return aoa;
  }

  function exportXlsx(rows, filename) {
    const aoa = buildAOA(rows);
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // RTL view
    if (!ws["!views"]) ws["!views"] = [{}];
    ws["!views"][0].rightToLeft = true;

    // Column widths (rough estimate)
    ws["!cols"] = COLUMNS.map((c) => ({ wch: Math.max(10, Math.min(28, c.header.length + 4)) }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "זכאות פסח");
    const fname = filename || `זכאות-פסח-${todayStr()}.xlsx`;
    XLSX.writeFile(wb, fname);
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
