/**
 * main.js - UI orchestration.
 */

(function () {
  // State
  const state = {
    files: {
      employees: null,
      three_months: null,
      april: null,
      march31: null,
      april9: null,
    },
    rows: null,
  };

  const FILE_KEYS = ["employees", "three_months", "april", "march31", "april9"];

  function $(sel) { return document.querySelector(sel); }
  function $$(sel) { return Array.from(document.querySelectorAll(sel)); }

  function setMessage(text, kind) {
    const el = $("#message");
    el.textContent = text || "";
    el.className = "message" + (kind ? " " + kind : "");
  }

  function updateButtons() {
    const allLoaded = FILE_KEYS.every((k) => !!state.files[k]);
    $("#calculate-btn").disabled = !allLoaded;
    $("#download-btn").disabled = !state.rows;
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  function setupFileInput(key) {
    const input = $(`#file-${key}`);
    const card = document.querySelector(`.upload-card[data-key="${key}"]`);
    const status = $(`#status-${key}`);
    input.addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        status.textContent = "טוען...";
        card.classList.remove("loaded", "error");
        const buf = await readFileAsArrayBuffer(file);
        state.files[key] = { name: file.name, buffer: buf };
        status.textContent = file.name;
        card.classList.add("loaded");
      } catch (err) {
        console.error(err);
        status.textContent = "שגיאה בטעינה";
        card.classList.add("error");
        state.files[key] = null;
      }
      state.rows = null;
      updateButtons();
    });
  }

  function calculate() {
    setMessage("מעבד נתונים...", "info");
    try {
      // Parse each file
      const empSheet = Parser.readSheet(state.files.employees.buffer, "גיליון1");
      const employees = Parser.parseEmployees(empSheet);

      const threeSheet = Parser.readSheet(state.files.three_months.buffer, "גיליון1");
      const threeMonths = Parser.parseThreeMonths(threeSheet);

      const aprilSheet = Parser.readSheet(state.files.april.buffer, "Sheet1");
      const april = Parser.parseDaily(aprilSheet);

      const m31Sheet = Parser.readSheet(state.files.march31.buffer, "Sheet1");
      const march31 = Parser.parseDaily(m31Sheet);

      const a9Sheet = Parser.readSheet(state.files.april9.buffer, "Sheet1");
      const april9 = Parser.parseDaily(a9Sheet);

      const rows = Eligibility.calculateAll({
        employees, threeMonths, march31, april9, april,
      });

      state.rows = rows;
      renderResults(rows);
      setMessage(`חושבו ${rows.length} עובדים בהצלחה.`, "success");
    } catch (err) {
      console.error(err);
      setMessage("שגיאה בעיבוד: " + err.message, "error");
      state.rows = null;
    }
    updateButtons();
  }

  function fmtCell(val, fmt) {
    if (val === null || val === undefined || val === "") return "—";
    if (fmt === "date") {
      if (val instanceof Date) {
        const dd = String(val.getDate()).padStart(2, "0");
        const mm = String(val.getMonth() + 1).padStart(2, "0");
        const yy = val.getFullYear();
        return `${dd}/${mm}/${yy}`;
      }
      return String(val);
    }
    if (fmt === "yesno") {
      return val ? '<span class="cell-yes">כן</span>' : '<span class="cell-no">לא</span>';
    }
    if (typeof val === "number") {
      return `<span class="cell-num">${val}</span>`;
    }
    return String(val);
  }

  function renderResults(rows) {
    const cols = Exporter.COLUMNS;
    const thead = $("#results-table thead");
    const tbody = $("#results-table tbody");

    thead.innerHTML = "<tr>" + cols.map((c) => `<th>${c.header}</th>`).join("") + "</tr>";
    tbody.innerHTML = rows.map((r) => {
      return "<tr>" + cols.map((c) => `<td>${fmtCell(r[c.key], c.fmt)}</td>`).join("") + "</tr>";
    }).join("");

    // Summary
    const eligibleA = rows.filter((r) => r.eligibleA).length;
    const eligibleB = rows.filter((r) => r.eligibleB).length;
    const totalDaysUsed = rows.reduce((s, r) => s + (r.daysToUse || 0), 0);
    const totalHours = rows.reduce((s, r) => s + (r.totalHolidayHours || 0), 0);

    $("#summary").innerHTML = `
      <div class="summary-card"><strong>${rows.length}</strong>סה"כ עובדים</div>
      <div class="summary-card"><strong>${eligibleA}</strong>זכאים לפסח א</div>
      <div class="summary-card"><strong>${eligibleB}</strong>זכאים לפסח ב</div>
      <div class="summary-card"><strong>${totalDaysUsed}</strong>ימי חופש לניצול</div>
      <div class="summary-card"><strong>${Math.round(totalHours * 100) / 100}</strong>סה"כ שעות חופש</div>
    `;
    $("#results-section").hidden = false;
  }

  function downloadXlsx() {
    if (!state.rows) return;
    Exporter.exportXlsx(state.rows);
  }

  // Init
  document.addEventListener("DOMContentLoaded", () => {
    FILE_KEYS.forEach(setupFileInput);
    $("#calculate-btn").addEventListener("click", calculate);
    $("#download-btn").addEventListener("click", downloadXlsx);
    updateButtons();
  });
})();
