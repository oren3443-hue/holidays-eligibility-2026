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
    filters: {
      eligibleA: "",
      eligibleB: "",
      eligibleAny: "",
      tenure: "",
      dpw: "",
      usage: "",
      search: "",
    },
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
    $("#download-import-btn").disabled = !state.rows;
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  function assignFileToSlot(key, file, buffer) {
    const card = document.querySelector(`.upload-card[data-key="${key}"]`);
    const status = $(`#status-${key}`);
    state.files[key] = { name: file.name, buffer };
    status.textContent = file.name;
    card.classList.remove("error");
    card.classList.add("loaded");
  }

  function clearSlot(key) {
    const card = document.querySelector(`.upload-card[data-key="${key}"]`);
    const status = $(`#status-${key}`);
    state.files[key] = null;
    status.textContent = "לא נבחר קובץ";
    card.classList.remove("loaded", "error");
  }

  function markSlotError(key, msg) {
    const card = document.querySelector(`.upload-card[data-key="${key}"]`);
    const status = $(`#status-${key}`);
    status.textContent = msg || "שגיאה בטעינה";
    card.classList.remove("loaded");
    card.classList.add("error");
    state.files[key] = null;
  }

  function setupFileInput(key) {
    const input = $(`#file-${key}`);
    input.addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        const buf = await readFileAsArrayBuffer(file);
        assignFileToSlot(key, file, buf);
      } catch (err) {
        console.error(err);
        markSlotError(key);
      }
      state.rows = null;
      updateButtons();
    });
  }

  /** Bulk upload: detect each file's type and route to the right slot. */
  async function handleBulkFiles(fileList) {
    const detected = $("#bulk-detected");
    detected.innerHTML = "";

    const files = Array.from(fileList);
    if (!files.length) return;

    const matchedKeys = new Set();
    const chips = [];

    for (const f of files) {
      try {
        const buf = await readFileAsArrayBuffer(f);
        const result = Detector.detect(buf);
        if (!result || !["employees", "three_months", "april", "march31", "april9"].includes(result.type)) {
          chips.push({ kind: "error", text: `❌ ${f.name} — לא זוהה` });
          continue;
        }
        if (matchedKeys.has(result.type)) {
          chips.push({
            kind: "warn",
            text: `⚠️ ${f.name} — שני קבצים זוהו כ"${Detector.LABEL[result.type]}"`,
          });
          continue;
        }
        assignFileToSlot(result.type, f, buf);
        matchedKeys.add(result.type);
        chips.push({
          kind: "ok",
          text: `✓ ${Detector.LABEL[result.type]}: ${f.name}`,
        });
      } catch (err) {
        console.error(err);
        chips.push({ kind: "error", text: `❌ ${f.name} — ${err.message}` });
      }
    }

    detected.innerHTML = chips
      .map((c) => `<span class="detected-chip ${c.kind === "ok" ? "" : c.kind}">${c.text}</span>`)
      .join("");

    // Note any slots that still missing
    const missing = FILE_KEYS.filter((k) => !state.files[k]);
    if (missing.length) {
      detected.innerHTML +=
        `<span class="detected-chip warn">חסרים: ${missing.map((k) => Detector.LABEL[k]).join(", ")}</span>`;
    }

    state.rows = null;
    updateButtons();
  }

  function setupBulkUpload() {
    const input = $("#bulk-input");
    const zone = $("#bulk-zone");
    input.addEventListener("change", (e) => {
      handleBulkFiles(e.target.files);
      // Reset so the same files can be re-selected later if needed
      e.target.value = "";
    });
    ["dragenter", "dragover"].forEach((ev) => {
      zone.addEventListener(ev, (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach((ev) => {
      zone.addEventListener(ev, (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.remove("dragover");
      });
    });
    zone.addEventListener("drop", (e) => {
      const fl = e.dataTransfer && e.dataTransfer.files;
      if (fl && fl.length) handleBulkFiles(fl);
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
      renderResults();
      setMessage(`חושבו ${rows.length} עובדים בהצלחה.`, "success");
    } catch (err) {
      console.error(err);
      setMessage("שגיאה בעיבוד: " + err.message, "error");
      state.rows = null;
    }
    updateButtons();
  }

  function fmtCell(val, col) {
    if (val === null || val === undefined || val === "") return "—";
    if (col.type === "date" || col.fmt === "date") {
      if (val instanceof Date) {
        const dd = String(val.getDate()).padStart(2, "0");
        const mm = String(val.getMonth() + 1).padStart(2, "0");
        const yy = val.getFullYear();
        return `${dd}/${mm}/${yy}`;
      }
      return String(val);
    }
    if (col.fmt === "yesno") {
      return val ? '<span class="cell-yes">כן</span>' : '<span class="cell-no">לא</span>';
    }
    if (typeof val === "number") {
      return `<span class="cell-num">${val}</span>`;
    }
    return String(val);
  }

  function applyFilters(rows) {
    const f = state.filters;
    const search = (f.search || "").trim().toLowerCase();
    return rows.filter((r) => {
      if (f.eligibleA === "yes" && !r.eligibleA) return false;
      if (f.eligibleA === "no"  &&  r.eligibleA) return false;
      if (f.eligibleB === "yes" && !r.eligibleB) return false;
      if (f.eligibleB === "no"  &&  r.eligibleB) return false;
      if (f.eligibleAny === "any"  && !(r.eligibleA || r.eligibleB)) return false;
      if (f.eligibleAny === "both" && !(r.eligibleA && r.eligibleB)) return false;
      if (f.eligibleAny === "none" &&  (r.eligibleA || r.eligibleB)) return false;
      if (f.tenure === "yes" && !r.tenureOk) return false;
      if (f.tenure === "no"  &&  r.tenureOk) return false;
      if (f.dpw === "5" && r.daysPerWeek !== 5) return false;
      if (f.dpw === "6" && r.daysPerWeek !== 6) return false;
      if (f.usage === "any"  && !(r.daysToUse > 0)) return false;
      if (f.usage === "none" &&  (r.daysToUse > 0)) return false;
      if (search) {
        const hay = `${r.empId} ${r.firstName || ""} ${r.lastName || ""}`.toLowerCase();
        if (!hay.includes(search)) return false;
      }
      return true;
    });
  }

  function renderResults() {
    if (!state.rows) return;
    const allRows = state.rows;
    const filteredRows = applyFilters(allRows);
    const cols = Exporter.COLUMNS;
    const thead = $("#results-table thead");
    const tbody = $("#results-table tbody");

    thead.innerHTML = "<tr>" + cols.map((c) => `<th>${c.header}</th>`).join("") + "</tr>";
    tbody.innerHTML = filteredRows.map((r) => {
      return "<tr>" + cols.map((c) => `<td>${fmtCell(r[c.key], c)}</td>`).join("") + "</tr>";
    }).join("");

    // Summary always over the FULL set, not the filtered view
    const eligibleA = allRows.filter((r) => r.eligibleA).length;
    const eligibleB = allRows.filter((r) => r.eligibleB).length;
    const totalDaysUsed = allRows.reduce((s, r) => s + (r.daysToUse || 0), 0);
    const totalHours = allRows.reduce((s, r) => s + (r.totalHolidayHours || 0), 0);

    $("#summary").innerHTML = `
      <div class="summary-card"><strong>${allRows.length}</strong>סה"כ עובדים</div>
      <div class="summary-card"><strong>${eligibleA}</strong>זכאים לפסח א</div>
      <div class="summary-card"><strong>${eligibleB}</strong>זכאים לפסח ב</div>
      <div class="summary-card"><strong>${totalDaysUsed}</strong>ימי חופש לניצול</div>
      <div class="summary-card"><strong>${Math.round(totalHours * 100) / 100}</strong>סה"כ שעות חופש</div>
    `;

    // Filter result counter
    $("#filter-count").innerHTML =
      filteredRows.length === allRows.length
        ? `מוצגים <strong>${allRows.length}</strong> עובדים`
        : `מוצגים <strong>${filteredRows.length}</strong> מתוך <strong>${allRows.length}</strong>`;

    $("#results-section").hidden = false;
  }

  function setupFilters() {
    document.querySelectorAll("#filters select[data-filter]").forEach((sel) => {
      sel.addEventListener("change", (e) => {
        state.filters[e.target.dataset.filter] = e.target.value;
        renderResults();
      });
    });
    const search = $("#filter-search");
    let searchTimer = null;
    search.addEventListener("input", (e) => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => {
        state.filters.search = e.target.value;
        renderResults();
      }, 150);
    });
    $("#reset-filters").addEventListener("click", () => {
      Object.keys(state.filters).forEach((k) => (state.filters[k] = ""));
      document.querySelectorAll("#filters select").forEach((s) => (s.value = ""));
      $("#filter-search").value = "";
      renderResults();
    });
  }

  function downloadXlsx() {
    if (!state.rows) return;
    // Export reflects what the user sees: filtered rows if filters are active.
    const rows = applyFilters(state.rows);
    Exporter.exportXlsx(rows);
  }

  function downloadImportXlsx() {
    if (!state.rows) return;
    // Import file: respect filters, then drop rows where ALL the four importable
    // payroll fields are zero/empty — those employees have nothing to import.
    const filtered = applyFilters(state.rows).filter((r) => {
      return (r.daysToUse > 0) || (r.totalHolidayHours > 0) ||
             (r.holidayDaysCount > 0) || (r.holidayPayHours > 0);
    });
    Exporter.exportImportXlsx(filtered);
  }

  // Init
  document.addEventListener("DOMContentLoaded", () => {
    FILE_KEYS.forEach(setupFileInput);
    setupBulkUpload();
    $("#calculate-btn").addEventListener("click", calculate);
    $("#download-btn").addEventListener("click", downloadXlsx);
    $("#download-import-btn").addEventListener("click", downloadImportXlsx);
    setupFilters();
    updateButtons();
  });
})();
