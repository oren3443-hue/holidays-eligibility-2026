/**
 * main.js - UI orchestration for both Pesach and Yom HaAtzmaut calculators.
 */

(function () {
  // ===== State =====
  const state = {
    holiday: "pesach", // "pesach" | "independence"

    // Pesach slot files
    files: {
      employees: null,
      three_months: null,
      april: null,
      march31: null,
      april9: null,
    },

    // Independence slot files (separate keys to avoid cross-mode contamination)
    indFiles: {
      ind_employees: null,
      ind_three_months: null,
    },
    // Independence multi-upload: array of { name, buffer }
    shiftReports: [],

    rows: null, // calculated results (shape depends on mode)

    filters: {
      pesach: {
        eligibleA: "", eligibleB: "", eligibleAny: "",
        tenure: "", dpw: "", usage: "", search: "",
      },
      independence: {
        mode: "", tenure: "", dpw: "", search: "",
      },
    },
  };

  const PESACH_KEYS = ["employees", "three_months", "april", "march31", "april9"];
  const IND_KEYS = ["ind_employees", "ind_three_months"];
  const PESACH_SUBTITLE = "פסח א · חול המועד · פסח ב — חישוב אוטומטי לכל עובד";
  const IND_SUBTITLE = "יום העצמאות — שעות חג בפועל ותשלום חג לכל עובד";

  function $(sel) { return document.querySelector(sel); }
  function $$(sel) { return Array.from(document.querySelectorAll(sel)); }

  function setMessage(text, kind) {
    const el = $("#message");
    el.textContent = text || "";
    el.className = "message" + (kind ? " " + kind : "");
  }

  // ===== Mode helpers =====
  function isPesach() { return state.holiday === "pesach"; }
  function activeColumns() { return isPesach() ? Exporter.COLUMNS : Exporter.COLUMNS_INDEPENDENCE; }

  function applyHolidayMode() {
    document.querySelectorAll("[data-mode]").forEach((el) => {
      const m = el.getAttribute("data-mode");
      el.hidden = m !== state.holiday;
    });
    document.querySelectorAll(".holiday-tab").forEach((btn) => {
      const active = btn.getAttribute("data-holiday") === state.holiday;
      btn.classList.toggle("active", active);
      btn.setAttribute("aria-selected", active ? "true" : "false");
    });
    $("#subtitle").textContent = isPesach() ? PESACH_SUBTITLE : IND_SUBTITLE;
  }

  function setHoliday(holiday) {
    if (state.holiday === holiday) return;
    state.holiday = holiday;
    state.rows = null;
    setMessage("");
    $("#results-section").hidden = true;
    applyHolidayMode();
    updateButtons();
  }

  function updateButtons() {
    let allLoaded;
    if (isPesach()) {
      allLoaded = PESACH_KEYS.every((k) => !!state.files[k]);
    } else {
      allLoaded = IND_KEYS.every((k) => !!state.indFiles[k]) && state.shiftReports.length > 0;
    }
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

  // ===== Single-file slot wiring (works for both Pesach and Independence) =====
  function getSlotStore(key) {
    if (PESACH_KEYS.includes(key)) return state.files;
    if (IND_KEYS.includes(key)) return state.indFiles;
    return null;
  }

  function assignFileToSlot(key, file, buffer) {
    const card = document.querySelector(`.upload-card[data-key="${key}"]`);
    const status = $(`#status-${key}`);
    const store = getSlotStore(key);
    if (!store) return;
    store[key] = { name: file.name, buffer };
    status.textContent = file.name;
    card.classList.remove("error");
    card.classList.add("loaded");
  }

  function markSlotError(key, msg) {
    const card = document.querySelector(`.upload-card[data-key="${key}"]`);
    const status = $(`#status-${key}`);
    status.textContent = msg || "שגיאה בטעינה";
    card.classList.remove("loaded");
    card.classList.add("error");
    const store = getSlotStore(key);
    if (store) store[key] = null;
  }

  function setupFileInput(key) {
    const input = $(`#file-${key}`);
    if (!input) return;
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

  // ===== Pesach bulk upload =====
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
        if (!result || !PESACH_KEYS.includes(result.type)) {
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

    const missing = PESACH_KEYS.filter((k) => !state.files[k]);
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
    if (!input || !zone) return;
    input.addEventListener("change", (e) => {
      handleBulkFiles(e.target.files);
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

  // ===== Independence: multi-upload for shift reports =====
  async function handleShiftReportFiles(fileList) {
    const detected = $("#shift-reports-detected");
    const files = Array.from(fileList);
    if (!files.length) return;

    const chips = [];
    const existingNames = new Set(state.shiftReports.map((f) => f.name));

    for (const f of files) {
      try {
        const buf = await readFileAsArrayBuffer(f);
        const result = Detector.detect(buf);
        if (!result || result.type !== "shift_report") {
          chips.push({ kind: "error", text: `❌ ${f.name} — לא זוהה כדוח שעות מפורט` });
          continue;
        }
        if (existingNames.has(f.name)) {
          chips.push({ kind: "warn", text: `⚠️ ${f.name} — כבר נטען` });
          continue;
        }
        state.shiftReports.push({ name: f.name, buffer: buf });
        existingNames.add(f.name);
        chips.push({ kind: "ok", text: `✓ ${f.name}` });
      } catch (err) {
        console.error(err);
        chips.push({ kind: "error", text: `❌ ${f.name} — ${err.message}` });
      }
    }

    renderShiftReportChips(chips);
    state.rows = null;
    updateButtons();
  }

  function renderShiftReportChips(newChips) {
    const detected = $("#shift-reports-detected");
    if (!detected) return;
    const loadedChips = state.shiftReports.map((f) => ({
      kind: "ok",
      text: `✓ ${f.name}`,
      removable: true,
      name: f.name,
    }));
    const allChips = newChips
      ? newChips.filter((c) => c.kind !== "ok").concat(loadedChips)
      : loadedChips;
    detected.innerHTML = allChips.map((c, i) => {
      const cls = c.kind === "ok" ? "" : c.kind;
      const removeBtn = c.removable
        ? ` <button type="button" class="chip-remove" data-name="${c.name}" title="הסר">✕</button>`
        : "";
      return `<span class="detected-chip ${cls}">${c.text}${removeBtn}</span>`;
    }).join("");

    detected.querySelectorAll(".chip-remove").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const name = btn.getAttribute("data-name");
        state.shiftReports = state.shiftReports.filter((f) => f.name !== name);
        state.rows = null;
        renderShiftReportChips();
        updateButtons();
      });
    });
  }

  function setupShiftReportsUpload() {
    const input = $("#shift-reports-input");
    const zone = $("#shift-reports-zone");
    if (!input || !zone) return;
    input.addEventListener("change", (e) => {
      handleShiftReportFiles(e.target.files);
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
      if (fl && fl.length) handleShiftReportFiles(fl);
    });
  }

  // ===== Holiday tabs =====
  function setupHolidayTabs() {
    document.querySelectorAll(".holiday-tab").forEach((btn) => {
      btn.addEventListener("click", () => {
        setHoliday(btn.getAttribute("data-holiday"));
      });
    });
  }

  // ===== Calculate =====
  function calculate() {
    setMessage("מעבד נתונים...", "info");
    try {
      if (isPesach()) {
        calculatePesach();
      } else {
        calculateIndependence();
      }
    } catch (err) {
      console.error(err);
      setMessage("שגיאה בעיבוד: " + err.message, "error");
      state.rows = null;
    }
    updateButtons();
  }

  function calculatePesach() {
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
  }

  /** Strip the boilerplate from a shift-report filename, leaving just the branch name.
   *  e.g. "דוח נוכחות אגמים - אשקלון 20_04_26 - 23_04_26" → "אגמים - אשקלון" */
  function branchNameFromFile(filename) {
    let label = filename.replace(/\.(xlsx|xls)$/i, "").trim();
    label = label.replace(/^דוח\s+נוכחות\s+/, "");
    label = label.replace(/\s+\d{1,2}[._\-/]\d{1,2}[._\-/]\d{2,4}.*$/, "");
    return label.trim() || filename;
  }

  function calculateIndependence() {
    const empSheet = Parser.readSheet(state.indFiles.ind_employees.buffer, "גיליון1");
    const employees = Parser.parseEmployees(empSheet);

    const threeSheet = Parser.readSheet(state.indFiles.ind_three_months.buffer, "גיליון1");
    const threeMonths = Parser.parseThreeMonths(threeSheet);

    const perFile = [];
    const fileErrors = [];
    for (const f of state.shiftReports) {
      try {
        const rows = Parser.readSheet(f.buffer);
        const label = branchNameFromFile(f.name);
        perFile.push(Parser.parseShiftReport(rows, label));
      } catch (err) {
        fileErrors.push({ file: f.name, message: err.message });
      }
    }
    if (perFile.length === 0) {
      throw new Error("אף קובץ דוח שעות לא נטען בהצלחה. שגיאות: " + fileErrors.map(e => e.message).join(" | "));
    }
    const shifts = Parser.combineShiftReports(perFile);

    const rows = IndependenceEligibility.calculateAll({
      employees, threeMonths, shifts,
    });

    state.rows = rows;
    renderResults();
    if (fileErrors.length) {
      const errs = fileErrors.map(e => e.message).join("\n");
      setMessage(`חושבו ${rows.length} עובדים בהצלחה. שים לב — קבצים עם שגיאות (לא נכללו): ${fileErrors.length}.\n${errs}`, "warning");
    } else {
      setMessage(`חושבו ${rows.length} עובדים בהצלחה.`, "success");
    }
  }

  // ===== Rendering =====
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
    if (isPesach()) return applyPesachFilters(rows);
    return applyIndependenceFilters(rows);
  }

  function applyPesachFilters(rows) {
    const f = state.filters.pesach;
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

  function applyIndependenceFilters(rows) {
    const f = state.filters.independence;
    const search = (f.search || "").trim().toLowerCase();
    return rows.filter((r) => {
      if (f.mode && r.mode !== f.mode) return false;
      if (f.tenure === "yes" && !r.tenureOk) return false;
      if (f.tenure === "no"  &&  r.tenureOk) return false;
      if (f.dpw === "5" && r.daysPerWeek !== 5) return false;
      if (f.dpw === "6" && r.daysPerWeek !== 6) return false;
      if (search) {
        const hay = `${r.empId} ${r.firstName || ""} ${r.lastName || ""} ${r.branches || ""}`.toLowerCase();
        if (!hay.includes(search)) return false;
      }
      return true;
    });
  }

  function renderResults() {
    if (!state.rows) return;
    const allRows = state.rows;
    const filteredRows = applyFilters(allRows);
    const cols = activeColumns();
    const thead = $("#results-table thead");
    const tbody = $("#results-table tbody");

    thead.innerHTML = "<tr>" + cols.map((c) => `<th>${c.header}</th>`).join("") + "</tr>";
    tbody.innerHTML = filteredRows.map((r) => {
      return "<tr>" + cols.map((c) => `<td>${fmtCell(r[c.key], c)}</td>`).join("") + "</tr>";
    }).join("");

    if (isPesach()) {
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
      $("#filter-count").innerHTML =
        filteredRows.length === allRows.length
          ? `מוצגים <strong>${allRows.length}</strong> עובדים`
          : `מוצגים <strong>${filteredRows.length}</strong> מתוך <strong>${allRows.length}</strong>`;
    } else {
      const worked = allRows.filter((r) => r.mode === "worked").length;
      const eligiblePay = allRows.filter((r) => r.mode === "holiday_pay").length;
      const ineligible = allRows.filter((r) => r.mode === "ineligible").length;
      const sumWorked = allRows.reduce((s, r) => s + (r.extraHolidayHours || 0), 0);
      const sumPay = allRows.reduce((s, r) => s + (r.holidayPayHours || 0), 0);
      $("#summary").innerHTML = `
        <div class="summary-card"><strong>${allRows.length}</strong>סה"כ עובדים</div>
        <div class="summary-card"><strong>${worked}</strong>עבדו בחג</div>
        <div class="summary-card"><strong>${eligiblePay}</strong>זכאים לתשלום חג</div>
        <div class="summary-card"><strong>${ineligible}</strong>לא זכאים</div>
        <div class="summary-card"><strong>${Math.round(sumWorked * 100) / 100}</strong>שעות חג 100%</div>
        <div class="summary-card"><strong>${Math.round(sumPay * 100) / 100}</strong>שעות תשלום חג</div>
      `;
      $("#filter-count-ind").innerHTML =
        filteredRows.length === allRows.length
          ? `מוצגים <strong>${allRows.length}</strong> עובדים`
          : `מוצגים <strong>${filteredRows.length}</strong> מתוך <strong>${allRows.length}</strong>`;
    }

    $("#results-section").hidden = false;
  }

  // ===== Filters wiring =====
  function setupFilters() {
    // Pesach filters
    document.querySelectorAll('#filters select[data-filter]').forEach((sel) => {
      sel.addEventListener("change", (e) => {
        state.filters.pesach[e.target.dataset.filter] = e.target.value;
        renderResults();
      });
    });
    let psearchTimer = null;
    const psearch = $("#filter-search");
    if (psearch) {
      psearch.addEventListener("input", (e) => {
        clearTimeout(psearchTimer);
        psearchTimer = setTimeout(() => {
          state.filters.pesach.search = e.target.value;
          renderResults();
        }, 150);
      });
    }
    const presetBtn = $("#reset-filters");
    if (presetBtn) {
      presetBtn.addEventListener("click", () => {
        Object.keys(state.filters.pesach).forEach((k) => (state.filters.pesach[k] = ""));
        document.querySelectorAll('#filters select').forEach((s) => (s.value = ""));
        if (psearch) psearch.value = "";
        renderResults();
      });
    }

    // Independence filters
    document.querySelectorAll('#filters-ind select[data-filter]').forEach((sel) => {
      sel.addEventListener("change", (e) => {
        state.filters.independence[e.target.dataset.filter] = e.target.value;
        renderResults();
      });
    });
    let isearchTimer = null;
    const isearch = $("#filter-search-ind");
    if (isearch) {
      isearch.addEventListener("input", (e) => {
        clearTimeout(isearchTimer);
        isearchTimer = setTimeout(() => {
          state.filters.independence.search = e.target.value;
          renderResults();
        }, 150);
      });
    }
    const iresetBtn = $("#reset-filters-ind");
    if (iresetBtn) {
      iresetBtn.addEventListener("click", () => {
        Object.keys(state.filters.independence).forEach((k) => (state.filters.independence[k] = ""));
        document.querySelectorAll('#filters-ind select').forEach((s) => (s.value = ""));
        if (isearch) isearch.value = "";
        renderResults();
      });
    }
  }

  // ===== Downloads =====
  function downloadXlsx() {
    if (!state.rows) return;
    const rows = applyFilters(state.rows);
    if (isPesach()) {
      Exporter.exportXlsx(rows);
    } else {
      Exporter.exportXlsxIndependence(rows);
    }
  }

  function downloadImportXlsx() {
    if (!state.rows) return;
    const filtered = applyFilters(state.rows);
    const meta = {
      companyNumber: parseInt($("#cfg-company").value, 10) || 10,
      year: parseInt($("#cfg-year").value, 10) || 2026,
      month: parseInt($("#cfg-month").value, 10) || 4,
    };
    if (isPesach()) {
      const onlyWithData = filtered.filter((r) => {
        return (r.daysToUse > 0) || (r.totalHolidayHours > 0) ||
               (r.holidayDaysCount > 0) || (r.holidayPayHours > 0);
      });
      Exporter.exportImportXlsx(onlyWithData, undefined, meta);
    } else {
      const onlyWithData = filtered.filter((r) => {
        return (r.holidayPayHours > 0) || (r.extraHolidayHours > 0) || (r.holidayDaysCount > 0);
      });
      Exporter.exportImportXlsxIndependence(onlyWithData, undefined, meta);
    }
  }

  // ===== Init =====
  document.addEventListener("DOMContentLoaded", () => {
    PESACH_KEYS.forEach(setupFileInput);
    IND_KEYS.forEach(setupFileInput);
    setupBulkUpload();
    setupShiftReportsUpload();
    setupHolidayTabs();
    $("#calculate-btn").addEventListener("click", calculate);
    $("#download-btn").addEventListener("click", downloadXlsx);
    $("#download-import-btn").addEventListener("click", downloadImportXlsx);
    setupFilters();
    applyHolidayMode();
    updateButtons();
  });
})();
