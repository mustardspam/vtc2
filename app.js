const SHEET_ALIASES = {
  safety: "Safety",
  schedule_adherence_raw: "Schedule_Adherence_Raw",
  scheduleraw: "Schedule_Adherence_Raw",
  schedule: "Schedule_Adherence_Raw",
  inspections: "Inspections",
  rework: "Rework",
  log: "Log",
  fieldlog: "Log",
  warranty: "Warranty",
  config: "Config",
  blacklist: "Black List",
  "black list": "Black List",
  workloadsl: "WorkloadSL",
  workloadaw: "WorkloadAW"
};

const DEFAULT_WEIGHTS = {
  Safety: 0.25,
  Schedule: 0.25,
  Inspections: 0,
  Rework: 0.125,
  Warranty: 0,
  Log: 0.375
};

const WEIGHT_METRICS = Object.keys(DEFAULT_WEIGHTS);
const STORAGE_KEY = "vtc-scorecard-state-v1";

const LOG_POINTS = {
  "Kudos|-": 100,
  "Complaint|Minor": 85,
  "Complaint|Major": 70,
  "Complaint|Critical": 50
};

const state = {
  workbook: null,
  fileName: "",
  tables: emptyTables(),
  vendors: [],
  weights: { ...DEFAULT_WEIGHTS },
  scores: [],
  activity: [],
  lastSaved: ""
};

const els = {};

window.addEventListener("DOMContentLoaded", () => {
  [
    "masterInput",
    "sourceInput",
    "exportBtn",
    "resetBtn",
    "vendorCount",
    "rowCount",
    "pendingCount",
    "lastUpdate",
    "weightsPanel",
    "scoreRows",
    "topBestTrades",
    "topWorstTrades",
    "vendorSearch",
    "categoryFilter",
    "activityLog",
    "feedbackText",
    "parseFeedbackBtn",
    "feedbackFileInput",
    "feedbackForm",
    "fbVendor",
    "fbSubmitted",
    "fbCategory",
    "fbSeverity",
    "fbCommunity",
    "fbNotes",
    "vendorList"
  ].forEach((id) => (els[id] = document.getElementById(id)));

  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => activateView(tab.dataset.view));
  });

  els.masterInput.addEventListener("change", (event) => loadMaster(event.target.files[0]));
  els.sourceInput.addEventListener("change", (event) => ingestFiles([...event.target.files]));
  els.feedbackFileInput.addEventListener("change", (event) => readFeedbackFiles([...event.target.files]));
  els.exportBtn.addEventListener("click", exportWorkbook);
  els.resetBtn.addEventListener("click", resetSession);
  els.vendorSearch.addEventListener("input", renderScores);
  els.categoryFilter.addEventListener("change", renderScores);
  els.parseFeedbackBtn.addEventListener("click", () => parseFeedback(els.feedbackText.value));
  els.feedbackForm.addEventListener("submit", addFeedbackToLog);

  wireDropZone(document.body);
  if (restoreState()) {
    logActivity("Restored saved scorecard data from this browser.");
  } else {
    logActivity("Ready. Load the master VTC scorecard to begin.");
  }
  render();
  if (window.lucide) window.lucide.createIcons();
});

function emptyTables() {
  return {
    Config: [],
    Safety: [],
    Schedule_Adherence_Raw: [],
    Inspections: [],
    Rework: [],
    Log: [],
    Warranty: [],
    WorkloadSL: [],
    WorkloadAW: [],
    "Black List": []
  };
}

async function loadMaster(file) {
  if (!file) return;
  state.workbook = await readWorkbook(file);
  state.fileName = file.name;
  state.tables = workbookToTables(state.workbook);
  state.vendors = readVendors(state.tables.Config);
  state.weights = readWeights(state.tables.Config);
  state.scores = calculateScores();
  logActivity(`Loaded master scorecard: ${file.name}`);
  saveState();
  render();
}

async function ingestFiles(files) {
  if (!files.length) return;
  if (!hasScorecardData()) {
    logActivity("Load the master scorecard before dropping update files.");
    return;
  }

  for (const file of files) {
    const wb = await readWorkbook(file);
    const imported = workbookToTables(wb);
    const matches = Object.entries(imported).filter(([, rows]) => rows.length);

    if (!matches.length) {
      logActivity(`No recognized scorecard tabs found in ${file.name}.`);
      continue;
    }

    for (const [sheetName, rows] of matches) {
      if (sheetName === "Config") {
        mergeConfig(rows);
      } else {
        state.tables[sheetName] = mergeRows(state.tables[sheetName], rows);
      }
      logActivity(`Ingested ${rows.length} row(s) into ${sheetName} from ${file.name}.`);
    }
  }

  state.vendors = readVendors(state.tables.Config);
  state.weights = readWeights(state.tables.Config);
  state.scores = calculateScores();
  saveState();
  render();
}

function workbookToTables(wb) {
  const tables = emptyTables();
  for (const sheetName of wb.SheetNames) {
    const normalized = normalizeSheetName(sheetName);
    if (!normalized) continue;
    const sheet = wb.Sheets[sheetName];
    tables[normalized] = normalized === "Config"
      ? readConfigSheet(sheet)
      : XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(cleanRow);
  }
  return tables;
}

function readConfigSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const configRows = [];
  const vendorHeader = findRowIndex(rows, ["Vendor_ID", "Vendor_Name", "Category"]);
  const metricHeader = findRowIndex(rows, ["Metric", "Weight"]);

  if (vendorHeader >= 0) {
    for (let index = vendorHeader + 1; index < rows.length; index += 1) {
      const row = rows[index];
      if (!row[0] || cleanText(row[0]).toLowerCase().includes("metric weights")) break;
      if (row[1]) {
        configRows.push({
          Vendor_ID: cleanText(row[0]),
          Vendor_Name: cleanText(row[1]),
          Category: cleanText(row[2])
        });
      }
    }
  }

  if (metricHeader >= 0) {
    for (let index = metricHeader + 1; index < rows.length; index += 1) {
      const row = rows[index];
      const metric = cleanText(row[0]);
      const weight = Number(row[1]);
      if (!metric || !Number.isFinite(weight)) {
        if (metric.toLowerCase() === "severity") break;
        continue;
      }
      configRows.push({ Metric: metric, Weight: weight });
    }
  }

  return configRows;
}

function findRowIndex(rows, labels) {
  return rows.findIndex((row) => labels.every((label, index) => cleanText(row[index]) === label));
}

function normalizeSheetName(name) {
  const compact = String(name).trim().toLowerCase().replace(/[\s-]+/g, "_");
  return SHEET_ALIASES[compact] || SHEET_ALIASES[compact.replaceAll("_", "")] || null;
}

function cleanRow(row) {
  const next = {};
  Object.entries(row).forEach(([key, value]) => {
    const cleanKey = String(key).trim();
    next[cleanKey] = typeof value === "string" ? cleanText(value) : value;
  });
  return next;
}

function cleanText(value) {
  return String(value).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function readVendors(configRows) {
  return configRows
    .filter((row) => row.Vendor_ID && row.Vendor_Name)
    .map((row) => ({
      id: cleanText(row.Vendor_ID),
      name: cleanText(row.Vendor_Name),
      category: cleanText(row.Category || "")
    }))
    .filter((vendor) => vendor.name);
}

function readWeights(configRows) {
  const weights = { ...DEFAULT_WEIGHTS };
  for (const row of configRows) {
    const metric = cleanText(row.Metric || "");
    const weight = Number(row.Weight ?? "");
    if (WEIGHT_METRICS.includes(metric) && Number.isFinite(weight)) weights[metric] = weight;
  }
  return normalizeWeights(weights);
}

function mergeConfig(rows) {
  const incomingVendors = readVendors(rows);
  const byId = new Map(state.vendors.map((vendor) => [vendor.id, vendor]));
  incomingVendors.forEach((vendor) => byId.set(vendor.id, vendor));
  state.vendors = [...byId.values()].sort((a, b) => a.name.localeCompare(b.name));
  state.tables.Config = vendorsToConfigRows(state.vendors, state.weights);
}

function mergeRows(existingRows, incomingRows) {
  const seen = new Set(existingRows.map(rowKey));
  const merged = [...existingRows];
  for (const row of incomingRows) {
    const key = rowKey(row);
    if (!seen.has(key)) {
      merged.push(row);
      seen.add(key);
    }
  }
  return merged;
}

function rowKey(row) {
  return JSON.stringify(Object.keys(row).sort().map((key) => [key, row[key]]));
}

function calculateScores() {
  const byVendorName = new Map(state.vendors.map((vendor) => [cleanText(vendor.name).toLowerCase(), vendor]));
  const scheduleById = groupBy(state.tables.Schedule_Adherence_Raw, (row) => cleanText(row.Vendor_ID));
  const safetyByVendor = groupBy(state.tables.Safety, (row) => cleanText(row.Vendor_Name).toLowerCase());
  const inspectionsByVendor = groupBy(state.tables.Inspections, (row) => cleanText(row.Vendor_Name).toLowerCase());
  const reworkByVendor = groupBy(state.tables.Rework, (row) => cleanText(row.Vendor_Name).toLowerCase());
  const warrantyByVendor = groupBy(state.tables.Warranty, (row) => cleanText(row.Vendor_Name).toLowerCase());
  const logByVendor = groupBy(state.tables.Log, (row) => cleanText(row["Vendor Name Clean"] || row["Vendor Name"]).toLowerCase());
  const workload = readWorkload();

  return state.vendors.map((vendor) => {
    const key = cleanText(vendor.name).toLowerCase();
    const safetyRows = safetyByVendor.get(key) || [];
    const scheduleRows = scheduleById.get(cleanText(vendor.id)) || [];
    const inspectionRows = inspectionsByVendor.get(key) || [];
    const reworkRows = reworkByVendor.get(key) || [];
    const warrantyRows = warrantyByVendor.get(key) || [];
    const logRows = logByVendor.get(key) || [];

    const safety = safetyRows.length
      ? Math.max(0, 100 - sum(safetyRows, (row) => Number(row.Severity_Score || 0)) * 10)
      : 100;
    const schedule = scheduleRows.length
      ? average(scheduleRows, (row) => Number(row.Adherence_Pct || 0)) * 100
      : 100;
    const inspections = inspectionRows.length
      ? average(inspectionRows, (row) => Number(row.Score_Pct || 0))
      : null;
    const rework = reworkRows.length
      ? Math.max(0, 100 - sum(reworkRows, (row) => Number(row.PenaltyPoints || 0)) * 5)
      : null;
    const warranty = warrantyRows.length
      ? average(warrantyRows, (row) => Number(row.Warranty_Score || 0))
      : null;
    const log = logRows.length ? average(logRows, (row) => Number(row.Points || 0)) : null;

    const components = {
      Safety: safety,
      Schedule: schedule,
      Inspections: inspections,
      Rework: rework,
      Warranty: warranty,
      Log: log
    };
    const overall = weightedScore(components);
    return {
      ...vendor,
      safety,
      schedule,
      inspections,
      rework,
      warranty,
      log,
      overall,
      workload: workload.get(key) ?? null,
      known: byVendorName.has(key)
    };
  }).sort((a, b) => (b.overall ?? -1) - (a.overall ?? -1));
}

function weightedScore(components) {
  let numerator = 0;
  let denominator = 0;
  for (const [metric, value] of Object.entries(components)) {
    const weight = Number(state.weights[metric] || 0);
    if (value !== null && value !== "" && Number.isFinite(value) && weight > 0) {
      numerator += value * weight;
      denominator += weight;
    }
  }
  return denominator ? numerator / denominator : null;
}

function readWorkload() {
  const map = new Map();
  for (const sheetName of ["WorkloadSL", "WorkloadAW"]) {
    for (const row of state.tables[sheetName]) {
      for (const [key, value] of Object.entries(row)) {
        if (key === "Cost Code" || !value) continue;
        const vendor = cleanText(value).toLowerCase();
        map.set(vendor, (map.get(vendor) || 0) + 1);
      }
    }
  }
  const max = Math.max(1, ...map.values());
  for (const [key, value] of map.entries()) map.set(key, value / max);
  return map;
}

function groupBy(rows, getKey) {
  const grouped = new Map();
  for (const row of rows) {
    const key = getKey(row);
    if (!key) continue;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(row);
  }
  return grouped;
}

function sum(rows, getValue) {
  return rows.reduce((total, row) => total + getValue(row), 0);
}

function average(rows, getValue) {
  const values = rows.map(getValue).filter((value) => Number.isFinite(value));
  return values.length ? values.reduce((total, value) => total + value, 0) / values.length : null;
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

function parseFeedback(text) {
  const cleaned = cleanText(text);
  if (!cleaned) return;
  const vendor = bestVendorMatch(cleaned);
  const submitted = matchValue(text, /(from|submitted by|builder)[:\s]+([^\n<]+)/i);
  const community = matchValue(text, /(community|neighborhood|subdivision)[:\s]+([^\n]+)/i);
  const severity = /critical|urgent|unsafe|failed|no show/i.test(text)
    ? "Critical"
    : /major|repeat|delay|reinspection|backcharge/i.test(text)
      ? "Major"
      : /kudos|great job|thank you|appreciate/i.test(text)
        ? "-"
        : "Minor";
  const category = severity === "-" ? "Kudos" : "Complaint";

  els.fbVendor.value = vendor?.name || "";
  els.fbSubmitted.value = submitted || "";
  els.fbCategory.value = category;
  els.fbSeverity.value = severity;
  els.fbCommunity.value = community || "";
  els.fbNotes.value = cleaned.slice(0, 700);
  logActivity("Parsed field feedback. Review it, then add it to the log.");
}

function bestVendorMatch(text) {
  const lower = text.toLowerCase();
  return state.vendors
    .map((vendor) => ({ vendor, index: lower.indexOf(vendor.name.toLowerCase()) }))
    .filter((match) => match.index >= 0)
    .sort((a, b) => a.index - b.index || b.vendor.name.length - a.vendor.name.length)[0]?.vendor;
}

function matchValue(text, regex) {
  const match = text.match(regex);
  return match ? cleanText(match[2]) : "";
}

async function readFeedbackFiles(files) {
  const chunks = [];
  for (const file of files) chunks.push(await file.text());
  els.feedbackText.value = [els.feedbackText.value, ...chunks].filter(Boolean).join("\n\n");
  parseFeedback(els.feedbackText.value);
}

function addFeedbackToLog(event) {
  event.preventDefault();
  if (!hasScorecardData()) {
    logActivity("Load the master scorecard before adding field feedback.");
    return;
  }
  const vendorName = cleanText(els.fbVendor.value);
  const vendor = state.vendors.find((item) => item.name.toLowerCase() === vendorName.toLowerCase());
  const category = els.fbCategory.value;
  const severity = els.fbSeverity.value;
  const points = LOG_POINTS[`${category}|${severity}`] ?? 70;
  state.tables.Log.push({
    Date: excelDate(new Date()),
    "Vendor Name": vendorName,
    Submitted: cleanText(els.fbSubmitted.value),
    Category: category,
    Severity: severity,
    Community: cleanText(els.fbCommunity.value),
    Notes: cleanText(els.fbNotes.value),
    Points: points,
    "Vendor Name Clean": vendor?.name || vendorName
  });
  state.scores = calculateScores();
  els.feedbackForm.reset();
  els.pendingCount.textContent = "0";
  logActivity(`Added field log entry for ${vendorName || "unknown vendor"} (${category}, ${severity}).`);
  saveState();
  render();
}

function exportWorkbook() {
  if (!hasScorecardData()) return;
  const wb = XLSX.utils.book_new();
  const tables = {
    ...state.tables,
    Scores: scoresToRows()
  };
  for (const [sheetName, rows] of Object.entries(tables)) {
    const sheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, sheet, sheetName.slice(0, 31));
  }
  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `VTC Scorecard Updated ${stamp}.xlsx`);
  logActivity("Exported updated scorecard workbook.");
}

function scoresToRows() {
  return state.scores.map((score, index, all) => ({
    Vendor_Name: score.name,
    Vendor_ID: score.id,
    Category: score.category,
    Safety_Score: round(score.safety),
    Schedule_Score: round(score.schedule),
    Inspection_Score: round(score.inspections),
    Rework_Score: round(score.rework),
    Warranty_Score: round(score.warranty),
    Log_Score: round(score.log),
    Overall_Score: round(score.overall),
    Overall_Rank: index + 1,
    Category_Rank: all.filter((row) => row.category === score.category && row.overall > score.overall).length + 1,
    Workload_Percent: round(score.workload)
  }));
}

function vendorsToConfigRows(vendors, weights) {
  const rows = vendors.map((vendor) => ({
    Vendor_ID: vendor.id,
    Vendor_Name: vendor.name,
    Category: vendor.category
  }));
  Object.entries(weights).forEach(([metric, weight]) => rows.push({ Metric: metric, Weight: weight }));
  return rows;
}

function excelDate(date) {
  return Math.floor((date - new Date(Date.UTC(1899, 11, 30))) / 86400000);
}

function round(value) {
  return value === null || value === "" || !Number.isFinite(value) ? "" : Math.round(value * 100) / 100;
}

function render() {
  els.vendorCount.textContent = state.vendors.length.toLocaleString();
  els.rowCount.textContent = Object.values(state.tables).flat().length.toLocaleString();
  els.lastUpdate.textContent = state.lastSaved ? new Date(state.lastSaved).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) : "None";
  els.exportBtn.disabled = !hasScorecardData();
  renderWeights();
  renderCategoryFilter();
  renderVendorList();
  renderTradeRankings();
  renderScores();
  renderActivity();
  if (window.lucide) window.lucide.createIcons();
}

function renderWeights() {
  const total = WEIGHT_METRICS.reduce((value, metric) => value + Number(state.weights[metric] || 0), 0);
  els.weightsPanel.innerHTML = `
    <div class="weights-head">
      <h2>Weights</h2>
      <strong>${formatPercent(total)}</strong>
    </div>
    ${WEIGHT_METRICS
    .map((metric) => `
      <label class="weight-row">
        <span>${metric}</span>
        <output id="weightValue${metric}">${formatPercent(state.weights[metric])}</output>
        <input
          type="range"
          min="0"
          max="100"
          step="0.5"
          value="${round((state.weights[metric] || 0) * 100)}"
          data-weight-metric="${metric}"
          aria-label="${metric} weight"
        />
      </label>
    `)
    .join("")}
  `;
  els.weightsPanel.querySelectorAll("input[type='range']").forEach((slider) => {
    slider.addEventListener("input", () => updateWeight(slider.dataset.weightMetric, Number(slider.value) / 100));
  });
}

function updateWeight(metric, nextWeight) {
  const current = { ...state.weights };
  const clamped = Math.max(0, Math.min(1, nextWeight));
  const otherMetrics = WEIGHT_METRICS.filter((item) => item !== metric);
  const remaining = 1 - clamped;
  const currentOtherTotal = otherMetrics.reduce((total, item) => total + Number(current[item] || 0), 0);

  current[metric] = clamped;
  if (remaining === 0) {
    otherMetrics.forEach((item) => {
      current[item] = 0;
    });
  } else if (currentOtherTotal > 0) {
    otherMetrics.forEach((item) => {
      current[item] = (Number(current[item] || 0) / currentOtherTotal) * remaining;
    });
  } else {
    otherMetrics.forEach((item) => {
      current[item] = remaining / otherMetrics.length;
    });
  }

  state.weights = normalizeWeights(current);
  state.scores = calculateScores();
  saveState();
  renderWeights();
  renderTradeRankings();
  renderScores();
}

function normalizeWeights(weights) {
  const normalized = {};
  const total = WEIGHT_METRICS.reduce((value, metric) => value + Number(weights[metric] || 0), 0);
  if (total <= 0) return { ...DEFAULT_WEIGHTS };
  WEIGHT_METRICS.forEach((metric) => {
    normalized[metric] = Number(weights[metric] || 0) / total;
  });
  return normalized;
}

function formatPercent(value) {
  const percent = round((Number(value) || 0) * 100);
  return `${percent}%`;
}

function renderCategoryFilter() {
  const current = els.categoryFilter.value;
  const categories = [...new Set(state.vendors.map((vendor) => vendor.category).filter(Boolean))].sort();
  els.categoryFilter.innerHTML = '<option value="">All categories</option>' + categories.map((category) => `<option>${escapeHtml(category)}</option>`).join("");
  els.categoryFilter.value = categories.includes(current) ? current : "";
}

function renderVendorList() {
  els.vendorList.innerHTML = state.vendors.map((vendor) => `<option value="${escapeHtml(vendor.name)}"></option>`).join("");
}

function renderScores() {
  const query = cleanText(els.vendorSearch.value).toLowerCase();
  const category = els.categoryFilter.value;
  const rows = state.scores
    .filter((score) => !category || score.category === category)
    .filter((score) => !query || `${score.name} ${score.category} ${score.id}`.toLowerCase().includes(query))
    .slice(0, 250);

  els.scoreRows.innerHTML = rows.map((score) => `
    <tr>
      <td>${escapeHtml(score.name)}</td>
      <td>${escapeHtml(score.category)}</td>
      <td class="number">${scorePill(score.overall)}</td>
      <td class="number">${formatScore(score.safety)}</td>
      <td class="number">${formatScore(score.schedule)}</td>
      <td class="number">${formatScore(score.rework)}</td>
      <td class="number">${formatScore(score.log)}</td>
      <td class="number">${score.workload === null ? "" : `${Math.round(score.workload * 100)}%`}</td>
    </tr>
  `).join("") || `<tr><td colspan="8">Load the master scorecard to calculate vendor scores.</td></tr>`;
}

function renderTradeRankings() {
  const trades = tradeRankings();
  const best = trades.slice(0, 5);
  const worst = [...trades].reverse().slice(0, 5);
  els.topBestTrades.innerHTML = renderTradeList(best, "No trade scores loaded yet.");
  els.topWorstTrades.innerHTML = renderTradeList(worst, "No trade scores loaded yet.");
}

function tradeRankings() {
  const grouped = groupBy(
    state.scores.filter((score) => score.category && Number.isFinite(score.overall)),
    (score) => score.category
  );
  return [...grouped.entries()]
    .map(([category, scores]) => ({
      category,
      count: scores.length,
      score: average(scores, (item) => item.overall)
    }))
    .filter((trade) => Number.isFinite(trade.score))
    .sort((a, b) => b.score - a.score || b.count - a.count || a.category.localeCompare(b.category));
}

function renderTradeList(trades, emptyText) {
  if (!trades.length) return `<li class="empty-trade">${emptyText}</li>`;
  return trades.map((trade) => `
    <li>
      <span>${escapeHtml(trade.category)}</span>
      <strong>${formatScore(trade.score)}</strong>
      <small>${trade.count} vendor${trade.count === 1 ? "" : "s"}</small>
    </li>
  `).join("");
}

function scorePill(value) {
  const score = round(value);
  if (score === "") return "";
  const tone = score >= 90 ? "good" : score >= 75 ? "warn" : "bad";
  return `<span class="score-pill ${tone}">${score}</span>`;
}

function formatScore(value) {
  const score = round(value);
  return score === "" ? "" : score;
}

function renderActivity() {
  els.activityLog.innerHTML = state.activity.map((item) => `<li>${escapeHtml(item)}</li>`).join("");
}

function logActivity(message) {
  state.activity.unshift(`${new Date().toLocaleString()}: ${message}`);
  state.activity = state.activity.slice(0, 80);
  renderActivity();
}

function activateView(viewName) {
  document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === viewName));
  document.querySelectorAll(".view").forEach((view) => view.classList.remove("active"));
  document.getElementById(`${viewName}View`).classList.add("active");
}

function resetSession() {
  state.workbook = null;
  state.fileName = "";
  state.tables = emptyTables();
  state.vendors = [];
  state.weights = { ...DEFAULT_WEIGHTS };
  state.scores = [];
  state.activity = [];
  state.lastSaved = "";
  localStorage.removeItem(STORAGE_KEY);
  logActivity("Session reset.");
  render();
}

function hasScorecardData() {
  return state.vendors.length > 0 || Object.values(state.tables).some((rows) => rows.length > 0);
}

function saveState() {
  try {
    state.lastSaved = new Date().toISOString();
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      fileName: state.fileName,
      tables: state.tables,
      vendors: state.vendors,
      weights: state.weights,
      activity: state.activity,
      lastSaved: state.lastSaved
    }));
  } catch (error) {
    console.warn("Unable to save scorecard state", error);
    logActivity("Browser storage is full, so this update was not saved for refresh.");
  }
}

function restoreState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return false;
    const parsed = JSON.parse(saved);
    state.fileName = parsed.fileName || "";
    state.tables = { ...emptyTables(), ...(parsed.tables || {}) };
    state.vendors = parsed.vendors?.length ? parsed.vendors : readVendors(state.tables.Config);
    state.weights = normalizeWeights({ ...DEFAULT_WEIGHTS, ...(parsed.weights || {}) });
    state.activity = Array.isArray(parsed.activity) ? parsed.activity.slice(0, 80) : [];
    state.lastSaved = parsed.lastSaved || "";
    state.scores = calculateScores();
    return hasScorecardData();
  } catch (error) {
    console.warn("Unable to restore scorecard state", error);
    localStorage.removeItem(STORAGE_KEY);
    return false;
  }
}

function wireDropZone(target) {
  target.addEventListener("dragover", (event) => {
    event.preventDefault();
  });
  target.addEventListener("drop", (event) => {
    event.preventDefault();
    const files = [...event.dataTransfer.files];
    const feedbackFiles = files.filter((file) => /\.(eml|txt)$/i.test(file.name));
    const workbookFiles = files.filter((file) => /\.(xlsx|xlsm|xlsb|xls|csv)$/i.test(file.name));
    if (!state.workbook && workbookFiles.length) {
      loadMaster(workbookFiles.shift());
    }
    if (workbookFiles.length) ingestFiles(workbookFiles);
    if (feedbackFiles.length) readFeedbackFiles(feedbackFiles);
  });
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  })[char]);
}
