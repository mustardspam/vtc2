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
  Rework: 0.125,
  Log: 0.375
};

const WEIGHT_METRICS = Object.keys(DEFAULT_WEIGHTS);
const STORAGE_KEY = "vtc-scorecard-state-v1";
const CLOUD_CONFIG_KEY = "vtc-scorecard-cloud-config-v1";
const APP_VERSION = "REST sync build 2026-05-11.11";

const SCHEDULE_COLUMN_ALIASES = {
  month: "Month",
  period: "Month",
  score_month: "Month",
  schedule_month: "Month",
  date: "Month",
  vendor_adj: "Vendor_Adj",
  vendor: "Vendor_Name",
  vendor_name: "Vendor_Name",
  vendorname: "Vendor_Name",
  trade_partner: "Vendor_Name",
  supplier: "Vendor_Name",
  vendor_id: "Vendor_ID",
  vendorid: "Vendor_ID",
  vendor_number: "Vendor_ID",
  vendor_no: "Vendor_ID",
  monthly_tasks: "Monthly_Tasks",
  monthly_task_count: "Monthly_Tasks",
  task_count: "Monthly_Tasks",
  tasks: "Monthly_Tasks",
  jobs: "Monthly_Tasks",
  job_count: "Monthly_Tasks",
  starts: "Monthly_Tasks",
  workload: "Monthly_Tasks",
  no_show_count: "No_Show_Count",
  no_shows: "No_Show_Count",
  noshows: "No_Show_Count",
  no_show: "No_Show_Count",
  noshow: "No_Show_Count",
  missed_appointments: "No_Show_Count",
  missed_jobs: "No_Show_Count",
  adherence_pct: "Adherence_Pct",
  adherence_percent: "Adherence_Pct",
  adherence: "Adherence_Pct"
};

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
  lastSaved: "",
  cloud: {
    url: "",
    anonKey: "",
    workspaceId: "vtc-main",
    connected: false,
    status: "Not connected"
  }
};

const els = {};
let cloudSaveTimer = null;

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
    "cloudStatus",
    "cloudUrl",
    "cloudAnonKey",
    "cloudWorkspace",
    "cloudSaveConfigBtn",
    "cloudLoadBtn",
    "cloudSaveBtn",
    "cloudShareBtn",
    "appVersion",
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
  els.cloudSaveConfigBtn.addEventListener("click", connectCloudFromForm);
  els.cloudLoadBtn.addEventListener("click", () => loadCloudState({ silent: false }));
  els.cloudSaveBtn.addEventListener("click", () => saveCloudState({ manual: true }));
  els.cloudShareBtn.addEventListener("click", copyCloudShareLink);

  wireDropZone(document.body);
  restoreCloudConfig();
  restoreCloudConfigFromUrl();
  if (restoreState()) {
    logActivity("Restored saved scorecard data from this browser.");
  } else {
    logActivity("Ready. Load the master VTC scorecard to begin.");
  }
  if (initCloudClient()) {
    loadCloudState({ silent: true });
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
    const sheet = wb.Sheets[sheetName];
    const rows = normalized === "Config"
      ? readConfigSheet(sheet)
      : XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(cleanRow);
    const detected = normalized || detectSheetType(rows);
    if (!detected) continue;
    const normalizedRows = detected === "Schedule_Adherence_Raw" ? normalizeScheduleRows(rows) : rows;
    tables[detected] = mergeRows(tables[detected], normalizedRows);
  }
  return tables;
}

function detectSheetType(rows) {
  if (!rows.length) return null;
  const canonicalHeaders = new Set();
  rows.slice(0, 10).forEach((row) => {
    Object.keys(row).forEach((key) => {
      const canonical = canonicalScheduleKey(key);
      if (canonical) canonicalHeaders.add(canonical);
    });
  });

  const hasVendor = canonicalHeaders.has("Vendor_ID") || canonicalHeaders.has("Vendor_Name");
  const hasScheduleMeasure = canonicalHeaders.has("Monthly_Tasks")
    || canonicalHeaders.has("No_Show_Count")
    || canonicalHeaders.has("Adherence_Pct");
  return hasVendor && hasScheduleMeasure ? "Schedule_Adherence_Raw" : null;
}

function normalizeScheduleRows(rows) {
  return rows
    .map((row) => {
      const normalized = {
        Month: "",
        Vendor_Adj: "",
        Vendor_ID: "",
        Vendor_Name: "",
        Monthly_Tasks: 0,
        No_Show_Count: 0,
        Adherence_Pct: ""
      };

      Object.entries(row).forEach(([key, value]) => {
        const canonical = canonicalScheduleKey(key);
        if (!canonical) return;
        if (canonical === "Monthly_Tasks" || canonical === "No_Show_Count") {
          normalized[canonical] = toNumber(value);
        } else if (canonical === "Adherence_Pct") {
          normalized[canonical] = toNumber(value);
        } else {
          normalized[canonical] = cleanText(value);
        }
      });

      if (!normalized.Vendor_Adj) normalized.Vendor_Adj = normalized.Vendor_Name;
      if (!normalized.Vendor_Name) normalized.Vendor_Name = normalized.Vendor_Adj;
      if (!hasExplicitScheduleCount(row, "No_Show_Count") && looksLikeNoShowRow(row)) {
        normalized.No_Show_Count = 1;
      }

      return normalized;
    })
    .filter((row) => row.Vendor_ID || row.Vendor_Name || row.Vendor_Adj);
}

function canonicalScheduleKey(key) {
  const compact = normalizeHeaderKey(key);
  return SCHEDULE_COLUMN_ALIASES[compact] || null;
}

function normalizeHeaderKey(key) {
  return cleanText(key)
    .toLowerCase()
    .replace(/[%#]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function hasExplicitScheduleCount(row, canonicalName) {
  return Object.keys(row).some((key) => canonicalScheduleKey(key) === canonicalName && cleanText(row[key]) !== "");
}

function looksLikeNoShowRow(row) {
  return Object.entries(row).some(([key, value]) => {
    const text = `${key} ${value}`.toLowerCase();
    return /\bno\s*-?\s*show\b|\bnoshow\b/.test(text) && !/^\s*(0|false|no)?\s*$/.test(cleanText(value).toLowerCase());
  });
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
  if (value === null || value === undefined) return "";
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

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = cleanText(value).replace(/[$,%(),]/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function calculateScores() {
  const byVendorName = new Map(state.vendors.map((vendor) => [cleanText(vendor.name).toLowerCase(), vendor]));
  const scheduleById = groupBy(state.tables.Schedule_Adherence_Raw, (row) => cleanText(row.Vendor_ID));
  const safetyByVendor = groupBy(state.tables.Safety, (row) => cleanText(row.Vendor_Name).toLowerCase());
  const reworkByVendor = groupBy(state.tables.Rework, (row) => cleanText(row.Vendor_Name).toLowerCase());
  const logByVendor = groupBy(state.tables.Log, (row) => cleanText(row["Vendor Name Clean"] || row["Vendor Name"]).toLowerCase());
  const workload = readWorkload();

  return state.vendors.map((vendor) => {
    const key = cleanText(vendor.name).toLowerCase();
    const safetyRows = safetyByVendor.get(key) || [];
    const scheduleRows = scheduleById.get(cleanText(vendor.id)) || [];
    const reworkRows = reworkByVendor.get(key) || [];
    const logRows = logByVendor.get(key) || [];

    const safety = safetyRows.length
      ? Math.max(0, 100 - sum(safetyRows, (row) => Number(row.Severity_Score || 0)) * 10)
      : 100;
    const scheduleTasks = sum(scheduleRows, (row) => Number(row.Monthly_Tasks || 0));
    const noShows = sum(scheduleRows, (row) => Number(row.No_Show_Count || 0));
    const schedule = scheduleRows.length && scheduleTasks > 0
      ? Math.max(0, Math.min(100, ((scheduleTasks - noShows) / scheduleTasks) * 100))
      : 100;
    const rework = reworkRows.length
      ? Math.max(0, 100 - sum(reworkRows, (row) => Number(row.PenaltyPoints || 0)) * 5)
      : null;
    const log = logRows.length ? average(logRows, (row) => Number(row.Points || 0)) : null;

    const components = {
      Safety: safety,
      Schedule: schedule,
      Rework: rework,
      Log: log
    };
    const overall = weightedScore(components);
    return {
      ...vendor,
      safety,
      schedule,
      rework,
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

function parseFeedback(text, meta = {}) {
  const cleaned = cleanText(text);
  if (!cleaned) return;
  const vendor = bestVendorMatch(cleaned);
  const submitted = meta.from || matchValue(text, /(from|submitted by|builder)[:\s]+([^\n<]+)/i);
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
  let lastMeta = {};
  for (const file of files) {
    const raw = await file.text();
    const parsed = /\.eml$/i.test(file.name) ? parseEml(raw) : { body: raw, from: "" };
    chunks.push(parsed.body);
    lastMeta = parsed;
  }
  els.feedbackText.value = [els.feedbackText.value, ...chunks].filter(Boolean).join("\n\n");
  parseFeedback(els.feedbackText.value, lastMeta);
}

function parseEml(raw) {
  const normalized = raw.replace(/\r\n/g, "\n");
  const headerEnd = normalized.indexOf("\n\n");
  const headerText = headerEnd >= 0 ? normalized.slice(0, headerEnd) : "";
  const bodyText = headerEnd >= 0 ? normalized.slice(headerEnd + 2) : normalized;
  const headers = parseEmailHeaders(headerText);
  const contentType = headers["content-type"] || "";
  const boundary = contentType.match(/boundary="?([^";\n]+)"?/i)?.[1];
  const parts = boundary ? splitMimeParts(bodyText, boundary) : [{ headers, body: bodyText }];
  const decodedParts = parts.map((part) => ({
    ...part,
    decodedBody: decodeMimeBody(part.body, part.headers["content-transfer-encoding"])
  }));
  const plain = decodedParts.find((part) => /text\/plain/i.test(part.headers["content-type"] || ""));
  const html = decodedParts.find((part) => /text\/html/i.test(part.headers["content-type"] || ""));
  const fallback = decodedParts.find((part) => cleanText(part.decodedBody));
  const body = cleanEmailBody(plain?.decodedBody || htmlToText(html?.decodedBody || "") || fallback?.decodedBody || bodyText);
  return {
    body,
    from: cleanEmailAddress(headers.from || ""),
    subject: decodeMimeWords(headers.subject || "")
  };
}

function parseEmailHeaders(headerText) {
  const unfolded = headerText.replace(/\n[ \t]+/g, " ");
  const headers = {};
  unfolded.split("\n").forEach((line) => {
    const separator = line.indexOf(":");
    if (separator <= 0) return;
    const key = line.slice(0, separator).trim().toLowerCase();
    headers[key] = decodeMimeWords(line.slice(separator + 1).trim());
  });
  return headers;
}

function splitMimeParts(bodyText, boundary) {
  return bodyText
    .split(`--${boundary}`)
    .filter((part) => part.trim() && !part.trim().startsWith("--"))
    .map((part) => {
      const cleanPart = part.replace(/^\n+|\n+$/g, "");
      const headerEnd = cleanPart.indexOf("\n\n");
      const headerText = headerEnd >= 0 ? cleanPart.slice(0, headerEnd) : "";
      const body = headerEnd >= 0 ? cleanPart.slice(headerEnd + 2) : cleanPart;
      return { headers: parseEmailHeaders(headerText), body };
    });
}

function decodeMimeBody(body, encoding = "") {
  const normalizedEncoding = cleanText(encoding).toLowerCase();
  if (normalizedEncoding === "base64") {
    return decodeBase64Text(body);
  }
  if (normalizedEncoding === "quoted-printable") {
    return decodeQuotedPrintable(body);
  }
  return body;
}

function decodeBase64Text(value) {
  try {
    const binary = atob(value.replace(/\s/g, ""));
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return new TextDecoder("utf-8").decode(bytes);
  } catch {
    return value;
  }
}

function decodeQuotedPrintable(value) {
  const withoutSoftBreaks = value.replace(/=\n/g, "");
  const bytes = [];
  for (let index = 0; index < withoutSoftBreaks.length; index += 1) {
    if (withoutSoftBreaks[index] === "=" && /^[0-9A-F]{2}$/i.test(withoutSoftBreaks.slice(index + 1, index + 3))) {
      bytes.push(parseInt(withoutSoftBreaks.slice(index + 1, index + 3), 16));
      index += 2;
    } else {
      bytes.push(withoutSoftBreaks.charCodeAt(index));
    }
  }
  try {
    return new TextDecoder("utf-8").decode(new Uint8Array(bytes));
  } catch {
    return withoutSoftBreaks;
  }
}

function decodeMimeWords(value) {
  return String(value || "").replace(/=\?([^?]+)\?([BQ])\?([^?]+)\?=/gi, (_, charset, encoding, text) => {
    if (encoding.toUpperCase() === "B") return decodeBase64Text(text);
    return decodeQuotedPrintable(text.replace(/_/g, " "));
  });
}

function htmlToText(html) {
  if (!html) return "";
  const doc = new DOMParser().parseFromString(html, "text/html");
  return doc.body?.innerText || "";
}

function cleanEmailBody(body) {
  return cleanText(body
    .replace(/\nOn .+ wrote:\n[\s\S]*$/i, "")
    .replace(/\nFrom: .+\nSent: .+\nTo: .+[\s\S]*$/i, "")
    .replace(/\n_{5,}[\s\S]*$/g, "")
    .replace(/\n-{5,}[\s\S]*$/g, ""));
}

function cleanEmailAddress(value) {
  const decoded = decodeMimeWords(value);
  const match = decoded.match(/^"?([^"<]+)"?\s*</) || decoded.match(/([^<>\s]+@[^<>\s]+)/);
  return cleanText(match ? match[1] : decoded);
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
    Rework_Score: round(score.rework),
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
  renderCloudStatus();
  if (els.appVersion) els.appVersion.textContent = APP_VERSION;
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
    <button id="resetWeightsBtn" class="secondary weight-reset" type="button">
      <i data-lucide="undo-2"></i>
      Reset to default
    </button>
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
  document.getElementById("resetWeightsBtn")?.addEventListener("click", resetWeightsToDefault);
}

function resetWeightsToDefault() {
  state.weights = { ...DEFAULT_WEIGHTS };
  state.scores = calculateScores();
  logActivity("Reset weights to default settings.");
  saveState();
  render();
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
  const vendors = tradeRankings();
  const best = vendors.slice(0, 5);
  const worst = [...vendors].reverse().slice(0, 5);
  els.topBestTrades.innerHTML = renderTradeList(best, "No vendor scores loaded yet.");
  els.topWorstTrades.innerHTML = renderTradeList(worst, "No vendor scores loaded yet.");
}

function tradeRankings() {
  return state.scores
    .filter((score) => score.name && Number.isFinite(score.overall))
    .map((score) => ({
      name: score.name,
      category: score.category,
      score: score.overall
    }))
    .sort((a, b) => b.score - a.score || a.name.localeCompare(b.name));
}

function renderTradeList(trades, emptyText) {
  if (!trades.length) return `<li class="empty-trade">${emptyText}</li>`;
  return trades.map((trade) => `
    <li>
      <span>${escapeHtml(trade.name)}</span>
      <strong>${formatScore(trade.score)}</strong>
      <small>${escapeHtml(trade.category || "Uncategorized")}</small>
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

function saveState(options = { syncCloud: true }) {
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
    if (options.syncCloud) queueCloudSave();
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

function getPersistedPayload() {
  return {
    fileName: state.fileName,
    tables: state.tables,
    vendors: state.vendors,
    weights: state.weights,
    activity: state.activity,
    lastSaved: state.lastSaved
  };
}

function applyPersistedPayload(payload) {
  state.fileName = payload.fileName || "";
  state.tables = { ...emptyTables(), ...(payload.tables || {}) };
  state.vendors = payload.vendors?.length ? payload.vendors : readVendors(state.tables.Config);
  state.weights = normalizeWeights({ ...DEFAULT_WEIGHTS, ...(payload.weights || {}) });
  state.activity = Array.isArray(payload.activity) ? payload.activity.slice(0, 80) : [];
  state.lastSaved = payload.lastSaved || new Date().toISOString();
  state.scores = calculateScores();
}

function restoreCloudConfig() {
  try {
    const saved = localStorage.getItem(CLOUD_CONFIG_KEY);
    if (!saved) return;
    const config = JSON.parse(saved);
    state.cloud.url = config.url || "";
    state.cloud.anonKey = config.anonKey || "";
    state.cloud.workspaceId = config.workspaceId || "vtc-main";
    els.cloudUrl.value = state.cloud.url;
    els.cloudAnonKey.value = state.cloud.anonKey;
    els.cloudWorkspace.value = state.cloud.workspaceId;
  } catch (error) {
    console.warn("Unable to restore Supabase config", error);
    localStorage.removeItem(CLOUD_CONFIG_KEY);
  }
}

function restoreCloudConfigFromUrl() {
  const hash = window.location.hash.startsWith("#") ? window.location.hash.slice(1) : "";
  if (!hash) return;
  const params = new URLSearchParams(hash);
  const cloudUrl = params.get("supabaseUrl");
  const cloudKey = params.get("supabaseKey");
  const workspaceId = params.get("workspaceId");
  if (!cloudUrl || !cloudKey) return;

  state.cloud.url = normalizeCloudUrl(cloudUrl);
  state.cloud.anonKey = normalizeCloudKey(cloudKey);
  state.cloud.workspaceId = cleanText(workspaceId) || "vtc-main";
  els.cloudUrl.value = state.cloud.url;
  els.cloudAnonKey.value = state.cloud.anonKey;
  els.cloudWorkspace.value = state.cloud.workspaceId;
  localStorage.setItem(CLOUD_CONFIG_KEY, JSON.stringify({
    url: state.cloud.url,
    anonKey: state.cloud.anonKey,
    workspaceId: state.cloud.workspaceId
  }));
  logActivity("Loaded cloud connection settings from share link.");
}

function connectCloudFromForm() {
  state.cloud.url = normalizeCloudUrl(els.cloudUrl.value);
  state.cloud.anonKey = normalizeCloudKey(els.cloudAnonKey.value);
  state.cloud.workspaceId = cleanText(els.cloudWorkspace.value) || "vtc-main";
  if (!state.cloud.url) {
    setCloudStatus("Missing URL", false, true);
    logActivity("Enter the Supabase Project URL before connecting.");
    return;
  }
  if (!state.cloud.anonKey) {
    setCloudStatus("Missing key", false, true);
    logActivity("Enter both the Supabase Project URL and anon/public key before connecting. Copy the key value only, not the label.");
    return;
  }
  localStorage.setItem(CLOUD_CONFIG_KEY, JSON.stringify({
    url: state.cloud.url,
    anonKey: state.cloud.anonKey,
    workspaceId: state.cloud.workspaceId
  }));
  if (initCloudClient()) {
    logActivity(`Connected cloud workspace: ${state.cloud.workspaceId}`);
    loadCloudState({ silent: false });
  }
  renderCloudStatus();
}

async function copyCloudShareLink() {
  state.cloud.url = normalizeCloudUrl(els.cloudUrl.value);
  state.cloud.anonKey = normalizeCloudKey(els.cloudAnonKey.value);
  state.cloud.workspaceId = cleanText(els.cloudWorkspace.value) || "vtc-main";
  if (!state.cloud.url || !state.cloud.anonKey) {
    setCloudStatus("Missing settings", false, true);
    logActivity("Enter the Supabase Project URL and key before copying a share link.");
    return;
  }

  const url = new URL(window.location.href);
  url.search = "";
  url.hash = new URLSearchParams({
    supabaseUrl: state.cloud.url,
    supabaseKey: state.cloud.anonKey,
    workspaceId: state.cloud.workspaceId
  }).toString();

  try {
    await navigator.clipboard.writeText(url.toString());
    logActivity("Copied team share link. Send that link after saving data to Supabase.");
  } catch {
    window.prompt("Copy this team share link:", url.toString());
  }
}

function initCloudClient() {
  if (!state.cloud.url || !state.cloud.anonKey) {
    setCloudStatus("Not connected", false);
    return false;
  }
  setCloudStatus("Connected", true);
  return true;
}

function normalizeCloudUrl(value) {
  const rawUrl = cleanText(value)
    .replace(/^project\s*url\s*[:=]\s*/i, "")
    .replace(/^url\s*[:=]\s*/i, "")
    .replace(/^["']|["']$/g, "")
    .replace(/\/+$/, "");
  if (!rawUrl) return "";
  if (/^[a-z0-9-]{15,}$/i.test(rawUrl) && !rawUrl.includes(".")) return `https://${rawUrl}.supabase.co`;
  try {
    const parsed = new URL(rawUrl.startsWith("http") ? rawUrl : `https://${rawUrl}`);
    return `${parsed.protocol}//${parsed.hostname}`;
  } catch {
    return rawUrl;
  }
}

function normalizeCloudKey(value) {
  return cleanText(value)
    .replace(/^anon\s*(public)?\s*key\s*[:=]\s*/i, "")
    .replace(/^publishable\s*key\s*[:=]\s*/i, "")
    .replace(/^["']|["']$/g, "");
}

function setCloudStatus(status, connected, isError = false) {
  state.cloud.status = status;
  state.cloud.connected = connected;
  state.cloud.error = isError;
  renderCloudStatus();
}

function renderCloudStatus() {
  if (!els.cloudStatus) return;
  els.cloudStatus.textContent = state.cloud.status;
  els.cloudStatus.classList.toggle("connected", Boolean(state.cloud.connected));
  els.cloudStatus.classList.toggle("error", Boolean(state.cloud.error));
  if (!els.cloudUrl.value && state.cloud.url) els.cloudUrl.value = state.cloud.url;
  if (!els.cloudAnonKey.value && state.cloud.anonKey) els.cloudAnonKey.value = state.cloud.anonKey;
  if (!els.cloudWorkspace.value) els.cloudWorkspace.value = state.cloud.workspaceId || "vtc-main";
}

function queueCloudSave() {
  if (!state.cloud.connected || !hasScorecardData()) return;
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = setTimeout(() => {
    saveCloudState({ manual: false });
  }, 900);
}

async function saveCloudState({ manual }) {
  if (!state.cloud.connected && !initCloudClient()) {
    if (manual) logActivity("Connect Supabase before saving to cloud.");
    return;
  }
  if (!hasScorecardData()) {
    if (manual) logActivity("Load scorecard data before saving to cloud.");
    return;
  }
  setCloudStatus("Saving...", true);
  const payload = getPersistedPayload();
  const { error } = await supabaseRequest("scorecard_states", {
    method: "POST",
    headers: {
      Prefer: "resolution=merge-duplicates,return=minimal"
    },
    body: JSON.stringify({
      id: state.cloud.workspaceId || "vtc-main",
      payload,
      updated_at: new Date().toISOString()
    })
  });

  if (error) {
    console.error(error);
    setCloudStatus("Save failed", false, true);
    logActivity(`Supabase save failed: ${formatCloudError(error)}`);
    return;
  }
  setCloudStatus("Saved", true);
  if (manual) logActivity("Saved scorecard data to Supabase.");
}

async function loadCloudState({ silent }) {
  if (!state.cloud.connected && !initCloudClient()) {
    if (!silent) logActivity("Connect Supabase before loading cloud data.");
    return;
  }
  setCloudStatus("Loading...", true);
  const id = encodeURIComponent(state.cloud.workspaceId || "vtc-main");
  const { data, error } = await supabaseRequest(`scorecard_states?select=payload,updated_at&id=eq.${id}&limit=1`);

  if (error) {
    console.error(error);
    setCloudStatus("Load failed", false, true);
    logActivity(`Supabase load failed: ${formatCloudError(error)}`);
    return;
  }
  const row = Array.isArray(data) ? data[0] : null;
  if (!row?.payload) {
    setCloudStatus("No cloud data", true);
    if (!silent) logActivity("No Supabase scorecard data exists for this workspace yet.");
    return;
  }

  const localSaved = state.lastSaved ? new Date(state.lastSaved).getTime() : 0;
  const cloudSaved = row.payload.lastSaved ? new Date(row.payload.lastSaved).getTime() : new Date(row.updated_at).getTime();
  if (silent && localSaved > cloudSaved) {
    setCloudStatus("Local newer", true);
    return;
  }

  applyPersistedPayload(row.payload);
  saveState({ syncCloud: false });
  setCloudStatus("Loaded", true);
  logActivity("Loaded scorecard data from Supabase.");
  render();
}

async function supabaseRequest(path, options = {}) {
  try {
    const response = await fetch(`${state.cloud.url}/rest/v1/${path}`, {
      method: options.method || "GET",
      headers: {
        apikey: state.cloud.anonKey,
        Authorization: `Bearer ${state.cloud.anonKey}`,
        "Content-Type": "application/json",
        ...(options.headers || {})
      },
      body: options.body
    });
    const text = await response.text();
    const data = parseJsonOrText(text);
    if (!response.ok) {
      return {
        data: null,
        error: {
          status: response.status,
          statusText: response.statusText,
          ...(typeof data === "object" && data !== null ? data : { message: String(data || response.statusText) })
        }
      };
    }
    return { data, error: null };
  } catch (error) {
    return { data: null, error };
  }
}

function parseJsonOrText(text) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function formatCloudError(error) {
  if (!error) return "Unknown error";
  const parts = [
    error.status ? `${error.status}` : "",
    error.message || error.statusText || "",
    error.details || "",
    error.hint || "",
    error.code || ""
  ].filter(Boolean);
  return parts.join(" - ") || String(error);
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
    const droppedOnSourceInput = event.target.closest?.(".dropzone")?.querySelector?.("#sourceInput");
    if (!state.workbook && workbookFiles.length && !droppedOnSourceInput) {
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
