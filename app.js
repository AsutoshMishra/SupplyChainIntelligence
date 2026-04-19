/** Tried in order: your root file first, then legacy paths. */
const SUPPLIER_CSV_BEFORE_CANDIDATES = [
  "SupplierCommitmentsBefore.csv",
  "SupplierCommitments.csv",
  "Assets/SupplierCommittments.csv",
  "assets/SupplierCommittments.csv",
];

const SUPPLIER_CSV_AFTER_CANDIDATES = [
  "SupplierCommitmentsAfter.csv",
  "SupplierCommitments.csv",
  "Assets/SupplierCommittments.csv",
  "assets/SupplierCommittments.csv",
];

const GLOBAL_SUPPLY_OUTLOOK_CSV_CANDIDATES = ["GlobalSupplyOutlook.csv", "Assets/GlobalSupplyOutlook.csv"];

const INVENTORY_STATUS_BEFORE_CSV_CANDIDATES = [
  "InventoryStatus.csv",
  "InventoryStatusBefore.csv",
  "Assets/InventoryStatus.csv",
];

const INVENTORY_STATUS_AFTER_CSV_CANDIDATES = [
  "InventoryStatusAfter.csv",
  "InventoryStatus.csv",
  "Assets/InventoryStatus.csv",
];
const MPS_BEFORE_CSV_CANDIDATES = ["MPSBefore.csv", "MPS.csv", "Assets/MPS.csv"];
const MPS_AFTER_CSV_CANDIDATES = ["MPSAfter.csv", "MPS.csv", "Assets/MPS.csv"];
const SEQUENCE_BEFORE_CSV_CANDIDATES = ["SequenceBefore.csv", "Assets/SequenceBefore.csv"];
const SEQUENCE_AFTER_CSV_CANDIDATES = ["SequenceAfter.csv", "Assets/SequenceAfter.csv"];
let cachedOutlookHeaders = [];
let cachedOutlookRows = [];
let hasAnalyzedOutlook = false;
let supplierCommitmentsPhase = "before";
let inventoryInsightsPhase = "before";
let mpsPhase = "before";
let sequenceCycleTicks = 0;
let sequenceCellTotal = 0;
let currentSequenceHeaders = [];
let currentSequenceBody = [];
let currentSequenceBaseSchedule = [];
/** Normalized rows: { date, shift, slots: string[9] } in display order (CSV columns 1–9 per shift) */
let currentSequenceSchedule = [];
/** Supplier matrix (Supply View): first row baseline B/R/W per column; extra rows render static */
let inventoryMatrixHeaders = [];
let inventoryMatrixBaseline = [];
let inventoryMatrixExtraRows = [];
/** After "Re-sequence Production", inventory shows dice restored for delayed suppliers + "(Replenished)" labels. Reset on Analyze. */
let inventoryReplenishedAfterPartnerResequence = false;
/** When true, `runLiveProductionLoop` advances ticks until stopped; wraps with reset at sequence end. */
let liveProductionRunActive = false;
let alertsUnreadCount = 0;
let alertsPanelOpen = false;
let alertsFeed = [];
let alertsFeedSeq = 0;
const APP_UI_STATE_KEY = "cfi_ui_state_v1";
const APP_UI_SYNC_CHANNEL = "cfi_ui_sync_v1";
let uiSyncChannel = null;

function currentAppUiState() {
  return {
    supplierCommitmentsPhase,
    inventoryInsightsPhase,
    mpsPhase,
    hasAnalyzedOutlook,
  };
}

function persistAppUiState() {
  try {
    const payload = JSON.stringify(currentAppUiState());
    window.localStorage.setItem(APP_UI_STATE_KEY, payload);
    if (uiSyncChannel) {
      uiSyncChannel.postMessage({ type: "ui-state-updated", payload });
    }
  } catch {
    /* ignore storage failures */
  }
}

function applyAppUiState(state) {
  if (!state || typeof state !== "object") return;
  supplierCommitmentsPhase = state.supplierCommitmentsPhase === "after" ? "after" : "before";
  inventoryInsightsPhase = state.inventoryInsightsPhase === "after" ? "after" : "before";
  mpsPhase = state.mpsPhase === "after" ? "after" : "before";
  hasAnalyzedOutlook = Boolean(state.hasAnalyzedOutlook);
}

function hydrateAppUiState() {
  try {
    const raw = window.localStorage.getItem(APP_UI_STATE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    applyAppUiState(parsed);
  } catch {
    /* ignore parse failures */
  }
}

function appUiStateSignature() {
  return JSON.stringify(currentAppUiState());
}

let lastAppliedUiStateSig = "";

async function reconcileUiStateFromStorage() {
  let parsed = null;
  try {
    const raw = window.localStorage.getItem(APP_UI_STATE_KEY);
    if (!raw) return;
    parsed = JSON.parse(raw);
  } catch {
    return;
  }
  if (!parsed || typeof parsed !== "object") return;
  const beforeSig = appUiStateSignature();
  applyAppUiState(parsed);
  const afterSig = appUiStateSignature();
  if (afterSig === beforeSig || afterSig === lastAppliedUiStateSig) return;
  lastAppliedUiStateSig = afterSig;
  if (!hasAnalyzedOutlook) {
    renderOutlookAlertsIdle("No impact analysis yet. Click Analyze in Global Supply Outlook.");
  }
  await loadSupplierCommitments();
  const invPanel = document.getElementById("panel-inventory");
  if (invPanel?.dataset.inventoryLoaded === "1") {
    await loadInventoryInsights();
  }
  const prodPanel = document.getElementById("panel-production");
  if (prodPanel?.dataset.mpsLoaded === "1") {
    await loadProductionPlanningInsights();
  }
}

async function fetchCsvFirstOk(candidates) {
  let lastStatus = 404;
  for (const path of candidates) {
    const url = new URL(path, import.meta.url);
    const res = await fetch(url.href, { cache: "no-store" });
    if (res.ok) return await res.text();
    lastStatus = res.status;
  }
  throw new Error(`Could not load CSV (last HTTP ${lastStatus}). Tried: ${candidates.join(", ")}`);
}

function fetchSupplierCsvText() {
  const candidates =
    supplierCommitmentsPhase === "after"
      ? SUPPLIER_CSV_AFTER_CANDIDATES
      : SUPPLIER_CSV_BEFORE_CANDIDATES;
  return fetchCsvFirstOk(candidates);
}

function fetchGlobalSupplyOutlookCsvText() {
  return fetchCsvFirstOk(GLOBAL_SUPPLY_OUTLOOK_CSV_CANDIDATES);
}

function fetchInventoryStatusCsvText() {
  const candidates =
    inventoryInsightsPhase === "after"
      ? INVENTORY_STATUS_AFTER_CSV_CANDIDATES
      : INVENTORY_STATUS_BEFORE_CSV_CANDIDATES;
  return fetchCsvFirstOk(candidates);
}

function fetchMpsCsvText() {
  const candidates = mpsPhase === "after" ? MPS_AFTER_CSV_CANDIDATES : MPS_BEFORE_CSV_CANDIDATES;
  return fetchCsvFirstOk(candidates);
}

function fetchSequenceCsvText() {
  const candidates = mpsPhase === "after" ? SEQUENCE_AFTER_CSV_CANDIDATES : SEQUENCE_BEFORE_CSV_CANDIDATES;
  return fetchCsvFirstOk(candidates);
}

function waitMs(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

/** Returns false if `shouldAbort()` is true before `ms` ms have elapsed (checked every 50ms). */
async function waitMsInterruptible(ms, shouldAbort) {
  const step = 50;
  let left = ms;
  while (left > 0) {
    if (shouldAbort()) return false;
    await waitMs(Math.min(step, left));
    left -= step;
  }
  return true;
}

/** Minimal RFC-style CSV parser (quoted fields, commas, CRLF/LF). */
function parseCSV(text) {
  const rows = [];
  let field = "";
  let row = [];
  let i = 0;
  let inQuotes = false;
  const s = text.replace(/^\uFEFF/, "");

  const pushField = () => {
    row.push(field);
    field = "";
  };
  const pushRow = () => {
    if (row.some((c) => c.trim() !== "")) rows.push(row);
    row = [];
  };

  while (i < s.length) {
    const c = s[i];
    if (inQuotes) {
      if (c === '"') {
        if (s[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += c;
      i++;
      continue;
    }
    if (c === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (c === ",") {
      pushField();
      i++;
      continue;
    }
    if (c === "\r") {
      if (s[i + 1] === "\n") i++;
      pushField();
      pushRow();
      i++;
      continue;
    }
    if (c === "\n") {
      pushField();
      pushRow();
      i++;
      continue;
    }
    field += c;
    i++;
  }
  pushField();
  pushRow();
  return rows;
}

function normKey(h) {
  return String(h).trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** First matching column index, or -1. */
function resolveCol(rawHeaders, ...canonicalNames) {
  const norms = rawHeaders.map(normKey);
  for (const name of canonicalNames) {
    const n = normKey(name);
    const i = norms.indexOf(n);
    if (i >= 0) return i;
  }
  return -1;
}

function seatPillClass(seatType) {
  const t = seatType.trim().toLowerCase();
  if (t.includes("blue")) return "pill pill--blue";
  if (t.includes("black")) return "pill pill--dark";
  if (t.includes("white")) return "pill pill--light";
  if (t.includes("all seat")) return "pill pill--neutral";
  return "pill pill--neutral";
}

function statusBadgeClass(status) {
  const t = status.trim().toLowerCase();
  if (t.includes("on track")) return "badge badge--green";
  if (t.includes("in transit")) return "badge badge--blue";
  if (t.includes("at risk")) return "badge badge--red-sm";
  if (t.includes("delay") || t.includes("minor")) return "badge badge--orange-sm";
  if (t.includes("critical") || t.includes("alert")) return "badge badge--red-sm";
  return "badge badge--green";
}

function formatQty(q) {
  const n = String(q).replace(/,/g, "").trim();
  if (n === "") return "";
  const num = Number(n);
  if (!Number.isFinite(num)) return escapeHtml(q);
  return escapeHtml(num.toLocaleString("en-US"));
}

function avatarLetter(supplier, avatarCol) {
  const a = (avatarCol || "").trim();
  if (a.length > 0) return escapeHtml(a.slice(0, 1).toUpperCase());
  const name = supplier.trim();
  if (!name) return "?";
  return escapeHtml(name.slice(0, 1).toUpperCase());
}

function monthIdxFromAbbr(mon) {
  const m = String(mon || "").slice(0, 3).toLowerCase();
  const map = {
    jan: 0,
    feb: 1,
    mar: 2,
    apr: 3,
    may: 4,
    jun: 5,
    jul: 6,
    aug: 7,
    sep: 8,
    oct: 9,
    nov: 10,
    dec: 11,
  };
  return Object.prototype.hasOwnProperty.call(map, m) ? map[m] : -1;
}

function formatDateAsMonDay(d) {
  const mons = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${mons[d.getMonth()]}-${d.getDate()}`;
}

function shiftNumberFromText(raw) {
  const m = String(raw || "").match(/(\d+)/);
  const n = m ? Number(m[1]) : 1;
  return Number.isFinite(n) ? Math.min(3, Math.max(1, n)) : 1;
}

function addDelayToCommitted(dateRaw, shiftRaw, delayShifts) {
  const shift = shiftNumberFromText(shiftRaw);
  const m = String(dateRaw || "").match(/^([A-Za-z]{3})[-\s](\d{1,2})$/);
  if (!m) {
    const nextShift = ((shift - 1 + Math.max(0, delayShifts)) % 3) + 1;
    return { dateText: String(dateRaw || "").trim(), shiftText: `Shift ${nextShift}` };
  }
  const mi = monthIdxFromAbbr(m[1]);
  const day = Number(m[2]);
  if (mi < 0 || !Number.isFinite(day)) {
    const nextShift = ((shift - 1 + Math.max(0, delayShifts)) % 3) + 1;
    return { dateText: String(dateRaw || "").trim(), shiftText: `Shift ${nextShift}` };
  }
  const year = new Date().getFullYear();
  const baseDate = new Date(year, mi, day);
  const total = (shift - 1) + Math.max(0, delayShifts);
  const dayAdd = Math.floor(total / 3);
  const nextShift = (total % 3) + 1;
  baseDate.setDate(baseDate.getDate() + dayAdd);
  return { dateText: formatDateAsMonDay(baseDate), shiftText: `Shift ${nextShift}` };
}

function rowFromRecord(rec) {
  const supplier = escapeHtml(rec.supplier || "");
  const location = escapeHtml(rec.location || "");
  const partName = escapeHtml(rec.partname || "");
  const partId = escapeHtml(rec.partid || "");
  const seatType = rec.seattype || "";
  const seatEscaped = escapeHtml(seatType);
  let poCell = formatQty(rec.poqty ?? "");
  if (rec.unit && poCell) {
    poCell += ` <span class="muted">${escapeHtml(rec.unit)}</span>`;
  } else if (rec.unit) {
    poCell = escapeHtml(rec.unit);
  }
  const liveDelay = supplierDelayShiftByName(rec.supplier || "");
  const committedAdjusted = addDelayToCommitted(rec.committed || "", rec.committedshift || "", liveDelay);
  const committedDate = escapeHtml(committedAdjusted.dateText || "");
  const committedShift = escapeHtml(committedAdjusted.shiftText || "");
  const status = liveDelay > 0 ? `Delayed (${liveDelay} Shift${liveDelay > 1 ? "s" : ""})` : rec.status || "";
  const statusEscaped = escapeHtml(status);
  const otif = escapeHtml(rec.otif || "");
  const av = avatarLetter(rec.supplier || "", rec.avatar || "");
  const desc = (rec.description || "").trim();
  const partTitle = desc ? escapeHtml(desc) : "";

  return `<tr>
    <td><div class="cell-supplier"><span class="avatar">${av}</span><span>${supplier}</span></div></td>
    <td>${location}</td>
    <td${partTitle ? ` title="${partTitle}"` : ""}>${partName} <span class="muted">(${partId})</span></td>
    <td><span class="${seatPillClass(seatType)}">${seatEscaped}</span></td>
    <td>${poCell}</td>
    <td>${committedDate}</td>
    <td>${committedShift}</td>
    <td><span class="${statusBadgeClass(status)}">${statusEscaped}</span></td>
    <td>${otif}</td>
  </tr>`;
}

function recordsFromRows(rows) {
  if (rows.length < 2) return [];
  const rawHeaders = rows[0];

  const iSupplier = resolveCol(rawHeaders, "Supplier");
  const iLocation = resolveCol(rawHeaders, "Location");
  const iPart = resolveCol(rawHeaders, "PartName", "MaterialPart", "Material / Part");
  const iPartId = resolveCol(rawHeaders, "PartId", "PartID", "Part ID");
  const iSeat = resolveCol(rawHeaders, "SeatType", "Seat Type");
  const iDesc = resolveCol(rawHeaders, "Description");
  const iPoQty = resolveCol(rawHeaders, "POQty", "PO Qty");
  const iUnit = resolveCol(rawHeaders, "Unit");
  const iCommitted = resolveCol(rawHeaders, "Committed", "Committed Date");
  const iCommittedShift = resolveCol(rawHeaders, "Committed Shift", "Commited Shift", "Shift");
  const iStatus = resolveCol(rawHeaders, "Status");
  const iOtif = resolveCol(rawHeaders, "OTIF", "OTIF %");
  const iAvatar = resolveCol(rawHeaders, "Avatar");
  const iLat = resolveCol(rawHeaders, "Latitude", "Lat");
  const iLng = resolveCol(rawHeaders, "Longitude", "Lng", "Lon");
  const iSupplyRisk = resolveCol(rawHeaders, "SupplyRisk", "Supply Risk", "Risk");

  if (iSupplier < 0 || iPart < 0) {
    throw new Error(
      "CSV needs a Supplier column and a part column (PartName or Material / Part). See SupplierCommitments.csv."
    );
  }

  const get = (cells, idx) => {
    if (idx < 0) return "";
    return cells[idx] != null ? String(cells[idx]).trim() : "";
  };

  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const cells = rows[r];
    if (!cells || cells.every((c) => !String(c).trim())) continue;
    out.push({
      supplier: get(cells, iSupplier),
      location: get(cells, iLocation),
      partname: get(cells, iPart),
      partid: get(cells, iPartId),
      seattype: get(cells, iSeat),
      description: get(cells, iDesc),
      poqty: get(cells, iPoQty),
      unit: get(cells, iUnit),
      committed: get(cells, iCommitted),
      committedshift: get(cells, iCommittedShift),
      status: get(cells, iStatus),
      otif: get(cells, iOtif),
      avatar: get(cells, iAvatar),
      latitude: get(cells, iLat),
      longitude: get(cells, iLng),
      supplyrisk: get(cells, iSupplyRisk),
    });
  }
  return out;
}

/** Country / region centroids for SupplierCommitments `Location` (optional Lat/Lng columns override). */
const SUPPLIER_LOCATION_COORDS = {
  poland: [52.2297, 21.0122],
  germany: [52.3759, 9.732],
  philippines: [14.5995, 120.9842],
  phillipines: [14.5995, 120.9842],
  phillippines: [14.5995, 120.9842],
  czechrepublic: [50.0755, 14.4378],
  czechia: [50.0755, 14.4378],
  italy: [41.9028, 12.4964],
  hungary: [47.4979, 19.0402],
  turkey: [39.9334, 32.8597],
  spain: [40.4168, -3.7038],
  vietnam: [21.0285, 105.8542],
  france: [48.8566, 2.3522],
  netherlands: [52.3676, 4.9041],
  belgium: [50.8503, 4.3517],
  austria: [48.2082, 16.3738],
  slovakia: [48.1486, 17.1077],
  romania: [44.4268, 26.1025],
  uk: [51.5074, -0.1278],
  unitedkingdom: [51.5074, -0.1278],
  usa: [38.9072, -77.0369],
  china: [39.9042, 116.4074],
  india: [28.6139, 77.209],
  ukraine: [50.4501, 30.5234],
  mexico: [19.4326, -99.1332],
  brazil: [-15.7939, -47.8828],
};

const RISK_MARKER_COLORS = { high: "#ef4444", medium: "#f59e0b", low: "#22c55e" };

function coordsFromLocationString(locationRaw) {
  const k = normKey(locationRaw || "");
  if (!k) return null;
  if (SUPPLIER_LOCATION_COORDS[k]) return SUPPLIER_LOCATION_COORDS[k];
  for (const [key, ll] of Object.entries(SUPPLIER_LOCATION_COORDS)) {
    if (k.includes(key) || key.includes(k)) return ll;
  }
  return null;
}

function coordsForSupplierRecord(rec) {
  const lat = parseFloat(String(rec.latitude || "").replace(",", "."));
  const lng = parseFloat(String(rec.longitude || "").replace(",", "."));
  if (Number.isFinite(lat) && Number.isFinite(lng)) return [lat, lng];
  return coordsFromLocationString(rec.location);
}

/** `high` | `medium` | `low` — CSV `Supply Risk` overrides; else status-aligned mapping. */
function supplierRiskLevel(rec) {
  const liveDelay = supplierDelayShiftByName(rec.supplier || "");
  if (liveDelay >= 5) return "high";
  if (liveDelay >= 3) return "medium";
  if (liveDelay >= 1) return "medium";

  const ex = String(rec.supplyrisk || "")
    .trim()
    .toLowerCase();
  if (ex === "high" || ex === "critical") return "high";
  if (ex === "medium" || ex === "moderate" || ex === "watch") return "medium";
  if (ex === "low" || ex === "stable" || ex === "minimal") return "low";
  const s = (rec.status || "").toLowerCase();
  if (s.includes("at risk") || s.includes("critical") || s.includes("alert")) return "high";
  if (s.includes("slight delay") || s.includes("minor delay")) return "medium";
  if (s.includes("delay")) return "high";
  if (s.includes("in transit")) return "low";
  return "low";
}

function supplierRecordKey(rec) {
  return `${rec.supplier}|${rec.partid || rec.partname || ""}`;
}

function jitterCoordsForSlot(lat, lng, slot, total) {
  if (total <= 1) return [lat, lng];
  const angle = (2 * Math.PI * slot) / Math.max(total, 1);
  return [lat + 0.07 * Math.cos(angle), lng + 0.11 * Math.sin(angle)];
}

function buildLocationSlotMap(records) {
  const byLoc = new Map();
  records.forEach((rec, idx) => {
    const c = coordsForSupplierRecord(rec);
    if (!c) return;
    const key = `${c[0].toFixed(2)}_${c[1].toFixed(2)}`;
    if (!byLoc.has(key)) byLoc.set(key, []);
    byLoc.get(key).push(idx);
  });
  const slotByIndex = {};
  byLoc.forEach((indices) => {
    indices.forEach((recIdx, j) => {
      slotByIndex[recIdx] = { slot: j, total: indices.length };
    });
  });
  return slotByIndex;
}

let cachedSupplierRecords = [];
let shipmentSearchQuery = "";
let shipmentCardViewportResizeWired = false;

function filterRecordsByShipmentSearch(records, searchQuery) {
  const q = String(searchQuery || "")
    .trim()
    .toLowerCase();
  if (!q) return [...records];
  return records.filter((r) => (r.supplier || "").toLowerCase().includes(q));
}

function shipmentDotIcon(fill) {
  return L.divIcon({
    className: "shipment-marker-wrap",
    html: `<span style="display:block;width:15px;height:15px;border-radius:50%;background:${fill};border:2px solid #f8fafc;box-shadow:0 0 0 1px rgba(0,0,0,.4)"></span>`,
    iconSize: [19, 19],
    iconAnchor: [9.5, 9.5],
  });
}

let shipmentMarkersLayer = null;

function syncBottomPanelsHeight() {
  const commitments = document.querySelector("#panel-scc .card--commitments");
  const shipment = document.querySelector("#panel-scc .card--map");
  if (!commitments || !shipment) return;
  if (window.matchMedia("(max-width: 1100px)").matches) {
    commitments.style.height = "";
    return;
  }
  const target = shipment.getBoundingClientRect().height;
  if (target > 0) {
    commitments.style.height = `${Math.round(target)}px`;
  }
}

function syncShipmentTrackingVisualization(records) {
  cachedSupplierRecords = Array.isArray(records) ? records : [];
  const q = shipmentSearchQuery.trim();
  if (q && filterRecordsByShipmentSearch(cachedSupplierRecords, q).length === 0) {
    shipmentSearchQuery = "";
    const inp = document.getElementById("shipment-supplier-search");
    if (inp) inp.value = "";
    const st = document.getElementById("shipment-search-status");
    if (st) {
      st.hidden = true;
      st.textContent = "";
      st.classList.remove("shipment-search-status--error");
    }
  }
  renderShipmentLocationCards(cachedSupplierRecords, shipmentSearchQuery);
  renderSupplierMarkersOnMap(cachedSupplierRecords, shipmentSearchQuery);
  updateShipmentCardsViewport();
  requestAnimationFrame(() => syncBottomPanelsHeight());
}

function renderSupplierMarkersOnMap(records, searchQuery) {
  if (!shipmentLeafletMap || !shipmentMarkersLayer || typeof L === "undefined") return;

  shipmentMarkersLayer.clearLayers();
  const filtered = filterRecordsByShipmentSearch(records, searchQuery);
  const slotMap = buildLocationSlotMap(filtered);

  const bounds = [];
  filtered.forEach((rec, idx) => {
    let coords = coordsForSupplierRecord(rec);
    if (!coords) return;
    const meta = slotMap[idx];
    if (meta) {
      coords = jitterCoordsForSlot(coords[0], coords[1], meta.slot, meta.total);
    }
    const risk = supplierRiskLevel(rec);
    const color = RISK_MARKER_COLORS[risk];
    const key = supplierRecordKey(rec);
    const m = L.marker(coords, {
      icon: shipmentDotIcon(color),
      title: `${rec.supplier} (${rec.location})`,
    });
    m._supplierKey = key;
    m._supplierName = rec.supplier || "";
    const popupHtml = [
      `<div class="shipment-popup"><strong>${escapeHtml(rec.supplier)}</strong>`,
      `<div class="shipment-popup__muted">${escapeHtml(rec.location)} · ${escapeHtml(rec.partname || "")} (${escapeHtml(rec.partid || "")})</div>`,
      `<div>Status: <strong>${escapeHtml(rec.status || "—")}</strong> · OTIF ${escapeHtml(rec.otif || "—")}</div>`,
      `<div class="shipment-popup__risk shipment-popup__risk--${risk}">Risk: ${risk === "high" ? "High" : risk === "medium" ? "Watch" : "Stable"}</div>`,
      `</div>`,
    ].join("");
    m.bindPopup(popupHtml);
    m.addTo(shipmentMarkersLayer);
    bounds.push(coords);
  });

  if (bounds.length === 0) return;
  const narrowed = String(searchQuery || "").trim().length > 0;
  if (bounds.length === 1) {
    shipmentLeafletMap.setView(bounds[0], narrowed ? 6 : 5, { animate: false });
  } else {
    shipmentLeafletMap.fitBounds(L.latLngBounds(bounds), { padding: [52, 52], maxZoom: 6, animate: false });
  }
  scheduleShipmentLeafletResize();
}

function renderShipmentLocationCards(records, searchQuery) {
  const host = document.getElementById("shipment-location-cards");
  if (!host) return;

  if (records.length === 0) {
    host.innerHTML =
      '<p class="shipment-cards-empty">No supplier rows to show. Add data to SupplierCommitments.csv.</p>';
    return;
  }

  const list = filterRecordsByShipmentSearch(records, searchQuery);
  const orderRank = { high: 0, medium: 1, low: 2 };
  list.sort((a, b) => {
    const ra = supplierRiskLevel(a);
    const rb = supplierRiskLevel(b);
    if (orderRank[ra] !== orderRank[rb]) return orderRank[ra] - orderRank[rb];
    return (a.supplier || "").localeCompare(b.supplier || "", undefined, { sensitivity: "base" });
  });

  const qTrim = String(searchQuery || "").trim();
  const allActive = qTrim.length === 0;
  const allCard = `<button type="button" class="shipment-loc-card shipment-loc-card--all${allActive ? " is-active" : ""}" data-shipment-all="1" aria-pressed="${allActive}">
    <span class="shipment-loc-card__badge">Overview</span>
    <span class="shipment-loc-card__title">All suppliers</span>
    <span class="shipment-loc-card__meta">Show every pin and card below</span>
    <span class="shipment-loc-card__hint">Resets the map filter</span>
  </button>`;

  const cardsHtml = list
    .map((rec) => {
      const risk = supplierRiskLevel(rec);
      const coords = coordsForSupplierRecord(rec);
      const hasPin = coords != null;
      const label = risk === "high" ? "High risk" : risk === "medium" ? "Watch" : "Stable";
      const key = supplierRecordKey(rec);
      const qty = formatQty(rec.poqty ?? "");
      const unit = rec.unit ? escapeHtml(rec.unit) : "";
      const qtyLine = qty ? `${qty}${unit ? ` ${unit}` : ""}` : "—";
      return `<button type="button" class="shipment-loc-card shipment-loc-card--${risk}" data-supplier-key="${encodeURIComponent(key)}" ${
        hasPin ? `data-lat="${coords[0]}" data-lng="${coords[1]}"` : ""
      }>
        <span class="shipment-loc-card__risk">${label}</span>
        <span class="shipment-loc-card__title">${escapeHtml(rec.supplier)}</span>
        <span class="shipment-loc-card__loc">${escapeHtml(rec.location || "—")}</span>
        <span class="shipment-loc-card__meta">${escapeHtml(rec.partname || "")} · ${escapeHtml(rec.status || "—")}</span>
        <span class="shipment-loc-card__qty">${qtyLine}</span>
        ${
          hasPin
            ? '<span class="shipment-loc-card__hint">Click to focus on map</span>'
            : '<span class="shipment-loc-card__hint shipment-loc-card__hint--warn">Add Latitude/Longitude or a known Location for a map pin</span>'
        }
      </button>`;
    })
    .join("");

  host.innerHTML = allCard + cardsHtml;

  host.querySelectorAll(".shipment-loc-card").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (btn.dataset.shipmentAll === "1") {
        shipmentSearchQuery = "";
        const inp = document.getElementById("shipment-supplier-search");
        if (inp) inp.value = "";
        const st = document.getElementById("shipment-search-status");
        if (st) {
          st.hidden = true;
          st.textContent = "";
          st.classList.remove("shipment-search-status--error");
        }
        const mapPanel = document.getElementById("shipment-map-panel");
        if (mapPanel) {
          delete mapPanel.dataset.selectedSupplier;
          mapPanel.removeAttribute("data-supplier-label");
        }
        syncShipmentTrackingVisualization(cachedSupplierRecords);
        scheduleShipmentLeafletResize();
        return;
      }

      const lat = parseFloat(btn.getAttribute("data-lat") || "");
      const lng = parseFloat(btn.getAttribute("data-lng") || "");
      let skey = "";
      try {
        skey = decodeURIComponent(btn.getAttribute("data-supplier-key") || "");
      } catch {
        skey = btn.getAttribute("data-supplier-key") || "";
      }
      if (!shipmentLeafletMap || !shipmentMarkersLayer) return;
      let opened = false;
      shipmentMarkersLayer.eachLayer((layer) => {
        if (layer._supplierKey === skey) {
          const ll = layer.getLatLng();
          shipmentLeafletMap.setView(ll, Math.max(shipmentLeafletMap.getZoom(), 6), { animate: true });
          layer.openPopup();
          opened = true;
        }
      });
      if (!opened && Number.isFinite(lat) && Number.isFinite(lng)) {
        shipmentLeafletMap.setView([lat, lng], 6, { animate: true });
      }
    });
  });
}

const SHIPMENT_CARD_MIN_W = 156;
const SHIPMENT_CARD_MAX_W = 228;

function updateShipmentCardsViewport() {
  const outer = document.getElementById("shipment-cards-outer");
  const scroll = document.getElementById("shipment-location-cards");
  if (!outer || !scroll) return;
  const w = outer.clientWidth;
  if (w <= 0) return;
  const rootFs = parseFloat(getComputedStyle(document.documentElement).fontSize) || 15;
  const gap = 0.65 * rootFs;
  let cols = Math.floor((w + gap) / (SHIPMENT_CARD_MIN_W + gap));
  cols = Math.max(2, Math.min(10, cols));
  let cardW = (w - (cols - 1) * gap) / cols;
  cardW = Math.min(SHIPMENT_CARD_MAX_W, Math.max(SHIPMENT_CARD_MIN_W, cardW));
  const rounded = Math.round(cardW * 10) / 10;
  scroll.style.setProperty("--shipment-card-width", `${rounded}px`);
}

function focusFirstMatchingSupplierMarker() {
  if (!shipmentLeafletMap || !shipmentMarkersLayer) return;
  const q = shipmentSearchQuery.trim().toLowerCase();
  if (!q) return;
  let target = null;
  shipmentMarkersLayer.eachLayer((layer) => {
    if (target) return;
    const name = String(layer._supplierName || "").toLowerCase();
    if (name.includes(q)) target = layer;
  });
  if (!target) return;
  const ll = target.getLatLng();
  shipmentLeafletMap.setView(ll, Math.max(shipmentLeafletMap.getZoom(), 6), { animate: true });
  target.openPopup();
}

function runShipmentSupplierSearch() {
  const inp = document.getElementById("shipment-supplier-search");
  const st = document.getElementById("shipment-search-status");
  const raw = inp?.value?.trim() || "";
  if (!raw) {
    shipmentSearchQuery = "";
    if (st) {
      st.hidden = true;
      st.textContent = "";
      st.classList.remove("shipment-search-status--error");
    }
    const mapPanel = document.getElementById("shipment-map-panel");
    if (mapPanel) {
      delete mapPanel.dataset.selectedSupplier;
      mapPanel.removeAttribute("data-supplier-label");
    }
    syncShipmentTrackingVisualization(cachedSupplierRecords);
    scheduleShipmentLeafletResize();
    return;
  }
  const q = raw.toLowerCase();
  const hasMatch = cachedSupplierRecords.some((r) => (r.supplier || "").toLowerCase().includes(q));
  if (!hasMatch) {
    if (st) {
      st.hidden = false;
      st.textContent = `No supplier name contains “${escapeHtml(raw)}”.`;
      st.classList.add("shipment-search-status--error");
    }
    return;
  }
  shipmentSearchQuery = raw;
  if (st) {
    st.hidden = true;
    st.textContent = "";
    st.classList.remove("shipment-search-status--error");
  }
  const mapPanel = document.getElementById("shipment-map-panel");
  if (mapPanel) {
    mapPanel.dataset.selectedSupplier = raw;
    mapPanel.setAttribute("data-supplier-label", raw);
  }
  syncShipmentTrackingVisualization(cachedSupplierRecords);
  scheduleShipmentLeafletResize();
  focusFirstMatchingSupplierMarker();
}

function wireShipmentSupplierSearch() {
  const inp = document.getElementById("shipment-supplier-search");
  const btn = document.getElementById("shipment-search-btn");
  const outer = document.getElementById("shipment-cards-outer");
  if (btn && btn.dataset.wired !== "1") {
    btn.dataset.wired = "1";
    btn.addEventListener("click", () => runShipmentSupplierSearch());
  }
  if (inp && inp.dataset.wired !== "1") {
    inp.dataset.wired = "1";
    inp.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        runShipmentSupplierSearch();
      }
    });
  }
  if (outer && outer.dataset.roWired !== "1") {
    outer.dataset.roWired = "1";
    new ResizeObserver(() => updateShipmentCardsViewport()).observe(outer);
  }
  if (!shipmentCardViewportResizeWired) {
    shipmentCardViewportResizeWired = true;
    window.addEventListener(
      "resize",
      () => {
        updateShipmentCardsViewport();
        syncBottomPanelsHeight();
      },
      { passive: true }
    );
  }
  requestAnimationFrame(() => updateShipmentCardsViewport());
}

function eventIconClass(severity) {
  const s = String(severity || "warn")
    .trim()
    .toLowerCase();
  if (s === "danger" || s === "critical" || s === "high") return "event-icon event-icon--danger";
  if (s === "warn" || s === "warning") return "event-icon event-icon--warn";
  return "event-icon event-icon--info";
}

function eventIconGlyph(severity) {
  const s = String(severity || "warn")
    .trim()
    .toLowerCase();
  if (s === "danger" || s === "critical" || s === "high") return "!";
  if (s === "warn" || s === "warning") return "◎";
  return "i";
}

function delayPillClass(delayClass, potentialDelayText) {
  const d = String(delayClass || "")
    .trim()
    .toLowerCase();
  if (d === "high") return "delay-pill delay-pill--high";
  if (d === "low") return "delay-pill delay-pill--low";
  if (d === "mid" || d === "medium") return "delay-pill delay-pill--mid";
  const t = String(potentialDelayText || "").toLowerCase();
  if (/no\s*change|buffer|0\s*day/i.test(t)) return "delay-pill delay-pill--low";
  if (/\+[4-9]\s|\+[1-9]\d/.test(t)) return "delay-pill delay-pill--high";
  if (/\+0|no\s+delay|none/i.test(t)) return "delay-pill delay-pill--low";
  return "delay-pill delay-pill--mid";
}

function impactBadgeClass(priority) {
  const p = String(priority || "")
    .trim()
    .toLowerCase();
  if (p === "none" || p === "n/a" || p === "minimal" || p === "no change") return "badge badge--green";
  if (p === "high" || p === "critical") return "badge badge--red-sm";
  if (p === "low") return "badge badge--orange-sm";
  if (p === "mid" || p === "medium") return "badge badge--mid-sm";
  return "badge badge--orange-sm";
}

/** Derive Original / Predicted ETA from outlook "ETA Impact" text, e.g. "Apr 16 (instead of Apr 12)". */
function parseOutlookEtaImpact(raw) {
  const t = String(raw ?? "").trim();
  if (!t) return { originalEta: "", predictedEta: "" };
  if (/no\s*change/i.test(t)) {
    return { originalEta: "—", predictedEta: "No change" };
  }
  const m = t.match(/^(.+?)\s*\(\s*instead\s+of\s+(.+?)\s*\)\s*$/i);
  if (m) {
    return {
      predictedEta: m[1].trim(),
      originalEta: m[2].trim(),
    };
  }
  return { originalEta: "—", predictedEta: t };
}

/** Only High / Low outlook rows feed Alerts & Impact Summary. */
function normalizeOutlookImpactForAlerts(raw) {
  const p = String(raw || "")
    .trim()
    .toLowerCase();
  if (p === "high" || p === "critical") return "high";
  if (p === "low") return "low";
  return null;
}

/** Stacked dd for alert metrics: bold header line + muted / italic sub-lines. */
function formatAlertMetricDd(label, value) {
  const v = String(value ?? "").trim();
  if (!v) return "<dd class=\"metric-dd\">—</dd>";
  const lk = normKey(label);

  if (lk === "material" || lk === "materialpart") {
    const { sku, descriptor, qualifier } = parseOutlookMaterialParts(v);
    if (!descriptor && !qualifier) {
      return `<dd class="metric-dd metric-dd--stacked"><span class="metric-dd__head">${escapeHtml(sku)}</span></dd>`;
    }
    return `<dd class="metric-dd metric-dd--stacked">
      <span class="metric-dd__head">${escapeHtml(sku)}</span>
      ${descriptor ? `<span class="metric-dd__sub">${escapeHtml(descriptor)}</span>` : ""}
      ${qualifier ? `<span class="metric-dd__paren">${escapeHtml(qualifier)}</span>` : ""}
    </dd>`;
  }

  if (lk === "supplier" || lk === "impactedsupplier") {
    const m = v.match(/^(.+?)\s+\(([^)]+)\)\s*$/);
    if (m) {
      return `<dd class="metric-dd metric-dd--stacked">
        <span class="metric-dd__head">${escapeHtml(m[1].trim())}</span>
        <span class="metric-dd__sub">${escapeHtml(m[2].trim())}</span>
      </dd>`;
    }
  }

  if (lk === "originaleta" || lk === "predictedeta") {
    return `<dd class="metric-dd metric-dd--compact">${v === "—" ? "—" : escapeHtml(v)}</dd>`;
  }

  return `<dd class="metric-dd">${escapeHtml(v)}</dd>`;
}

function truncateAlertHeadline(s, max = 56) {
  const t = String(s).trim();
  if (t.length <= max) return t;
  return `${t.slice(0, max - 1)}…`;
}

/** Card summary title: delay + material focus (not event narrative). */
function alertCardImpactHeadline(delay, sku, descriptor) {
  const d = String(delay || "").trim();
  const matFocus = truncateAlertHeadline(String(descriptor || sku || "").trim() || "key material");
  if (d && matFocus) return `Delay of ${d} · ${matFocus}!`;
  if (d) return `Delay of ${d}!`;
  if (matFocus) return `Material impact: ${matFocus}!`;
  return "Supply impact alert!";
}

function alertCardImpactAnalysisHtml(level, delay, supplier, sku, descriptor, etaParts) {
  const { originalEta, predictedEta } = etaParts;
  const matE = escapeHtml(descriptor || sku || "—");
  const delayE = escapeHtml(String(delay || "").trim() || "as listed in outlook");
  const origE = escapeHtml(originalEta || "—");
  const predE = escapeHtml(predictedEta || "—");
  const supE = escapeHtml(String(supplier || "").trim() || "—");
  const tier = level === "high" ? "High" : "Low";

  const analysis = `${tier} impact concentration on ${matE}. Reported delay ${delayE}; original ETA ${origE}, predicted ${predE}. Primary supplier lane: ${supE}.`;

  const workaround =
    level === "high"
      ? "Workaround: Secure alternate lots or approved substitutes, rebalance final assembly sequence, lock feeder dock slots, and brief program management within the current shift window."
      : "Workaround: Continue planned buffer consumption, confirm the next two inbound milestones, keep expedite budget staged if slip widens beyond one day, and maintain daily tier-2 visibility until the lane clears.";

  return `<div class="what-means what-means--impact">
    <h3>Impact Analysis and Workaround recommendations</h3>
    <p class="what-means__analysis">${analysis}</p>
    <p class="what-means__workaround">${escapeHtml(workaround)}</p>
  </div>`;
}

function outlookAlertCardHtml(headers, cells, level, impactRaw, startOpen) {
  const openAttr = startOpen ? " open" : "";
  const cardClass =
    level === "high" ? "impact-card impact-card--impact-high" : "impact-card impact-card--impact-low";
  const badgeClass = impactBadgeClass(impactRaw);
  const badgeText = escapeHtml(String(impactRaw || (level === "high" ? "High" : "Low")).trim());

  const supplier = outlookCellByAliases(cells, headers, "Impacted Supplier", "Supplier");
  const material = outlookCellByAliases(cells, headers, "Material / Part", "Material/Part", "Material", "Part");
  const delay = outlookCellByAliases(cells, headers, "Delay");
  const etaRaw = outlookCellByAliases(cells, headers, "ETA Impact", "ETA");
  const etaParts = parseOutlookEtaImpact(etaRaw);

  const { sku, descriptor } = parseOutlookMaterialParts(material);
  const impactTitle = alertCardImpactHeadline(delay, sku, descriptor);
  const hintParts = [];
  if (supplier) hintParts.push(supplier);
  if (delay) hintParts.push(delay);
  const hint = hintParts.join(" · ");

  const originalEta = etaParts.originalEta || "—";
  const predictedEta = etaParts.predictedEta || "—";

  const metrics = [
    ["Material", material || "—"],
    ["Supplier", supplier || "—"],
    ["Original ETA", originalEta],
    ["Predicted ETA", predictedEta],
    ["Delay", delay || "—"],
  ];

  const metricsHtml = metrics
    .map(([dt, dd]) => `<div class="metric-cell"><dt>${escapeHtml(dt)}</dt>${formatAlertMetricDd(dt, dd)}</div>`)
    .join("");

  const narrative = alertCardImpactAnalysisHtml(level, delay, supplier, sku, descriptor, etaParts);

  return `<details class="${cardClass}"${openAttr}>
  <summary class="impact-card__summary">
    <span class="impact-card__summary-text">
      <span class="${badgeClass} impact-card__impact-badge">${badgeText}</span>
      <span class="impact-card__title">${escapeHtml(impactTitle)}</span>
      ${hint ? `<span class="impact-card__hint">${escapeHtml(hint)}</span>` : ""}
    </span>
    <span class="impact-card__chevron" aria-hidden="true">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M6 9l6 6 6-6" />
      </svg>
    </span>
  </summary>
  <div class="impact-card__body">
    <dl class="metric-list metric-list--impact-grid">${metricsHtml}</dl>
    ${narrative}
  </div>
</details>`;
}

function renderOutlookImpactAlerts(headers, bodyRows) {
  const host = document.getElementById("alert-impact-cards");
  if (!host) return;

  if (!headers || headers.length === 0) {
    host.innerHTML =
      '<p class="alert-cards-empty">Outlook data unavailable. Refresh Global Supply Outlook after fixing the CSV or connection.</p>';
    return;
  }

  if (!Array.isArray(bodyRows) || bodyRows.length === 0) {
    host.innerHTML =
      '<p class="alert-cards-empty">No outlook rows loaded. High and low impact alerts appear when the CSV includes matching Impact values.</p>';
    return;
  }

  const tagged = [];
  bodyRows.forEach((cells, idx) => {
    const impactRaw = outlookCellByAliases(cells, headers, "Impact", "ImpactPriority", "Priority");
    const level = normalizeOutlookImpactForAlerts(impactRaw);
    if (!level) return;
    tagged.push({ cells, level, impactRaw, idx });
  });

  tagged.sort((a, b) => {
    const order = { high: 0, low: 1 };
    if (order[a.level] !== order[b.level]) return order[a.level] - order[b.level];
    return a.idx - b.idx;
  });

  if (tagged.length === 0) {
    host.innerHTML =
      '<p class="alert-cards-empty">No high or low impact rows in Global Supply Outlook. None and other levels appear only in the outlook table.</p>';
    return;
  }

  host.innerHTML = tagged
    .map((row, i) => outlookAlertCardHtml(headers, row.cells, row.level, row.impactRaw, i === 0))
    .join("");
}

function renderOutlookAlertsIdle(message) {
  const host = document.getElementById("alert-impact-cards");
  if (!host) return;
  host.innerHTML = `<p class="alert-cards-empty">${escapeHtml(message)}</p>`;
}

function outlookCellByAliases(cells, headers, ...aliases) {
  const norms = headers.map(normKey);
  for (const a of aliases) {
    const n = normKey(a);
    const i = norms.indexOf(n);
    if (i >= 0 && cells[i] != null) return String(cells[i]).trim();
  }
  return "";
}

/** Map Impact column to icon severity when Severity column is absent. */
function outlookSeverityFromImpact(impact) {
  const p = String(impact || "")
    .trim()
    .toLowerCase();
  if (p === "high" || p === "critical") return "danger";
  if (p === "low") return "warn";
  if (p === "none" || p === "n/a" || p === "") return "info";
  return "warn";
}

function padOutlookRow(cells, len) {
  const out = (cells || []).map((c) => (c != null ? String(c) : ""));
  while (out.length < len) out.push("");
  return out.slice(0, len);
}

function parseOutlookTable(parsedRows) {
  if (parsedRows.length < 1) return { headers: [], body: [] };
  const headers = parsedRows[0].map((h) => String(h).trim());
  const norms = headers.map(normKey);
  if (!norms.includes("event")) {
    throw new Error("GlobalSupplyOutlook.csv must include an Event column.");
  }
  const body = [];
  for (let r = 1; r < parsedRows.length; r++) {
    const cells = parsedRows[r];
    if (!cells || cells.every((c) => !String(c ?? "").trim())) continue;
    body.push(padOutlookRow(cells, headers.length));
  }
  return { headers, body };
}

function isSequenceProductLetter(v) {
  const s = String(v || "")
    .trim()
    .toUpperCase();
  return s === "B" || s === "K" || s === "W";
}

function countSequenceProductionSlots(schedule) {
  let n = 0;
  for (const row of schedule || []) {
    for (const slot of row.slots || []) {
      if (isSequenceProductLetter(slot)) n++;
    }
  }
  return n;
}

function firstColIndexByAliases(headers, aliases) {
  const norms = headers.map((h) => normKey(h));
  for (const a of aliases) {
    const want = normKey(a);
    const idx = norms.findIndex((n) => n === want);
    if (idx >= 0) return idx;
  }
  return -1;
}

function randomIntInclusive(min, max) {
  return min + Math.floor(Math.random() * (max - min + 1));
}

function impactFromDelayShifts(shifts) {
  if (shifts >= 5) return "High";
  if (shifts >= 3) return "Medium";
  return "Low";
}

function locationKeyCandidatesFromText(text) {
  const nk = normKey(text || "");
  if (!nk) return [];
  const candidates = [];
  for (const key of Object.keys(SUPPLIER_LOCATION_COORDS)) {
    if (nk.includes(key)) candidates.push(key);
  }
  return [...new Set(candidates)];
}

function pickEventCoordsFromText(eventText) {
  const keys = locationKeyCandidatesFromText(eventText);
  if (!keys.length) return null;
  // Prefer the last mentioned location token in the sentence.
  const key = keys[keys.length - 1];
  return SUPPLIER_LOCATION_COORDS[key] || null;
}

function distanceKm(a, b) {
  const toRad = (x) => (x * Math.PI) / 180;
  const [lat1, lon1] = a;
  const [lat2, lon2] = b;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const s1 = Math.sin(dLat / 2);
  const s2 = Math.sin(dLon / 2);
  const h =
    s1 * s1 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * s2 * s2;
  return 6371 * 2 * Math.atan2(Math.sqrt(h), Math.sqrt(1 - h));
}

function nearestSupplierRecordsByEvent(eventText, limit = 2) {
  const valid = cachedSupplierRecords
    .map((r) => ({ rec: r, coords: coordsForSupplierRecord(r) }))
    .filter((x) => Array.isArray(x.coords) && x.coords.length === 2);
  if (!valid.length) return [];

  const eventCoords = pickEventCoordsFromText(eventText);
  if (!eventCoords) {
    // No detectable location token: fallback to first suppliers in current commitments.
    return valid.slice(0, Math.max(1, limit)).map((x) => x.rec);
  }

  valid.sort((a, b) => distanceKm(eventCoords, a.coords) - distanceKm(eventCoords, b.coords));
  const picked = [];
  const seenSupplier = new Set();
  for (const item of valid) {
    const name = String(item.rec.supplier || "").trim();
    if (!name || seenSupplier.has(name)) continue;
    seenSupplier.add(name);
    picked.push(item.rec);
    if (picked.length >= Math.max(1, limit)) break;
  }
  return picked;
}

function generatedOutlookRowFromEvent(headers, eventText) {
  const row = Array.from({ length: headers.length }, () => "");
  const iEvent = firstColIndexByAliases(headers, ["Event"]);
  const iSupplier = firstColIndexByAliases(headers, ["Impacted Supplier", "Supplier"]);
  const iMaterial = firstColIndexByAliases(headers, ["Material / Part", "Material", "Part / Material"]);
  const iDelay = firstColIndexByAliases(headers, ["Delay", "Potential Delay"]);
  const iImpact = firstColIndexByAliases(headers, ["Impact", "Impact Priority", "Priority"]);

  const nearest = nearestSupplierRecordsByEvent(eventText, 1);
  const suppliers = nearest
    .map((r) => String(r.supplier || "").trim())
    .filter(Boolean);
  const materials = nearest
    .map((r) => String(r.partname || "").trim())
    .filter(Boolean);

  const delayShifts = randomIntInclusive(1, 6);
  const impact = impactFromDelayShifts(delayShifts);

  if (iEvent >= 0) row[iEvent] = String(eventText || "").trim();
  if (iSupplier >= 0) row[iSupplier] = suppliers.length ? suppliers.join(", ") : "Unknown supplier";
  if (iMaterial >= 0) row[iMaterial] = materials.length ? [...new Set(materials)].join(", ") : "Unknown material";
  if (iDelay >= 0) row[iDelay] = `${delayShifts} Shift${delayShifts > 1 ? "s" : ""}`;
  if (iImpact >= 0) row[iImpact] = impact;

  return row;
}

function parseDelayShiftCount(raw) {
  const m = String(raw || "").match(/(\d+)/);
  if (!m) return 0;
  const n = Number(m[1]);
  return Number.isFinite(n) ? Math.max(0, n) : 0;
}

function splitSupplierNames(raw) {
  return String(raw || "")
    .split(/,|&| and /i)
    .map((s) => s.trim())
    .filter(Boolean);
}

/** Map inventory baseline cell (B / R / K / W) to production sequence slot letters (B / K / W). */
function sequenceLetterFromInventoryBaselineCell(rawCell) {
  const s = String(rawCell ?? "")
    .trim()
    .toUpperCase();
  if (s === "B") return "B";
  if (s === "R" || s === "K") return "K";
  if (s === "W") return "W";
  return "";
}

function supplierSeatLetterByName(name) {
  const key = normKey(name || "");
  if (!key) return "";
  const rec = cachedSupplierRecords.find((r) => normKey(r.supplier || "") === key);
  if (rec) {
    const fromSeat = seatTokenLetterFromCell(rec.seattype || "");
    if (isSequenceProductLetter(fromSeat)) return fromSeat;
  }
  const ci = supplierColumnIndexFromName(name);
  if (ci >= 0 && inventoryMatrixBaseline.length > ci) {
    const fromBaseline = sequenceLetterFromInventoryBaselineCell(inventoryMatrixBaseline[ci]);
    if (fromBaseline) return fromBaseline;
  }
  return "";
}

function outlookSupplierDelayImpacts() {
  if (!cachedOutlookHeaders.length || !cachedOutlookRows.length) return [];
  const iSupplier = firstColIndexByAliases(cachedOutlookHeaders, ["Impacted Supplier", "Supplier"]);
  const iDelay = firstColIndexByAliases(cachedOutlookHeaders, ["Delay", "Potential Delay"]);
  if (iSupplier < 0 || iDelay < 0) return [];
  const bySupplier = new Map();
  for (const row of cachedOutlookRows) {
    const delay = parseDelayShiftCount(row[iDelay]);
    if (delay <= 0) continue;
    const suppliers = splitSupplierNames(row[iSupplier]);
    suppliers.forEach((s) => {
      const name = String(s || "").trim();
      if (!name) return;
      const key = normKey(name);
      const prev = bySupplier.get(key);
      if (!prev || delay > prev.delayShifts) {
        bySupplier.set(key, { supplier: name, delayShifts: delay });
      }
    });
  }
  return [...bySupplier.values()];
}

function supplierDelayShiftByName(name) {
  const key = normKey(name || "");
  if (!key) return 0;
  const aliasKeys = new Set([key]);
  const m = String(name || "").match(/A\s*(\d+)/i);
  if (m) aliasKeys.add(`a${m[1]}`);

  const impacts = outlookSupplierDelayImpacts();
  let best = 0;
  for (const imp of impacts) {
    const supplierRaw = String(imp.supplier || "").trim();
    if (!supplierRaw) continue;
    const impKey = normKey(supplierRaw);
    const impAlias = new Set([impKey]);
    const mm = supplierRaw.match(/A\s*(\d+)/i);
    if (mm) impAlias.add(`a${mm[1]}`);
    let matched = false;
    for (const a of aliasKeys) {
      if (impAlias.has(a)) {
        matched = true;
        break;
      }
      for (const b of impAlias) {
        if (a.includes(b) || b.includes(a)) {
          matched = true;
          break;
        }
      }
      if (matched) break;
    }
    if (!matched) continue;
    const d = Math.max(0, Number(imp.delayShifts) || 0);
    if (d > best) best = d;
  }
  return best;
}

function supplierColumnIndexFromName(name) {
  const raw = String(name || "").trim();
  if (!raw || !inventoryMatrixHeaders.length) return -1;
  const maxCol = Math.min(9, inventoryMatrixHeaders.length);
  const m = raw.match(/A\s*(\d+)/i);
  if (m) {
    const n = Number(m[1]);
    if (Number.isFinite(n) && n >= 1 && n <= inventoryMatrixHeaders.length) return n - 1;
  }
  // CSV often uses "Supplier 1" … "Supplier 9" (same as Supplier A1…A9 in the matrix).
  const mSn = raw.match(/\bsupplier\s*(\d+)\b/i);
  if (mSn) {
    const n = Number(mSn[1]);
    if (Number.isFinite(n) && n >= 1 && n <= maxCol) return n - 1;
  }
  const mBare = raw.match(/^\s*(\d+)\s*$/);
  if (mBare) {
    const n = Number(mBare[1]);
    if (Number.isFinite(n) && n >= 1 && n <= maxCol) return n - 1;
  }
  const nk = normKey(raw);
  for (let i = 0; i < inventoryMatrixHeaders.length; i++) {
    const hk = normKey(inventoryMatrixHeaders[i]);
    if (hk && (hk.includes(nk) || nk.includes(hk))) return i;
  }
  return -1;
}

function blockedInventoryColumnsByShift(totalShifts) {
  if (inventoryReplenishedAfterPartnerResequence) {
    return Array.from({ length: Math.max(0, totalShifts) }, () => new Set());
  }
  const out = Array.from({ length: Math.max(0, totalShifts) }, () => new Set());
  const impacts = outlookSupplierDelayImpacts();
  for (const imp of impacts) {
    const idx = supplierColumnIndexFromName(imp.supplier);
    if (idx < 0) continue;
    const lim = Math.min(out.length, Math.max(0, imp.delayShifts));
    for (let s = 0; s < lim; s++) out[s].add(idx);
  }
  return out;
}

function currentSequenceShiftIndex() {
  if (!currentSequenceSchedule.length || sequenceCellTotal <= 0) return 0;
  let remaining = sequenceCycleTicks;
  for (let r = 0; r < currentSequenceSchedule.length; r++) {
    const cap = (currentSequenceSchedule[r].slots || []).filter((s) => isSequenceProductLetter(s)).length;
    const rowCap = Math.max(1, cap);
    if (remaining < rowCap) return r;
    remaining -= rowCap;
  }
  return Math.max(0, currentSequenceSchedule.length - 1);
}

function impactedSupplierColumnSet(currentShiftIdx = 0) {
  if (inventoryReplenishedAfterPartnerResequence) return new Set();
  const out = new Set();
  const impacts = outlookSupplierDelayImpacts();
  for (const imp of impacts) {
    if ((Number(imp.delayShifts) || 0) <= 0) continue;
    if (Number(imp.delayShifts) <= currentShiftIdx) continue;
    const idx = supplierColumnIndexFromName(imp.supplier);
    if (idx >= 0) out.add(idx);
  }
  return out;
}

function activeDelayedSupplierCount(currentShiftIdx = 0) {
  if (inventoryReplenishedAfterPartnerResequence) return 0;
  const impacts = outlookSupplierDelayImpacts();
  const active = new Set();
  for (const imp of impacts) {
    const delay = Number(imp.delayShifts) || 0;
    if (delay <= currentShiftIdx) continue;
    const key = normKey(imp.supplier || "");
    if (key) active.add(key);
  }
  return active.size;
}

function applyOutlookImpactsToSequence(baseSchedule) {
  const schedule = (baseSchedule || []).map((r) => ({ ...r, slots: [...(r.slots || [])] }));
  const impacts = outlookSupplierDelayImpacts();
  if (!impacts.length) return schedule;

  // Supplier column A1…A9 maps to OEM slots by shared dice color (e.g. A1→slot1, A2→slot4, A3→slot7 for blues).
  impacts.forEach((imp) => {
    const lim = Math.min(schedule.length, Math.max(0, Number(imp.delayShifts) || 0));
    if (lim <= 0) return;
    const ci = supplierColumnIndexFromName(imp.supplier);
    if (ci >= 0 && ci < 9) {
      for (let r = 0; r < lim; r++) {
        const slots = schedule[r].slots;
        if (!Array.isArray(slots)) continue;
        while (slots.length < 9) slots.push("");
        const baseRow = (baseSchedule[r] && baseSchedule[r].slots) || [];
        const si = sequenceSlotIndexForSupplierColumnInRow(baseRow, ci);
        if (si < slots.length) slots[si] = "";
      }
    }
  });

  // rowImpact[rowIdx][letter] = count of unavailable suppliers for that color in this shift
  // (used when no A1…A9 column mapping; avoids double-removing impacts already cleared by column).
  const rowImpact = Array.from({ length: schedule.length }, () => ({ B: 0, K: 0, W: 0 }));
  impacts.forEach((imp) => {
    const ci = supplierColumnIndexFromName(imp.supplier);
    if (ci >= 0 && ci < 9) return;
    const letter = supplierSeatLetterByName(imp.supplier);
    if (!isSequenceProductLetter(letter)) return;
    const lim = Math.min(schedule.length, Math.max(0, Number(imp.delayShifts) || 0));
    for (let r = 0; r < lim; r++) rowImpact[r][letter] += 1;
  });

  for (let r = 0; r < schedule.length; r++) {
    const slots = schedule[r].slots;
    if (!Array.isArray(slots)) continue;
    for (const letter of ["B", "K", "W"]) {
      let remaining = rowImpact[r][letter] || 0;
      while (remaining > 0) {
        // remove from rightmost to mimic higher-number supplier (e.g., A9 white) dropping first
        let found = -1;
        for (let i = slots.length - 1; i >= 0; i--) {
          if (String(slots[i] || "").trim().toUpperCase() === letter) {
            found = i;
            break;
          }
        }
        if (found < 0) break;
        slots[found] = "";
        remaining--;
      }
    }

    // Gaps stay in column position (no shift-left); Re-sequence Production applies the partner tail order.
    while (slots.length < 9) slots.push("");
    schedule[r].slots = slots.length > 9 ? slots.slice(0, 9) : slots;
  }

  return schedule;
}

function delayShiftsForSupplierColumn(ci) {
  let max = 0;
  for (const imp of outlookSupplierDelayImpacts()) {
    if (supplierColumnIndexFromName(imp.supplier) === ci) {
      max = Math.max(max, Number(imp.delayShifts) || 0);
    }
  }
  return max;
}

function isSupplierColumnDelayedForRow(rowIdx, ci) {
  return rowIdx < delayShiftsForSupplierColumn(ci);
}

/** Same leather-color supplier columns as `ci` (nearest index first), excluding `ci`. */
function partnerSupplierColumnIndicesSameColor(ci) {
  const invSeq = sequenceLetterFromInventoryBaselineCell(inventoryMatrixBaseline[ci]);
  if (!isSequenceProductLetter(invSeq)) return [];
  const idxs = [];
  for (let j = 0; j < Math.min(9, inventoryMatrixBaseline.length); j++) {
    if (j === ci) continue;
    if (sequenceLetterFromInventoryBaselineCell(inventoryMatrixBaseline[j]) === invSeq) idxs.push(j);
  }
  idxs.sort((a, b) => Math.abs(a - ci) - Math.abs(b - ci));
  return idxs;
}

function firstSuggestedPartnerHeaderLabel(impactedCi) {
  if (impactedCi < 0 || !inventoryMatrixHeaders.length) return "";
  for (const p of partnerSupplierColumnIndicesSameColor(impactedCi)) {
    if (delayShiftsForSupplierColumn(p) > 0) continue;
    const h = String(inventoryMatrixHeaders[p] ?? "").trim();
    if (h) return h;
  }
  return "";
}

function buildSupplyInvDisclaimerPromptText() {
  const iSupplier = firstColIndexByAliases(cachedOutlookHeaders, ["Impacted Supplier", "Supplier"]);
  if (iSupplier < 0 || !cachedOutlookRows.length) return "";
  const lastRow = cachedOutlookRows[cachedOutlookRows.length - 1];
  const primary = splitSupplierNames(String(lastRow[iSupplier] ?? ""))[0] || "Supplier";
  const ci = supplierColumnIndexFromName(primary);
  const partnerLabel = ci >= 0 ? firstSuggestedPartnerHeaderLabel(ci) : "";
  const replen = partnerLabel || "a nearby facility";
  return `${primary} has been impacted. System suggests replenishment from ${replen} and resequence the production to accommodate new lead times.`;
}

function renderSupplyInvDisclaimer() {
  const el = document.getElementById("supply-inv-disclaimer-text");
  if (!el) return;
  if (inventoryReplenishedAfterPartnerResequence) {
    el.textContent = "Disclaimer: Inventory is replenished every shift.";
    return;
  }
  if (hasAnalyzedOutlook && cachedOutlookRows.length) {
    const prompt = buildSupplyInvDisclaimerPromptText();
    if (prompt) {
      el.textContent = prompt;
      return;
    }
  }
  el.textContent = "Disclaimer: Inventory is replenished every shift.";
}

/**
 * Maps inventory supplier column (A1…A9) to OEM sequence slot index (0…8) for this row.
 * Same leather-color suppliers share one dice color in the row (e.g. 1st/2nd/3rd blue → slots 1,4,7).
 */
function sequenceSlotIndexForSupplierColumnInRow(baseRow, ci) {
  if (ci < 0 || ci >= 9) return ci;
  const invSeq = sequenceLetterFromInventoryBaselineCell(inventoryMatrixBaseline[ci]);
  if (!isSequenceProductLetter(invSeq)) return ci;
  const sameCols = [];
  for (let j = 0; j < Math.min(9, inventoryMatrixBaseline.length); j++) {
    if (sequenceLetterFromInventoryBaselineCell(inventoryMatrixBaseline[j]) === invSeq) sameCols.push(j);
  }
  sameCols.sort((a, b) => a - b);
  const rank = sameCols.indexOf(ci);
  if (rank < 0) return ci;
  const br = (baseRow || []).map((s) =>
    String(s ?? "")
      .trim()
      .toUpperCase()
  );
  const positions = [];
  for (let p = 0; p < 9; p++) {
    if (String(br[p] || "") === invSeq) positions.push(p);
  }
  positions.sort((a, b) => a - b);
  if (rank < positions.length) return positions[rank];
  return ci;
}

/**
 * Partner re-sequence: dice for delayed suppliers (by slot 1/4/7 style mapping) move to the **tail**;
 * remaining slots stay in order (shift left). Multiset stays 3×B / 3×K / 3×W.
 */
function buildPartnerResequenceRowForDelays(baseSlots, rowIdx, delayedColsSorted) {
  const base = baseSlots.map((s) =>
    String(s ?? "")
      .trim()
      .toUpperCase()
  );
  while (base.length < 9) base.push("");
  const active = delayedColsSorted.filter((ci) => isSupplierColumnDelayedForRow(rowIdx, ci));
  if (!active.length) return base.slice();
  const deferredSlotIdx = active.map((ci) => sequenceSlotIndexForSupplierColumnInRow(base, ci));
  const slotSet = new Set(deferredSlotIdx);
  const middle = [];
  for (let j = 0; j < 9; j++) {
    if (!slotSet.has(j)) middle.push(base[j]);
  }
  const tail = active.map((ci) => base[sequenceSlotIndexForSupplierColumnInRow(base, ci)]);
  return [...middle, ...tail];
}

/**
 * Full 9 slots: permutation of baseline — deferred supplier tokens grouped at columns 8–9 (last slots).
 */
function applyPartnerResequenceToSequence(baseSchedule) {
  const schedule = (baseSchedule || []).map((r) => ({ ...r, slots: [...(r.slots || [])] }));
  const delayedCols = [];
  const seen = new Set();
  for (const imp of outlookSupplierDelayImpacts()) {
    const ci = supplierColumnIndexFromName(imp.supplier);
    if (ci >= 0 && ci < 9 && !seen.has(ci)) {
      seen.add(ci);
      delayedCols.push(ci);
    }
  }
  delayedCols.sort((a, b) => a - b);
  if (!delayedCols.length) return schedule;

  for (let r = 0; r < schedule.length; r++) {
    const baseSlots = (baseSchedule[r].slots || []).map((s) =>
      String(s ?? "")
        .trim()
        .toUpperCase()
    );
    while (baseSlots.length < 9) baseSlots.push("");
    schedule[r].slots = buildPartnerResequenceRowForDelays(baseSlots, r, delayedCols);
  }
  return schedule;
}

/** Split long Event strings on en/em dash (CSV may use mojibake U+FFFD). */
function parseOutlookEventParts(raw) {
  const t = String(raw ?? "")
    .replace(/\uFFFD/g, "–")
    .trim();
  if (!t) return { headline: "—", paren: "", body: "" };
  const parts = t
    .split(/\s*[–—]\s*/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (parts.length >= 3) {
    return {
      headline: parts[0],
      paren: parts[1],
      body: parts.slice(2).join(" – "),
    };
  }
  if (parts.length === 2) {
    return { headline: parts[0], paren: "", body: parts[1] };
  }
  return { headline: t, paren: "", body: "" };
}

function outlookEventColumnTh(rawHeader) {
  const full = String(rawHeader || "Event").trim() || "Event";
  return `<th scope="col" class="events-th events-th--event-stack" title="${escapeHtml(full)}">Event</th>`;
}

function outlookMaterialColumnTh(rawHeader) {
  const full = String(rawHeader || "Material / Part").trim() || "Material / Part";
  return `<th scope="col" class="events-th events-th--material-stack" title="${escapeHtml(full)}">Material/Part</th>`;
}

/** Split Material/Part: optional trailing (category), middle dot or P-xxx prefix. */
function parseOutlookMaterialParts(raw) {
  let t = String(raw ?? "")
    .replace(/\uFFFD/g, "·")
    .trim();
  if (!t) return { sku: "—", descriptor: "", qualifier: "" };

  let qualifier = "";
  const pm = t.match(/\(([^)]+)\)\s*$/);
  if (pm) {
    qualifier = pm[1].trim();
    t = t.slice(0, pm.index).trim();
  }

  const segs = t
    .split(/\s*[·•]\s*/)
    .map((s) => s.trim())
    .filter(Boolean);
  if (segs.length >= 2) {
    return { sku: segs[0], descriptor: segs.slice(1).join(" · "), qualifier };
  }

  const m = t.match(/^(P-\d+[A-Za-z-]*)\s+(.+)$/i);
  if (m) return { sku: m[1], descriptor: m[2], qualifier };

  return { sku: t, descriptor: "", qualifier };
}

function outlookTheadRowHtml(headers) {
  return `<tr>${headers
    .map((h) => {
      const nk = normKey(h);
      if (nk === "event") return outlookEventColumnTh(h);
      if (nk === "materialpart" || nk === "material") return outlookMaterialColumnTh(h);
      return `<th scope="col">${escapeHtml(h)}</th>`;
    })
    .join("")}</tr>`;
}

function outlookRowHtml(headers, cells) {
  const delayClass = outlookCellByAliases(cells, headers, "DelayClass", "DelayLevel");
  const explicitSeverity = outlookCellByAliases(cells, headers, "Severity", "EventSeverity", "Level");
  const impactVal = outlookCellByAliases(cells, headers, "Impact", "ImpactPriority", "Priority");
  const severity = explicitSeverity || outlookSeverityFromImpact(impactVal);

  const tds = headers.map((rawH, idx) => {
    const nk = normKey(rawH);
    const val = cells[idx] != null ? String(cells[idx]).trim() : "";
    const escaped = escapeHtml(val);

    if (nk === "event") {
      const iconClass = eventIconClass(severity);
      const glyph = eventIconGlyph(severity);
      const { headline, paren, body } = parseOutlookEventParts(val);
      const parenLine = paren
        ? `<span class="event-cell__paren">(${escapeHtml(paren)})</span>`
        : "";
      const bodyBlock = body ? `<p class="event-cell__detail">${escapeHtml(body)}</p>` : "";
      return `<td class="events-td-event">
        <div class="event-cell event-cell--stacked event-cell--wide">
          <span class="${iconClass}" aria-hidden="true">${glyph}</span>
          <div class="event-cell__stack">
            <span class="event-cell__headline">${escapeHtml(headline)}</span>
            ${parenLine}
            ${bodyBlock}
          </div>
        </div>
      </td>`;
    }
    if (nk === "materialpart" || nk === "material") {
      const { sku, descriptor, qualifier } = parseOutlookMaterialParts(val);
      const parenLine = qualifier
        ? `<span class="event-cell__paren">(${escapeHtml(qualifier)})</span>`
        : "";
      const detailBlock = descriptor
        ? `<p class="event-cell__detail">${escapeHtml(descriptor)}</p>`
        : "";
      return `<td class="events-td-material">
        <div class="event-cell__stack">
          <span class="event-cell__headline">${escapeHtml(sku)}</span>
          ${parenLine}
          ${detailBlock}
        </div>
      </td>`;
    }
    if (nk === "delay" || nk === "potentialdelay") {
      const pill = delayPillClass(delayClass, val);
      return `<td><span class="${pill}">${escaped}</span></td>`;
    }
    if (nk === "impact" || nk === "impactpriority" || nk === "priority") {
      return `<td><span class="${impactBadgeClass(val)}">${escaped}</span></td>`;
    }
    return `<td class="events-td-wrap">${escaped}</td>`;
  });

  return `<tr>${tds.join("")}</tr>`;
}

async function loadSupplierCommitments() {
  const tbody = document.getElementById("supplier-commitments-tbody");
  const statusEl = document.getElementById("supplier-commitments-status");
  if (!tbody) return;

  const setStatus = (msg, isError) => {
    if (!statusEl) return;
    statusEl.hidden = !msg;
    statusEl.textContent = msg;
    statusEl.classList.toggle("table-status--error", Boolean(isError));
  };

  setStatus("Loading supplier data…", false);

  try {
    const text = await fetchSupplierCsvText();
    const rows = parseCSV(text);
    const records = recordsFromRows(rows);
    if (records.length === 0) {
      tbody.innerHTML = "";
      syncShipmentTrackingVisualization([]);
      setStatus("No data rows found in the supplier CSV.", true);
      return;
    }
    tbody.innerHTML = records.map(rowFromRecord).join("");
    syncShipmentTrackingVisualization(records);
    setStatus("", false);
  } catch (e) {
    tbody.innerHTML = "";
    syncShipmentTrackingVisualization([]);
    const hint =
      window.location.protocol === "file:"
        ? " Browsers block file:// fetches; use python -m http.server in this folder, then open http://localhost:8080/index.html"
        : "";
    setStatus((e instanceof Error ? e.message : "Failed to load CSV.") + hint, true);
  }
}

let shipmentLeafletMap = null;
let shipmentMapResizeTimer = null;

function resizeShipmentLeafletMap() {
  shipmentLeafletMap?.invalidateSize({ animate: false });
}

function scheduleShipmentLeafletResize() {
  window.clearTimeout(shipmentMapResizeTimer);
  shipmentMapResizeTimer = window.setTimeout(() => resizeShipmentLeafletMap(), 160);
}

/**
 * Free map: Leaflet + Carto Dark Matter tiles (OpenStreetMap data). No API key.
 * @see https://leafletjs.com/
 */
function initShipmentLeafletMap() {
  const panel = document.getElementById("shipment-map-panel");
  const el = document.getElementById("shipment-leaflet-map");
  const hint = document.getElementById("shipment-maps-hint");
  if (!panel || !el) return;

  if (typeof window.L === "undefined") {
    if (hint) {
      hint.hidden = false;
      hint.textContent =
        "Interactive map could not load (Leaflet missing). Use http://localhost and allow the CDN, or check your network.";
    }
    return;
  }

  if (shipmentLeafletMap) {
    scheduleShipmentLeafletResize();
    return;
  }

  try {
    el.hidden = false;
    el.removeAttribute("aria-hidden");
    panel.classList.add("map-panel--leaflet-active");
    if (hint) hint.hidden = true;

    shipmentLeafletMap = L.map(el, {
      zoomControl: true,
      attributionControl: true,
    }).setView([48.5, 19], 4);

    L.tileLayer("https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png", {
      attribution:
        '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/attributions">CARTO</a>',
      subdomains: "abcd",
      maxZoom: 20,
    }).addTo(shipmentLeafletMap);

    shipmentMarkersLayer = L.layerGroup().addTo(shipmentLeafletMap);

    window.addEventListener("resize", scheduleShipmentLeafletResize);
    new ResizeObserver(() => scheduleShipmentLeafletResize()).observe(panel);
    requestAnimationFrame(() => scheduleShipmentLeafletResize());

    syncShipmentTrackingVisualization(cachedSupplierRecords);
  } catch (err) {
    console.error(err);
    panel.classList.remove("map-panel--leaflet-active");
    el.hidden = true;
    el.setAttribute("aria-hidden", "true");
    shipmentLeafletMap = null;
    shipmentMarkersLayer = null;
    if (hint) {
      hint.hidden = false;
      hint.textContent = "Map failed to initialize. See the browser console.";
    }
  }
}

async function loadGlobalSupplyOutlook() {
  const tbody = document.getElementById("global-supply-outlook-tbody");
  const statusEl = document.getElementById("global-supply-status");
  const countEl = document.getElementById("global-supply-active-count");
  if (!tbody) return;

  const setStatus = (msg, isError) => {
    if (!statusEl) return;
    statusEl.hidden = !msg;
    statusEl.textContent = msg;
    statusEl.classList.toggle("table-status--error", Boolean(isError));
  };

  setStatus("Loading outlook data…", false);
  await waitMs(1000);
  hasAnalyzedOutlook = false;
  inventoryReplenishedAfterPartnerResequence = false;
  alertsFeed = [];
  alertsPanelOpen = false;
  persistAppUiState();

  try {
    const text = await fetchGlobalSupplyOutlookCsvText();
    const rows = parseCSV(text);
    const { headers } = parseOutlookTable(rows);
    const theadEl = document.getElementById("global-supply-outlook-thead");
    if (theadEl) {
      theadEl.innerHTML = outlookTheadRowHtml(headers);
    }
    tbody.innerHTML = "";
    cachedOutlookHeaders = headers;
    cachedOutlookRows = [];
    alertsUnreadCount = 0;
    alertsFeed = [];
    alertsPanelOpen = false;
    renderAlertsPopover();
    setAlertsPanelOpen(false);
    if (countEl) countEl.textContent = "(0)";
    renderOutlookAlertsIdle("Outlook header loaded. Enter an event note, then click Analyze.");
    setStatus("", false);
    renderSupplyInvDisclaimer();
  } catch (e) {
    cachedOutlookHeaders = [];
    cachedOutlookRows = [];
    alertsUnreadCount = 0;
    alertsFeed = [];
    alertsPanelOpen = false;
    renderAlertsPopover();
    setAlertsPanelOpen(false);
    tbody.innerHTML = "";
    if (countEl) countEl.textContent = "(0)";
    renderOutlookAlertsIdle("Analyze unavailable until Global Supply Outlook loads.");
    const hint =
      window.location.protocol === "file:"
        ? " Use python -m http.server in this folder, then open http://localhost:8080/index.html"
        : "";
    setStatus((e instanceof Error ? e.message : "Failed to load outlook CSV.") + hint, true);
    renderSupplyInvDisclaimer();
  }
}

function alertsImpactClass(impact) {
  const norm = String(impact || "").trim().toLowerCase();
  if (norm === "high") return "alerts-item__impact alerts-item__impact--high";
  if (norm === "medium") return "alerts-item__impact alerts-item__impact--medium";
  if (norm === "info") return "alerts-item__impact alerts-item__impact--info";
  return "alerts-item__impact alerts-item__impact--low";
}

function alertItemsFromOutlookRows(headers, rows) {
  const iEvent = firstColIndexByAliases(headers, ["event"]);
  const iSupplier = firstColIndexByAliases(headers, ["impacted supplier", "supplier"]);
  const iDelay = firstColIndexByAliases(headers, ["delay"]);
  const iImpact = firstColIndexByAliases(headers, ["impact"]);
  return rows.map((r, idx) => {
    const event = iEvent >= 0 ? String(r[iEvent] || "").trim() : "";
    const supplier = iSupplier >= 0 ? String(r[iSupplier] || "").trim() : "";
    const delay = iDelay >= 0 ? String(r[iDelay] || "").trim() : "";
    const impact = iImpact >= 0 ? String(r[iImpact] || "").trim() : "Low";
    return {
      id: `evt-${idx + 1}`,
      event: event || "Unnamed event",
      supplier: supplier || "Unknown supplier",
      delay: delay || "Delay unknown",
      impact: impact || "Low",
    };
  });
}

function pushAlertFeedItem(item, options = {}) {
  const markUnread = options.markUnread !== false;
  alertsFeedSeq += 1;
  alertsFeed.unshift({
    id: `a-${alertsFeedSeq}`,
    title: String(item.title || "Alert"),
    meta: String(item.meta || ""),
    impact: String(item.impact || "Info"),
  });
  if (markUnread) alertsUnreadCount += 1;
}

function updateAlertsUnreadBadge() {
  const badge = document.getElementById("alerts-unread-count");
  const totalEl = document.getElementById("alerts-total-count");
  const badgeCount = alertsFeed.length;
  if (badge) {
    const show = badgeCount > 0;
    badge.hidden = !show;
    badge.textContent = badgeCount > 99 ? "99+" : String(badgeCount);
  }
  if (totalEl) totalEl.textContent = String(alertsFeed.length || 0);
}

function renderAlertsPopover() {
  const list = document.getElementById("alerts-popover-list");
  if (!list) return;
  if (!alertsFeed.length) {
    list.innerHTML = '<p class="alerts-popover__empty">No disruption events yet.</p>';
    updateAlertsUnreadBadge();
    return;
  }
  list.innerHTML = alertsFeed
    .map(
      (item) => `<article class="alerts-item" data-alert-id="${escapeHtml(item.id || "")}">
        <div class="alerts-item__top">
          <span class="alerts-item__event">${escapeHtml(item.title)}</span>
          <span class="${alertsImpactClass(item.impact)}">${escapeHtml(item.impact)}</span>
        </div>
        <p class="alerts-item__meta">${escapeHtml(item.meta)}</p>
      </article>`
    )
    .join("");
  updateAlertsUnreadBadge();
}

function setAlertsPanelOpen(nextOpen) {
  const pop = document.getElementById("alerts-popover");
  const btn = document.getElementById("btn-alerts-panel");
  if (!pop || !btn) return;
  alertsPanelOpen = Boolean(nextOpen);
  pop.hidden = !alertsPanelOpen;
  btn.setAttribute("aria-expanded", alertsPanelOpen ? "true" : "false");
}

function wireAlertsPopover() {
  const btn = document.getElementById("btn-alerts-panel");
  const pop = document.getElementById("alerts-popover");
  const clearBtn = document.getElementById("btn-alerts-clear");
  if (!btn || !pop || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";

  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    setAlertsPanelOpen(!alertsPanelOpen);
  });

  document.addEventListener("click", (e) => {
    if (!alertsPanelOpen) return;
    const target = e.target;
    if (!(target instanceof Element)) return;
    if (target.closest("#alerts-popover") || target.closest("#btn-alerts-panel")) return;
    setAlertsPanelOpen(false);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && alertsPanelOpen) setAlertsPanelOpen(false);
  });

  clearBtn?.addEventListener("click", () => {
    cachedOutlookRows = [];
    alertsFeed = [];
    const tbody = document.getElementById("global-supply-outlook-tbody");
    const countEl = document.getElementById("global-supply-active-count");
    if (tbody) tbody.innerHTML = "";
    if (countEl) countEl.textContent = "(0)";
    hasAnalyzedOutlook = false;
    alertsUnreadCount = 0;
    renderOutlookAlertsIdle("No impact analysis yet. Click Analyze in Global Supply Outlook.");
    renderAlertsPopover();
    persistAppUiState();
  });
  setAlertsPanelOpen(false);
  renderAlertsPopover();
}

function parseInventoryCoverageDays(val) {
  const m = String(val || "").match(/(\d+(?:\.\d+)?)/);
  return m ? parseFloat(m[1]) : 0;
}

function inventoryRiskMeta(raw) {
  const s = String(raw || "")
    .trim()
    .toLowerCase();
  if (s.includes("high"))
    return { cls: "inv-pill inv-pill--critical", short: "▲", title: String(raw || "").trim() };
  if (s.includes("critical"))
    return { cls: "inv-pill inv-pill--critical", short: "!", title: String(raw || "").trim() };
  if (s.includes("medium"))
    return { cls: "inv-pill inv-pill--medium", short: "◐", title: String(raw || "").trim() };
  if (s.includes("low"))
    return { cls: "inv-pill inv-pill--low", short: "◇", title: String(raw || "").trim() };
  if (s.includes("safe"))
    return { cls: "inv-pill inv-pill--safe", short: "✓", title: String(raw || "").trim() };
  return { cls: "inv-pill inv-pill--neutral", short: "·", title: String(raw || "").trim() };
}

function inventoryCoverageFillClass(days) {
  if (days < 2.5) return "inv-coverage-viz__fill--critical";
  if (days < 4) return "inv-coverage-viz__fill--warn";
  return "inv-coverage-viz__fill--ok";
}

const INV_SVG = {
  part: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M4 7h16M4 12h16M4 17h8"/></svg>`,
  material: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5M2 12l10 5 10-5"/></svg>`,
  seat: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="4" y="10" width="16" height="8" rx="2"/><path d="M6 10V8a3 3 0 016 0v2"/></svg>`,
  stock: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M21 16V8l-9-5-9 5v8l9 5 9-5z"/><path d="M3.27 6.96L12 12l8.73-5.05"/></svg>`,
  usage: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2"/><path d="M9 5a2 2 0 012-2h2a2 2 0 012 2v0a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/><path d="M9 12h6M9 16h6"/></svg>`,
  clock: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>`,
  truck: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M1 12h11l2 3h6v3H1z"/><circle cx="7" cy="18" r="2"/><circle cx="17" cy="18" r="2"/><path d="M14 15V9h4l3 3v3"/></svg>`,
  supplier: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M3 21h18"/><path d="M5 21V7l8-4v18"/><path d="M19 21V11l-6-4"/><path d="M9 9v0M9 13v0M9 17v0"/></svg>`,
  calendar: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/><path d="M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01"/></svg>`,
  risk: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><path d="M12 9v4M12 17h.01"/></svg>`,
  dot: `<svg class="inv-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="3"/></svg>`,
};

function inventoryThIcon(rawHeader) {
  const nk = normKey(rawHeader);
  if (nk.includes("partid") || (nk.startsWith("part") && nk.includes("id"))) return INV_SVG.part;
  if (nk.includes("supplier")) return INV_SVG.supplier;
  if (nk.includes("material")) return INV_SVG.material;
  if (nk.includes("seat")) return INV_SVG.seat;
  if (nk.includes("currentstock") || (nk.includes("stock") && !nk.includes("incoming")))
    return INV_SVG.stock;
  if (nk.includes("daily") && (nk.includes("usage") || nk.includes("consumption"))) return INV_SVG.usage;
  if (nk.includes("coverage")) return INV_SVG.clock;
  if (nk.includes("incoming") || nk.includes("nextpo")) return INV_SVG.truck;
  if (nk.includes("eta") || nk.includes("adjusted")) return INV_SVG.calendar;
  if (nk.includes("risk")) return INV_SVG.risk;
  return INV_SVG.dot;
}

function inventoryHeadRow(headers) {
  return `<tr>${headers
    .map(
      (h) =>
        `<th scope="col"><span class="inv-th">${inventoryThIcon(h)}<span class="inv-th__txt">${escapeHtml(
          String(h).trim()
        )}</span></span></th>`
    )
    .join("")}</tr>`;
}

function inventoryCellHtml(header, val, scales) {
  const nk = normKey(header);
  const raw = String(val ?? "").trim();
  const escaped = escapeHtml(raw);

  if (nk.includes("risk")) {
    const m = inventoryRiskMeta(raw);
    return `<td class="inv-cell inv-cell--risk"><span class="${m.cls}" title="${escapeHtml(m.title)}"><span class="inv-pill__glyph" aria-hidden="true">${escapeHtml(m.short)}</span><span class="inv-pill__txt">${escaped}</span></span></td>`;
  }

  if (nk.includes("coverage")) {
    const days = parseInventoryCoverageDays(raw);
    const cap = Math.max(scales.maxCoverage, 0.001);
    const pct = Math.min(100, (days / cap) * 100);
    const fillCls = inventoryCoverageFillClass(days);
    const shortVal = raw.replace(/\s*days?\s*/i, " d").trim();
    return `<td class="inv-cell inv-cell--coverage">
      <div class="inv-coverage-viz" title="${escaped}">
        <span class="inv-coverage-viz__label">${escapeHtml(shortVal || "—")}</span>
        <div class="inv-coverage-viz__track" role="presentation"><div class="inv-coverage-viz__fill ${fillCls}" style="width:${pct.toFixed(1)}%"></div></div>
      </div>
    </td>`;
  }

  if (
    (nk.includes("daily") && (nk.includes("usage") || nk.includes("consumption"))) ||
    nk.includes("dailyusage") ||
    nk.includes("dailyconsumption")
  ) {
    return `<td class="inv-cell inv-cell--usage">${escaped}</td>`;
  }

  if (nk.includes("incoming") || (nk.includes("supply") && nk.includes("next"))) {
    return `<td class="inv-cell inv-cell--inbound">
      <span class="inv-inbound">${INV_SVG.truck}<span class="inv-inbound__txt">${escaped}</span></span>
    </td>`;
  }

  if (nk.includes("supplier")) {
    return `<td class="inv-cell inv-cell--supplier">
      <span class="inv-supplier">${INV_SVG.supplier}<span class="inv-supplier__txt">${escaped}</span></span>
    </td>`;
  }

  if (nk.includes("material") && !nk.includes("risk")) {
    return `<td class="inv-cell inv-cell--material">
      <span class="inv-material">${INV_SVG.material}<span class="inv-material__txt">${escaped}</span></span>
    </td>`;
  }

  if (nk.includes("adjustedeta") || (nk.includes("eta") && !nk.includes("risk"))) {
    return `<td class="inv-cell inv-cell--eta">
      <span class="inv-eta">${INV_SVG.calendar}<span class="inv-eta__txt">${escaped}</span></span>
    </td>`;
  }

  if (nk.includes("seattype") || (nk.includes("seat") && nk.includes("type"))) {
    return `<td class="inv-cell inv-cell--seat">
      <span class="inv-seat">${INV_SVG.seat}<span class="inv-seat__txt">${escaped}</span></span>
    </td>`;
  }

  if (nk.includes("partid") || (nk.includes("part") && nk.includes("id"))) {
    return `<td class="inv-cell"><span class="inv-chip">${escaped}</span></td>`;
  }

  if (nk.includes("currentstock")) {
    return `<td class="inv-cell inv-cell--stock">
      <span class="inv-stock">${INV_SVG.stock}<span class="inv-stock__txt">${escaped}</span></span>
    </td>`;
  }

  return `<td class="inv-cell">${escaped}</td>`;
}

function inventoryDataRow(headers, cells, scales) {
  const tds = headers.map((h, i) => inventoryCellHtml(h, cells[i] != null ? cells[i] : "", scales));
  return `<tr>${tds.join("")}</tr>`;
}

function parseInventoryTable(parsedRows) {
  if (parsedRows.length < 1) return { headers: [], body: [] };
  const headers = parsedRows[0].map((h) => String(h).trim());
  const body = [];
  for (let r = 1; r < parsedRows.length; r++) {
    const row = parsedRows[r];
    if (!row || row.every((c) => !String(c ?? "").trim())) continue;
    const cells = row.map((c) => (c != null ? String(c) : ""));
    while (cells.length < headers.length) cells.push("");
    body.push(cells.slice(0, headers.length));
  }
  return { headers, body };
}

function inventoryScales(headers, body) {
  let maxCoverage = 1;
  const covIdx = headers.findIndex((h) => normKey(h).includes("coverage"));
  for (const row of body) {
    if (covIdx >= 0) maxCoverage = Math.max(maxCoverage, parseInventoryCoverageDays(row[covIdx]));
  }
  return { maxCoverage };
}

function inventoryFindCol(headers, test) {
  return headers.findIndex((h) => test(normKey(String(h).trim())));
}

function inventoryColumnIndices(headers) {
  return {
    partId: inventoryFindCol(headers, (n) => n.includes("partid") || (n.includes("part") && n.includes("id"))),
    material: inventoryFindCol(headers, (n) => n.includes("material") && !n.includes("risk")),
    supplier: inventoryFindCol(headers, (n) => n.includes("supplier")),
    stock: inventoryFindCol(
      headers,
      (n) => n.includes("currentstock") || (n.includes("stock") && !n.includes("incoming"))
    ),
    dailyUsage: inventoryFindCol(
      headers,
      (n) =>
        (n.includes("daily") && (n.includes("usage") || n.includes("consumption"))) ||
        n.includes("dailyusage") ||
        n.includes("dailyconsumption")
    ),
    coverage: inventoryFindCol(headers, (n) => n.includes("coverage")),
    incoming: inventoryFindCol(
      headers,
      (n) => n.includes("incoming") || (n.includes("supply") && n.includes("next"))
    ),
    eta: inventoryFindCol(
      headers,
      (n) => n.includes("adjustedeta") || (n.includes("eta") && !n.includes("risk"))
    ),
    risk: inventoryFindCol(headers, (n) => n.includes("risk")),
  };
}

function supplyViewColumnMaxima(headers, body) {
  const ix = inventoryColumnIndices(headers);
  const out = { maxStock: 1, maxIncoming: 1, maxCoverage: 1 };
  for (const row of body) {
    if (ix.stock >= 0) {
      const n = Number(String(row[ix.stock] ?? "").replace(/[^\d.\-]/g, ""));
      if (Number.isFinite(n)) out.maxStock = Math.max(out.maxStock, n);
    }
    if (ix.incoming >= 0) {
      const n = Number(String(row[ix.incoming] ?? "").replace(/[^\d.\-]/g, ""));
      if (Number.isFinite(n)) out.maxIncoming = Math.max(out.maxIncoming, n);
    }
    if (ix.coverage >= 0) {
      out.maxCoverage = Math.max(out.maxCoverage, parseInventoryCoverageDays(row[ix.coverage]));
    }
  }
  return out;
}

function supplyViewFillHeightPct(value, maxVal) {
  const n = Number(value);
  if (!Number.isFinite(n) || maxVal <= 0) return 8;
  return Math.max(8, Math.min(100, (n / maxVal) * 100));
}

function seatTokenLetterFromCell(raw) {
  const s = String(raw || "").toLowerCase();
  if (s.includes("all seat")) return "";
  if (s.includes("blue")) return "B";
  if (s.includes("red") || s.includes("black")) return "K";
  if (s.includes("white")) return "W";
  return "";
}

function supplySeatSeqTokenClass(letter) {
  const L = String(letter || "").trim().toUpperCase();
  if (L === "B") return "seq-token seq-token--b";
  if (L === "K") return "seq-token seq-token--k";
  if (L === "W") return "seq-token seq-token--w";
  return "seq-token seq-token--n";
}

function seqTokenClassForRisk(raw) {
  const s = String(raw || "").toLowerCase();
  if (s.includes("critical") || s.includes("high")) return "seq-token seq-token--risk-crit";
  if (s.includes("medium")) return "seq-token seq-token--risk-med";
  if (s.includes("low") || s.includes("watch")) return "seq-token seq-token--risk-low";
  if (s.includes("safe")) return "seq-token seq-token--risk-safe";
  return "seq-token seq-token--n";
}

function supplyViewInventoryCellHtml(header, val, maxima) {
  const nk = normKey(header);
  const raw = String(val ?? "").trim();
  const escaped = escapeHtml(raw);
  const row = (inner) => `<span class="supply-inv-cyl-row">${inner}</span>`;

  if (nk.includes("risk")) {
    const tok = `<span class="${seqTokenClassForRisk(raw)}" title="${escaped}" aria-label="${escaped}"></span>`;
    return `<td class="supply-inv-cell supply-inv-cell--risk">${row(
      `${tok}<span class="supply-inv-val">${escaped}</span>`
    )}</td>`;
  }

  if (nk.includes("coverage")) {
    const days = parseInventoryCoverageDays(raw);
    const pct = supplyViewFillHeightPct(days, maxima.maxCoverage);
    const toneCls =
      days < 2.5 ? "inv-fill-cyl__inner--crit" : days < 4 ? "inv-fill-cyl__inner--warn" : "inv-fill-cyl__inner--ok";
    const shortVal = raw.replace(/\s*days?\s*/i, " d").trim();
    return `<td class="supply-inv-cell supply-inv-cell--coverage">${row(
      `<span class="inv-fill-cyl" title="${escaped}"><span class="inv-fill-cyl__inner ${toneCls}" style="height:${pct.toFixed(
        1
      )}%"></span></span><span class="supply-inv-val">${escapeHtml(shortVal || "—")}</span>`
    )}</td>`;
  }

  if (nk.includes("seattype") || (nk.includes("seat") && nk.includes("type"))) {
    const letter = seatTokenLetterFromCell(raw);
    const tok = letter
      ? `<span class="${supplySeatSeqTokenClass(letter)}" title="${escaped}" aria-label="${escaped}"></span>`
      : `<span class="seq-token seq-token--n" title="${escaped}" aria-label="${escaped}"></span>`;
    return `<td class="supply-inv-cell supply-inv-cell--seat">${row(
      `${tok}<span class="supply-inv-val">${escaped}</span>`
    )}</td>`;
  }

  if (nk.includes("currentstock") || (nk.includes("stock") && !nk.includes("incoming"))) {
    const n = Number(String(raw).replace(/[^\d.\-]/g, ""));
    const pct = supplyViewFillHeightPct(n, maxima.maxStock);
    return `<td class="supply-inv-cell supply-inv-cell--stock">${row(
      `<span class="inv-fill-cyl inv-fill-cyl--wide" title="${escaped}"><span class="inv-fill-cyl__inner inv-fill-cyl__inner--stock" style="height:${pct.toFixed(
        1
      )}%"></span></span><span class="supply-inv-val supply-inv-val--num">${escaped}</span>`
    )}</td>`;
  }

  if (nk.includes("incoming") || (nk.includes("supply") && nk.includes("next"))) {
    const n = Number(String(raw).replace(/[^\d.\-]/g, ""));
    const pct = Number.isFinite(n) ? supplyViewFillHeightPct(n, maxima.maxIncoming) : 8;
    return `<td class="supply-inv-cell supply-inv-cell--inbound">${row(
      `<span class="inv-fill-cyl inv-fill-cyl--wide" title="${escaped}"><span class="inv-fill-cyl__inner inv-fill-cyl__inner--inbound" style="height:${pct.toFixed(
        1
      )}%"></span></span><span class="supply-inv-val">${escaped}</span>`
    )}</td>`;
  }

  if (
    (nk.includes("daily") && (nk.includes("usage") || nk.includes("consumption"))) ||
    nk.includes("dailyusage") ||
    nk.includes("dailyconsumption")
  ) {
    return `<td class="supply-inv-cell supply-inv-cell--usage">${row(
      `<span class="seq-token seq-token--usage" title="Daily usage" aria-hidden="true"></span><span class="supply-inv-val">${escaped}</span>`
    )}</td>`;
  }

  if (nk.includes("material") && !nk.includes("risk")) {
    return `<td class="supply-inv-cell supply-inv-cell--material">${row(
      `<span class="seq-token seq-token--material" title="Material" aria-hidden="true"></span><span class="supply-inv-val">${escaped}</span>`
    )}</td>`;
  }

  if (nk.includes("supplier")) {
    return `<td class="supply-inv-cell supply-inv-cell--supplier">${row(
      `<span class="seq-token seq-token--supplier" title="Supplier" aria-hidden="true"></span><span class="supply-inv-val">${escaped}</span>`
    )}</td>`;
  }

  if (nk.includes("adjustedeta") || (nk.includes("eta") && !nk.includes("risk"))) {
    return `<td class="supply-inv-cell supply-inv-cell--eta">${row(
      `<span class="seq-token seq-token--eta" title="ETA" aria-hidden="true"></span><span class="supply-inv-val">${escaped}</span>`
    )}</td>`;
  }

  if (nk.includes("partid") || (nk.includes("part") && nk.includes("id"))) {
    return `<td class="supply-inv-cell supply-inv-cell--part">${row(
      `<span class="seq-token seq-token--part" title="Part" aria-hidden="true"></span><span class="supply-inv-val supply-inv-val--mono">${escaped}</span>`
    )}</td>`;
  }

  return `<td class="supply-inv-cell">${row(
    `<span class="seq-token seq-token--n" aria-hidden="true"></span><span class="supply-inv-val">${escaped}</span>`
  )}</td>`;
}

function supplyViewInventoryDataRow(headers, cells, maxima) {
  const tds = headers.map((h, i) =>
    supplyViewInventoryCellHtml(h, cells[i] != null ? cells[i] : "", maxima)
  );
  return `<tr>${tds.join("")}</tr>`;
}

function inventoryCellAt(row, idx) {
  if (idx < 0) return "";
  return String(row[idx] ?? "").trim();
}

/** @returns {"high" | "medium" | "low" | null} */
function inventoryAlertRiskTier(raw) {
  const s = String(raw ?? "")
    .trim()
    .toLowerCase();
  if (s === "high" || s.startsWith("high ")) return "high";
  if (s === "medium" || s.startsWith("medium ") || s === "watch" || s.startsWith("watch ")) return "medium";
  if (s === "low" || s.startsWith("low ")) return "low";
  if (s === "safe" || s.startsWith("safe ")) return null;
  return null;
}

function escapeAttr(s) {
  return String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

function inventoryStockoutCardHtml(row, ix, tier) {
  const partRaw = inventoryCellAt(row, ix.partId);
  const part = escapeHtml(partRaw || "—");
  const material = escapeHtml(inventoryCellAt(row, ix.material));
  const supplier = escapeHtml(inventoryCellAt(row, ix.supplier) || "—");
  const stock = escapeHtml(inventoryCellAt(row, ix.stock) || "—");
  const incoming = escapeHtml(inventoryCellAt(row, ix.incoming) || "—");
  const eta = escapeHtml(inventoryCellAt(row, ix.eta) || "");
  const usage = escapeHtml(inventoryCellAt(row, ix.dailyUsage) || "—");
  const daysRaw = inventoryCellAt(row, ix.coverage);
  const daysDisp = escapeHtml(daysRaw || "—");
  const mod =
    tier === "high"
      ? "inv-stockout-card--high"
      : tier === "medium"
        ? "inv-stockout-card--medium"
        : "inv-stockout-card--low";
  const badge = tier === "high" ? "High risk" : tier === "medium" ? "Watch" : "Low risk";
  const titleMat =
    material && material !== "—"
      ? `<span class="inv-stockout-card__mat">${material}</span>`
      : "";
  const hintParts = [];
  if (daysRaw) hintParts.push(`${escapeHtml(daysRaw)} coverage`);
  const supRaw = inventoryCellAt(row, ix.supplier);
  if (supRaw) hintParts.push(escapeHtml(supRaw));
  const hint = hintParts.join(" · ");

  return `<details class="inv-stockout-card ${mod}" data-part-id="${escapeAttr(partRaw || "")}">
  <summary class="inv-stockout-card__summary">
    <span class="inv-stockout-card__summary-text">
      <span class="inv-stockout-card__badge inv-stockout-card__badge--${tier}">${escapeHtml(badge)}</span>
      <span class="inv-stockout-card__title"><span class="inv-stockout-card__part">${part}</span>${titleMat}</span>
      ${hint ? `<span class="inv-stockout-card__hint">${hint}</span>` : ""}
    </span>
    <span class="inv-stockout-card__chevron" aria-hidden="true">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M6 9l6 6 6-6" />
      </svg>
    </span>
  </summary>
  <div class="inv-stockout-card__body">
    <dl class="inv-stockout-card__dl">
      <div><dt>Supplier</dt><dd>${supplier}</dd></div>
      <div><dt>Coverage</dt><dd>${daysDisp}</dd></div>
      <div><dt>Stock</dt><dd>${stock}</dd></div>
      <div><dt>Usage</dt><dd>${usage}</dd></div>
      <div><dt>Inbound</dt><dd>${incoming}${eta ? ` <span class="inv-stockout-card__eta">· ${eta}</span>` : ""}</dd></div>
    </dl>
  </div>
</details>`;
}

function renderInventoryStockoutSection(headers, body) {
  const root = document.getElementById("inventory-alert-cards");
  if (!root) return;
  if (!headers || headers.length === 0) {
    root.innerHTML =
      '<p class="alert-cards-empty">Load inventory data to see stockout alerts.</p>';
    return;
  }
  if (!body || body.length === 0) {
    root.innerHTML =
      '<p class="alert-cards-empty">No data rows in InventoryStatus.csv — no stockout alerts to show.</p>';
    return;
  }
  const ix = inventoryColumnIndices(headers);
  if (ix.risk < 0) {
    root.innerHTML =
      '<p class="alert-cards-empty">Add a Risk Level column to surface High and Low stockout alerts.</p>';
    return;
  }

  const alerts = body
    .map((row) => ({ row, tier: inventoryAlertRiskTier(inventoryCellAt(row, ix.risk)) }))
    .filter((x) => x.tier !== null);

  if (alerts.length === 0) {
    root.innerHTML =
      '<p class="alert-cards-empty">No non-safe rows in this file — no stockout alerts in this view.</p>';
    return;
  }

  alerts.sort((a, b) => {
    const order = { high: 0, medium: 1, low: 2 };
    const oa = order[a.tier] ?? 3;
    const ob = order[b.tier] ?? 3;
    if (oa !== ob) return oa - ob;
    const da = parseInventoryCoverageDays(inventoryCellAt(a.row, ix.coverage));
    const db = parseInventoryCoverageDays(inventoryCellAt(b.row, ix.coverage));
    return da - db;
  });

  root.innerHTML = alerts.map(({ row, tier }) => inventoryStockoutCardHtml(row, ix, tier)).join("");
}

async function loadInventoryInsights() {
  const tbody = document.getElementById("inventory-insights-tbody");
  const thead = document.getElementById("inventory-insights-thead");
  const statusEl = document.getElementById("inventory-insights-status");
  if (!tbody || !thead) return;

  const setStatus = (msg, isError) => {
    if (!statusEl) return;
    statusEl.hidden = !msg;
    statusEl.textContent = msg;
    statusEl.classList.toggle("table-status--error", Boolean(isError));
  };

  setStatus("Loading inventory…", false);
  const alertRoot = document.getElementById("inventory-alert-cards");
  if (alertRoot) {
    alertRoot.innerHTML = '<p class="alert-cards-empty">Loading stockout alerts…</p>';
  }
  try {
    const text = await fetchInventoryStatusCsvText();
    const rows = parseCSV(text);
    const { headers, body } = parseInventoryTable(rows);
    if (headers.length === 0) {
      thead.innerHTML = '<tr><th class="inventory-th-placeholder">No columns</th></tr>';
      tbody.innerHTML = "";
      renderInventoryStockoutSection([], []);
      setStatus("Inventory file has no header row.", true);
      return;
    }
    if (body.length === 0) {
      thead.innerHTML = inventoryHeadRow(headers);
      tbody.innerHTML = "";
      renderInventoryStockoutSection(headers, []);
      setStatus("No data rows in InventoryStatus.csv.", true);
      return;
    }
    const scales = inventoryScales(headers, body);
    thead.innerHTML = inventoryHeadRow(headers);
    tbody.innerHTML = body.map((cells) => inventoryDataRow(headers, cells, scales)).join("");
    renderInventoryStockoutSection(headers, body);
    setStatus("", false);
  } catch (e) {
    document.getElementById("panel-inventory")?.removeAttribute("data-inventory-loaded");
    thead.innerHTML = '<tr><th class="inventory-th-placeholder">—</th></tr>';
    tbody.innerHTML = "";
    if (alertRoot) {
      alertRoot.innerHTML =
        '<p class="alert-cards-empty">Could not load stockout alerts — fix CSV access and refresh.</p>';
    }
    const hint =
      window.location.protocol === "file:"
        ? " Use python -m http.server in this folder, then open http://localhost:8080/index.html"
        : "";
    setStatus((e instanceof Error ? e.message : "Failed to load inventory CSV.") + hint, true);
  }
}

function parseMpsTable(parsedRows) {
  if (!Array.isArray(parsedRows) || parsedRows.length < 1) return { headers: [], body: [] };
  const headers = parsedRows[0].map((h) => String(h ?? "").trim());
  const body = [];
  for (let i = 1; i < parsedRows.length; i++) {
    const row = parsedRows[i];
    if (!row || row.every((c) => !String(c ?? "").trim())) continue;
    const cells = headers.map((_, idx) => String(row[idx] ?? "").trim());
    body.push(cells);
  }
  return { headers, body };
}

function mpsRowMapBySeat(headers, body) {
  const out = new Map();
  const seatIdx = 0;
  for (const row of body) {
    const key = normKey(String(row[seatIdx] || ""));
    if (key) out.set(key, row);
  }
  return out;
}

function mpsValueToneClass(raw, maxVal) {
  const n = Number(String(raw ?? "").trim());
  if (!Number.isFinite(n)) return "mps-cell--txt";
  if (maxVal <= 0) return "mps-cell--l1";
  const ratio = n / maxVal;
  if (ratio >= 0.85) return "mps-cell--l4";
  if (ratio >= 0.65) return "mps-cell--l3";
  if (ratio >= 0.45) return "mps-cell--l2";
  return "mps-cell--l1";
}

function mpsHeatStyle(seatRaw, rawVal, rowMax) {
  const n = Number(String(rawVal ?? "").trim());
  if (!Number.isFinite(n)) return "";
  const cap = Math.max(rowMax, 1);
  const ratio = Math.max(0, Math.min(1, n / cap));
  const key = normKey(seatRaw);

  let hue = 190; // blue-ish default
  let sat = 58;
  if (key.includes("red") || key.includes("black")) {
    hue = 3;
    sat = 72;
  } else if (key.includes("white")) {
    hue = 220;
    sat = 22;
  } else if (key.includes("total")) {
    hue = 24;
    sat = 78;
  }

  const lightTop = key.includes("white") ? 26 + ratio * 22 : 20 + ratio * 18;
  const lightBottom = key.includes("white") ? 18 + ratio * 14 : 14 + ratio * 14;
  const glowAlpha = 0.08 + ratio * 0.2;
  return `background: radial-gradient(130% 120% at 16% 10%, hsla(${hue}, ${sat}%, 72%, ${glowAlpha.toFixed(
    2
  )}) 0%, transparent 68%), linear-gradient(180deg, hsl(${hue} ${sat}% ${lightTop.toFixed(
    1
  )}%), hsl(${hue} ${sat}% ${lightBottom.toFixed(1)}%));`;
}

function mpsSeatDotClass(seat) {
  const n = normKey(seat);
  if (n.includes("blue")) return "mps-dot mps-dot--blue";
  if (n.includes("red")) return "mps-dot mps-dot--red";
  if (n.includes("black")) return "mps-dot mps-dot--red";
  if (n.includes("white")) return "mps-dot mps-dot--white";
  return "mps-dot mps-dot--neutral";
}

function buildMpsTableHtml(headers, body, diffMap = new Map()) {
  const seatCol = 0;
  let maxVal = 0;
  body.forEach((row) => {
    for (let c = 1; c < headers.length; c++) {
      const n = Number(String(row[c] ?? "").trim());
      if (Number.isFinite(n)) maxVal = Math.max(maxVal, n);
    }
  });

  const head = `<tr>${headers
    .map((h, i) =>
      i === 0
        ? `<th scope="col" class="mps-th mps-th--seat">${escapeHtml(h)}</th>`
        : `<th scope="col" class="mps-th">${escapeHtml(h)}</th>`
    )
    .join("")}</tr>`;

  const bodyHtml = body
    .map((row) => {
      const seatRaw = String(row[seatCol] ?? "").trim();
      const seatCell = `<th scope="row" class="mps-seat"><span class="${mpsSeatDotClass(
        seatRaw
      )}" aria-hidden="true"></span><span class="mps-seat__txt">${escapeHtml(seatRaw || "—")}</span></th>`;
      let rowMax = 0;
      for (let c = 1; c < headers.length; c++) {
        const n = Number(String(row[c] ?? "").trim());
        if (Number.isFinite(n)) rowMax = Math.max(rowMax, n);
      }
      const vals = headers
        .slice(1)
        .map((_, cIdx) => {
          const v = String(row[cIdx + 1] ?? "").trim();
          const key = `${normKey(seatRaw)}|${cIdx + 1}`;
          const delta = diffMap.get(key);
          const hasDelta = Number.isFinite(delta) && delta !== 0;
          const trendCls = hasDelta ? (delta > 0 ? "mps-delta mps-delta--up" : "mps-delta mps-delta--down") : "";
          const trendTxt = hasDelta ? (delta > 0 ? `+${delta}` : `${delta}`) : "";
          const heatStyle = mpsHeatStyle(seatRaw, v, rowMax || maxVal);
          return `<td class="mps-cell ${mpsValueToneClass(v, maxVal)}${hasDelta ? " mps-cell--alerted" : ""}" style="${heatStyle}">
            <span class="mps-cell__val">${escapeHtml(v || "—")}</span>
            ${hasDelta ? `<span class="${trendCls}" title="Changed vs baseline">${escapeHtml(trendTxt)}</span>` : ""}
          </td>`;
        })
        .join("");
      return `<tr>${seatCell}${vals}</tr>`;
    })
    .join("");

  return { head, bodyHtml };
}

/** Supplier grid: columns like Supplier A1…A9 with B/R/W (or K) seat codes per cell. */
function isSupplierMatrixInventory(headers) {
  const h = headers.map((x) => String(x ?? "").trim()).filter(Boolean);
  if (h.length < 2) return false;
  return h.every((cell) => /supplier\s*a\s*\d+/i.test(cell));
}

function supplierMatrixTokenClass(raw) {
  const s = String(raw ?? "")
    .trim()
    .toUpperCase();
  if (s === "B") return supplySeatSeqTokenClass("B");
  if (s === "R" || s === "K") return supplySeatSeqTokenClass("K");
  if (s === "W") return supplySeatSeqTokenClass("W");
  return "seq-token seq-token--n";
}

/** Map sequence slot letter to inventory bucket (K → red / R). */
function supplyMatrixSequenceWantLetter(v) {
  const s = String(v ?? "")
    .trim()
    .toUpperCase();
  if (s === "B") return "B";
  if (s === "K") return "R";
  if (s === "W") return "W";
  return "";
}

/**
 * Remove one stocked unit matching the production slot color (first free slot left→right).
 * Baseline holds B/R/W per supplier column; state[i] is '' when consumed.
 */
function supplyMatrixConsumeOne(state, baseline, rawSlot) {
  const want = supplyMatrixSequenceWantLetter(rawSlot);
  if (!want) return;
  for (let i = 0; i < 9; i++) {
    if (state[i]) {
      const b = String(baseline[i] ?? "")
        .trim()
        .toUpperCase();
      if (b === want) {
        state[i] = "";
        return;
      }
    }
  }
}

/**
 * After T production ticks: consume from baseline in global sequence order; refill all
 * suppliers at the end of each shift row (after column 9). After the T-th tick on the last
 * column of a shift, refill so the next shift starts full.
 */
function computeInventorySlotsAfterTicks(baseline, schedule, T) {
  const B = baseline.map((c) => String(c ?? "").trim().toUpperCase());
  if (!schedule?.length) return B.slice();

  const blockedByShift = blockedInventoryColumnsByShift(schedule.length);
  const stateForShift = (rowIdx) => {
    const st = B.slice();
    const blocked = blockedByShift[rowIdx] || new Set();
    blocked.forEach((ci) => {
      if (ci >= 0 && ci < st.length) st[ci] = "";
    });
    return st;
  };

  if (T <= 0) return stateForShift(0);

  let state = stateForShift(0);
  let tickNum = 0;
  for (let r = 0; r < schedule.length; r++) {
    const row = schedule[r];
    state = stateForShift(r);
    const produced = (row.slots || []).filter((slot) => isSequenceProductLetter(slot));
    if (produced.length === 0) {
      state = stateForShift(Math.min(r + 1, schedule.length - 1));
      continue;
    }
    for (let i = 0; i < produced.length; i++) {
      tickNum++;
      supplyMatrixConsumeOne(state, B, produced[i]);
      if (tickNum === T) {
        if (i === produced.length - 1) {
          const nextRow = Math.min(r + 1, schedule.length - 1);
          state = stateForShift(nextRow);
        }
        return state;
      }
      if (i === produced.length - 1) {
        const nextRow = Math.min(r + 1, schedule.length - 1);
        state = stateForShift(nextRow);
      }
    }
  }
  return stateForShift(schedule.length - 1);
}

function supplyMatrixCellHtmlForSlot(currentLetter, baselineLetterForTone) {
  const base = String(baselineLetterForTone ?? "")
    .trim()
    .toUpperCase();
  const raw = String(currentLetter ?? "")
    .trim()
    .toUpperCase();
  if (!raw) {
    const tone =
      base === "B" ? "b" : base === "R" || base === "K" ? "r" : base === "W" ? "w" : "n";
    const aria =
      base === "B"
        ? "Consumed (blue)"
        : base === "R" || base === "K"
          ? "Consumed (red)"
          : base === "W"
            ? "Consumed (white)"
            : "Consumed";
    return `<td class="supply-mx-cell"><div class="supply-mx-cyl"><span class="supply-mx-empty-slot supply-mx-empty-slot--${tone}" title="Consumed" aria-label="${escapeHtml(
      aria
    )}"></span></div></td>`;
  }
  const cls = supplierMatrixTokenClass(raw);
  const label =
    raw === "B" ? "Blue" : raw === "R" || raw === "K" ? "Red" : raw === "W" ? "White" : "Empty";
  const tok = `<span class="${cls}" title="${escapeHtml(label)}" aria-label="${escapeHtml(label)}"></span>`;
  return `<td class="supply-mx-cell"><div class="supply-mx-cyl">${tok}</div></td>`;
}

function buildSupplyMatrixDataRowFromSlots(headers, slotLetters, rowIdx, hasRowLabels, baselineRow) {
  const tds = headers
    .map((_, ci) => supplyMatrixCellHtmlForSlot(slotLetters[ci] ?? "", baselineRow[ci] ?? ""))
    .join("");
  const rowHead = hasRowLabels
    ? `<th scope="row" class="supply-mx-rowlabel">${escapeHtml(`Row ${rowIdx + 1}`)}</th>`
    : "";
  return `<tr class="supply-mx-row">${rowHead}${tds}</tr>`;
}

function buildSupplyMatrixTbodyHtml(headers, primarySlotLetters, extraStaticRows) {
  const rows = [];
  const hasRowLabels = extraStaticRows.length > 0;
  rows.push(
    buildSupplyMatrixDataRowFromSlots(headers, primarySlotLetters, 0, hasRowLabels, inventoryMatrixBaseline)
  );
  extraStaticRows.forEach((row, ri) => {
    const slotLetters = headers.map((_, ci) => String(row[ci] ?? "").trim());
    rows.push(
      buildSupplyMatrixDataRowFromSlots(headers, slotLetters, ri + 1, hasRowLabels, slotLetters)
    );
  });
  return rows.join("");
}

function supplyMatrixThDisplayLabel(rawHeader, colIndex) {
  const base = String(rawHeader ?? "").trim();
  if (!inventoryReplenishedAfterPartnerResequence) return base;
  if (delayShiftsForSupplierColumn(colIndex) <= 0) return base;
  return `${base} (Replenished)`;
}

function buildSupplyMatrixTheadHtml(headers, hasRowLabels) {
  const impactedCols = impactedSupplierColumnSet(currentSequenceShiftIndex());
  const thCells = headers
    .map((h, i) => {
      const impacted = impactedCols.has(i);
      const label = supplyMatrixThDisplayLabel(h, i);
      return `<th scope="col" class="supply-mx-th${
        impacted ? " supply-mx-th--impacted" : ""
      }${inventoryReplenishedAfterPartnerResequence && delayShiftsForSupplierColumn(i) > 0 ? " supply-mx-th--replenished" : ""}"><span class="supply-mx-th__txt">${escapeHtml(label)}</span>${
        impacted
          ? '<span class="supply-mx-warn" title="Impacted supplier delay" aria-label="Impacted supplier delay">⚠</span>'
          : ""
      }</th>`;
    })
    .join("");
  return hasRowLabels
    ? `<tr><th class="supply-mx-corner" scope="col"><span class="visually-hidden">Row</span></th>${thCells}</tr>`
    : `<tr>${thCells}</tr>`;
}

/** Updates matrix tbody from sequence ticks (no-op if not in supplier-matrix mode). */
function refreshInventorySupplyViewForSequence() {
  const tbody = document.getElementById("supply-inv-tbody");
  const thead = document.getElementById("supply-inv-thead");
  if (!tbody || !thead || !inventoryMatrixBaseline.length || !inventoryMatrixHeaders.length) return;
  const tableEl = thead.closest("table");
  if (!tableEl?.classList.contains("supply-inv-table--matrix")) return;
  thead.innerHTML = buildSupplyMatrixTheadHtml(
    inventoryMatrixHeaders,
    inventoryMatrixExtraRows.length > 0
  );
  const slots = computeInventorySlotsAfterTicks(
    inventoryMatrixBaseline,
    currentSequenceSchedule,
    sequenceCycleTicks
  );
  tbody.innerHTML = buildSupplyMatrixTbodyHtml(inventoryMatrixHeaders, slots, inventoryMatrixExtraRows);
}

async function loadMpsInsights() {
  const thead = document.getElementById("supply-inv-thead");
  const tbody = document.getElementById("supply-inv-tbody");
  const statusEl = document.getElementById("supply-inv-status");
  if (!thead || !tbody) return;

  const setStatus = (msg, isError) => {
    if (!statusEl) return;
    statusEl.hidden = !msg;
    statusEl.textContent = msg;
    statusEl.classList.toggle("table-status--error", Boolean(isError));
  };

  setStatus("Loading inventory…", false);
  try {
    const text = await fetchInventoryStatusCsvText();
    const rows = parseCSV(text);
    const { headers, body } = parseInventoryTable(rows);
    inventoryMatrixHeaders = [];
    inventoryMatrixBaseline = [];
    inventoryMatrixExtraRows = [];
    const tableEl = thead.closest("table");
    if (tableEl) tableEl.classList.remove("supply-inv-table--matrix");

    if (headers.length === 0) {
      thead.innerHTML = '<tr><th class="inventory-th-placeholder">No columns in inventory file.</th></tr>';
      tbody.innerHTML = "";
      setStatus("Inventory file has no header row.", true);
      return;
    }
    if (body.length === 0) {
      thead.innerHTML = '<tr><th class="inventory-th-placeholder">No inventory rows found.</th></tr>';
      tbody.innerHTML = "";
      setStatus("No data rows in inventory file.", true);
      return;
    }

    if (isSupplierMatrixInventory(headers)) {
      if (tableEl) tableEl.classList.add("supply-inv-table--matrix");
      inventoryMatrixHeaders = headers;
      inventoryMatrixBaseline = headers.map((_, ci) => String(body[0][ci] ?? "").trim().toUpperCase());
      inventoryMatrixExtraRows = body.slice(1);
      thead.innerHTML = buildSupplyMatrixTheadHtml(headers, inventoryMatrixExtraRows.length > 0);
      refreshInventorySupplyViewForSequence();
      setStatus("", false);
      renderSupplyInvDisclaimer();
      return;
    }

    const maxima = supplyViewColumnMaxima(headers, body);
    thead.innerHTML = inventoryHeadRow(headers);
    tbody.innerHTML = body.map((cells) => supplyViewInventoryDataRow(headers, cells, maxima)).join("");
    setStatus("", false);
    renderSupplyInvDisclaimer();
  } catch (e) {
    inventoryMatrixHeaders = [];
    inventoryMatrixBaseline = [];
    inventoryMatrixExtraRows = [];
    thead.closest("table")?.classList.remove("supply-inv-table--matrix");
    thead.innerHTML = '<tr><th class="inventory-th-placeholder">Could not load inventory.</th></tr>';
    tbody.innerHTML = "";
    const hint =
      window.location.protocol === "file:"
        ? " Use python -m http.server in this folder, then open http://localhost:8080/index.html"
        : "";
    setStatus((e instanceof Error ? e.message : "Failed to load inventory CSV.") + hint, true);
    renderSupplyInvDisclaimer();
  }
}

function parseSequenceTable(parsedRows) {
  if (!Array.isArray(parsedRows) || parsedRows.length < 1) return { headers: [], body: [] };
  const headers = parsedRows[0].map((h) => String(h ?? "").trim());
  const body = [];
  for (let i = 1; i < parsedRows.length; i++) {
    const row = parsedRows[i];
    if (!row || row.every((c) => !String(c ?? "").trim())) continue;
    body.push(headers.map((_, idx) => String(row[idx] ?? "").trim()));
  }
  return { headers, body };
}

/**
 * Unify SequenceBefore (Date + Shift + 1..9) and SequenceAfter (Date + 1..9, no Shift).
 * Each output row is one shift line with 9 product slots (CSV columns per shift).
 */
function normalizeSequenceSchedule(headers, body) {
  const h = headers.map((x) => String(x ?? "").trim());
  const shiftIdx = h.findIndex((x) => x.toLowerCase() === "shift");
  const hasShift = shiftIdx >= 0;
  const out = [];
  let lastDate = "";

  if (hasShift) {
    for (const row of body) {
      let d = String(row[0] ?? "").trim();
      if (d) lastDate = d;
      const shift = String(row[shiftIdx] ?? "").trim() || "Shift";
      const slots = [];
      for (let col = shiftIdx + 1; col < h.length && slots.length < 9; col++) {
        slots.push(String(row[col] ?? "").trim());
      }
      while (slots.length < 9) slots.push("");
      out.push({ date: lastDate, shift, slots });
    }
  } else {
    for (const row of body) {
      const d = String(row[0] ?? "").trim();
      if (!d) continue;
      lastDate = d;
      const slots = [];
      for (let col = 1; col < h.length && slots.length < 9; col++) {
        slots.push(String(row[col] ?? "").trim());
      }
      while (slots.length < 9) slots.push("");
      for (let s = 1; s <= 3; s++) {
        out.push({ date: lastDate, shift: `Shift ${s}`, slots: [...slots] });
      }
    }
  }
  return out;
}

function buildSequenceDiffMapFromSchedules(beforeSched, afterSched) {
  const beforeMap = new Map();
  for (const r of beforeSched) {
    beforeMap.set(`${normKey(r.date)}|${normKey(r.shift)}`, r.slots);
  }
  const seqDiffMap = new Map();
  for (const ar of afterSched) {
    const key = `${normKey(ar.date)}|${normKey(ar.shift)}`;
    const bs = beforeMap.get(key);
    if (!bs) continue;
    for (let i = 0; i < 9; i++) {
      if (String(bs[i] ?? "").trim() !== String(ar.slots[i] ?? "").trim()) {
        seqDiffMap.set(`${key}|${i}`, true);
      }
    }
  }
  return seqDiffMap;
}

function sequenceLetterClass(v) {
  const s = String(v || "")
    .trim()
    .toUpperCase();
  if (s === "B") return "seq-token seq-token--b";
  if (s === "K") return "seq-token seq-token--k";
  if (s === "W") return "seq-token seq-token--w";
  return "seq-token seq-token--n";
}

function renderSequenceTable(schedule, seqDiffMap = new Map()) {
  const root = document.getElementById("sequence-grid");
  if (!root) return;
  if (!schedule.length) {
    root.innerHTML = '<p class="events-th-placeholder">No sequence rows found.</p>';
    return;
  }

  const head = `<thead><tr>
    <th class="seq-th seq-th--corner" colspan="2" aria-label="Date and shift columns"></th>
    ${Array.from({ length: 9 })
      .map((_, i) => `<th class="seq-th seq-th--num" scope="col">${i + 1}</th>`)
      .join("")}
  </tr></thead>`;

  let tickCounter = 0;
  const bodyHtml = [];
  let i = 0;
  while (i < schedule.length) {
    const dateLabel = schedule[i].date || "—";
    const dateKey = normKey(String(dateLabel));
    const chunk = [];
    while (i < schedule.length && schedule[i].date === dateLabel) {
      chunk.push(schedule[i]);
      i++;
    }
    chunk.forEach((row, rowIdx) => {
      const shiftKey = normKey(row.shift);
      const shiftCells = row.slots
        .map((v, slotIdx) => {
          const diffKey = `${dateKey}|${shiftKey}|${slotIdx}`;
          const changed = seqDiffMap.get(diffKey) === true;
          const vv = String(v || "")
            .trim()
            .toUpperCase();
          const label = vv === "B" ? "Blue" : vv === "K" ? "Red" : vv === "W" ? "White" : "Unknown";
          const isProd = isSequenceProductLetter(vv);
          if (isProd) tickCounter += 1;
          const ticked = isProd && tickCounter <= sequenceCycleTicks;
          return `<td class="seq-slot-td">
            <span class="seq-cell${changed ? " seq-cell--changed" : ""}${ticked ? " seq-cell--ticked" : ""}${!isProd ? " seq-cell--unavailable" : ""}">
              <span class="${isProd ? sequenceLetterClass(v) : "seq-token seq-token--unavailable"}" title="${escapeHtml(
                isProd ? label : "Unavailable due to supplier delay"
              )}" aria-label="${escapeHtml(isProd ? label : "Unavailable due to supplier delay")}"></span>
              ${changed ? '<span class="seq-changed-dot" title="Changed vs baseline">!</span>' : ""}
              ${ticked ? '<span class="seq-cycle-tick" title="Cycle completed">✓</span>' : ""}
            </span>
          </td>`;
        })
        .join("");
      const dateCell =
        rowIdx === 0
          ? `<td rowspan="${chunk.length}" class="seq-date-cell"><span class="seq-date">${escapeHtml(
              dateLabel
            )}</span></td>`
          : "";
      bodyHtml.push(`<tr class="seq-data-row${rowIdx === chunk.length - 1 ? " seq-data-row--date-last" : ""}">
        ${dateCell}
        <th scope="row" class="seq-shift-label">${escapeHtml(row.shift)}</th>
        ${shiftCells}
      </tr>`);
    });
  }

  root.innerHTML = `<table class="sequence-grid-table" role="grid" aria-label="Production sequence">${head}<tbody>${bodyHtml.join("")}</tbody></table>`;
}

async function loadSequenceInsights() {
  const root = document.getElementById("sequence-grid");
  const statusEl = document.getElementById("sequence-status");
  if (!root) return;

  const setStatus = (msg, isError) => {
    if (!statusEl) return;
    statusEl.hidden = !msg;
    statusEl.textContent = msg;
    statusEl.classList.toggle("table-status--error", Boolean(isError));
  };

  setStatus("Loading sequence…", false);
  try {
    const text = await fetchSequenceCsvText();
    const rows = parseCSV(text);
    const { headers, body } = parseSequenceTable(rows);
    if (headers.length === 0 || body.length === 0) {
      root.innerHTML = '<p class="events-th-placeholder">No sequence rows found.</p>';
      setStatus("No data rows in sequence file.", true);
      return;
    }

    const schedule = normalizeSequenceSchedule(headers, body);
    currentSequenceHeaders = headers;
    currentSequenceBody = body;
    currentSequenceBaseSchedule = schedule;
    currentSequenceSchedule = schedule.map((r) => ({ ...r, slots: [...(r.slots || [])] }));
    sequenceCellTotal = countSequenceProductionSlots(currentSequenceSchedule);
    sequenceCycleTicks = 0;

    let seqDiffMap = new Map();
    if (mpsPhase === "after") {
      const beforeText = await fetchCsvFirstOk(SEQUENCE_BEFORE_CSV_CANDIDATES);
      const beforeRows = parseCSV(beforeText);
      const before = parseSequenceTable(beforeRows);
      const beforeSched = normalizeSequenceSchedule(before.headers, before.body);
      seqDiffMap = buildSequenceDiffMapFromSchedules(beforeSched, currentSequenceSchedule);
    }

    renderSequenceTable(currentSequenceSchedule, seqDiffMap);
    refreshInventorySupplyViewForSequence();
    updateProductionToolbarMetrics();

    setStatus("", false);
  } catch (e) {
    root.innerHTML = '<p class="events-th-placeholder">Could not load sequence.</p>';
    const hint =
      window.location.protocol === "file:"
        ? " Use python -m http.server in this folder, then open http://localhost:8080/index.html"
        : "";
    setStatus((e instanceof Error ? e.message : "Failed to load sequence CSV.") + hint, true);
  }
}

/**
 * Rebuilds `currentSequenceSchedule` from baseline CSV + active Global Supply Outlook delays.
 * Call after loading sequence/MPS or when outlook rows change so Production Planning stays in sync.
 */
async function syncSequenceScheduleWithOutlookImpacts() {
  if (!currentSequenceBaseSchedule.length) return;
  if (!inventoryMatrixHeaders.length) {
    await loadMpsInsights();
  }
  currentSequenceSchedule = applyOutlookImpactsToSequence(currentSequenceBaseSchedule);
  sequenceCellTotal = countSequenceProductionSlots(currentSequenceSchedule);
  sequenceCycleTicks = 0;
  const seqDiffMap = await currentSequenceDiffMap();
  renderSequenceTable(currentSequenceSchedule, seqDiffMap);
  refreshInventorySupplyViewForSequence();
  updateProductionToolbarMetrics();
}

/**
 * Applies Global Supply Outlook delays to the production sequence and refreshes inventory/KPIs.
 * @param {{ skipAlert?: boolean }} [options]
 */
async function applyResequenceFromInventoryButton(options = {}) {
  const skipAlert = options.skipAlert === true;
  if (!currentSequenceBaseSchedule.length) {
    await loadSequenceInsights();
    if (!currentSequenceBaseSchedule.length) return;
  }
  if (!inventoryMatrixHeaders.length) {
    await loadMpsInsights();
  }
  await syncSequenceScheduleWithOutlookImpacts();
  if (!skipAlert) {
    pushAlertFeedItem({
      title: "Production resequence activated",
      meta: `Applied outlook impacts at ${currentCycleLabel()}`,
      impact: "Info",
    });
    renderAlertsPopover();
  }
}

/** Re-sequence: same-color partners (e.g. A1/A3 for A2) fill delayed slots; deferred build moves to slot 9. */
async function applyPartnerResequenceFromButton() {
  if (!currentSequenceBaseSchedule.length) {
    await loadSequenceInsights();
    if (!currentSequenceBaseSchedule.length) return;
  }
  if (!inventoryMatrixHeaders.length) {
    await loadMpsInsights();
  }
  inventoryReplenishedAfterPartnerResequence = true;
  currentSequenceSchedule = applyPartnerResequenceToSequence(currentSequenceBaseSchedule);
  sequenceCellTotal = countSequenceProductionSlots(currentSequenceSchedule);
  sequenceCycleTicks = 0;
  const seqDiffMap = await currentSequenceDiffMap();
  renderSequenceTable(currentSequenceSchedule, seqDiffMap);
  refreshInventorySupplyViewForSequence();
  updateProductionToolbarMetrics();
  pushAlertFeedItem({
    title: "Partner re-sequence applied",
    meta: `3×B / 3×R / 3×W preserved; delayed suppliers’ planned dice move to the last slots in the row.`,
    impact: "Info",
  });
  renderAlertsPopover();
  persistAppUiState();
  renderSupplyInvDisclaimer();
}

function wireInventoryResequenceButton() {
  const btn = document.getElementById("btn-resequence-inventory");
  if (!btn || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    const prev = btn.textContent;
    btn.textContent = "Re-sequencing…";
    try {
      await applyPartnerResequenceFromButton();
    } finally {
      btn.textContent = prev;
      btn.disabled = false;
    }
  });
}

function productionAlertCardHtml(type, title, details, severity) {
  const cls = severity === "high" ? "impact-card impact-card--impact-high" : "impact-card impact-card--impact-low";
  return `<article class="${cls} production-alert-card">
    <div class="production-alert-card__head">
      <span class="impact-badge ${severity === "high" ? "impact-badge--high" : "impact-badge--low"}">${escapeHtml(
    severity === "high" ? "High" : "Low"
  )}</span>
      <h3>${escapeHtml(title)}</h3>
    </div>
    <p class="production-alert-card__type">${escapeHtml(type)}</p>
    <p class="production-alert-card__details">${escapeHtml(details)}</p>
  </article>`;
}

async function loadProductionAlerts() {
  const root = document.getElementById("production-alert-cards");
  if (!root) return;
  if (mpsPhase !== "after") {
    root.innerHTML =
      '<p class="alert-cards-empty">Run Simulate in Supply Chain Control Tower to generate production alerts.</p>';
    return;
  }

  try {
    const [mpsBeforeText, mpsAfterText, seqBeforeText, seqAfterText] = await Promise.all([
      fetchCsvFirstOk(MPS_BEFORE_CSV_CANDIDATES),
      fetchCsvFirstOk(MPS_AFTER_CSV_CANDIDATES),
      fetchCsvFirstOk(SEQUENCE_BEFORE_CSV_CANDIDATES),
      fetchCsvFirstOk(SEQUENCE_AFTER_CSV_CANDIDATES),
    ]);

    const mpsBefore = parseMpsTable(parseCSV(mpsBeforeText));
    const mpsAfter = parseMpsTable(parseCSV(mpsAfterText));
    const seqBefore = parseSequenceTable(parseCSV(seqBeforeText));
    const seqAfter = parseSequenceTable(parseCSV(seqAfterText));

    const cards = [];

    const mpsCols = mpsAfter.headers.slice(1);
    const mpsBeforeMap = mpsRowMapBySeat(mpsBefore.headers, mpsBefore.body);
    const mpsAfterMap = mpsRowMapBySeat(mpsAfter.headers, mpsAfter.body);
    for (const seatKey of ["blackseat", "whiteseat", "blueseat"]) {
      const bRow = mpsBeforeMap.get(seatKey);
      const aRow = mpsAfterMap.get(seatKey);
      if (!bRow || !aRow) continue;
      let largestDrop = { day: "", delta: 0 };
      for (let i = 1; i < aRow.length; i++) {
        const b = Number(String(bRow[i] || "").trim());
        const a = Number(String(aRow[i] || "").trim());
        if (!Number.isFinite(b) || !Number.isFinite(a)) continue;
        const d = a - b;
        if (d < largestDrop.delta) largestDrop = { day: mpsCols[i - 1] || `Col ${i}`, delta: d };
      }
      if (largestDrop.delta < 0) {
        const seatName = seatKey === "blackseat" ? "Black Seat" : seatKey === "whiteseat" ? "White Seat" : "Blue Seat";
        cards.push(
          productionAlertCardHtml(
            "MPS Capacity Shift",
            `${seatName} output drops on ${largestDrop.day}`,
            `Planned units reduced by ${Math.abs(largestDrop.delta)} versus baseline.`,
            Math.abs(largestDrop.delta) >= 2 ? "high" : "low"
          )
        );
      }
    }

    const seqBeforeSched = normalizeSequenceSchedule(seqBefore.headers, seqBefore.body);
    const seqAfterSched = normalizeSequenceSchedule(seqAfter.headers, seqAfter.body);
    const beforeSeqMap = new Map();
    for (const r of seqBeforeSched) {
      beforeSeqMap.set(`${normKey(r.date)}|${normKey(r.shift)}`, r.slots);
    }
    const changedByDate = new Map();
    for (const ar of seqAfterSched) {
      const key = `${normKey(ar.date)}|${normKey(ar.shift)}`;
      const bs = beforeSeqMap.get(key);
      if (!bs) continue;
      let changedSlots = 0;
      for (let i = 0; i < 9; i++) {
        if (String(bs[i] ?? "").trim() !== String(ar.slots[i] ?? "").trim()) changedSlots++;
      }
      if (changedSlots > 0) {
        const prev = changedByDate.get(ar.date) || 0;
        changedByDate.set(ar.date, prev + changedSlots);
      }
    }
    for (const [dateLabel, changedSlots] of changedByDate.entries()) {
      cards.push(
        productionAlertCardHtml(
          "Sequence Resequencing",
          `${dateLabel} has ${changedSlots} sequence slot changes`,
          "Sequence pattern diverges from baseline to absorb supply disruption.",
          changedSlots >= 4 ? "high" : "low"
        )
      );
    }

    if (cards.length === 0) {
      root.innerHTML = '<p class="alert-cards-empty">No production deltas detected versus baseline.</p>';
      return;
    }

    root.innerHTML = cards.join("");
  } catch (e) {
    root.innerHTML = `<p class="alert-cards-empty">${escapeHtml(
      e instanceof Error ? e.message : "Could not load production alert datasets."
    )}</p>`;
  }
}

async function loadProductionPlanningInsights() {
  await Promise.all([loadMpsInsights(), loadSequenceInsights(), loadProductionAlerts()]);
  await syncSequenceScheduleWithOutlookImpacts();
}

async function currentSequenceDiffMap() {
  if (mpsPhase !== "after" || !currentSequenceSchedule.length) {
    return new Map();
  }
  try {
    const beforeText = await fetchCsvFirstOk(SEQUENCE_BEFORE_CSV_CANDIDATES);
    const beforeRows = parseCSV(beforeText);
    const before = parseSequenceTable(beforeRows);
    const beforeSched = normalizeSequenceSchedule(before.headers, before.body);
    return buildSequenceDiffMapFromSchedules(beforeSched, currentSequenceSchedule);
  } catch {
    return new Map();
  }
}

/**
 * One Next Cycle step: 500ms wait, then tick + render + inventory refresh.
 * @returns {"ok" | "aborted"}
 */
async function runSingleNextCycleTickCore(options = {}) {
  const interruptCheck = options.interruptCheck;
  const delayMs = Number.isFinite(options.delayMs) ? Number(options.delayMs) : 500;
  if (typeof interruptCheck === "function") {
    if (!(await waitMsInterruptible(delayMs, interruptCheck))) return "aborted";
  } else {
    await waitMs(delayMs);
  }
  if (sequenceCycleTicks < sequenceCellTotal) {
    sequenceCycleTicks += 1;
  }
  const seqDiffMap = await currentSequenceDiffMap();
  renderSequenceTable(currentSequenceSchedule, seqDiffMap);
  refreshInventorySupplyViewForSequence();
  updateProductionToolbarMetrics();
  return "ok";
}

function updateProductionToolbarMetrics() {
  const dateEl = document.getElementById("production-kpi-date");
  const cycleEl = document.getElementById("production-kpi-cycle");
  const riskEl = document.getElementById("production-kpi-risk");
  const curShiftIdx = currentSequenceShiftIndex();
  if (riskEl) riskEl.textContent = String(activeDelayedSupplierCount(curShiftIdx));
  if (!dateEl || !cycleEl) return;
  if (!currentSequenceSchedule.length || sequenceCellTotal <= 0) {
    dateEl.textContent = "--";
    cycleEl.textContent = "--";
    return;
  }
  const cycleInfo = currentCycleInfo();
  const currentDate = String(cycleInfo.date || "--").replace(/-/g, " ");
  dateEl.textContent = currentDate;
  cycleEl.textContent = currentCycleLabel(cycleInfo);
}

function currentCycleInfo() {
  if (!currentSequenceSchedule.length || sequenceCellTotal <= 0) {
    return { date: "--", shift: "Shift --", slot: 1 };
  }
  const rowCount = currentSequenceSchedule.length;
  let remaining = sequenceCycleTicks;
  let currentRow = 0;
  let slotProgress = 1;
  for (let r = 0; r < rowCount; r++) {
    const cap = (currentSequenceSchedule[r].slots || []).filter((s) => isSequenceProductLetter(s)).length;
    const rowCap = Math.max(1, cap);
    if (remaining < rowCap) {
      currentRow = r;
      slotProgress = remaining + 1;
      break;
    }
    remaining -= rowCap;
    if (r === rowCount - 1) {
      currentRow = r;
      slotProgress = rowCap;
    }
  }
  const row = currentSequenceSchedule[currentRow] || { date: "--", shift: "Shift --" };
  return { date: String(row.date || "--"), shift: String(row.shift || "Shift --"), slot: slotProgress };
}

function currentCycleLabel(info = currentCycleInfo()) {
  return `${info.shift} - Slot ${info.slot}`;
}

function compactIsoForFilename(d = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(
    d.getMinutes()
  )}${pad(d.getSeconds())}`;
}

function sequenceSnapshotPayload() {
  return {
    generatedAt: new Date().toISOString(),
    source: "Re-sequence Production",
    totalProductionSlots: sequenceCellTotal,
    consumedTicks: sequenceCycleTicks,
    currentCycle: currentCycleLabel(),
    schedule: currentSequenceSchedule.map((row) => ({
      date: String(row.date || ""),
      shift: String(row.shift || ""),
      slots: Array.isArray(row.slots) ? [...row.slots] : [],
    })),
  };
}

async function saveSequenceSnapshotJson() {
  const payload = sequenceSnapshotPayload();
  const json = JSON.stringify(payload, null, 2);
  const filename = `sequence_snapshot_${compactIsoForFilename()}.json`;

  // Preferred: save directly with picker (Chromium-based browsers)
  if ("showSaveFilePicker" in window) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{ description: "JSON Files", accept: { "application/json": [".json"] } }],
      });
      const writable = await handle.createWritable();
      await writable.write(json);
      await writable.close();
      return;
    } catch {
      // Fallback to browser download below.
    }
  }

  // Fallback: trigger a local download.
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function wireExportSequenceJsonButton() {
  const btn = document.getElementById("btn-export-sequence-json");
  if (!btn || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    const prev = btn.textContent;
    btn.textContent = "Exporting…";
    try {
      if (!currentSequenceSchedule.length) {
        await loadSequenceInsights();
      }
      if (!currentSequenceSchedule.length) return;
      await saveSequenceSnapshotJson();
      pushAlertFeedItem({
        title: "Sequence JSON exported",
        meta: `Snapshot exported at ${currentCycleLabel()}`,
        impact: "Info",
      });
      renderAlertsPopover();
    } finally {
      btn.textContent = prev;
      btn.disabled = false;
    }
  });
}

function syncLiveProductionRunControls() {
  const liveBtn = document.getElementById("btn-live-production-run");
  const nextBtn = document.getElementById("btn-next-cycle-sequence");
  if (liveBtn) {
    liveBtn.setAttribute("aria-pressed", liveProductionRunActive ? "true" : "false");
    liveBtn.classList.toggle("production-tab-toolbar__live-run--active", liveProductionRunActive);
  }
  if (nextBtn) {
    nextBtn.disabled = liveProductionRunActive;
  }
  updateProductionToolbarMetrics();
}

async function runLiveProductionLoop() {
  while (liveProductionRunActive) {
    if (!currentSequenceSchedule.length) {
      await loadSequenceInsights();
      if (!liveProductionRunActive) break;
      continue;
    }
    if (sequenceCellTotal <= 0) {
      liveProductionRunActive = false;
      break;
    }

    if (sequenceCycleTicks >= sequenceCellTotal) {
      sequenceCycleTicks = 0;
      const seqDiffMap = await currentSequenceDiffMap();
      renderSequenceTable(currentSequenceSchedule, seqDiffMap);
      refreshInventorySupplyViewForSequence();
      updateProductionToolbarMetrics();
      if (!(await waitMsInterruptible(1000, () => !liveProductionRunActive))) break;
      continue;
    }

    const step = await runSingleNextCycleTickCore({ delayMs: 1000, interruptCheck: () => !liveProductionRunActive });
    if (step === "aborted") break;
  }
  syncLiveProductionRunControls();
}

function wireLiveProductionRunButton() {
  const btn = document.getElementById("btn-live-production-run");
  if (!btn || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";
  btn.addEventListener("click", () => {
    liveProductionRunActive = !liveProductionRunActive;
    syncLiveProductionRunControls();
    if (liveProductionRunActive) {
      void runLiveProductionLoop();
    }
  });
  syncLiveProductionRunControls();
}

function wireSequenceNextCycleButton() {
  const btn = document.getElementById("btn-next-cycle-sequence");
  if (!btn || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";
  const loadingHtml =
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin" aria-hidden="true"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg> Working…';

  btn.addEventListener("click", async () => {
    if (liveProductionRunActive) return;
    if (!currentSequenceSchedule.length) {
      btn.disabled = true;
      btn.setAttribute("aria-busy", "true");
      const prev = btn.innerHTML;
      btn.innerHTML = loadingHtml;
      try {
        await loadSequenceInsights();
      } finally {
        btn.innerHTML = prev;
        btn.removeAttribute("aria-busy");
        btn.disabled = false;
      }
      return;
    }
    if (sequenceCellTotal <= 0) return;
    if (sequenceCycleTicks >= sequenceCellTotal) return;

    btn.disabled = true;
    btn.setAttribute("aria-busy", "true");
    const prev = btn.innerHTML;
    btn.innerHTML = loadingHtml;
    try {
      await runSingleNextCycleTickCore();
    } finally {
      btn.innerHTML = prev;
      btn.removeAttribute("aria-busy");
      btn.disabled = false;
    }
  });
}

function wireImpactAlertSimulate() {
  const btn = document.getElementById("btn-simulate-alerts");
  if (!btn || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    const prev = btn.textContent;
    btn.textContent = "Simulating…";
    supplierCommitmentsPhase = "after";
    inventoryInsightsPhase = "after";
    mpsPhase = "after";
    persistAppUiState();
    await loadSupplierCommitments();
    const invPanel = document.getElementById("panel-inventory");
    if (invPanel?.dataset.inventoryLoaded === "1") {
      await loadInventoryInsights();
    }
    const prodPanel = document.getElementById("panel-production");
    if (prodPanel?.dataset.mpsLoaded === "1") {
      await loadProductionPlanningInsights();
    }
    renderSupplyInvDisclaimer();
    window.setTimeout(() => {
      btn.disabled = false;
      btn.textContent = prev;
    }, 1100);
  });
}

function wireOutlookAnalyzeButton() {
  const btn = document.getElementById("btn-analyze-outlook");
  const noteInput = document.getElementById("outlook-notes-input");
  if (!btn || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    const prev = btn.textContent;
    btn.textContent = "Analyzing…";
    renderOutlookAlertsIdle("Analyzing impact rows…");
    await waitMs(1000);
    if (!cachedOutlookHeaders.length) {
      renderOutlookAlertsIdle("Load Global Supply Outlook first, then click Analyze.");
      btn.disabled = false;
      btn.textContent = prev;
      return;
    }
    const note = String(noteInput?.value || "").trim();
    if (!note) {
      renderOutlookAlertsIdle("Enter an event note first, then click Analyze.");
      btn.disabled = false;
      btn.textContent = prev;
      return;
    }

    inventoryReplenishedAfterPartnerResequence = false;

    const generated = generatedOutlookRowFromEvent(cachedOutlookHeaders, note);
    cachedOutlookRows.push(generated);
    const alertItem = alertItemsFromOutlookRows(cachedOutlookHeaders, [generated])[0];
    if (alertItem) {
      pushAlertFeedItem({
        title: `Global event: ${alertItem.event}`,
        meta: `${alertItem.supplier} | ${alertItem.delay}`,
        impact: alertItem.impact,
      });
    }
    const tbody = document.getElementById("global-supply-outlook-tbody");
    if (tbody) {
      tbody.innerHTML = cachedOutlookRows.map((cells) => outlookRowHtml(cachedOutlookHeaders, cells)).join("");
    }
    const countEl = document.getElementById("global-supply-active-count");
    if (countEl) countEl.textContent = `(${cachedOutlookRows.length})`;

    // Re-render SCC downstream panels so supplier status + map risk colors
    // reflect newly generated outlook impacts immediately.
    await loadSupplierCommitments();

    await applyResequenceFromInventoryButton({ skipAlert: true });
    if (!currentSequenceBaseSchedule.length) {
      refreshInventorySupplyViewForSequence();
      updateProductionToolbarMetrics();
    }

    renderOutlookImpactAlerts(cachedOutlookHeaders, cachedOutlookRows);
    renderAlertsPopover();
    hasAnalyzedOutlook = true;
    persistAppUiState();
    renderSupplyInvDisclaimer();
    btn.disabled = false;
    btn.textContent = prev;
  });
}

function initGlobalOutlookEmptyState() {
  const theadEl = document.getElementById("global-supply-outlook-thead");
  const tbody = document.getElementById("global-supply-outlook-tbody");
  const countEl = document.getElementById("global-supply-active-count");
  const statusEl = document.getElementById("global-supply-status");
  if (theadEl) {
    theadEl.innerHTML =
      '<tr><th class="events-th-placeholder">No events loaded yet.</th></tr>';
  }
  if (tbody) tbody.innerHTML = "";
  if (countEl) countEl.textContent = "(0)";
  if (statusEl) {
    statusEl.hidden = true;
    statusEl.textContent = "";
    statusEl.classList.remove("table-status--error");
  }
  alertsUnreadCount = 0;
  alertsFeed = [];
  alertsPanelOpen = false;
  inventoryReplenishedAfterPartnerResequence = false;
  renderAlertsPopover();
  setAlertsPanelOpen(false);
  renderOutlookAlertsIdle("No impact analysis yet. Click Analyze in Global Supply Outlook.");
  renderSupplyInvDisclaimer();
}

async function hardReloadAppBypassingCache() {
  try {
    if ("caches" in window) {
      const keys = await caches.keys();
      await Promise.all(keys.map((k) => caches.delete(k)));
    }
  } catch {
    /* ignore cache API failures */
  }
  const url = new URL(window.location.href);
  url.searchParams.set("_hard", Date.now().toString());
  window.location.replace(url.toString());
}

function wireResetButton() {
  const btn = document.getElementById("btn-reset-scenarios");
  if (!btn || btn.dataset.wired === "1") return;
  btn.dataset.wired = "1";
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    const prev = btn.textContent;
    btn.textContent = "Resetting…";
    liveProductionRunActive = false;
    syncLiveProductionRunControls();
    supplierCommitmentsPhase = "before";
    inventoryInsightsPhase = "before";
    mpsPhase = "before";
    hasAnalyzedOutlook = false;
    cachedOutlookHeaders = [];
    cachedOutlookRows = [];
    persistAppUiState();
    await hardReloadAppBypassingCache();
  });
}

function wireCrossTabStateSync() {
  if ("BroadcastChannel" in window) {
    try {
      uiSyncChannel = new BroadcastChannel(APP_UI_SYNC_CHANNEL);
      uiSyncChannel.addEventListener("message", (e) => {
        const data = e.data || {};
        if (data.type !== "ui-state-updated") return;
        reconcileUiStateFromStorage();
      });
    } catch {
      uiSyncChannel = null;
    }
  }

  window.addEventListener("storage", async (e) => {
    if (e.key !== APP_UI_STATE_KEY || !e.newValue) return;
    await reconcileUiStateFromStorage();
  });

  window.addEventListener("focus", () => {
    reconcileUiStateFromStorage();
  });

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      reconcileUiStateFromStorage();
    }
  });

  window.addEventListener("pageshow", () => {
    reconcileUiStateFromStorage();
  });
}

function attachRefreshButton(button, loader) {
  button?.addEventListener("click", async () => {
    button.disabled = true;
    const prev = button.innerHTML;
    button.innerHTML =
      '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg> Refreshing…';
    try {
      await loader();
    } finally {
      window.setTimeout(() => {
        button.disabled = false;
        button.innerHTML = prev;
      }, 400);
    }
  });
}

attachRefreshButton(document.getElementById("btn-refresh-supplier"), loadSupplierCommitments);
attachRefreshButton(document.getElementById("btn-refresh-outlook"), loadGlobalSupplyOutlook);
attachRefreshButton(document.getElementById("btn-refresh-inventory"), loadInventoryInsights);

function wireAppTabs() {
  const tablist = document.querySelector('[role="tablist"]');
  if (!tablist || tablist.dataset.wired === "1") return;
  tablist.dataset.wired = "1";
  const tabs = [...tablist.querySelectorAll('[role="tab"]')];

  function selectTab(tab) {
    const panelId = tab.getAttribute("aria-controls");
    if (!panelId) return;
    tabs.forEach((t) => {
      const sel = t === tab;
      t.setAttribute("aria-selected", sel ? "true" : "false");
      t.classList.toggle("app-tab--active", sel);
      t.tabIndex = sel ? 0 : -1;
    });
    document.querySelectorAll('[role="tabpanel"]').forEach((p) => {
      const show = p.id === panelId;
      p.hidden = !show;
      p.classList.toggle("app-tab-panel--active", show);
    });
    if (panelId === "panel-scc") scheduleShipmentLeafletResize();
    if (panelId === "panel-inventory") {
      const invPanel = document.getElementById("panel-inventory");
      if (invPanel?.dataset.inventoryLoaded !== "1") {
        invPanel.dataset.inventoryLoaded = "1";
        loadInventoryInsights();
      }
    }
    if (panelId === "panel-production") {
      const prodPanel = document.getElementById("panel-production");
      if (prodPanel?.dataset.mpsLoaded !== "1") {
        prodPanel.dataset.mpsLoaded = "1";
        loadProductionPlanningInsights();
      }
    }
  }

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => selectTab(tab));
  });

  tablist.addEventListener("keydown", (e) => {
    const i = tabs.indexOf(document.activeElement);
    if (i < 0) return;
    if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
      e.preventDefault();
      const next =
        e.key === "ArrowRight" ? (i + 1) % tabs.length : (i - 1 + tabs.length) % tabs.length;
      tabs[next].focus();
      selectTab(tabs[next]);
    } else if (e.key === "Home") {
      e.preventDefault();
      tabs[0].focus();
      selectTab(tabs[0]);
    } else if (e.key === "End") {
      e.preventDefault();
      tabs[tabs.length - 1].focus();
      selectTab(tabs[tabs.length - 1]);
    }
  });
}

function enforceSccPanelCardOrder() {
  const grid = document.querySelector("#panel-scc main.grid");
  if (!grid) return;
  const events = grid.querySelector(".card--events");
  const commitments = grid.querySelector(".card--commitments");
  const map = grid.querySelector(".card--map");
  if (!events || !commitments || !map) return;
  grid.append(events, commitments, map);
}

hydrateAppUiState();
lastAppliedUiStateSig = appUiStateSignature();
wireAppTabs();
enforceSccPanelCardOrder();
wireShipmentSupplierSearch();
if (typeof window.L !== "undefined") {
  initShipmentLeafletMap();
} else {
  window.addEventListener("load", () => initShipmentLeafletMap(), { once: true });
}
wireCrossTabStateSync();
wireResetButton();
wireAlertsPopover();
initGlobalOutlookEmptyState();
loadGlobalSupplyOutlook();
wireOutlookAnalyzeButton();
wireImpactAlertSimulate();
wireInventoryResequenceButton();
wireExportSequenceJsonButton();
wireLiveProductionRunButton();
wireSequenceNextCycleButton();
updateProductionToolbarMetrics();
void loadSupplierCommitments().then(() => renderSupplyInvDisclaimer());
