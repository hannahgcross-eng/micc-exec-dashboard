/* =====================================================================
   MICC USA Playbook Executive & Distributor Dashboard
   Pure client-side: SheetJS + Chart.js
   ===================================================================== */

/* ---------- 1. Field aliases (Playbook export tolerance) ---------- */
const ALIAS = {
  // Missed Opportunity fields
  client:        ['Client Name','Distributor','Client'],
  outlet:        ['Outlet','Outlet Name','Store'],
  outletCode:    ['Outlet Code','Store Code'],
  product:       ['Product','SKU','Product Name'],
  brand:         ['Brand'],
  asset:         ['Asset Serial Number','Asset Serial #','Asset Serial','Asset'],
  casesPerFacing:['Cases Per Facing'],
  casePrice:     ['Case Price'],
  unitPrice:     ['Unit Price'],
  caseSize:      ['Case Size'],
  inStockDays:   ['Total InStock Days','InStock Days'],
  oosDays:       ['OOS Days','Out of Stock Days'],
  casesSold:     ['Total Cases Sold'],
  unitsSold:     ['Total Units Sold'],
  timesRefilled: ['Times Refilled'],
  avgDailyCases: ['Avg Daily Cases Sold'],
  avgDailyUnits: ['Avg Daily Units Sold'],
  missedUnits:   ['Missed Opportunity Units','Missed Units'],
  missedRev:     ['Missed Revenue'],
  actualRev:     ['Actual Revenue'],
  potentialRev:  ['Potential Revenue'],
  planogram:     ['Planogram'],
  salesRep:      ['Primary Sales Rep','Sales Rep'],
  market:        ['Market Name','Market'],
  channel:       ['Channel Name','Channel'],
  classification:['Classification'],
  route:         ['Route'],
  // Asset Performance fields
  cabinetType:   ['Cabinet Type','Camera Type','Hardware Type'],
  city:          ['City'],
  street:        ['Street'],
  deviceNumber:  ['Device Number'],
  oosPct:        ['Out of Stock (%)','OOS %','Empty %'],
  facings:       ['Recognized Facings','Total Facings','Facings'],
  installedOn:   ['Installed On','Install Date'],
  devicePing:    ['Device Ping','Last Device Ping'],
  doorClose:     ['Door Close','Last Door Close'],
  imagesPct7:    ['Percent Images (7-Days)','Image %'],
  missingImg7:   ['Missing Images (7-Days)'],
  diagnosis:     ['Diagnosis'],
  status:        ['Status','Asset Status'],
  lastImg:       ['Last Image Received','Last Image'],
  lastImgProv:   ['Last Image Provisioned'],
  gps:           ['GPS'],
  latitude:      ['Latitude','Lat'],
  longitude:     ['Longitude','Lon','Long','Lng'],
  temperature:   ['Temperature'],
  emptySoS:      ['Empty SoS','Empty Share of Shelf','Empty %'],
  foreignSoS:    ['Foreign SoS','Foreign Share of Shelf','Foreign %'],
  hardwareStatus:['Hardware Status'],
};

const FIELD_KIND = {
  MISSED_OPP: 'missed_opp',
  ASSET_PERF: 'asset_perf',
  UNKNOWN: 'unknown',
};

const REQ_MO = ['client','outlet','outletCode','product','asset','missedRev','actualRev','potentialRev','oosDays'];
const REQ_AP = ['client','outlet','asset','emptySoS','foreignSoS','status'];

// Normalize a header string for fuzzy matching: lowercase, strip punctuation/spaces/units
function normalizeHeader(h) {
  return String(h || '')
    .toLowerCase()
    .replace(/\(.*?\)/g, '')      // strip parenthetical units like "(%)" or "(7-Days)"
    .replace(/[^a-z0-9]/g, '')    // strip all non-alphanumerics (spaces, underscores, hyphens, slashes)
    .trim();
}

function aliasMatch(headers, key) {
  const aliases = ALIAS[key] || [];
  const normalizedHeaders = headers.map(normalizeHeader);
  for (const a of aliases) {
    const normA = normalizeHeader(a);
    const idx = normalizedHeaders.findIndex(h => h === normA);
    if (idx >= 0) return idx;
  }
  // Second pass: substring match for very forgiving fallback
  for (const a of aliases) {
    const normA = normalizeHeader(a);
    if (normA.length < 4) continue; // too short to substring-match safely
    const idx = normalizedHeaders.findIndex(h => h && (h.includes(normA) || normA.includes(h)));
    if (idx >= 0) return idx;
  }
  return -1;
}

function buildIndex(headers) {
  const idx = {};
  for (const k in ALIAS) idx[k] = aliasMatch(headers, k);
  return idx;
}

function detectKind(idx) {
  const moHits = REQ_MO.filter(k => idx[k] >= 0).length;
  const apHits = REQ_AP.filter(k => idx[k] >= 0).length;
  // Use whichever has the higher match score; require at least 50% match
  const moScore = moHits / REQ_MO.length;
  const apScore = apHits / REQ_AP.length;
  if (moScore < 0.5 && apScore < 0.5) return FIELD_KIND.UNKNOWN;
  return moScore >= apScore ? FIELD_KIND.MISSED_OPP : FIELD_KIND.ASSET_PERF;
}

/* ---------- 2. Excel date helper ---------- */
function excelToDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    const ms = (v - 25569) * 86400 * 1000;
    return new Date(ms);
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function fmtDate(d) {
  if (!d) return '—';
  const dd = (d instanceof Date) ? d : excelToDate(d);
  if (!dd) return String(d);
  return dd.toISOString().slice(0,10);
}
function daysBetween(a, b) {
  if (!a || !b) return null;
  const ms = b.getTime() - a.getTime();
  return Math.round(ms / 86400000);
}

/* ---------- 3. Number helpers ---------- */
const num = v => {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g,''));
  return isFinite(n) ? n : 0;
};
const fmtMoney = (v, dec=0) => {
  if (v == null || !isFinite(v)) return '—';
  const n = num(v);
  const abs = Math.abs(n);
  if (abs >= 1e6) return `$${(n/1e6).toFixed(dec===0?1:dec)}M`;
  if (abs >= 1e3) return `$${(n/1e3).toFixed(dec===0?1:dec)}K`;
  return `$${n.toFixed(dec)}`;
};
const fmtMoneyFull = v => {
  if (v == null || !isFinite(v)) return '—';
  return '$' + num(v).toLocaleString(undefined, {maximumFractionDigits:2});
};
const fmtPct = (v, dec=1) => {
  if (v == null || !isFinite(v)) return '—';
  return num(v).toFixed(dec) + '%';
};
const fmtNum = (v, dec=0) => {
  if (v == null || !isFinite(v)) return '—';
  return num(v).toLocaleString(undefined, {maximumFractionDigits:dec, minimumFractionDigits:dec});
};
const stripPrefix = s => {
  if (!s) return s;
  return String(s).replace(/^MICC USA\s+/i,'').trim();
};

/* ---------- 3b. CSV export utilities ---------- */
function tableToCSV(tableEl) {
  const escape = s => `"${String(s ?? '').replace(/"/g, '""').trim()}"`;
  const rows = [];
  const headers = [...tableEl.querySelectorAll('thead th')].map(th => escape(th.textContent));
  if (headers.length) rows.push(headers.join(','));
  tableEl.querySelectorAll('tbody tr').forEach(tr => {
    if (tr.classList.contains('expanded-row')) return; // skip expand panels
    const cells = [...tr.querySelectorAll('td')].map(td => escape(td.textContent));
    if (cells.length) rows.push(cells.join(','));
  });
  return rows.join('\r\n');
}

function downloadCSV(csv, filename) {
  const bom = '﻿'; // UTF-8 BOM so Excel opens it correctly
  const blob = new Blob([bom + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.style.display = 'none';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function safeFilename(s) {
  return String(s || 'export').replace(/[^a-z0-9\-_ ]/gi, '').trim().replace(/\s+/g, '_').toLowerCase() || 'export';
}

function exportBtn(label, getTableFn, filename) {
  const btn = document.createElement('button');
  btn.className = 'tbl-export-btn';
  btn.innerHTML = `⬇ ${label}`;
  btn.addEventListener('click', e => {
    e.stopPropagation();
    const tbl = typeof getTableFn === 'function' ? getTableFn() : getTableFn;
    if (!tbl) return;
    downloadCSV(tableToCSV(tbl), safeFilename(filename) + '.csv');
  });
  return btn;
}

/* ---------- 4. Global state ---------- */
const state = {
  missedOpp: [],
  assetPerf: [],
  periods: {},
  primaryPeriod: 'preset:mom',
  comparisonPeriod: 'auto',
  // Computed from primaryPeriod / comparisonPeriod each render cycle
  allMonths: [],
  primaryMonths: [],
  comparisonMonths: [],
  customPrimaryMonths: [],
  customCompareMonths: [],
  filters: {
    distributor: 'all',
    market: 'all',
    channel: 'all',
    classification: 'all',
    route: 'all',
    outlet: 'all',
    product: 'all',
    cameraType: 'all',
    assetStatus: 'all',
  },
  view: 'executive',
  isSample: false,
  charts: {},
  assetLeafletMap: null,
  currentAssets: [],
  outletLeafletMap: null,
  currentOutletList: [],
  thresholds: {
    moq: 0,
    mov: 250,
    routeMov: 1000,
  },
};

/* ---------- 5. Normalize one row ---------- */
function normMORow(row, idx, monthLabel) {
  return {
    monthLabel,
    distributor: stripPrefix(row[idx.client]),
    distributorRaw: row[idx.client],
    outlet: row[idx.outlet],
    outletCode: row[idx.outletCode],
    product: row[idx.product],
    brand: row[idx.brand] || '',
    asset: row[idx.asset],
    casesPerFacing: num(row[idx.casesPerFacing]),
    casePrice: num(row[idx.casePrice]),
    unitPrice: num(row[idx.unitPrice]),
    caseSize: num(row[idx.caseSize]),
    inStockDays: num(row[idx.inStockDays]),
    oosDays: num(row[idx.oosDays]),
    casesSold: num(row[idx.casesSold]),
    unitsSold: num(row[idx.unitsSold]),
    timesRefilled: num(row[idx.timesRefilled]),
    avgDailyCases: num(row[idx.avgDailyCases]),
    avgDailyUnits: num(row[idx.avgDailyUnits]),
    missedUnits: num(row[idx.missedUnits]),
    missedRev: num(row[idx.missedRev]),
    actualRev: num(row[idx.actualRev]),
    potentialRev: num(row[idx.potentialRev]),
    planogram: row[idx.planogram] || '',
    salesRep: row[idx.salesRep] || '',
    market: row[idx.market] || '',
    channel: row[idx.channel] || '',
    classification: row[idx.classification] || '',
    route: row[idx.route] || '',
  };
}
function parseGPS(row, idx) {
  const tryLatLon = (a, b) => {
    const lat = parseFloat(a), lon = parseFloat(b);
    if (isFinite(lat) && isFinite(lon) && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180 && (lat !== 0 || lon !== 0))
      return { lat, lon };
    return null;
  };
  // Try combined GPS field ("lat,lon" or "lat lon")
  const gpsRaw = idx.gps >= 0 ? String(row[idx.gps] ?? '').trim() : '';
  if (gpsRaw) {
    const parts = gpsRaw.split(/[\s,]+/);
    if (parts.length >= 2) { const r = tryLatLon(parts[0], parts[1]); if (r) return r; }
  }
  // Try separate Latitude / Longitude columns
  if (idx.latitude >= 0 && idx.longitude >= 0) {
    const r = tryLatLon(row[idx.latitude], row[idx.longitude]);
    if (r) return r;
  }
  return { lat: null, lon: null };
}

function normAPRow(row, idx, dailyImgCols) {
  const status = row[idx.status] || '';
  const diagnosis = row[idx.diagnosis] || '';
  const lastImg = excelToDate(row[idx.lastImg]) || excelToDate(row[idx.lastImgProv]);
  const ping = excelToDate(row[idx.devicePing]);
  return {
    distributor: stripPrefix(row[idx.client]),
    distributorRaw: row[idx.client],
    asset: row[idx.asset],
    cabinetType: row[idx.cabinetType] || '',
    outlet: row[idx.outlet],
    outletCode: row[idx.outletCode],
    city: row[idx.city] || '',
    street: row[idx.street] || '',
    classification: row[idx.classification] || '',
    deviceNumber: row[idx.deviceNumber] || '',
    oosPct: num(row[idx.oosPct]),
    facings: num(row[idx.facings]),
    installedOn: excelToDate(row[idx.installedOn]),
    devicePing: ping,
    doorClose: excelToDate(row[idx.doorClose]),
    imagesPct7: num(row[idx.imagesPct7]),
    missingImg7: num(row[idx.missingImg7]),
    diagnosis: diagnosis,
    status: status,
    lastImg: lastImg,
    ...parseGPS(row, idx),
    temperature: num(row[idx.temperature]),
    emptySoS: num(row[idx.emptySoS]),
    foreignSoS: num(row[idx.foreignSoS]),
    hardwareStatus: row[idx.hardwareStatus] || '',
    dailyImg: dailyImgCols.map(c => ({ date: c.date, doors: num(row[c.doorIdx]), images: num(row[c.imgIdx]) })),
    noCommunication: status.toLowerCase().includes('no communication') || hardwareIsNoComm(diagnosis),
    noImage: status.toLowerCase().includes('no image') || diagnosis.toLowerCase().includes('no image'),
    noDoor: status.toLowerCase().includes('no door') || diagnosis.toLowerCase().includes('no door'),
  };
}
function hardwareIsNoComm(d) {
  const s = String(d || '').toLowerCase();
  return s.includes('no communication') || s.includes('offline');
}

/* ---------- 6. Detect daily image columns ---------- */
function findDailyImgCols(headers) {
  const out = [];
  const doorRe = /^(\d{1,2}[-\s]?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec))\s+Doors$/i;
  const imgRe  = /^(\d{1,2}[-\s]?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec))\s+Images$/i;
  const doors = {}, imgs = {};
  headers.forEach((h, i) => {
    const hs = String(h || '').trim();
    let m = doorRe.exec(hs); if (m) { doors[m[1]] = i; return; }
    m = imgRe.exec(hs); if (m) { imgs[m[1]] = i; }
  });
  for (const date of Object.keys(imgs)) {
    out.push({ date, doorIdx: doors[date], imgIdx: imgs[date] });
  }
  return out;
}

/* ---------- 7. Filter helpers ---------- */
function rowsInScope(rows, f) {
  return rows.filter(r => {
    if (f.distributor !== 'all' && r.distributor !== f.distributor) return false;
    if (f.market !== 'all' && r.market !== f.market) return false;
    if (f.channel !== 'all' && r.channel !== f.channel) return false;
    if (f.classification !== 'all' && r.classification !== f.classification) return false;
    if (f.route !== 'all' && String(r.route) !== String(f.route)) return false;
    if (f.outlet !== 'all' && r.outlet !== f.outlet) return false;
    if (f.product !== 'all' && r.product !== f.product) return false;
    return true;
  });
}
function assetsInScope(rows, f) {
  return rows.filter(r => {
    if (f.distributor !== 'all' && r.distributor !== f.distributor) return false;
    if (f.classification !== 'all' && r.classification !== f.classification) return false;
    if (f.outlet !== 'all' && r.outlet !== f.outlet) return false;
    if (f.cameraType !== 'all' && r.cabinetType !== f.cameraType) return false;
    if (f.assetStatus !== 'all') {
      if (f.assetStatus === 'No Communication' && !r.noCommunication) return false;
      else if (f.assetStatus === 'No Image' && !r.noImage) return false;
      else if (f.assetStatus === 'No Door' && !r.noDoor) return false;
      else if (f.assetStatus === 'OK' && (r.noCommunication || r.noImage || r.noDoor)) return false;
    }
    return true;
  });
}

/* ---------- 7b. Period expansion helpers ---------- */
const MONTH_IDX = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};

function expandPeriod(value, allMonths) {
  if (!value || value === 'none' || !allMonths.length) return [];
  if (value === 'preset:mom') {
    return allMonths.length ? [allMonths[allMonths.length - 1]] : [];
  }
  if (value === 'preset:ytd') {
    const latestYear = allMonths[allMonths.length - 1].split(' ')[1];
    return allMonths.filter(m => m.split(' ')[1] === latestYear);
  }
  if (value === 'preset:qoq') {
    const latest = allMonths[allMonths.length - 1];
    const [latestMon, latestYear] = latest.split(' ');
    const qStart = Math.floor(MONTH_IDX[latestMon] / 3) * 3;
    const qNames = Object.keys(MONTH_IDX).slice(qStart, qStart + 3);
    return allMonths.filter(m => {
      const [mon, yr] = m.split(' ');
      return yr === latestYear && qNames.includes(mon);
    });
  }
  if (value === 'custom') return state.customPrimaryMonths || [];
  if (value === 'custom:compare') return state.customCompareMonths || [];
  if (value === 'auto') return getAutoPriorMonths(expandPeriod(state.primaryPeriod, allMonths), allMonths);
  // Single month label
  return allMonths.includes(value) ? [value] : [];
}

function getAutoPriorMonths(primaryMonths, allMonths) {
  if (!primaryMonths.length) return [];
  const n = primaryMonths.length;
  const firstIdx = allMonths.indexOf(primaryMonths[0]);
  if (firstIdx < n) return [];
  return allMonths.slice(firstIdx - n, firstIdx);
}

function getRangeMonths(from, to, allMonths) {
  const a = allMonths.indexOf(from), b = allMonths.indexOf(to);
  if (a < 0 || b < 0) return [];
  return allMonths.slice(Math.min(a, b), Math.max(a, b) + 1);
}

function getRowsForMonths(months) {
  if (!months.length) return [];
  const set = new Set(months);
  return rowsInScope(state.missedOpp.filter(r => set.has(r.monthLabel)), state.filters);
}

function getPeriodLabel(months) {
  if (!months || !months.length) return '—';
  if (months.length === 1) return months[0];
  return `${months[0]} – ${months[months.length - 1]} (${months.length}mo)`;
}

function computeCurrentPeriods() {
  state.primaryMonths = expandPeriod(state.primaryPeriod, state.allMonths);
  if (state.comparisonPeriod === 'auto') {
    state.comparisonMonths = getAutoPriorMonths(state.primaryMonths, state.allMonths);
  } else if (state.comparisonPeriod === 'custom:compare') {
    state.comparisonMonths = state.customCompareMonths || [];
  } else {
    state.comparisonMonths = expandPeriod(state.comparisonPeriod, state.allMonths);
  }
}

function updatePeriodInfoChip() {
  const el = document.getElementById('periodInfoChip');
  if (!el) return;
  const p = getPeriodLabel(state.primaryMonths);
  const c = state.comparisonMonths.length ? ` vs ${getPeriodLabel(state.comparisonMonths)}` : '';
  el.textContent = p + c;
  el.style.display = state.primaryMonths.length ? 'inline-flex' : 'none';
}

/* ---------- 8. KPI calculations (Playbook-aligned) ---------- */
function computeKPIs(moRows, apRows) {
  const totalMissedRev = moRows.reduce((s,r)=>s+r.missedRev, 0);
  const totalActualRev = moRows.reduce((s,r)=>s+r.actualRev, 0);
  const totalPotentialRev = moRows.reduce((s,r)=>s+r.potentialRev, 0);
  const totalMissedUnits = moRows.reduce((s,r)=>s+r.missedUnits, 0);
  const uniqueProducts = new Set(moRows.map(r=>r.product)).size;
  const avgOOSDaysPerProduct = moRows.length ? moRows.reduce((s,r)=>s+r.oosDays,0)/moRows.length : 0;
  const uniqueAssets = new Set(moRows.map(r=>r.asset)).size;
  const uniqueOutlets = new Set(moRows.map(r=>r.outletCode||r.outlet)).size;

  let emptySoS=null, foreignSoS=null, ownSoS=null;
  if (apRows.length) {
    emptySoS = apRows.reduce((s,r)=>s+r.emptySoS,0)/apRows.length;
    foreignSoS = apRows.reduce((s,r)=>s+r.foreignSoS,0)/apRows.length;
    ownSoS = Math.max(0, 100 - emptySoS - foreignSoS);
  }

  const totalAssets = apRows.length;
  const noCommAssets = apRows.filter(r=>r.noCommunication).length;
  const noImageAssets = apRows.filter(r=>r.noImage).length;
  const noDoorAssets = apRows.filter(r=>r.noDoor).length;
  const assetsWithImg7 = apRows.filter(r=>r.imagesPct7 > 0).length;

  return {
    totalMissedRev, totalActualRev, totalPotentialRev, totalMissedUnits,
    uniqueProducts, avgOOSDaysPerProduct, uniqueAssets, uniqueOutlets,
    emptySoS, foreignSoS, ownSoS,
    totalAssets, noCommAssets, noImageAssets, noDoorAssets, assetsWithImg7,
    moRowCount: moRows.length, apRowCount: apRows.length,
  };
}

function pctDelta(current, prior) {
  if (prior == null || current == null) return null;
  if (prior === 0) return current === 0 ? 0 : null;
  return ((current - prior) / Math.abs(prior)) * 100;
}

/* ---------- 9. Sample data generator (Playbook-shaped) ---------- */
function buildSampleData() {
  const distributors = ['Dixie Belle','Glacier Point','Mountain Peak','Coastal Foods','Heartland Distribution'];
  const markets = ['Southeast','Northeast','Midwest','West','South Central'];
  const channels = ['C-Store','Grocery','Mass','Drug'];
  const classifications = ['Platinum','Gold','Silver','Bronze'];
  const products = [
    {p:'GH Giant Vanilla Sandwich 1-24BK',  brand:'Magnum',  cs:24, cp:30.96, up:1.29},
    {p:'GH Magnum Almond 3.4oz 1-18BK',     brand:'Magnum',  cs:18, cp:35.10, up:1.95},
    {p:'GH Magnum Double Caramel 1-18BK',   brand:'Magnum',  cs:18, cp:36.50, up:2.03},
    {p:'GH Mini Bar Variety 1-30BK',        brand:'Magnum',  cs:30, cp:42.00, up:1.40},
    {p:'GH Klondike Bar 1-24BK',            brand:'Klondike',cs:24, cp:28.50, up:1.19},
    {p:'GH Klondike Cookie Sandwich 1-24BK',brand:'Klondike',cs:24, cp:29.99, up:1.25},
    {p:'GH Magnum Pint Vanilla 1-8BK',      brand:'Magnum',  cs:8,  cp:36.00, up:4.50},
    {p:'GH Magnum Pint Choc 1-8BK',         brand:'Magnum',  cs:8,  cp:36.00, up:4.50},
    {p:'GH Snicker Bar 1-24BK',             brand:'Snickers',cs:24, cp:30.00, up:1.25},
    {p:'GH M&M Cookie Sandwich 1-24BK',     brand:'M&M',     cs:24, cp:30.00, up:1.25},
  ];
  const cabinetTypes = ['Retrofit','OEM','None'];
  const routes = ['101','102','103','110','115','120','125'];
  const statuses = ['Fine - Images in last 3 days','No Image','No Communication','No Door'];

  const mo = [];
  const ap = [];
  const months = ['Jan 2026','Feb 2026','Mar 2026','Apr 2026'];

  let assetCounter = 600000000;
  let outletCounter = 1000;

  distributors.forEach((dist, di) => {
    const numOutlets = 8 + Math.floor(Math.random()*6);
    for (let o=0; o<numOutlets; o++) {
      const outletCode = `D${di+1}-${outletCounter++}`;
      const outletName = `${dist} Outlet ${o+1}`;
      const market = markets[di % markets.length];
      const channel = channels[Math.floor(Math.random()*channels.length)];
      const classification = classifications[Math.floor(Math.random()*classifications.length)];
      const route = routes[Math.floor(Math.random()*routes.length)];
      const cabinetType = cabinetTypes[Math.floor(Math.random()*cabinetTypes.length)];
      const assetSerial = String(assetCounter++);

      // Asset perf row (one per outlet)
      const statusRoll = Math.random();
      let status = statuses[0], diagnosis = 'Fine - Images in last 3 days';
      let imagesPct = 80 + Math.random()*20;
      if (statusRoll < 0.08) { status='No Communication'; diagnosis='No Communication'; imagesPct = 0; }
      else if (statusRoll < 0.22) { status='No Image'; diagnosis='Other'; imagesPct = Math.random()*30; }
      else if (statusRoll < 0.28) { status='No Door'; diagnosis='No Door'; }

      const emptySoS = Math.max(0, Math.random()*18 + (statusRoll < 0.3 ? 8 : 0));
      const foreignSoS = Math.random()*12;
      const facings = 18 + Math.floor(Math.random()*12);
      const today = new Date();
      const lastImgDays = status === 'No Communication' ? 4+Math.floor(Math.random()*30) : Math.floor(Math.random()*3);
      const lastImg = new Date(today.getTime() - lastImgDays*86400000);

      const dailyImg = [];
      for (let d=6; d>=0; d--) {
        const dt = new Date(today.getTime() - d*86400000);
        const day = `${dt.getDate()}-${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][dt.getMonth()]}`;
        const hasImages = (status === 'No Communication') ? 0 : (Math.random() < 0.85 ? Math.floor(Math.random()*8)+1 : 0);
        const doors = (status === 'No Door') ? 0 : Math.floor(Math.random()*15)+1;
        dailyImg.push({ date: day, doors, images: hasImages });
      }
      ap.push({
        distributor: dist, distributorRaw: `MICC USA ${dist}`,
        asset: assetSerial, cabinetType, outlet: outletName, outletCode,
        city:'', street:'', classification, deviceNumber:`DEV${assetSerial.slice(-5)}`,
        oosPct: emptySoS * (0.6 + Math.random()*0.4),
        facings, installedOn: new Date(today.getTime() - (180+Math.floor(Math.random()*500))*86400000),
        devicePing: status==='No Communication'? new Date(today.getTime() - lastImgDays*86400000) : new Date(today.getTime() - Math.random()*3600000),
        doorClose: lastImg,
        imagesPct7: imagesPct, missingImg7: Math.floor((100-imagesPct)/14),
        diagnosis, status, lastImg, gps:0, temperature: -2 + Math.random()*5,
        emptySoS, foreignSoS, hardwareStatus: status==='No Communication'? 'No Image':'OK',
        dailyImg,
        noCommunication: status==='No Communication',
        noImage: status==='No Image',
        noDoor: status==='No Door',
      });

      // MO rows per outlet x product x month
      months.forEach(monthLabel => {
        const productSubset = products.filter(()=> Math.random() < 0.75);
        productSubset.forEach(pd => {
          const oosBase = Math.floor(Math.random() * (statusRoll < 0.3 ? 18 : 8));
          const inStockDays = 30 - oosBase;
          const avgDailyUnits = 0.3 + Math.random()*3;
          const unitsSold = inStockDays * avgDailyUnits;
          const missedUnits = oosBase * avgDailyUnits;
          const actualRev = unitsSold * pd.up;
          const missedRev = missedUnits * pd.up;
          mo.push({
            monthLabel,
            distributor: dist, distributorRaw: `MICC USA ${dist}`,
            outlet: outletName, outletCode,
            product: pd.p, brand: pd.brand, asset: assetSerial,
            casesPerFacing: 1, casePrice: pd.cp, unitPrice: pd.up, caseSize: pd.cs,
            inStockDays, oosDays: oosBase,
            casesSold: unitsSold/pd.cs, unitsSold,
            timesRefilled: 1 + Math.floor(Math.random()*3),
            avgDailyCases: avgDailyUnits/pd.cs, avgDailyUnits,
            missedUnits, missedRev, actualRev, potentialRev: actualRev + missedRev,
            planogram: '5 Basket With Pint Ind Accounts_10 Apr 26',
            salesRep:'', market, channel, classification, route,
          });
        });
      });
    }
  });

  return { mo, ap, isSample: true };
}

/* ---------- 10. File handlers ---------- */
function setupDropzones() {
  ['moDrop','apDrop'].forEach(id => {
    const dz = document.getElementById(id);
    const fileInput = document.getElementById(id+'-input');
    dz.addEventListener('click', () => fileInput.click());
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
    dz.addEventListener('drop', e => {
      e.preventDefault(); dz.classList.remove('drag');
      handleFiles(id, e.dataTransfer.files);
    });
    fileInput.addEventListener('change', e => handleFiles(id, e.target.files));
  });
  document.getElementById('useSampleBtn').addEventListener('click', () => {
    const sd = buildSampleData();
    state.missedOpp = sd.mo; state.assetPerf = sd.ap; state.isSample = true;
    boot();
  });
  document.getElementById('generateBtn').addEventListener('click', () => {
    if (!state.missedOpp.length && !state.assetPerf.length) {
      alert('Please upload at least one Playbook export, or use sample data.');
      return;
    }
    state.isSample = false;
    boot();
  });
}

const uploadedFiles = { mo: [], ap: [] };

// Status feedback in console & UI
function logUpload(msg, level = 'info') {
  console.log(`[Upload] ${msg}`);
  // also push to a small status banner if it exists
  const banner = document.getElementById('uploadStatusBanner');
  if (banner) {
    const div = document.createElement('div');
    div.className = `upload-status upload-status-${level}`;
    div.textContent = msg;
    banner.appendChild(div);
    // auto-fade after 8s
    setTimeout(() => { div.style.opacity = '0'; setTimeout(() => div.remove(), 1000); }, 8000);
  }
}

function handleFiles(slot, files) {
  const slotKey = slot === 'moDrop' ? 'mo' : 'ap';
  for (const f of files) {
    addFilePill(slotKey, f.name); // initial placement uses dropped slot; updateFilePill moves it after detection
    if (f.name.toLowerCase().endsWith('.zip')) {
      parseZip(f);
    } else {
      parseFile(f);
    }
  }
}

async function parseZip(file) {
  // Use SheetJS's CFB.read for zip? No — we need JSZip. Fall back to dropping a clear error if JSZip not loaded.
  if (typeof JSZip === 'undefined') {
    logUpload(`Cannot open ${file.name}: ZIP support not loaded. Please extract the zip and upload individual files.`, 'error');
    alert(`This file is a ZIP archive. Please extract it first and drop the individual .xlsx files, or ask DJ to add JSZip support to the dashboard.`);
    return;
  }
  try {
    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    const entries = Object.values(zip.files).filter(e => !e.dir && /\.xlsx?$/i.test(e.name));
    logUpload(`Found ${entries.length} spreadsheet(s) in ${file.name}`);
    for (const entry of entries) {
      const data = await entry.async('uint8array');
      // wrap as a pseudo-file for the existing parser
      const pseudo = { name: entry.name, _arrayBuffer: data };
      parseFile(pseudo);
    }
  } catch (err) {
    logUpload(`Failed to read ZIP ${file.name}: ${err.message}`, 'error');
  }
}

function parseFile(file) {
  const onData = (data) => {
    let wb;
    try {
      wb = XLSX.read(data, { type: 'array', cellDates: true });
    } catch (err) {
      logUpload(`Could not parse ${file.name}: ${err.message}`, 'error');
      alert(`Could not parse ${file.name}: ${err.message}`);
      return;
    }
    const monthLabel = inferMonthLabel(file.name);
    let totalParsed = 0;
    let detectedKind = null;
    const sheetResults = [];

    for (const sheetName of wb.SheetNames) {
      const sheet = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!json.length) continue;
      const headers = json[0].map(h => String(h || '').trim());
      if (!headers.some(h => h.length > 0)) continue; // empty sheet
      const idx = buildIndex(headers);
      const kind = detectKind(idx);

      if (kind === FIELD_KIND.UNKNOWN) {
        sheetResults.push({ sheetName, status: 'skipped', reason: 'no recognized columns' });
        continue;
      }

      let rowsAdded = 0;
      if (kind === FIELD_KIND.MISSED_OPP) {
        for (let r = 1; r < json.length; r++) {
          if (!json[r] || !json[r][idx.client]) continue;
          state.missedOpp.push(normMORow(json[r], idx, monthLabel));
          rowsAdded++;
        }
      } else if (kind === FIELD_KIND.ASSET_PERF) {
        const dailyImgCols = findDailyImgCols(headers);
        for (let r = 1; r < json.length; r++) {
          if (!json[r] || !json[r][idx.client]) continue;
          state.assetPerf.push(normAPRow(json[r], idx, dailyImgCols));
          rowsAdded++;
        }
      }
      sheetResults.push({ sheetName, status: 'ok', kind, rowsAdded });
      totalParsed += rowsAdded;
      detectedKind = kind;
    }

    // Update file pill with detected kind (and move it to the correct slot if needed)
    const kindLabel = detectedKind === FIELD_KIND.MISSED_OPP ? 'Missed Opp'
                    : detectedKind === FIELD_KIND.ASSET_PERF ? 'Asset Performance'
                    : 'Unrecognized';
    updateFilePill(file.name, kindLabel, totalParsed, detectedKind);

    if (totalParsed === 0) {
      logUpload(`⚠️ ${file.name}: no rows parsed. Sheets checked: ${sheetResults.map(s => `${s.sheetName} (${s.status})`).join(', ')}`, 'warn');
      alert(`${file.name}: I couldn't find Missed Opportunity or Asset Performance columns in any sheet. The file may be in an unexpected format — check the column headers match Playbook's export.`);
    } else {
      logUpload(`✓ ${file.name}: parsed ${totalParsed.toLocaleString()} ${kindLabel} rows`);
    }
    updateUploadStatus();
  };

  if (file._arrayBuffer) {
    // already-extracted ZIP entry
    onData(file._arrayBuffer);
    return;
  }
  const reader = new FileReader();
  reader.onload = e => onData(new Uint8Array(e.target.result));
  reader.readAsArrayBuffer(file);
}

function addFilePill(slot, name) {
  // Place pill in the slot it was dropped on; updateFilePill will move it after parsing if it was the wrong slot.
  const targetId = slot === 'mo' ? 'moFiles' : 'apFiles';
  const target = document.getElementById(targetId);
  if (!target) return;
  const pill = document.createElement('div');
  pill.className = 'file-pill';
  pill.dataset.filename = name;
  pill.dataset.droppedSlot = slot;
  pill.innerHTML = `
    <span class="pill-name" style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;margin-right:6px" title="${name}">${name}</span>
    <span class="pill-kind muted" style="font-size:9.5px;margin-right:6px">parsing…</span>
    <span class="remove" title="Remove" style="cursor:pointer">✕</span>`;
  pill.querySelector('.remove').addEventListener('click', () => pill.remove());
  target.appendChild(pill);
}

function updateFilePill(name, kindLabel, rowCount, detectedKind) {
  const pills = document.querySelectorAll('.file-pill');
  pills.forEach(p => {
    if (p.dataset.filename !== name) return;
    const kindEl = p.querySelector('.pill-kind');
    const droppedSlot = p.dataset.droppedSlot;
    let correctSlotId = null;
    if (detectedKind === FIELD_KIND.MISSED_OPP) correctSlotId = 'moFiles';
    else if (detectedKind === FIELD_KIND.ASSET_PERF) correctSlotId = 'apFiles';

    if (kindEl) {
      if (rowCount > 0 && correctSlotId) {
        kindEl.innerHTML = `<span class="pill ${correctSlotId === 'moFiles' ? 'blue' : 'green'}" style="font-size:9.5px">${kindLabel} • ${rowCount.toLocaleString()} rows</span>`;
      } else {
        kindEl.innerHTML = `<span class="pill red" style="font-size:9.5px">unrecognized</span>`;
      }
    }

    // Move pill to the correct slot if it was dropped on the wrong one.
    if (correctSlotId && p.parentElement && p.parentElement.id !== correctSlotId) {
      const oldSlot = p.parentElement;
      const correctSlot = document.getElementById(correctSlotId);
      if (correctSlot) {
        correctSlot.appendChild(p);
        // Flash a brief note in the slot where it was dropped, explaining the routing.
        const expectedHere = oldSlot.id === 'moFiles' ? 'Missed Opportunity' : 'Asset Performance';
        const movedTo     = correctSlotId === 'moFiles' ? 'Missed Opportunity' : 'Asset Performance';
        const note = document.createElement('div');
        note.className = 'muted small';
        note.style.cssText = 'font-style:italic;font-size:10.5px;color:#FF8500;padding:4px 0;line-height:1.3';
        note.textContent = `"${name}" was dropped on ${expectedHere} but detected as ${movedTo} — moved.`;
        oldSlot.appendChild(note);
        setTimeout(() => note.remove(), 7000);
      }
    }
  });
}
function inferMonthLabel(name) {
  // Handles patterns like:
  //   MICC_MO_ALL_April_2026_PB_Export.xlsx       -> "April 2026"
  //   MICC USA Dixie Belle_April_2025.xlsx        -> "April 2025"
  //   Asset_Performance-2026-05-20T13-00-12.xlsx  -> "May 2026" (from ISO date)
  //   anything_Jan2026_something                  -> "Jan 2026"
  // 1. Full or abbreviated month name + year (with optional separator)
  const m1 = name.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s_\-]*(\d{4})/i);
  if (m1) return `${capitalize(m1[1])} ${m1[2]}`;
  // 2. ISO date like 2026-05-20 -> "May 2026"
  const m2 = name.match(/(\d{4})-(\d{2})-\d{2}/);
  if (m2) {
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const monthNum = parseInt(m2[2], 10);
    if (monthNum >= 1 && monthNum <= 12) return `${months[monthNum-1]} ${m2[1]}`;
  }
  // 3. YYYYMM or YYYY_MM
  const m3 = name.match(/(\d{4})[_\-]?(\d{2})(?!\d)/);
  if (m3) {
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const monthNum = parseInt(m3[2], 10);
    if (monthNum >= 1 && monthNum <= 12) return `${months[monthNum-1]} ${m3[1]}`;
  }
  return 'Uploaded';
}
function capitalize(s) {
  // Normalize to 3-letter month abbreviation to match monthSort() expectations
  return s.charAt(0).toUpperCase() + s.slice(1, 3).toLowerCase();
}
function updateUploadStatus() {
  const moEl = document.getElementById('moCount');
  const apEl = document.getElementById('apCount');
  if (moEl) moEl.textContent = `${state.missedOpp.length.toLocaleString()} rows parsed`;
  if (apEl) apEl.textContent = `${state.assetPerf.length.toLocaleString()} rows parsed`;
}

/* ---------- 11. Boot ---------- */
function boot() {
  document.getElementById('landing').style.display = 'none';
  document.getElementById('mainApp').style.display = 'block';
  document.getElementById('sampleBadge').style.display = state.isSample ? 'inline-block' : 'none';
  document.getElementById('liveBadge').style.display = state.isSample ? 'none' : 'inline-block';
  populateFilters();
  setDefaultPeriod();
  renderAll();
}

/* ---------- 12. Populate filter dropdowns ---------- */
function unique(rows, key) {
  return Array.from(new Set(rows.map(r=>r[key]).filter(v=>v!==undefined && v!==null && v!==''))).sort();
}
function populateFilters() {
  const filtSpecs = [
    {id:'fDistributor', key:'distributor', src:state.missedOpp.concat(state.assetPerf)},
    {id:'fMarket', key:'market', src:state.missedOpp},
    {id:'fChannel', key:'channel', src:state.missedOpp},
    {id:'fClassification', key:'classification', src:state.missedOpp.concat(state.assetPerf)},
    {id:'fRoute', key:'route', src:state.missedOpp},
    {id:'fOutlet', key:'outlet', src:state.missedOpp.concat(state.assetPerf)},
    {id:'fProduct', key:'product', src:state.missedOpp},
    {id:'fCameraType', key:'cabinetType', src:state.assetPerf},
    {id:'fAssetStatus', key:null, src:null, opts:['all','OK','No Communication','No Image','No Door']},
    {id:'fPeriod', key:null, src:null, opts:[]},
    {id:'fCompare', key:null, src:null, opts:[]},
  ];
  filtSpecs.forEach(spec => {
    const el = document.getElementById(spec.id);
    if (!el) return;
    el.innerHTML = '';
    if (spec.opts) {
      spec.opts.forEach(v => {
        const o = document.createElement('option');
        o.value = v; o.textContent = (v === 'all' ? 'All' : v);
        el.appendChild(o);
      });
    } else {
      const all = document.createElement('option');
      all.value = 'all'; all.textContent = 'All';
      el.appendChild(all);
      unique(spec.src, spec.key).forEach(v => {
        const o = document.createElement('option'); o.value=v; o.textContent=v;
        el.appendChild(o);
      });
    }
    el.addEventListener('change', onFilterChange);
  });

  // ── Period selects (smart presets) ───────────────────────────────────────
  const months = Array.from(new Set(state.missedOpp.map(r=>r.monthLabel))).sort(monthSort);
  state.allMonths = months;

  const pEl = document.getElementById('fPeriod');
  const cEl = document.getElementById('fCompare');
  if (!pEl || !cEl) return;

  // Build primary period dropdown
  pEl.innerHTML = '';
  const pPresets = [
    {value:'preset:mom', label:'Month over Month'},
    {value:'preset:qoq', label:'Quarter over Quarter'},
    {value:'preset:ytd', label:'Year to Date'},
    {value:'custom',     label:'Custom range…'},
  ];
  addOptGroup(pEl, '— Presets —', pPresets);
  addOptGroup(pEl, '— Single month —', months.slice().reverse().map(m=>({value:m,label:m})));
  pEl.value = 'preset:mom';
  pEl.addEventListener('change', onFilterChange);

  // Build compare-to dropdown
  cEl.innerHTML = '';
  const cPresets = [
    {value:'auto',           label:'Prior period (auto)'},
    {value:'none',           label:'— none —'},
    {value:'custom:compare', label:'Custom range…'},
  ];
  addOptGroup(cEl, '— Presets —', cPresets);
  addOptGroup(cEl, '— Single month —', months.slice().reverse().map(m=>({value:m,label:m})));
  cEl.value = 'auto';
  cEl.addEventListener('change', onFilterChange);

  // Populate custom range month selectors
  ['fPeriodFrom','fPeriodTo','fCompareFrom','fCompareTo'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = '';
    months.slice().reverse().forEach(m => {
      const o = document.createElement('option'); o.value=m; o.textContent=m; el.appendChild(o);
    });
    el.addEventListener('change', onFilterChange);
  });
  // Sensible custom range defaults
  const last = months[months.length - 1] || '';
  const prev = months.length > 1 ? months[months.length - 2] : last;
  const pFrom = document.getElementById('fPeriodFrom');
  const pTo   = document.getElementById('fPeriodTo');
  const cFrom = document.getElementById('fCompareFrom');
  const cTo   = document.getElementById('fCompareTo');
  if (pFrom) pFrom.value = last;
  if (pTo)   pTo.value   = last;
  if (cFrom) cFrom.value = prev;
  if (cTo)   cTo.value   = prev;
}

function addOptGroup(selectEl, label, items) {
  if (!items.length) return;
  const grp = document.createElement('optgroup'); grp.label = label;
  items.forEach(({value, label}) => {
    const o = document.createElement('option'); o.value = value; o.textContent = label;
    grp.appendChild(o);
  });
  selectEl.appendChild(grp);
}
function monthSort(a,b) {
  const M={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const pa=a.split(' '), pb=b.split(' ');
  return (parseInt(pa[1])-parseInt(pb[1]))*12 + (M[pa[0]]-M[pb[0]]);
}
function setDefaultPeriod() {
  state.primaryPeriod = 'preset:mom';
  state.comparisonPeriod = 'auto';
}
function onFilterChange() {
  state.filters.distributor = document.getElementById('fDistributor').value;
  state.filters.market = document.getElementById('fMarket').value;
  state.filters.channel = document.getElementById('fChannel').value;
  state.filters.classification = document.getElementById('fClassification').value;
  state.filters.route = document.getElementById('fRoute').value;
  state.filters.outlet = document.getElementById('fOutlet').value;
  state.filters.product = document.getElementById('fProduct').value;
  state.filters.cameraType = document.getElementById('fCameraType').value;
  state.filters.assetStatus = document.getElementById('fAssetStatus').value;
  state.primaryPeriod = document.getElementById('fPeriod').value;
  state.comparisonPeriod = document.getElementById('fCompare').value;

  // Show/hide custom range pickers
  const showPrimCustom = state.primaryPeriod === 'custom';
  const showCompCustom = state.comparisonPeriod === 'custom:compare';
  const pcr = document.getElementById('periodCustomRange');
  const ccr = document.getElementById('compareCustomRange');
  if (pcr) pcr.style.display = showPrimCustom ? 'flex' : 'none';
  if (ccr) ccr.style.display = showCompCustom ? 'flex' : 'none';

  // Build custom month arrays
  if (showPrimCustom) {
    const from = document.getElementById('fPeriodFrom').value;
    const to   = document.getElementById('fPeriodTo').value;
    state.customPrimaryMonths = getRangeMonths(from, to, state.allMonths);
  }
  if (showCompCustom) {
    const from = document.getElementById('fCompareFrom').value;
    const to   = document.getElementById('fCompareTo').value;
    state.customCompareMonths = getRangeMonths(from, to, state.allMonths);
  }

  renderAll();
}
function clearFilters() {
  ['fDistributor','fMarket','fChannel','fClassification','fRoute','fOutlet','fProduct','fCameraType','fAssetStatus'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = 'all';
  });
  const pEl = document.getElementById('fPeriod');
  const cEl = document.getElementById('fCompare');
  if (pEl) pEl.value = 'preset:mom';
  if (cEl) cEl.value = 'auto';
  const pcr = document.getElementById('periodCustomRange');
  const ccr = document.getElementById('compareCustomRange');
  if (pcr) pcr.style.display = 'none';
  if (ccr) ccr.style.display = 'none';
  onFilterChange();
}

/* ---------- 13. Render ---------- */
function getPeriodRows(period) {
  // Legacy shim — used by drill-down callbacks that capture state.primaryPeriod
  if (!period || period === 'none') return [];
  if (period === 'custom') return getRowsForMonths(state.customPrimaryMonths || []);
  return getRowsForMonths(expandPeriod(period, state.allMonths));
}
function getAssetRows() {
  return assetsInScope(state.assetPerf, state.filters);
}

function renderAll() {
  computeCurrentPeriods();
  updatePeriodInfoChip();
  if (state.view === 'executive') renderExecutive();
  else renderDistributor();
}

function viewToggle(target) {
  state.view = target;
  document.querySelectorAll('.view-toggle button').forEach(b=>{
    b.classList.toggle('active', b.dataset.view === target);
  });
  document.getElementById('execView').style.display = target==='executive' ? 'block':'none';
  document.getElementById('distView').style.display = target==='distributor' ? 'block':'none';
  renderAll();
}

/* ---------- 14. EXECUTIVE VIEW ---------- */
function renderExecutive() {
  const primary = getRowsForMonths(state.primaryMonths);
  const comparison = getRowsForMonths(state.comparisonMonths);
  const assets = getAssetRows();
  const k = computeKPIs(primary, assets);
  const kp = comparison.length ? computeKPIs(comparison, assets) : null;

  renderExecKPIs(k, kp);
  renderTopDistributorsByMissedRev(primary);
  renderEmptyShelfChart(primary, comparison, assets);
  renderFSOSChart(assets, primary);
  renderAssetHealthPanel(assets);
  renderMissedRevDrill(primary);
  renderTPC(primary);
  renderRootCausePanel(primary, assets);
  renderBeforeAfter(k, kp);
}

function renderExecKPIs(k, kp) {
  const grid = document.getElementById('execKPIs');
  // desired='up'  → increase is good (green ▲, red ▼)
  // desired='down' → decrease is good (green ▼, red ▲)
  const delta = (cur, prior, fmt, desired='up') => {
    if (prior == null || cur == null) return '';
    const d = pctDelta(cur, prior);
    if (d == null) return '';
    const direction = d > 0.5 ? 'up' : d < -0.5 ? 'down' : 'flat';
    const isGood = direction === 'flat' || direction === desired;
    const colorClass = direction === 'flat' ? 'flat' : isGood ? 'up' : 'down';
    const arrow = d > 0.5 ? '▲' : d < -0.5 ? '▼' : '◆';
    const diff = cur - prior;
    return `<div class="kpi-delta ${colorClass}">${arrow} ${Math.abs(d).toFixed(1)}% (${fmt(diff)})</div>`;
  };
  const deltaNum = (cur, prior, desired='up') => {
    if (prior == null) return '';
    const d = pctDelta(cur, prior);
    if (d == null) return '';
    const direction = d > 0.5 ? 'up' : d < -0.5 ? 'down' : 'flat';
    const isGood = direction === 'flat' || direction === desired;
    const colorClass = direction === 'flat' ? 'flat' : isGood ? 'up' : 'down';
    const arrow = d > 0.5 ? '▲' : d < -0.5 ? '▼' : '◆';
    return `<div class="kpi-delta ${colorClass}">${arrow} ${Math.abs(d).toFixed(1)}%</div>`;
  };

  grid.innerHTML = '';
  const _mo = () => getRowsForMonths(state.primaryMonths);
  const _ap = () => getAssetRows();
  grid.appendChild(kpiTile('Actual Revenue', fmtMoney(k.totalActualRev), kp ? delta(k.totalActualRev, kp.totalActualRev, fmtMoney, 'up') : '',
    `${getPeriodLabel(state.primaryMonths)} · ${k.uniqueProducts} SKUs across ${k.uniqueOutlets} outlets`,
    'green', () => openMetricDrill('actualRev', _mo(), _ap())));
  grid.appendChild(kpiTile('Missed Revenue', fmtMoney(k.totalMissedRev), kp ? delta(k.totalMissedRev, kp.totalMissedRev, fmtMoney, 'down') : '',
    `${fmtNum(k.totalMissedUnits)} missed units · avg ${fmtMoney(k.totalMissedRev / Math.max(1,k.uniqueAssets))} per asset`,
    'red', () => openMetricDrill('missedRev', _mo(), _ap())));
  grid.appendChild(kpiTile('Empty Share of Shelf',
    k.emptySoS != null ? fmtPct(k.emptySoS) : '—', '',
    k.emptySoS != null ? 'Empty Facings / Total Facings × 100 (Playbook)' : 'No Asset Performance data uploaded',
    'orange', () => openMetricDrill('emptySoS', _mo(), _ap())));
  grid.appendChild(kpiTile('Foreign Share of Shelf',
    k.foreignSoS != null ? fmtPct(k.foreignSoS) : '—', '',
    k.foreignSoS != null ? 'Foreign Facings / Total Facings × 100' : 'No Asset Performance data uploaded',
    'purple', () => openMetricDrill('foreignSoS', _mo(), _ap())));
  grid.appendChild(kpiTile('Avg Days SKU OOS / Product',
    fmtNum(k.avgOOSDaysPerProduct, 1) + ' d',
    kp ? deltaNum(k.avgOOSDaysPerProduct, kp.avgOOSDaysPerProduct, 'down') : '',
    `Average over ${k.moRowCount.toLocaleString()} SKU-asset rows · ${getPeriodLabel(state.primaryMonths)}`,
    'pink', () => openMetricDrill('oosDays', _mo(), _ap())));
  const activeAssets = k.totalAssets - k.noCommAssets - k.noImageAssets - k.noDoorAssets;
  const activePct = k.totalAssets > 0 ? (activeAssets / k.totalAssets) * 100 : null;
  const assetColorClass = activePct == null ? '' : activePct >= 90 ? 'green' : activePct >= 75 ? 'orange' : 'red';
  grid.appendChild(kpiTile('Asset Health',
    activePct != null ? fmtPct(activePct) : '—',
    kp ? (() => {
      const priorActive = kp.totalAssets - kp.noCommAssets - kp.noImageAssets - kp.noDoorAssets;
      const priorPct = kp.totalAssets > 0 ? (priorActive / kp.totalAssets) * 100 : null;
      return deltaNum(activePct, priorPct, 'up');
    })() : '',
    `${activeAssets.toLocaleString()} of ${k.totalAssets.toLocaleString()} active · ${k.noCommAssets} No Comm · ${k.noImageAssets} No Image · ${k.noDoorAssets} No Door`,
    assetColorClass, () => openMetricDrill('assetHealth', _mo(), _ap())));
}

function kpiTile(label, value, deltaHTML, sub, colorClass, onClick) {
  const div = document.createElement('div');
  div.className = 'kpi-tile' + (colorClass ? ' '+colorClass : '');
  div.innerHTML = `
    <div class="drill">drill ▸</div>
    <div class="kpi-label">${label}</div>
    <div class="kpi-value">${value}</div>
    ${deltaHTML || ''}
    <div class="kpi-sub">${sub}</div>
  `;
  if (onClick) div.addEventListener('click', onClick);
  return div;
}

/* ---------- 15. Charts ---------- */
function destroyChart(id) {
  if (state.charts[id]) { state.charts[id].destroy(); delete state.charts[id]; }
}
function renderTopDistributorsByMissedRev(rows) {
  const byDist = {};
  rows.forEach(r => {
    if (!byDist[r.distributor]) byDist[r.distributor] = { missed: 0, actual: 0 };
    byDist[r.distributor].missed += r.missedRev;
    byDist[r.distributor].actual += r.actualRev;
  });
  const sorted = Object.entries(byDist)
    .sort((a, b) => (b[1].missed + b[1].actual) - (a[1].missed + a[1].actual))
    .slice(0, 10);
  destroyChart('chTopDist');
  const ctx = document.getElementById('chTopDist').getContext('2d');
  state.charts.chTopDist = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sorted.map(s => s[0]),
      datasets: [
        {
          label: 'Actual Revenue',
          data: sorted.map(s => s[1].actual),
          backgroundColor: '#5AC581',
          stack: 'rev',
          borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 3, bottomRight: 3 },
          borderSkipped: false,
        },
        {
          label: 'Missed Revenue',
          data: sorted.map(s => s[1].missed),
          backgroundColor: '#E93D31',
          stack: 'rev',
          borderRadius: { topLeft: 3, topRight: 3, bottomLeft: 0, bottomRight: 0 },
          borderSkipped: false,
        },
      ],
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { font: { size: 11 }, boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            label: c => `${c.dataset.label}: ${fmtMoneyFull(c.parsed.x)}`,
            footer: items => `Potential total: ${fmtMoneyFull(items.reduce((s, i) => s + i.parsed.x, 0))}`,
          },
        },
      },
      scales: {
        x: { stacked: true, ticks: { callback: v => fmtMoney(v) }, grid: { color: '#EEE' } },
        y: { stacked: true, ticks: { font: { size: 11 } }, grid: { display: false } },
      },
    },
  });
}

function renderEmptyShelfChart(primary, comparison, assets) {
  const distMap = {};
  assets.forEach(a => {
    if (!distMap[a.distributor]) distMap[a.distributor] = { sum: 0, n: 0 };
    distMap[a.distributor].sum += a.emptySoS;
    distMap[a.distributor].n += 1;
  });
  const actualByDist = {};
  primary.forEach(r => { actualByDist[r.distributor] = (actualByDist[r.distributor] || 0) + r.actualRev; });

  const arr = Object.entries(distMap)
    .map(([k, v]) => ({ dist: k, sos: v.sum / v.n, actual: actualByDist[k] || 0 }))
    .sort((a, b) => b.sos - a.sos).slice(0, 12);

  destroyChart('chEmpty');
  const el = document.getElementById('chEmpty'); if (!el) return;
  if (!arr.length) { el.parentElement.innerHTML = '<div class="empty-state">No Asset Performance data — upload an Asset_Performance Playbook export.</div>'; return; }
  const ctx = el.getContext('2d');
  state.charts.chEmpty = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: arr.map(s => s.dist),
      datasets: [
        {
          label: 'Empty SoS %',
          data: arr.map(s => +s.sos.toFixed(2)),
          backgroundColor: '#FF8500',
          borderRadius: 3,
          yAxisID: 'y',
          order: 2,
        },
        {
          label: 'Actual Revenue',
          data: arr.map(s => s.actual),
          type: 'line',
          borderColor: '#1A29B6',
          backgroundColor: 'rgba(26,41,182,0.08)',
          pointBackgroundColor: '#1A29B6',
          pointRadius: 4,
          borderWidth: 2,
          fill: false,
          tension: 0.3,
          yAxisID: 'y2',
          order: 1,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { font: { size: 11 }, boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            label: c => c.dataset.label === 'Empty SoS %'
              ? `Empty SoS: ${fmtPct(c.parsed.y)}`
              : `Actual Rev: ${fmtMoneyFull(c.parsed.y)}`,
          },
        },
      },
      scales: {
        y: {
          type: 'linear', position: 'left',
          ticks: { callback: v => v + '%' }, grid: { color: '#EEE' },
          title: { display: true, text: 'Empty SoS %', color: '#FF8500', font: { size: 10, weight: '600' } },
        },
        y2: {
          type: 'linear', position: 'right',
          ticks: { callback: v => fmtMoney(v) }, grid: { display: false },
          title: { display: true, text: 'Actual Revenue', color: '#1A29B6', font: { size: 10, weight: '600' } },
        },
        x: { ticks: { font: { size: 10 }, maxRotation: 60, minRotation: 30 }, grid: { display: false } },
      },
    },
  });
}

function renderFSOSChart(assets, rows) {
  const el = document.getElementById('chFSOS'); if (!el) return;
  destroyChart('chFSOS');
  if (!assets.length) { el.parentElement.innerHTML = '<div class="empty-state">No Asset Performance data uploaded.</div>'; return; }
  const distMap = {};
  assets.forEach(a => {
    if (!distMap[a.distributor]) distMap[a.distributor] = { sum: 0, n: 0 };
    distMap[a.distributor].sum += a.foreignSoS;
    distMap[a.distributor].n += 1;
  });
  const actualByDist = {};
  (rows || []).forEach(r => { actualByDist[r.distributor] = (actualByDist[r.distributor] || 0) + r.actualRev; });

  const arr = Object.entries(distMap)
    .map(([k, v]) => ({ dist: k, sos: v.sum / v.n, actual: actualByDist[k] || 0 }))
    .sort((a, b) => b.sos - a.sos).slice(0, 12);

  const ctx = el.getContext('2d');
  state.charts.chFSOS = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: arr.map(s => s.dist),
      datasets: [
        {
          label: 'Foreign SoS %',
          data: arr.map(s => +s.sos.toFixed(2)),
          backgroundColor: '#9F36C4',
          borderRadius: 3,
          yAxisID: 'y',
          order: 2,
        },
        {
          label: 'Actual Revenue',
          data: arr.map(s => s.actual),
          type: 'line',
          borderColor: '#1A29B6',
          backgroundColor: 'rgba(26,41,182,0.08)',
          pointBackgroundColor: '#1A29B6',
          pointRadius: 4,
          borderWidth: 2,
          fill: false,
          tension: 0.3,
          yAxisID: 'y2',
          order: 1,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { font: { size: 11 }, boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            label: c => c.dataset.label === 'Foreign SoS %'
              ? `Foreign SoS: ${fmtPct(c.parsed.y)}`
              : `Actual Rev: ${fmtMoneyFull(c.parsed.y)}`,
          },
        },
      },
      scales: {
        y: {
          type: 'linear', position: 'left',
          ticks: { callback: v => v + '%' }, grid: { color: '#EEE' },
          title: { display: true, text: 'Foreign SoS %', color: '#9F36C4', font: { size: 10, weight: '600' } },
        },
        y2: {
          type: 'linear', position: 'right',
          ticks: { callback: v => fmtMoney(v) }, grid: { display: false },
          title: { display: true, text: 'Actual Revenue', color: '#1A29B6', font: { size: 10, weight: '600' } },
        },
        x: { ticks: { font: { size: 10 }, maxRotation: 60, minRotation: 30 }, grid: { display: false } },
      },
    },
  });
}

function renderAssetHealthPanel(assets) {
  state.currentAssets = assets;
  const el = document.getElementById('assetMatrix');
  if (!el) return;

  if (!assets.length) {
    el.innerHTML = '<div class="empty-state">Upload an Asset_Performance Playbook export to see asset health detail.</div>';
    return;
  }

  // ── Grid (dot matrix) ────────────────────────────────────────────────────
  el.innerHTML = '';
  const ordered = assets.slice().sort((a, b) => {
    const rank = r => r.noCommunication ? 0 : r.noImage ? 1 : r.noDoor ? 2 : 3;
    return rank(a) - rank(b);
  });
  ordered.slice(0, 600).forEach(a => {
    const d = document.createElement('div');
    d.className = 'a';
    if (a.noCommunication) d.classList.add('bad');
    else if (a.noImage) d.classList.add('warn');
    else if (a.noDoor) d.classList.add('gray');
    d.title = `${a.outlet} · ${a.asset} · ${a.status}`;
    d.addEventListener('click', () => openAssetDetail(a));
    el.appendChild(d);
  });
  if (assets.length > 600) {
    const more = document.createElement('div');
    more.className = 'muted small'; more.style.gridColumn = '1/-1'; more.style.marginTop = '6px';
    more.textContent = `Showing first 600 of ${assets.length.toLocaleString()} assets — apply filters to narrow.`;
    el.appendChild(more);
  }

  // ── Summary tile counts & click handlers ────────────────────────────────
  const tiles = {
    'apTile-all':      { filter: () => true,                                         label: 'All Installed Assets',      val: 'apTotalAssets', count: assets.length },
    'apTile-nocomm':   { filter: a => a.noCommunication,                             label: 'No Communication Assets',   val: 'apNoComm',      count: assets.filter(a=>a.noCommunication).length },
    'apTile-noimage':  { filter: a => a.noImage,                                     label: 'No Image Assets',           val: 'apNoImage',     count: assets.filter(a=>a.noImage).length },
    'apTile-nodoor':   { filter: a => a.noDoor,                                      label: 'No Door Assets',            val: 'apNoDoor',      count: assets.filter(a=>a.noDoor).length },
    'apTile-oem':      { filter: a => String(a.cabinetType).toUpperCase()==='OEM',   label: 'OEM Camera Assets',         val: 'apOEM',         count: assets.filter(a=>String(a.cabinetType).toUpperCase()==='OEM').length },
    'apTile-retrofit': { filter: a => String(a.cabinetType).toUpperCase()==='RETROFIT', label: 'Retrofit Camera Assets', val: 'apRetrofit',    count: assets.filter(a=>String(a.cabinetType).toUpperCase()==='RETROFIT').length },
  };

  Object.entries(tiles).forEach(([tileId, cfg]) => {
    const valEl = document.getElementById(cfg.val);
    if (valEl) valEl.textContent = cfg.count.toLocaleString();
    const tile = document.getElementById(tileId);
    if (!tile) return;
    // Remove old listener by cloning
    const fresh = tile.cloneNode(true);
    tile.parentNode.replaceChild(fresh, tile);
    // Re-set value (cloneNode copies the DOM, value already set above — refresh)
    const v = fresh.querySelector(`#${cfg.val}`);
    if (v) v.textContent = cfg.count.toLocaleString();
    fresh.addEventListener('click', () => {
      openDrill(cfg.label, () => buildAssetDetailTable(assets.filter(cfg.filter), cfg.label));
    });
  });
}

/* ---------- 15a. Address geocoding (Nominatim / localStorage cache) ---------- */
const GEOCACHE_KEY = 'micc_geocache_v2';

function loadGeoCache() {
  try { return JSON.parse(localStorage.getItem(GEOCACHE_KEY) || '{}'); } catch { return {}; }
}
function saveGeoCache(cache) {
  try { localStorage.setItem(GEOCACHE_KEY, JSON.stringify(cache)); } catch(e) {
    // localStorage full — clear old cache and retry
    try { localStorage.removeItem(GEOCACHE_KEY); localStorage.setItem(GEOCACHE_KEY, JSON.stringify(cache)); } catch {}
  }
}

async function geocodeCityUS(city, cache) {
  const key = city.trim().toLowerCase();
  if (key in cache) return cache[key];
  try {
    const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(city + ', USA')}&countrycodes=us&format=json&limit=1&addressdetails=0`;
    const resp = await fetch(url, { headers: { 'Accept-Language': 'en-US,en', 'User-Agent': 'MICC-USA-Dashboard/1.0' } });
    if (!resp.ok) { cache[key] = null; return null; }
    const data = await resp.json();
    const result = data.length ? { lat: parseFloat(data[0].lat), lon: parseFloat(data[0].lon) } : null;
    cache[key] = result;
    saveGeoCache(cache);
    return result;
  } catch { cache[key] = null; return null; }
}

async function geocodeAssetsFromCity(assets, onProgress) {
  const cache = loadGeoCache();
  const cities = [...new Set(assets.map(a => a.city).filter(c => c && c.trim().length > 1))];
  const uncached = cities.filter(c => !(c.trim().toLowerCase() in cache));
  let done = 0;
  for (const city of uncached) {
    await geocodeCityUS(city, cache);
    done++;
    if (onProgress) onProgress(done, uncached.length, cities.length - uncached.length);
    if (done < uncached.length) await new Promise(r => setTimeout(r, 1100)); // Nominatim: max 1 req/sec
  }
  // Assign coordinates + tiny jitter so assets in the same city don't perfectly overlap
  return assets.map(a => {
    if (a.lat != null) return a;
    if (!a.city || !a.city.trim()) return a;
    const geo = cache[a.city.trim().toLowerCase()];
    if (!geo) return a;
    const jitter = () => (Math.random() - 0.5) * 0.04; // ~2km spread
    return { ...a, lat: geo.lat + jitter(), lon: geo.lon + jitter() };
  });
}

/* ---------- 15b. Asset detail table & network map ---------- */
function buildAssetDetailTable(assets, title) {
  const wrap = document.createElement('div');
  if (!assets.length) {
    wrap.innerHTML = '<div class="empty-state">No assets in this category.</div>';
    return wrap;
  }
  const sorted = assets.slice().sort((a, b) => {
    const rank = r => r.noCommunication ? 0 : r.noImage ? 1 : r.noDoor ? 2 : 3;
    return rank(a) - rank(b);
  });
  const hdrRow = document.createElement('div');
  hdrRow.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:8px';
  const hdr = document.createElement('div');
  hdr.className = 'muted small';
  hdr.textContent = `${sorted.length.toLocaleString()} asset${sorted.length !== 1 ? 's' : ''} — click any row for full detail`;
  hdrRow.appendChild(hdr);
  // Export button added here — modal's auto-wiring also catches it, but this ensures correct filename
  const expBtn = exportBtn('Export CSV', () => tbl, safeFilename(title || 'assets'));
  hdrRow.appendChild(expBtn);
  wrap.appendChild(hdrRow);
  const tblWrap = document.createElement('div'); tblWrap.className = 'tbl-wrap';
  const tbl = document.createElement('table'); tbl.className = 'tbl';
  tbl.innerHTML = `<thead><tr>
    <th>Outlet</th><th>Outlet Code</th><th>Distributor</th>
    <th>Asset Serial</th><th>Camera</th><th>Status</th>
    <th>Last Image</th><th>Last Ping</th>
    <th class="num">Empty SoS</th><th class="num">Foreign SoS</th>
    <th>Diagnosis</th>
  </tr></thead>`;
  const tb = document.createElement('tbody');
  sorted.forEach(a => {
    const tr = document.createElement('tr');
    tr.className = 'expandable';
    tr.innerHTML = `
      <td><b>${a.outlet || '—'}</b></td>
      <td class="tight">${a.outletCode || '—'}</td>
      <td>${a.distributor || '—'}</td>
      <td class="tight">${a.asset || '—'}</td>
      <td>${a.cabinetType || '—'}</td>
      <td>${statusPill(a.status)}</td>
      <td class="tight">${fmtDate(a.lastImg)}</td>
      <td class="tight">${fmtDate(a.devicePing)}</td>
      <td class="num${a.emptySoS > 20 ? ' warn' : ''}">${fmtPct(a.emptySoS)}</td>
      <td class="num${a.foreignSoS > 20 ? ' warn' : ''}">${fmtPct(a.foreignSoS)}</td>
      <td style="font-size:11px">${a.diagnosis || '—'}</td>
    `;
    tr.addEventListener('click', () => openAssetDetail(a));
    tb.appendChild(tr);
  });
  tbl.appendChild(tb); tblWrap.appendChild(tbl); wrap.appendChild(tblWrap);
  return wrap;
}

async function renderAssetMap(assets) {
  const el = document.getElementById('assetMap');
  if (!el) return;

  if (state.assetLeafletMap) { state.assetLeafletMap.remove(); state.assetLeafletMap = null; }

  let positioned = assets.filter(a => a.lat != null && a.lon != null);

  // No GPS — fall back to city-level geocoding
  if (!positioned.length) {
    const hasCities = assets.some(a => a.city && a.city.trim().length > 1);
    if (!hasCities) {
      el.style.height = 'auto';
      el.innerHTML = `<div class="outlet-map-no-gps">
        <div style="font-size:28px;margin-bottom:8px">📍</div>
        <b>No location data found</b><br>
        <span class="muted">Asset Performance data needs a GPS column ("lat,lon") or a City column for address-based mapping.</span>
      </div>`;
      return;
    }
    const cache = loadGeoCache();
    const cities = [...new Set(assets.map(a => a.city).filter(c => c && c.trim().length > 1))];
    const uncachedCount = cities.filter(c => !(c.trim().toLowerCase() in cache)).length;

    el.style.height = '560px'; el.innerHTML = '';
    el.innerHTML = `<div class="geocode-progress">
      <div class="gp-icon">🗺️</div>
      <div class="gp-msg">Locating ${cities.length} cities from address data…</div>
      <div class="gp-bar-wrap"><div class="gp-bar" id="gpBar"></div></div>
      <div class="gp-sub" id="gpSub">${uncachedCount > 0 ? `${uncachedCount} new cities to look up (~${Math.ceil(uncachedCount / 55)} min first time)` : 'Loading from cache…'}</div>
      ${uncachedCount > 0 ? `<div class="gp-cache-note">Results are saved locally — subsequent loads will be instant.</div>` : ''}
    </div>`;

    const geocoded = await geocodeAssetsFromCity(assets, (done, total, fromCache) => {
      const pct = Math.round((done / total) * 100);
      const bar = document.getElementById('gpBar'); if (bar) bar.style.width = pct + '%';
      const sub = document.getElementById('gpSub'); if (sub) sub.textContent = `${done + fromCache} of ${total + fromCache} cities resolved`;
    });

    el.innerHTML = '';
    positioned = geocoded.filter(a => a.lat != null && a.lon != null);
    if (!positioned.length) {
      el.style.height = 'auto';
      el.innerHTML = `<div class="outlet-map-no-gps"><b>Could not geocode any cities.</b><br><span class="muted">Check that City values are valid US city names (e.g. "Atlanta" not "ATL").</span></div>`;
      return;
    }
  }

  // ── Draw map ──────────────────────────────────────────────────────────────
  el.style.height = '560px';
  const map = L.map(el, { preferCanvas: true });
  state.assetLeafletMap = map;

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>', maxZoom: 18,
  }).addTo(map);

  const statusColor = a => a.noCommunication ? '#E93D31' : a.noImage ? '#FF8500' : a.noDoor ? '#7A8AA6' : '#5AC581';

  positioned.forEach(a => {
    L.circleMarker([a.lat, a.lon], {
      radius: 7, fillColor: statusColor(a),
      color: 'rgba(0,0,0,0.18)', weight: 1, fillOpacity: 0.82,
    })
    .bindTooltip(`<b>${a.outlet}</b><br>${a.city ? a.city + ' · ' : ''}${a.distributor}<br>${statusPill(a.status)}<br>
      <span style="font-size:10.5px">Empty SoS: ${fmtPct(a.emptySoS)} · Foreign SoS: ${fmtPct(a.foreignSoS)}</span>`,
      { direction: 'top' })
    .on('click', () => openAssetDetail(a))
    .addTo(map);
  });

  map.fitBounds(L.latLngBounds(positioned.map(a => [a.lat, a.lon])), { padding: [30, 30] });
  setTimeout(() => map.invalidateSize(), 80);

  const legEl = document.getElementById('assetMapLegend');
  if (legEl) {
    const ct = { noComm: assets.filter(a=>a.noCommunication).length, noImage: assets.filter(a=>a.noImage).length, noDoor: assets.filter(a=>a.noDoor).length };
    ct.ok = assets.length - ct.noComm - ct.noImage - ct.noDoor;
    legEl.innerHTML = `
      <div class="leg-item"><div class="leg-dot" style="background:#E93D31"></div> No Communication (${ct.noComm.toLocaleString()})</div>
      <div class="leg-item"><div class="leg-dot" style="background:#FF8500"></div> No Image (${ct.noImage.toLocaleString()})</div>
      <div class="leg-item"><div class="leg-dot" style="background:#7A8AA6"></div> No Door (${ct.noDoor.toLocaleString()})</div>
      <div class="leg-item"><div class="leg-dot" style="background:#5AC581"></div> Active (${ct.ok.toLocaleString()})</div>
      <div class="leg-item" style="margin-left:auto;font-size:10.5px;color:var(--text-mute)">
        ${positioned.length.toLocaleString()} of ${assets.length.toLocaleString()} assets mapped · click any dot for detail
      </div>`;
  }
}

/* ---------- 16. Missed Revenue Drill Table ---------- */
const sortState = { tbl: null, key: null, dir: -1 };
function renderMissedRevDrill(rows) {
  const wrap = document.getElementById('missedRevTblWrap');
  if (!wrap) return;
  if (!rows.length) { wrap.innerHTML = '<div class="empty-state">No Missed Opportunity data in selected scope.</div>'; return; }
  // sort by missedRev desc
  const sorted = rows.slice().sort((a,b)=>b.missedRev-a.missedRev).slice(0, 300);
  const cols = [
    {h:'Distributor', k:'distributor'},
    {h:'Route', k:'route', cls:'tight'},
    {h:'Outlet', k:'outlet'},
    {h:'Outlet Code', k:'outletCode', cls:'tight'},
    {h:'Asset', k:'asset', cls:'tight'},
    {h:'Product', k:'product'},
    {h:'Case Price', k:'casePrice', cls:'num', fmt:fmtMoneyFull},
    {h:'Unit Price', k:'unitPrice', cls:'num', fmt:fmtMoneyFull},
    {h:'Avg Daily Units', k:'avgDailyUnits', cls:'num', fmt:v=>fmtNum(v,2)},
    {h:'OOS Days', k:'oosDays', cls:'num', fmt:v=>fmtNum(v,1)},
    {h:'Missed Units', k:'missedUnits', cls:'num', fmt:v=>fmtNum(v,1)},
    {h:'Missed Rev', k:'missedRev', cls:'num bad', fmt:fmtMoneyFull},
    {h:'Actual Rev', k:'actualRev', cls:'num', fmt:fmtMoneyFull},
    {h:'Potential Rev', k:'potentialRev', cls:'num', fmt:fmtMoneyFull},
    {h:'Planogram', k:'planogram'},
  ];
  wrap.innerHTML = '';
  const tblWrap = buildTable('missedRevTbl', sorted, cols, true);
  wrap.appendChild(tblWrap);

  // Export button in card-head
  const cardHead = wrap.closest('.card')?.querySelector('.card-head');
  if (cardHead) {
    cardHead.querySelectorAll('.tbl-export-btn').forEach(b => b.remove());
    cardHead.appendChild(exportBtn('Export CSV', () => document.getElementById('missedRevTbl'), 'missed_revenue_detail'));
  }
}

function buildTable(tblId, rows, cols, expandable=false) {
  const w = document.createElement('div'); w.className='tbl-wrap';
  const t = document.createElement('table'); t.className='tbl'; t.id=tblId;
  const thead = document.createElement('thead');
  const trH = document.createElement('tr');
  cols.forEach(c => {
    const th = document.createElement('th');
    th.innerHTML = `${c.h}<span class="sort-ind">⇅</span>`;
    if (c.cls && c.cls.indexOf('num') >= 0) th.classList.add('num');
    th.title = `Sort by ${c.h}`;
    th.addEventListener('click', () => sortTable(tblId, rows, cols, c.k, t));
    trH.appendChild(th);
  });
  thead.appendChild(trH);
  t.appendChild(thead);
  const tbody = document.createElement('tbody');
  rows.forEach((r, i) => {
    const tr = document.createElement('tr');
    if (expandable) {
      tr.className = 'expandable';
      tr.addEventListener('click', () => toggleExpand(tr, r));
    }
    cols.forEach(c => {
      const td = document.createElement('td');
      if (c.cls) td.className = c.cls;
      const val = r[c.k];
      td.textContent = c.fmt ? c.fmt(val) : (val == null ? '' : String(val));
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  t.appendChild(tbody);
  w.appendChild(t);
  return w;
}
function sortTable(tblId, rows, cols, key, t) {
  const dir = (sortState.tbl === tblId && sortState.key === key) ? -sortState.dir : -1;
  sortState.tbl = tblId; sortState.key = key; sortState.dir = dir;
  rows.sort((a,b) => {
    const va = a[key], vb = b[key];
    if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
    return String(va||'').localeCompare(String(vb||'')) * dir;
  });
  // Update sort indicators
  t.querySelectorAll('thead th').forEach((th, i) => {
    const ind = th.querySelector('.sort-ind');
    const isActive = cols[i] && cols[i].k === key;
    th.classList.toggle('sort-active', isActive);
    if (ind) ind.textContent = isActive ? (dir === -1 ? '▼' : '▲') : '⇅';
  });
  // re-render body
  const oldTbody = t.querySelector('tbody');
  const newTbody = document.createElement('tbody');
  rows.forEach(r => {
    const tr = document.createElement('tr');
    tr.className = 'expandable';
    tr.addEventListener('click', () => toggleExpand(tr, r));
    cols.forEach(c => {
      const td = document.createElement('td');
      if (c.cls) td.className = c.cls;
      const val = r[c.k];
      td.textContent = c.fmt ? c.fmt(val) : (val == null ? '' : String(val));
      tr.appendChild(td);
    });
    newTbody.appendChild(tr);
  });
  oldTbody.replaceWith(newTbody);
}

function toggleExpand(tr, r) {
  const next = tr.nextSibling;
  if (next && next.classList && next.classList.contains('expanded-row')) {
    next.remove();
    return;
  }
  const expRow = document.createElement('tr');
  expRow.className = 'expanded-row';
  const td = document.createElement('td');
  td.colSpan = tr.children.length;
  td.innerHTML = `
    <div class="expand-panel">
      <div class="grid-3" style="gap:18px">
        <div>
          <div class="muted small">Product</div>
          <div style="font-weight:700;font-size:13px">${r.product}</div>
          <div class="small muted">${r.brand} · Case size ${r.caseSize}</div>
        </div>
        <div>
          <div class="muted small">Pricing</div>
          <div>Case price: <b>${fmtMoneyFull(r.casePrice)}</b></div>
          <div>Unit price: <b>${fmtMoneyFull(r.unitPrice)}</b></div>
        </div>
        <div>
          <div class="muted small">Velocity</div>
          <div>Avg daily units: <b>${fmtNum(r.avgDailyUnits,2)}</b></div>
          <div>Avg daily cases: <b>${fmtNum(r.avgDailyCases,2)}</b></div>
          <div>Times refilled: <b>${r.timesRefilled}</b></div>
        </div>
        <div>
          <div class="muted small">Availability</div>
          <div>In-stock days: <b>${fmtNum(r.inStockDays,1)}</b></div>
          <div>OOS days: <b class="bad">${fmtNum(r.oosDays,1)}</b></div>
        </div>
        <div>
          <div class="muted small">Revenue</div>
          <div>Actual: <b>${fmtMoneyFull(r.actualRev)}</b></div>
          <div>Missed: <b class="bad">${fmtMoneyFull(r.missedRev)}</b></div>
          <div>Potential: <b>${fmtMoneyFull(r.potentialRev)}</b></div>
        </div>
        <div>
          <div class="muted small">Planogram</div>
          <div style="font-size:11.5px">${r.planogram || '—'}</div>
        </div>
      </div>
    </div>`;
  expRow.appendChild(td);
  tr.parentNode.insertBefore(expRow, tr.nextSibling);
}

/* ---------- 17. Top 10 products by avg OOS days ---------- */
function renderTPC(rows) {
  const el = document.getElementById('tpcCard');
  if (!el) return;

  el.innerHTML = `
    <div class="card-head">
      <div>
        <h3>Top 10 products — Avg OOS days</h3>
        <div class="descr">Top sellers by units sold · color = severity (green &lt;5d · orange 5–10d · red &gt;10d)</div>
      </div>
    </div>
    <div id="topOosWrap"></div>
  `;

  const wrap = document.getElementById('topOosWrap');
  if (!rows.length) {
    wrap.innerHTML = '<div class="empty-state">No Missed Opportunity data in selected scope.</div>';
    return;
  }

  const byProduct = {};
  rows.forEach(r => {
    if (!byProduct[r.product]) byProduct[r.product] = { brand: r.brand, unitsSold: 0, oosDaysSum: 0, n: 0, missedRev: 0 };
    byProduct[r.product].unitsSold  += r.unitsSold;
    byProduct[r.product].oosDaysSum += r.oosDays;
    byProduct[r.product].n          += 1;
    byProduct[r.product].missedRev  += r.missedRev;
  });

  const top10 = Object.entries(byProduct)
    .map(([p, v]) => ({ product: p, brand: v.brand, unitsSold: v.unitsSold, avgOosDays: v.oosDaysSum / v.n, missedRev: v.missedRev }))
    .sort((a, b) => b.unitsSold - a.unitsSold)
    .slice(0, 10);

  const maxOos = Math.max(...top10.map(p => p.avgOosDays), 1);

  wrap.innerHTML = top10.map((p, i) => {
    const barPct = (p.avgOosDays / maxOos) * 100;
    const color = p.avgOosDays >= 10 ? 'var(--red)' : p.avgOosDays >= 5 ? 'var(--orange)' : 'var(--green)';
    const shortName = p.product.length > 34 ? p.product.slice(0, 32) + '…' : p.product;
    return `
      <div class="oos-row">
        <div class="oos-rank">${i + 1}</div>
        <div class="oos-name" title="${p.product}">${shortName}</div>
        <div class="oos-bar-wrap">
          <div class="oos-bar" style="width:${barPct.toFixed(1)}%;background:${color}"></div>
        </div>
        <div class="oos-val" style="color:${color}">${fmtNum(p.avgOosDays, 1)}d</div>
        <div class="oos-miss muted">${fmtMoney(p.missedRev)}</div>
      </div>`;
  }).join('');
}

/* ---------- 18. Root Cause Relationship Panel ---------- */
function renderRootCausePanel(rows, assets) {
  const el = document.getElementById('rcPanel');
  if (!el) return;
  const k = computeKPIs(rows, assets);
  const scope = describeScope();
  el.innerHTML = `
    <div class="ctx-line">For <b>${scope}</b> during <b>${getPeriodLabel(state.primaryMonths)}</b>, the following related factors are observed (Playbook fields only — no inferred causes):</div>
    <div class="rc-grid">
      ${rcFact('Missed Revenue', fmtMoney(k.totalMissedRev), 'SUM(Missed Revenue) across ' + k.moRowCount + ' SKU-asset rows')}
      ${rcFact('Avg Days SKU OOS / Product', fmtNum(k.avgOOSDaysPerProduct,1)+' d', 'Mean OOS Days field')}
      ${rcFact('Empty Share of Shelf', k.emptySoS!=null ? fmtPct(k.emptySoS) : '—', 'Empty Facings / Total Facings')}
      ${rcFact('Foreign Share of Shelf', k.foreignSoS!=null ? fmtPct(k.foreignSoS) : '—', 'Foreign Facings / Total Facings')}
      ${rcFact('No Communication Assets', k.noCommAssets+' / '+k.totalAssets, 'Playbook status field')}
      ${rcFact('No Image Assets', k.noImageAssets+' / '+k.totalAssets, 'Playbook status field')}
      ${rcFact('Outlets in Scope', k.uniqueOutlets, 'Unique Outlet Codes')}
      ${rcFact('SKUs in Scope', k.uniqueProducts, 'Unique Products')}
    </div>
    <div class="muted small" style="margin-top:10px">These are related factors observed in the data, not causal claims.</div>
  `;
}
function rcFact(lbl, val, sub) { return `<div class="rc-fact"><div class="lbl">${lbl}</div><div class="val">${val}</div><div class="sub">${sub}</div></div>`; }
function describeScope() {
  const parts = [];
  const f = state.filters;
  if (f.distributor !== 'all') parts.push(`distributor ${f.distributor}`);
  if (f.outlet !== 'all') parts.push(`outlet ${f.outlet}`);
  if (f.route !== 'all') parts.push(`route ${f.route}`);
  if (f.product !== 'all') parts.push(`product ${f.product}`);
  if (!parts.length) parts.push('all MICC USA');
  return parts.join(', ');
}

/* ---------- 19. Before / After ---------- */
function renderBeforeAfter(k, kp) {
  const el = document.getElementById('beforeAfter');
  if (!el) return;
  if (!kp) {
    el.innerHTML = '<div class="empty-state">Pick a comparison period to enable before/after.</div>';
    return;
  }
  const lines = [
    ['Actual Revenue', kp.totalActualRev, k.totalActualRev, fmtMoney, 'up'],
    ['Missed Revenue', kp.totalMissedRev, k.totalMissedRev, fmtMoney, 'down'],
    ['Empty Share of Shelf', kp.emptySoS, k.emptySoS, v=>v==null?'—':fmtPct(v), 'down'],
    ['Foreign Share of Shelf', kp.foreignSoS, k.foreignSoS, v=>v==null?'—':fmtPct(v), 'down'],
    ['Own Share of Shelf', kp.ownSoS, k.ownSoS, v=>v==null?'—':fmtPct(v), 'up'],
    ['Avg Days SKU OOS / Product', kp.avgOOSDaysPerProduct, k.avgOOSDaysPerProduct, v=>fmtNum(v,1)+' d', 'down'],
    ['Total Assets Monitored', kp.totalAssets, k.totalAssets, v=>fmtNum(v,0), 'flat'],
    ['No Communication', kp.noCommAssets, k.noCommAssets, v=>fmtNum(v,0), 'down'],
    ['No Image', kp.noImageAssets, k.noImageAssets, v=>fmtNum(v,0), 'down'],
    ['No Door', kp.noDoorAssets, k.noDoorAssets, v=>fmtNum(v,0), 'down'],
  ];
  el.innerHTML = lines.map(([lbl, prior, cur, fmt, desired]) => {
    const d = pctDelta(cur, prior);
    let cls = 'flat', arrow = '◆';
    if (d != null && Math.abs(d) > 0.5) {
      const isGood = (desired === 'up' && d > 0) || (desired === 'down' && d < 0);
      cls = isGood ? 'up' : 'down';
      arrow = d > 0 ? '▲' : '▼';
    }
    const dStr = d == null ? '' : `<div class="kpi-delta ${cls}">${arrow} ${Math.abs(d).toFixed(1)}%</div>`;
    return `
      <div class="compare-row">
        <div class="compare-side">
          <div class="when">${getPeriodLabel(state.comparisonMonths)}</div>
          <div class="val">${fmt(prior)}</div>
        </div>
        <div class="compare-arrow">→</div>
        <div class="compare-side">
          <div class="when">${getPeriodLabel(state.primaryMonths)}</div>
          <div class="val">${fmt(cur)}</div>
          ${dStr}
        </div>
      </div>
      <div class="muted small" style="margin:-4px 4px 12px 4px">${lbl}</div>
    `;
  }).join('');
}

/* ---------- 20. DISTRIBUTOR VIEW ---------- */
function renderDistributor() {
  const primary = getRowsForMonths(state.primaryMonths);
  const assets = getAssetRows();
  renderOutletPriority(primary, assets);
  renderRouteView(primary, assets);
  renderMOQPanel(primary);
}

function renderOutletPriority(rows, assets) {
  const wrap = document.getElementById('outletPriorityWrap');
  if (!wrap) return;
  if (!rows.length) { wrap.innerHTML = '<div class="empty-state">No Missed Opportunity data in selected scope.</div>'; return; }
  // group by outletCode
  const byOutlet = {};
  rows.forEach(r => {
    const k = r.outletCode || r.outlet;
    if (!byOutlet[k]) byOutlet[k] = {
      outletCode: r.outletCode, outlet: r.outlet, distributor: r.distributor, route: r.route, asset: r.asset,
      products: new Set(), missedRev: 0, actualRev: 0, potentialRev: 0, oosDaysSum: 0, oosCount: 0,
    };
    byOutlet[k].products.add(r.product);
    byOutlet[k].missedRev += r.missedRev;
    byOutlet[k].actualRev += r.actualRev;
    byOutlet[k].potentialRev += r.potentialRev;
    byOutlet[k].oosDaysSum += r.oosDays;
    byOutlet[k].oosCount += 1;
  });
  // join asset perf info
  const assetByKey = {};
  assets.forEach(a => { assetByKey[a.asset] = a; });
  const list = Object.values(byOutlet).map(o => {
    const a = assetByKey[o.asset];
    return {
      ...o,
      productCount: o.products.size,
      avgOosDays: o.oosDaysSum / Math.max(1,o.oosCount),
      emptySoS: a ? a.emptySoS : null,
      foreignSoS: a ? a.foreignSoS : null,
      ownSoS: a ? Math.max(0,100-a.emptySoS-a.foreignSoS) : null,
      lastImg: a ? a.lastImg : null,
      status: a ? a.status : '',
      noCommunication: a ? a.noCommunication : false,
      noImage: a ? a.noImage : false,
    };
  }).sort((a,b)=>b.missedRev-a.missedRev);

  const cols = [
    {h:'Outlet', k:'outlet'},
    {h:'Outlet Code', k:'outletCode', cls:'tight'},
    {h:'Distributor', k:'distributor'},
    {h:'Route', k:'route', cls:'tight'},
    {h:'Asset', k:'asset', cls:'tight'},
    {h:'SKUs Affected', k:'productCount', cls:'num', fmt:v=>fmtNum(v,0)},
    {h:'Missed Rev', k:'missedRev', cls:'num bad', fmt:fmtMoneyFull},
    {h:'Actual Rev', k:'actualRev', cls:'num', fmt:fmtMoneyFull},
    {h:'Potential Rev', k:'potentialRev', cls:'num', fmt:fmtMoneyFull},
    {h:'Avg OOS Days', k:'avgOosDays', cls:'num', fmt:v=>fmtNum(v,1)},
    {h:'Empty SoS %', k:'emptySoS', cls:'num warn', fmt:v=>v==null?'—':fmtPct(v)},
    {h:'Foreign SoS %', k:'foreignSoS', cls:'num', fmt:v=>v==null?'—':fmtPct(v)},
    {h:'Last Image', k:'lastImg', cls:'tight', fmt:fmtDate},
    {h:'Status', k:'status', fmt: v => statusPill(v)},
  ];
  state.currentOutletList = list;

  // Show map toggle when GPS data exists
  const hasGps = list.some(o => { const a = state.assetPerf.find(ap => ap.asset === o.asset); return a && a.lat != null; });
  const toggle = document.getElementById('outletViewToggle');
  if (toggle) toggle.style.display = hasGps ? 'flex' : 'none';

  wrap.innerHTML = '';
  wrap.appendChild(buildTableWithExpand('outletPriorityTbl', list, cols, (o) => buildOutletExpand(o, rows, assets)));

  // Export button
  const cardHead = wrap.closest('.card')?.querySelector('.card-head');
  if (cardHead) {
    cardHead.querySelectorAll('.tbl-export-btn').forEach(b => b.remove());
    cardHead.appendChild(exportBtn('Export CSV', () => document.getElementById('outletPriorityTbl'), 'outlet_priority'));
  }
}

function statusPill(v) {
  if (!v) return '';
  const s = String(v).toLowerCase();
  if (s.includes('no communication')) return `<span class="pill red">${v}</span>`;
  if (s.includes('no image')) return `<span class="pill orange">${v}</span>`;
  if (s.includes('no door')) return `<span class="pill gray">${v}</span>`;
  if (s.includes('fine')) return `<span class="pill green">${v}</span>`;
  return `<span class="pill blue">${v}</span>`;
}

function buildTableWithExpand(tblId, rows, cols, expandBuilder) {
  const w = document.createElement('div'); w.className='tbl-wrap';
  const t = document.createElement('table'); t.className='tbl'; t.id=tblId;
  const thead = document.createElement('thead');
  const trH = document.createElement('tr');
  cols.forEach(c => {
    const th = document.createElement('th');
    th.innerHTML = `${c.h}<span class="sort-ind">⇅</span>`;
    if (c.cls && c.cls.indexOf('num') >= 0) th.classList.add('num');
    th.title = `Sort by ${c.h}`;
    th.addEventListener('click', () => sortGenericTable(rows, cols, c.k, t, expandBuilder));
    trH.appendChild(th);
  });
  thead.appendChild(trH);
  t.appendChild(thead);
  const tbody = document.createElement('tbody');
  rows.forEach(r => {
    const tr = buildRow(r, cols);
    tr.addEventListener('click', () => toggleGenericExpand(tr, r, expandBuilder));
    tbody.appendChild(tr);
  });
  t.appendChild(tbody);
  w.appendChild(t);
  return w;
}
function buildRow(r, cols) {
  const tr = document.createElement('tr');
  tr.className = 'expandable';
  cols.forEach(c => {
    const td = document.createElement('td');
    if (c.cls) td.className = c.cls;
    const val = r[c.k];
    const html = c.fmt ? c.fmt(val) : (val == null ? '' : String(val));
    if (html != null && /<[a-z]/.test(html)) td.innerHTML = html;
    else td.textContent = html;
    tr.appendChild(td);
  });
  return tr;
}
function sortGenericTable(rows, cols, key, t, expandBuilder) {
  const dir = (sortState.key === key) ? -sortState.dir : -1;
  sortState.key = key; sortState.dir = dir;
  rows.sort((a,b) => {
    const va = a[key], vb = b[key];
    if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
    return String(va||'').localeCompare(String(vb||'')) * dir;
  });
  // Update sort indicators
  t.querySelectorAll('thead th').forEach((th, i) => {
    const ind = th.querySelector('.sort-ind');
    const isActive = cols[i] && cols[i].k === key;
    th.classList.toggle('sort-active', isActive);
    if (ind) ind.textContent = isActive ? (dir === -1 ? '▼' : '▲') : '⇅';
  });
  const oldTbody = t.querySelector('tbody');
  const newTbody = document.createElement('tbody');
  rows.forEach(r => {
    const tr = buildRow(r, cols);
    tr.addEventListener('click', () => toggleGenericExpand(tr, r, expandBuilder));
    newTbody.appendChild(tr);
  });
  oldTbody.replaceWith(newTbody);
}
function toggleGenericExpand(tr, r, expandBuilder) {
  const next = tr.nextSibling;
  if (next && next.classList && next.classList.contains('expanded-row')) { next.remove(); return; }
  const expRow = document.createElement('tr');
  expRow.className = 'expanded-row';
  const td = document.createElement('td');
  td.colSpan = tr.children.length;
  td.innerHTML = `<div class="expand-panel">${expandBuilder(r)}</div>`;
  expRow.appendChild(td);
  tr.parentNode.insertBefore(expRow, tr.nextSibling);
}

function buildOutletExpand(o, rows, assets) {
  const moRows = rows.filter(r => (r.outletCode||r.outlet) === (o.outletCode||o.outlet)).sort((a,b)=>b.missedRev-a.missedRev);
  const asset = assets.find(a => a.asset === o.asset);
  let imgTimeline = '';
  if (asset && asset.dailyImg && asset.dailyImg.length) {
    imgTimeline = `<div class="muted small" style="margin-bottom:4px">Last 7 days image activity:</div><div class="img-timeline">` +
      asset.dailyImg.map(d => {
        const cls = d.images === 0 ? 'none' : d.images < 2 ? 'low' : d.images < 5 ? 'ok' : 'high';
        return `<div class="day ${cls}" title="${d.date}: ${d.images} images / ${d.doors} doors">${d.images}</div>`;
      }).join('') + '</div>';
  }
  const top = moRows.slice(0,8).map(r =>
    `<tr>
      <td>${r.product}</td>
      <td class="num">${fmtMoneyFull(r.unitPrice)}</td>
      <td class="num">${fmtNum(r.avgDailyUnits,2)}</td>
      <td class="num">${fmtNum(r.oosDays,1)}</td>
      <td class="num bad">${fmtMoneyFull(r.missedRev)}</td>
      <td class="num">${fmtMoneyFull(r.actualRev)}</td>
    </tr>`).join('');
  return `
    <div class="grid-2-1">
      <div>
        <div class="muted small" style="margin-bottom:4px">Top SKUs by missed revenue at ${o.outlet}:</div>
        <table class="tbl" style="font-size:11.5px">
          <thead><tr><th>Product</th><th class="num">Unit Price</th><th class="num">Avg Daily Units</th><th class="num">OOS Days</th><th class="num">Missed Rev</th><th class="num">Actual Rev</th></tr></thead>
          <tbody>${top}</tbody>
        </table>
      </div>
      <div>
        ${imgTimeline}
        <div class="muted small" style="margin-top:10px">Asset details:</div>
        <div class="small" style="line-height:1.7">
          ${asset ? `
            Serial: <b>${asset.asset}</b><br>
            Camera type: <b>${asset.cabinetType || '—'}</b><br>
            Last image: <b>${fmtDate(asset.lastImg)}</b><br>
            Empty SoS: <b>${fmtPct(asset.emptySoS)}</b><br>
            Foreign SoS: <b>${fmtPct(asset.foreignSoS)}</b><br>
            Status: ${statusPill(asset.status)}
          ` : '<span class="muted">No Asset Performance row found for this asset.</span>'}
        </div>
      </div>
    </div>`;
}

/* ---------- 20b. Outlet Map ---------- */
function renderOutletMap(list) {
  const el = document.getElementById('outletMap');
  if (!el) return;

  if (state.outletLeafletMap) { state.outletLeafletMap.remove(); state.outletLeafletMap = null; }

  // Join outlet list with asset GPS data
  const assetGps = {};
  state.assetPerf.forEach(a => { if (a.lat != null) assetGps[a.asset] = a; });
  const mappable = list.map(o => ({ ...o, _ap: assetGps[o.asset] })).filter(o => o._ap && o._ap.lat != null);

  if (!mappable.length) {
    el.style.height = 'auto';
    el.innerHTML = `<div class="outlet-map-no-gps">
      <b>No GPS coordinates found</b><br>
      <span class="muted">Upload Asset Performance data with GPS / Latitude / Longitude columns to enable the map view.</span>
    </div>`;
    return;
  }
  el.style.height = '520px'; el.innerHTML = '';

  const map = L.map(el, { preferCanvas: true });
  state.outletLeafletMap = map;
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>', maxZoom: 18,
  }).addTo(map);

  const maxMissed = Math.max(...mappable.map(o => o.missedRev), 1);
  mappable.forEach(o => {
    const a = o._ap;
    const color = a.noCommunication ? '#E93D31' : a.noImage ? '#FF8500' : a.noDoor ? '#7A8AA6' : '#1A29B6';
    const r = 6 + Math.round((o.missedRev / maxMissed) * 10);
    const popup = L.popup({ maxWidth: 280 }).setContent(`
      <div class="map-popup-head">${o.outlet}</div>
      <div class="map-popup-body">
        <div class="map-popup-row"><span class="lbl">Distributor</span><span class="val">${o.distributor}</span></div>
        <div class="map-popup-row"><span class="lbl">Missed Rev</span><span class="val bad">${fmtMoneyFull(o.missedRev)}</span></div>
        <div class="map-popup-row"><span class="lbl">Actual Rev</span><span class="val">${fmtMoneyFull(o.actualRev)}</span></div>
        <div class="map-popup-row"><span class="lbl">Avg OOS Days</span><span class="val">${fmtNum(o.avgOosDays,1)}</span></div>
        <div class="map-popup-row"><span class="lbl">Empty SoS</span><span class="val${a.emptySoS>15?' warn':''}">${fmtPct(a.emptySoS)}</span></div>
        <div class="map-popup-row"><span class="lbl">Status</span><span class="val">${statusPill(a.status)}</span></div>
      </div>`);
    L.circleMarker([a.lat, a.lon], {
      radius: r, fillColor: color, color: 'rgba(0,0,0,0.2)', weight: 1, fillOpacity: 0.85,
    }).bindPopup(popup).addTo(map);
  });

  map.fitBounds(L.latLngBounds(mappable.map(o => [o._ap.lat, o._ap.lon])), { padding: [30, 30] });

  const legEl = document.getElementById('outletMapLegend') || (() => {
    const d = document.createElement('div'); d.id = 'outletMapLegend'; d.className = 'outlet-map-legend';
    el.parentElement.appendChild(d); return d;
  })();
  legEl.innerHTML = `
    <div class="leg-item"><div class="leg-dot" style="background:#E93D31"></div> No Communication</div>
    <div class="leg-item"><div class="leg-dot" style="background:#FF8500"></div> No Image</div>
    <div class="leg-item"><div class="leg-dot" style="background:#7A8AA6"></div> No Door</div>
    <div class="leg-item"><div class="leg-dot" style="background:#1A29B6"></div> Active</div>
    <div class="leg-item" style="margin-left:auto;font-size:10.5px;color:var(--text-mute)">Dot size = Missed Revenue · ${mappable.length} outlets mapped</div>`;
}

/* ---------- 21. Route view ---------- */
function renderRouteView(rows, assets) {
  const wrap = document.getElementById('routeViewWrap');
  if (!wrap) return;
  const hasRoutes = rows.some(r => r.route);
  if (!rows.length || !hasRoutes) { wrap.innerHTML = '<div class="empty-state">No Route information available in current scope.</div>'; return; }
  const byRoute = {};
  rows.forEach(r => {
    const k = String(r.route || 'Unassigned');
    if (!byRoute[k]) byRoute[k] = {
      route: k, distributor: r.distributor,
      outlets: new Set(), products: new Set(),
      missedRev: 0, actualRev: 0, oosDaysSum: 0, oosCount: 0,
    };
    byRoute[k].outlets.add(r.outletCode || r.outlet);
    byRoute[k].products.add(r.product);
    byRoute[k].missedRev += r.missedRev;
    byRoute[k].actualRev += r.actualRev;
    byRoute[k].oosDaysSum += r.oosDays;
    byRoute[k].oosCount += 1;
  });
  // asset perf summary by joining via outletCode
  const assetsByOutlet = {};
  assets.forEach(a => { assetsByOutlet[a.outletCode || a.outlet] = a; });

  const list = Object.values(byRoute).map(r => {
    const empties = [...r.outlets].map(o => assetsByOutlet[o]).filter(Boolean).map(a => a.emptySoS);
    const fsos = [...r.outlets].map(o => assetsByOutlet[o]).filter(Boolean).map(a => a.foreignSoS);
    return {
      ...r,
      outletCount: r.outlets.size,
      productCount: r.products.size,
      avgOosDays: r.oosDaysSum / Math.max(1, r.oosCount),
      avgEmptySoS: empties.length ? empties.reduce((s,v)=>s+v,0)/empties.length : null,
      avgFSOS: fsos.length ? fsos.reduce((s,v)=>s+v,0)/fsos.length : null,
    };
  }).sort((a,b)=>b.missedRev-a.missedRev);

  const cols = [
    {h:'Route', k:'route', cls:'tight'},
    {h:'Distributor', k:'distributor'},
    {h:'Outlets', k:'outletCount', cls:'num', fmt:v=>fmtNum(v,0)},
    {h:'SKUs', k:'productCount', cls:'num', fmt:v=>fmtNum(v,0)},
    {h:'Missed Rev', k:'missedRev', cls:'num bad', fmt:fmtMoneyFull},
    {h:'Actual Rev', k:'actualRev', cls:'num', fmt:fmtMoneyFull},
    {h:'Avg OOS Days', k:'avgOosDays', cls:'num', fmt:v=>fmtNum(v,1)},
    {h:'Avg Empty SoS', k:'avgEmptySoS', cls:'num', fmt:v=>v==null?'—':fmtPct(v)},
    {h:'Avg Foreign SoS', k:'avgFSOS', cls:'num', fmt:v=>v==null?'—':fmtPct(v)},
  ];
  wrap.innerHTML = '';
  wrap.appendChild(buildTable('routeTbl', list, cols, false));

  const cardHead = wrap.closest('.card')?.querySelector('.card-head');
  if (cardHead) {
    cardHead.querySelectorAll('.tbl-export-btn').forEach(b => b.remove());
    cardHead.appendChild(exportBtn('Export CSV', () => document.getElementById('routeTbl'), 'route_view'));
  }
}

/* ---------- 22. MOQ / MOV panel ---------- */
function renderMOQPanel(rows) {
  const moqEl = document.getElementById('moqOutlets');
  const movEl = document.getElementById('movRoutes');
  if (!moqEl || !movEl) return;
  const movThreshold = num(document.getElementById('movInput').value);
  const routeMov = num(document.getElementById('routeMovInput').value);
  state.thresholds.mov = movThreshold;
  state.thresholds.routeMov = routeMov;

  // outlets below threshold by potentialRev (we use missedRev as proxy for required-order indicator)
  const byOutlet = {};
  rows.forEach(r => {
    const k = r.outletCode || r.outlet;
    if (!byOutlet[k]) byOutlet[k] = { outlet:r.outlet, outletCode:r.outletCode, distributor:r.distributor, route:r.route, missedRev:0, actualRev:0, productCount:0, products:new Set() };
    byOutlet[k].missedRev += r.missedRev;
    byOutlet[k].actualRev += r.actualRev;
    byOutlet[k].products.add(r.product);
  });
  const outlets = Object.values(byOutlet).map(o => ({ ...o, productCount:o.products.size }));
  const belowMov = outlets.filter(o => o.missedRev > 0 && o.missedRev < movThreshold).sort((a,b)=>b.missedRev-a.missedRev);
  const aboveMov = outlets.filter(o => o.missedRev >= movThreshold).sort((a,b)=>b.missedRev-a.missedRev);

  moqEl.innerHTML = `<div class="muted small" style="margin-bottom:6px">${belowMov.length} outlet(s) below $${movThreshold} missed-revenue threshold (cannot justify individual service)</div>`;
  if (belowMov.length) {
    const tblWrap = buildTable('movTbl', belowMov.slice(0,30), [
      {h:'Outlet', k:'outlet'},{h:'Route', k:'route', cls:'tight'},
      {h:'Distributor', k:'distributor'},{h:'Missed Rev', k:'missedRev', cls:'num bad', fmt:fmtMoneyFull},
      {h:'SKUs', k:'productCount', cls:'num', fmt:v=>fmtNum(v,0)},
    ]);
    const row = document.createElement('div'); row.style.cssText = 'display:flex;justify-content:flex-end;margin-bottom:4px';
    row.appendChild(exportBtn('Export CSV', () => document.getElementById('movTbl'), 'outlets_below_mov'));
    moqEl.appendChild(row); moqEl.appendChild(tblWrap);
  } else {
    moqEl.innerHTML += '<div class="empty-state small">None below threshold.</div>';
  }

  // routes above threshold (potential bundled service)
  const byRoute = {};
  outlets.forEach(o => {
    const k = String(o.route || 'Unassigned');
    if (!byRoute[k]) byRoute[k] = { route:k, distributor:o.distributor, missedRev:0, outletCount:0 };
    byRoute[k].missedRev += o.missedRev;
    byRoute[k].outletCount += 1;
  });
  const routesAbove = Object.values(byRoute).filter(r => r.missedRev >= routeMov).sort((a,b)=>b.missedRev-a.missedRev);

  movEl.innerHTML = `<div class="muted small" style="margin-bottom:6px">${routesAbove.length} route(s) above $${routeMov} bundled missed-revenue threshold (may justify route-level service)</div>`;
  if (routesAbove.length) {
    const tblWrap = buildTable('routeMovTbl', routesAbove.slice(0,30), [
      {h:'Route', k:'route', cls:'tight'},{h:'Distributor', k:'distributor'},
      {h:'Outlets', k:'outletCount', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'Bundled Missed Rev', k:'missedRev', cls:'num bad', fmt:fmtMoneyFull},
    ]);
    const row = document.createElement('div'); row.style.cssText = 'display:flex;justify-content:flex-end;margin-bottom:4px';
    row.appendChild(exportBtn('Export CSV', () => document.getElementById('routeMovTbl'), 'routes_above_mov'));
    movEl.appendChild(row); movEl.appendChild(tblWrap);
  } else {
    movEl.innerHTML += '<div class="empty-state small">No routes above threshold.</div>';
  }
}

/* ---------- 23. Modal drill-downs ---------- */
function openDrill(title, contentBuilder) {
  document.getElementById('modalTitle').textContent = title;
  const body = document.getElementById('modalBody');
  body.innerHTML = '';
  const content = contentBuilder();
  if (typeof content === 'string') body.innerHTML = content;
  else if (content) body.appendChild(content);

  // Wire modal export button — finds the first table in the modal body
  const exportBtn = document.getElementById('modalExportBtn');
  if (exportBtn) {
    const tbl = body.querySelector('table.tbl');
    if (tbl) {
      exportBtn.style.display = 'inline-flex';
      exportBtn.onclick = () => downloadCSV(tableToCSV(tbl), safeFilename(title) + '.csv');
    } else {
      exportBtn.style.display = 'none';
    }
  }

  document.getElementById('modal').classList.add('show');
}
function closeDrill() {
  document.getElementById('modal').classList.remove('show');
  const exportBtn = document.getElementById('modalExportBtn');
  if (exportBtn) exportBtn.style.display = 'none';
}

function drillRevenue(rows, key) {
  const wrap = document.createElement('div');
  // by distributor
  const byD = {};
  rows.forEach(r => { byD[r.distributor] = (byD[r.distributor]||0) + r[key]; });
  const sorted = Object.entries(byD).sort((a,b)=>b[1]-a[1]);
  const tbl = document.createElement('table'); tbl.className='tbl';
  tbl.innerHTML = `<thead><tr><th>Distributor</th><th class="num">${key === 'missedRev' ? 'Missed Revenue' : 'Actual Revenue'}</th><th class="num">% of total</th></tr></thead>`;
  const tb = document.createElement('tbody');
  const tot = sorted.reduce((s,v)=>s+v[1],0);
  sorted.forEach(([d,v]) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${d}</td><td class="num">${fmtMoneyFull(v)}</td><td class="num">${((v/tot)*100).toFixed(1)}%</td>`;
    tb.appendChild(tr);
  });
  tbl.appendChild(tb);
  wrap.appendChild(tbl);
  return wrap;
}
function drillSoS(assets, field, label) {
  const wrap = document.createElement('div');
  if (!assets.length) { wrap.innerHTML = '<div class="empty-state">No Asset Performance data uploaded.</div>'; return wrap; }
  const byD = {};
  assets.forEach(a => {
    let val;
    if (field === 'ownSoS') val = Math.max(0,100 - a.emptySoS - a.foreignSoS);
    else val = a[field];
    if (!byD[a.distributor]) byD[a.distributor] = { sum:0, n:0 };
    byD[a.distributor].sum += val;
    byD[a.distributor].n += 1;
  });
  const sorted = Object.entries(byD).map(([k,v])=>[k, v.sum/v.n]).sort((a,b)=>b[1]-a[1]);
  const tbl = document.createElement('table'); tbl.className='tbl';
  tbl.innerHTML = `<thead><tr><th>Distributor</th><th class="num">Avg ${label}</th><th class="num">Assets</th></tr></thead>`;
  const tb = document.createElement('tbody');
  sorted.forEach(([d,v]) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${d}</td><td class="num">${fmtPct(v)}</td><td class="num">${byD[d].n}</td>`;
    tb.appendChild(tr);
  });
  tbl.appendChild(tb); wrap.appendChild(tbl);
  return wrap;
}
function drillOOS(rows) {
  const wrap = document.createElement('div');
  const byProduct = {};
  rows.forEach(r => {
    if (!byProduct[r.product]) byProduct[r.product] = { sum:0, n:0, missed:0 };
    byProduct[r.product].sum += r.oosDays;
    byProduct[r.product].n += 1;
    byProduct[r.product].missed += r.missedRev;
  });
  const sorted = Object.entries(byProduct).map(([k,v])=>[k, v.sum/v.n, v.missed, v.n]).sort((a,b)=>b[1]-a[1]).slice(0,40);
  const tbl = document.createElement('table'); tbl.className='tbl';
  tbl.innerHTML = `<thead><tr><th>Product</th><th class="num">Avg OOS Days</th><th class="num">SKU-asset rows</th><th class="num">Total Missed Rev</th></tr></thead>`;
  const tb = document.createElement('tbody');
  sorted.forEach(([p, avg, missed, n]) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${p}</td><td class="num">${fmtNum(avg,1)}</td><td class="num">${n}</td><td class="num bad">${fmtMoneyFull(missed)}</td>`;
    tb.appendChild(tr);
  });
  tbl.appendChild(tb); wrap.appendChild(tbl);
  return wrap;
}
function drillAssets(assets) {
  const wrap = document.createElement('div');
  if (!assets.length) { wrap.innerHTML = '<div class="empty-state">No Asset Performance data uploaded.</div>'; return wrap; }
  const sorted = assets.slice().sort((a,b) => {
    const rank = r => r.noCommunication ? 0 : r.noImage ? 1 : r.noDoor ? 2 : 3;
    return rank(a) - rank(b);
  }).slice(0,200);
  const tbl = document.createElement('table'); tbl.className='tbl';
  tbl.innerHTML = `<thead><tr><th>Asset</th><th>Outlet</th><th>Distributor</th><th>Camera</th><th>Status</th><th>Last Image</th><th class="num">Empty SoS</th><th class="num">Foreign SoS</th></tr></thead>`;
  const tb = document.createElement('tbody');
  sorted.forEach(a => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="tight">${a.asset}</td><td>${a.outlet}</td><td>${a.distributor}</td><td>${a.cabinetType||'—'}</td><td>${statusPill(a.status)}</td><td class="tight">${fmtDate(a.lastImg)}</td><td class="num">${fmtPct(a.emptySoS)}</td><td class="num">${fmtPct(a.foreignSoS)}</td>`;
    tb.appendChild(tr);
  });
  tbl.appendChild(tb); wrap.appendChild(tbl);
  return wrap;
}
function openAssetDetail(a) {
  document.getElementById('modalTitle').textContent = `Asset ${a.asset} · ${a.outlet}`;
  let dailyHtml = '';
  if (a.dailyImg && a.dailyImg.length) {
    dailyHtml = `<div class="muted small" style="margin:14px 0 4px">Rolling 7-day image activity:</div><div class="img-timeline">` +
      a.dailyImg.map(d => {
        const cls = d.images === 0 ? 'none' : d.images < 2 ? 'low' : d.images < 5 ? 'ok' : 'high';
        return `<div class="day ${cls}" title="${d.date}: ${d.images} images / ${d.doors} doors">${d.images}</div>`;
      }).join('') + '</div>';
  }
  document.getElementById('modalBody').innerHTML = `
    <div class="grid-3" style="gap:18px">
      <div><div class="muted small">Distributor</div><div><b>${a.distributor}</b></div></div>
      <div><div class="muted small">Outlet</div><div><b>${a.outlet}</b> · ${a.outletCode||'—'}</div></div>
      <div><div class="muted small">Status</div><div>${statusPill(a.status)}</div></div>
      <div><div class="muted small">Camera Type</div><div>${a.cabinetType||'—'}</div></div>
      <div><div class="muted small">Hardware Status</div><div>${a.hardwareStatus||'—'}</div></div>
      <div><div class="muted small">Diagnosis</div><div>${a.diagnosis||'—'}</div></div>
      <div><div class="muted small">Last Image</div><div>${fmtDate(a.lastImg)}</div></div>
      <div><div class="muted small">Last Device Ping</div><div>${fmtDate(a.devicePing)}</div></div>
      <div><div class="muted small">Installed</div><div>${fmtDate(a.installedOn)}</div></div>
      <div><div class="muted small">Empty SoS</div><div><b>${fmtPct(a.emptySoS)}</b></div></div>
      <div><div class="muted small">Foreign SoS</div><div><b>${fmtPct(a.foreignSoS)}</b></div></div>
      <div><div class="muted small">Total Facings</div><div>${a.facings}</div></div>
    </div>
    ${dailyHtml}
  `;
  document.getElementById('modal').classList.add('show');
}

/* ---------- 23b. Multi-level KPI drill-down navigator ---------- */

const drillNav = { stack: [], metric: null, moRows: [], apRows: [] };

const DRILL_LEVELS = {
  actualRev:   ['National', 'Distributor', 'Outlet', 'SKU Detail'],
  missedRev:   ['National', 'Distributor', 'Outlet', 'SKU Detail'],
  oosDays:     ['National', 'Distributor', 'Outlet', 'SKU Detail'],
  emptySoS:    ['National', 'Distributor', 'Outlet', 'Asset Detail'],
  foreignSoS:  ['National', 'Distributor', 'Outlet', 'Asset Detail'],
  assetHealth: ['National', 'Distributor', 'Outlet', 'Asset Detail'],
};
const DRILL_LABELS = {
  actualRev:'Actual Revenue', missedRev:'Missed Revenue', oosDays:'Avg OOS Days',
  emptySoS:'Empty Share of Shelf', foreignSoS:'Foreign Share of Shelf', assetHealth:'Asset Health',
};

function openMetricDrill(metric, moRows, apRows) {
  drillNav.stack = [];
  drillNav.metric = metric;
  drillNav.moRows = moRows;
  drillNav.apRows = apRows;
  drillNav.stack.push({ title: 'All Distributors', fn: () => drillLevel1(metric, moRows, apRows) });
  renderDrillModal();
  document.getElementById('modal').classList.add('show');
}

function drillPush(title, fn) {
  drillNav.stack.push({ title, fn });
  renderDrillModal();
}

function renderDrillModal() {
  const cur = drillNav.stack[drillNav.stack.length - 1];
  const metric = drillNav.metric;
  const body = document.getElementById('modalBody');
  body.innerHTML = '';
  document.getElementById('modalTitle').textContent = DRILL_LABELS[metric] || metric;

  // ── Path progress bar ────────────────────────────────────────────────────
  const levels = DRILL_LEVELS[metric] || [];
  const pathBar = document.createElement('div');
  pathBar.className = 'drill-path-bar';
  levels.forEach((lbl, i) => {
    const step = document.createElement('div');
    step.className = 'dpi-step' + (i < drillNav.stack.length - 1 ? ' done' : i === drillNav.stack.length - 1 ? ' current' : '');
    step.innerHTML = `<div class="dpi-dot"></div><div class="dpi-label">${lbl}</div>`;
    // Click done steps to go back to that level
    if (i < drillNav.stack.length - 1) {
      step.style.cursor = 'pointer';
      step.title = `Back to ${lbl}`;
      step.addEventListener('click', () => { drillNav.stack.splice(i + 1); renderDrillModal(); });
    }
    pathBar.appendChild(step);
    if (i < levels.length - 1) {
      const conn = document.createElement('div');
      conn.className = 'dpi-connector' + (i < drillNav.stack.length - 1 ? ' done' : '');
      pathBar.appendChild(conn);
    }
  });
  body.appendChild(pathBar);

  // ── Breadcrumb back nav ──────────────────────────────────────────────────
  if (drillNav.stack.length > 1) {
    const nav = document.createElement('div');
    nav.className = 'drill-nav-bar';
    const back = document.createElement('button');
    back.className = 'drill-back-btn';
    back.textContent = '← Back';
    back.onclick = () => { drillNav.stack.pop(); renderDrillModal(); };
    nav.appendChild(back);
    const crumbs = document.createElement('div');
    crumbs.className = 'drill-crumbs';
    drillNav.stack.forEach((item, i) => {
      if (i > 0) crumbs.insertAdjacentHTML('beforeend', '<span class="drill-sep">›</span>');
      if (i < drillNav.stack.length - 1) {
        const btn = document.createElement('button');
        btn.className = 'drill-crumb-link'; btn.textContent = item.title;
        btn.onclick = () => { drillNav.stack.splice(i + 1); renderDrillModal(); };
        crumbs.appendChild(btn);
      } else {
        crumbs.insertAdjacentHTML('beforeend', `<span class="drill-crumb-current">${item.title}</span>`);
      }
    });
    nav.appendChild(crumbs);
    body.appendChild(nav);
  }

  // ── Content ───────────────────────────────────────────────────────────────
  const content = cur.fn();
  if (content) body.appendChild(content);

  // ── Export wiring ─────────────────────────────────────────────────────────
  const expBtn = document.getElementById('modalExportBtn');
  if (expBtn) {
    const tbl = body.querySelector('table.tbl');
    expBtn.style.display = tbl ? 'inline-flex' : 'none';
    if (tbl) expBtn.onclick = () => downloadCSV(tableToCSV(tbl), safeFilename(cur.title) + '.csv');
  }
}

// ── Drill table builder — sortable columns + optional row drill ───────────────
function buildDrillTable(id, rows, cols, onRowClick) {
  const localSort = { key: null, dir: -1 };
  const wrap = document.createElement('div'); wrap.className = 'drill-tbl-wrap';
  const tbl  = document.createElement('table'); tbl.className = 'tbl'; tbl.id = id;
  const thead = document.createElement('thead');
  const trH   = document.createElement('tr');

  const rebuildBody = () => {
    const tbody = document.createElement('tbody');
    rows.forEach(r => {
      const tr = document.createElement('tr');
      if (onRowClick) tr.className = 'drill-row';
      cols.forEach(c => {
        const td = document.createElement('td'); if (c.cls) td.className = c.cls;
        const v = r[c.k]; const html = c.fmt ? c.fmt(v) : (v == null ? '' : String(v));
        if (html && /<[a-z]/i.test(String(html))) td.innerHTML = html; else td.textContent = html;
        tr.appendChild(td);
      });
      if (onRowClick) {
        const arrowTd = document.createElement('td');
        arrowTd.className = 'drill-arrow'; arrowTd.textContent = '›';
        tr.appendChild(arrowTd);
        tr.addEventListener('click', () => onRowClick(r));
      }
      tbody.appendChild(tr);
    });
    const old = tbl.querySelector('tbody');
    if (old) old.replaceWith(tbody); else tbl.appendChild(tbody);
  };

  cols.forEach(c => {
    const th = document.createElement('th');
    th.innerHTML = `${c.h}<span class="sort-ind">⇅</span>`;
    if (c.cls && c.cls.includes('num')) th.classList.add('num');
    th.title = `Sort by ${c.h}`;
    th.addEventListener('click', () => {
      const dir = localSort.key === c.k ? -localSort.dir : -1;
      localSort.key = c.k; localSort.dir = dir;
      rows.sort((a, b) => {
        const va = a[c.k], vb = b[c.k];
        if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
        return String(va || '').localeCompare(String(vb || '')) * dir;
      });
      trH.querySelectorAll('th').forEach(h => {
        h.classList.remove('sort-active');
        const s = h.querySelector('.sort-ind'); if (s) s.textContent = '⇅';
      });
      th.classList.add('sort-active');
      th.querySelector('.sort-ind').textContent = dir === -1 ? '▼' : '▲';
      rebuildBody();
      // Re-wire export button with current sort
      const expBtn = document.getElementById('modalExportBtn');
      if (expBtn && expBtn.style.display !== 'none') {
        const curTbl = document.getElementById(id);
        if (curTbl) expBtn.onclick = () => downloadCSV(tableToCSV(curTbl), safeFilename(id) + '.csv');
      }
    });
    trH.appendChild(th);
  });

  if (onRowClick) trH.insertAdjacentHTML('beforeend', '<th style="width:28px"></th>');
  thead.appendChild(trH); tbl.appendChild(thead);
  rebuildBody();
  wrap.appendChild(tbl);
  return wrap;
}

// ── LEVEL 1: National → by Distributor ───────────────────────────────────────
function drillLevel1(metric, moRows, apRows) {
  const wrap = document.createElement('div');

  if (['actualRev','missedRev','oosDays'].includes(metric)) {
    const byDist = {};
    moRows.forEach(r => {
      if (!byDist[r.distributor]) byDist[r.distributor] = { distributor:r.distributor, actual:0, missed:0, potential:0, oosSum:0, n:0, outlets:new Set(), skus:new Set() };
      byDist[r.distributor].actual   += r.actualRev;
      byDist[r.distributor].missed   += r.missedRev;
      byDist[r.distributor].potential+= r.potentialRev;
      byDist[r.distributor].oosSum   += r.oosDays;
      byDist[r.distributor].n++;
      byDist[r.distributor].outlets.add(r.outletCode || r.outlet);
      byDist[r.distributor].skus.add(r.product);
    });
    const list = Object.values(byDist).map(d => ({
      ...d, avgOos: d.oosSum/Math.max(1,d.n), outletCount:d.outlets.size, skuCount:d.skus.size
    })).sort((a,b) => metric==='oosDays' ? b.avgOos-a.avgOos : b[metric==='actualRev'?'actual':'missed']-a[metric==='actualRev'?'actual':'missed']);

    const cols = [
      {h:'Distributor', k:'distributor'},
      {h:'Outlets', k:'outletCount', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'SKUs', k:'skuCount', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'Actual Rev', k:'actual', cls:'num', fmt:fmtMoneyFull},
      {h:'Missed Rev', k:'missed', cls:'num bad', fmt:fmtMoneyFull},
      {h:'Potential Rev', k:'potential', cls:'num', fmt:fmtMoneyFull},
      {h:'Avg OOS Days', k:'avgOos', cls:'num', fmt:v=>fmtNum(v,1)},
    ];
    wrap.insertAdjacentHTML('beforeend', `<div class="drill-count">${list.length} distributors — click any row to drill into outlet level</div>`);
    wrap.appendChild(buildDrillTable('dL1', list, cols, row => {
      const distMo = moRows.filter(r=>r.distributor===row.distributor);
      const distAp = apRows.filter(a=>a.distributor===row.distributor);
      drillPush(row.distributor, () => drillLevel2(metric, row.distributor, distMo, distAp));
    }));

  } else {
    // SoS / Asset Health
    const byDist = {};
    apRows.forEach(a => {
      if (!byDist[a.distributor]) byDist[a.distributor] = { distributor:a.distributor, eSoSSum:0, fSoSSum:0, n:0, noComm:0, noImage:0, noDoor:0 };
      byDist[a.distributor].eSoSSum += a.emptySoS; byDist[a.distributor].fSoSSum += a.foreignSoS;
      byDist[a.distributor].n++;
      if (a.noCommunication) byDist[a.distributor].noComm++;
      if (a.noImage) byDist[a.distributor].noImage++;
      if (a.noDoor) byDist[a.distributor].noDoor++;
    });
    const revByDist = {}; moRows.forEach(r => { revByDist[r.distributor]=(revByDist[r.distributor]||0)+r.actualRev; });
    const list = Object.values(byDist).map(d => ({
      ...d, avgESoS:d.eSoSSum/Math.max(1,d.n), avgFSoS:d.fSoSSum/Math.max(1,d.n),
      activePct:((d.n-d.noComm-d.noImage-d.noDoor)/Math.max(1,d.n))*100,
      actual:revByDist[d.distributor]||0,
    })).sort((a,b) => metric==='emptySoS' ? b.avgESoS-a.avgESoS : metric==='foreignSoS' ? b.avgFSoS-a.avgFSoS : a.activePct-b.activePct);

    const cols = [
      {h:'Distributor', k:'distributor'},
      {h:'Total Assets', k:'n', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'Active %', k:'activePct', cls:'num', fmt:v=>fmtPct(v)},
      {h:'No Comm', k:'noComm', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'No Image', k:'noImage', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'No Door', k:'noDoor', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'Avg Empty SoS', k:'avgESoS', cls:'num', fmt:v=>fmtPct(v)},
      {h:'Avg Foreign SoS', k:'avgFSoS', cls:'num', fmt:v=>fmtPct(v)},
      {h:'Actual Rev', k:'actual', cls:'num', fmt:fmtMoneyFull},
    ];
    wrap.insertAdjacentHTML('beforeend', `<div class="drill-count">${list.length} distributors — click any row to drill into outlet level</div>`);
    wrap.appendChild(buildDrillTable('dL1', list, cols, row => {
      const distMo = moRows.filter(r=>r.distributor===row.distributor);
      const distAp = apRows.filter(a=>a.distributor===row.distributor);
      drillPush(row.distributor, () => drillLevel2(metric, row.distributor, distMo, distAp));
    }));
  }
  return wrap;
}

// ── LEVEL 2: Distributor → by Outlet ─────────────────────────────────────────
function drillLevel2(metric, distributor, moRows, apRows) {
  const wrap = document.createElement('div');

  if (['actualRev','missedRev','oosDays'].includes(metric)) {
    const byOutlet = {};
    moRows.forEach(r => {
      const k = r.outletCode||r.outlet;
      if (!byOutlet[k]) byOutlet[k] = { outlet:r.outlet, outletCode:r.outletCode, route:r.route, actual:0, missed:0, potential:0, oosSum:0, n:0, skus:new Set() };
      byOutlet[k].actual   += r.actualRev; byOutlet[k].missed  += r.missedRev;
      byOutlet[k].potential+= r.potentialRev; byOutlet[k].oosSum += r.oosDays; byOutlet[k].n++;
      byOutlet[k].skus.add(r.product);
    });
    const list = Object.values(byOutlet).map(o => ({
      ...o, avgOos:o.oosSum/Math.max(1,o.n), skuCount:o.skus.size
    })).sort((a,b) => metric==='oosDays' ? b.avgOos-a.avgOos : b[metric==='actualRev'?'actual':'missed']-a[metric==='actualRev'?'actual':'missed']);

    const cols = [
      {h:'Outlet', k:'outlet'}, {h:'Code', k:'outletCode', cls:'tight'}, {h:'Route', k:'route', cls:'tight'},
      {h:'SKUs', k:'skuCount', cls:'num', fmt:v=>fmtNum(v,0)},
      {h:'Actual Rev', k:'actual', cls:'num', fmt:fmtMoneyFull},
      {h:'Missed Rev', k:'missed', cls:'num bad', fmt:fmtMoneyFull},
      {h:'Avg OOS Days', k:'avgOos', cls:'num', fmt:v=>fmtNum(v,1)},
    ];
    wrap.insertAdjacentHTML('beforeend', `<div class="drill-count">${list.length} outlets — click any row to see SKU detail</div>`);
    wrap.appendChild(buildDrillTable('dL2', list, cols, row => {
      const k = row.outletCode||row.outlet;
      const outletMo = moRows.filter(r=>(r.outletCode||r.outlet)===k);
      const outletAp = apRows.filter(a=>(a.outletCode||a.outlet)===k);
      drillPush(row.outlet, () => drillLevel3Revenue(metric, outletMo, outletAp));
    }));

  } else {
    // SoS / Asset Health: outlet = asset row
    const sorted = apRows.slice().sort((a,b) =>
      metric==='emptySoS' ? b.emptySoS-a.emptySoS :
      metric==='foreignSoS' ? b.foreignSoS-a.foreignSoS :
      (a.noCommunication?0:a.noImage?1:a.noDoor?2:3)-(b.noCommunication?0:b.noImage?1:b.noDoor?2:3));
    const revByOutlet = {}; moRows.forEach(r => { const k=r.outletCode||r.outlet; revByOutlet[k]=(revByOutlet[k]||0)+r.actualRev; });

    const cols = [
      {h:'Outlet', k:'outlet'}, {h:'Code', k:'outletCode', cls:'tight'}, {h:'Asset', k:'asset', cls:'tight'},
      {h:'Camera', k:'cabinetType'}, {h:'Status', k:'status', fmt:v=>statusPill(v)},
      {h:'Last Image', k:'lastImg', cls:'tight', fmt:fmtDate},
      {h:'Empty SoS', k:'emptySoS', cls:'num', fmt:v=>fmtPct(v)},
      {h:'Foreign SoS', k:'foreignSoS', cls:'num', fmt:v=>fmtPct(v)},
    ];
    wrap.insertAdjacentHTML('beforeend', `<div class="drill-count">${sorted.length} outlets — click any row for full asset detail</div>`);
    wrap.appendChild(buildDrillTable('dL2', sorted, cols, row => {
      drillPush(row.outlet, () => drillLevel3Asset(row, moRows));
    }));
  }
  return wrap;
}

// ── LEVEL 3a: Outlet → SKU Detail (revenue / OOS) ────────────────────────────
function drillLevel3Revenue(metric, moRows, apRows) {
  const wrap = document.createElement('div');
  const sorted = moRows.slice().sort((a,b) =>
    metric==='oosDays' ? b.oosDays-a.oosDays : metric==='actualRev' ? b.actualRev-a.actualRev : b.missedRev-a.missedRev);

  // Asset info if available
  const asset = apRows[0];
  if (asset) {
    wrap.insertAdjacentHTML('beforeend', `
      <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:12px;padding:10px 12px;background:var(--surface-2);border:1px solid var(--border);border-radius:7px;font-size:12px">
        <div><span class="muted">Asset</span> <b>${asset.asset}</b></div>
        <div><span class="muted">Camera</span> <b>${asset.cabinetType||'—'}</b></div>
        <div><span class="muted">Status</span> ${statusPill(asset.status)}</div>
        <div><span class="muted">Last Image</span> <b>${fmtDate(asset.lastImg)}</b></div>
        <div><span class="muted">Empty SoS</span> <b>${fmtPct(asset.emptySoS)}</b></div>
        <div><span class="muted">Foreign SoS</span> <b>${fmtPct(asset.foreignSoS)}</b></div>
      </div>`);
  }

  const cols = [
    {h:'Product', k:'product'}, {h:'Brand', k:'brand', cls:'tight'},
    {h:'Unit Price', k:'unitPrice', cls:'num', fmt:fmtMoneyFull},
    {h:'Units Sold', k:'unitsSold', cls:'num', fmt:v=>fmtNum(v,0)},
    {h:'Avg Daily Units', k:'avgDailyUnits', cls:'num', fmt:v=>fmtNum(v,2)},
    {h:'OOS Days', k:'oosDays', cls:'num', fmt:v=>fmtNum(v,1)},
    {h:'Actual Rev', k:'actualRev', cls:'num', fmt:fmtMoneyFull},
    {h:'Missed Rev', k:'missedRev', cls:'num bad', fmt:fmtMoneyFull},
    {h:'Potential Rev', k:'potentialRev', cls:'num', fmt:fmtMoneyFull},
  ];
  wrap.insertAdjacentHTML('beforeend', `<div class="drill-count">${sorted.length} SKU rows</div>`);
  wrap.appendChild(buildDrillTable('dL3', sorted, cols, null)); // no further drill
  return wrap;
}

// ── LEVEL 3b: Outlet → Asset Detail (SoS / health) ───────────────────────────
function drillLevel3Asset(asset, moRows) {
  const wrap = document.createElement('div');
  // Full asset detail card
  let dailyHtml = '';
  if (asset.dailyImg && asset.dailyImg.length) {
    dailyHtml = `<div class="muted small" style="margin:12px 0 4px">Rolling 7-day image activity:</div><div class="img-timeline">` +
      asset.dailyImg.map(d => {
        const cls = d.images===0?'none':d.images<2?'low':d.images<5?'ok':'high';
        return `<div class="day ${cls}" title="${d.date}: ${d.images} images / ${d.doors} doors">${d.images}</div>`;
      }).join('') + '</div>';
  }
  // Revenue from MO rows for this outlet
  const outletMo = moRows.filter(r=>(r.outletCode||r.outlet)===(asset.outletCode||asset.outlet));
  const totalActual = outletMo.reduce((s,r)=>s+r.actualRev,0);
  const totalMissed = outletMo.reduce((s,r)=>s+r.missedRev,0);
  wrap.innerHTML = `
    <div class="grid-3" style="gap:14px;margin-bottom:14px">
      <div><div class="muted small">Outlet</div><div><b>${asset.outlet}</b> · ${asset.outletCode||'—'}</div></div>
      <div><div class="muted small">Status</div>${statusPill(asset.status)}</div>
      <div><div class="muted small">Camera Type</div>${asset.cabinetType||'—'}</div>
      <div><div class="muted small">Asset Serial</div>${asset.asset}</div>
      <div><div class="muted small">Last Image</div>${fmtDate(asset.lastImg)}</div>
      <div><div class="muted small">Last Ping</div>${fmtDate(asset.devicePing)}</div>
      <div><div class="muted small">Empty SoS</div><b>${fmtPct(asset.emptySoS)}</b></div>
      <div><div class="muted small">Foreign SoS</div><b>${fmtPct(asset.foreignSoS)}</b></div>
      <div><div class="muted small">Diagnosis</div>${asset.diagnosis||'—'}</div>
      ${totalActual ? `<div><div class="muted small">Actual Revenue</div><b>${fmtMoneyFull(totalActual)}</b></div>` : ''}
      ${totalMissed ? `<div><div class="muted small">Missed Revenue</div><b style="color:var(--red)">${fmtMoneyFull(totalMissed)}</b></div>` : ''}
    </div>${dailyHtml}`;
  if (outletMo.length) {
    wrap.insertAdjacentHTML('beforeend', `<div class="muted small" style="margin:12px 0 6px">SKU detail (${outletMo.length} rows):</div>`);
    wrap.appendChild(buildDrillTable('dL3a', outletMo.sort((a,b)=>b.missedRev-a.missedRev), [
      {h:'Product',k:'product'},{h:'OOS Days',k:'oosDays',cls:'num',fmt:v=>fmtNum(v,1)},
      {h:'Actual Rev',k:'actualRev',cls:'num',fmt:fmtMoneyFull},
      {h:'Missed Rev',k:'missedRev',cls:'num bad',fmt:fmtMoneyFull},
    ], null));
  }
  return wrap;
}

/* ---------- 24. Wire-up ---------- */
document.addEventListener('DOMContentLoaded', () => {
  setupDropzones();
  document.querySelectorAll('.view-toggle button').forEach(b => {
    b.addEventListener('click', () => viewToggle(b.dataset.view));
  });
  document.getElementById('clearFiltersBtn').addEventListener('click', clearFilters);
  document.getElementById('exitBtn').addEventListener('click', () => {
    if (confirm('Return to upload screen? Loaded data will be cleared.')) {
      state.missedOpp = []; state.assetPerf = [];
      uploadedFiles.mo.length = 0; uploadedFiles.ap.length = 0;
      document.getElementById('moFiles').innerHTML = '';
      document.getElementById('apFiles').innerHTML = '';
      document.getElementById('landing').style.display = 'flex';
      document.getElementById('mainApp').style.display = 'none';
    }
  });
  document.getElementById('modalClose').addEventListener('click', closeDrill);
  document.getElementById('modal').addEventListener('click', e => { if (e.target.id === 'modal') closeDrill(); });
  // MOV inputs trigger re-render
  ['movInput','routeMovInput'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => {
      if (state.view === 'distributor') renderDistributor();
    });
  });
  // ── Asset health Grid / Map toggle ───────────────────────────────────────
  document.querySelectorAll('[data-asset-view]').forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.dataset.assetView;
      document.querySelectorAll('[data-asset-view]').forEach(b => b.classList.toggle('active', b === btn));
      const gridWrap = document.getElementById('assetGridWrap');
      const mapWrap  = document.getElementById('assetMapWrap');
      if (view === 'map') {
        if (gridWrap) gridWrap.style.display = 'none';
        if (mapWrap)  mapWrap.style.display  = 'block';
        renderAssetMap(state.currentAssets || []);
        setTimeout(() => { if (state.assetLeafletMap) state.assetLeafletMap.invalidateSize(); }, 120);
      } else {
        if (gridWrap) gridWrap.style.display = 'block';
        if (mapWrap)  mapWrap.style.display  = 'none';
      }
    });
  });

  // ── Outlet priority Table / Map toggle ───────────────────────────────────
  document.querySelectorAll('[data-outlet-view]').forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.dataset.outletView;
      document.querySelectorAll('[data-outlet-view]').forEach(b => b.classList.toggle('active', b === btn));
      const tblWrap = document.getElementById('outletPriorityWrap');
      const mapWrap = document.getElementById('outletMapWrap');
      if (view === 'map') {
        if (tblWrap) tblWrap.style.display = 'none';
        if (mapWrap) mapWrap.style.display = 'block';
        renderOutletMap(state.currentOutletList || []);
        setTimeout(() => { if (state.outletLeafletMap) state.outletLeafletMap.invalidateSize(); }, 120);
      } else {
        if (tblWrap) tblWrap.style.display = 'block';
        if (mapWrap) mapWrap.style.display = 'none';
      }
    });
  });

  // (preset-pills removed — period presets now live in the Period dropdown)
});
