// ---------- constants ----------
const COLORS = {
  JLD: '#C9852E',
  ZPM: '#3F9FA8',
  budget: '#8B93A0',
  pos: '#6FBF8B',
  neg: '#E0645A',
};
const BRAND_NAME = { JLD: "Jul's", ZPM: 'Zephyr' };
const DAY_MS = 24*60*60*1000;

// Country -> local currency, used to know which FX rate applies to each outlet.
// Add entries here if outlets are added in new countries.
const COUNTRY_CURRENCY = {
  'UK': 'GBP', 'United Kingdom': 'GBP',
  'Monaco': 'EUR', 'France': 'EUR', 'Italy': 'EUR', 'Spain': 'EUR', 'Germany': 'EUR',
  'Thailand': 'THB',
};
const REPORTING_CURRENCY = 'THB';

let RAW = null; // { actual: [...], budget: [...], outlets: [...], firstActualDate, lastActualDate }
let currentOutletFilter = 'all'; // 'all' | 'JLD' | 'ZPM'
let currentPeriodType = 'DAY'; // 'DAY' | 'MTD' | 'QTD' | 'YTD'
let currentAsOfDate = null; // 'YYYY-MM-DD', defaults to lastActualDate on load
const charts = {};

// ---------- formatting helpers ----------
function fmtMoney(n){
  const sign = n < 0 ? '-' : '';
  const abs = Math.abs(n);
  if (abs >= 1000000) return sign + '฿' + (abs/1000000).toFixed(2) + 'M';
  if (abs >= 1000) return sign + '฿' + (abs/1000).toFixed(1) + 'k';
  return sign + '฿' + abs.toFixed(0);
}
function fmtNum(n){ return Math.round(n).toLocaleString(); }
function fmtPct(n){ const s = n > 0 ? '+' : ''; return s + (n*100).toFixed(1) + '%'; }
function addDays(dateStr, days){
  const d = new Date(dateStr+'T00:00:00Z');
  d.setUTCDate(d.getUTCDate()+days);
  return d.toISOString().slice(0,10);
}

// ---------- excel date helpers ----------
function toDateStr(v){
  if (v instanceof Date) {
    // Shift by 12h before reading the calendar date. This makes the result immune to
    // any near-midnight drift in how the source file's date serials were encoded
    // (e.g. a date meant to be a clean midnight sometimes lands a few hours either
    // side of 00:00:00 UTC depending on what tool/timezone wrote the spreadsheet).
    // A clean midnight UTC date is unaffected; a drifted one gets pulled back onto
    // the correct calendar day.
    const shifted = new Date(v.getTime() + 12*60*60*1000);
    const y = shifted.getUTCFullYear();
    const m = String(shifted.getUTCMonth()+1).padStart(2,'0');
    const d = String(shifted.getUTCDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  return String(v);
}
function toMonthStr(dateStr){ return dateStr.slice(0,7); }

// ---------- parse raw sheets into RAW (no aggregation yet, but revenue is converted to THB) ----------
function buildRaw(sheets){
  const revRaw = sheets['Revenue_Daily'] || [];
  const budRaw = sheets['Budget'] || [];
  const outletRaw = sheets['Outlet'] || [];
  const fxRaw = sheets['FX'] || [];

  // FX rates: Currency -> rate to THB (if multiple years given, last one wins)
  const fxRates = {};
  fxRaw.forEach(r => { if (r['Currency']) fxRates[r['Currency']] = Number(r['FX to THB']) || 1; });
  fxRates[REPORTING_CURRENCY] = 1; // THB to THB is always 1

  const outletMap = {};       // Outlet code -> Brand
  const outletCurrency = {};  // Outlet code -> currency
  const outletMeta = [];
  const unmappedCountries = new Set();

  outletRaw.forEach(r => {
    outletMap[r['Outlet']] = r['Brand'];
    const country = r['Country'];
    let currency = COUNTRY_CURRENCY[country];
    if (!currency) { unmappedCountries.add(country); currency = REPORTING_CURRENCY; }
    outletCurrency[r['Outlet']] = currency;
    outletMeta.push({
      Outlet: r['Outlet'], Country: country, Brand: r['Brand'], Currency: currency,
      'Opening Date': r['Opening Date'] ? toDateStr(r['Opening Date']) : null
    });
  });

  function fxRateFor(outletCode){
    const currency = outletCurrency[outletCode] || REPORTING_CURRENCY;
    return fxRates[currency] !== undefined ? fxRates[currency] : 1;
  }

  const actual = [];
  revRaw.forEach(r => {
    const revVal = r['Revenue'];
    if (revVal === undefined || revVal === null || revVal === '') return;
    const dateStr = toDateStr(r['Date']);
    const rate = fxRateFor(r['Outlet']);
    actual.push({
      Date: dateStr, Month: toMonthStr(dateStr), Outlet: r['Outlet'],
      Brand: outletMap[r['Outlet']] || r['Outlet'], MealPeriod: r['Meal Period'],
      RevenueLocal: Number(revVal) || 0,
      Revenue: (Number(revVal) || 0) * rate,
      Covers: Number(r['Covers']) || 0,
    });
  });

  if (actual.length === 0) {
    throw new Error('No recorded (non-blank) Revenue rows found in Revenue_Daily — check the file structure.');
  }

  const budget = [];
  budRaw.forEach(r => {
    const dateStr = toDateStr(r['Date']);
    const rate = fxRateFor(r['Outlet']);
    budget.push({
      Date: dateStr, Month: toMonthStr(dateStr), Outlet: r['Outlet'],
      Brand: outletMap[r['Outlet']] || r['Outlet'], MealPeriod: r['Meal Period'],
      BudgetRevenueLocal: Number(r['Revenue Budget']) || 0,
      BudgetRevenue: (Number(r['Revenue Budget']) || 0) * rate,
      BudgetCovers: Number(r['Covers Budget']) || 0,
    });
  });

  const allDates = actual.map(r => r.Date).sort();
  if (unmappedCountries.size) {
    console.warn('Unmapped country -> currency (defaulted to THB, 1:1):', [...unmappedCountries]);
  }
  return {
    actual, budget, outlets: outletMeta,
    fxRates,
    firstActualDate: allDates[0],
    lastActualDate: allDates[allDates.length-1],
    unmappedCountries: [...unmappedCountries],
  };
}

// ---------- period range resolution (MTD / QTD / YTD, relative to the "as of" date) ----------
function resolveDateRange(){
  const asOf = currentAsOfDate || RAW.lastActualDate;
  if (currentPeriodType === 'DAY') {
    return { from: asOf, to: asOf };
  }
  const d = new Date(asOf+'T00:00:00Z');
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth(); // 0-indexed
  let fromDate;
  if (currentPeriodType === 'MTD') {
    fromDate = new Date(Date.UTC(y, m, 1));
  } else if (currentPeriodType === 'QTD') {
    const qStartMonth = Math.floor(m/3)*3;
    fromDate = new Date(Date.UTC(y, qStartMonth, 1));
  } else { // YTD
    fromDate = new Date(Date.UTC(y, 0, 1));
  }
  // don't go earlier than the data actually starts
  const fromStr = fromDate.toISOString().slice(0,10);
  const clampedFrom = fromStr < RAW.firstActualDate ? RAW.firstActualDate : fromStr;
  return { from: clampedFrom, to: asOf };
}

function filteredOutlets(){
  return currentOutletFilter === 'all' ? ['JLD','ZPM'] : [currentOutletFilter];
}

// ---------- derived aggregates, computed fresh from RAW on every render ----------
function getFilteredActual(){
  const {from, to} = resolveDateRange();
  const outlets = filteredOutlets();
  return RAW.actual.filter(r => outlets.includes(r.Outlet) && r.Date >= from && r.Date <= to);
}
function getFilteredBudget(){
  const {from, to} = resolveDateRange();
  const outlets = filteredOutlets();
  return RAW.budget.filter(r => outlets.includes(r.Outlet) && r.Date >= from && r.Date <= to);
}

function computeKPIs(){
  const act = getFilteredActual();
  const bud = getFilteredBudget();
  const revenue = act.reduce((s,r)=>s+r.Revenue,0);
  const covers = act.reduce((s,r)=>s+r.Covers,0);
  const budgetRevenue = bud.reduce((s,r)=>s+r.BudgetRevenue,0);
  const budgetCovers = bud.reduce((s,r)=>s+r.BudgetCovers,0);
  const avgCheck = covers ? revenue/covers : 0;
  const budgetAvgCheck = budgetCovers ? budgetRevenue/budgetCovers : 0;
  const revVsBudget = budgetRevenue ? (revenue-budgetRevenue)/budgetRevenue : 0;
  const coversVsBudget = budgetCovers ? (covers-budgetCovers)/budgetCovers : 0;
  return {revenue, covers, avgCheck, budgetRevenue, budgetCovers, budgetAvgCheck, revVsBudget, coversVsBudget};
}

function computeMonthlyByOutlet(){
  const act = getFilteredActual();
  const bud = getFilteredBudget();
  const outlets = filteredOutlets();
  const months = [...new Set([...act.map(r=>r.Month), ...bud.map(r=>r.Month)])].sort();
  const perOutlet = {};
  outlets.forEach(o => {
    perOutlet[o] = months.map(m => act.filter(r=>r.Month===m && r.Outlet===o).reduce((s,r)=>s+r.Revenue,0));
  });
  const budgetByMonth = months.map(m => bud.filter(r=>r.Month===m).reduce((s,r)=>s+r.BudgetRevenue,0));
  return { months, outlets, perOutlet, budgetByMonth };
}

function computeDaily(){
  const act = getFilteredActual();
  const outlets = filteredOutlets();
  const dates = [...new Set(act.map(r=>r.Date))].sort();
  return { dates, outlets, rows: act };
}

function computeMealSplit(){
  const act = getFilteredActual();
  const outlets = filteredOutlets();
  return outlets.map(o => ({
    Outlet: o,
    Lunch: act.filter(r=>r.Outlet===o && r.MealPeriod==='Lunch').reduce((s,r)=>s+r.Revenue,0),
    Dinner: act.filter(r=>r.Outlet===o && r.MealPeriod==='Dinner').reduce((s,r)=>s+r.Revenue,0),
  }));
}

function computeOutletSplit(){
  const act = getFilteredActual();
  const outlets = filteredOutlets();
  return outlets.map(o => ({
    Outlet: o,
    Revenue: act.filter(r=>r.Outlet===o).reduce((s,r)=>s+r.Revenue,0),
  })).filter(r => r.Revenue > 0 || outlets.length <= 1);
}

// ---------- UI: filter pills ----------
function buildOutletPills(){
  const el = document.getElementById('filterPills');
  const options = [
    {key:'all', label:'All Outlets', color:null},
    {key:'JLD', label:"Jul's — London", color:COLORS.JLD},
    {key:'ZPM', label:'Zephyr — Monaco', color:COLORS.ZPM},
  ];
  el.innerHTML = options.map(o => `
    <div class="pill ${o.key===currentOutletFilter?'active':''}" data-key="${o.key}">
      ${o.color ? `<span class="dot" style="background:${o.color}"></span>` : ''}
      ${o.label}
    </div>
  `).join('');
  el.querySelectorAll('.pill').forEach(p => {
    p.addEventListener('click', () => { currentOutletFilter = p.getAttribute('data-key'); renderAll(); });
  });
}

function buildPeriodPills(){
  const el = document.getElementById('periodPills');
  const options = [
    {key:'DAY', label:'Daily'},
    {key:'MTD', label:'Month to Date'},
    {key:'QTD', label:'Quarter to Date'},
    {key:'YTD', label:'Year to Date'},
  ];
  el.innerHTML = options.map(o => `
    <div class="pill ${o.key===currentPeriodType?'active':''}" data-key="${o.key}">${o.label}</div>
  `).join('');
  el.querySelectorAll('.pill').forEach(p => {
    p.addEventListener('click', () => { currentPeriodType = p.getAttribute('data-key'); renderAll(); });
  });
}

document.getElementById('asOfDate').addEventListener('change', (e) => {
  currentAsOfDate = e.target.value || RAW.lastActualDate;
  renderAll();
});

// ---------- KPI rendering ----------
function renderKPIs(){
  const k = computeKPIs();
  const el = document.getElementById('kpiRow');
  el.innerHTML = `
    <div class="kpi">
      <div class="kpi-label">Total Revenue</div>
      <div class="kpi-value">${fmtMoney(k.revenue)}</div>
      <div class="kpi-sub ${k.revVsBudget>=0?'pos':'neg'}">${fmtPct(k.revVsBudget)} vs budget (${fmtMoney(k.budgetRevenue)})</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Total Covers</div>
      <div class="kpi-value">${fmtNum(k.covers)}</div>
      <div class="kpi-sub ${k.coversVsBudget>=0?'pos':'neg'}">${fmtPct(k.coversVsBudget)} vs budget (${fmtNum(k.budgetCovers)})</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Average Check</div>
      <div class="kpi-value">฿${k.avgCheck.toFixed(0)}</div>
      <div class="kpi-sub">Budgeted: ฿${k.budgetAvgCheck.toFixed(0)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Revenue per Cover Trend</div>
      <div class="kpi-value">${k.avgCheck > k.budgetAvgCheck ? '↑' : '↓'} ฿${Math.abs(k.avgCheck-k.budgetAvgCheck).toFixed(0)}</div>
      <div class="kpi-sub">${k.avgCheck >= k.budgetAvgCheck ? 'Above' : 'Below'} budgeted spend per guest</div>
    </div>
  `;
}

function destroyChart(id){ if (charts[id]) { charts[id].destroy(); charts[id] = null; } }

function chartBaseOptions(hideLegendLabelsSmall, stacked){
  return {
    responsive:true,
    maintainAspectRatio:false,
    interaction:{mode:'index', intersect:false},
    plugins:{
      legend:{ display:true, position:'top', align:'end',
        labels:{color:'#8B93A0', font:{size:11, family:'Inter'}, boxWidth:10, boxHeight:10, usePointStyle:true} },
      tooltip:{
        backgroundColor:'#222A36', titleColor:'#F1EDE4', bodyColor:'#F1EDE4',
        borderColor:'#2C3542', borderWidth:1, padding:10,
        callbacks:{ label: (c) => `${c.dataset.label}: ${fmtMoney(c.raw)}` }
      }
    },
    scales:{
      x:{ grid:{color:'#232B36', display:false}, stacked: !!stacked,
        ticks:{color:'#8B93A0', font:{size:10, family:'Inter'}, maxRotation:0, autoSkip:true, maxTicksLimit: hideLegendLabelsSmall ? 10 : undefined} },
      y:{ grid:{color:'#232B36'}, stacked: !!stacked,
        ticks:{color:'#8B93A0', font:{size:10, family:'Inter'}, callback:(v)=>fmtMoney(v)} }
    }
  };
}

function renderMonthlyChart(){
  const { months, outlets, perOutlet, budgetByMonth } = computeMonthlyByOutlet();
  const monthLabels = months.map(m => new Date(m+'-01T00:00:00Z').toLocaleDateString('en-US',{month:'short', year:'2-digit', timeZone:'UTC'}));

  const barDatasets = outlets.map(o => ({
    type:'bar',
    label: BRAND_NAME[o]||o,
    data: perOutlet[o],
    backgroundColor: COLORS[o]||'#C9A24B',
    borderRadius:6,
    maxBarThickness:42,
    stack:'actual',
  }));

  const budgetDataset = {
    type:'line', label:'Budget', data: budgetByMonth, borderColor:COLORS.budget, borderDash:[5,4],
    borderWidth:2, pointRadius:3, pointBackgroundColor:COLORS.budget, tension:0.25, fill:false,
  };

  destroyChart('monthly');
  charts.monthly = new Chart(document.getElementById('monthlyChart'), {
    data: { labels: monthLabels, datasets: [...barDatasets, budgetDataset] },
    options: chartBaseOptions(true, true)
  });
}

function renderOutletDonut(){
  const rows = computeOutletSplit();
  destroyChart('donut');
  charts.donut = new Chart(document.getElementById('outletDonut'), {
    type:'doughnut',
    data:{ labels: rows.map(r=>BRAND_NAME[r.Outlet]||r.Outlet),
      datasets:[{ data: rows.map(r=>r.Revenue), backgroundColor: rows.map(r=>COLORS[r.Outlet]||'#8B93A0'), borderColor:'#1B212B', borderWidth:3 }] },
    options:{ responsive:true, maintainAspectRatio:false, cutout:'68%',
      plugins:{ legend:{display:false}, tooltip:{callbacks:{label:(c)=> (BRAND_NAME[rows[c.dataIndex].Outlet]||rows[c.dataIndex].Outlet) + ': ' + fmtMoney(c.raw) }} } }
  });
  document.getElementById('outletLegend').innerHTML = rows.map(r => `
    <div class="legend-item"><span class="sw" style="background:${COLORS[r.Outlet]||'#8B93A0'}"></span>${BRAND_NAME[r.Outlet]||r.Outlet} — ${fmtMoney(r.Revenue)}</div>
  `).join('');
}

function renderDailyChart(){
  const { dates, outlets, rows } = computeDaily();
  destroyChart('daily');
  const datasets = outlets.map(o => ({
    label: BRAND_NAME[o]||o,
    data: dates.map(d => {
      const matches = rows.filter(r=>r.Date===d && r.Outlet===o);
      return matches.length ? matches.reduce((s,r)=>s+r.Revenue,0) : null;
    }),
    borderColor: COLORS[o]||'#8B93A0', backgroundColor: COLORS[o]||'#8B93A0',
    borderWidth:2, pointRadius: dates.length <= 31 ? 2 : 0, tension:0.3, spanGaps:true,
  }));
  charts.daily = new Chart(document.getElementById('dailyChart'), {
    type:'line',
    data:{ labels: dates.map(d => new Date(d+'T00:00:00Z').toLocaleDateString('en-US',{month:'short', day:'numeric', timeZone:'UTC'})), datasets },
    options: chartBaseOptions(false)
  });
}

function renderMealChart(){
  const rows = computeMealSplit();
  destroyChart('meal');
  charts.meal = new Chart(document.getElementById('mealChart'), {
    type:'bar',
    data:{ labels: rows.map(r=>BRAND_NAME[r.Outlet]||r.Outlet), datasets:[
      { label:'Lunch', data: rows.map(r=>r.Lunch), backgroundColor:'#C9A24B', borderRadius:6, maxBarThickness:56 },
      { label:'Dinner', data: rows.map(r=>r.Dinner), backgroundColor:'#4A5568', borderRadius:6, maxBarThickness:56 },
    ]},
    options: chartBaseOptions(false)
  });
}

function renderFooter(){
  const {from, to} = resolveDateRange();
  const periodLabels = {DAY:'Daily', MTD:'Month to Date', QTD:'Quarter to Date', YTD:'Year to Date'};
  const rateLines = Object.entries(RAW.fxRates || {})
    .filter(([cur]) => cur !== REPORTING_CURRENCY)
    .map(([cur, rate]) => `${cur}→THB ${rate}`).join(', ');
  const warning = (RAW.unmappedCountries && RAW.unmappedCountries.length)
    ? ` <span style="color:var(--neg)">Note: currency for ${RAW.unmappedCountries.join(', ')} not recognized — defaulted to THB 1:1, please verify.</span>`
    : '';
  document.getElementById('footerNote').innerHTML = `
    <strong>About this view —</strong>
    Showing ${from} through ${to} (${periodLabels[currentPeriodType] || currentPeriodType}), all figures consolidated into <strong>Thai Baht (THB)</strong> using flat annual FX rates (${rateLines}). Jul's is recorded in GBP and Zephyr in EUR at source; conversion happens automatically on load or upload.${warning}
    Full recorded actuals span ${RAW.firstActualDate} through ${RAW.lastActualDate}; budget figures are compared over the same selected window.
    To refresh with new data, upload a new export using the button above — everything recalculates instantly in your browser and nothing is sent anywhere.
  `;
}

function renderAll(){
  buildOutletPills();
  buildPeriodPills();
  document.getElementById('asOfDate').value = currentAsOfDate || RAW.lastActualDate;
  renderKPIs();
  renderMonthlyChart();
  renderOutletDonut();
  renderDailyChart();
  renderMealChart();
  document.getElementById('periodRange').textContent = `${RAW.firstActualDate} → ${RAW.lastActualDate}`;
  document.getElementById('statusDate').textContent = RAW.lastActualDate;
  renderFooter();
}

const STORAGE_KEY = 'fb_dashboard_raw_data';
const hasSharedStorage = (typeof window !== 'undefined' && !!window.storage);

async function loadSharedData(){
  if (!hasSharedStorage) return null;
  try {
    const result = await window.storage.get(STORAGE_KEY, true);
    if (result && result.value) return JSON.parse(result.value);
    return null;
  } catch (err) {
    return null; // key doesn't exist yet, or storage unavailable
  }
}

async function saveSharedData(raw, meta){
  if (!hasSharedStorage) return false;
  try {
    const payload = JSON.stringify({ ...raw, _uploadedAt: new Date().toISOString(), _uploadedFileName: meta && meta.fileName });
    const result = await window.storage.set(STORAGE_KEY, payload, true);
    return !!result;
  } catch (err) {
    console.error('Shared storage save failed:', err);
    return false;
  }
}
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const fileStatus = document.getElementById('fileStatus');

uploadBtn.addEventListener('click', () => {
  fileInput.value = ''; // reset so selecting the same filename again still fires 'change'
  fileInput.click();
});

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  fileStatus.textContent = `Reading ${file.name}...`;
  fileStatus.className = '';
  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      const wb = XLSX.read(evt.target.result, {type:'array', cellDates:true});
      const sheets = {};
      wb.SheetNames.forEach(name => { sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], {defval:null}); });
      RAW = buildRaw(sheets);
      currentOutletFilter = 'all';
      currentPeriodType = 'DAY';
      currentAsOfDate = RAW.lastActualDate;
      document.getElementById('asOfDate').min = RAW.firstActualDate;
      document.getElementById('asOfDate').max = RAW.lastActualDate;
      renderAll();

      if (hasSharedStorage) {
        fileStatus.textContent = `Saving ${file.name} for everyone with this link...`;
        const saved = await saveSharedData(RAW, { fileName: file.name });
        if (saved) {
          fileStatus.textContent = `✓ Updated from ${file.name} — showing data through ${RAW.lastActualDate}. Everyone opening this link will now see this update.`;
        } else {
          fileStatus.textContent = `✓ Updated locally through ${RAW.lastActualDate}, but couldn't save for other viewers — try again.`;
        }
      } else {
        fileStatus.textContent = `✓ Updated from ${file.name} — showing data through ${RAW.lastActualDate} (this view only; open via a published Claude link for shared updates).`;
      }
      fileStatus.className = 'ok';
    } catch (err) {
      console.error(err);
      fileStatus.textContent = `Could not read that file: ${err.message}`;
      fileStatus.className = 'err';
    }
  };
  reader.onerror = () => { fileStatus.textContent = 'Error reading file.'; fileStatus.className = 'err'; };
  reader.readAsArrayBuffer(file);
});

// ---------- initial load: shared storage first (if available), else embedded default ----------
(async function init(){
  const embedded = JSON.parse(document.getElementById('dashboard-data-json').textContent);
  let initial = { actual: embedded.actual, budget: embedded.budget, outlets: embedded.outlets,
    fxRates: embedded.fxRates, firstActualDate: embedded.firstActualDate, lastActualDate: embedded.lastActualDate,
    unmappedCountries: [] };
  let sharedNote = '';

  const shared = await loadSharedData();
  if (shared && shared.actual && shared.actual.length) {
    initial = shared;
    sharedNote = shared._uploadedFileName
      ? ` (last updated from ${shared._uploadedFileName} on ${new Date(shared._uploadedAt).toLocaleString()})`
      : '';
  }

  RAW = initial;
  currentAsOfDate = RAW.lastActualDate;
  document.getElementById('asOfDate').min = RAW.firstActualDate;
  document.getElementById('asOfDate').max = RAW.lastActualDate;
  renderAll();

  if (hasSharedStorage) {
    fileStatus.textContent = shared
      ? `Showing shared data${sharedNote}.`
      : `Showing starting data. Upload a file to set the shared data everyone sees on this link.`;
    fileStatus.className = '';
  }
})();
