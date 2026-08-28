// app.js
// Parses the first sheet of an Excel file and renders charts using Chart.js

const fileInput = document.getElementById('file-input');
const dataTable = document.getElementById('data-table');
const summaryDiv = document.getElementById('summary');

let charts = [];

fileInput.addEventListener('change', (ev) => {
  const file = ev.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    handleData(json);
  };
  reader.readAsArrayBuffer(file);
});

function normalizeKey(k) {
  return String(k || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function keyMap(row) {
  // map headers to canonical keys (lowercase names matching expected columns)
  const mapped = {};
  for (const key in row) {
    const nk = normalizeKey(key);
    mapped[nk] = row[key];
  }
  return mapped;
}

function handleData(rows) {
  if (!rows || rows.length === 0) {
    summaryDiv.textContent = 'No rows found in the first sheet.';
    return;
  }

  const mappedRows = rows.map(r => keyMap(r));

  // create canonical column accessors
  // columns expected (case-insensitive): SR, Site, Sources, Shift, Year, Date, Incident Category, Criticality, Type of Incident, Event Title, GEHSMS Standard, Detailed Observation, What is the Action?, Priority, Owner, Status, New Timeline, Department, Area, Contractor Name, Responsible

  // For each row, extract important fields with fallbacks
  const data = mappedRows.map(r => ({
    sr: r['sr'] || r['s r'] || r['#'] || '',
    site: r['site'] || '',
    sources: r['sources'] || '',
    shift: r['shift'] || '',
    year: r['year'] || (r['date'] ? parseYear(r['date']) : ''),
    date: r['date'] || '',
    incident_category: r['incident category'] || r['incident_category'] || r['category'] || '',
    criticality: r['criticality'] || '',
    type_of_incident: r['type of incident'] || r['type'] || '',
    event_title: r['event title'] || '',
    gehsms_standard: r['gehsms standard'] || '',
    detailed_observation: r['detailed observation'] || '',
    what_is_the_action: r['what is the action?'.toLowerCase()] || r['what is the action'] || '',
    priority: r['priority'] || '',
    owner: r['owner'] || '',
    status: r['status'] || '',
    new_timeline: r['new timeline'] || '',
    department: r['department'] || '',
    area: r['area'] || '',
    contractor_name: r['contractor name'] || r['contractor'] || '',
    responsible: r['responsible'] || ''
  }));

  renderSummary(data);
  renderTable(data);
  renderCharts(data);
}

function parseYear(dateVal) {
  // try to parse a year from a date cell
  if (!dateVal) return '';
  try {
    const d = new Date(dateVal);
    if (!isNaN(d.getFullYear())) return String(d.getFullYear());
  } catch (e) {}
  // fallback: look for 4-digit year
  const m = String(dateVal).match(/(20\d{2}|19\d{2})/);
  return m ? m[0] : '';
}

function aggCount(rows, key) {
  const counts = {};
  rows.forEach(r => {
    const v = (r[key] || '').toString().trim() || 'Unknown';
    counts[v] = (counts[v] || 0) + 1;
  });
  return counts;
}

function toSortedLabelsAndCounts(counts, topN) {
  const entries = Object.entries(counts).sort((a,b) => b[1]-a[1]);
  if (topN) {
    const other = entries.slice(topN).reduce((s, e) => s + e[1], 0);
    const sliced = entries.slice(0, topN);
    if (other > 0) sliced.push(['Other', other]);
    return sliced.map(e => e[0]), sliced.map(e => e[1]);
  }
  return [entries.map(e => e[0]), entries.map(e => e[1])];
}

function renderSummary(rows) {
  const total = rows.length;
  const open = rows.filter(r => (r.status || '').toLowerCase() !== 'closed').length;
  summaryDiv.innerHTML = `<strong>Total incidents:</strong> ${total} &nbsp; | &nbsp; <strong>Open:</strong> ${open}`;
}

function renderTable(rows) {
  const headers = ['SR','Site','Date','Year','Incident Category','Criticality','Type of Incident','Priority','Owner','Status','Department','Area','Contractor Name'];
  const thead = `<thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>`;
  const tbody = rows.map(r => `<tr>${[
    r.sr, r.site, r.date, r.year, r.incident_category, r.criticality, r.type_of_incident, r.priority, r.owner, r.status, r.department, r.area, r.contractor_name
  ].map(c => `<td>${escapeHtml(c)}</td>`).join('')}</tr>`).join('');
  dataTable.innerHTML = thead + `<tbody>${tbody}</tbody>`;
}

function escapeHtml(s) { return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function destroyCharts() {
  charts.forEach(c => c.destroy());
  charts = [];
}

function renderCharts(rows) {
  destroyCharts();
  // Site chart
  const siteCounts = aggCount(rows, 'site');
  const [siteLabels, siteData] = Object.entries(siteCounts).sort((a,b)=>b[1]-a[1]).reduce((acc, e) => { acc[0].push(e[0]); acc[1].push(e[1]); return acc; }, [[],[]]);

  const ctxSite = document.getElementById('chart-site').getContext('2d');
  charts.push(new Chart(ctxSite, {
    type: 'bar',
    data: { labels: siteLabels, datasets: [{ label: 'Incidents', data: siteData, backgroundColor: '#007bff' }] },
    options: { responsive: true, maintainAspectRatio: false }
  }));

  // Year chart (line)
  const yearCounts = aggCount(rows, 'year');
  const yearEntries = Object.entries(yearCounts).filter(e=>e[0] && e[0] !== 'Unknown').sort((a,b)=>a[0]-b[0]);
  const yearLabels = yearEntries.map(e=>e[0]);
  const yearData = yearEntries.map(e=>e[1]);
  const ctxYear = document.getElementById('chart-year').getContext('2d');
  charts.push(new Chart(ctxYear, { type: 'line', data: { labels: yearLabels, datasets:[{ label: 'Incidents', data: yearData, borderColor: '#28a745', backgroundColor: 'rgba(40,167,69,0.1)', fill: true }] }, options: { responsive: true, maintainAspectRatio: false } }));

  // Category (pie)
  const catCounts = aggCount(rows, 'incident_category');
  const catEntries = Object.entries(catCounts).sort((a,b)=>b[1]-a[1]);
  const catLabels = catEntries.map(e=>e[0]);
  const catData = catEntries.map(e=>e[1]);
  const ctxCat = document.getElementById('chart-category').getContext('2d');
  charts.push(new Chart(ctxCat, { type: 'pie', data: { labels: catLabels, datasets:[{ data: catData, backgroundColor: generateColors(catData.length) }] }, options: { responsive: true, maintainAspectRatio: false } }));

  // Criticality
  const critCounts = aggCount(rows, 'criticality');
  const critEntries = Object.entries(critCounts).sort((a,b)=>b[1]-a[1]);
  const critLabels = critEntries.map(e=>e[0]);
  const critData = critEntries.map(e=>e[1]);
  const ctxCrit = document.getElementById('chart-criticality').getContext('2d');
  charts.push(new Chart(ctxCrit, { type: 'bar', data: { labels: critLabels, datasets:[{ label:'Incidents', data: critData, backgroundColor: '#dc3545' }] }, options: { responsive: true, maintainAspectRatio: false } }));

  // Status
  const statusCounts = aggCount(rows, 'status');
  const statusEntries = Object.entries(statusCounts).sort((a,b)=>b[1]-a[1]);
  const statusLabels = statusEntries.map(e=>e[0]);
  const statusData = statusEntries.map(e=>e[1]);
  const ctxStatus = document.getElementById('chart-status').getContext('2d');
  charts.push(new Chart(ctxStatus, { type: 'doughnut', data: { labels: statusLabels, datasets:[{ data: statusData, backgroundColor: generateColors(statusData.length) }] }, options: { responsive: true, maintainAspectRatio: false } }));
}

function generateColors(n) {
  const palette = ['#3366CC','#DC3912','#FF9900','#109618','#990099','#0099C6','#DD4477','#66AA00','#B82E2E','#316395'];
  const out = [];
  for (let i=0;i<n;i++) out.push(palette[i % palette.length]);
  return out;
}
