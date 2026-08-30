const INJURY_COLUMNS = ['Event ID', 'Event Date', 'Site (Entity\\Location)', 'Location Area', 'Injury Occurrence Status', 'Injury Type', 'Body Position', 'Body Part', 'Actual Consequence', 'Actual Severity', 'Primary Cause', 'Root Cause', 'Body Position.1', 'Injury Location', 'Type', 'Applicable EHS Standards', 'Event Title', 'Title', 'Description', 'Person Involved', 'Gender', 'Departments'];
const REGIONS = ['head', 'shoulder', 'torso', 'arm', 'hand', 'leg', 'foot'];
const REGION_LABELS = { head: 'Head', shoulder: 'Shoulder', torso: 'Torso / back', arm: 'Arm', hand: 'Hand', leg: 'Leg', foot: 'Foot' };
let charts = [];
const normalize = value => String(value ?? '').trim().toLowerCase().replace(/[._]/g, ' ').replace(/\s+/g, ' ');
const read = (row, column) => { const key = Object.keys(row).find(item => normalize(item) === normalize(column)); return key === undefined ? '' : row[key]; };
const escapeHtml = value => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
function counts(rows, column) { return rows.reduce((all, row) => { const key = String(read(row, column)).trim() || 'Not specified'; all[key] = (all[key] || 0) + 1; return all; }, {}); }
function bodyCounts(rows) {
  const result = Object.fromEntries(REGIONS.map(region => [region, 0]));
  rows.forEach(row => {
    const text = normalize([read(row, 'Body Part'), read(row, 'Body Position'), read(row, 'Body Position.1'), read(row, 'Injury Location')].join(' '));
    const matched = new Set();
    if (/head|face|eye|ear|nose|skull|neck/.test(text)) matched.add('head');
    if (/shoulder|clavicle/.test(text)) matched.add('shoulder');
    if (/torso|back|chest|abdomen|trunk|rib|hip/.test(text)) matched.add('torso');
    if (/arm|elbow|wrist/.test(text)) matched.add('arm');
    if (/hand|finger|thumb|palm/.test(text)) matched.add('hand');
    if (/leg|knee|thigh|calf|ankle/.test(text)) matched.add('leg');
    if (/foot|feet|toe|heel/.test(text)) matched.add('foot');
    matched.forEach(region => result[region]++);
  });
  return result;
}
function renderBody(rows) {
  const data = bodyCounts(rows), max = Math.max(...Object.values(data), 1);
  const legend = document.getElementById('body-legend');
  legend.innerHTML = REGIONS.map(region => `<button data-region="${region}"><span>${REGION_LABELS[region]}</span><strong>${data[region]}</strong></button>`).join('');
  document.querySelectorAll('.hit-regions [data-region]').forEach(shape => { const count = data[shape.dataset.region]; shape.style.setProperty('--heat', String(.12 + .78 * count / max)); });
  const callouts = document.getElementById('body-callouts');
  callouts.innerHTML = REGIONS.map(region => `<button data-region="${region}" class="callout callout-${region}"><span>${REGION_LABELS[region]}</span><strong>${data[region]}</strong></button>`).join('');
  const activate = region => {
    document.querySelectorAll('[data-region]').forEach(item => item.classList.toggle('active', item.dataset.region === region));
  };
  [...legend.querySelectorAll('button'), ...callouts.querySelectorAll('button')].forEach(button => button.addEventListener('click', () => activate(button.dataset.region)));
  document.querySelectorAll('.hit-regions [data-region]').forEach(shape => shape.addEventListener('click', () => activate(shape.dataset.region)));
}
function renderChart(id, rows, column, type) {
  const entries = Object.entries(counts(rows, column)).sort((a,b) => b[1]-a[1]).slice(0, 10);
  return new Chart(document.getElementById(id), { type, data: { labels: entries.map(x => x[0]), datasets: [{ label: column, data: entries.map(x => x[1]), backgroundColor: ['#0e7490','#f59e0b','#ef4444','#16a34a','#7c3aed','#0284c7','#db2777','#65a30d','#ea580c','#4f46e5'], borderWidth: 0, borderRadius: type === 'bar' ? 5 : 0 }] }, options: { responsive:true, maintainAspectRatio:false, indexAxis:type === 'bar'?'y':'x', plugins:{legend:{display:type!=='bar',position:'bottom'}}, scales:type==='bar'?{x:{beginAtZero:true,ticks:{precision:0}},y:{grid:{display:false}}}:undefined } });
}
function renderTable(rows) {
  const shown = rows.slice(0, 100); document.getElementById('injury-table').innerHTML = `<thead><tr>${INJURY_COLUMNS.map(c=>`<th>${escapeHtml(c)}</th>`).join('')}</tr></thead><tbody>${shown.map(row=>`<tr>${INJURY_COLUMNS.map(c=>`<td>${escapeHtml(read(row,c))}</td>`).join('')}</tr>`).join('')}</tbody>`;
  document.getElementById('injury-preview-note').textContent = rows.length > 100 ? `Showing the first 100 of ${rows.length} records.` : `Showing all ${rows.length} records.`;
}
function daysSince(rows) { const dates=rows.map(row=>new Date(read(row,'Event Date'))).filter(date=>!Number.isNaN(date.getTime())); if(!dates.length)return '—'; const latest=new Date(Math.max(...dates)); return Math.max(0,Math.floor((Date.now()-latest)/86400000)); }
function render(rows, name) {
  document.getElementById('injury-empty').hidden=true; document.getElementById('injury-dashboard').hidden=false; document.getElementById('injury-file-name').textContent=name; document.getElementById('injury-total').textContent=rows.length; document.getElementById('injury-days').textContent=daysSince(rows);
  document.getElementById('injury-severe').textContent=rows.filter(row=>/high|major|severe|critical|fatal/i.test(String(read(row,'Actual Severity')))).length;
  renderBody(rows); renderTable(rows); charts.forEach(chart=>chart.destroy()); charts=[renderChart('injury-type-chart',rows,'Injury Type','doughnut'),renderChart('injury-severity-chart',rows,'Actual Severity','bar'),renderChart('injury-cause-chart',rows,'Primary Cause','bar')];
}
const input=document.getElementById('injury-input'), status=document.getElementById('injury-upload-status');
input.addEventListener('change',()=>{const file=input.files[0];if(!file)return;status.textContent=`Reading ${file.name}…`;const reader=new FileReader();reader.onload=e=>{try{const workbook=XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true});const rows=XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]],{defval:'',raw:false});if(!rows.length)throw new Error('No records found');render(rows,file.name);status.textContent=`${rows.length} records loaded`;}catch(error){status.textContent=error.message||'Could not read file';status.classList.add('text-danger');}};reader.readAsArrayBuffer(file);});
