// Client-side Excel tracker dashboard. Each supported tracker gets its own panel.

const TRACKERS = {
  incident: {
    title: 'Incident Tracker',
    description: 'Incidents by site, year, category and status.',
    columns: ['SR', 'Site', 'Sources', 'Shift', 'Year', 'Date', 'Incident Category', 'Criticality', 'Type of Incident', 'Event Title', 'GEHSMS Standard', 'Detailed Observation', 'What is the Action?', 'Priority', 'Owner', 'Status', 'New Timeline', 'Department', 'Area', 'Contractor Name', 'Responsible'],
    charts: [['Site', 'bar'], ['Year', 'line'], ['Incident Category', 'doughnut'], ['Status', 'doughnut']]
  },
  bbs: {
    title: 'BBS Tracker',
    description: 'Behaviour-based safety observations and completion progress.',
    columns: ['Completion Date', 'Site (Entity\\Location)', 'Similar Exposure Group', 'Observation Type', 'Created on', 'GPID field', 'Comment'],
    charts: [['Site (Entity\\Location)', 'bar'], ['Observation Type', 'doughnut'], ['Similar Exposure Group', 'bar'], ['Completion Date', 'timeline']]
  },
  moc: {
    title: 'MOC Tracker',
    description: 'Management of change workload and completion status.',
    columns: ['Title', 'MOC Number', 'Initiation Date', 'Initiator', 'Department', 'Status', 'Stage', 'Completion Date', 'Type of Change'],
    charts: [['Status', 'doughnut'], ['Stage', 'bar'], ['Department', 'bar'], ['Type of Change', 'doughnut']]
  },
  gehms: {
    title: 'GEHMS Tracker',
    description: 'GEHSMS assessments, actions, priorities and timelines.',
    columns: ['Action ID', 'Sources', 'Year', 'Date', 'Assessment', 'GEHSMS Standard', 'Detail of clause', 'What is the Action?', 'Priority', 'GPID', 'Name', 'Status', 'Current/Revised Timeline', 'Stage', 'Due/OverDue'],
    charts: [['Status', 'doughnut'], ['Priority', 'bar'], ['GEHSMS Standard', 'bar'], ['Due/OverDue', 'doughnut']]
  },
  observations: {
    title: 'EHS Observations Tracker',
    description: 'Observation trends, responsibilities and corrective actions.',
    columns: ['MM Date 1', 'MM Date 2', 'Shift', 'Observer name', 'Area', 'Permit Applicable', 'Observation Description', 'Relevant GEHSMS standard', 'Why it happened?', 'Action Taken(Corrective)', 'Action to be Taken(Preventive Action)', 'Department', 'Responsibility', 'Criticality', 'Status', 'Timelines', 'EHS Standard Name', 'Type'],
    charts: [['Status', 'doughnut'], ['Criticality', 'bar'], ['Department', 'bar'], ['Type', 'doughnut']]
  },
  ptw: {
    title: 'PTW Tracker',
    description: 'Permit-to-work activity, categories and verification status.',
    columns: ['Serial No', 'PTW no.', 'Permit Type', 'Date', 'Permit Details', 'MOC No.', 'Department', 'Area/Location', 'Shift', 'Start Time', 'End Time', 'Initiator', 'Authorizer', 'Executer', 'EHS verification', 'Working Group', 'Contractor company', 'Contact Number', 'No. of workers', 'LOTO required', 'Permit Extension', 'General work', 'Hot Work', 'Electrical Work', 'Work At height', 'Confined Space', 'Heavy Lifting', 'Demolition /Excavation', 'PTW status', 'Observation of PTW', 'Observation Category'],
    charts: [['Permit Type', 'doughnut'], ['PTW status', 'bar'], ['Department', 'bar'], ['Observation Category', 'doughnut']]
  }
};

const chartInstances = new Map();

function normalizeKey(value) {
  return String(value || '').trim().toLowerCase().replace(/[._]/g, ' ').replace(/\s+/g, ' ');
}

function readValue(row, column) {
  const wanted = normalizeKey(column);
  const actualKey = Object.keys(row).find(key => normalizeKey(key) === wanted);
  return actualKey === undefined ? '' : row[actualKey];
}

function escapeHtml(value) {
  return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

function displayValue(value) {
  if (typeof value === 'number' && value > 25000 && value < 100000 && window.XLSX?.SSF) {
    return XLSX.SSF.format('yyyy-mm-dd', value);
  }
  return value ?? '';
}

function countBy(rows, column) {
  return rows.reduce((counts, row) => {
    const value = String(displayValue(readValue(row, column))).trim() || 'Not specified';
    counts[value] = (counts[value] || 0) + 1;
    return counts;
  }, {});
}

function monthLabel(value) {
  if (typeof value === 'number' && window.XLSX?.SSF) value = XLSX.SSF.format('yyyy-mm-dd', value);
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 'Not specified' : date.toLocaleDateString('en', { year: 'numeric', month: 'short' });
}

function chartData(rows, column, chartType) {
  let counts;
  if (chartType === 'timeline') {
    counts = rows.reduce((result, row) => {
      const label = monthLabel(readValue(row, column));
      result[label] = (result[label] || 0) + 1;
      return result;
    }, {});
  } else {
    counts = countBy(rows, column);
  }
  let entries = Object.entries(counts);
  if (chartType === 'timeline') entries.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  else entries.sort((a, b) => b[1] - a[1]);
  return { labels: entries.map(([label]) => label), values: entries.map(([, value]) => value) };
}

function chartOptions(horizontal = false) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    indexAxis: horizontal ? 'y' : 'x',
    plugins: {
      legend: { position: 'bottom', labels: { boxWidth: 12 } },
      tooltip: { callbacks: { label: context => `${context.label}: ${context.raw}` } }
    },
    scales: horizontal
      ? { x: { beginAtZero: true, ticks: { precision: 0 } }, y: { grid: { display: false } } }
      : undefined
  };
}

function renderCharts(type, rows) {
  const config = TRACKERS[type];
  (chartInstances.get(type) || []).forEach(chart => chart.destroy());
  const instances = config.charts.map(([column, requestedType], index) => {
    const { labels, values } = chartData(rows, column, requestedType);
    const canvas = document.getElementById(`${type}-chart-${index}`);
    const horizontal = requestedType === 'bar';
    const actualType = requestedType === 'timeline' ? 'line' : requestedType;
    return new Chart(canvas, {
      type: actualType,
      data: {
        labels,
        datasets: [{
          label: column,
          data: values,
          backgroundColor: actualType === 'line' ? 'rgba(14, 116, 144, .16)' : colors(values.length),
          borderColor: actualType === 'line' ? '#0e7490' : '#ffffff',
          borderWidth: actualType === 'line' ? 2 : 1,
          fill: actualType === 'line',
          tension: .25
        }]
      },
      options: chartOptions(horizontal)
    });
  });
  chartInstances.set(type, instances);
}

function colors(size) {
  const palette = ['#0e7490', '#f59e0b', '#16a34a', '#7c3aed', '#dc2626', '#0284c7', '#db2777', '#65a30d', '#ea580c', '#4f46e5'];
  return Array.from({ length: size }, (_, index) => palette[index % palette.length]);
}

function renderTable(type, rows) {
  const columns = TRACKERS[type].columns;
  const table = document.getElementById(`${type}-table`);
  const preview = rows.slice(0, 100);
  table.innerHTML = `<thead><tr>${columns.map(column => `<th>${escapeHtml(column)}</th>`).join('')}</tr></thead><tbody>${preview.map(row => `<tr>${columns.map(column => `<td>${escapeHtml(displayValue(readValue(row, column)))}</td>`).join('')}</tr>`).join('')}</tbody>`;
  document.getElementById(`${type}-preview-note`).textContent = rows.length > 100 ? `Showing the first 100 of ${rows.length} rows.` : `Showing all ${rows.length} rows.`;
}

function renderTracker(type, rows, fileName) {
  const section = document.getElementById(`${type}-section`);
  section.hidden = false;
  document.getElementById(`${type}-file-name`).textContent = fileName;
  document.getElementById(`${type}-total`).textContent = rows.length;
  renderTable(type, rows);
  renderCharts(type, rows);
  section.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function showUploadState(type, message, isError = false) {
  const status = document.getElementById(`${type}-upload-status`);
  status.textContent = message;
  status.classList.toggle('text-danger', isError);
}

function loadWorkbook(type, file) {
  showUploadState(type, `Reading ${file.name}…`);
  const reader = new FileReader();
  reader.onload = event => {
    try {
      const workbook = XLSX.read(new Uint8Array(event.target.result), { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
      if (!rows.length) throw new Error('No data rows were found in the first worksheet.');
      renderTracker(type, rows, file.name);
      showUploadState(type, `${rows.length} rows loaded`);
    } catch (error) {
      showUploadState(type, error.message || 'This workbook could not be read.', true);
    }
  };
  reader.onerror = () => showUploadState(type, 'This file could not be read.', true);
  reader.readAsArrayBuffer(file);
}

function trackerMarkup(type, config) {
  return `<section id="${type}-section" class="tracker-section" hidden>
    <div class="section-heading">
      <div><p class="eyebrow">Tracker dashboard</p><h2>${escapeHtml(config.title)}</h2><p>${escapeHtml(config.description)}</p></div>
      <div class="total-pill"><strong id="${type}-total">0</strong><span>records</span></div>
    </div>
    <p class="source-file">Source: <strong id="${type}-file-name"></strong></p>
    <div class="charts-grid">${config.charts.map(([column], index) => `<article class="chart-card"><h3>${escapeHtml(column)}</h3><div class="canvas-wrap"><canvas id="${type}-chart-${index}"></canvas></div></article>`).join('')}</div>
    <details class="data-preview"><summary>View uploaded data</summary><p id="${type}-preview-note" class="preview-note"></p><div class="table-wrap"><table id="${type}-table"></table></div></details>
  </section>`;
}

function initialize() {
  const uploadGrid = document.getElementById('upload-grid');
  const dashboards = document.getElementById('dashboards');
  uploadGrid.innerHTML = Object.entries(TRACKERS).map(([type, config]) => `<label class="upload-card" for="${type}-input"><span class="upload-icon">↗</span><strong>${escapeHtml(config.title)}</strong><span>${config.columns.length} recognized columns</span><input id="${type}-input" type="file" accept=".xlsx,.xls" data-tracker="${type}"><small id="${type}-upload-status">Choose an Excel file</small></label>`).join('');
  dashboards.innerHTML = Object.entries(TRACKERS).map(([type, config]) => trackerMarkup(type, config)).join('');
  uploadGrid.addEventListener('change', event => {
    const input = event.target.closest('input[type="file"]');
    if (input?.files[0]) loadWorkbook(input.dataset.tracker, input.files[0]);
  });
}

initialize();
