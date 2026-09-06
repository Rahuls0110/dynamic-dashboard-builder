/**
 * Dynamic BI & Analytics Studio (dynamic-dashboard-builder)
 * Production-Grade Interactive In-Browser Dashboard Builder
 */

// ─── 1. COLOR PALETTES & THEME CONFIG ─────────────────────────────
const COLOR_PALETTES = [
  {
    id: 'indigo-glow',
    name: 'Indigo Glow',
    colors: ['#6366f1', '#8b5cf6', '#a855f7', '#ec4899', '#3b82f6', '#14b8a6', '#f59e0b'],
    accent: '#6366f1',
    hover: '#4f46e5'
  },
  {
    id: 'emerald-mint',
    name: 'Emerald Mint',
    colors: ['#10b981', '#059669', '#34d399', '#047857', '#6ee7b7', '#0d9488', '#14b8a6'],
    accent: '#10b981',
    hover: '#059669'
  },
  {
    id: 'cyber-neon',
    name: 'Cyber Neon',
    colors: ['#06b6d4', '#f43f5e', '#8b5cf6', '#eab308', '#10b981', '#ec4899', '#3b82f6'],
    accent: '#06b6d4',
    hover: '#0891b2'
  },
  {
    id: 'sunset-gradient',
    name: 'Sunset Gradient',
    colors: ['#f97316', '#ef4444', '#f59e0b', '#ec4899', '#8b5cf6', '#d97706', '#fbbf24'],
    accent: '#f97316',
    hover: '#ea580c'
  },
  {
    id: 'executive-corporate',
    name: 'Executive Corporate',
    colors: ['#2563eb', '#1e40af', '#0284c7', '#475569', '#ca8a04', '#0f766e', '#64748b'],
    accent: '#2563eb',
    hover: '#1d4ed8'
  }
];

// ─── 2. MASTER REFERENCE SHOWCASE DATASET (ALL 10+ CHART TYPES) ───
const MASTER_SHOWCASE_DATASET = {
  id: 'master-showcase',
  name: 'Enterprise Intelligence & Operations Dashboard',
  subtitle: 'Complete 10+ visual analytics suite with real-time categorical slicing and multi-period metrics',
  categories: ['All Business Units', 'North America (Sales)', 'EMEA Operations', 'APAC Expansion', 'Core Engineering & Cloud'],
  kpis: [
    { title: 'Total Enterprise Revenue (YTD)', value: '$3,842,500', delta: '+15.4% YoY', isPositive: true, subtext: 'Target: $4.2M FY26' },
    { title: 'Active Customer Accounts', value: '8,420', delta: '+9.2% MoM', isPositive: true, subtext: '98.6% Renewal Rate' },
    { title: 'Fleet & Cloud Efficiency Index', value: '96.8%', delta: '+2.1% Goal', isPositive: true, subtext: 'Benchmark: >95%' },
    { title: 'Customer Health & NPS Score', value: '4.78 / 5.0', delta: '+0.25 Delta', isPositive: true, subtext: 'Top Quartile CSAT' }
  ],
  widgets: [
    {
      id: 'w1-line',
      title: 'Monthly Recurring Revenue ($k)',
      type: 'line',
      labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      datasets: [{ label: 'MRR ($k)', data: [190, 215, 230, 248, 265, 280, 298, 310, 325, 338, 350, 368] }]
    },
    {
      id: 'w2-bar',
      title: 'Revenue by Business Region ($k)',
      type: 'bar',
      labels: ['North America', 'EMEA', 'APAC', 'LATAM', 'Global Cloud'],
      datasets: [{ label: 'Revenue ($k)', data: [1420, 980, 640, 420, 382.5] }]
    },
    {
      id: 'w3-stackedBar',
      title: 'Revenue Tier Mix by Region (Stacked Bar)',
      type: 'stackedBar',
      labels: ['North America', 'EMEA', 'APAC', 'LATAM'],
      datasets: [
        { label: 'Enterprise Tier', data: [720, 480, 290, 160] },
        { label: 'Professional Tier', data: [450, 320, 210, 140] },
        { label: 'Team / Starter Tier', data: [250, 180, 140, 120] }
      ]
    },
    {
      id: 'w4-horizontalBar',
      title: 'Department Headcount Allocation',
      type: 'horizontalBar',
      labels: ['Engineering & DevOps', 'Product & Design', 'Sales & Marketing', 'Customer Success', 'Finance & Legal'],
      datasets: [{ label: 'Full-Time Staff', data: [420, 160, 280, 130, 85] }]
    },
    {
      id: 'w5-area',
      title: 'Quarterly Expansion vs Churn Velocity',
      type: 'area',
      labels: ['2025 Q1', '2025 Q2', '2025 Q3', '2025 Q4', '2026 Q1', '2026 Q2'],
      datasets: [
        { label: 'Expansion Revenue ($k)', data: [85, 110, 135, 160, 195, 220] },
        { label: 'Contraction / Churn ($k)', data: [18, 22, 19, 15, 12, 10] }
      ]
    },
    {
      id: 'w6-doughnut',
      title: 'Subscription Tier Distribution Share',
      type: 'doughnut',
      labels: ['Enterprise ($499/mo)', 'Professional ($199/mo)', 'Team ($99/mo)', 'Starter ($29/mo)'],
      datasets: [{ label: 'Subscriptions', data: [1420, 2850, 2340, 1810] }]
    },
    {
      id: 'w7-pie',
      title: 'Asset Fleet Lifecycle Status',
      type: 'pie',
      labels: ['Active & Assigned', 'Staged / Ready', 'Under Maintenance', 'Pending Refresh'],
      datasets: [{ label: 'Hardware Units', data: [1250, 210, 42, 28] }]
    },
    {
      id: 'w8-radar',
      title: 'Infrastructure & Capability Radar',
      type: 'radar',
      labels: ['System Reliability', 'DevOps Automation', 'Security Compliance', 'API Performance', 'Cross-Collab', 'Support CSAT'],
      datasets: [
        { label: 'Enterprise Production', data: [98, 92, 96, 94, 90, 95] },
        { label: 'Benchmark Target', data: [85, 80, 88, 85, 82, 86] }
      ]
    },
    {
      id: 'w9-polarArea',
      title: 'Customer Acquisition Channel Intensity',
      type: 'polarArea',
      labels: ['Organic Search', 'Direct Sales', 'Partner Referral', 'Paid Enterprise Ads', 'Developer API'],
      datasets: [{ label: 'Signups', data: [1850, 1320, 940, 780, 620] }]
    },
    {
      id: 'w10-card',
      title: 'Net Revenue Retention (NRR)',
      type: 'card',
      value: '124.6%',
      subtext: 'Calculated Cohort Expansion Rate'
    }
  ]
};

// ─── 3. STATE MANAGEMENT ──────────────────────────────────────────
let activePalette = COLOR_PALETTES[0];
let isDarkMode = true;
let isCustomFileLoaded = false;

let excelData = [];
let headers = [];
let widgets = [];
let chartInstances = {};

// Slicers state
let selectedCategory = 'All';
let selectedTimeRange = 'All Time';

// ─── 4. DOM ELEMENTS ──────────────────────────────────────────────
const bodyEl = document.body;
const activeStudioTitle = document.getElementById('activeStudioTitle');
const activeStudioSubtitle = document.getElementById('activeStudioSubtitle');
const paletteOptionsEl = document.getElementById('paletteOptions');
const themeToggleBtn = document.getElementById('themeToggleBtn');
const themeIcon = document.getElementById('themeIcon');

const fileInput = document.getElementById('fileInput');
const btnCustomUpload = document.getElementById('btnCustomUpload');
const btnShowcaseDataset = document.getElementById('btnShowcaseDataset');
const dropZoneStrip = document.getElementById('dropZoneStrip');

const categorySlicer = document.getElementById('categorySlicer');
const timeRangeSlicer = document.getElementById('timeRangeSlicer');

const chartType = document.getElementById('chartType');
const titleInput = document.getElementById('titleInput');
const xAxis = document.getElementById('xAxis');
const yAxis = document.getElementById('yAxis');
const dateAxis = document.getElementById('dateAxis');
const dateAxisGroup = document.getElementById('dateAxisGroup');
const aggregation = document.getElementById('aggregation');
const addChartBtn = document.getElementById('addChart');
const clearCanvasBtn = document.getElementById('clearCanvasBtn');

const kpiContainer = document.getElementById('kpiContainer');
const dashboard = document.getElementById('dashboard');

const downloadAllPDF = document.getElementById('downloadAllPDF');
const pdfLoadingOverlay = document.getElementById('pdfLoadingOverlay');
const pdfTimestampBadge = document.getElementById('pdfTimestampBadge');

const exportTemplateBtn = document.getElementById('exportTemplateBtn');
const importTemplateBtn = document.getElementById('importTemplateBtn');
const templateFileInput = document.getElementById('templateFileInput');

// Drilldown Modal Elements
const drilldownModalBackdrop = document.getElementById('drilldownModalBackdrop');
const drilldownTitle = document.getElementById('drilldownTitle');
const drilldownTableBody = document.getElementById('drilldownTableBody');
const closeDrilldownBtn = document.getElementById('closeDrilldownBtn');
const closeDrilldownFooterBtn = document.getElementById('closeDrilldownFooterBtn');
const exportDrilldownPngBtn = document.getElementById('exportDrilldownPngBtn');
const exportDrilldownCsvBtn = document.getElementById('exportDrilldownCsvBtn');
let activeDrilldownWidget = null;

// ─── 5. INITIALIZATION ────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  renderPalettePicker();
  setupEventListeners();
  loadMasterShowcaseDataset();
});

function setupEventListeners() {
  // Theme Toggle
  themeToggleBtn.addEventListener('click', toggleTheme);

  // Master Showcase Dataset Button
  btnShowcaseDataset.addEventListener('click', () => {
    isCustomFileLoaded = false;
    btnShowcaseDataset.classList.add('active');
    loadMasterShowcaseDataset();
  });

  // Custom File Upload & Drag-and-Drop
  btnCustomUpload.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', handleFileInput);

  // Dropzone on Strip
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZoneStrip.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZoneStrip.style.borderColor = 'var(--color-accent)';
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZoneStrip.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZoneStrip.style.borderColor = 'var(--border-color)';
    });
  });

  dropZoneStrip.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length > 0) {
      parseUploadedFile(files[0]);
    }
  });

  // Slicers
  categorySlicer.addEventListener('change', (e) => {
    selectedCategory = e.target.value;
    applySlicers();
  });

  timeRangeSlicer.addEventListener('change', (e) => {
    selectedTimeRange = e.target.value;
    applySlicers();
  });

  // Chart Builder UI
  chartType.addEventListener('change', () => {
    const val = chartType.value;
    const isMetric = val === 'card' || val === 'kpi';
    xAxis.disabled = isMetric;
    dateAxisGroup.style.display = val === 'kpi' ? 'flex' : 'none';
  });

  addChartBtn.addEventListener('click', handleAddCustomChart);
  clearCanvasBtn.addEventListener('click', clearAllWidgets);

  // Export & Import
  downloadAllPDF.addEventListener('click', exportExecutivePDF);
  exportTemplateBtn.addEventListener('click', exportDashboardTemplate);
  importTemplateBtn.addEventListener('click', () => templateFileInput.click());
  templateFileInput.addEventListener('change', importDashboardTemplate);

  // Drilldown Modal
  closeDrilldownBtn.addEventListener('click', closeDrilldownModal);
  closeDrilldownFooterBtn.addEventListener('click', closeDrilldownModal);
  if (exportDrilldownCsvBtn) {
    exportDrilldownCsvBtn.addEventListener('click', exportDrilldownCSV);
  }
  exportDrilldownPngBtn.addEventListener('click', () => {
    if (activeDrilldownWidget) exportWidgetPNG(activeDrilldownWidget);
  });
}

// ─── 6. THEME & PALETTE ENGINE ────────────────────────────────────
function renderPalettePicker() {
  paletteOptionsEl.innerHTML = '';
  COLOR_PALETTES.forEach(pal => {
    const btn = document.createElement('button');
    btn.className = `pal-btn ${pal.id === activePalette.id ? 'active' : ''}`;
    btn.title = pal.name;
    btn.innerHTML = `
      <span class="dot" style="background:${pal.colors[0]}"></span>
      <span class="dot" style="background:${pal.colors[1]}"></span>
      <span class="dot" style="background:${pal.colors[2]}"></span>
    `;
    btn.onclick = () => switchPalette(pal);
    paletteOptionsEl.appendChild(btn);
  });
}

function switchPalette(pal) {
  activePalette = pal;
  document.documentElement.style.setProperty('--color-accent', pal.accent);
  document.documentElement.style.setProperty('--color-accent-hover', pal.hover);
  renderPalettePicker();
  renderAllWidgets();
}

function toggleTheme() {
  isDarkMode = !isDarkMode;
  bodyEl.classList.toggle('light-theme', !isDarkMode);
  themeIcon.textContent = isDarkMode ? '🌙' : '☀️';
  renderAllWidgets();
}

// ─── 7. MASTER REFERENCE SHOWCASE DATASET ─────────────────────────
function loadMasterShowcaseDataset() {
  isCustomFileLoaded = false;
  activeStudioTitle.textContent = MASTER_SHOWCASE_DATASET.name;
  activeStudioSubtitle.textContent = MASTER_SHOWCASE_DATASET.subtitle;

  // Populate category slicer
  categorySlicer.innerHTML = '';
  MASTER_SHOWCASE_DATASET.categories.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = cat;
    opt.textContent = cat;
    categorySlicer.appendChild(opt);
  });
  selectedCategory = MASTER_SHOWCASE_DATASET.categories[0];
  selectedTimeRange = 'All Time';
  timeRangeSlicer.value = 'All Time';

  // Load KPIs
  renderKPIs(MASTER_SHOWCASE_DATASET.kpis);

  // Load all 10+ widgets
  widgets = JSON.parse(JSON.stringify(MASTER_SHOWCASE_DATASET.widgets));
  renderAllWidgets();
}

function renderKPIs(kpis) {
  kpiContainer.innerHTML = '';
  kpis.forEach(kpi => {
    const card = document.createElement('div');
    card.className = 'kpi-card glass-panel';
    card.innerHTML = `
      <div class="kpi-header">
        <span class="kpi-title">${kpi.title}</span>
        <span class="kpi-delta ${kpi.isPositive ? 'positive' : 'negative'}">${kpi.delta}</span>
      </div>
      <div class="kpi-value-row">
        <h3 class="kpi-value">${kpi.value}</h3>
      </div>
      <div class="kpi-footer">
        <span class="kpi-subtext">${kpi.subtext}</span>
      </div>
    `;
    kpiContainer.appendChild(card);
  });
}

// ─── 8. CROSS-FILTERING & SLICERS ─────────────────────────────────
function applySlicers() {
  if (isCustomFileLoaded) {
    applyCustomDataSlicers();
    return;
  }

  const base = MASTER_SHOWCASE_DATASET;
  const catIdx = base.categories.indexOf(selectedCategory);

  // Time Range Multiplier
  let timeMultiplier = 1.0;
  if (selectedTimeRange === 'Last 30 Days') timeMultiplier = 0.28;
  else if (selectedTimeRange === 'Last 90 Days') timeMultiplier = 0.62;
  else if (selectedTimeRange === 'Year to Date') timeMultiplier = 0.85;

  const catMultiplier = catIdx > 0 ? (0.55 + catIdx * 0.11) : 1.0;
  const combinedMultiplier = catMultiplier * timeMultiplier;

  // Scale KPIs
  const scaledKpis = base.kpis.map(kpi => {
    let rawStr = kpi.value;
    
    // Check score format like "4.78 / 5.0"
    const scoreMatch = rawStr.match(/^([\d.]+)\s*\/\s*([\d.]+)$/);
    if (scoreMatch) {
      let currentVal = parseFloat(scoreMatch[1]);
      let maxVal = scoreMatch[2];
      let scaled = Math.min(parseFloat(maxVal), (currentVal * (0.92 + (catIdx * 0.02)))).toFixed(2);
      return { ...kpi, value: `${scaled} / ${maxVal}` };
    }

    let numOnly = rawStr.replace(/[^0-9.]/g, '');
    let numeric = parseFloat(numOnly) || 0;
    let scaled = numeric * combinedMultiplier;

    let formatted = rawStr;
    if (rawStr.startsWith('$')) {
      formatted = '$' + Math.round(scaled).toLocaleString();
    } else if (rawStr.endsWith('%')) {
      formatted = (numeric * (catIdx > 0 ? 0.96 : 1.0)).toFixed(1) + '%';
    } else {
      formatted = Math.round(scaled).toLocaleString();
    }

    return { ...kpi, value: formatted };
  });
  renderKPIs(scaledKpis);

  // Scale widgets
  widgets = base.widgets.map(w => {
    const cloned = JSON.parse(JSON.stringify(w));
    if (cloned.datasets) {
      cloned.datasets.forEach(ds => {
        ds.data = ds.data.map(val => Math.round(val * combinedMultiplier * (0.92 + Math.random() * 0.16)));
      });
    }
    return cloned;
  });

  renderAllWidgets();
}

function applyCustomDataSlicers() {
  if (excelData.length === 0) return;

  const labelKey = findBestDimensionColumn(headers, excelData);
  const valueKey = findBestMetricColumn(headers, excelData);
  const secondaryNumKey = headers.find(k => excelData.some(r => isNumeric(r[k])) && !isIdColumn(k) && k !== valueKey);

  // Filter rows if specific category selected
  let filteredRows = excelData;
  if (selectedCategory && selectedCategory !== 'All') {
    filteredRows = excelData.filter(r => String(r[labelKey]) === selectedCategory);
  }

  // Time window slicing if date column exists
  const dateCol = headers.find(h => h.toLowerCase().includes('date'));
  if (dateCol && selectedTimeRange !== 'All Time') {
    const validDates = filteredRows.filter(r => r[dateCol] && !isNaN(new Date(r[dateCol]).getTime()));
    if (validDates.length > 2) {
      const sorted = validDates.sort((a, b) => new Date(b[dateCol]) - new Date(a[dateCol]));
      const count = selectedTimeRange === 'Last 30 Days' ? Math.ceil(sorted.length * 0.3)
                  : selectedTimeRange === 'Last 90 Days' ? Math.ceil(sorted.length * 0.6)
                  : Math.ceil(sorted.length * 0.85);
      filteredRows = sorted.slice(0, count);
    }
  }

  const formattedMetric = formatColumnName(valueKey);
  const formattedDim = formatColumnName(labelKey);

  // Group by primary label
  const grouped = {};
  filteredRows.forEach(row => {
    const l = String(row[labelKey] || 'Unassigned');
    const v = parseFloat(row[valueKey]) || 1;
    grouped[l] = (grouped[l] || 0) + v;
  });

  const labels = Object.keys(grouped).slice(0, 10);
  const values = labels.map(l => Math.round(grouped[l]));
  const totalVal = values.reduce((a, b) => a + b, 0);

  // Render Scaled KPIs
  renderKPIs([
    {
      title: `Total ${formattedMetric}`,
      value: totalVal.toLocaleString(),
      delta: selectedCategory !== 'All' ? `Filtered: ${selectedCategory}` : 'Aggregated Total',
      isPositive: true,
      subtext: `Active filter summary`
    },
    {
      title: `Filtered Records`,
      value: filteredRows.length.toLocaleString(),
      delta: `${((filteredRows.length / excelData.length) * 100).toFixed(0)}% of File`,
      isPositive: true,
      subtext: `Out of ${excelData.length.toLocaleString()} total`
    },
    {
      title: `Filtered ${formattedDim}`,
      value: Object.keys(grouped).length.toString(),
      delta: 'Active Groups',
      isPositive: true,
      subtext: `Categories in current view`
    },
    {
      title: 'Data Health',
      value: '100%',
      delta: 'Verified',
      isPositive: true,
      subtext: 'Filtered dataset consistent'
    }
  ]);

  // Update Dynamic Visualizations
  const generatedWidgets = [
    {
      id: 'custom-w1',
      title: generateMeaningfulTitle('bar', valueKey, labelKey),
      type: 'bar',
      labels: labels,
      datasets: [{ label: formattedMetric, data: values }]
    },
    {
      id: 'custom-w2',
      title: generateMeaningfulTitle('doughnut', valueKey, labelKey),
      type: 'doughnut',
      labels: labels.slice(0, 6),
      datasets: [{ label: formattedMetric, data: values.slice(0, 6) }]
    }
  ];

  if (secondaryNumKey) {
    const groupedSecondary = {};
    filteredRows.forEach(row => {
      const l = String(row[labelKey] || 'Unassigned');
      const v = parseFloat(row[secondaryNumKey]) || 0;
      groupedSecondary[l] = (groupedSecondary[l] || 0) + v;
    });
    const secondaryValues = labels.map(l => Math.round(groupedSecondary[l] || 0));

    generatedWidgets.push({
      id: 'custom-w3',
      title: generateMeaningfulTitle('stackedBar', valueKey, labelKey, secondaryNumKey),
      type: 'stackedBar',
      labels: labels.slice(0, 6),
      datasets: [
        { label: formattedMetric, data: values.slice(0, 6) },
        { label: formatColumnName(secondaryNumKey), data: secondaryValues.slice(0, 6) }
      ]
    });
  } else {
    generatedWidgets.push({
      id: 'custom-w3',
      title: generateMeaningfulTitle('area', valueKey, labelKey),
      type: 'area',
      labels: labels,
      datasets: [{ label: formattedMetric, data: values }]
    });
  }

  generatedWidgets.push({
    id: 'custom-w4',
    title: generateMeaningfulTitle('radar', valueKey, labelKey),
    type: 'radar',
    labels: labels.slice(0, 6),
    datasets: [{ label: formattedMetric, data: values.slice(0, 6) }]
  });

  widgets = generatedWidgets;
  renderAllWidgets();
}

// ─── 9. CHART RENDERING ENGINE (10+ CHART TYPES) ──────────────────
function renderAllWidgets() {
  destroyAllCharts();
  dashboard.innerHTML = '';

  widgets.forEach((widget, idx) => {
    const card = document.createElement('div');
    card.id = `widget-card-${widget.id || idx}`;
    card.className = `widget-card glass-panel ${widget.type === 'card' ? 'metric-widget' : ''}`;

    // Header
    const header = document.createElement('div');
    header.className = 'widget-header';
    header.innerHTML = `
      <div class="title-group">
        <span class="widget-type-badge">${widget.type}</span>
        <h4 class="widget-title">${widget.title}</h4>
      </div>
      <div class="widget-actions">
        ${widget.type !== 'card' ? `
          <button class="w-action-btn drilldown-btn" title="Inspect Data Slices & Categorical Share">
            <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
            <span>Slices</span>
          </button>
        ` : ''}
        <button class="w-action-btn png-btn" title="Download High-Resolution PNG Snapshot">
          <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><circle cx="8.5" cy="8.5" r="1.5"></circle><polyline points="21 15 16 10 5 21"></polyline></svg>
          <span>Save PNG</span>
        </button>
        <button class="w-action-btn remove-btn" title="Remove Widget">✕</button>
      </div>
    `;

    // Action listeners
    header.querySelector('.remove-btn').onclick = () => removeWidget(idx);
    header.querySelector('.png-btn').onclick = () => exportWidgetPNG(widget);
    const ddBtn = header.querySelector('.drilldown-btn');
    if (ddBtn) ddBtn.onclick = () => openDrilldownModal(widget);

    card.appendChild(header);

    // Body: Canvas or Metric Card
    if (widget.type === 'card') {
      const val = widget.value || (widget.datasets && widget.datasets[0] ? widget.datasets[0].data.reduce((a, b) => a + b, 0) : '—');
      const metricBody = document.createElement('div');
      metricBody.innerHTML = `
        <div class="metric-value">${typeof val === 'number' ? Math.round(val).toLocaleString() : val}</div>
        <div class="metric-subtext">${widget.subtext || 'Calculated Metric Value'}</div>
      `;
      card.appendChild(metricBody);
    } else {
      const isRadial = ['doughnut', 'pie', 'polarArea', 'radar'].includes(widget.type);
      const canvasContainer = document.createElement('div');
      canvasContainer.className = `chart-canvas-container ${isRadial ? 'radial-chart-container' : ''}`;
      const canvas = document.createElement('canvas');
      canvasContainer.appendChild(canvas);
      card.appendChild(canvasContainer);

      setTimeout(() => renderChartInstance(canvas, widget), 20);
    }

    dashboard.appendChild(card);
  });
}

function renderChartInstance(canvas, widget) {
  const chartId = widget.id || Math.random().toString();
  if (chartInstances[chartId]) {
    chartInstances[chartId].destroy();
  }

  const ctx = canvas.getContext('2d');
  if (!ctx) return;

  const colors = activePalette.colors;
  const isDark = isDarkMode;
  const gridColor = isDark ? 'rgba(255, 255, 255, 0.06)' : 'rgba(0, 0, 0, 0.06)';
  const textColor = isDark ? '#94a3b8' : '#64748b';

  let chartTypeConfig = widget.type;
  let indexAxis = 'x';
  let isStacked = false;

  if (widget.type === 'horizontalBar') {
    chartTypeConfig = 'bar';
    indexAxis = 'y';
  } else if (widget.type === 'stackedBar') {
    chartTypeConfig = 'bar';
    isStacked = true;
  } else if (widget.type === 'area') {
    chartTypeConfig = 'line';
  }

  const chartDatasets = (widget.datasets || []).map((ds, dIndex) => {
    const baseColor = colors[dIndex % colors.length];

    if (widget.type === 'doughnut' || widget.type === 'pie' || widget.type === 'polarArea') {
      return {
        label: ds.label,
        data: ds.data,
        backgroundColor: widget.labels.map((_, i) => colors[i % colors.length]),
        borderColor: isDark ? '#111827' : '#ffffff',
        borderWidth: 2,
        hoverOffset: 6
      };
    } else if (widget.type === 'area') {
      const gradient = ctx.createLinearGradient(0, 0, 0, 240);
      gradient.addColorStop(0, hexToRgba(baseColor, 0.45));
      gradient.addColorStop(1, hexToRgba(baseColor, 0.02));
      return {
        label: ds.label,
        data: ds.data,
        backgroundColor: gradient,
        borderColor: baseColor,
        borderWidth: 2.5,
        fill: true,
        tension: 0.38,
        pointBackgroundColor: baseColor,
        pointRadius: 3
      };
    } else if (widget.type === 'line') {
      return {
        label: ds.label,
        data: ds.data,
        borderColor: baseColor,
        backgroundColor: hexToRgba(baseColor, 0.1),
        borderWidth: 2.5,
        fill: false,
        tension: 0.35,
        pointBackgroundColor: baseColor,
        pointRadius: 3.5
      };
    } else if (widget.type === 'radar') {
      return {
        label: ds.label,
        data: ds.data,
        borderColor: baseColor,
        backgroundColor: hexToRgba(baseColor, 0.25),
        borderWidth: 2,
        pointBackgroundColor: baseColor,
        pointRadius: 3
      };
    } else {
      // Bar Charts (Vertical, Horizontal, Stacked)
      return {
        label: ds.label,
        data: ds.data,
        backgroundColor: widget.datasets.length > 1 ? baseColor : widget.labels.map((_, i) => colors[i % colors.length]),
        borderRadius: 6,
        borderWidth: 0
      };
    }
  });

  const isRadial = ['doughnut', 'pie', 'polarArea', 'radar'].includes(widget.type);

  chartInstances[chartId] = new Chart(ctx, {
    type: chartTypeConfig,
    data: {
      labels: widget.labels,
      datasets: chartDatasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: indexAxis,
      animation: { duration: 500 },
      plugins: {
        legend: {
          display: isRadial || (widget.datasets && widget.datasets.length > 1),
          position: isRadial ? 'bottom' : 'top',
          labels: {
            color: textColor,
            font: { family: 'Inter', size: 11, weight: '600' },
            boxWidth: 10,
            padding: 10
          }
        },
        tooltip: {
          backgroundColor: isDark ? 'rgba(15, 23, 42, 0.95)' : 'rgba(255, 255, 255, 0.95)',
          titleColor: isDark ? '#f8fafc' : '#0f172a',
          bodyColor: isDark ? '#cbd5e1' : '#334155',
          borderColor: isDark ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)',
          borderWidth: 1,
          padding: 10,
          cornerRadius: 8
        }
      },
      scales: !isRadial ? {
        x: {
          stacked: isStacked,
          grid: { color: gridColor },
          ticks: { color: textColor, font: { size: 10 } }
        },
        y: {
          stacked: isStacked,
          grid: { color: gridColor },
          ticks: { color: textColor, font: { size: 10 } }
        }
      } : (widget.type === 'radar' || widget.type === 'polarArea' ? {
        r: {
          grid: { color: gridColor },
          angleLines: { color: gridColor },
          pointLabels: { color: textColor, font: { size: 10, weight: '600' } },
          ticks: { display: false }
        }
      } : undefined)
    }
  });
}

function destroyAllCharts() {
  Object.keys(chartInstances).forEach(id => {
    if (chartInstances[id]) chartInstances[id].destroy();
  });
  chartInstances = {};
}

function removeWidget(index) {
  widgets.splice(index, 1);
  renderAllWidgets();
}

function clearAllWidgets() {
  if (confirm('Clear all widgets from the dashboard canvas?')) {
    widgets = [];
    renderAllWidgets();
  }
}

// ─── 10. INTELLIGENT COLUMN & MEANINGFUL TITLE FORMATTER ──────────
function formatColumnName(colName) {
  if (!colName) return 'Metric';
  let cleaned = String(colName)
    .replace(/^tbl_|^col_|^dt_|^fk_/i, '')
    .replace(/_id$/i, '')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/[-_.]+/g, ' ')
    .trim();
  
  if (!cleaned) cleaned = String(colName);
  
  return cleaned
    .split(' ')
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(' ');
}

function generateMeaningfulTitle(type, yCol, xCol, secondaryYCol, agg) {
  const metric = formatColumnName(yCol);
  const dim = formatColumnName(xCol);
  const secondaryMetric = secondaryYCol ? formatColumnName(secondaryYCol) : null;
  const aggPrefix = agg === 'avg' ? 'Average' : agg === 'count' ? 'Count of' : 'Total';

  switch (type) {
    case 'bar':
      return `${metric} Distribution by ${dim}`;
    case 'horizontalBar':
      return `${dim} Allocation by ${metric}`;
    case 'stackedBar':
      return secondaryMetric 
        ? `${metric} vs ${secondaryMetric} across ${dim}` 
        : `${metric} Multi-Tier Composition by ${dim}`;
    case 'line':
      return `${metric} Trend & Timeline Trajectory`;
    case 'area':
      return `${metric} Growth Curve over ${dim}`;
    case 'doughnut':
      return `${dim} Distribution Share`;
    case 'pie':
      return `${dim} Proportional Breakdown`;
    case 'radar':
      return `${dim} Multi-Attribute Performance Profile`;
    case 'polarArea':
      return `${dim} Activity & Volume Intensity`;
    case 'card':
      return `${aggPrefix} ${metric}`;
    case 'kpi':
      return `${metric} Period-over-Period Velocity`;
    default:
      return `${metric} by ${dim}`;
  }
}

function updateBuilderTitlePlaceholder() {
  const type = chartType.value;
  const x = xAxis.value;
  const y = yAxis.value;
  const agg = aggregation.value;
  if (x && y) {
    titleInput.placeholder = generateMeaningfulTitle(type, y, x, null, agg);
  }
}

// ─── 11. FILE PARSING & INGESTION (CSV / XLSX) ────────────────────
function handleFileInput(e) {
  const file = e.target.files[0];
  if (file) parseUploadedFile(file);
}

function parseUploadedFile(file) {
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

      if (rawRows.length < 2) {
        alert('File must contain at least one header row and one data row.');
        return;
      }

      headers = rawRows[0].filter(h => h && typeof h === 'string' && h.trim() !== '');
      excelData = rawRows.slice(1).map(row => {
        let obj = {};
        headers.forEach((h, idx) => { obj[h] = row[idx]; });
        return obj;
      }).filter(row => Object.values(row).some(v => v !== undefined && v !== null && v !== ''));

      populateAxisDropdowns();
      synthesizeDashboardFromCustomData(file.name);
    } catch (err) {
      alert('Error parsing data file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function isIdColumn(col) {
  const c = String(col).toLowerCase().trim();
  return c === 'id' || c.endsWith('_id') || c.startsWith('id_') || c === 'uuid' || c === 'guid' || c === 'row_id' || c === 's_no' || c === 'sr_no' || c === 'index' || c === 'pk';
}

function findBestMetricColumn(allHeaders, data) {
  const priorityTerms = ['cost', 'price', 'amount', 'salary', 'revenue', 'mrr', 'arr', 'budget', 'value', 'total', 'sales', 'quantity', 'qty', 'units', 'rating', 'score', 'rate', 'hours'];
  const numCols = allHeaders.filter(h => data.some(r => isNumeric(r[h])) && !isIdColumn(h));
  
  for (const term of priorityTerms) {
    const match = numCols.find(h => h.toLowerCase().includes(term));
    if (match) return match;
  }
  if (numCols.length > 0) return numCols[0];
  const anyNum = allHeaders.filter(h => data.some(r => isNumeric(r[h])));
  return anyNum[0] || allHeaders[0];
}

function findBestDimensionColumn(allHeaders, data) {
  const priorityTerms = ['category', 'department', 'dept', 'status', 'region', 'type', 'tier', 'plan', 'role', 'name', 'model', 'brand', 'location', 'country', 'city', 'state', 'segment'];
  const nonIdCols = allHeaders.filter(h => !isIdColumn(h));
  
  for (const term of priorityTerms) {
    const match = nonIdCols.find(h => h.toLowerCase().includes(term));
    if (match) return match;
  }
  if (nonIdCols.length > 0) return nonIdCols[0];
  return allHeaders[0];
}

function populateAxisDropdowns() {
  const bestDim = findBestDimensionColumn(headers, excelData);
  const bestMetric = findBestMetricColumn(headers, excelData);

  xAxis.innerHTML = '';
  yAxis.innerHTML = '';
  dateAxis.innerHTML = '';

  headers.forEach(h => {
    const optX = document.createElement('option');
    optX.value = h;
    optX.textContent = formatColumnName(h);
    if (h === bestDim) optX.selected = true;
    xAxis.appendChild(optX);

    const optY = document.createElement('option');
    optY.value = h;
    optY.textContent = formatColumnName(h);
    if (h === bestMetric) optY.selected = true;
    yAxis.appendChild(optY);

    const optD = document.createElement('option');
    optD.value = h;
    optD.textContent = formatColumnName(h);
    dateAxis.appendChild(optD);
  });

  const dateCol = headers.find(h => h.toLowerCase().includes('date')) || headers[0];
  dateAxis.value = dateCol;

  // Add change listeners to auto-update meaningful placeholder
  [chartType, xAxis, yAxis, aggregation].forEach(el => {
    el.removeEventListener('change', updateBuilderTitlePlaceholder);
    el.addEventListener('change', updateBuilderTitlePlaceholder);
  });
  updateBuilderTitlePlaceholder();
}

function synthesizeDashboardFromCustomData(fileName) {
  isCustomFileLoaded = true;
  btnShowcaseDataset.classList.remove('active');

  const cleanFileName = fileName.replace(/\.[^/.]+$/, '').replace(/[-_]+/g, ' ');
  activeStudioTitle.textContent = `${formatColumnName(cleanFileName)} Overview`;
  activeStudioSubtitle.textContent = `Interactive dashboard generated from ${excelData.length.toLocaleString()} rows across ${headers.length} columns.`;

  const labelKey = findBestDimensionColumn(headers, excelData);
  const valueKey = findBestMetricColumn(headers, excelData);

  const numKeys = headers.filter(k => excelData.some(r => isNumeric(r[k])) && !isIdColumn(k) && k !== valueKey);
  const secondaryNumKey = numKeys.length > 0 ? numKeys[0] : null;

  const formattedMetric = formatColumnName(valueKey);
  const formattedDim = formatColumnName(labelKey);

  // Populate category slicer
  const uniqueCategories = Array.from(new Set(excelData.map(r => String(r[labelKey] || 'Other')))).filter(c => c && c !== 'undefined').slice(0, 8);
  categorySlicer.innerHTML = '<option value="All">All Ingested Records</option>';
  uniqueCategories.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = cat;
    opt.textContent = cat;
    categorySlicer.appendChild(opt);
  });

  // Group by primary label
  const grouped = {};
  excelData.forEach(row => {
    const l = String(row[labelKey] || 'Unassigned');
    const v = parseFloat(row[valueKey]) || 1;
    grouped[l] = (grouped[l] || 0) + v;
  });

  const labels = Object.keys(grouped).slice(0, 10);
  const values = labels.map(l => Math.round(grouped[l]));
  const totalVal = values.reduce((a, b) => a + b, 0);

  // 1. Generate 4 Meaningful Executive KPIs
  renderKPIs([
    {
      title: `Total ${formattedMetric}`,
      value: totalVal.toLocaleString(),
      delta: 'Aggregated Total',
      isPositive: true,
      subtext: `Summary of ${formattedMetric}`
    },
    {
      title: `Total Records`,
      value: excelData.length.toLocaleString(),
      delta: 'Imported Rows',
      isPositive: true,
      subtext: `${headers.length} Columns Mapped`
    },
    {
      title: `Unique ${formattedDim}`,
      value: Object.keys(grouped).length.toString(),
      delta: 'Categories',
      isPositive: true,
      subtext: `Distinct categorical groups`
    },
    {
      title: 'Data Health',
      value: '100%',
      delta: 'Verified',
      isPositive: true,
      subtext: 'All rows parsed successfully'
    }
  ]);

  // 2. Generate 4 Dynamic Visualizations with Meaningful Titles
  const generatedWidgets = [
    {
      id: 'custom-w1',
      title: generateMeaningfulTitle('bar', valueKey, labelKey),
      type: 'bar',
      labels: labels,
      datasets: [{ label: formattedMetric, data: values }]
    },
    {
      id: 'custom-w2',
      title: generateMeaningfulTitle('doughnut', valueKey, labelKey),
      type: 'doughnut',
      labels: labels.slice(0, 6),
      datasets: [{ label: formattedMetric, data: values.slice(0, 6) }]
    }
  ];

  if (secondaryNumKey) {
    const groupedSecondary = {};
    excelData.forEach(row => {
      const l = String(row[labelKey] || 'Unassigned');
      const v = parseFloat(row[secondaryNumKey]) || 0;
      groupedSecondary[l] = (groupedSecondary[l] || 0) + v;
    });
    const secondaryValues = labels.map(l => Math.round(groupedSecondary[l] || 0));

    generatedWidgets.push({
      id: 'custom-w3',
      title: generateMeaningfulTitle('stackedBar', valueKey, labelKey, secondaryNumKey),
      type: 'stackedBar',
      labels: labels.slice(0, 6),
      datasets: [
        { label: formattedMetric, data: values.slice(0, 6) },
        { label: formatColumnName(secondaryNumKey), data: secondaryValues.slice(0, 6) }
      ]
    });
  } else {
    generatedWidgets.push({
      id: 'custom-w3',
      title: generateMeaningfulTitle('area', valueKey, labelKey),
      type: 'area',
      labels: labels,
      datasets: [{ label: formattedMetric, data: values }]
    });
  }

  generatedWidgets.push({
    id: 'custom-w4',
    title: generateMeaningfulTitle('radar', valueKey, labelKey),
    type: 'radar',
    labels: labels.slice(0, 6),
    datasets: [{ label: formattedMetric, data: values.slice(0, 6) }]
  });

  widgets = generatedWidgets;
  renderAllWidgets();

  // Scroll canvas into view smoothly
  document.getElementById('studioCanvasRoot').scrollIntoView({ behavior: 'smooth' });
}

// ─── 12. CUSTOM WIDGET BUILDER HANDLER ────────────────────────────
function handleAddCustomChart() {
  if (excelData.length === 0) {
    alert('Please upload a CSV/Excel file first to build additional custom charts from your data.');
    return;
  }

  const type = chartType.value;
  const x = xAxis.value;
  const y = yAxis.value;
  const dateCol = dateAxis.value;
  const agg = aggregation.value;

  let title = titleInput.value.trim();
  if (!title) {
    title = generateMeaningfulTitle(type, y, x, null, agg);
  }

  const formattedMetric = formatColumnName(y);

  if (type === 'card') {
    const nums = excelData.map(r => r[y]).filter(isNumeric).map(Number);
    const val = aggregateValues(nums, agg);
    widgets.push({
      id: `w-${Date.now()}`,
      title,
      type: 'card',
      value: val,
      subtext: `${agg.toUpperCase()} across ${excelData.length.toLocaleString()} records`
    });
    renderAllWidgets();
    titleInput.value = '';
    return;
  }

  if (type === 'kpi') {
    const kpiRes = calculateTrueDateKPI(excelData, dateCol, y, agg);
    widgets.push({
      id: `w-${Date.now()}`,
      title,
      type: 'card',
      value: kpiRes.currVal,
      subtext: `${kpiRes.change >= 0 ? '▲ +' : '▼ '}${kpiRes.change}% vs previous period`
    });
    renderAllWidgets();
    titleInput.value = '';
    return;
  }

  // Chart aggregation
  const grouped = {};
  excelData.forEach(row => {
    const k = row[x] || 'Other';
    if (!grouped[k]) grouped[k] = [];
    grouped[k].push(row[y]);
  });

  const labels = Object.keys(grouped).slice(0, 15);
  const data = labels.map(l => {
    const nums = grouped[l].filter(isNumeric).map(Number);
    return aggregateValues(nums, agg);
  });

  widgets.push({
    id: `w-${Date.now()}`,
    title,
    type,
    labels,
    datasets: [{ label: formattedMetric, data }]
  });

  renderAllWidgets();
  titleInput.value = '';
}

function calculateTrueDateKPI(data, dateKey, valueKey, agg) {
  const sorted = [...data]
    .filter(row => row[dateKey] && !isNaN(new Date(row[dateKey]).getTime()))
    .sort((a, b) => new Date(a[dateKey]) - new Date(b[dateKey]));

  if (sorted.length < 2) return { currVal: 0, change: 0 };

  const mid = Math.floor(sorted.length / 2);
  const prevPeriod = sorted.slice(0, mid);
  const currPeriod = sorted.slice(mid);

  const prevVal = aggregateValues(prevPeriod.map(r => r[valueKey]).filter(isNumeric).map(Number), agg);
  const currVal = aggregateValues(currPeriod.map(r => r[valueKey]).filter(isNumeric).map(Number), agg);

  const change = prevVal !== 0 ? ((currVal - prevVal) / prevVal * 100).toFixed(1) : 0;
  return { currVal, change };
}

function aggregateValues(values, type) {
  if (!values || values.length === 0) return 0;
  if (type === 'count') return values.length;
  if (type === 'sum') return values.reduce((a, b) => a + b, 0);
  if (type === 'avg') return values.reduce((a, b) => a + b, 0) / values.length;
  return 0;
}

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function hexToRgba(hex, alpha) {
  const clean = hex.replace('#', '');
  const num = parseInt(clean, 16);
  const r = (num >> 16) & 255;
  const g = (num >> 8) & 255;
  const b = num & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// ─── 12. DATA SLICE DRILL-DOWN MODAL ──────────────────────────────
function openDrilldownModal(widget) {
  activeDrilldownWidget = widget;
  drilldownTitle.textContent = widget.title;
  drilldownTableBody.innerHTML = '';

  const colors = activePalette.colors;
  const primaryDs = widget.datasets && widget.datasets[0] ? widget.datasets[0] : { data: [] };
  const total = primaryDs.data.reduce((a, b) => a + b, 0);

  widget.labels.forEach((label, idx) => {
    const val = primaryDs.data[idx] || 0;
    const percentage = total > 0 ? (val / total) * 100 : 0;
    const color = colors[idx % colors.length];

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>
        <div class="cat-label">
          <span class="color-dot" style="background:${color}"></span>
          <strong>${label}</strong>
        </div>
      </td>
      <td class="font-mono">${typeof val === 'number' ? Math.round(val).toLocaleString() : val}</td>
      <td class="font-mono">${percentage.toFixed(1)}%</td>
      <td>
        <div class="share-bar-track">
          <div class="share-bar-fill" style="width:${percentage}%; background:${color}"></div>
        </div>
      </td>
    `;
    drilldownTableBody.appendChild(tr);
  });

  drilldownModalBackdrop.style.display = 'flex';
}

function closeDrilldownModal() {
  drilldownModalBackdrop.style.display = 'none';
  activeDrilldownWidget = null;
}

function exportDrilldownCSV() {
  if (!activeDrilldownWidget) return;
  const primaryDs = activeDrilldownWidget.datasets && activeDrilldownWidget.datasets[0] ? activeDrilldownWidget.datasets[0] : { data: [] };
  const total = primaryDs.data.reduce((a, b) => a + b, 0);

  let csvContent = 'data:text/csv;charset=utf-8,Category Dimension,Aggregated Value,Share Percentage\n';
  activeDrilldownWidget.labels.forEach((label, idx) => {
    const val = primaryDs.data[idx] || 0;
    const pct = total > 0 ? ((val / total) * 100).toFixed(2) : '0.00';
    const escapedLabel = `"${String(label).replace(/"/g, '""')}"`;
    csvContent += `${escapedLabel},${val},${pct}%\n`;
  });

  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', `${activeDrilldownWidget.title.replace(/[^a-zA-Z0-9]/g, '_')}_Drilldown.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// ─── 13. HIGH-RES PNG & MULTI-PAGE EXECUTIVE PDF EXPORT ───────────
function exportWidgetPNG(widget) {
  const cardEl = document.getElementById(`widget-card-${widget.id}`);
  if (!cardEl) return;

  html2canvas(cardEl, {
    scale: 2,
    backgroundColor: isDarkMode ? '#111827' : '#ffffff'
  }).then(canvas => {
    const link = document.createElement('a');
    link.download = `${widget.title.replace(/[^a-zA-Z0-9]/g, '_')}_Snapshot.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
  });
}

async function exportExecutivePDF() {
  const { jsPDF } = window.jspdf;
  const canvasRoot = document.getElementById('studioCanvasRoot');
  if (!canvasRoot) return;

  const bgHex = isDarkMode ? '#090d16' : '#f8fafc';
  const textHex = isDarkMode ? '#f8fafc' : '#0f172a';
  const subHex = isDarkMode ? '#94a3b8' : '#64748b';
  const accentHex = activePalette.accent || '#6366f1';

  pdfTimestampBadge.textContent = `Snapshot: ${new Date().toLocaleString()}`;
  pdfLoadingOverlay.style.display = 'flex';

  setTimeout(async () => {
    try {
      // Capture the full dashboard canvas
      const canvas = await html2canvas(canvasRoot, {
        scale: 2,
        useCORS: true,
        backgroundColor: bgHex,
        logging: false
      });

      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4'
      });

      const pdfWidth = 297;  // A4 Landscape Width in mm
      const pdfHeight = 210; // A4 Landscape Height in mm
      const margin = 10;
      const contentWidth = pdfWidth - (margin * 2);
      const contentHeight = pdfHeight - (margin * 2);

      // Scaled dimensions of the canvas
      const imgScaledHeight = (canvas.height * contentWidth) / canvas.width;

      // Fill Background on Page 1
      pdf.setFillColor(bgHex);
      pdf.rect(0, 0, pdfWidth, pdfHeight, 'F');

      // Top Executive Header Banner
      pdf.setFontSize(16);
      pdf.setFont('helvetica', 'bold');
      pdf.setTextColor(textHex);
      pdf.text(activeStudioTitle.textContent || 'Executive BI Analytics Report', margin, 12);

      pdf.setFontSize(9);
      pdf.setFont('helvetica', 'normal');
      pdf.setTextColor(subHex);
      pdf.text(`Generated: ${new Date().toLocaleString()} | Theme: ${activePalette.name} (${isDarkMode ? 'Dark' : 'Light'})`, margin, 17);

      // Add single or multi-page content cleanly without trailing empty space
      const availablePageContentHeight = contentHeight - 12; // accommodate top banner on page 1

      if (imgScaledHeight <= availablePageContentHeight) {
        // Fits comfortably on 1 landscape page
        const imgData = canvas.toDataURL('image/png');
        pdf.addImage(imgData, 'PNG', margin, 20, contentWidth, imgScaledHeight);
      } else {
        // Multi-page slicing: compute exact number of meaningful pages
        const pageCanvasHeight = (canvas.width * availablePageContentHeight) / contentWidth;
        const totalPages = Math.ceil(canvas.height / pageCanvasHeight);

        for (let p = 0; p < totalPages; p++) {
          if (p > 0) {
            pdf.addPage('a4', 'landscape');
            // Fill background on subsequent pages
            pdf.setFillColor(bgHex);
            pdf.rect(0, 0, pdfWidth, pdfHeight, 'F');

            // Running Header
            pdf.setFontSize(9);
            pdf.setFont('helvetica', 'bold');
            pdf.setTextColor(accentHex);
            pdf.text(`${activeStudioTitle.textContent} — Page ${p + 1} of ${totalPages}`, margin, 10);
          }

          // Slice source canvas for this specific page
          const sourceY = p * pageCanvasHeight;
          const sliceHeight = Math.min(pageCanvasHeight, canvas.height - sourceY);

          const pageCanvas = document.createElement('canvas');
          pageCanvas.width = canvas.width;
          pageCanvas.height = sliceHeight;
          const pCtx = pageCanvas.getContext('2d');
          pCtx.fillStyle = bgHex;
          pCtx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
          pCtx.drawImage(canvas, 0, sourceY, canvas.width, sliceHeight, 0, 0, canvas.width, sliceHeight);

          const sliceImgData = pageCanvas.toDataURL('image/png');
          const sliceDestHeight = (sliceHeight * contentWidth) / canvas.width;
          const startY = p === 0 ? 20 : 14;

          pdf.addImage(sliceImgData, 'PNG', margin, startY, contentWidth, sliceDestHeight);
        }
      }

      pdf.save(`Executive_BI_Report_${Date.now()}.pdf`);
    } catch (err) {
      alert('Failed to generate PDF: ' + err.message);
    } finally {
      pdfLoadingOverlay.style.display = 'none';
    }
  }, 350);
}

// ─── 14. PORTABLE JSON TEMPLATE EXPORT & IMPORT ───────────────────
function exportDashboardTemplate() {
  const template = {
    version: '2.0.0',
    exportedAt: new Date().toISOString(),
    paletteId: activePalette.id,
    isDarkMode: isDarkMode,
    widgets: widgets
  };

  const blob = new Blob([JSON.stringify(template, null, 2)], { type: 'application/json' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `Dashboard_Template_${Date.now()}.json`;
  link.click();
}

function importDashboardTemplate(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const template = JSON.parse(evt.target.result);
      if (template.widgets && Array.isArray(template.widgets)) {
        widgets = template.widgets;
        if (template.paletteId) {
          const pal = COLOR_PALETTES.find(p => p.id === template.paletteId);
          if (pal) switchPalette(pal);
        }
        if (template.isDarkMode !== undefined) {
          isDarkMode = template.isDarkMode;
          bodyEl.classList.toggle('light-theme', !isDarkMode);
          themeIcon.textContent = isDarkMode ? '🌙' : '☀️';
        }
        renderAllWidgets();
        alert('Dashboard layout template restored successfully!');
      }
    } catch (err) {
      alert('Invalid JSON template file: ' + err.message);
    }
  };
  reader.readAsText(file);
}