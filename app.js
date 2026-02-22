// ==================== STATE ====================
let excelData = [];
let headers = [];
let charts = [];
let widgets = [];
const dashboard = document.getElementById('dashboard');

const fileInput = document.getElementById('fileInput');
const chartType = document.getElementById('chartType');
const titleInput = document.getElementById('titleInput');
const xAxis = document.getElementById('xAxis');
const yAxis = document.getElementById('yAxis');
const dateAxis = document.getElementById('dateAxis');
const aggregation = document.getElementById('aggregation');
const addChartBtn = document.getElementById('addChart');
const downloadAllBtn = document.getElementById('downloadAllPDF');

// ==================== FILE HANDLING ====================
fileInput.addEventListener('change', handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

      if (excelData.length < 2) {
        alert('File must contain at least one header row and one data row.');
        return;
      }

      headers = excelData[0].filter(h => h && typeof h === 'string' && h.trim() !== '');
      excelData = excelData.slice(1).map(row => {
        let obj = {};
        headers.forEach((h, idx) => {
          obj[h] = row[idx];
        });
        return obj;
      }).filter(row => Object.values(row).some(v => v !== undefined && v !== null && v !== ''));

      if (excelData.length === 0) {
        alert('No valid data rows found.');
        return;
      }

      populateAxes();
    } catch (err) {
      alert('Error parsing file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function populateAxes() {
  [xAxis, yAxis, dateAxis].forEach(select => {
    select.innerHTML = '';
    headers.forEach(h => {
      const option = document.createElement('option');
      option.value = h;
      option.textContent = h;
      select.appendChild(option);
    });
  });
  const dateCol = headers.find(h => h.toLowerCase().includes('date')) || headers[0];
  dateAxis.value = dateCol;
}

// ==================== UI BEHAVIOR ====================
chartType.addEventListener('change', () => {
  const type = chartType.value;
  xAxis.disabled = (type === 'card' || type === 'kpi');
  dateAxis.disabled = (type !== 'kpi');
});

function updateDownloadAllButton() {
  downloadAllBtn.disabled = widgets.length < 2;
}

// ==================== DATA VALIDATION HELPERS ====================
function isNumeric(value) {
  return !isNaN(parseFloat(value)) && isFinite(value);
}

function getNumericValues(arr) {
  return arr.filter(v => isNumeric(v)).map(Number);
}

// ==================== AGGREGATION ====================
function aggregate(values, type) {
  if (type === 'count') return values.length;
  const nums = getNumericValues(values);
  if (nums.length === 0) return 0;
  if (type === 'sum') return nums.reduce((a, b) => a + b, 0);
  if (type === 'avg') return nums.reduce((a, b) => a + b, 0) / nums.length;
  return 0;
}

// ==================== CLEAN / AUTO TITLE ====================
function cleanName(name) {
  return name.replace(/id$/i, '').trim();
}

// ==================== KPI CALCULATION ====================
function calculateKPI(data, dateKey, valueKey, agg) {
  const sorted = [...data]
    .filter(row => {
      const dateVal = row[dateKey];
      return dateVal && !isNaN(new Date(dateVal).getTime());
    })
    .sort((a, b) => new Date(a[dateKey]) - new Date(b[dateKey]));

  if (sorted.length < 2) return { currVal: 0, change: 0 };

  const mid = Math.floor(sorted.length / 2);
  const prevPeriod = sorted.slice(0, mid);
  const currPeriod = sorted.slice(mid);

  const getVal = arr => aggregate(arr.map(r => r[valueKey]), agg);
  const prevVal = getVal(prevPeriod);
  const currVal = getVal(currPeriod);
  const change = prevVal ? ((currVal - prevVal) / prevVal * 100).toFixed(1) : 0;

  return { currVal, change };
}

// ==================== ANALYSIS TEXT FOR PDF ====================
function getAnalysisText(widget) {
  if (widget.type === 'card') {
    return `Total: ${widget.value}`;
  } else if (widget.type === 'kpi') {
    return `Current: ${widget.value} (${widget.change}% vs previous period)`;
  } else {
    const labels = widget.labels;
    const data = widget.data;
    if (!labels || !data || labels.length === 0) return 'No data';
    const maxIdx = data.indexOf(Math.max(...data));
    const minIdx = data.indexOf(Math.min(...data));
    const maxLabel = labels[maxIdx];
    const minLabel = labels[minIdx];
    const maxVal = data[maxIdx].toFixed(1);
    const minVal = data[minIdx].toFixed(1);
    return `Highest: ${maxLabel} (${maxVal}), Lowest: ${minLabel} (${minVal})`;
  }
}

// ==================== DOWNLOAD INDIVIDUAL WIDGET PDF ====================
async function downloadWidgetAsPDF(widgetElement, filename, widgetMeta) {
  const { jsPDF } = window.jspdf;
  let imgData;

  const canvas = widgetElement.querySelector('canvas');
  if (canvas) {
    // For charts, use canvas directly (buttons not in canvas)
    imgData = canvas.toDataURL('image/png');
  } else {
    // For card/KPI, hide buttons before screenshot
    const buttons = widgetElement.querySelector('.card-buttons');
    if (buttons) buttons.style.visibility = 'hidden';

    await html2canvas(widgetElement, { scale: 2, backgroundColor: '#ffffff' })
      .then(canvas2 => {
        imgData = canvas2.toDataURL('image/png');
      });

    if (buttons) buttons.style.visibility = 'visible';
  }

  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  pdf.setFontSize(16);
  pdf.text(widgetMeta.title, 20, 20);
  pdf.setFontSize(11);
  const analysis = getAnalysisText(widgetMeta);
  pdf.text(analysis, 20, 30);

  if (imgData) {
    const imgProps = pdf.getImageProperties(imgData);
    const imgWidth = imgProps.width;
    const imgHeight = imgProps.height;
    const ratio = Math.min((pageWidth - 40) / imgWidth, (pageHeight - 50) / imgHeight);
    const w = imgWidth * ratio;
    const h = imgHeight * ratio;
    const x = (pageWidth - w) / 2;
    const y = 40;
    pdf.addImage(imgData, 'PNG', x, y, w, h);
  }

  pdf.save(filename);
}

// ==================== RESET CONTROLS ====================
function resetControls() {
  chartType.value = 'bar';
  chartType.dispatchEvent(new Event('change'));
  titleInput.value = '';
  if (headers.length > 0) {
    if (xAxis.options.length > 0) xAxis.selectedIndex = 0;
    if (yAxis.options.length > 0) yAxis.selectedIndex = 0;
    if (dateAxis.options.length > 0) {
      const dateCol = headers.find(h => h.toLowerCase().includes('date')) || headers[0];
      dateAxis.value = dateCol;
    }
  }
  aggregation.value = 'sum';
}

// ==================== ADD CHART ====================
addChartBtn.addEventListener('click', addChart);

function addChart() {
  if (!excelData.length) {
    alert('Please upload an Excel file first.');
    return;
  }

  const type = chartType.value;
  const x = xAxis.value;
  const y = yAxis.value;
  const date = dateAxis.value;
  const agg = aggregation.value;

  // Generate title
  let title = titleInput.value.trim();
  if (!title) {
    const metric = cleanName(y);
    if (type === 'card' || type === 'kpi') {
      // For card/KPI, omit the "by X" part
      if (agg === 'sum') title = `Total ${metric}`;
      else if (agg === 'avg') title = `Average ${metric}`;
      else title = `Count of ${metric}`;
    } else {
      // For charts, include the X axis
      const dim = cleanName(x);
      if (agg === 'sum') title = `Total ${metric} by ${dim}`;
      else if (agg === 'avg') title = `Average ${metric} by ${dim}`;
      else title = `Count of ${metric} by ${dim}`;
    }
  }

  if (!headers.includes(x) || !headers.includes(y) || !headers.includes(date)) {
    alert('Selected axes are not valid headers.');
    return;
  }

  const wrapper = document.createElement('div');
  wrapper.className = `chart-card ${type === 'card' ? 'card' : type === 'kpi' ? 'kpi' : 'chart'}`;

  // Buttons container
  const btnContainer = document.createElement('div');
  btnContainer.className = 'card-buttons';

  // Remove button
  const removeBtn = document.createElement('span');
  removeBtn.textContent = '✕';
  removeBtn.className = 'remove-btn';
  removeBtn.title = 'Remove';
  removeBtn.onclick = () => {
    const idx = widgets.findIndex(w => w.element === wrapper);
    if (idx !== -1) {
      if (widgets[idx].chartInstance) {
        widgets[idx].chartInstance.destroy();
      }
      widgets.splice(idx, 1);
    }
    wrapper.remove();
    updateDownloadAllButton();
  };
  btnContainer.appendChild(removeBtn);

  // Download button for ALL widget types
  const downloadBtn = document.createElement('span');
  downloadBtn.textContent = '⬇️ PDF';
  downloadBtn.className = 'download-btn';
  downloadBtn.title = 'Download as PDF';
  downloadBtn.onclick = async (e) => {
    e.stopPropagation();
    const meta = widgets.find(w => w.element === wrapper);
    if (meta) {
      // Create filename: dashboard_{type}_{title}.pdf
      const sanitizedTitle = title.replace(/\s+/g, '_');
      const filename = `dashboard_${meta.type}_${sanitizedTitle}.pdf`;
      await downloadWidgetAsPDF(wrapper, filename, meta);
    }
  };
  btnContainer.appendChild(downloadBtn);

  wrapper.appendChild(btnContainer);

  const titleEl = document.createElement('div');
  titleEl.className = 'chart-title';
  titleEl.textContent = title;
  wrapper.appendChild(titleEl);

  let widgetMeta = { type, title, element: wrapper };

  // ===== CARD TYPE =====
  if (type === 'card') {
    const values = excelData.map(r => r[y]);
    const numericValues = getNumericValues(values);
    const total = numericValues.length > 0 ? aggregate(numericValues, agg) : 0;
    const valueDiv = document.createElement('div');
    valueDiv.className = 'card-value';
    valueDiv.textContent = Math.round(total).toLocaleString();
    wrapper.appendChild(valueDiv);

    if (numericValues.length === 0) {
      const warning = document.createElement('div');
      warning.className = 'data-warning';
      warning.textContent = '⚠️ No numeric data';
      wrapper.appendChild(warning);
    }

    dashboard.appendChild(wrapper);

    widgetMeta.value = Math.round(total).toLocaleString();
    widgets.push(widgetMeta);
    updateDownloadAllButton();
    resetControls();
    return;
  }

  // ===== KPI TYPE =====
  if (type === 'kpi') {
    const { currVal, change } = calculateKPI(excelData, date, y, agg);
    const valueDiv = document.createElement('div');
    valueDiv.className = 'card-value';
    valueDiv.textContent = Math.round(currVal).toLocaleString();
    const deltaDiv = document.createElement('div');
    deltaDiv.className = 'kpi-delta';
    const sign = change >= 0 ? '▲' : '▼';
    deltaDiv.textContent = `${sign} ${Math.abs(change)}% vs previous period`;
    deltaDiv.style.color = change >= 0 ? '#16a34a' : '#dc2626';
    wrapper.appendChild(valueDiv);
    wrapper.appendChild(deltaDiv);

    const dateRows = excelData.filter(row => row[date] && !isNaN(new Date(row[date]).getTime()));
    if (dateRows.length < 2) {
      const warning = document.createElement('div');
      warning.className = 'data-warning';
      warning.textContent = '⚠️ Not enough date rows for comparison';
      wrapper.appendChild(warning);
    }

    dashboard.appendChild(wrapper);

    widgetMeta.value = Math.round(currVal).toLocaleString();
    widgetMeta.change = change;
    widgets.push(widgetMeta);
    updateDownloadAllButton();
    resetControls();
    return;
  }

  // ===== CHART TYPES =====
  const grouped = {};
  excelData.forEach(row => {
    const key = row[x];
    if (key === undefined || key === null) return;
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(row[y]);
  });

  const labels = Object.keys(grouped);
  const data = labels.map(label => {
    const nums = getNumericValues(grouped[label]);
    return aggregate(nums, agg);
  });

  const hasAnyNumeric = data.some(v => v !== 0);

  const canvas = document.createElement('canvas');
  wrapper.appendChild(canvas);
  dashboard.appendChild(wrapper);

  // Create chart with beginAtZero for y-axis
  const chart = new Chart(canvas, {
    type,
    data: {
      labels,
      datasets: [{
        label: y,
        data,
        backgroundColor: type === 'pie' 
          ? ['#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6']
          : '#3b82f6',
        borderColor: '#2563ed',
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        title: {
          display: true,
          text: title,
          font: { size: 16, weight: 'bold' },
          padding: { top: 10, bottom: 20 }
        },
        legend: {
          display: type === 'pie'
        }
      },
      scales: type !== 'pie' ? {
        x: { title: { display: true, text: x } },
        y: {
          beginAtZero: true,
          title: {
            display: true,
            text: agg === 'count' ? 'Count' : agg === 'sum' ? 'Sum' : 'Average'
          }
        }
      } : {}
    }
  });

  if (!hasAnyNumeric) {
    const warning = document.createElement('div');
    warning.className = 'data-warning';
    warning.textContent = '⚠️ No numeric data for selected Y axis';
    wrapper.appendChild(warning);
  }

  widgetMeta.chartInstance = chart;
  widgetMeta.labels = labels;
  widgetMeta.data = data;
  widgetMeta.x = x;
  widgetMeta.y = y;
  widgetMeta.agg = agg;
  widgets.push(widgetMeta);
  updateDownloadAllButton();
  resetControls();
}

// ==================== DOWNLOAD ALL PDF ====================
downloadAllBtn.addEventListener('click', downloadAllChartsPDF);

async function downloadAllChartsPDF() {
  if (widgets.length < 2) {
    alert('Please add at least two widgets before downloading.');
    return;
  }

  const { jsPDF } = window.jspdf;
  let pdf;

  for (let i = 0; i < widgets.length; i++) {
    const widget = widgets[i];
    const element = widget.element;

    let imgData;
    const canvas = element.querySelector('canvas');
    if (canvas) {
      imgData = canvas.toDataURL('image/png');
    } else {
      // For card/KPI, hide buttons before screenshot
      const buttons = element.querySelector('.card-buttons');
      if (buttons) buttons.style.visibility = 'hidden';

      await html2canvas(element, { scale: 2, backgroundColor: '#ffffff' })
        .then(canvas2 => {
          imgData = canvas2.toDataURL('image/png');
        });

      if (buttons) buttons.style.visibility = 'visible';
    }

    const img = new Image();
    img.src = imgData;
    await img.decode();

    const isLandscape = img.width > img.height;

    if (i === 0) {
      pdf = new jsPDF({
        orientation: isLandscape ? 'landscape' : 'portrait',
        unit: 'mm',
        format: 'a4'
      });
    } else {
      pdf.addPage(isLandscape ? 'landscape' : 'portrait');
    }

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    pdf.setFontSize(16);
    pdf.text(widget.title, 20, 20);
    pdf.setFontSize(11);
    const analysis = getAnalysisText(widget);
    pdf.text(analysis, 20, 30);

    const imgWidth = img.width;
    const imgHeight = img.height;
    const ratio = Math.min((pageWidth - 40) / imgWidth, (pageHeight - 50) / imgHeight);
    const w = imgWidth * ratio;
    const h = imgHeight * ratio;
    const x = (pageWidth - w) / 2;
    const y = 40;

    pdf.addImage(imgData, 'PNG', x, y, w, h);
  }

  pdf.save('dynamic_dashboard.pdf');
}