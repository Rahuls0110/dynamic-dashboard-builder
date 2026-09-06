# 📊 Interactive Business Intelligence & Dashboard Studio (dynamic-dashboard-builder)

> **Professional in-browser Business Intelligence & Interactive Dashboard Builder built with Vanilla HTML5, CSS3, ES6 JavaScript, Chart.js 4.4, SheetJS, and jsPDF.**

---

## 🚀 Key Highlights & Capabilities

### 1. 🌟 Master Reference Showcase Dataset (All 10+ Chart Types on Load)
When opened, the dashboard **instantly renders a complete executive operations dashboard** displaying all 10+ visualization types and KPI cards simultaneously:
* 📊 **Vertical Bar Chart**: Revenue by Business Region
* 📊 **Stacked Bar Chart**: Multi-Series Revenue Tier Mix
* 📊 **Horizontal Bar Chart**: Department Headcount Allocation
* 📈 **Line Trend Chart**: Monthly Recurring Revenue (MRR) Growth
* 🌊 **Gradient Area Chart**: Quarterly Expansion vs Churn Velocity
* 🍩 **Doughnut Chart**: Subscription Tier Distribution Share
* 🥧 **Pie Chart**: Asset Fleet Lifecycle Status
* 🕸️ **Radar Chart**: Infrastructure & Capability Matrix
* 🎯 **Polar Area Chart**: Acquisition Channel Intensity
* 📋 **Executive Metric Cards & KPI Tiles**: Delta trend indicators (`+15.4% YoY`, `+9.2% MoM`)

### 2. ⚡ Automated Dashboard Generation on File Upload
* **Drag-and-Drop or File Picker**: Upload any `.xlsx`, `.xls`, or `.csv` file.
* **Instant Dashboard Creation**: Automatically inspects column headers and numeric metrics to **create executive KPIs and interactive charts immediately on upload** with zero setup needed.
* **Custom Widget Builder Bar**: Allows appending extra custom charts, switching axes, or tweaking aggregation methods (Sum, Avg, Count) dynamically.

### 3. 🎯 Global Cross-Filtering & Slicers
* **Dimension Slicers**: Interactive category dropdown that filters and recalculates metrics across all canvas cards in real-time.
* **Time Period Slicers**: All Time, Year to Date (YTD), Last 90 Days, Last 30 Days with automatic metric volume scaling.

### 4. 🔍 Interactive Slice Drill-Down Modal
* Deep-dive into any widget to inspect raw aggregated data slices, category percentages, and proportional distribution bars.

### 5. 🎨 Theme Engine & 5 Curated Palettes
* **Dark & Light Mode Toggle** with glassmorphism panels.
* **5 Curated Color Schemes**:
  * 🟣 **Indigo Glow**: Deep slate, neon indigo, violet
  * 🟢 **Emerald Mint**: Dark teal, mint green, emerald
  * 🔵 **Cyber Neon**: Midnight, electric cyan, hot magenta
  * 🟠 **Sunset Gradient**: Warm twilight, amber, coral rose
  * 🔷 **Executive Corporate**: Navy blue, steel blue, gold accent

### 6. 📄 PDF, PNG & JSON Export Pipelines
* **Executive Multi-Widget PDF Report**: Generates high-resolution multi-page executive PDF reports using `html2canvas` + `jsPDF`.
* **Individual Widget PNG Snapshot**: Download crisp PNG snapshots of any chart or KPI card.
* **JSON Layout Template Save & Load**: Export custom dashboard configurations to portable `.json` files and restore them in one click.

---

## 🛠️ Technology Stack
* **Core**: Pure Vanilla HTML5 & ES6 JavaScript (Zero framework overhead)
* **Styling**: Vanilla CSS3 with CSS Custom Properties, Glassmorphism, and Flexbox/Grid
* **Visualization Engine**: [Chart.js 4.4.0](https://www.chartjs.org/) (Local bundled with CDN fallback)
* **Spreadsheet Ingestion**: [SheetJS (xlsx 0.20.2)](https://sheetjs.com/) (Local bundled with CDN fallback)
* **PDF & Canvas Export**: [jsPDF 2.5.1](https://github.com/parallax/jsPDF) & [html2canvas 1.4.1](https://html2canvas.hertzen.com/)

---

## 🏁 Quick Start Guide

### Running Locally
1. Open `index.html` directly in any web browser or serve it with any local server:
   ```bash
   npx http-server . -p 8085
   ```
2. Navigate to `http://localhost:8085`.
3. **Works 100% offline with zero build steps!**
