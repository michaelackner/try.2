class ExcelProcessor {
    constructor() {
        this.workbook = null;
        this.processedData = null;
        this.summaryMetrics = {};
        // API Configuration
        this.apiBaseUrl = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1'
            ? 'http://localhost:8000'
            : window.location.origin;
        // Map state
        this.map = null;
        this.mapLayerGroup = null;
        this.heatLayer = null;
        this.choroplethLayer = null;
        this._worldGeo = null; // cached GeoJSON
        this.mapMode = 'bubbles';
        this._geoLookup = this.buildGeoLookup();
        this._countryLookup = this.buildCountryLookup();
        // Chart state
        this.productChart = null;
        this.productChartType = 'bar';
        this.setupEventListeners();

        // Initialize theme from storage or system
        this.initTheme();
    }

    setupEventListeners() {
        const fileInput = document.getElementById('fileInput');
        const clearFile = document.getElementById('clearFile');
        const processButton = document.getElementById('processButton');
        const downloadButton = document.getElementById('downloadButton');

        fileInput.addEventListener('change', this.handleFileSelect.bind(this));
        clearFile.addEventListener('click', this.clearFile.bind(this));
        processButton.addEventListener('click', this.processFile.bind(this));
        downloadButton.addEventListener('click', this.downloadFile.bind(this));

        // Map mode toggles
        document.addEventListener('click', (e) => {
            const btn = e.target.closest('.map-mode-btn');
            if (!btn) return;
            const mode = btn.getAttribute('data-mode');
            this.setMapMode(mode);
        });

        // Chart mode toggles
        document.addEventListener('click', (e) => {
            const btn = e.target.closest('.chart-mode-btn');
            if (!btn) return;
            const type = btn.getAttribute('data-chart');
            this.setChartMode(type);
        });

        // Theme toggle
        const themeToggle = document.getElementById('themeToggle');
        if (themeToggle) {
            themeToggle.addEventListener('click', () => this.toggleTheme());
        }
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const processButton = document.getElementById('processButton');

        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
        processButton.disabled = false;

        this.hideError();
        this.hideResults();
    }

    clearFile() {
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const processButton = document.getElementById('processButton');

        fileInput.value = '';
        fileInfo.style.display = 'none';
        processButton.disabled = true;

        this.hideError();
        this.hideResults();
    }

    async processFile() {
        this.showLoading();
        this.hideError();
        this.hideResults();

        try {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];

            if (!file) {
                throw new Error('No file selected');
            }

            // Validate file type
            if (!file.name.match(/\.(xlsx|xls)$/i)) {
                throw new Error('Please select a valid Excel file (.xlsx or .xls)');
            }

            // Read the Excel file for dashboard metrics only
            const arrayBuffer = await file.arrayBuffer();
            this.workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

            // Validate file structure
            this.validateFile();

            // Calculate dashboard metrics only
            await new Promise(resolve => setTimeout(resolve, 100));
            await this.calculateDashboardMetrics();

            // Set a flag to enable download
            this.processedData = true; // Just a flag for download button

            // Show results
            this.showResults();

        } catch (error) {
            console.error('Processing error:', error);
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }

    validateFile() {
        const settings = this.getSettings();

        // Check if required sheets exist
        const sheetNames = this.workbook.SheetNames;

        const rawSheet1 = settings.rawSheet1Name || sheetNames[0];
        const rawSheet2 = settings.rawSheet2Name || sheetNames[1];
        const rawSheet3 = settings.rawSheet3Name || sheetNames[2];

        if (!this.workbook.Sheets[rawSheet1]) {
            throw new Error(`Raw Sheet 1 "${rawSheet1}" not found`);
        }
        if (!this.workbook.Sheets[rawSheet2]) {
            throw new Error(`Raw Sheet 2 "${rawSheet2}" not found`);
        }
        if (!this.workbook.Sheets[rawSheet3]) {
            throw new Error(`Raw Sheet 3 "${rawSheet3}" not found`);
        }

        // Validate essential columns in each sheet
        this.validateSheetColumns(rawSheet1, ['B', 'AA', 'M', 'L', 'Q', 'AB', 'AD', 'AL', 'X', 'BZ']);
        this.validateSheetColumns(rawSheet2, ['N', 'AQ', 'AV']);
        this.validateSheetColumns(rawSheet3, ['M', 'BR', 'CN']);

        // Check if Raw Sheet 1 has data rows beyond header
        const sheet1Data = XLSX.utils.sheet_to_json(this.workbook.Sheets[rawSheet1], { header: 1 });
        if (sheet1Data.length < 2) {
            throw new Error('Raw Sheet 1 must have at least one data row');
        }
    }

    validateSheetColumns(sheetName, requiredColumns) {
        const sheet = this.workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sheet['!ref']);

        for (const col of requiredColumns) {
            const colIndex = XLSX.utils.decode_col(col);
            if (colIndex > range.e.c) {
                throw new Error(`Required column ${col} not found in sheet ${sheetName}`);
            }
        }
    }

    async calculateDashboardMetrics() {
        const settings = this.getSettings();

        // Get raw data from Sheet 1 for metrics calculation
        const rawSheet1Name = settings.rawSheet1Name || this.workbook.SheetNames[0];
        const rawSheet1 = this.workbook.Sheets[rawSheet1Name];
        const rawData = XLSX.utils.sheet_to_json(rawSheet1, { header: 1 });

        // Skip first row (titles) and process data
        const dataRows = rawData.slice(1).filter(row => row && row.length > 0);

        // Build raw records for drill-down and date filtering
        this.rawRecords = dataRows.map(row => ({
            deal: this.normalize(row[5] || ''),
            product: this.normalize(row[12] || ''),
            volume: parseFloat(row[16]) || 0,
            date: this.parseDate(row[38]), // AM
            location: this.normalize(row[29] || '')
        })).filter(r => r.product && r.volume);

        // Default range preset
        if (!this.datePreset) this.datePreset = 'ALL';

        // Calculate metrics from current preset
        this.refreshMetricsFromRecords();
    }

    refreshMetricsFromRecords() {
        const records = this.getRecordsByPreset(this.datePreset);
        this.filteredRecords = records;

        let totalDeals = 0;
        let totalVolume = 0;
        const productDistribution = {};
        const uniqueDeals = new Set();

        records.forEach(r => {
            if (r.deal) uniqueDeals.add(r.deal);
            totalVolume += r.volume;
            productDistribution[r.product] = (productDistribution[r.product] || 0) + r.volume;
        });
        totalDeals = uniqueDeals.size;

        this.summaryMetrics = { totalDeals, totalVolume, productDistribution };
    }

    getQuarter(date) {
        return Math.floor(date.getMonth() / 3) + 1;
    }

    getRecordsByPreset(preset) {
        const all = this.rawRecords || [];
        if (!preset || preset === 'ALL') return all;
        const now = new Date();
        const y = now.getFullYear();
        const m = now.getMonth();
        if (preset === 'MTD') {
            return all.filter(r => r.date && r.date.getFullYear() === y && r.date.getMonth() === m);
        }
        if (preset === 'YTD') {
            return all.filter(r => r.date && r.date.getFullYear() === y);
        }
        if (preset === 'QTD') {
            const q = this.getQuarter(now);
            return all.filter(r => r.date && r.date.getFullYear() === y && this.getQuarter(r.date) === q);
        }
        return all;
    }

    // Excel processing methods removed - now handled by Python backend

    // Step 2 enrichment now handled by Python backend

    // Lookup tables now built by Python backend

    // All Excel formatting now handled by Python backend with openpyxl

    // All formatting and calculation methods removed - handled by Python backend

    normalize(value) {
        return value ? String(value).trim().toUpperCase() : '';
    }

    parseDate(value) {
        if (!value) return null;

        // Try parsing as Excel date number
        if (typeof value === 'number') {
            return new Date((value - 25569) * 86400 * 1000);
        }

        // Try parsing as date string
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date;
    }

    formatDateDDMMYYYY(date) {
        if (!date) return '';
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
    }

    getSettings() {
        return {
            outputSheetName: document.getElementById('outputSheetName').value,
            rawSheet1Name: document.getElementById('rawSheet1Name').value,
            rawSheet2Name: document.getElementById('rawSheet2Name').value,
            rawSheet3Name: document.getElementById('rawSheet3Name').value,
            dealColumnName: document.getElementById('dealColumnName').value
        };
    }

    showLoading() {
        document.getElementById('processButton').style.display = 'none';
        document.getElementById('loadingSpinner').style.display = 'flex';
    }

    hideLoading() {
        document.getElementById('processButton').style.display = 'inline-block';
        document.getElementById('loadingSpinner').style.display = 'none';
    }

    showError(message) {
        const errorSection = document.getElementById('errorSection');
        const errorMessage = document.getElementById('errorMessage');
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
    }

    hideError() {
        document.getElementById('errorSection').style.display = 'none';
    }

    showResults() {
        const resultsSection = document.getElementById('resultsSection');

        // Show summary metrics
        this.displayMetrics();

        resultsSection.style.display = 'block';
        // Ensure analytics are visible without manual scrolling
        resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    hideResults() {
        document.getElementById('resultsSection').style.display = 'none';
    }

    displayMetrics() {
        // Update the key metrics cards
        document.getElementById('totalDeals').textContent = this.summaryMetrics.totalDeals || 0;
        document.getElementById('totalVolume').textContent = this.formatNumber(this.summaryMetrics.totalVolume || 0);

        // Generate product distribution chart
        this.displayProductDistribution();
        // Map removed from dashboard
    }

    displayProductDistribution() {
        const canvas = document.getElementById('productChartCanvas');
        if (!canvas || typeof Chart === 'undefined') return;

        const products = this.summaryMetrics.productDistribution || {};
        const sorted = Object.entries(products).sort(([,a],[,b]) => b - a);
        const top = sorted.slice(0, 6);
        const rest = sorted.slice(6);
        const otherSum = rest.reduce((sum, [,v]) => sum + v, 0);
        const grouped = otherSum > 0 ? [...top, ['Other', otherSum]] : top;
        const labels = grouped.map(([k]) => k);
        const data = grouped.map(([,v]) => v);
        if (!data.length) {
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0,0,canvas.width,canvas.height);
            ctx.fillStyle = '#94a3b8';
            ctx.font = '14px Inter, system-ui, sans-serif';
            ctx.fillText('No product data available', 24, 24);
            return;
        }

        // Build palette from CSS variables
        const cs = getComputedStyle(document.documentElement);
        const palette = [
            cs.getPropertyValue('--accent-orange-600').trim() || '#f97316',
            cs.getPropertyValue('--accent-purple-600').trim() || '#8b5cf6',
            cs.getPropertyValue('--accent-green-600').trim() || '#10b981',
            cs.getPropertyValue('--accent-pink-600').trim() || '#f43f5e',
            cs.getPropertyValue('--accent-blue-600').trim() || '#3b82f6',
            '#f59e0b','#22c55e','#ef4444','#06b6d4','#eab308'
        ];
        const gray = (cs.getPropertyValue('--gray-400').trim() || '#9ca3af');
        const bgColors = labels.map((lbl, i) => lbl === 'Other' ? gray : palette[i % palette.length]);

        const datasets = [{
            label: 'Volume (BBL)'
            , data
            , backgroundColor: bgColors
            , borderColor: '#ffffff'
            , borderWidth: this.productChartType === 'doughnut' ? 2 : 0
            , hoverOffset: 8
        }];

        const config = {
            type: this.productChartType,
            data: { labels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                onClick: (evt, els) => {
                    if (els && els.length > 0) {
                        const idx = els[0].index;
                        const label = labels[idx];
                        this.selectedProduct = this.selectedProduct === label ? null : label;
                        this.renderDetailsTable();
                    }
                },
                plugins: {
                    legend: { 
                        display: true, 
                        position: 'right',
                        labels: {
                            usePointStyle: true,
                            generateLabels: (chart) => {
                                const ds = chart.data.datasets[0] || {};
                                const colors = Array.isArray(ds.backgroundColor) ? ds.backgroundColor : [];
                                return chart.data.labels.map((label, i) => ({
                                    text: label,
                                    fillStyle: colors[i] || ds.backgroundColor || '#999',
                                    strokeStyle: '#fff',
                                    lineWidth: 1,
                                    hidden: false,
                                    datasetIndex: 0,
                                    index: i
                                }));
                            }
                        },
                        onClick: () => {}
                    },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => {
                                const val = ctx.parsed;
                                const total = data.reduce((a,b)=>a+b,0);
                                const pct = total ? ((val/total)*100).toFixed(1) : 0;
                                return `${ctx.label}: ${this.formatNumber(val)} BBL (${pct}%)`;
                            }
                        }
                    }
                },
                scales: this.productChartType === 'bar' ? {
                    x: { ticks: { color: cs.getPropertyValue('--gray-700') || '#334155' } },
                    y: { beginAtZero: true, ticks: { color: cs.getPropertyValue('--gray-700') || '#334155' } }
                } : {}
            }
        };

        if (this.productChart) this.productChart.destroy();
        this.productChart = new Chart(canvas.getContext('2d'), config);
        this.renderDetailsTable();
    }

    setChartMode(type) {
        if (!type || (type !== 'bar' && type !== 'doughnut')) return;
        if (this.productChartType === type) return;
        this.productChartType = type;
        document.querySelectorAll('.chart-mode-btn').forEach(btn => {
            const isActive = btn.getAttribute('data-chart') === type;
            btn.classList.toggle('active', isActive);
            btn.setAttribute('aria-selected', String(isActive));
        });
        // Re-render chart with new type
        this.displayProductDistribution();
    }

    renderDetailsTable() {
        const table = document.getElementById('detailsTable');
        if (!table) return;
        const tbody = table.querySelector('tbody');
        tbody.innerHTML = '';

        const records = (this.filteredRecords || []).filter(r => !this.selectedProduct || r.product === this.selectedProduct);
        const rows = records.map(r => `
            <tr>
                <td>${r.deal || ''}</td>
                <td>${r.product || ''}</td>
                <td>${this.formatNumber(r.volume || 0)}</td>
                <td>${r.date ? this.formatDateDDMMYYYY(r.date) : ''}</td>
                <td>${r.location || ''}</td>
            </tr>
        `).join('');
        tbody.innerHTML = rows || '<tr><td colspan="5" style="color: var(--gray-500)">No rows</td></tr>';

        // Sorting handlers
        table.querySelectorAll('th').forEach(th => {
            th.onclick = () => {
                const key = th.getAttribute('data-sort');
                const dir = th.getAttribute('data-dir') === 'asc' ? 'desc' : 'asc';
                th.setAttribute('data-dir', dir);
                const sorted = [...records].sort((a,b) => {
                    let va = a[key]; let vb = b[key];
                    if (key === 'date') { va = a.date ? a.date.getTime() : 0; vb = b.date ? b.date.getTime() : 0; }
                    if (key === 'volume') { va = a.volume; vb = b.volume; }
                    if (typeof va === 'string') va = va.toString();
                    if (typeof vb === 'string') vb = vb.toString();
                    return dir === 'asc' ? (va > vb ? 1 : va < vb ? -1 : 0) : (va < vb ? 1 : va > vb ? -1 : 0);
                });
                this.filteredRecords = sorted; // temporarily render sorted
                this.selectedProduct = this.selectedProduct; // keep filter
                this.renderDetailsTable();
            };
        });

        // Wire actions
        const resetBtn = document.getElementById('resetFilterBtn');
        if (resetBtn) resetBtn.onclick = () => { this.selectedProduct = null; this.renderDetailsTable(); };
        const pngBtn = document.getElementById('exportPngBtn');
        if (pngBtn) pngBtn.onclick = () => this.exportChartPng();
        const csvBtn = document.getElementById('copyCsvBtn');
        if (csvBtn) csvBtn.onclick = () => this.copyChartCsv();

        // Range buttons
        document.querySelectorAll('.range-btn').forEach(btn => {
            btn.onclick = () => {
                const range = btn.getAttribute('data-range');
                this.datePreset = range;
                document.querySelectorAll('.range-btn').forEach(b => {
                    const active = b.getAttribute('data-range') === range;
                    b.classList.toggle('active', active);
                    b.setAttribute('aria-selected', String(active));
                });
                this.refreshMetricsFromRecords();
                this.displayMetrics();
            };
        });
    }

    exportChartPng() {
        if (!this.productChart) return;
        const url = this.productChart.toBase64Image('image/png', 1);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'product_distribution.png';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    copyChartCsv() {
        const products = this.summaryMetrics.productDistribution || {};
        const sorted = Object.entries(products).sort(([,a],[,b]) => b - a);
        const top = sorted.slice(0, 6);
        const rest = sorted.slice(6);
        const otherSum = rest.reduce((sum, [,v]) => sum + v, 0);
        const grouped = otherSum > 0 ? [...top, ['Other', otherSum]] : top;
        const rows = [['Product','Volume']].concat(grouped.map(([k,v]) => [k, v]));
        const csv = rows.map(r => r.map(x => typeof x === 'string' && x.includes(',') ? '"'+x+'"' : x).join(',')).join('\n');
        navigator.clipboard.writeText(csv).then(() => this.toast('Copied CSV to clipboard'));
    }

    toast(msg) {
        const t = document.getElementById('toast');
        if (!t) return;
        t.textContent = msg;
        t.style.display = 'block';
        setTimeout(() => { t.style.display = 'none'; t.textContent=''; }, 1800);
    }

    formatNumber(num) {
        if (num >= 1000000) {
            return (num / 1000000).toFixed(1) + 'M';
        } else if (num >= 1000) {
            return (num / 1000).toFixed(1) + 'K';
        } else {
            return num.toLocaleString();
        }
    }

    // Preview functionality removed - dashboard now focuses on key metrics only

    async downloadFile() {
        if (!this.processedData) return;

        try {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];

            if (!file) {
                throw new Error('No file selected');
            }

            const settings = this.getSettings();

            // Create FormData for the Python backend
            const formData = new FormData();
            formData.append('file', file);
            formData.append('output_sheet_name', settings.outputSheetName || 'Q1-Q2-Q3-Q4-2024');
            formData.append('raw_sheet1_name', settings.rawSheet1Name || '');
            formData.append('raw_sheet2_name', settings.rawSheet2Name || '');
            formData.append('raw_sheet3_name', settings.rawSheet3Name || '');
            formData.append('deal_column_name', settings.dealColumnName || 'N');

            // Send to Python backend
            const response = await fetch(`${this.apiBaseUrl}/process`, {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || 'Server processing failed');
            }

            // Get the processed Excel file
            const blob = await response.blob();

            // Download the file
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'formatted_output.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

        } catch (error) {
            console.error('Download error:', error);
            this.showError(`Download failed: ${error.message}`);
        }
    }

    // xlsx-js-style loader removed - formatting handled by Python backend
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new ExcelProcessor();
});

    // ---- Mapping helpers ----
ExcelProcessor.prototype.initTheme = function() {
    const stored = localStorage.getItem('theme');
    let theme = stored;
    if (!theme) {
        const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
        theme = prefersDark ? 'dark' : 'light';
    }
    this.setTheme(theme);
};

ExcelProcessor.prototype.setTheme = function(theme) {
    const root = document.documentElement;
    if (theme === 'dark') {
        root.setAttribute('data-theme', 'dark');
    } else {
        root.removeAttribute('data-theme');
        theme = 'light';
    }
    localStorage.setItem('theme', theme);
    const btn = document.getElementById('themeToggle');
    if (btn) {
        const icon = btn.querySelector('.theme-icon');
        if (icon) icon.textContent = theme === 'dark' ? '☀️' : '🌙';
        btn.setAttribute('aria-label', theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode');
        btn.title = theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode';
    }
};

ExcelProcessor.prototype.toggleTheme = function() {
    const current = localStorage.getItem('theme') || 'light';
    this.setTheme(current === 'dark' ? 'light' : 'dark');
};

ExcelProcessor.prototype.buildGeoLookup = function() {
    // Minimal built-in geocoder for common ports/cities.
    // Extend this as needed or wire to a backend geocoder.
    return {
        'SINGAPORE': [1.3521, 103.8198],
        'ROTTERDAM': [51.9244, 4.4777],
        'HOUSTON': [29.7604, -95.3698],
        'NEW YORK': [40.7128, -74.0060],
        'ANTWERP': [51.2194, 4.4025],
        'HAMBURG': [53.5511, 9.9937],
        'GENOA': [44.4056, 8.9463],
        'SINES': [37.9561, -8.8697],
        'FUJAIRAH': [25.1288, 56.3265],
        'SUEZ': [29.9668, 32.5498],
        'JEBEL ALI': [25.0108, 55.0617],
        'MARSEILLE': [43.2965, 5.3698],
        'BARCELONA': [41.3851, 2.1734],
        'VALENCIA': [39.4699, -0.3763],
        'ALGECIRAS': [36.1408, -5.4562],
        'TANGIER': [35.7595, -5.8340],
        'PANAMA': [8.9824, -79.5199],
        'SANTOS': [-23.967, -46.328],
        'BUENOS AIRES': [-34.6037, -58.3816],
        'DURBAN': [-29.8587, 31.0218],
        'MUMBAI': [19.0760, 72.8777],
        'SINGAPURA': [1.3521, 103.8198], // common alt
        'PORT SAID': [31.2653, 32.3019]
    };
};

ExcelProcessor.prototype.buildCountryLookup = function() {
    // Map known locations to countries (approx)
    return {
        'SINGAPORE': 'Singapore',
        'SINGAPURA': 'Singapore',
        'ROTTERDAM': 'Netherlands',
        'ANTWERP': 'Belgium',
        'HAMBURG': 'Germany',
        'GENOA': 'Italy',
        'SINES': 'Portugal',
        'FUJAIRAH': 'United Arab Emirates',
        'SUEZ': 'Egypt',
        'JEBEL ALI': 'United Arab Emirates',
        'MARSEILLE': 'France',
        'BARCELONA': 'Spain',
        'VALENCIA': 'Spain',
        'ALGECIRAS': 'Spain',
        'TANGIER': 'Morocco',
        'PANAMA': 'Panama',
        'SANTOS': 'Brazil',
        'BUENOS AIRES': 'Argentina',
        'DURBAN': 'South Africa',
        'MUMBAI': 'India',
        'HOUSTON': 'United States of America',
        'NEW YORK': 'United States of America',
        'PORT SAID': 'Egypt'
    };
};

ExcelProcessor.prototype.displayLocationDensity = function() {
    const container = document.getElementById('locationMap');
    if (!container) return;

    const locations = this.summaryMetrics.locationCounts || {};
    const countries = this.summaryMetrics.countryCounts || {};

    // Lazy init map
    if (!this.map) {
        this.map = L.map('locationMap', {
            zoomControl: true,
            attributionControl: false
        }).setView([20, 0], 2);

        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            maxZoom: 18
        }).addTo(this.map);

        // Attribution (styled minimal)
        L.control.attribution({ prefix: false }).addAttribution('© OpenStreetMap contributors').addTo(this.map);
    }

    // Render by selected map mode
    this.clearMapLayers();
    if (this.mapMode === 'heat') {
        this.renderHeatmap(locations);
    } else if (this.mapMode === 'choropleth') {
        this.renderChoropleth(countries);
    } else {
        this.renderBubbles(locations);
    }
};

ExcelProcessor.prototype._renderMapLegend = function(maxCount) {
    if (this._legendControl) {
        this.map.removeControl(this._legendControl);
    }

    const legend = L.control({ position: 'bottomright' });
    legend.onAdd = () => {
        const div = L.DomUtil.create('div', 'map-legend');
        div.style.background = 'rgba(255,255,255,0.9)';
        div.style.borderRadius = '8px';
        div.style.padding = '8px 10px';
        div.style.boxShadow = '0 2px 8px rgba(0,0,0,0.08)';
        div.innerHTML = `<div style="font-weight:600;margin-bottom:6px">Deal Density</div>
            <div style="display:flex;align-items:center;gap:6px">
                <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#FFDD57"></span>
                <span style="font-size:12px;color:#666">Low</span>
                <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#EA580C;margin-left:12px"></span>
                <span style="font-size:12px;color:#666">High</span>
            </div>`;
        return div;
    };
    legend.addTo(this.map);
    this._legendControl = legend;
};

ExcelProcessor.prototype.setMapMode = function(mode) {
    if (!mode || this.mapMode === mode) return;
    this.mapMode = mode;
    document.querySelectorAll('.map-mode-btn').forEach(btn => {
        const isActive = btn.getAttribute('data-mode') === mode;
        btn.classList.toggle('active', isActive);
        btn.setAttribute('aria-selected', String(isActive));
    });
    // Re-render with new mode if metrics loaded
    if (this.summaryMetrics && (this.summaryMetrics.locationCounts || this.summaryMetrics.countryCounts)) {
        this.displayLocationDensity();
    }
};

ExcelProcessor.prototype.clearMapLayers = function() {
    if (this.mapLayerGroup) {
        this.map.removeLayer(this.mapLayerGroup);
        this.mapLayerGroup = null;
    }
    if (this.heatLayer) {
        this.map.removeLayer(this.heatLayer);
        this.heatLayer = null;
    }
    if (this.choroplethLayer) {
        this.map.removeLayer(this.choroplethLayer);
        this.choroplethLayer = null;
    }
};

ExcelProcessor.prototype.renderBubbles = function(locations) {
    this.mapLayerGroup = L.layerGroup().addTo(this.map);
    const counts = Object.values(locations);
    const maxCount = counts.length ? Math.max(...counts) : 0;
    if (maxCount === 0) return;
    const colorFor = (count) => {
        const t = Math.max(0, Math.min(1, count / maxCount));
        const c1 = [255, 221, 87];
        const c2 = [234, 88, 12];
        const mix = (a,b) => Math.round(a + (b-a)*t);
        const [r,g,b] = [mix(c1[0],c2[0]), mix(c1[1],c2[1]), mix(c1[2],c2[2])];
        return `rgb(${r}, ${g}, ${b})`;
    };
    const latlngs = [];
    Object.entries(locations).forEach(([name, count]) => {
        const coords = this._geoLookup[name];
        if (!coords) return;
        latlngs.push(L.latLng(coords[0], coords[1]));
        const radius = 6 + (28 * (count / maxCount));
        const color = colorFor(count);
        const circle = L.circleMarker(coords, {
            radius,
            color,
            weight: 1,
            opacity: 0.9,
            fillOpacity: 0.35,
            fillColor: color
        });
        circle.bindPopup(`<strong>${name}</strong><br/>Deals: ${count}`);
        circle.addTo(this.mapLayerGroup);
    });
    if (latlngs.length) {
        this.map.fitBounds(L.latLngBounds(latlngs).pad(0.25));
    }
    this._renderMapLegend(maxCount);
};

ExcelProcessor.prototype.renderHeatmap = function(locations) {
    const points = [];
    let maxCount = 0;
    Object.entries(locations).forEach(([name, count]) => {
        const coords = this._geoLookup[name];
        if (!coords) return;
        maxCount = Math.max(maxCount, count);
        points.push([coords[0], coords[1], count]);
    });
    if (!points.length) return;
    this.heatLayer = L.heatLayer(points, {
        radius: 28,
        blur: 22,
        maxZoom: 6,
        max: maxCount,
        gradient: {
            0.2: '#FFDD57',
            0.4: '#fbbf24',
            0.6: '#f97316',
            0.8: '#ef4444',
            1.0: '#dc2626'
        }
    }).addTo(this.map);
};

ExcelProcessor.prototype.getWorldGeo = async function() {
    if (this._worldGeo) return this._worldGeo;
    // Fetch and cache simplified world countries TopoJSON and convert to GeoJSON
    const url = 'https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json';
    const res = await fetch(url);
    const topo = await res.json();
    const geo = topojson.feature(topo, topo.objects.countries);
    this._worldGeo = geo;
    return geo;
};

ExcelProcessor.prototype.renderChoropleth = async function(countryCounts) {
    const geo = await this.getWorldGeo();
    // Compute scale
    const counts = Object.values(countryCounts);
    const maxCount = counts.length ? Math.max(...counts) : 0;
    const minCount = counts.length ? Math.min(...counts) : 0;
    const colorFor = (v) => {
        if (!v || maxCount === 0) return '#f3f4f6';
        const t = (v - minCount) / (maxCount - minCount || 1);
        const c1 = [255, 221, 87];
        const c2 = [234, 88, 12];
        const mix = (a,b) => Math.round(a + (b-a)*t);
        const [r,g,b] = [mix(c1[0],c2[0]), mix(c1[1],c2[1]), mix(c1[2],c2[2])];
        return `rgb(${r}, ${g}, ${b})`;
    };

    // Country name matching helper (simple normalization)
    const norm = (s) => (s || '').toString().trim().toUpperCase();
    const layer = L.geoJSON(geo, {
        style: (feature) => {
            const name = feature.properties.name;
            const v = countryCounts[name] || countryCounts[name + ' (the)'] || countryCounts[name.replace(' and', ' &')];
            return {
                color: '#ffffff',
                weight: 0.8,
                fillColor: colorFor(v),
                fillOpacity: 0.85
            };
        },
        onEachFeature: (feature, lyr) => {
            const name = feature.properties.name;
            const v = countryCounts[name] || 0;
            lyr.bindTooltip(`${name}: ${v} deals`, { sticky: true, direction: 'center', className: 'country-tip' });
            lyr.on({
                mouseover: (e) => e.target.setStyle({ weight: 2, color: '#111827' }),
                mouseout: (e) => layer.resetStyle(e.target)
            });
        }
    });
    this.choroplethLayer = layer.addTo(this.map);
    this.map.fitBounds(layer.getBounds().pad(0.1));
};
