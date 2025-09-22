class ExcelProcessor {
    constructor() {
        this.workbook = null;
        this.processedData = null;
        this.summaryMetrics = {};
        // API Configuration
        this.apiBaseUrl = this.getDefaultApiBaseUrl();
        this.apiBaseUrlReady = this.initializeApiBaseUrl();
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

    }

    getDefaultApiBaseUrl() {
        try {
            const { protocol, hostname, port, origin } = window.location;

            if (protocol === 'file:' || !hostname) {
                // Opening the dashboard directly from the filesystem.
                return 'http://127.0.0.1:8000';
            }

            const normalizedProtocol = protocol === 'https:' ? 'https:' : 'http:';

            if (['localhost', '127.0.0.1', '::1'].includes(hostname)) {
                return `${normalizedProtocol}//${hostname}:8000`;
            }

            if (port && !['80', '443', '8000', ''].includes(port)) {
                return `${normalizedProtocol}//${hostname}:8000`;
            }

            if (origin && origin !== 'null') {
                return origin;
            }

            return `${normalizedProtocol}//${hostname}${port ? `:${port}` : ''}`;
        } catch (error) {
            console.warn('Failed to determine default API base URL', error);
            return 'http://127.0.0.1:8000';
        }
    }

    buildApiBaseCandidates() {
        const candidates = [];

        const pushCandidate = (value) => {
            if (!value || typeof value !== 'string') return;
            const trimmed = value.trim();
            if (trimmed) candidates.push(trimmed);
        };

        // Manual overrides (global, query parameter, meta tag)
        if (typeof window !== 'undefined' && window.API_BASE_URL) {
            pushCandidate(window.API_BASE_URL);
        }

        try {
            const params = new URLSearchParams(window.location.search || '');
            pushCandidate(params.get('api'));
            pushCandidate(params.get('apiBase'));
            pushCandidate(params.get('api_base'));
        } catch (error) {
            console.warn('Unable to parse query parameters for API base override', error);
        }

        const metaTag = document.querySelector('meta[name="api-base-url"]');
        if (metaTag && metaTag.content) {
            pushCandidate(metaTag.content);
        }

        pushCandidate(this.getDefaultApiBaseUrl());

        try {
            const { protocol, hostname, port, origin } = window.location;
            const normalizedProtocol = protocol === 'https:' ? 'https:' : 'http:';

            if (origin && origin !== 'null') {
                pushCandidate(origin);
            }

            if (!hostname || protocol === 'file:') {
                pushCandidate('http://127.0.0.1:8000');
                pushCandidate('http://localhost:8000');
            } else {
                if (port && !['', '80', '443', '8000'].includes(port)) {
                    pushCandidate(`${normalizedProtocol}//${hostname}:8000`);
                }

                if (!['localhost', '127.0.0.1', '::1'].includes(hostname)) {
                    pushCandidate(`${normalizedProtocol}//localhost:8000`);
                    pushCandidate(`${normalizedProtocol}//127.0.0.1:8000`);
                }
            }
        } catch (error) {
            console.warn('Failed to gather API base candidates', error);
        }

        // Remove duplicates while preserving order
        return [...new Set(candidates.map(c => c.replace(/\/+$/, '')))]
            .filter(Boolean);
    }

    async detectApiBaseUrl() {
        const candidates = this.buildApiBaseCandidates();

        for (const candidate of candidates) {
            const base = candidate.replace(/\/+$/, '');
            try {
                const response = await fetch(`${base}/health`, {
                    method: 'GET',
                    mode: 'cors',
                    cache: 'no-store'
                });

                if (response.ok) {
                    return base;
                }
            } catch (error) {
                console.warn('API base candidate unreachable', base, error);
            }
        }

        if (candidates.length > 0) {
            return candidates[0];
        }

        throw new Error('No API base URL candidates available');
    }

    async initializeApiBaseUrl() {
        try {
            const base = await this.detectApiBaseUrl();
            this.apiBaseUrl = base;
            return base;
        } catch (error) {
            console.error('Unable to detect API base URL, using fallback', error);
            this.apiBaseUrl = this.getDefaultApiBaseUrl();
            return this.apiBaseUrl;
        }
    }

    async ensureApiBaseUrl() {
        if (this.apiBaseUrl) {
            return this.apiBaseUrl;
        }

        if (this.apiBaseUrlReady) {
            const resolved = await this.apiBaseUrlReady;
            if (resolved) {
                this.apiBaseUrl = resolved;
                return resolved;
            }
        }

        this.apiBaseUrl = this.getDefaultApiBaseUrl();
        return this.apiBaseUrl;
    }

    setupEventListeners() {
        const rawFileInput = document.getElementById('rawFileInput');
        const formattedFileInput = document.getElementById('formattedFileInput');
        const clearRawFile = document.getElementById('clearRawFile');
        const clearFormattedFile = document.getElementById('clearFormattedFile');
        const processButton = document.getElementById('processButton');
        const downloadButton = document.getElementById('downloadButton');

        rawFileInput.addEventListener('change', (event) => this.handleFileSelect(event, 'raw'));
        formattedFileInput.addEventListener('change', (event) => this.handleFileSelect(event, 'formatted'));
        clearRawFile.addEventListener('click', () => this.clearFile('raw'));
        clearFormattedFile.addEventListener('click', () => this.clearFile('formatted'));
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

    }

    handleFileSelect(event, type) {
        const file = event.target.files[0];
        const isRaw = type === 'raw';

        if (!file) {
            this.clearFile(type);
            return;
        }

        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            this.showError('Please select a valid Excel file (.xlsx or .xls)');
            event.target.value = '';
            this.clearFile(type);
            return;
        }

        const fileInfo = document.getElementById(isRaw ? 'rawFileInfo' : 'formattedFileInfo');
        const fileName = document.getElementById(isRaw ? 'rawFileName' : 'formattedFileName');

        if (fileName) fileName.textContent = file.name;
        if (fileInfo) fileInfo.style.display = 'flex';

        if (isRaw) {
            const processButton = document.getElementById('processButton');
            processButton.disabled = false;
            this.hideError();
            this.hideResults();
        }
    }

    clearFile(type) {
        const isRaw = type === 'raw';
        const fileInput = document.getElementById(isRaw ? 'rawFileInput' : 'formattedFileInput');
        const fileInfo = document.getElementById(isRaw ? 'rawFileInfo' : 'formattedFileInfo');

        if (fileInput) fileInput.value = '';
        if (fileInfo) fileInfo.style.display = 'none';

        if (isRaw) {
            const processButton = document.getElementById('processButton');
            processButton.disabled = true;
            this.processedData = null;
            this.hideError();
            this.hideResults();
        }
    }

    async processFile() {
        this.showLoading();
        this.hideError();
        this.hideResults();

        try {
            const fileInput = document.getElementById('rawFileInput');
            const file = fileInput.files[0];

            if (!file) {
                throw new Error('No file selected');
            }

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
            const rawFileInput = document.getElementById('rawFileInput');
            const formattedFileInput = document.getElementById('formattedFileInput');
            const rawFile = rawFileInput.files[0];
            const formattedFile = formattedFileInput.files[0];

            if (!rawFile) {
                throw new Error('No raw data file selected');
            }

            const settings = this.getSettings();

            // Create FormData for the Python backend
            const formData = new FormData();
            formData.append('file', rawFile);
            if (formattedFile) {
                formData.append('existing_file', formattedFile);
            }
            formData.append('output_sheet_name', settings.outputSheetName || 'Q1-Q2-Q3-Q4-2024');
            formData.append('raw_sheet1_name', settings.rawSheet1Name || '');
            formData.append('raw_sheet2_name', settings.rawSheet2Name || '');
            formData.append('raw_sheet3_name', settings.rawSheet3Name || '');
            formData.append('deal_column_name', settings.dealColumnName || 'N');

            const apiBaseUrl = await this.ensureApiBaseUrl();
            if (!apiBaseUrl) {
                throw new Error('API backend is not reachable. Please ensure the server is running.');
            }

            // Send to Python backend
            const response = await fetch(`${apiBaseUrl}/process`, {
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

// Deal comparison dashboard module
class DealComparisonDashboard {
    constructor(processor) {
        this.processor = processor;
        this.apiBaseUrl = null;
        this.apiReady = this.initializeApi();

        this.formattedFile = null;
        this.referenceFile = null;
        this.analysisToken = null;
        this.analysisData = null;
        this.charts = {
            variance: null,
            waterfall: null,
            heatmap: null,
            treemap: null
        };
        this.filters = {
            deals: new Set(),
            costTypes: new Set()
        };
        this.defaultPlotlyConfig = {
            responsive: true,
            displaylogo: false,
            displayModeBar: true,
            scrollZoom: true,
            modeBarButtonsToRemove: ['toImage']
        };

        this.elements = this.cacheElements();
        this.bindEvents();
    }

    async initializeApi() {
        try {
            this.apiBaseUrl = await this.processor.ensureApiBaseUrl();
        } catch (error) {
            console.warn('Comparison module using fallback API base URL', error);
            this.apiBaseUrl = this.processor.apiBaseUrl || 'http://127.0.0.1:8000';
        }
        return this.apiBaseUrl;
    }

    cacheElements() {
        return {
            module: document.getElementById('comparisonModule'),
            formattedInput: document.getElementById('comparisonFormattedInput'),
            formattedInfo: document.getElementById('comparisonFormattedInfo'),
            formattedName: document.getElementById('comparisonFormattedName'),
            clearFormatted: document.getElementById('clearComparisonFormatted'),
            referenceInput: document.getElementById('comparisonReferenceInput'),
            referenceInfo: document.getElementById('comparisonReferenceInfo'),
            referenceName: document.getElementById('comparisonReferenceName'),
            clearReference: document.getElementById('clearComparisonReference'),
            compareButton: document.getElementById('compareButton'),
            loading: document.getElementById('comparisonLoading'),
            errorCard: document.getElementById('comparisonError'),
            errorMessage: document.getElementById('comparisonErrorMessage'),
            formattedSheet: document.getElementById('comparisonFormattedSheet'),
            referenceSheet: document.getElementById('comparisonReferenceSheet'),
            quantityColumn: document.getElementById('comparisonQuantityColumn'),
            formattedLetter: document.getElementById('comparisonFormattedLetter'),
            results: document.getElementById('comparisonResults'),
            kpiDeals: document.getElementById('kpiDeals'),
            kpiDifference: document.getElementById('kpiDifference'),
            kpiAverageVariance: document.getElementById('kpiAverageVariance'),
            kpiUnregistered: document.getElementById('kpiUnregistered'),
            filters: document.getElementById('comparisonFilters'),
            headline: document.getElementById('comparisonHeadline'),
            topDeals: document.getElementById('comparisonTopDeals'),
            costSummary: document.getElementById('comparisonCostSummary'),
            recommendations: document.getElementById('comparisonRecommendations'),
            anomalies: document.getElementById('comparisonAnomalies'),
            patterns: document.getElementById('comparisonPatterns'),
            exportButtons: document.querySelectorAll('.export-actions [data-export]'),
            varianceChart: document.getElementById('dealVarianceChart'),
            waterfallChart: document.getElementById('costWaterfallChart'),
            treemapChart: document.getElementById('unregisteredTreemap'),
            heatmapChart: document.getElementById('costHeatmap')
        };
    }

    bindEvents() {
        const el = this.elements;
        if (!el.module) {
            return;
        }

        const handleFile = (input, type) => {
            const file = input.files && input.files[0] ? input.files[0] : null;
            if (type === 'formatted') {
                this.formattedFile = file;
                if (file && el.formattedName) {
                    el.formattedName.textContent = `${file.name} (${this.formatSize(file.size)})`;
                }
                if (el.formattedInfo) {
                    el.formattedInfo.style.display = file ? 'block' : 'none';
                }
            } else {
                this.referenceFile = file;
                if (file && el.referenceName) {
                    el.referenceName.textContent = `${file.name} (${this.formatSize(file.size)})`;
                }
                if (el.referenceInfo) {
                    el.referenceInfo.style.display = file ? 'block' : 'none';
                }
            }
            this.updateActionState();
        };

        if (el.formattedInput) {
            el.formattedInput.addEventListener('change', () => handleFile(el.formattedInput, 'formatted'));
        }
        if (el.referenceInput) {
            el.referenceInput.addEventListener('change', () => handleFile(el.referenceInput, 'reference'));
        }
        if (el.clearFormatted) {
            el.clearFormatted.addEventListener('click', () => {
                if (el.formattedInput) {
                    el.formattedInput.value = '';
                }
                this.formattedFile = null;
                if (el.formattedInfo) {
                    el.formattedInfo.style.display = 'none';
                }
                this.updateActionState();
            });
        }
        if (el.clearReference) {
            el.clearReference.addEventListener('click', () => {
                if (el.referenceInput) {
                    el.referenceInput.value = '';
                }
                this.referenceFile = null;
                if (el.referenceInfo) {
                    el.referenceInfo.style.display = 'none';
                }
                this.updateActionState();
            });
        }
        if (el.compareButton) {
            el.compareButton.addEventListener('click', () => this.analyze());
        }

        if (el.exportButtons && el.exportButtons.length) {
            el.exportButtons.forEach((button) => {
                button.addEventListener('click', () => {
                    const type = button.getAttribute('data-export');
                    this.downloadExport(type);
                });
            });
        }
    }

    updateActionState() {
        if (!this.elements.compareButton) return;
        const ready = Boolean(this.formattedFile && this.referenceFile);
        this.elements.compareButton.disabled = !ready;
    }

    setLoading(isLoading) {
        if (!this.elements.compareButton || !this.elements.loading) return;
        this.elements.compareButton.disabled = isLoading || !this.formattedFile || !this.referenceFile;
        this.elements.loading.style.display = isLoading ? 'flex' : 'none';
    }

    clearError() {
        if (this.elements.errorCard) {
            this.elements.errorCard.style.display = 'none';
        }
    }

    showError(message) {
        if (this.elements.errorCard && this.elements.errorMessage) {
            this.elements.errorMessage.textContent = message;
            this.elements.errorCard.style.display = 'block';
        }
    }

    async analyze() {
        if (!this.formattedFile || !this.referenceFile) {
            return;
        }
        await this.apiReady;
        this.clearError();
        this.setLoading(true);

        const formData = new FormData();
        formData.append('formatted_file', this.formattedFile);
        formData.append('comparison_file', this.referenceFile);
        if (this.elements.formattedSheet && this.elements.formattedSheet.value) {
            formData.append('formatted_sheet', this.elements.formattedSheet.value);
        }
        if (this.elements.referenceSheet && this.elements.referenceSheet.value) {
            formData.append('comparison_sheet', this.elements.referenceSheet.value);
        }
        if (this.elements.quantityColumn && this.elements.quantityColumn.value) {
            formData.append('comparison_quantity_column', this.elements.quantityColumn.value);
        }
        if (this.elements.formattedLetter && this.elements.formattedLetter.value) {
            formData.append('formatted_quantity_letter', this.elements.formattedLetter.value);
        }

        try {
            const response = await fetch(`${this.apiBaseUrl}/compare-deals`, {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const error = await response.json().catch(() => ({ error: 'Comparison failed' }));
                throw new Error(error.error || 'Comparison failed');
            }

            const data = await response.json();
            this.analysisData = data;
            this.analysisToken = data.token || null;
            this.filters.deals.clear();
            this.filters.costTypes.clear();
            this.renderResults();
        } catch (error) {
            console.error('Comparison analysis failed', error);
            this.showError(error.message || 'Comparison analysis failed');
        } finally {
            this.setLoading(false);
        }
    }

    renderResults() {
        if (!this.analysisData || !this.elements.results) {
            return;
        }
        this.elements.results.style.display = 'flex';
        this.renderKpis();
        this.updateFiltersUI();
        this.renderInsights();
        this.renderCharts();
    }

    renderKpis() {
        const overview = this.analysisData?.overview || {};
        if (this.elements.kpiDeals) {
            this.elements.kpiDeals.textContent = this.formatNumber(overview.total_deals || 0);
        }
        if (this.elements.kpiDifference) {
            this.elements.kpiDifference.textContent = this.formatCurrency(overview.total_difference || 0);
        }
        if (this.elements.kpiAverageVariance) {
            const value = typeof overview.average_variance === 'number' ? overview.average_variance : 0;
            this.elements.kpiAverageVariance.textContent = `${value.toFixed(2)}%`;
        }
        if (this.elements.kpiUnregistered) {
            this.elements.kpiUnregistered.textContent = this.formatNumber(overview.unregistered_cost_types || 0);
        }
    }

    renderInsights() {
        const summary = this.analysisData?.summary_report;
        if (!summary) return;
        if (this.elements.headline) {
            this.elements.headline.textContent = summary.headline || '';
        }
        if (this.elements.topDeals) {
            this.elements.topDeals.textContent = summary.top_contributors || '';
        }
        if (this.elements.costSummary) {
            this.elements.costSummary.textContent = summary.unregistered_costs || '';
        }

        if (this.elements.recommendations) {
            this.elements.recommendations.innerHTML = '';
            (summary.recommended_actions || []).forEach((action) => {
                const li = document.createElement('li');
                li.textContent = action;
                this.elements.recommendations.appendChild(li);
            });
        }

        if (this.elements.anomalies) {
            this.elements.anomalies.innerHTML = '';
            const anomalies = this.analysisData?.anomalies || [];
            if (!anomalies.length) {
                const li = document.createElement('li');
                li.textContent = 'No anomaly-level deal variances detected.';
                this.elements.anomalies.appendChild(li);
            } else {
                anomalies.forEach((item) => {
                    const li = document.createElement('li');
                    li.textContent = `${item.deal_id}: difference ${this.formatCurrency(item.difference)} (${this.formatCurrency(item.comparison_quantity)} → ${this.formatCurrency(item.formatted_quantity)})`;
                    li.dataset.deal = item.deal_id;
                    li.addEventListener('click', () => this.toggleDealFilter(item.deal_id));
                    this.elements.anomalies.appendChild(li);
                });
            }
        }

        if (this.elements.patterns) {
            this.elements.patterns.innerHTML = '';
            const patterns = this.analysisData?.patterns || {};
            const statusCounts = patterns.status_counts || {};
            const statusSummary = `Registered: ${this.formatNumber(statusCounts.Registered || 0)}, Partial: ${this.formatNumber(statusCounts.Partial || 0)}, Unregistered: ${this.formatNumber(statusCounts.Unregistered || 0)}`;
            const statusLi = document.createElement('li');
            statusLi.textContent = `Deal registry status mix — ${statusSummary}`;
            this.elements.patterns.appendChild(statusLi);

            (patterns.repeating_patterns || []).forEach((pattern) => {
                const li = document.createElement('li');
                li.textContent = `${pattern.cost_types.join(', ')} missing for deals ${pattern.deals.join(', ')}`;
                li.dataset.cost = pattern.cost_types[0];
                li.addEventListener('click', () => {
                    pattern.cost_types.forEach((cost) => this.toggleCostFilter(cost));
                });
                this.elements.patterns.appendChild(li);
            });
        }
    }

    renderCharts() {
        const filtered = this.getFilteredDeals();
        this.renderVarianceChart(filtered);
        this.renderWaterfallChart(filtered);
        this.renderTreemap(filtered);
        this.renderHeatmap();
    }

    getPlotlyConfig(overrides = {}) {
        const base = { ...this.defaultPlotlyConfig };
        if (overrides.modeBarButtonsToRemove) {
            const merged = new Set([
                ...(this.defaultPlotlyConfig.modeBarButtonsToRemove || []),
                ...overrides.modeBarButtonsToRemove
            ]);
            overrides = { ...overrides, modeBarButtonsToRemove: Array.from(merged) };
        }
        return { ...base, ...overrides };
    }

    renderVarianceChart(filteredDeals) {
        if (!this.elements.varianceChart) return;
        const container = this.elements.varianceChart;
        if (!filteredDeals.length) {
            container.innerHTML = '<div class="empty-state">No qualifying deals for selection.</div>';
            return;
        }
        const topDeals = filteredDeals.slice().sort((a, b) => b.difference - a.difference).slice(0, 20);
        const labels = topDeals.map((deal) => deal.deal_id);
        const formatted = topDeals.map((deal) => deal.formatted_quantity);
        const comparison = topDeals.map((deal) => deal.comparison_quantity);
        const difference = topDeals.map((deal) => Math.max(deal.difference, 0));

        const traces = [
            {
                type: 'bar',
                name: 'Comparison',
                orientation: 'h',
                x: comparison,
                y: labels,
                marker: { color: '#9ca3af' },
                hovertemplate: 'Deal %{y}<br>Comparison: %{x:$,.2f}<extra></extra>'
            },
            {
                type: 'bar',
                name: 'Formatted',
                orientation: 'h',
                x: formatted,
                y: labels,
                marker: { color: '#2563eb', opacity: 0.55 },
                hovertemplate: 'Deal %{y}<br>Formatted: %{x:$,.2f}<extra></extra>'
            },
            {
                type: 'bar',
                name: 'Difference',
                orientation: 'h',
                x: difference,
                y: labels,
                base: comparison,
                marker: { color: '#ef4444' },
                hovertemplate: 'Deal %{y}<br>Difference: %{x:$,.2f}<extra></extra>'
            }
        ];

        const layout = {
            barmode: 'overlay',
            hovermode: 'closest',
            dragmode: 'zoom',
            margin: { l: 140, r: 30, t: 20, b: 40 },
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            xaxis: {
                title: 'Total USD',
                tickprefix: '$',
                hoverformat: '$,.2f',
                zeroline: false,
                automargin: true
            },
            yaxis: {
                automargin: true
            },
            legend: {
                orientation: 'h',
                y: -0.25
            }
        };

        Plotly.react(
            container,
            traces,
            layout,
            this.getPlotlyConfig({ modeBarButtonsToAdd: ['hovercompare', 'hoverclosest'] })
        );
        container.on('plotly_click', (event) => {
            if (!event || !event.points || !event.points[0]) return;
            const dealId = event.points[0].y;
            this.toggleDealFilter(dealId);
        });
        container.on('plotly_doubleclick', () => {
            Plotly.relayout(container, { 'xaxis.autorange': true, 'yaxis.autorange': true });
            this.resetFilters();
        });
    }

    renderWaterfallChart(filteredDeals) {
        if (!this.elements.waterfallChart) return;
        const container = this.elements.waterfallChart;
        if (!filteredDeals.length) {
            container.innerHTML = '<div class="empty-state">No cost differences available.</div>';
            return;
        }

        const breakdown = this.aggregateCostBreakdown(filteredDeals);
        const labels = ['Comparison Total'];
        const measures = ['absolute'];
        const values = [breakdown.comparisonTotal];
        const text = [this.formatCurrency(breakdown.comparisonTotal)];
        const colors = ['#1e3a8a'];

        breakdown.costs.forEach((cost) => {
            labels.push(cost.cost_type);
            measures.push('relative');
            values.push(cost.difference);
            text.push(this.formatCurrency(cost.difference));
            if (cost.status === 'Unregistered') {
                colors.push('#fb923c');
            } else if (cost.difference >= 0) {
                colors.push('#ef4444');
            } else {
                colors.push('#10b981');
            }
        });

        labels.push('Formatted Total');
        measures.push('total');
        values.push(breakdown.formattedTotal);
        text.push(this.formatCurrency(breakdown.formattedTotal));
        colors.push('#1e40af');

        const trace = {
            type: 'waterfall',
            orientation: 'v',
            measure: measures,
            x: labels,
            text: text,
            textposition: 'outside',
            y: values,
            connector: { line: { color: '#94a3b8' } },
            decreasing: { marker: { color: '#10b981' } },
            increasing: { marker: { color: '#ef4444' } },
            totals: { marker: { color: '#1e3a8a' } },
            marker: { color: colors },
            hovertemplate: '%{x}<br>%{y:$,.2f}<extra></extra>'
        };

        const layout = {
            margin: { t: 20, l: 60, r: 30, b: 60 },
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            dragmode: 'zoom',
            yaxis: {
                title: 'USD',
                tickprefix: '$',
                hoverformat: '$,.2f'
            }
        };

        Plotly.react(container, [trace], layout, this.getPlotlyConfig());
        container.on('plotly_click', (event) => {
            if (!event || !event.points || !event.points[0]) return;
            const label = event.points[0].x;
            if (label && label !== 'Comparison Total' && label !== 'Formatted Total') {
                this.toggleCostFilter(label);
            }
        });
        container.on('plotly_doubleclick', () => {
            Plotly.relayout(container, { 'xaxis.autorange': true, 'yaxis.autorange': true });
            this.resetFilters();
        });
    }

    renderTreemap(filteredDeals) {
        if (!this.elements.treemapChart) return;
        const container = this.elements.treemapChart;
        const unregistered = this.aggregateUnregistered(filteredDeals);
        if (!unregistered.length) {
            container.innerHTML = '<div class="empty-state">No unregistered costs detected.</div>';
            return;
        }

        const totalImpact = unregistered.reduce((sum, item) => sum + item.impact, 0);
        const labels = ['Unregistered Costs', ...unregistered.map((item) => item.cost_type)];
        const parents = ['', ...unregistered.map(() => 'Unregistered Costs')];
        const values = [totalImpact, ...unregistered.map((item) => item.impact)];
        const colors = [0, ...unregistered.map((item) => item.deal_count)];
        const text = ['Total', ...unregistered.map((item) => `${item.deal_count} deals`)].map((info, idx) => `${labels[idx]}<br>${info}`);

        const trace = {
            type: 'treemap',
            labels,
            parents,
            values,
            textinfo: 'label+value',
            hovertemplate: '<b>%{label}</b><br>Impact: %{value:$,.2f}<extra>%{customdata}</extra>',
            customdata: ['Overall', ...unregistered.map((item) => `${item.deal_count} deals`)],
            marker: {
                colors,
                colorscale: 'YlOrRd'
            }
        };

        const layout = {
            margin: { t: 20, l: 20, r: 20, b: 20 },
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)'
        };

        Plotly.react(
            container,
            [trace],
            layout,
            this.getPlotlyConfig({ modeBarButtonsToRemove: ['select2d', 'lasso2d'] })
        );
        container.on('plotly_click', (event) => {
            if (!event || !event.points || !event.points[0]) return;
            const label = event.points[0].label;
            if (label && label !== 'Unregistered Costs') {
                this.toggleCostFilter(label);
            }
        });
        container.on('plotly_doubleclick', () => this.resetFilters());
    }

    renderHeatmap() {
        if (!this.elements.heatmapChart || !this.analysisData?.heatmap) return;
        const container = this.elements.heatmapChart;
        const data = this.analysisData.heatmap;
        const dealFilter = this.filters.deals;
        const costFilter = this.filters.costTypes;

        const selectedDeals = data.deal_ids.filter((deal) => !dealFilter.size || dealFilter.has(deal));
        const selectedCosts = data.cost_types.filter((cost) => !costFilter.size || costFilter.has(cost));

        if (!selectedDeals.length || !selectedCosts.length) {
            container.innerHTML = '<div class="empty-state">No matrix values for current selection.</div>';
            return;
        }

        const dealIndex = new Map(data.deal_ids.map((deal, index) => [deal, index]));
        const costIndex = new Map(data.cost_types.map((cost, index) => [cost, index]));

        const matrix = selectedDeals.map((deal) => {
            const rowIndex = dealIndex.get(deal);
            return selectedCosts.map((cost) => data.matrix[rowIndex][costIndex.get(cost)]);
        });

        const statusMatrix = selectedDeals.map((deal) => {
            const rowIndex = dealIndex.get(deal);
            return selectedCosts.map((cost) => data.status_matrix[rowIndex][costIndex.get(cost)]);
        });

        const hover = selectedDeals.map((deal) => {
            const rowIndex = dealIndex.get(deal);
            return selectedCosts.map((cost) => data.hover[rowIndex][costIndex.get(cost)]);
        });

        const colorscale = [
            [0.0, '#9ca3af'],
            [0.25, '#1e3a8a'],
            [0.35, '#60a5fa'],
            [0.5, '#ffffff'],
            [0.75, '#fca5a5'],
            [1.0, '#b91c1c']
        ];

        const trace = {
            type: 'heatmap',
            x: selectedCosts,
            y: selectedDeals,
            z: matrix,
            text: statusMatrix,
            hoverinfo: 'text',
            hovertext: hover,
            colorscale,
            zmin: -10,
            zmax: 2,
            showscale: false
        };

        const layout = {
            margin: { t: 30, l: 140, r: 20, b: 80 },
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            dragmode: 'zoom',
            xaxis: {
                tickangle: -45,
                automargin: true
            },
            yaxis: {
                automargin: true
            }
        };

        Plotly.react(
            container,
            [trace],
            layout,
            this.getPlotlyConfig({ modeBarButtonsToAdd: ['hovercompare'], modeBarButtonsToRemove: ['select2d', 'lasso2d'] })
        );
        container.on('plotly_click', (event) => {
            if (!event || !event.points || !event.points[0]) return;
            const point = event.points[0];
            const dealId = point.y;
            const costType = point.x;
            if (dealId) {
                this.toggleDealFilter(dealId);
            }
            if (costType) {
                this.toggleCostFilter(costType);
            }
        });
        container.on('plotly_doubleclick', () => {
            Plotly.relayout(container, { 'xaxis.autorange': true, 'yaxis.autorange': true });
            this.resetFilters();
        });
    }

    aggregateCostBreakdown(deals) {
        const breakdownMap = new Map();
        let comparisonTotal = 0;
        let formattedTotal = 0;

        deals.forEach((deal) => {
            comparisonTotal += deal.comparison_quantity || 0;
            formattedTotal += deal.formatted_quantity || 0;
            (deal.costs || []).forEach((cost) => {
                const current = breakdownMap.get(cost.cost_type) || {
                    cost_type: cost.cost_type,
                    formatted_total: 0,
                    comparison_total: 0,
                    difference: 0,
                    status: 'Registered'
                };
                current.formatted_total += cost.formatted || 0;
                current.comparison_total += cost.comparison || 0;
                current.difference += cost.difference || 0;
                if (cost.status === 'Unregistered') {
                    current.status = 'Unregistered';
                } else if (cost.status === 'Partial' && current.status !== 'Unregistered') {
                    current.status = 'Partial';
                }
                breakdownMap.set(cost.cost_type, current);
            });
        });

        const costs = Array.from(breakdownMap.values()).sort((a, b) => Math.abs(b.difference) - Math.abs(a.difference));
        return { costs, comparisonTotal, formattedTotal };
    }

    aggregateUnregistered(deals) {
        const registry = new Map();
        deals.forEach((deal) => {
            (deal.costs || []).forEach((cost) => {
                if (cost.status === 'Unregistered') {
                    const current = registry.get(cost.cost_type) || { cost_type: cost.cost_type, impact: 0, deal_count: 0, deals: new Set() };
                    current.impact += Math.abs(cost.difference || 0);
                    current.deals.add(deal.deal_id);
                    current.deal_count = current.deals.size;
                    registry.set(cost.cost_type, current);
                }
            });
        });
        return Array.from(registry.values()).map((item) => ({
            cost_type: item.cost_type,
            impact: item.impact,
            deal_count: item.deal_count,
            deals: Array.from(item.deals)
        })).sort((a, b) => b.impact - a.impact);
    }

    getFilteredDeals() {
        if (!this.analysisData?.deals) {
            return [];
        }
        const deals = this.analysisData.deals.map((deal) => ({ ...deal, costs: (deal.costs || []).map((cost) => ({ ...cost })) }));
        const positiveDeals = deals.filter((deal) => (deal.difference || 0) > 0);
        let filtered = positiveDeals;
        if (this.filters.deals.size) {
            filtered = filtered.filter((deal) => this.filters.deals.has(deal.deal_id));
        }
        if (this.filters.costTypes.size) {
            filtered = filtered
                .map((deal) => ({
                    ...deal,
                    costs: deal.costs.filter((cost) => this.filters.costTypes.has(cost.cost_type))
                }))
                .filter((deal) => deal.costs.length > 0);
        }
        return filtered;
    }

    updateFiltersUI() {
        if (!this.elements.filters) return;
        const container = this.elements.filters;
        container.innerHTML = '';

        const hasFilters = this.filters.deals.size || this.filters.costTypes.size;
        if (!hasFilters) {
            const placeholder = document.createElement('span');
            placeholder.textContent = 'Click on any chart element to drill down.';
            placeholder.style.color = 'var(--gray-500)';
            container.appendChild(placeholder);
            return;
        }

        const createChip = (label, type) => {
            const chip = document.createElement('span');
            chip.className = 'filter-chip';
            chip.textContent = label;
            const remove = document.createElement('button');
            remove.type = 'button';
            remove.textContent = '×';
            remove.addEventListener('click', () => {
                if (type === 'deal') {
                    this.filters.deals.delete(label);
                } else {
                    this.filters.costTypes.delete(label);
                }
                this.refresh();
            });
            chip.appendChild(remove);
            return chip;
        };

        this.filters.deals.forEach((deal) => container.appendChild(createChip(deal, 'deal')));
        this.filters.costTypes.forEach((cost) => container.appendChild(createChip(cost, 'cost')));

        const reset = document.createElement('button');
        reset.type = 'button';
        reset.className = 'btn ghost';
        reset.textContent = 'Clear filters';
        reset.addEventListener('click', () => this.resetFilters());
        container.appendChild(reset);
    }

    refresh() {
        this.updateFiltersUI();
        this.renderCharts();
    }

    resetFilters() {
        this.filters.deals.clear();
        this.filters.costTypes.clear();
        this.refresh();
    }

    toggleDealFilter(dealId) {
        if (!dealId) return;
        if (this.filters.deals.has(dealId)) {
            this.filters.deals.delete(dealId);
        } else {
            this.filters.deals.add(dealId);
        }
        this.refresh();
    }

    toggleCostFilter(costType) {
        if (!costType) return;
        if (this.filters.costTypes.has(costType)) {
            this.filters.costTypes.delete(costType);
        } else {
            this.filters.costTypes.add(costType);
        }
        this.refresh();
    }

    async downloadExport(exportType) {
        if (!exportType || !this.analysisToken) {
            return;
        }
        await this.apiReady;
        try {
            const response = await fetch(`${this.apiBaseUrl}/compare-deals/export/${exportType}?token=${this.analysisToken}`);
            if (!response.ok) {
                throw new Error('Unable to export comparison report');
            }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            const filenames = {
                excel: 'deal_comparison.xlsx',
                csv: 'deal_comparison.csv',
                pdf: 'deal_comparison_summary.pdf'
            };
            a.href = url;
            a.download = filenames[exportType] || 'comparison_export';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Export failed', error);
            this.showError(error.message || 'Export failed');
        }
    }

    formatCurrency(value) {
        return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 }).format(value || 0);
    }

    formatNumber(value) {
        return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(value || 0);
    }

    formatSize(bytes) {
        if (!bytes && bytes !== 0) return '';
        const units = ['B', 'KB', 'MB', 'GB'];
        let size = bytes;
        let unitIndex = 0;
        while (size >= 1024 && unitIndex < units.length - 1) {
            size /= 1024;
            unitIndex += 1;
        }
        return `${size.toFixed(1)} ${units[unitIndex]}`;
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    const processor = new ExcelProcessor();
    processor.ensureApiBaseUrl().finally(() => {
        new DealComparisonDashboard(processor);
    });
});

// ---- Mapping helpers ----
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
