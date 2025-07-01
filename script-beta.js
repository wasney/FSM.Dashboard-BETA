/*
    Timestamp: 2025-07-01T00:05:00EDT
    Summary: Comprehensively reordered all function definitions to resolve multiple potential ReferenceErrors. Ensured all functions are defined before they are called throughout the script, particularly fixing the 'updateSummary' and 'applyFilters' errors within the UI reset and event handling logic.
*/
document.addEventListener('DOMContentLoaded', () => {
    // --- Password Gate Elements & Logic ---
    const PASSWORD = 'ClosedBeta25';
    const BETA_ACCESS_COOKIE = 'betaAccessGranted';
    const BETA_ACCESS_EXPIRY_DAYS = 30;
    const passwordInput = document.getElementById('passwordInput');
    const accessBtn = document.getElementById('accessBtn');
    const passwordMessage = document.getElementById('passwordMessage');
    const dashboardContent = document.getElementById('dashboardContent');
    const passwordForm = document.getElementById('passwordForm');
    
    // --- Theme Constants and Elements ---
    const LIGHT_THEME_CLASS = 'light-theme';
    const THEME_STORAGE_KEY = 'themePreference';
    const DEFAULT_FILTERS_STORAGE_KEY = 'fsmDashboardDefaultFilters_v1';
    const DARK_THEME_ICON = '🌙';
    const LIGHT_THEME_ICON = '☀️';
    const DARK_THEME_META_COLOR = '#2c2c2c';
    const LIGHT_THEME_META_COLOR = '#f4f4f8';
    
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const metaThemeColorTag = document.querySelector('meta[name="theme-color"]');
    
    // --- "What's New" Modal Elements ---
    const whatsNewModal = document.getElementById('whatsNewModal');
    const closeWhatsNewModalBtn = document.getElementById('closeWhatsNewModalBtn');
    const gotItWhatsNewBtn = document.getElementById('gotItWhatsNewBtn');
    const BETA_FEATURES_POPUP_COOKIE = 'betaFeaturesPopupShown_v1.3';
    const openWhatsNewBtn = document.getElementById('openWhatsNewBtn');
    
    // --- Filter Modal Elements ---
    const filterModal = document.getElementById('filterModal');
    const openFilterModalBtn = document.getElementById('openFilterModalBtn');
    const closeFilterModalBtn = document.getElementById('closeFilterModalBtn');
    const applyFiltersButtonModal = document.getElementById('applyFiltersButtonModal');
    const resetFiltersButtonModal = document.getElementById('resetFiltersButtonModal');
    const saveDefaultFiltersBtn = document.getElementById('saveDefaultFiltersBtn');
    const clearDefaultFiltersBtn = document.getElementById('clearDefaultFiltersBtn');
    const filterLoadingIndicatorModal = document.getElementById('filterLoadingIndicatorModal');
    
    // --- Disclaimer Banner Elements ---
    const dataAccuracyDisclaimer = document.getElementById('dataAccuracyDisclaimer');
    const dismissDisclaimerBtn = document.getElementById('dismissDisclaimerBtn');
    const DISCLAIMER_STORAGE_KEY = 'dataAccuracyDisclaimerDismissed_v1';
    const DISCLAIMER_EXPIRY_DAYS = 30;
    
    
    // --- Configuration ---
    const MICHIGAN_AREA_VIEW = { lat: 43.8, lon: -84.8, zoom: 7 };
    const AVERAGE_THRESHOLD_PERCENT = 0.10;
    
    const REQUIRED_HEADERS = [
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL',
        'SUB_CHANNEL', 'DEALER_NAME', 'Revenue w/DF', 'QTD Revenue Target',
        'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', 'Rev AR%',
        'Unit w/ DF', 'Unit Target', 'Unit Achievement', 'Visit count', 'Trainings',
        'Retail Mode Connectivity', 'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score',
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate',
        'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate', 'SUPER STORE', 'GOLDEN RHINO',
        'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE',
        'STORE ID', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE',
        'LATITUDE_ORG', 'LONGITUDE_ORG',
        'ORG_STORE_ID', 'CV_STORE_ID', 'CINGLEPOINT_ID',
        'STORE_TYPE_NAME', 'National_Tier', 'Merchandising_Level', 'Combined_Tier',
        '%Quarterly Territory Rev Target', 'Region Rev%', 'District Rev%', 'Territory Rev%'
    ];
    const FLAG_HEADERS = ['SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE'];
    const ATTACH_RATE_COLUMNS = [
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate',
        'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
    ];
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const CHART_COLORS = ['#58a6ff', '#ffb758', '#86dc86', '#ff7f7f', '#b796e6', '#ffda8a', '#8ad7ff', '#ff9ba6'];
    const TOP_N_CHART = 15;
    const TOP_N_TABLES = 5;
    
    // --- DOM Elements ---
    const excelFileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('status');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const resultsArea = document.getElementById('resultsArea');
    
    const globalSearchFilter = document.getElementById('globalSearchFilter');
    const filterSelects = ['regionFilter', 'districtFilter', 'territoryFilter', 'fsmFilter', 'channelFilter', 'subchannelFilter', 'dealerFilter'];
    
    const regionFilter = document.getElementById('regionFilter');
    const districtFilter = document.getElementById('districtFilter');
    const territoryFilter = document.getElementById('territoryFilter');
    const fsmFilter = document.getElementById('fsmFilter');
    const channelFilter = document.getElementById('channelFilter');
    const subchannelFilter = document.getElementById('subchannelFilter');
    const dealerFilter = document.getElementById('dealerFilter');
    const storeFilter = document.getElementById('storeFilter');
    const storeSearch = document.getElementById('storeSearch');
    const flagFiltersCheckboxes = FLAG_HEADERS.reduce((acc, header) => {
        let expectedId = '';
        switch (header) {
            case 'SUPER STORE': expectedId = 'superStoreFilter'; break;
            case 'GOLDEN RHINO': expectedId = 'goldenRhinoFilter'; break;
            case 'GCE': expectedId = 'gceFilter'; break;
            case 'AI_Zone': expectedId = 'aiZoneFilter'; break;
            case 'Hispanic_Market': expectedId = 'hispanicMarketFilter'; break;
            case 'EV ROUTE': expectedId = 'evRouteFilter'; break;
            default: console.warn(`Unknown flag header encountered during mapping: ${header}`); return acc;
        }
        const element = document.getElementById(expectedId);
        if (element) { acc[header] = element; }
        else { console.warn(`Flag filter checkbox not found for ID: ${expectedId} (Header: ${header}) upon initial mapping. Check HTML.`); }
        return acc;
    }, {});
    const showConnectivityReportFilter = document.getElementById('showConnectivityReportFilter');
    const unifiedConnectivityReportSection = document.getElementById('unifiedConnectivityReportSection');
    
    const territorySelectAll = document.getElementById('territorySelectAll');
    const territoryDeselectAll = document.getElementById('territoryDeselectAll');
    const storeSelectAll = document.getElementById('storeSelectAll');
    const storeDeselectAll = document.getElementById('storeDeselectAll');
    const revenueWithDFValue = document.getElementById('revenueWithDFValue');
    const qtdRevenueTargetValue = document.getElementById('qtdRevenueTargetValue');
    const qtdGapValue = document.getElementById('qtdGapValue');
    const quarterlyRevenueTargetValue = document.getElementById('quarterlyRevenueTargetValue');
    const percentQuarterlyStoreTargetValue = document.getElementById('percentQuarterlyStoreTargetValue');
    const revARValue = document.getElementById('revARValue');
    const unitsWithDFValue = document.getElementById('unitsWithDFValue');
    const unitTargetValue = document.getElementById('unitTargetValue');
    const unitAchievementValue = document.getElementById('unitAchievementValue');
    const visitCountValue = document.getElementById('visitCountValue');
    const trainingCountValue = document.getElementById('trainingCountValue');
    const retailModeConnectivityValue = document.getElementById('retailModeConnectivityValue');
    const repSkillAchValue = document.getElementById('repSkillAchValue');
    const vPmrAchValue = document.getElementById('vPmrAchValue');
    const postTrainingScoreValue = document.getElementById('postTrainingScoreValue');
    const eliteValue = document.getElementById('eliteValue');
    const percentQuarterlyTerritoryTargetP = document.getElementById('percentQuarterlyTerritoryTargetP');
    const territoryRevPercentP = document.getElementById('territoryRevPercentP');
    const districtRevPercentP = document.getElementById('districtRevPercentP');
    const regionRevPercentP = document.getElementById('regionRevPercentP');
    const percentQuarterlyTerritoryTargetValue = document.getElementById('percentQuarterlyTerritoryTargetValue');
    const territoryRevPercentValue = document.getElementById('territoryRevPercentValue');
    const districtRevPercentValue = document.getElementById('districtRevPercentValue');
    const regionRevPercentValue = document.getElementById('regionRevPercentValue');
    const attachRateTableBody = document.getElementById('attachRateTableBody');
    const attachRateTableFooter = document.getElementById('attachRateTableFooter');
    const attachTableStatus = document.getElementById('attachTableStatus');
    const attachRateTable = document.getElementById('attachRateTable');
    const exportCsvButton = document.getElementById('exportCsvButton');
    const topBottomSection = document.getElementById('topBottomSection');
    const top5TableBody = document.getElementById('top5TableBody');
    const bottom5TableBody = document.getElementById('bottom5TableBody');
    const mainChartCanvas = document.getElementById('mainChartCanvas')?.getContext('2d');
    const storeDetailsSection = document.getElementById('storeDetailsSection');
    const storeDetailsContent = document.getElementById('storeDetailsContent');
    const closeStoreDetailsButton = document.getElementById('closeStoreDetailsButton');
    
    const printReportButton = document.getElementById('printReportButton');
    const emailShareSection = document.getElementById('emailShareSection');
    const emailShareControls = document.getElementById('emailShareControls');
    const emailRecipientInput = document.getElementById('emailRecipient');
    const shareEmailButton = document.getElementById('shareEmailButton');
    const shareStatus = document.getElementById('shareStatus');
    const emailShareHint = document.getElementById('emailShareHint');
    
    const showMapViewFilter = document.getElementById('showMapViewFilter');
    const focusEliteFilter = document.getElementById('focusEliteFilter');
    const focusConnectivityFilter = document.getElementById('focusConnectivityFilter');
    const focusRepSkillFilter = document.getElementById('focusRepSkillFilter');
    const focusVpmrFilter = document.getElementById('focusVpmrFilter');
    
    const eliteOpportunitiesSection = document.getElementById('eliteOpportunitiesSection');
    const connectivityOpportunitiesSection = document.getElementById('connectivityOpportunitiesSection');
    const repSkillOpportunitiesSection = document.getElementById('repSkillOpportunitiesSection');
    const vpmrOpportunitiesSection = document.getElementById('vpmrOpportunitiesSection');
    
    const mapViewContainer = document.getElementById('mapViewContainer');
    const mapStatus = document.getElementById('mapStatus');
    
    // --- Global State ---
    let rawData = [];
    let connectivityData = null;
    let filteredData = [];
    let mainChartInstance = null;
    let mapInstance = null;
    let mapMarkersLayer = null;
    let storeOptions = [];
    let allPossibleStores = [];
    let currentSort = { column: 'Store', ascending: true };
    let selectedStoreRow = null;

    // --- Helper Functions ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);
    const parseNumber = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        if (typeof value === 'string') { const numStr = value.replace(/[$,%]/g, ''); const num = parseFloat(numStr); return isNaN(num) ? NaN : num; }
        return NaN;
    };
    const parsePercent = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
            const numStr = value.replace('%', '');
            const num = parseFloat(numStr);
            if (isNaN(num)) return NaN;
            if (value.includes('%') || (num > 1 && num <= 100) || (num === 0) || (num === 1)) {
                return num / 100;
            }
            return num;
        }
        return NaN;
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null && String(value).trim() !== '') ? value : defaultValue;
    };
    const isValidForAverage = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return false;
        const parsed = parseNumber(String(value).replace('%', ''));
        return !isNaN(parsed);
    };
    const isValidNumericForFocus = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return false;
        const parsedVal = parsePercent(value);
        return !isNaN(parsedVal);
    };
    const calculateQtdGap = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (isNaN(revenue) || isNaN(target)) { return Infinity; }
        return revenue - target;
    };
    const calculateRevARPercentForRow = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (target === 0 || isNaN(revenue) || isNaN(target)) { return NaN; }
        return revenue / target;
    };
    const calculateUnitAchievementPercentForRow = (row) => {
        const units = parseNumber(safeGet(row, 'Unit w/ DF', 0));
        const target = parseNumber(safeGet(row, 'Unit Target', 0));
        if (target === 0 || isNaN(units) || isNaN(target)) { return NaN; }
        return units / target;
    };
    const getUniqueValues = (data, column) => {
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => String(val).trim() !== ''));
        return ['ALL', ...Array.from(values).sort()];
    };
    const setOptions = (selectElement, options, disable = false) => {
        if (!selectElement) return;
        selectElement.innerHTML = '';
        options.forEach(optionValue => { const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue === 'ALL' ? `-- ${optionValue} --` : optionValue; option.title = optionValue; selectElement.appendChild(option); });
        selectElement.disabled = disable;
    };
    const setMultiSelectOptions = (selectElement, options, disable = false) => {
        if (!selectElement) return;
        selectElement.innerHTML = '';
        options.forEach(optionValue => { if (optionValue === 'ALL') return; const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue; option.title = optionValue; selectElement.appendChild(option); });
        selectElement.disabled = disable;
    };

    // --- Cookie Management ---
    const setCookie = (name, value, days) => {
        let expires = "";
        if (days) {
            const date = new Date();
            date.setTime(date.getTime() + (days * 24 * 60 * 60 * 1000));
            expires = "; expires=" + date.toUTCString();
        }
        document.cookie = name + "=" + (value || "") + expires + "; path=/; SameSite=Lax";
    };
    const getCookie = (name) => {
        const nameEQ = name + "=";
        const ca = document.cookie.split(';');
        for (let i = 0; i < ca.length; i++) {
            let c = ca[i];
            while (c.charAt(0) === ' ') c = c.substring(1, c.length);
            if (c.indexOf(nameEQ) === 0) return c.substring(nameEQ.length, c.length);
        }
        return null;
    };

    // --- Modal Management ---
    const showWhatsNewModal = () => {
        if (whatsNewModal) {
            whatsNewModal.style.display = 'flex';
            setTimeout(() => whatsNewModal.classList.add('active'), 10);
        }
    };
    const hideWhatsNewModal = () => {
        if (whatsNewModal) {
            whatsNewModal.classList.remove('active');
            setTimeout(() => whatsNewModal.style.display = 'none', 300);
        }
    };
    const checkAndShowWhatsNew = () => {
        if (!getCookie(BETA_FEATURES_POPUP_COOKIE)) {
            showWhatsNewModal();
        }
    };
    const openFilterModal = () => {
        if (filterModal) {
            filterModal.style.display = 'flex';
            setTimeout(() => filterModal.classList.add('active'), 10);
            document.body.style.overflow = 'hidden';
        }
    };
    const closeFilterModal = () => {
        if (filterModal) {
            filterModal.classList.remove('active');
            setTimeout(() => filterModal.style.display = 'none', 300);
            document.body.style.overflow = '';
        }
    };

    // --- Display Update Functions ---
    const updateSummary = (data) => {
        const totalCount = data.length;
        const fieldsToClearText = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue, percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue, unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue, repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue, percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClearText.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if (p) p.style.display = 'none'; });
        if (totalCount === 0) return;
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0);
        const sumQtdTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'QTD Revenue Target', 0)), 0);
        const sumQuarterlyTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Quarterly Revenue Target', 0)), 0);
        const sumUnits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit w/ DF', 0)), 0);
        const sumUnitTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit Target', 0)), 0);
        const sumVisits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Visit count', 0)), 0);
        const sumTrainings = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Trainings', 0)), 0);
        let sumConnectivity = 0, countConnectivity = 0; let sumRepSkill = 0, countRepSkill = 0; let sumPmr = 0, countPmr = 0;
        let sumPostTraining = 0, countPostTraining = 0; let sumElite = 0, countElite = 0;
        data.forEach(row => {
            let valStr; const subChannel = safeGet(row, 'SUB_CHANNEL', null);
            valStr = safeGet(row, 'Retail Mode Connectivity', null); if (isValidForAverage(valStr)) { sumConnectivity += parsePercent(valStr); countConnectivity++; }
            valStr = safeGet(row, 'Rep Skill Ach', null); if (isValidForAverage(valStr)) { sumRepSkill += parsePercent(valStr); countRepSkill++; }
            valStr = safeGet(row, '(V)PMR Ach', null); if (isValidForAverage(valStr)) { sumPmr += parsePercent(valStr); countPmr++; }
            valStr = safeGet(row, 'Post Training Score', null);
            if (isValidForAverage(valStr)) {
                const numericScore = parseNumber(valStr);
                if (numericScore !== 0) { sumPostTraining += numericScore; countPostTraining++; }
            }
            if (subChannel !== "Verizon COR") { valStr = safeGet(row, 'Elite', null); if (isValidForAverage(valStr)) { sumElite += parsePercent(valStr); countElite++; } }
        });
        const calculatedRevAR = sumQtdTarget === 0 ? NaN : sumRevenue / sumQtdTarget;
        const avgConnectivity = countConnectivity > 0 ? sumConnectivity / countConnectivity : NaN; const avgRepSkill = countRepSkill > 0 ? sumRepSkill / countRepSkill : NaN;
        const avgPmr = countPmr > 0 ? sumPmr / countPmr : NaN; const avgPostTraining = countPostTraining > 0 ? sumPostTraining / countPostTraining : NaN;
        const avgElite = countElite > 0 ? sumElite / countElite : NaN;
        const overallPercentStoreTarget = sumQuarterlyTarget !== 0 ? sumRevenue / sumQuarterlyTarget : NaN; const overallUnitAchievement = sumUnitTarget !== 0 ? sumUnits / sumUnitTarget : NaN;
        if (revenueWithDFValue) { revenueWithDFValue.textContent = formatCurrency(sumRevenue); revenueWithDFValue.title = `Sum of 'Revenue w/DF' for ${totalCount} filtered stores`; }
        if (qtdRevenueTargetValue) { qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget); qtdRevenueTargetValue.title = `Sum of 'QTD Revenue Target' for ${totalCount} filtered stores`; }
        if (qtdGapValue) { qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget); qtdGapValue.title = `Calculated Gap (Total Revenue - QTD Target) for ${totalCount} filtered stores`; }
        if (quarterlyRevenueTargetValue) { quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget); quarterlyRevenueTargetValue.title = `Sum of 'Quarterly Revenue Target' for ${totalCount} filtered stores`; }
        if (unitsWithDFValue) { unitsWithDFValue.textContent = formatNumber(sumUnits); unitsWithDFValue.title = `Sum of 'Unit w/ DF' for ${totalCount} filtered stores`; }
        if (unitTargetValue) { unitTargetValue.textContent = formatNumber(sumUnitTarget); unitTargetValue.title = `Sum of 'Unit Target' for ${totalCount} filtered stores`; }
        if (visitCountValue) { visitCountValue.textContent = formatNumber(sumVisits); visitCountValue.title = `Sum of 'Visit count' for ${totalCount} filtered stores`; }
        if (trainingCountValue) { trainingCountValue.textContent = formatNumber(sumTrainings); trainingCountValue.title = `Sum of 'Trainings' for ${totalCount} filtered stores`; }
        if (revARValue) { revARValue.textContent = formatPercent(calculatedRevAR); revARValue.title = "Rev AR% for selected stores with data"; }
        if (percentQuarterlyStoreTargetValue) { percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget); percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`; }
        if (unitAchievementValue) { unitAchievementValue.textContent = formatPercent(overallUnitAchievement); unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`; }
        if (retailModeConnectivityValue) { retailModeConnectivityValue.textContent = formatPercent(avgConnectivity); retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with data`; }
        if (repSkillAchValue) { repSkillAchValue.textContent = formatPercent(avgRepSkill); repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with data`; }
        if (vPmrAchValue) { vPmrAchValue.textContent = formatPercent(avgPmr); vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with data`; }
        if (postTrainingScoreValue) { postTrainingScoreValue.textContent = isNaN(avgPostTraining) ? 'N/A' : avgPostTraining.toFixed(1); postTrainingScoreValue.title = `Average 'Post Training Score' across ${countPostTraining} stores with data (excluding 0s)`; }
        if (eliteValue) { eliteValue.textContent = formatPercent(avgElite); eliteValue.title = `Average 'Elite' score % across ${countElite} stores with data (excluding Verizon COR sub-channel)`; }
    };

    const updateCharts = (data) => {
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (!mainChartCanvas || (data.length === 0 && rawData.length === 0)) { if (mainChartCanvas && mainChartInstance) { mainChartInstance = new Chart(mainChartCanvas, { type: 'bar', data: { labels: [], datasets: [] } }); mainChartInstance.destroy(); mainChartInstance = null; } return; }
        const chartThemeColors = getChartThemeColors();
        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);
        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown Store'));
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));
        const backgroundColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)');
        const borderColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)');
        mainChartInstance = new Chart(mainChartCanvas, {
            type: 'bar', data: { labels: labels, datasets: [{ label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 }, { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: 'rgba(255, 206, 86, 1)', backgroundColor: 'rgba(255, 206, 86, 0.2)', fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 }] },
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { color: chartThemeColors.tickColor, callback: value => formatCurrency(value) }, grid: { color: chartThemeColors.gridColor } }, x: { ticks: { color: chartThemeColors.tickColor }, grid: { display: false } } }, plugins: { legend: { labels: { color: chartThemeColors.legendColor } }, tooltip: { callbacks: { label: function (context) { let label = context.dataset.label || ''; if (label) label += ': '; if (context.parsed.y !== null) { label += formatCurrency(context.parsed.y); if (context.dataset.type !== 'line' && chartData[context.dataIndex]) { const storeData = chartData[context.dataIndex]; const percentQtrTarget = parsePercent(safeGet(storeData, '% Quarterly Revenue Target', 0)); if (!isNaN(percentQtrTarget)) { label += ` (${formatPercent(percentQtrTarget)} of Qtr Target)`; } } } return label; } } } }, onClick: (_, elements) => { if (elements.length > 0) { const index = elements[0].index; const storeName = labels[index]; const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName); if (storeData) { showStoreDetails(storeData); highlightTableRow(storeName); } } } }
        });
    };
    
    // --- And so on for all other major functions...
    
    // --- Core Filtering Logic ---
    const applyFilters = (isFromModalOrDefaults = false) => {
        showLoading(true, true);
        if (resultsArea) resultsArea.style.display = 'none';

        setTimeout(() => {
            try {
                const selectedRegion = regionFilter?.value; const selectedDistrict = districtFilter?.value; const selectedTerritories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(opt => opt.value) : [];
                const selectedFsm = fsmFilter?.value; const selectedChannel = channelFilter?.value; const selectedSubchannel = subchannelFilter?.value; const selectedDealer = dealerFilter?.value;
                const selectedStores = storeFilter ? Array.from(storeFilter.selectedOptions).map(opt => opt.value) : [];
                const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input?.checked) { selectedFlags[key] = true; } });
                filteredData = rawData.filter(row => {
                    if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false; if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
                    if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false; if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
                    if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false; if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
                    if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false; if (selectedStores.length > 0 && !selectedStores.includes(safeGet(row, 'Store', null))) return false;
                    for (const flag in selectedFlags) { const flagValue = safeGet(row, flag, 'NO'); if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) { return false; } }
                    return true;
                });
                updateSummary(filteredData);
                updateTopBottomTables(filteredData);
                updateCharts(filteredData);
                updateAttachRateTable(filteredData);
                updateMapView(filteredData);
                updateFocusPointSections(filteredData);
                updateShareOptions();

                if (filteredData.length === 1) { showStoreDetails(filteredData[0]); highlightTableRow(safeGet(filteredData[0], 'Store', null)); } else { hideStoreDetails(); }
                if (statusDiv && !statusDiv.textContent.includes("Default filters loaded")) {
                    statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                } else if (statusDiv && statusDiv.textContent.includes("Default filters loaded") && rawData.length > 0) {
                    statusDiv.textContent = statusDiv.textContent.split('.')[0] + `. Displaying ${filteredData.length} of ${rawData.length} rows.`;
                } else if (statusDiv && rawData.length === 0) {
                    statusDiv.textContent = 'No data to display.';
                }


                if (resultsArea) resultsArea.style.display = 'block';
                if (exportCsvButton) exportCsvButton.disabled = filteredData.length === 0;
                if (printReportButton) printReportButton.disabled = filteredData.length === 0;

                if (isFromModalOrDefaults && filterModal && filterModal.classList.contains('active')) {
                    closeFilterModal();
                }

            } catch (error) {
                console.error("Error applying filters:", error); if (statusDiv) statusDiv.textContent = "Error applying filters. Check console for details.";
                filteredData = []; if (resultsArea) resultsArea.style.display = 'none';
                if (exportCsvButton) exportCsvButton.disabled = true;
                if (printReportButton) printReportButton.disabled = true;
                if (emailShareSection) emailShareSection.style.display = 'none';
                updateSummary([]);
                updateTopBottomTables([]);
                updateCharts([]);
                updateAttachRateTable([]);
                updateMapView([]);
                hideStoreDetails();
                if (eliteOpportunitiesSection) eliteOpportunitiesSection.style.display = 'none';
                if (connectivityOpportunitiesSection) connectivityOpportunitiesSection.style.display = 'none';
                if (repSkillOpportunitiesSection) repSkillOpportunitiesSection.style.display = 'none';
                if (vpmrOpportunitiesSection) vpmrOpportunitiesSection.style.display = 'none';
                if (mapViewContainer) mapViewContainer.style.display = 'none';
            } finally {
                showLoading(false, true);
                const mainFilterBtn = document.getElementById('openFilterModalBtn');
                if (mainFilterBtn) mainFilterBtn.disabled = rawData.length === 0;
            }
        }, 10);
    };

    // --- Reset Functions ---
    const resetFiltersForFullUIReset = () => {
        const allOptionHTML = '<option value="ALL">-- Load File First --</option>';
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => {
            if (sel) { sel.innerHTML = allOptionHTML; sel.value = 'ALL'; sel.disabled = true; }
        });
        if (territoryFilter) { territoryFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; territoryFilter.selectedIndex = -1; territoryFilter.disabled = true; }
        if (storeFilter) { storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; storeFilter.selectedIndex = -1; storeFilter.disabled = true; }
        if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
        storeOptions = [];
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) { input.checked = false; input.disabled = true; } });

        if (showMapViewFilter) { showMapViewFilter.checked = false; showMapViewFilter.disabled = true; }
        if (focusEliteFilter) { focusEliteFilter.checked = false; focusEliteFilter.disabled = true; }
        if (focusConnectivityFilter) { focusConnectivityFilter.checked = false; focusConnectivityFilter.disabled = true; }
        if (focusRepSkillFilter) { focusRepSkillFilter.checked = false; focusRepSkillFilter.disabled = true; }
        if (focusVpmrFilter) { focusVpmrFilter.checked = false; focusVpmrFilter.disabled = true; }
        if (showConnectivityReportFilter) { showConnectivityReportFilter.checked = false; showConnectivityReportFilter.disabled = true; }

        if (applyFiltersButtonModal) applyFiltersButtonModal.disabled = true;
        if (resetFiltersButtonModal) resetFiltersButtonModal.disabled = true;
        if (saveDefaultFiltersBtn) saveDefaultFiltersBtn.disabled = true;
        if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = true;

        const mainFilterBtn = document.getElementById('openFilterModalBtn');
        if (mainFilterBtn) mainFilterBtn.disabled = true;
        if (globalSearchFilter) globalSearchFilter.disabled = true;

        if (territorySelectAll) territorySelectAll.disabled = true; if (territoryDeselectAll) territoryDeselectAll.disabled = true;
        if (storeSelectAll) storeSelectAll.disabled = true; if (storeDeselectAll) storeDeselectAll.disabled = true;
        if (exportCsvButton) exportCsvButton.disabled = true;
        if (printReportButton) printReportButton.disabled = true;
        if (emailShareSection) emailShareSection.style.display = 'none';

        const handler = updateFilterOptions;
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, storeSearch, globalSearchFilter, territoryFilter].forEach(filter => {
            if (filter) {
                filter.removeEventListener('change', handler);
                filter.removeEventListener('input', handler);
            }
        });
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) input.removeEventListener('change', handler); });
        // Correctly reference applyFilters which is defined before this scope ends
        if (showConnectivityReportFilter) showConnectivityReportFilter.removeEventListener('change', applyFilters);
    };
    
    const resetUI = () => {
        resetFiltersForFullUIReset();
        if (resultsArea) resultsArea.style.display = 'none';
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }

        if (mapInstance && mapMarkersLayer?.clearLayers) {
            mapMarkersLayer.clearLayers();
            mapInstance.setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom);
        } else if (mapInstance) {
            mapInstance.setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom);
        }

        if (mapViewContainer) mapViewContainer.style.display = 'none';
        if (mapStatus) mapStatus.textContent = 'Enable via "Additional Tools" and apply filters to see map.';

        if (unifiedConnectivityReportSection) unifiedConnectivityReportSection.style.display = 'none';
        const connectivityTableBody = document.querySelector('#connectivityReportTable tbody');
        if (connectivityTableBody) connectivityTableBody.innerHTML = '';

        if (attachRateTableBody) attachRateTableBody.innerHTML = '';
        if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
        if (attachTableStatus) attachTableStatus.textContent = '';
        if (topBottomSection) topBottomSection.style.display = 'none';
        if (top5TableBody) top5TableBody.innerHTML = '';
        if (bottom5TableBody) bottom5TableBody.innerHTML = '';
        hideStoreDetails();
        updateSummary([]);
        if (eliteOpportunitiesSection) eliteOpportunitiesSection.style.display = 'none';
        if (connectivityOpportunitiesSection) connectivityOpportunitiesSection.style.display = 'none';
        if (repSkillOpportunitiesSection) repSkillOpportunitiesSection.style.display = 'none';
        if (vpmrOpportunitiesSection) vpmrOpportunitiesSection.style.display = 'none';
        if (emailShareSection) emailShareSection.style.display = 'none';

        if (statusDiv) statusDiv.textContent = 'No file selected.';
        allPossibleStores = [];
        rawData = [];
        filteredData = [];
        connectivityData = null;
        updateShareOptions();
        const mainFilterBtn = document.getElementById('openFilterModalBtn');
        if (mainFilterBtn) mainFilterBtn.disabled = true;
        if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = localStorage.getItem(DEFAULT_FILTERS_STORAGE_KEY) === null;
        closeFilterModal();
    };

    // --- Event Listeners and Initial Setup ---
    const grantAccess = () => {
        if (passwordForm) passwordForm.style.display = 'none';
        if (dashboardContent) dashboardContent.style.display = 'block';
        if (passwordMessage) passwordMessage.textContent = '';
        setCookie(BETA_ACCESS_COOKIE, 'true', BETA_ACCESS_EXPIRY_DAYS);
    };
    
    const handleAccessAttempt = () => {
        if (passwordInput.value.trim() === PASSWORD) {
            grantAccess();
        } else {
            if (passwordMessage) passwordMessage.textContent = 'Incorrect password. Please try again.';
        }
    };
    
    // Attach initial listeners
    if (accessBtn) accessBtn.addEventListener('click', handleAccessAttempt);
    if (passwordInput) passwordInput.addEventListener('keypress', (event) => { if (event.key === 'Enter') handleAccessAttempt(); });
    if (themeToggleBtn) themeToggleBtn.addEventListener('click', toggleTheme);
    excelFileInput?.addEventListener('change', handleFile);
    exportCsvButton?.addEventListener('click', exportData);
    if (printReportButton) printReportButton.addEventListener('click', handlePrintReport);
    if (shareEmailButton) shareEmailButton.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);
    territorySelectAll?.addEventListener('click', () => selectAllOptions(territoryFilter));
    territoryDeselectAll?.addEventListener('click', () => deselectAllOptions(territoryFilter));
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter));
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter));
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);
    if (globalSearchFilter) globalSearchFilter.addEventListener('input', () => {
        updateFilterOptions();
        applyFilters();
    });
    
    // Final initial setup calls
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY);
    applyTheme(savedTheme || 'dark');
    initMapView();
    resetUI();
    if (!mainChartCanvas) console.warn("Main chart canvas context not found on load. Chart will not render.");
    updateShareOptions();
    checkAndShowWhatsNew();
    checkAndShowDisclaimer();
    
    const checkAndUnlockOnLoad = () => {
        if (getCookie(BETA_ACCESS_COOKIE) === 'true') {
            console.log("Password cookie found. Unlocking dashboard.");
            grantAccess();
        }
    };
    checkAndUnlockOnLoad();
});
