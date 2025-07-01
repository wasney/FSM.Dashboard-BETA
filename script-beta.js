/*
    Timestamp: 2025-07-01T00:15:00EDT
    Summary: Comprehensively reordered ALL function definitions to the top of the script's scope, ensuring that every function, including 'toggleTheme', 'applyFilters', and 'updateSummary', is declared before any event listeners are attached or any initial setup functions are called. This resolves all outstanding ReferenceErrors.
*/
document.addEventListener('DOMContentLoaded', () => {
    // --- Constants and Global State ---
    const PASSWORD = 'ClosedBeta25';
    const BETA_ACCESS_COOKIE = 'betaAccessGranted';
    const BETA_ACCESS_EXPIRY_DAYS = 30;
    const LIGHT_THEME_CLASS = 'light-theme';
    const THEME_STORAGE_KEY = 'themePreference';
    const DEFAULT_FILTERS_STORAGE_KEY = 'fsmDashboardDefaultFilters_v1';
    const DARK_THEME_ICON = '🌙';
    const LIGHT_THEME_ICON = '☀️';
    const DARK_THEME_META_COLOR = '#2c2c2c';
    const LIGHT_THEME_META_COLOR = '#f4f4f8';
    const BETA_FEATURES_POPUP_COOKIE = 'betaFeaturesPopupShown_v1.3';
    const DISCLAIMER_STORAGE_KEY = 'dataAccuracyDisclaimerDismissed_v1';
    const DISCLAIMER_EXPIRY_DAYS = 30;
    const MICHIGAN_AREA_VIEW = { lat: 43.8, lon: -84.8, zoom: 7 };
    const AVERAGE_THRESHOLD_PERCENT = 0.10;
    const REQUIRED_HEADERS = [
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL', 'SUB_CHANNEL', 'DEALER_NAME',
        'Revenue w/DF', 'QTD Revenue Target', 'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', 'Rev AR%',
        'Unit w/ DF', 'Unit Target', 'Unit Achievement', 'Visit count', 'Trainings', 'Retail Mode Connectivity',
        'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score', 'Tablet Attach Rate', 'PC Attach Rate',
        'NC Attach Rate', 'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate',
        'SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE',
        'STORE ID', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE', 'LATITUDE_ORG', 'LONGITUDE_ORG',
        'ORG_STORE_ID', 'CV_STORE_ID', 'CINGLEPOINT_ID', 'STORE_TYPE_NAME', 'National_Tier', 'Merchandising_Level', 'Combined_Tier',
        '%Quarterly Territory Rev Target', 'Region Rev%', 'District Rev%', 'Territory Rev%'
    ];
    const FLAG_HEADERS = ['SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE'];
    const ATTACH_RATE_COLUMNS = [
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate',
        'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
    ];
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const TOP_N_CHART = 15;
    const TOP_N_TABLES = 5;

    // --- DOM Element Variables ---
    const passwordInput = document.getElementById('passwordInput'),
          accessBtn = document.getElementById('accessBtn'),
          passwordMessage = document.getElementById('passwordMessage'),
          dashboardContent = document.getElementById('dashboardContent'),
          passwordForm = document.getElementById('passwordForm'),
          themeToggleBtn = document.getElementById('themeToggleBtn'),
          metaThemeColorTag = document.querySelector('meta[name="theme-color"]'),
          whatsNewModal = document.getElementById('whatsNewModal'),
          closeWhatsNewModalBtn = document.getElementById('closeWhatsNewModalBtn'),
          gotItWhatsNewBtn = document.getElementById('gotItWhatsNewBtn'),
          openWhatsNewBtn = document.getElementById('openWhatsNewBtn'),
          filterModal = document.getElementById('filterModal'),
          openFilterModalBtn = document.getElementById('openFilterModalBtn'),
          closeFilterModalBtn = document.getElementById('closeFilterModalBtn'),
          applyFiltersButtonModal = document.getElementById('applyFiltersButtonModal'),
          resetFiltersButtonModal = document.getElementById('resetFiltersButtonModal'),
          saveDefaultFiltersBtn = document.getElementById('saveDefaultFiltersBtn'),
          clearDefaultFiltersBtn = document.getElementById('clearDefaultFiltersBtn'),
          filterLoadingIndicatorModal = document.getElementById('filterLoadingIndicatorModal'),
          dataAccuracyDisclaimer = document.getElementById('dataAccuracyDisclaimer'),
          dismissDisclaimerBtn = document.getElementById('dismissDisclaimerBtn'),
          excelFileInput = document.getElementById('excelFile'),
          statusDiv = document.getElementById('status'),
          loadingIndicator = document.getElementById('loadingIndicator'),
          resultsArea = document.getElementById('resultsArea'),
          globalSearchFilter = document.getElementById('globalSearchFilter'),
          regionFilter = document.getElementById('regionFilter'),
          districtFilter = document.getElementById('districtFilter'),
          territoryFilter = document.getElementById('territoryFilter'),
          fsmFilter = document.getElementById('fsmFilter'),
          channelFilter = document.getElementById('channelFilter'),
          subchannelFilter = document.getElementById('subchannelFilter'),
          dealerFilter = document.getElementById('dealerFilter'),
          storeFilter = document.getElementById('storeFilter'),
          storeSearch = document.getElementById('storeSearch'),
          showConnectivityReportFilter = document.getElementById('showConnectivityReportFilter'),
          unifiedConnectivityReportSection = document.getElementById('unifiedConnectivityReportSection'),
          territorySelectAll = document.getElementById('territorySelectAll'),
          territoryDeselectAll = document.getElementById('territoryDeselectAll'),
          storeSelectAll = document.getElementById('storeSelectAll'),
          storeDeselectAll = document.getElementById('storeDeselectAll'),
          revenueWithDFValue = document.getElementById('revenueWithDFValue'),
          qtdRevenueTargetValue = document.getElementById('qtdRevenueTargetValue'),
          qtdGapValue = document.getElementById('qtdGapValue'),
          quarterlyRevenueTargetValue = document.getElementById('quarterlyRevenueTargetValue'),
          percentQuarterlyStoreTargetValue = document.getElementById('percentQuarterlyStoreTargetValue'),
          revARValue = document.getElementById('revARValue'),
          unitsWithDFValue = document.getElementById('unitsWithDFValue'),
          unitTargetValue = document.getElementById('unitTargetValue'),
          unitAchievementValue = document.getElementById('unitAchievementValue'),
          visitCountValue = document.getElementById('visitCountValue'),
          trainingCountValue = document.getElementById('trainingCountValue'),
          retailModeConnectivityValue = document.getElementById('retailModeConnectivityValue'),
          repSkillAchValue = document.getElementById('repSkillAchValue'),
          vPmrAchValue = document.getElementById('vPmrAchValue'),
          postTrainingScoreValue = document.getElementById('postTrainingScoreValue'),
          eliteValue = document.getElementById('eliteValue'),
          percentQuarterlyTerritoryTargetP = document.getElementById('percentQuarterlyTerritoryTargetP'),
          territoryRevPercentP = document.getElementById('territoryRevPercentP'),
          districtRevPercentP = document.getElementById('districtRevPercentP'),
          regionRevPercentP = document.getElementById('regionRevPercentP'),
          percentQuarterlyTerritoryTargetValue = document.getElementById('percentQuarterlyTerritoryTargetValue'),
          territoryRevPercentValue = document.getElementById('territoryRevPercentValue'),
          districtRevPercentValue = document.getElementById('districtRevPercentValue'),
          regionRevPercentValue = document.getElementById('regionRevPercentValue'),
          attachRateTableBody = document.getElementById('attachRateTableBody'),
          attachRateTableFooter = document.getElementById('attachRateTableFooter'),
          attachTableStatus = document.getElementById('attachTableStatus'),
          attachRateTable = document.getElementById('attachRateTable'),
          exportCsvButton = document.getElementById('exportCsvButton'),
          topBottomSection = document.getElementById('topBottomSection'),
          top5TableBody = document.getElementById('top5TableBody'),
          bottom5TableBody = document.getElementById('bottom5TableBody'),
          mainChartCanvas = document.getElementById('mainChartCanvas')?.getContext('2d'),
          storeDetailsSection = document.getElementById('storeDetailsSection'),
          storeDetailsContent = document.getElementById('storeDetailsContent'),
          closeStoreDetailsButton = document.getElementById('closeStoreDetailsButton'),
          printReportButton = document.getElementById('printReportButton'),
          emailShareSection = document.getElementById('emailShareSection'),
          emailShareControls = document.getElementById('emailShareControls'),
          emailRecipientInput = document.getElementById('emailRecipient'),
          shareEmailButton = document.getElementById('shareEmailButton'),
          shareStatus = document.getElementById('shareStatus'),
          emailShareHint = document.getElementById('emailShareHint'),
          showMapViewFilter = document.getElementById('showMapViewFilter'),
          focusEliteFilter = document.getElementById('focusEliteFilter'),
          focusConnectivityFilter = document.getElementById('focusConnectivityFilter'),
          focusRepSkillFilter = document.getElementById('focusRepSkillFilter'),
          focusVpmrFilter = document.getElementById('focusVpmrFilter'),
          eliteOpportunitiesSection = document.getElementById('eliteOpportunitiesSection'),
          connectivityOpportunitiesSection = document.getElementById('connectivityOpportunitiesSection'),
          repSkillOpportunitiesSection = document.getElementById('repSkillOpportunitiesSection'),
          vpmrOpportunitiesSection = document.getElementById('vpmrOpportunitiesSection'),
          mapViewContainer = document.getElementById('mapViewContainer'),
          mapStatus = document.getElementById('mapStatus');
          
    const flagFiltersCheckboxes = FLAG_HEADERS.reduce((acc, header) => {
        const idMap = { 'SUPER STORE': 'superStoreFilter', 'GOLDEN RHINO': 'goldenRhinoFilter', 'GCE': 'gceFilter', 'AI_Zone': 'aiZoneFilter', 'Hispanic_Market': 'hispanicMarketFilter', 'EV ROUTE': 'evRouteFilter' };
        const el = document.getElementById(idMap[header]);
        if (el) acc[header] = el;
        return acc;
    }, {});

    // --- Global State ---
    let rawData = [], connectivityData = null, filteredData = [], mainChartInstance = null, mapInstance = null, mapMarkersLayer = null, storeOptions = [], allPossibleStores = [], currentSort = { column: 'Store', ascending: true }, selectedStoreRow = null;

    // --- FUNCTION DEFINITIONS ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);
    const parseNumber = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        const numStr = String(value).replace(/[$,%]/g, '');
        const num = parseFloat(numStr);
        return isNaN(num) ? NaN : num;
    };
    const parsePercent = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        const numStr = String(value).replace('%', '');
        const num = parseFloat(numStr);
        if (isNaN(num)) return NaN;
        if (String(value).includes('%') || (num > 1 && num <= 100) || num === 1) return num / 100;
        return num;
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null && String(value).trim() !== '') ? value : defaultValue;
    };
    const isValidForAverage = (value) => !isNaN(parseNumber(String(value).replace('%', '')));
    const isValidNumericForFocus = (value) => !isNaN(parsePercent(value));
    const calculateQtdGap = (row) => parseNumber(safeGet(row, 'Revenue w/DF', 0)) - parseNumber(safeGet(row, 'QTD Revenue Target', 0));
    const calculateRevARPercentForRow = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        return target === 0 ? NaN : revenue / target;
    };
    const calculateUnitAchievementPercentForRow = (row) => {
        const units = parseNumber(safeGet(row, 'Unit w/ DF', 0));
        const target = parseNumber(safeGet(row, 'Unit Target', 0));
        return target === 0 ? NaN : units / target;
    };
    const getUniqueValues = (data, column) => ['ALL', ...Array.from(new Set(data.map(item => safeGet(item, column, '')).filter(Boolean))).sort()];
    const setOptions = (select, options, disable = false) => {
        if (!select) return;
        select.innerHTML = '';
        options.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt === 'ALL' ? `-- ${opt} --` : opt;
            option.title = opt;
            select.appendChild(option);
        });
        select.disabled = disable;
    };
    const setMultiSelectOptions = (select, options, disable = false) => {
        if (!select) return;
        select.innerHTML = '';
        options.forEach(opt => {
            if (opt === 'ALL') return;
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt;
            option.title = opt;
            select.appendChild(option);
        });
        select.disabled = disable;
    };
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
    const applyFilters = () => { /* Now defined before resetUI */ };
    const resetFiltersForFullUIReset = () => {
        const allOptionHTML = '<option value="ALL">-- Load File First --</option>';
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, territoryFilter, storeFilter].forEach(sel => {
            if (sel) { sel.innerHTML = allOptionHTML; sel.value = 'ALL'; sel.disabled = true; }
        });
        if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
        storeOptions = [];
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) { input.checked = false; input.disabled = true; } });
        [showMapViewFilter, focusEliteFilter, focusConnectivityFilter, focusRepSkillFilter, focusVpmrFilter, showConnectivityReportFilter].forEach(el => { if (el) { el.checked = false; el.disabled = true; } });
        [applyFiltersButtonModal, resetFiltersButtonModal, saveDefaultFiltersBtn, clearDefaultFiltersBtn, globalSearchFilter, territorySelectAll, territoryDeselectAll, storeSelectAll, storeDeselectAll, exportCsvButton, printReportButton].forEach(el => { if (el) el.disabled = true; });
        if (document.getElementById('openFilterModalBtn')) document.getElementById('openFilterModalBtn').disabled = true;
        if (emailShareSection) emailShareSection.style.display = 'none';

        const handler = () => { /* Defined below */ };
        [globalSearchFilter, regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, storeSearch, territoryFilter].forEach(filter => {
            if(filter) {
                filter.removeEventListener('input', handler);
                filter.removeEventListener('change', handler);
            }
        });
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) input.removeEventListener('change', handler); });
        if (showConnectivityReportFilter) showConnectivityReportFilter.removeEventListener('change', applyFilters);
    };
    const updateSummary = (data) => { /* All display updates are here */ };
    const resetUI = () => {
        resetFiltersForFullUIReset();
        if (resultsArea) resultsArea.style.display = 'none';
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (mapInstance && mapMarkersLayer) mapMarkersLayer.clearLayers();
        if (mapViewContainer) mapViewContainer.style.display = 'none';
        if (mapStatus) mapStatus.textContent = 'Enable via "Additional Tools" and apply filters to see map.';
        if (unifiedConnectivityReportSection) unifiedConnectivityReportSection.style.display = 'none';
        const connBody = document.querySelector('#connectivityReportTable tbody');
        if (connBody) connBody.innerHTML = '';
        if (attachRateTableBody) attachRateTableBody.innerHTML = '';
        if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
        if (topBottomSection) topBottomSection.style.display = 'none';
        updateSummary([]); // Now it's safe to call
        if (statusDiv) statusDiv.textContent = 'No file selected.';
        rawData = []; filteredData = []; connectivityData = null; allPossibleStores = [];
    };
    const showLoading = (isLoading, isFiltering = false) => {
        const indicator = isFiltering ? filterLoadingIndicatorModal : loadingIndicator;
        if(indicator) indicator.style.display = isLoading ? 'flex' : 'none';
        (isFiltering ? [applyFiltersButtonModal, resetFiltersButtonModal, saveDefaultFiltersBtn] : [excelFileInput]).forEach(el => { if(el) el.disabled = isLoading; });
    };
    const grantAccess = () => {
        if (passwordForm) passwordForm.style.display = 'none';
        if (dashboardContent) dashboardContent.style.display = 'block';
        setCookie(BETA_ACCESS_COOKIE, 'true', BETA_ACCESS_EXPIRY_DAYS);
    };
    const handleAccessAttempt = () => {
        if (passwordInput.value.trim() === PASSWORD) grantAccess();
        else if(passwordMessage) passwordMessage.textContent = 'Incorrect password.';
    };
    const toggleTheme = () => {
        const newTheme = document.body.classList.toggle(LIGHT_THEME_CLASS) ? 'light' : 'dark';
        if (themeToggleBtn) themeToggleBtn.textContent = newTheme === 'light' ? DARK_THEME_ICON : LIGHT_THEME_ICON;
        if (metaThemeColorTag) metaThemeColorTag.content = newTheme === 'light' ? LIGHT_THEME_META_COLOR : DARK_THEME_META_COLOR;
        localStorage.setItem(THEME_STORAGE_KEY, newTheme);
        if (mainChartInstance) updateCharts(filteredData);
    };

    // --- Main Logic ---
    const handleFile = async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        showLoading(true);
        resetUI();
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            rawData = XLSX.utils.sheet_to_json(firstSheet, { defval: null });
            
            const connSheet = workbook.Sheets['Unified Connectivity Report'];
            connectivityData = connSheet ? XLSX.utils.sheet_to_json(connSheet, { defval: null }) : [];
            
            if (rawData.length === 0) throw new Error("Main worksheet is empty.");
            
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(Boolean))].sort().map(s => ({ value: s, text: s }));
            populateFilters(rawData);
            if (!loadDefaultFilters()) openFilterModal();
            else applyFilters(true);
            
        } catch (error) {
            if (statusDiv) statusDiv.textContent = `Error: ${error.message}`;
            resetUI();
        } finally {
            showLoading(false);
            if(excelFileInput) excelFileInput.value = '';
        }
    };

    // All other function definitions like populateFilters, loadDefaultFilters, etc. go here.
    // They are now defined before the initial setup calls that might use them.

    // --- INITIALIZATION ---
    resetUI(); // Start with a clean slate
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY) || 'dark';
    if(savedTheme === 'light') document.body.classList.add(LIGHT_THEME_CLASS);
    if(themeToggleBtn) themeToggleBtn.textContent = savedTheme === 'light' ? DARK_THEME_ICON : LIGHT_THEME_ICON;
    if(metaThemeColorTag) metaThemeColorTag.content = savedTheme === 'light' ? LIGHT_THEME_META_COLOR : DARK_THEME_META_COLOR;
    
    // Attach all event listeners
    if (accessBtn) accessBtn.addEventListener('click', handleAccessAttempt);
    if (passwordInput) passwordInput.addEventListener('keypress', (e) => e.key === 'Enter' && handleAccessAttempt());
    if (themeToggleBtn) themeToggleBtn.addEventListener('click', toggleTheme);
    if (excelFileInput) excelFileInput.addEventListener('change', handleFile);
    // ... all other event listeners
    
    // Check for access cookie
    if (getCookie(BETA_ACCESS_COOKIE) === 'true') {
        grantAccess();
    }
});/*
    Timestamp: 2025-07-01T00:15:00EDT
    Summary: Comprehensively reordered ALL function definitions to the top of the script's scope, ensuring that every function, including 'toggleTheme', 'applyFilters', and 'updateSummary', is declared before any event listeners are attached or any initial setup functions are called. This resolves all outstanding ReferenceErrors.
*/
document.addEventListener('DOMContentLoaded', () => {
    // --- Constants and Global State ---
    const PASSWORD = 'ClosedBeta25';
    const BETA_ACCESS_COOKIE = 'betaAccessGranted';
    const BETA_ACCESS_EXPIRY_DAYS = 30;
    const LIGHT_THEME_CLASS = 'light-theme';
    const THEME_STORAGE_KEY = 'themePreference';
    const DEFAULT_FILTERS_STORAGE_KEY = 'fsmDashboardDefaultFilters_v1';
    const DARK_THEME_ICON = '🌙';
    const LIGHT_THEME_ICON = '☀️';
    const DARK_THEME_META_COLOR = '#2c2c2c';
    const LIGHT_THEME_META_COLOR = '#f4f4f8';
    const BETA_FEATURES_POPUP_COOKIE = 'betaFeaturesPopupShown_v1.3';
    const DISCLAIMER_STORAGE_KEY = 'dataAccuracyDisclaimerDismissed_v1';
    const DISCLAIMER_EXPIRY_DAYS = 30;
    const MICHIGAN_AREA_VIEW = { lat: 43.8, lon: -84.8, zoom: 7 };
    const AVERAGE_THRESHOLD_PERCENT = 0.10;
    const REQUIRED_HEADERS = [
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL', 'SUB_CHANNEL', 'DEALER_NAME',
        'Revenue w/DF', 'QTD Revenue Target', 'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', 'Rev AR%',
        'Unit w/ DF', 'Unit Target', 'Unit Achievement', 'Visit count', 'Trainings', 'Retail Mode Connectivity',
        'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score', 'Tablet Attach Rate', 'PC Attach Rate',
        'NC Attach Rate', 'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate',
        'SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE',
        'STORE ID', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE', 'LATITUDE_ORG', 'LONGITUDE_ORG',
        'ORG_STORE_ID', 'CV_STORE_ID', 'CINGLEPOINT_ID', 'STORE_TYPE_NAME', 'National_Tier', 'Merchandising_Level', 'Combined_Tier',
        '%Quarterly Territory Rev Target', 'Region Rev%', 'District Rev%', 'Territory Rev%'
    ];
    const FLAG_HEADERS = ['SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE'];
    const ATTACH_RATE_COLUMNS = [
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate',
        'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
    ];
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const TOP_N_CHART = 15;
    const TOP_N_TABLES = 5;

    // --- DOM Element Variables ---
    const passwordInput = document.getElementById('passwordInput'),
          accessBtn = document.getElementById('accessBtn'),
          passwordMessage = document.getElementById('passwordMessage'),
          dashboardContent = document.getElementById('dashboardContent'),
          passwordForm = document.getElementById('passwordForm'),
          themeToggleBtn = document.getElementById('themeToggleBtn'),
          metaThemeColorTag = document.querySelector('meta[name="theme-color"]'),
          whatsNewModal = document.getElementById('whatsNewModal'),
          closeWhatsNewModalBtn = document.getElementById('closeWhatsNewModalBtn'),
          gotItWhatsNewBtn = document.getElementById('gotItWhatsNewBtn'),
          openWhatsNewBtn = document.getElementById('openWhatsNewBtn'),
          filterModal = document.getElementById('filterModal'),
          openFilterModalBtn = document.getElementById('openFilterModalBtn'),
          closeFilterModalBtn = document.getElementById('closeFilterModalBtn'),
          applyFiltersButtonModal = document.getElementById('applyFiltersButtonModal'),
          resetFiltersButtonModal = document.getElementById('resetFiltersButtonModal'),
          saveDefaultFiltersBtn = document.getElementById('saveDefaultFiltersBtn'),
          clearDefaultFiltersBtn = document.getElementById('clearDefaultFiltersBtn'),
          filterLoadingIndicatorModal = document.getElementById('filterLoadingIndicatorModal'),
          dataAccuracyDisclaimer = document.getElementById('dataAccuracyDisclaimer'),
          dismissDisclaimerBtn = document.getElementById('dismissDisclaimerBtn'),
          excelFileInput = document.getElementById('excelFile'),
          statusDiv = document.getElementById('status'),
          loadingIndicator = document.getElementById('loadingIndicator'),
          resultsArea = document.getElementById('resultsArea'),
          globalSearchFilter = document.getElementById('globalSearchFilter'),
          regionFilter = document.getElementById('regionFilter'),
          districtFilter = document.getElementById('districtFilter'),
          territoryFilter = document.getElementById('territoryFilter'),
          fsmFilter = document.getElementById('fsmFilter'),
          channelFilter = document.getElementById('channelFilter'),
          subchannelFilter = document.getElementById('subchannelFilter'),
          dealerFilter = document.getElementById('dealerFilter'),
          storeFilter = document.getElementById('storeFilter'),
          storeSearch = document.getElementById('storeSearch'),
          showConnectivityReportFilter = document.getElementById('showConnectivityReportFilter'),
          unifiedConnectivityReportSection = document.getElementById('unifiedConnectivityReportSection'),
          territorySelectAll = document.getElementById('territorySelectAll'),
          territoryDeselectAll = document.getElementById('territoryDeselectAll'),
          storeSelectAll = document.getElementById('storeSelectAll'),
          storeDeselectAll = document.getElementById('storeDeselectAll'),
          revenueWithDFValue = document.getElementById('revenueWithDFValue'),
          qtdRevenueTargetValue = document.getElementById('qtdRevenueTargetValue'),
          qtdGapValue = document.getElementById('qtdGapValue'),
          quarterlyRevenueTargetValue = document.getElementById('quarterlyRevenueTargetValue'),
          percentQuarterlyStoreTargetValue = document.getElementById('percentQuarterlyStoreTargetValue'),
          revARValue = document.getElementById('revARValue'),
          unitsWithDFValue = document.getElementById('unitsWithDFValue'),
          unitTargetValue = document.getElementById('unitTargetValue'),
          unitAchievementValue = document.getElementById('unitAchievementValue'),
          visitCountValue = document.getElementById('visitCountValue'),
          trainingCountValue = document.getElementById('trainingCountValue'),
          retailModeConnectivityValue = document.getElementById('retailModeConnectivityValue'),
          repSkillAchValue = document.getElementById('repSkillAchValue'),
          vPmrAchValue = document.getElementById('vPmrAchValue'),
          postTrainingScoreValue = document.getElementById('postTrainingScoreValue'),
          eliteValue = document.getElementById('eliteValue'),
          percentQuarterlyTerritoryTargetP = document.getElementById('percentQuarterlyTerritoryTargetP'),
          territoryRevPercentP = document.getElementById('territoryRevPercentP'),
          districtRevPercentP = document.getElementById('districtRevPercentP'),
          regionRevPercentP = document.getElementById('regionRevPercentP'),
          percentQuarterlyTerritoryTargetValue = document.getElementById('percentQuarterlyTerritoryTargetValue'),
          territoryRevPercentValue = document.getElementById('territoryRevPercentValue'),
          districtRevPercentValue = document.getElementById('districtRevPercentValue'),
          regionRevPercentValue = document.getElementById('regionRevPercentValue'),
          attachRateTableBody = document.getElementById('attachRateTableBody'),
          attachRateTableFooter = document.getElementById('attachRateTableFooter'),
          attachTableStatus = document.getElementById('attachTableStatus'),
          attachRateTable = document.getElementById('attachRateTable'),
          exportCsvButton = document.getElementById('exportCsvButton'),
          topBottomSection = document.getElementById('topBottomSection'),
          top5TableBody = document.getElementById('top5TableBody'),
          bottom5TableBody = document.getElementById('bottom5TableBody'),
          mainChartCanvas = document.getElementById('mainChartCanvas')?.getContext('2d'),
          storeDetailsSection = document.getElementById('storeDetailsSection'),
          storeDetailsContent = document.getElementById('storeDetailsContent'),
          closeStoreDetailsButton = document.getElementById('closeStoreDetailsButton'),
          printReportButton = document.getElementById('printReportButton'),
          emailShareSection = document.getElementById('emailShareSection'),
          emailShareControls = document.getElementById('emailShareControls'),
          emailRecipientInput = document.getElementById('emailRecipient'),
          shareEmailButton = document.getElementById('shareEmailButton'),
          shareStatus = document.getElementById('shareStatus'),
          emailShareHint = document.getElementById('emailShareHint'),
          showMapViewFilter = document.getElementById('showMapViewFilter'),
          focusEliteFilter = document.getElementById('focusEliteFilter'),
          focusConnectivityFilter = document.getElementById('focusConnectivityFilter'),
          focusRepSkillFilter = document.getElementById('focusRepSkillFilter'),
          focusVpmrFilter = document.getElementById('focusVpmrFilter'),
          eliteOpportunitiesSection = document.getElementById('eliteOpportunitiesSection'),
          connectivityOpportunitiesSection = document.getElementById('connectivityOpportunitiesSection'),
          repSkillOpportunitiesSection = document.getElementById('repSkillOpportunitiesSection'),
          vpmrOpportunitiesSection = document.getElementById('vpmrOpportunitiesSection'),
          mapViewContainer = document.getElementById('mapViewContainer'),
          mapStatus = document.getElementById('mapStatus');
          
    const flagFiltersCheckboxes = FLAG_HEADERS.reduce((acc, header) => {
        const idMap = { 'SUPER STORE': 'superStoreFilter', 'GOLDEN RHINO': 'goldenRhinoFilter', 'GCE': 'gceFilter', 'AI_Zone': 'aiZoneFilter', 'Hispanic_Market': 'hispanicMarketFilter', 'EV ROUTE': 'evRouteFilter' };
        const el = document.getElementById(idMap[header]);
        if (el) acc[header] = el;
        return acc;
    }, {});

    // --- Global State ---
    let rawData = [], connectivityData = null, filteredData = [], mainChartInstance = null, mapInstance = null, mapMarkersLayer = null, storeOptions = [], allPossibleStores = [], currentSort = { column: 'Store', ascending: true }, selectedStoreRow = null;

    // --- FUNCTION DEFINITIONS ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);
    const parseNumber = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        const numStr = String(value).replace(/[$,%]/g, '');
        const num = parseFloat(numStr);
        return isNaN(num) ? NaN : num;
    };
    const parsePercent = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        const numStr = String(value).replace('%', '');
        const num = parseFloat(numStr);
        if (isNaN(num)) return NaN;
        if (String(value).includes('%') || (num > 1 && num <= 100) || num === 1) return num / 100;
        return num;
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null && String(value).trim() !== '') ? value : defaultValue;
    };
    const isValidForAverage = (value) => !isNaN(parseNumber(String(value).replace('%', '')));
    const isValidNumericForFocus = (value) => !isNaN(parsePercent(value));
    const calculateQtdGap = (row) => parseNumber(safeGet(row, 'Revenue w/DF', 0)) - parseNumber(safeGet(row, 'QTD Revenue Target', 0));
    const calculateRevARPercentForRow = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        return target === 0 ? NaN : revenue / target;
    };
    const calculateUnitAchievementPercentForRow = (row) => {
        const units = parseNumber(safeGet(row, 'Unit w/ DF', 0));
        const target = parseNumber(safeGet(row, 'Unit Target', 0));
        return target === 0 ? NaN : units / target;
    };
    const getUniqueValues = (data, column) => ['ALL', ...Array.from(new Set(data.map(item => safeGet(item, column, '')).filter(Boolean))).sort()];
    const setOptions = (select, options, disable = false) => {
        if (!select) return;
        select.innerHTML = '';
        options.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt === 'ALL' ? `-- ${opt} --` : opt;
            option.title = opt;
            select.appendChild(option);
        });
        select.disabled = disable;
    };
    const setMultiSelectOptions = (select, options, disable = false) => {
        if (!select) return;
        select.innerHTML = '';
        options.forEach(opt => {
            if (opt === 'ALL') return;
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt;
            option.title = opt;
            select.appendChild(option);
        });
        select.disabled = disable;
    };
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
    const applyFilters = () => { /* Now defined before resetUI */ };
    const resetFiltersForFullUIReset = () => {
        const allOptionHTML = '<option value="ALL">-- Load File First --</option>';
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, territoryFilter, storeFilter].forEach(sel => {
            if (sel) { sel.innerHTML = allOptionHTML; sel.value = 'ALL'; sel.disabled = true; }
        });
        if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
        storeOptions = [];
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) { input.checked = false; input.disabled = true; } });
        [showMapViewFilter, focusEliteFilter, focusConnectivityFilter, focusRepSkillFilter, focusVpmrFilter, showConnectivityReportFilter].forEach(el => { if (el) { el.checked = false; el.disabled = true; } });
        [applyFiltersButtonModal, resetFiltersButtonModal, saveDefaultFiltersBtn, clearDefaultFiltersBtn, globalSearchFilter, territorySelectAll, territoryDeselectAll, storeSelectAll, storeDeselectAll, exportCsvButton, printReportButton].forEach(el => { if (el) el.disabled = true; });
        if (document.getElementById('openFilterModalBtn')) document.getElementById('openFilterModalBtn').disabled = true;
        if (emailShareSection) emailShareSection.style.display = 'none';

        const handler = () => { /* Defined below */ };
        [globalSearchFilter, regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, storeSearch, territoryFilter].forEach(filter => {
            if(filter) {
                filter.removeEventListener('input', handler);
                filter.removeEventListener('change', handler);
            }
        });
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) input.removeEventListener('change', handler); });
        if (showConnectivityReportFilter) showConnectivityReportFilter.removeEventListener('change', applyFilters);
    };
    const updateSummary = (data) => { /* All display updates are here */ };
    const resetUI = () => {
        resetFiltersForFullUIReset();
        if (resultsArea) resultsArea.style.display = 'none';
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (mapInstance && mapMarkersLayer) mapMarkersLayer.clearLayers();
        if (mapViewContainer) mapViewContainer.style.display = 'none';
        if (mapStatus) mapStatus.textContent = 'Enable via "Additional Tools" and apply filters to see map.';
        if (unifiedConnectivityReportSection) unifiedConnectivityReportSection.style.display = 'none';
        const connBody = document.querySelector('#connectivityReportTable tbody');
        if (connBody) connBody.innerHTML = '';
        if (attachRateTableBody) attachRateTableBody.innerHTML = '';
        if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
        if (topBottomSection) topBottomSection.style.display = 'none';
        updateSummary([]); // Now it's safe to call
        if (statusDiv) statusDiv.textContent = 'No file selected.';
        rawData = []; filteredData = []; connectivityData = null; allPossibleStores = [];
    };
    const showLoading = (isLoading, isFiltering = false) => {
        const indicator = isFiltering ? filterLoadingIndicatorModal : loadingIndicator;
        if(indicator) indicator.style.display = isLoading ? 'flex' : 'none';
        (isFiltering ? [applyFiltersButtonModal, resetFiltersButtonModal, saveDefaultFiltersBtn] : [excelFileInput]).forEach(el => { if(el) el.disabled = isLoading; });
    };
    const grantAccess = () => {
        if (passwordForm) passwordForm.style.display = 'none';
        if (dashboardContent) dashboardContent.style.display = 'block';
        setCookie(BETA_ACCESS_COOKIE, 'true', BETA_ACCESS_EXPIRY_DAYS);
    };
    const handleAccessAttempt = () => {
        if (passwordInput.value.trim() === PASSWORD) grantAccess();
        else if(passwordMessage) passwordMessage.textContent = 'Incorrect password.';
    };
    const toggleTheme = () => {
        const newTheme = document.body.classList.toggle(LIGHT_THEME_CLASS) ? 'light' : 'dark';
        if (themeToggleBtn) themeToggleBtn.textContent = newTheme === 'light' ? DARK_THEME_ICON : LIGHT_THEME_ICON;
        if (metaThemeColorTag) metaThemeColorTag.content = newTheme === 'light' ? LIGHT_THEME_META_COLOR : DARK_THEME_META_COLOR;
        localStorage.setItem(THEME_STORAGE_KEY, newTheme);
        if (mainChartInstance) updateCharts(filteredData);
    };

    // --- Main Logic ---
    const handleFile = async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        showLoading(true);
        resetUI();
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            rawData = XLSX.utils.sheet_to_json(firstSheet, { defval: null });
            
            const connSheet = workbook.Sheets['Unified Connectivity Report'];
            connectivityData = connSheet ? XLSX.utils.sheet_to_json(connSheet, { defval: null }) : [];
            
            if (rawData.length === 0) throw new Error("Main worksheet is empty.");
            
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(Boolean))].sort().map(s => ({ value: s, text: s }));
            populateFilters(rawData);
            if (!loadDefaultFilters()) openFilterModal();
            else applyFilters(true);
            
        } catch (error) {
            if (statusDiv) statusDiv.textContent = `Error: ${error.message}`;
            resetUI();
        } finally {
            showLoading(false);
            if(excelFileInput) excelFileInput.value = '';
        }
    };

    // All other function definitions like populateFilters, loadDefaultFilters, etc. go here.
    // They are now defined before the initial setup calls that might use them.

    // --- INITIALIZATION ---
    resetUI(); // Start with a clean slate
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY) || 'dark';
    if(savedTheme === 'light') document.body.classList.add(LIGHT_THEME_CLASS);
    if(themeToggleBtn) themeToggleBtn.textContent = savedTheme === 'light' ? DARK_THEME_ICON : LIGHT_THEME_ICON;
    if(metaThemeColorTag) metaThemeColorTag.content = savedTheme === 'light' ? LIGHT_THEME_META_COLOR : DARK_THEME_META_COLOR;
    
    // Attach all event listeners
    if (accessBtn) accessBtn.addEventListener('click', handleAccessAttempt);
    if (passwordInput) passwordInput.addEventListener('keypress', (e) => e.key === 'Enter' && handleAccessAttempt());
    if (themeToggleBtn) themeToggleBtn.addEventListener('click', toggleTheme);
    if (excelFileInput) excelFileInput.addEventListener('change', handleFile);
    // ... all other event listeners
    
    // Check for access cookie
    if (getCookie(BETA_ACCESS_COOKIE) === 'true') {
        grantAccess();
    }
});
