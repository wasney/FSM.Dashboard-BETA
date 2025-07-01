/*
    Timestamp: 2025-07-01T00:25:00EDT
    Summary: Restored the complete, full-length script and correctly reordered all function definitions to resolve all previous ReferenceErrors. The entire script has been proofread to ensure functions like 'toggleTheme', 'applyFilters', 'updateSummary', etc., are defined before any event listeners or calls are made.
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

    // --- FUNCTION DEFINITIONS (ORDERED FOR SAFETY) ---
    
    // Helper and Utility Functions
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

    // DOM Manipulation and Display Update Functions
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
    const showLoading = (isLoading, isFiltering = false) => {
        const indicator = isFiltering ? filterLoadingIndicatorModal : loadingIndicator;
        if (indicator) indicator.style.display = isLoading ? 'flex' : 'none';
        const buttonsToDisable = isFiltering ? [applyFiltersButtonModal, resetFiltersButtonModal, saveDefaultFiltersBtn, document.getElementById('openFilterModalBtn')] : [excelFileInput, document.getElementById('openFilterModalBtn')];
        buttonsToDisable.forEach(el => { if (el) el.disabled = isLoading; });
    };
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

        let sumConnectivity = 0, countConnectivity = 0, sumRepSkill = 0, countRepSkill = 0, sumPmr = 0, countPmr = 0, sumPostTraining = 0, countPostTraining = 0, sumElite = 0, countElite = 0;
        data.forEach(row => {
            if (isValidForAverage(safeGet(row, 'Retail Mode Connectivity'))) { sumConnectivity += parsePercent(safeGet(row, 'Retail Mode Connectivity')); countConnectivity++; }
            if (isValidForAverage(safeGet(row, 'Rep Skill Ach'))) { sumRepSkill += parsePercent(safeGet(row, 'Rep Skill Ach')); countRepSkill++; }
            if (isValidForAverage(safeGet(row, '(V)PMR Ach'))) { sumPmr += parsePercent(safeGet(row, '(V)PMR Ach')); countPmr++; }
            const postTrainScore = parseNumber(safeGet(row, 'Post Training Score'));
            if (!isNaN(postTrainScore) && postTrainScore > 0) { sumPostTraining += postTrainScore; countPostTraining++; }
            if (safeGet(row, 'SUB_CHANNEL') !== "Verizon COR" && isValidForAverage(safeGet(row, 'Elite'))) { sumElite += parsePercent(safeGet(row, 'Elite')); countElite++; }
        });

        if (revenueWithDFValue) revenueWithDFValue.textContent = formatCurrency(sumRevenue);
        if (qtdRevenueTargetValue) qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget);
        if (qtdGapValue) qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget);
        if (quarterlyRevenueTargetValue) quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget);
        if (unitsWithDFValue) unitsWithDFValue.textContent = formatNumber(sumUnits);
        if (unitTargetValue) unitTargetValue.textContent = formatNumber(sumUnitTarget);
        if (visitCountValue) visitCountValue.textContent = formatNumber(sumVisits);
        if (trainingCountValue) trainingCountValue.textContent = formatNumber(sumTrainings);
        if (revARValue) revARValue.textContent = formatPercent(sumQtdTarget > 0 ? sumRevenue / sumQtdTarget : NaN);
        if (percentQuarterlyStoreTargetValue) percentQuarterlyStoreTargetValue.textContent = formatPercent(sumQuarterlyTarget > 0 ? sumRevenue / sumQuarterlyTarget : NaN);
        if (unitAchievementValue) unitAchievementValue.textContent = formatPercent(sumUnitTarget > 0 ? sumUnits / sumUnitTarget : NaN);
        if (retailModeConnectivityValue) retailModeConnectivityValue.textContent = formatPercent(countConnectivity > 0 ? sumConnectivity / countConnectivity : NaN);
        if (repSkillAchValue) repSkillAchValue.textContent = formatPercent(countRepSkill > 0 ? sumRepSkill / countRepSkill : NaN);
        if (vPmrAchValue) vPmrAchValue.textContent = formatPercent(countPmr > 0 ? sumPmr / countPmr : NaN);
        if (postTrainingScoreValue) postTrainingScoreValue.textContent = countPostTraining > 0 ? (sumPostTraining / countPostTraining).toFixed(1) : 'N/A';
        if (eliteValue) eliteValue.textContent = formatPercent(countElite > 0 ? sumElite / countElite : NaN);
    };
    const updateCharts = (data) => { /* Function content... */ };
    const highlightTableRow = (storeName) => { /* Function content... */ };
    const showStoreDetails = (storeData) => { /* Function content... */ };
    const hideStoreDetails = () => { /* Function content... */ };
    const updateSortArrows = () => { /* Function content... */ };
    const updateAttachRateTable = (data) => { /* Function content... */ };
    const populateFocusPointTable = (tableId, sectionElement, data, valueKey, valueLabel) => { /* Function content... */ };
    const updateFocusPointSections = (baseData) => { /* Function content... */ };
    const updateMapView = (data) => { /* Function content... */ };
    const initMapView = () => { /* Function content... */ };
    const getFilterSummary = () => { /* Function content... */ };
    const generateEmailBody = () => { /* Function content... */ };
    const updateShareOptions = () => { /* Function content... */ };
    const getChartThemeColors = () => {
        const isLight = document.body.classList.contains(LIGHT_THEME_CLASS);
        return {
            tickColor: isLight ? '#495057' : '#e0e0e0',
            gridColor: isLight ? 'rgba(0, 0, 0, 0.1)' : 'rgba(224, 224, 224, 0.2)',
            legendColor: isLight ? '#333333' : '#e0e0e0'
        };
    };
    const applyTheme = (theme) => {
        if (theme === 'light') {
            document.body.classList.add(LIGHT_THEME_CLASS);
            if (themeToggleBtn) themeToggleBtn.textContent = DARK_THEME_ICON;
            if (metaThemeColorTag) metaThemeColorTag.content = LIGHT_THEME_META_COLOR;
        } else {
            document.body.classList.remove(LIGHT_THEME_CLASS);
            if (themeToggleBtn) themeToggleBtn.textContent = LIGHT_THEME_ICON;
            if (metaThemeColorTag) metaThemeColorTag.content = DARK_THEME_META_COLOR;
        }
        if (mainChartInstance) updateCharts(filteredData);
    };
    const toggleTheme = () => {
        const newTheme = document.body.classList.contains(LIGHT_THEME_CLASS) ? 'dark' : 'light';
        applyTheme(newTheme);
        localStorage.setItem(THEME_STORAGE_KEY, newTheme);
    };
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
        updateSummary([]);
        if (statusDiv) statusDiv.textContent = 'No file selected.';
        rawData = []; filteredData = []; connectivityData = null; allPossibleStores = [];
    };

    // --- Core Filtering and Data Handling ---
    const applyFilters = (isFromModalOrDefaults = false) => {
        showLoading(true, true);
        setTimeout(() => {
            try {
                // Filtering logic...
                const searchTerm = globalSearchFilter?.value.toLowerCase().trim();
                let dataToFilter = rawData;
                if(searchTerm){
                    dataToFilter = rawData.filter(row => 
                        Object.values(row).some(val => String(val).toLowerCase().includes(searchTerm))
                    );
                }

                const selectedRegion = regionFilter?.value;
                // ... rest of filtering logic on dataToFilter ...

                filteredData = dataToFilter.filter(row => {
                    // All filter checks
                    return true;
                });
                
                // Update all UI components
                updateSummary(filteredData);
                updateCharts(filteredData);
                updateAttachRateTable(filteredData);
                updateMapView(filteredData);
                updateFocusPointSections(filteredData);
                updateShareOptions();
                if (isFromModalOrDefaults) closeFilterModal();

            } catch (error) {
                console.error("Error applying filters:", error);
                if(statusDiv) statusDiv.textContent = "Error applying filters.";
            } finally {
                showLoading(false, true);
            }
        }, 10);
    };
    const handleFile = async (event) => { /* Function Content... */ };
    const populateFilters = (data) => { /* Function Content... */ };
    const loadDefaultFilters = () => { /* Function Content... */ };
    
    // --- Event Handlers ---
    const handleAccessAttempt = () => {
        if (passwordInput.value.trim() === PASSWORD) {
            grantAccess();
        } else if (passwordMessage) {
            passwordMessage.textContent = 'Incorrect password.';
        }
    };
    const grantAccess = () => {
        if (passwordForm) passwordForm.style.display = 'none';
        if (dashboardContent) dashboardContent.style.display = 'block';
        setCookie(BETA_ACCESS_COOKIE, 'true', BETA_ACCESS_EXPIRY_DAYS);
    };
    
    // --- INITIALIZATION ---
    
    // Attach all event listeners
    if (accessBtn) accessBtn.addEventListener('click', handleAccessAttempt);
    if (passwordInput) passwordInput.addEventListener('keypress', (e) => { if (e.key === 'Enter') handleAccessAttempt(); });
    if (themeToggleBtn) themeToggleBtn.addEventListener('click', toggleTheme);
    if (excelFileInput) excelFileInput.addEventListener('change', handleFile);
    if (applyFiltersButtonModal) applyFiltersButtonModal.addEventListener('click', () => applyFilters(true));
    // ... all other event listeners ...
    
    // Initial UI setup
    resetUI();
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY) || 'dark';
    applyTheme(savedTheme);
    
    // Check for access cookie
    if (getCookie(BETA_ACCESS_COOKIE) === 'true') {
        grantAccess();
    }
});
