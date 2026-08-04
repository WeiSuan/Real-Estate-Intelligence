(function () {
    "use strict";

    const DEFAULT_EXCEL_PATH = "data/總表_202605.xlsx";
    const TIMELINE_FILE_PREFIX = "總表_";
    const TAIWAN_MAP_NAME = "taiwan-counties";    const COUNTY_ALIAS = {
        "臺北市": "台北市",
        "臺中市": "台中市",
        "臺南市": "台南市",
        "臺東縣": "台東縣",
        "桃園縣": "桃園市"
    };
    const COMMON_COUNTY_ORDER = [
        "台北市", "新北市", "桃園市", "台中市", "台南市", "高雄市",
        "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣",
        "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
        "台東縣", "澎湖縣", "金門縣", "連江縣"
    ];
    const FIELD = {
        group: "額度關係人",
        applicant: "申購人",
        caseId: "初審編號(歸案)",
        totalLimit: "總額度",
        limit: "額度",
        industry: "業別第一碼(擔保品)分類",
        activeLimit: "有效額度",
        disbursed: "已撥金額",
        principal: "剩餘本金",
        collateral: "擔保值"
    };

    const RISK_META = {
        safe: { label: "安全", color: "#16835f", className: "risk-safe" },
        normal: { label: "正常", color: "#1769aa", className: "risk-normal" },
        caution: { label: "注意", color: "#c46a12", className: "risk-caution" },
        danger: { label: "危險", color: "#b92d2b", className: "risk-danger" }
    };

    const LTV_BUCKETS = [
        { id: "0-60", label: "0-60%", min: 0, max: 0.6 },
        { id: "60-80", label: "60-80%", min: 0.6, max: 0.8 },
        { id: "80-100", label: "80-100%", min: 0.8, max: 1 },
        { id: "100+", label: "100%+", min: 1, max: Infinity }
    ];

    const COVERAGE_BUCKETS = [
        { id: "under-100", label: "≤ 100%", min: 0, max: 1 },
        { id: "100-120", label: "100-120%", min: 1, max: 1.2 },
        { id: "120-150", label: "120-150%", min: 1.2, max: 1.5 },
        { id: "150+", label: "≥ 150%", min: 1.5, max: Infinity }
    ];
    const MERGED_CASE_REASON = "合併案件";
    const MATRIX_META = {
        "0-60": { id: "0-60", label: "0-60%", className: "matrix-weak" },
        "60-80": { id: "60-80", label: "60-80%", className: "matrix-tight" },
        "80-100": { id: "80-100", label: "80-100%", className: "matrix-tight" },
        "100+": { id: "100+", label: "100%+", className: "matrix-strong" }
    };
    const SUPPLY_PRESSURE_META = {
        both: { id: "both", label: "第一象限-供需雙高警戒區", className: "supply-both" },
        unsold: { id: "unsold", label: "第二象限-建商激戰區", className: "supply-unsold" },
        vacancy: { id: "vacancy", label: "第四象限-存量去化區", className: "supply-vacancy" },
        low: { id: "low", label: "第三象限-剛需熱區", className: "supply-low" },
        unknown: { id: "unknown", label: "供給資料不足", className: "supply-unknown" }
    };
    const COLLATERAL_LEVEL_FILTERS = [
        { id: "0-60", label: "0-60%", description: "擔保覆蓋率低" },
        { id: "60-80", label: "60-80%", description: "擔保覆蓋率中低" },
        { id: "80-100", label: "80-100%", description: "擔保覆蓋率中高" },
        { id: "100+", label: "100%+", description: "擔保覆蓋率高" }
    ];
    const QUADRANT_FILTERS = [
        { id: "q1", supplyId: "both", label: "第一象限", description: "供需雙高警戒區" },
        { id: "q2", supplyId: "unsold", label: "第二象限", description: "建商激戰區" },
        { id: "q3", supplyId: "low", label: "第三象限", description: "剛需熱區" },
        { id: "q4", supplyId: "vacancy", label: "第四象限", description: "存量去化區" }
    ];
    const NEWS_STORAGE_KEY = "fmg_regional_news_v1";
    const TABLE_WIDTH_STORAGE_KEY = "fmg_table_widths_v1";
    const TABLE_WIDTH_LOCK_STORAGE_KEY = "fmg_table_width_locks_v1";
    const UI_MODE_STORAGE_KEY = "fmg_ui_mode_v1";
    const THEME_MODE_STORAGE_KEY = "fmg_theme_mode_v1";
    const CUSTOMER_MASK_STORAGE_KEY = "fmg_customer_mask_v1";
    const BAR_PALETTE_STORAGE_KEY = "fmg_bar_palette_v1";
    const EXPOSURE_BASIS_STORAGE_KEY = "fmg_exposure_basis_v1";
    const TIMELINE_MODE_STORAGE_KEY = "fmg_timeline_mode_v1";
    const SELECTED_PERIOD_STORAGE_KEY = "fmg_selected_period_v1";
    const COMPARE_PERIOD_STORAGE_KEY = "fmg_compare_period_v1";
    const TIMELINE_METRIC_STORAGE_KEY = "fmg_timeline_metric_v1";
    const FIXED_TIMELINE_METRIC = "principalBalance";
    const FRONT_PAGE_DATA_PATH = "data/front_page_data.xlsx";
    const LOAN_REMAIN_CAPITAL_SHEET = "loan_remain_capital";
    const LOAN_REMAIN_CAPITAL_AXIS_MAX = 40000000000;
    const LOAN_REMAIN_CAPITAL_STRUCTURE_AXIS_MAX = 0.2;
    const VACANCY_DATA_DATE_LABEL = window.VACANCY_DATA_DATE_LABEL || "資料日期未提供";
    const TIMELINE_CHANGE_MODE_STORAGE_KEY = "fmg_timeline_change_mode_v1";
    const AMOUNT_UNIT_STORAGE_KEY = "fmg_amount_unit_v1";
    const MAIN_PANEL_SLOTS_STORAGE_KEY = "fmg_main_panel_slots_v1";
    const RANK_TOPN_STORAGE_KEY = "fmg_rank_topn_v1";
    const DEVELOPER_MODE_STORAGE_KEY = "fmg_developer_mode_v1";
    const DEV_HIDDEN_BLOCKS_STORAGE_KEY = "fmg_dev_hidden_blocks_v1";
    const DISPLAY_NAME_STORAGE_KEY = "fmg_component_display_names_v1";
    const DISCONNECT_TIMELINE_FILTERS_STORAGE_KEY = "fmg_disconnect_timeline_filters_v1";
    const DEFAULT_DEV_HIDDEN_BLOCKS = [
        "timelineOverviewPanel",
        "timelineTrendPanel",
        "timelineCustomerChangePanel",
        "timelineFlowPanel"
    ];
    const FONT_SIZE_CONFIG_STORAGE_KEY = "fmg_font_size_config_v1";
    const DEFAULT_FONT_SIZE_CONFIG = { title: 16, body: 13, chart: 11, map: 12, family: "system" };
    const FONT_SIZE_LIMITS = {
        title: { min: 10, max: 28, fallback: DEFAULT_FONT_SIZE_CONFIG.title },
        body: { min: 10, max: 28, fallback: DEFAULT_FONT_SIZE_CONFIG.body },
        chart: { min: 9, max: 20, fallback: DEFAULT_FONT_SIZE_CONFIG.chart },
        map: { min: 9, max: 20, fallback: DEFAULT_FONT_SIZE_CONFIG.map }
    };
    const FONT_FAMILY_OPTIONS = {
        system: '"Noto Sans TC", "Microsoft JhengHei", "PingFang TC", system-ui, sans-serif',
        "microsoft-jhenghei": '"Microsoft JhengHei", "Noto Sans TC", "PingFang TC", system-ui, sans-serif',
        kai: '"DFKai-SB", "標楷體", "KaiTi", serif',
        "noto-sans-hk": '"Noto Sans HK", "Noto Sans TC", "Microsoft JhengHei", system-ui, sans-serif'
    };
    const RISK_NOTE_DEFINITION = "風險等級定義（現行）：安全 <70%、正常 70-85%、注意 85-100%、危險 >100% 或擔保值=0。";
    const DEVELOPER_HIDE_GROUPS = [
        {
            title: "主標題工具列",
            items: [
                { key: "kpiPanel", label: "區域指標面板" },
                { key: "filterPanel", label: "篩選條件" },
                { key: "countyDrawerPanel", label: "縣市彙總表" },
                { key: "caseDrawerPanel", label: "案件明細表" },
                { key: "newsPanel", label: "相關新聞區" }
            ]
        },
        {
            title: "當前面板",
            items: [
                { key: "mapPanel", label: "台澎金馬曝險地圖" },
                { key: "mainSlot1", label: "區域授信排行" },
                { key: "mainSlot2", label: "前十大授信戶" },
                { key: "mainSlot3", label: "供需四象限摘要" },
                { key: "mainSlot4", label: "LTV 區間分布" }
            ]
        },
        {
            title: "趨勢面板",
            items: [
                { key: "timelineOverviewPanel", label: "整體趨勢卡" },
                { key: "timelineTrendPanel", label: "趨勢直條圖" },
                { key: "timelineLoanPanel", label: "本金餘額結構趨勢" },
                { key: "timelineCountyRankPanel", label: "區域授信排行" },
                { key: "timelineCountyChangePanel", label: "區域變化排行" },
                { key: "timelineCustomerChangePanel", label: "TOP10 客戶變化排行" },
                { key: "timelineFlowPanel", label: "淨增加 / 淨減少概況" }
            ]
        },
        {
            title: "分析圖表",
            items: [
                { key: "analysisDrawerPanel", label: "進階分析圖表面板" },
                { key: "analysisIndustryPie", label: "產品/業別佔比" },
                { key: "analysisRiskPie", label: "風險等級分布" },
                { key: "analysisGroup", label: "前十大曝險集團" },
                { key: "analysisCountyRank", label: "區域曝險排行" },
                { key: "analysisTopCustomer", label: "前十大客戶曝險排行" },
                { key: "analysisQuadrantMatrix", label: "象限矩陣散佈圖" },
                { key: "analysisQuadrantSummary", label: "供需四象限摘要" },
                { key: "analysisLtv2", label: "LTV 區間分布" }
            ]
        }
    ];
    const BAR_PALETTE = {
        ocean: {
            countyRank: ["#3B82F6", "#2563EB"],
            group: ["#3B82F6", "#2563EB"],
            coverage: ["#60A5FA", "#2563EB"]
        },
        violet: {
            countyRank: ["#9b7dff", "#6153f7"],
            group: ["#9b7dff", "#6153f7"],
            coverage: ["#a994ff", "#7467f7"]
        },
        sunset: {
            countyRank: ["#ffc26d", "#ff8a26"],
            group: ["#ffc26d", "#ff8a26"],
            coverage: ["#ffd084", "#ff9c36"]
        },
        forest: {
            countyRank: ["#64dfad", "#1aae74"],
            group: ["#64dfad", "#1aae74"],
            coverage: ["#7ce6bb", "#28ba82"]
        }
    };

    const state = {
        rows: [],
        timelinePeriods: [],
        periodDatasets: new Map(),
        frontPageLoanRemainCapitalRows: [],
        dataSourceText: "",
        cityFields: new Map(),
        displayCountyByKey: new Map(),
        geoNameByCountyKey: new Map(),
        districtDetailsByCase: new Map(),
        districtOptionsByCounty: new Map(),
        districtValidation: null,
        districtSearch: "",
        activeDistrictCounty: "",
        districtPopupCounty: "",
        districtPopupAnchor: null,
        mapClickTimer: null,
        hoveredDistrictCounty: "",
        countyRankClickTimer: null,
        countyRankDrilldownCounty: "",
        timelineRankDrilldownCounty: "",
        vacancyByCounty: new Map(),
        vacancyByDistrict: new Map(),
        geoJson: null,
        tableSearch: "",
        filters: {
            county: [],
            districts: {},
            industry: [],
            industryLabel: "",
            ltvBucket: [],
            coverageBucket: "",
            anomalyMode: "",
            collateralLevels: [],
            quadrants: [],
            group: "",
            customer: "",
            riskLevel: ""
        },
        showTop5Labels: true,
        mainGridMode: "default",
        newsItems: loadNewsItems(),
        showNewsForm: false,
        editingNewsId: "",
        uiMode: loadUiMode(),
        themeMode: loadThemeMode(),
        customerMasked: loadCustomerMask(),
        tableWidthLocks: loadTableWidthLocks(),
        tableColumnWidths: loadTableColumnWidths(),
        selectedBarPalette: loadBarPalette(),
        exposureBasis: loadExposureBasis(),
        amountUnit: loadAmountUnit(),
        timelineMode: loadTimelineMode(),
        selectedPeriodId: loadSelectedPeriodId(),
        comparePeriodId: loadComparePeriodId(),
        timelineMetric: FIXED_TIMELINE_METRIC,
        timelineChangeMode: loadTimelineChangeMode(),
        fontSizeConfig: loadFontSizeConfig(),
        mainPanelChartSlots: loadMainPanelChartSlots(),
        rankTopN: loadRankTopN(),
        developerMode: loadDeveloperMode(),
        devHiddenBlocks: loadDeveloperHiddenBlocks(),
        componentDisplayNames: loadComponentDisplayNames(),
        displacementMeasure: {
            active: false,
            dragging: false,
            cancelled: false,
            startX: 0,
            startY: 0,
            currentX: 0,
            currentY: 0,
            startPageX: 0,
            startPageY: 0,
            currentPageX: 0,
            currentPageY: 0,
            targetBox: null,
            endTargetBox: null
        },
        analysisLtvBasis: "count",
        disconnectTimelineFilters: loadDisconnectTimelineFilters(),
        caseSort: { column: 8, direction: "desc" },
        countySort: { column: 0, direction: "asc" },
        coverageMode: "case"
    };

    const charts = {};

    const els = {};
    const elementIds = [
        "dataStatus", "filterSummary", "dataSourceBadge", "dependencyNotice", "countyFilter", "countyFilterSummary", "industryFilter", "industryFilterSummary",
        "ltvFilter", "ltvFilterSummary", "anomalyFilter", "matrixCollateralFilters", "matrixQuadrantFilters", "clearFilters",
        "kpiPanelTitle", "kpiPrincipal", "kpiApprovedUnused", "kpiHousehold", "kpiCollateral", "kpiCases",
        "kpiVacancyHomes", "kpiUnsoldHomes", "kpiAvgLtv", "kpiAvgCoverage",
        "currentRegionBadge", "capturePagePngBtn", "mapChart", "kinmenInsetBtn", "ltvBucketChart", "coverageChartLegacy", "concentrationList",
        "districtPanel", "districtTitle", "districtSummary", "districtSearch", "districtSelectAllBtn", "districtClearBtn", "districtBackBtn", "districtList",
        "activeFiltersBar", "activeFilterChips", "refreshFiltersBtn", "uiModeDefaultBtn", "uiModeIosBtn", "themeLightBtn", "themeDarkBtn", "themeToggleBtn", "exposurePrincipalBtn", "exposureHouseholdBtn", "timelineExposurePrincipalBtn", "timelineExposureHouseholdBtn", "rankExposurePrincipalBtn", "rankExposureHouseholdBtn", "amountUnitToggleBtn", "viewModeToggleBtn", "selectedPeriodSelect", "comparePeriodSelect", "barPaletteSelect", "fontFamilySelect", "titleFontSizeInput", "bodyFontSizeInput", "chartFontSizeInput", "mapFontSizeInput", "resetFontSizeConfigBtn", "customerMaskToggle", "developerModeToggle", "timelineFilterLinkToggle", "developerPanel", "developerHiddenList", "resetComponentDisplayNamesBtn",
        "caseFilterChips", "countyFilterChips",
        "toggleNewsPanel", "newsPanel", "newsRegionBadge", "toggleNewsFormBtn", "newsForm", "cancelNewsFormBtn", "newsCountySelect", "newsTitleInput", "newsDateInput", "newsUrlInput", "newsList",
        "tableSearch", "clearTableSearch", "exportCaseDetailExcel", "exportCountySummaryExcel", "toggleCaseTable", "toggleCountySummary",
        "tableSummary", "riskRows", "countySummaryRows", "caseTableWrap", "countySummaryWrap",
        "openCaseDrawer", "closeCaseDrawer", "caseDrawerPanel", "drawerBackdrop", "top5Toggle",
        "filterPanel", "closeFilterPanel",
        "toggleCountyDrawer", "closeCountyDrawer", "countyDrawerPanel",
        "lockCaseColumns", "lockCountyColumns",
        "toggleToolbarBtn",
        "toggleAnalysisDrawer", "closeAnalysisDrawer", "analysisDrawerPanel",
        "mainPanelChartPicker", "riskDefinitionNote",
        "timelinePanel", "timelineCompareBadge", "timelineMetricBadge", "timelineTrendCards", "timelineTrendChart", "loanRemainCapitalChart", "timelineCountyRankTitle", "timelineCountyRankBackBtn", "timelineCountyRankChart", "timelineCountyChangeChart", "timelineCustomerChangeChart", "timelineChangeModeSwitch", "timelineFlowCards", "timelineFlowList",
        "industryPieChart", "riskPieChart",
        "topCustomerChart", "quadrantMatrixChart", "quadrantSummaryChart", "analysisLtvBasisToggle",
        "ltvBucketChart2",
        "analysisGroupChart", "analysisCountyRankChart",
        "mainSlotTitle1", "mainSlotTitle2", "mainSlotTitle3", "mainSlotTitle4", "mainSlotMarketMeta3", "countyRankBackBtn",
        "mainSlotBadge1", "mainSlotBadge2", "mainSlotBadge3", "mainSlotBadge4",
        "mainSlotCoverageToggle1", "mainSlotCoverageToggle2", "mainSlotCoverageToggle3", "mainSlotCoverageToggle4",
        "mainSlotLtvBasisToggle4",
        "mainSlotChart1", "mainSlotChart2", "mainSlotChart3", "mainSlotChart4",
        "displacementMeasureOverlay", "displacementMeasureSvg", "displacementMeasureLine", "displacementMeasureStart", "displacementMeasureEnd", "displacementMeasureHud", "displacementMeasureDx", "displacementMeasureDy", "displacementMeasureDirection", "displacementMeasureStartCoord", "displacementMeasureCurrentCoord", "displacementMeasureTargetBox", "displacementMeasureEndTargetBox", "displacementMeasureTransform"
    ];

    document.addEventListener("DOMContentLoaded", init);

    async function init() {
        cacheElements();
        applyThemeMode();
        applyUiMode();
        applyFontSize();
        applyToolbarState();
        syncExposureBasisControls();
        syncAmountUnitControl();
        bindUiEvents();
        initAltDisplacementMeasure();

        if (!window.XLSX || !window.echarts) {
            showNotice("找不到離線套件。請確認 libs/xlsx.full.min.js 與 libs/echarts.min.js 已放在專案資料夾。");
            return;
        }

        initCharts();
        initSizeHoverHints();
        await loadGeoJson();
        await loadWorkbook();
        await loadFrontPageLoanRemainCapital();
        prepareVacancyRows();
        populateFilters();
        populateNewsCountyOptions();
        render();
        initResizableTables();
        initTableSorting();
        applyStoredTableWidths();
        syncTableWidthLockButtons();
        initCollapsedSections();
        bindChartEvents();
        initDashboardResizeObserver();
        updateStickyOffsets();

        window.addEventListener("resize", debounce(function () {
            updateStickyOffsets();
            fitMainSlotChartsToContainers();
            Object.values(charts).forEach(function (chart) {
                if (chart && typeof chart.resize === "function") {
                    chart.resize();
                }
            });
        }, 160));
    }

    function cacheElements() {
        elementIds.forEach(function (id) {
            els[id] = document.getElementById(id);
        });
    }

    function bindFontSizeInput(input, key) {
        if (!input) {
            return;
        }
        const limits = FONT_SIZE_LIMITS[key] || FONT_SIZE_LIMITS.body;
        input.min = String(limits.min);
        input.max = String(limits.max);
        input.step = "1";
        input.value = String(clampFontSizeValue(key, state.fontSizeConfig && state.fontSizeConfig[key]));
        input.addEventListener("change", function () {
            const nextConfig = normalizeFontSizeConfig(state.fontSizeConfig);
            nextConfig[key] = clampFontSizeValue(key, input.value);
            state.fontSizeConfig = normalizeFontSizeConfig(nextConfig);
            input.value = String(state.fontSizeConfig[key]);
            saveFontSizeConfig();
            applyFontSize();
            render();
            animateChartsResize();
        });
    }

    function bindFontFamilySelect(select) {
        if (!select) {
            return;
        }
        select.value = normalizeFontFamily(state.fontSizeConfig && state.fontSizeConfig.family);
        select.addEventListener("change", function () {
            const nextConfig = normalizeFontSizeConfig(state.fontSizeConfig);
            nextConfig.family = normalizeFontFamily(select.value);
            state.fontSizeConfig = normalizeFontSizeConfig(nextConfig);
            select.value = state.fontSizeConfig.family;
            saveFontSizeConfig();
            applyFontSize();
            render();
            animateChartsResize();
        });
    }

    function bindUiEvents() {
        bindFilterDropdownFloatBehavior();
        els.countyFilter.addEventListener("change", function (event) {
            const checkbox = event.target.closest("[data-county-filter]");
            if (!checkbox) {
                return;
            }
            setArrayFilterValue("county", checkbox.value, checkbox.checked);
            render();
        });
        if (els.districtSearch) {
            els.districtSearch.addEventListener("input", function () {
                state.districtSearch = this.value;
                renderDistrictPanel();
            });
        }
        if (els.districtSelectAllBtn) {
            els.districtSelectAllBtn.addEventListener("click", function () {
                selectAllActiveCountyDistricts();
                syncFilterControls();
                render();
            });
        }
        if (els.districtClearBtn) {
            els.districtClearBtn.addEventListener("click", function () {
                clearActiveCountyDistricts();
                syncFilterControls();
                render();
            });
        }
        if (els.districtBackBtn) {
            els.districtBackBtn.addEventListener("click", function () {
                closeDistrictPopup();
            });
        }
        if (els.districtList) {
            els.districtList.addEventListener("change", function (event) {
                const checkbox = event.target.closest("[data-district-filter]");
                if (!checkbox) {
                    return;
                }
                setDistrictSelection(checkbox.getAttribute("data-county"), checkbox.value, checkbox.checked);
                render();
            });
            els.districtList.addEventListener("mouseover", function (event) {
                const row = event.target.closest("[data-district-hover-county]");
                if (!row) {
                    return;
                }
                state.hoveredDistrictCounty = normalizeCountyName(row.getAttribute("data-district-hover-county"));
                renderMap(getFilteredRows());
            });
            els.districtList.addEventListener("mouseout", function (event) {
                if (!event.target.closest("[data-district-hover-county]")) {
                    return;
                }
                state.hoveredDistrictCounty = "";
                renderMap(getFilteredRows());
            });
        }
        els.industryFilter.addEventListener("change", function (event) {
            const checkbox = event.target.closest("[data-industry-filter]");
            if (!checkbox) {
                return;
            }
            setArrayFilterValue("industry", checkbox.value, checkbox.checked);
            state.filters.industryLabel = "";
            render();
        });
        els.ltvFilter.addEventListener("change", function (event) {
            const checkbox = event.target.closest("[data-ltv-filter]");
            if (!checkbox) {
                return;
            }
            setArrayFilterValue("ltvBucket", checkbox.value, checkbox.checked);
            render();
        });
        els.anomalyFilter.addEventListener("change", function () {
            state.filters.anomalyMode = this.value;
            render();
        });
        els.clearFilters.addEventListener("click", function () {
            state.filters = createEmptyFilters();
            state.districtSearch = "";
            state.countyRankDrilldownCounty = "";
            state.timelineRankDrilldownCounty = "";
            syncFilterControls();
            render();
        });
        els.matrixCollateralFilters.addEventListener("change", function (event) {
            const checkbox = event.target.closest("[data-collateral-level]");
            if (!checkbox) {
                return;
            }
            toggleArrayFilter(state.filters.collateralLevels, checkbox.getAttribute("data-collateral-level"));
            render();
        });
        els.matrixQuadrantFilters.addEventListener("change", function (event) {
            const checkbox = event.target.closest("[data-quadrant-filter]");
            if (!checkbox) {
                return;
            }
            toggleArrayFilter(state.filters.quadrants, checkbox.getAttribute("data-quadrant-filter"));
            render();
        });
        els.tableSearch.addEventListener("input", function () {
            state.tableSearch = this.value;
            renderTable(getFilteredRows());
        });
        els.clearTableSearch.addEventListener("click", function () {
            state.tableSearch = "";
            els.tableSearch.value = "";
            renderTable(getFilteredRows());
        });
        if (els.exportCaseDetailExcel) {
            els.exportCaseDetailExcel.addEventListener("click", exportCaseDetailExcel);
        }
        if (els.exportCountySummaryExcel) {
            els.exportCountySummaryExcel.addEventListener("click", exportCountySummaryExcel);
        }
        els.toggleCaseTable.addEventListener("click", function () {
            toggleSection("caseTableWrap", "toggleCaseTable");
        });
        els.toggleCountySummary.addEventListener("click", function () {
            toggleSection("countySummaryWrap", "toggleCountySummary");
        });
        if (els.toggleCountyDrawer && els.countyDrawerPanel) {
            els.toggleCountyDrawer.addEventListener("click", function () {
                toggleCountyDrawer();
            });
        }
        if (els.closeCountyDrawer && els.countyDrawerPanel) {
            els.closeCountyDrawer.addEventListener("click", function () {
                setCountyDrawer(false);
            });
        }
        if (els.top5Toggle) {
            els.top5Toggle.addEventListener("change", function () {
                state.showTop5Labels = !!this.checked;
                render();
            });
        }
        if (els.toggleNewsPanel) {
            els.toggleNewsPanel.addEventListener("click", function () {
                state.mainGridMode = state.mainGridMode === "default" ? "news" : "default";
                renderMainGridMode();
            });
        }
        if (els.refreshFiltersBtn) {
            els.refreshFiltersBtn.addEventListener("click", function () {
                render();
            });
        }
        if (els.uiModeDefaultBtn) {
            els.uiModeDefaultBtn.addEventListener("click", function () {
                state.uiMode = "default";
                saveUiMode();
                applyUiMode();
            });
        }
        if (els.uiModeIosBtn) {
            els.uiModeIosBtn.addEventListener("click", function () {
                state.uiMode = "ios";
                saveUiMode();
                applyUiMode();
            });
        }
        if (els.themeLightBtn) {
            els.themeLightBtn.addEventListener("click", function () {
                state.themeMode = "light";
                saveThemeMode();
                applyThemeMode();
                render();
                animateChartsResize();
            });
        }
        if (els.themeDarkBtn) {
            els.themeDarkBtn.addEventListener("click", function () {
                state.themeMode = "dark";
                saveThemeMode();
                applyThemeMode();
                render();
                animateChartsResize();
            });
        }
        if (els.themeToggleBtn) {
            els.themeToggleBtn.addEventListener("click", function () {
                state.themeMode = state.themeMode === "dark" ? "light" : "dark";
                saveThemeMode();
                applyThemeMode();
                render();
                animateChartsResize();
            });
        }
        if (els.capturePagePngBtn) {
            els.capturePagePngBtn.addEventListener("click", function () {
                captureFullPagePng();
            });
        }
        if (els.barPaletteSelect) {
            els.barPaletteSelect.value = state.selectedBarPalette;
            els.barPaletteSelect.addEventListener("change", function () {
                const next = text(this.value);
                state.selectedBarPalette = BAR_PALETTE[next] ? next : "ocean";
                saveBarPalette();
                render();
            });
        }
        if (els.exposurePrincipalBtn) {
            els.exposurePrincipalBtn.addEventListener("click", function () {
                state.exposureBasis = "principal";
                saveExposureBasis();
                syncExposureBasisControls();
                render();
            });
        }
        if (els.exposureHouseholdBtn) {
            els.exposureHouseholdBtn.addEventListener("click", function () {
                state.exposureBasis = "household";
                saveExposureBasis();
                syncExposureBasisControls();
                render();
            });
        }
        if (els.timelineExposurePrincipalBtn) {
            els.timelineExposurePrincipalBtn.addEventListener("click", function () {
                state.exposureBasis = "principal";
                saveExposureBasis();
                syncExposureBasisControls();
                render();
            });
        }
        if (els.timelineExposureHouseholdBtn) {
            els.timelineExposureHouseholdBtn.addEventListener("click", function () {
                state.exposureBasis = "household";
                saveExposureBasis();
                syncExposureBasisControls();
                render();
            });
        }
        if (els.rankExposurePrincipalBtn) {
            els.rankExposurePrincipalBtn.addEventListener("click", function () {
                state.exposureBasis = "principal";
                saveExposureBasis();
                syncExposureBasisControls();
                render();
            });
        }
        if (els.rankExposureHouseholdBtn) {
            els.rankExposureHouseholdBtn.addEventListener("click", function () {
                state.exposureBasis = "household";
                saveExposureBasis();
                syncExposureBasisControls();
                render();
            });
        }
        if (els.timelineCountyRankBackBtn) {
            els.timelineCountyRankBackBtn.addEventListener("click", function () {
                closeTimelineRankDrilldown();
            });
        }
        if (els.countyRankBackBtn) {
            els.countyRankBackBtn.addEventListener("click", function () {
                closeCountyRankDrilldown();
            });
        }
        if (els.amountUnitToggleBtn) {
            els.amountUnitToggleBtn.addEventListener("click", function (event) {
                const unitButton = event.target.closest("[data-amount-unit]");
                const nextUnit = unitButton ? unitButton.getAttribute("data-amount-unit") : (state.amountUnit === "yi" ? "k" : "yi");
                state.amountUnit = nextUnit === "yi" ? "yi" : "k";
                saveAmountUnit();
                syncAmountUnitControl();
                renderDistrictPanel();
                render();
                animateChartsResize();
            });
        }
        if (els.kinmenInsetBtn) {
            els.kinmenInsetBtn.addEventListener("click", function () {
                const county = normalizeCountyName("金門縣");
                selectCountyFromMap(county);
                syncFilterControls();
                render();
            });
        }
        if (els.viewModeToggleBtn) {
            els.viewModeToggleBtn.addEventListener("click", function (event) {
                const modeButton = event.target.closest("[data-view-mode]");
                const nextMode = modeButton ? modeButton.getAttribute("data-view-mode") : (state.timelineMode === "timeline" ? "snapshot" : "timeline");
                state.timelineMode = nextMode === "timeline" ? "timeline" : "snapshot";
                saveTimelineMode();
                syncTimelineControls();
                render();
                animateChartsResize();
            });
        }
        if (els.selectedPeriodSelect) {
            els.selectedPeriodSelect.addEventListener("change", function () {
                state.selectedPeriodId = text(this.value);
                state.timelineRankDrilldownCounty = "";
                ensureComparePeriod();
                applySelectedPeriodDataset();
                saveSelectedPeriodId();
                saveComparePeriodId();
                populatePeriodControls();
                render();
            });
        }
        if (els.comparePeriodSelect) {
            els.comparePeriodSelect.addEventListener("change", function () {
                state.comparePeriodId = text(this.value);
                saveComparePeriodId();
                render();
            });
        }
        if (els.timelineChangeModeSwitch) {
            els.timelineChangeModeSwitch.addEventListener("click", function (event) {
                const button = event.target.closest("[data-value]");
                if (!button) {
                    return;
                }
                const next = text(button.getAttribute("data-value"));
                state.timelineChangeMode = next === "decrease" ? "decrease" : "increase";
                saveTimelineChangeMode();
                syncTimelineControls();
                renderTimelineDashboard();
            });
        }
        bindFontFamilySelect(els.fontFamilySelect);
        bindFontSizeInput(els.titleFontSizeInput, "title");
        bindFontSizeInput(els.bodyFontSizeInput, "body");
        bindFontSizeInput(els.chartFontSizeInput, "chart");
        bindFontSizeInput(els.mapFontSizeInput, "map");
        if (els.resetFontSizeConfigBtn) {
            els.resetFontSizeConfigBtn.addEventListener("click", function () {
                resetFontSizeConfig();
            });
        }
        if (els.customerMaskToggle) {
            syncCustomerMaskToggle();
            els.customerMaskToggle.addEventListener("change", function () {
                state.customerMasked = !!this.checked;
                saveCustomerMask();
                syncCustomerMaskToggle();
                render();
            });
        }
        if (els.developerModeToggle) {
            els.developerModeToggle.addEventListener("change", function () {
                state.developerMode = !!this.checked;
                saveDeveloperMode();
                syncDeveloperMode();
                scheduleChartReflow();
            });
        }
        if (els.timelineFilterLinkToggle) {
            els.timelineFilterLinkToggle.addEventListener("change", function () {
                state.disconnectTimelineFilters = !!this.checked;
                saveDisconnectTimelineFilters();
                syncDeveloperMode();
                if (state.timelineMode === "timeline") {
                    render();
                }
            });
        }
        if (els.developerHiddenList) {
            els.developerHiddenList.addEventListener("change", function (event) {
                const checkbox = event.target.closest("input[type='checkbox'][data-dev-hide-option]");
                if (!checkbox) {
                    return;
                }
                const key = text(checkbox.getAttribute("data-dev-hide-option"));
                if (!key) {
                    return;
                }
                if (checkbox.checked) {
                    state.devHiddenBlocks.add(key);
                } else {
                    state.devHiddenBlocks.delete(key);
                }
                saveDeveloperHiddenBlocks();
                syncDeveloperMode();
                scheduleChartReflow();
            });
        }
        if (els.developerDisplayNameList) {
            els.developerDisplayNameList.addEventListener("input", function (event) {
                const input = event.target.closest("input[data-display-name-key]");
                if (!input) {
                    return;
                }
                updateComponentDisplayName(text(input.getAttribute("data-display-name-key")), input.value);
            });
            els.developerDisplayNameList.addEventListener("click", function (event) {
                const button = event.target.closest("button[data-display-name-reset]");
                if (!button) {
                    return;
                }
                resetComponentDisplayName(text(button.getAttribute("data-display-name-reset")));
            });
        }
        if (els.resetComponentDisplayNamesBtn) {
            els.resetComponentDisplayNamesBtn.addEventListener("click", function () {
                resetAllComponentDisplayNames();
            });
        }
        if (els.newsForm) {
            els.newsForm.addEventListener("submit", function (event) {
                event.preventDefault();
                const county = text(els.newsCountySelect.value);
                const title = text(els.newsTitleInput.value);
                const date = text(els.newsDateInput.value);
                const url = text(els.newsUrlInput.value);
                if (!county || !title || !date || !url) {
                    showNotice("請完整填寫新聞區域、標題、日期與連結。");
                    return;
                }
                if (state.editingNewsId) {
                    state.newsItems = state.newsItems.map(function (item) {
                        if (item.id !== state.editingNewsId) {
                            return item;
                        }
                        return {
                            id: item.id,
                            county: county,
                            title: title,
                            date: date,
                            url: url
                        };
                    });
                } else {
                    state.newsItems.unshift({
                        id: `news_${Date.now()}`,
                        county: county,
                        title: title,
                        date: date,
                        url: url
                    });
                }
                saveNewsItems();
                resetNewsFormFields();
                state.showNewsForm = false;
                state.editingNewsId = "";
                syncNewsFormVisibility();
                renderNewsPanel();
            });
        }
        if (els.toggleNewsFormBtn) {
            els.toggleNewsFormBtn.addEventListener("click", function () {
                state.showNewsForm = !state.showNewsForm;
                if (state.showNewsForm && !state.editingNewsId && !text(els.newsDateInput.value)) {
                    els.newsDateInput.value = new Date().toISOString().slice(0, 10);
                }
                syncNewsFormVisibility();
            });
        }
        if (els.cancelNewsFormBtn) {
            els.cancelNewsFormBtn.addEventListener("click", function () {
                state.showNewsForm = false;
                state.editingNewsId = "";
                resetNewsFormFields();
                syncNewsFormVisibility();
            });
        }
        if (els.newsList) {
            els.newsList.addEventListener("click", function (event) {
                const editButton = event.target.closest("[data-edit-news-id]");
                if (editButton) {
                    const id = editButton.getAttribute("data-edit-news-id");
                    const item = state.newsItems.find(function (row) { return row.id === id; });
                    if (item) {
                        state.editingNewsId = item.id;
                        state.showNewsForm = true;
                        els.newsCountySelect.value = item.county;
                        els.newsTitleInput.value = item.title;
                        els.newsDateInput.value = item.date || "";
                        els.newsUrlInput.value = item.url;
                        syncNewsFormVisibility();
                    }
                    return;
                }
                const deleteButton = event.target.closest("[data-delete-news-id]");
                if (!deleteButton) {
                    return;
                }
                const id = deleteButton.getAttribute("data-delete-news-id");
                state.newsItems = state.newsItems.filter(function (item) { return item.id !== id; });
                saveNewsItems();
                if (state.editingNewsId === id) {
                    state.editingNewsId = "";
                    resetNewsFormFields();
                    syncNewsFormVisibility();
                }
                renderNewsPanel();
            });
        }
        if (els.lockCaseColumns) {
            els.lockCaseColumns.addEventListener("click", function () {
                state.tableWidthLocks.case = !state.tableWidthLocks.case;
                persistTableColumnWidths();
                syncTableWidthLockButtons();
            });
        }
        if (els.lockCountyColumns) {
            els.lockCountyColumns.addEventListener("click", function () {
                state.tableWidthLocks.county = !state.tableWidthLocks.county;
                persistTableColumnWidths();
                syncTableWidthLockButtons();
            });
        }
        if (els.closeFilterPanel && els.filterPanel) {
            els.closeFilterPanel.addEventListener("click", function () {
                setFilterDrawer(false);
            });
        }
        if (els.openCaseDrawer && els.caseDrawerPanel) {
            els.openCaseDrawer.addEventListener("click", function () {
                toggleCaseDrawer();
            });
        }
        if (els.closeCaseDrawer && els.caseDrawerPanel) {
            els.closeCaseDrawer.addEventListener("click", function () {
                setCaseDrawer(false);
            });
        }
        if (els.toggleAnalysisDrawer) {
            els.toggleAnalysisDrawer.addEventListener("click", function () {
                toggleAnalysisDrawer();
            });
        }
        if (els.closeAnalysisDrawer) {
            els.closeAnalysisDrawer.addEventListener("click", function () {
                setAnalysisDrawer(false);
            });
        }
        if (els.drawerBackdrop) {
            els.drawerBackdrop.addEventListener("click", function () {
                setCaseDrawer(false);
                setFilterDrawer(false);
                setCountyDrawer(false);
                setAnalysisDrawer(false);
            });
        }
        els.riskRows.addEventListener("click", function (event) {
            const button = event.target.closest("[data-customer-filter]");
            if (!button) {
                return;
            }
            const customer = button.getAttribute("data-customer-filter") || "";
            state.filters.customer = state.filters.customer === customer ? "" : customer;
            render();
        });
        if (els.caseTableWrap) {
            els.caseTableWrap.addEventListener("click", function (event) {
                if (event.target.closest(".resize-handle")) {
                    return;
                }
                const th = event.target.closest("th[data-sort-table='case']");
                if (!th) {
                    return;
                }
                toggleTableSort("case", toNumber(th.getAttribute("data-sort-index")));
                renderTable(getFilteredRows());
            });
        }
        if (els.countySummaryWrap) {
            els.countySummaryWrap.addEventListener("click", function (event) {
                if (event.target.closest(".resize-handle")) {
                    return;
                }
                const th = event.target.closest("th[data-sort-table='county']");
                if (!th) {
                    return;
                }
                toggleTableSort("county", toNumber(th.getAttribute("data-sort-index")));
                renderCountySummaryTable(getFilteredRows());
            });
        }
        ["activeFilterChips", "caseFilterChips", "countyFilterChips"].forEach(function (key) {
            if (!els[key]) {
                return;
            }
            els[key].addEventListener("click", function (event) {
                const removeButton = event.target.closest("[data-remove-filter]");
                if (!removeButton) {
                    return;
                }
                removeFilterByChip(removeButton);
                syncFilterControls();
                render();
            });
        });
        document.querySelectorAll(".slot-coverage-toggle").forEach(function (toggleEl) {
            toggleEl.addEventListener("click", function (event) {
                const btn = event.target.closest(".segmented-btn");
                if (!btn) {
                    return;
                }
                const nextMode = btn.getAttribute("data-value");
                if (state.coverageMode === nextMode) {
                    return;
                }
                state.coverageMode = nextMode;
                document.querySelectorAll(".slot-coverage-toggle .segmented-btn").forEach(function (el) {
                    el.classList.toggle("active", el.getAttribute("data-value") === nextMode);
                });
                render();
            });
        });
        document.querySelectorAll(".ltv-basis-toggle").forEach(function (toggleEl) {
            toggleEl.addEventListener("click", function (event) {
                const btn = event.target.closest("[data-ltv-basis]");
                if (!btn) {
                    return;
                }
                const nextBasis = btn.getAttribute("data-ltv-basis") === "amount" ? "amount" : "count";
                if (state.analysisLtvBasis === nextBasis) {
                    return;
                }
                state.analysisLtvBasis = nextBasis;
                renderAnalysisCharts(getFilteredRows());
                renderMainPanelCharts(getFilteredRows());
            });
        });
        if (els.mainPanelChartPicker) {
            els.mainPanelChartPicker.addEventListener("change", function (event) {
                const checkbox = event.target.closest("input[type='checkbox'][data-main-chart-id]");
                if (!checkbox) {
                    return;
                }
                const chartId = text(checkbox.getAttribute("data-main-chart-id"));
                const selected = state.mainPanelChartSlots.slice();
                const existingIndex = selected.indexOf(chartId);
                if (checkbox.checked) {
                    if (existingIndex === -1) {
                        if (selected.length >= 4) {
                            checkbox.checked = false;
                            showNotice("主面板最多只能選 4 個圖表。");
                            return;
                        }
                        selected.push(chartId);
                    }
                } else if (existingIndex !== -1) {
                    selected.splice(existingIndex, 1);
                }
                state.mainPanelChartSlots = normalizeMainPanelChartSlots(selected);
                saveMainPanelChartSlots();
                renderMainPanelChartPicker();
                render();
                animateChartsResize();
            });
        }
        document.addEventListener("click", function (event) {
            if (state.districtPopupCounty && els.districtPanel && !event.target.closest("#districtPanel") && !event.target.closest("#mapChart") && !event.target.closest("#amountUnitToggleBtn")) {
                closeDistrictPopup();
            }
            const topNButton = event.target.closest("[data-topn-target]");
            if (!topNButton) {
                return;
            }
            const target = topNButton.getAttribute("data-topn-target");
            const current = state.rankTopN[target] === 10 ? 10 : 5;
            state.rankTopN[target] = current === 5 ? 10 : 5;
            saveRankTopN();
            render();
            animateChartsResize();
        });

        if (els.toggleToolbarBtn) {
            els.toggleToolbarBtn.addEventListener("click", function () {
                const shouldOpen = !document.body.classList.contains("toolbar-open");
                document.body.classList.toggle("toolbar-open", shouldOpen);
                localStorage.setItem("toolbarState", shouldOpen ? "expanded" : "collapsed");
                syncToolbarControl();
                animateChartsResize();
            });
        }
    }

    function initCharts() {
        // 註冊一個自訂字型主題，讓圖表文字跟隨全站字體設定。
        if (window.echarts && typeof window.echarts.registerTheme === "function") {
            window.echarts.registerTheme("custom_font", {
                textStyle: {
                    fontFamily: getAppFontFamily(),
                    fontSize: getFSD(12)
                }
            });
        }

        var initOpts = { renderer: "svg" };
        charts.map = createDashboardChart(els.mapChart, initOpts);
        charts.mainSlots = [els.mainSlotChart1, els.mainSlotChart2, els.mainSlotChart3, els.mainSlotChart4].map(function (el) {
            return createDashboardChart(el, initOpts);
        });
        if (els.ltvBucketChart) charts.ltv = createDashboardChart(els.ltvBucketChart, initOpts);
        if (els.coverageChartLegacy) charts.coverage = createDashboardChart(els.coverageChartLegacy, initOpts);
        if (els.timelineTrendChart) charts.timelineTrend = createDashboardChart(els.timelineTrendChart, initOpts);
        if (els.loanRemainCapitalChart) charts.loanRemainCapital = createDashboardChart(els.loanRemainCapitalChart, initOpts);
        if (els.timelineCountyRankChart) charts.timelineCountyRank = createDashboardChart(els.timelineCountyRankChart, initOpts);
        if (els.timelineCountyChangeChart) charts.timelineCountyChange = createDashboardChart(els.timelineCountyChangeChart, initOpts);
        if (els.timelineCustomerChangeChart) charts.timelineCustomerChange = createDashboardChart(els.timelineCustomerChangeChart, initOpts);
        charts.analysisInited = false;
    }

    function initAnalysisCharts() {
        if (charts.analysisInited) return;
        var initOpts = { renderer: "svg" };
        if (els.analysisGroupChart) charts.group = createDashboardChart(els.analysisGroupChart, initOpts);
        if (els.analysisCountyRankChart) charts.countyRank = createDashboardChart(els.analysisCountyRankChart, initOpts);
        if (els.industryPieChart) charts.industryPie = createDashboardChart(els.industryPieChart, initOpts);
        if (els.riskPieChart) charts.riskPie = createDashboardChart(els.riskPieChart, initOpts);
        if (els.topCustomerChart) charts.topCustomer = createDashboardChart(els.topCustomerChart, initOpts);
        if (els.quadrantMatrixChart) charts.quadrantMatrix = createDashboardChart(els.quadrantMatrixChart, initOpts);
        if (els.quadrantSummaryChart) charts.quadrantSummary = createDashboardChart(els.quadrantSummaryChart, initOpts);
        if (els.ltvBucketChart2) charts.ltv2 = createDashboardChart(els.ltvBucketChart2, initOpts);
        charts.analysisInited = true;
        bindAnalysisChartEvents();
    }

    function createDashboardChart(el, initOpts) {
        if (!el) {
            return null;
        }
        const chart = echarts.init(el, "custom_font", initOpts);
        patchChartSetOption(chart);
        return chart;
    }

    function patchChartSetOption(chart) {
        if (!chart || chart.__fontFamilyPatched) {
            return;
        }
        const originalSetOption = chart.setOption.bind(chart);
        chart.setOption = function (option) {
            applyChartFontFamily(option);
            return originalSetOption.apply(chart, arguments);
        };
        chart.__fontFamilyPatched = true;
    }

    function applyChartFontFamily(option) {
        if (!option || typeof option !== "object") {
            return option;
        }
        const family = getAppFontFamily();
        const styleKeys = new Set(["textStyle", "label", "axisLabel", "nameTextStyle"]);
        option.textStyle = Object.assign({}, option.textStyle || {}, { fontFamily: family });

        function visit(value, key) {
            if (!value || typeof value !== "object") {
                return;
            }
            if (Array.isArray(value)) {
                value.forEach(function (item) {
                    visit(item, key);
                });
                return;
            }
            if (styleKeys.has(key) || value.fontSize || value.fontWeight || value.color) {
                value.fontFamily = family;
            }
            if (key === "style" && (value.text || value.font || value.fontSize)) {
                value.fontFamily = family;
            }
            Object.keys(value).forEach(function (childKey) {
                if (typeof value[childKey] !== "function") {
                    visit(value[childKey], childKey);
                }
            });
        }

        Object.keys(option).forEach(function (key) {
            visit(option[key], key);
        });
        return option;
    }

    function bindChartEvents() {
        charts.map.off("click");
        charts.map.off("dblclick");
        charts.map.on("click", function (params) {
            const county = normalizeCountyName(params.name || "");
            if (!county || (params.event && params.event.event && params.event.event.detail > 1)) {
                return;
            }
            clearMapClickTimer();
            state.mapClickTimer = window.setTimeout(function () {
                selectCountyFromMap(county);
                syncFilterControls();
                render();
                state.mapClickTimer = null;
            }, 220);
        });
    }

    function bindAnalysisChartEvents() {
        if (charts.group) {
            charts.group.off("click");
            charts.group.on("click", function (params) {
                const group = params.data && params.data.rawName ? params.data.rawName : (params.name || "");
                state.filters.group = state.filters.group === group ? "" : group;
                render();
            });
        }

        if (charts.countyRank) {
            charts.countyRank.off("click");
            charts.countyRank.off("dblclick");
            charts.countyRank.on("click", function (params) {
                const county = normalizeCountyName(params.name || "");
                if (state.countyRankDrilldownCounty) {
                    if (params.event && params.event.event && params.event.event.detail > 1) {
                        return;
                    }
                    const district = (params.data && params.data.rawName) || params.name || "";
                    toggleDistrictFromCountyRank(state.countyRankDrilldownCounty, district);
                    return;
                }
                if (!county || state.countyRankDrilldownCounty || (params.event && params.event.event && params.event.event.detail > 1)) {
                    return;
                }
                clearCountyRankClickTimer();
                state.countyRankClickTimer = window.setTimeout(function () {
                    toggleArrayFilterValue("county", county);
                    syncFilterControls();
                    render();
                    state.countyRankClickTimer = null;
                }, 220);
            });
            charts.countyRank.on("dblclick", function (params) {
                handleCountyRankDrilldownDblclick(params);
            });
        }

        if (charts.timelineCountyRank) {
            charts.timelineCountyRank.off("dblclick");
            charts.timelineCountyRank.on("dblclick", function (params) {
                if (state.timelineRankDrilldownCounty) {
                    return;
                }
                const county = normalizeCountyName((params.data && params.data.rawName) || params.name || "");
                if (!canDrilldownCountyRank(county) || !(state.districtOptionsByCounty.get(county) || []).length) {
                    return;
                }
                state.timelineRankDrilldownCounty = county;
                closeDistrictPopup();
                renderTimelineDashboard();
            });
            const zr = charts.timelineCountyRank.getZr && charts.timelineCountyRank.getZr();
            if (zr && !charts.timelineCountyRank.__rankBlankDblclickBound) {
                zr.on("dblclick", function (event) {
                    if (state.timelineRankDrilldownCounty && !event.target) {
                        closeTimelineRankDrilldown();
                    }
                });
                charts.timelineCountyRank.__rankBlankDblclickBound = true;
            }
        }

        if (charts.industryPie) {
            charts.industryPie.off("click");
            charts.industryPie.on("click", function (params) {
                const name = params.name || "";
                state.filters.industry = [];
                state.filters.industryLabel = state.filters.industryLabel === name ? "" : name;
                syncFilterControls();
                render();
            });
        }

        if (charts.riskPie) {
            charts.riskPie.off("click");
            charts.riskPie.on("click", function (params) {
                if (params.data && params.data.key) {
                    const key = params.data.key;
                    state.filters.riskLevel = state.filters.riskLevel === key ? "" : key;
                    render();
                }
            });
        }

        if (charts.topCustomer) {
            charts.topCustomer.off("click");
            charts.topCustomer.on("click", function (params) {
                const rawName = params.data && params.data.rawName ? params.data.rawName : (params.name || "");
                if (rawName) {
                    state.filters.customer = state.filters.customer === rawName ? "" : rawName;
                    render();
                }
            });
        }

        if (charts.quadrantMatrix) {
            charts.quadrantMatrix.off("click");
            charts.quadrantMatrix.on("click", function (params) {
                if (params.data && params.data.raw && params.data.raw.applicant) {
                    const customer = params.data.raw.applicant;
                    state.filters.customer = state.filters.customer === customer ? "" : customer;
                    render();
                }
            });
        }

        if (charts.quadrantSummary) {
            charts.quadrantSummary.off("click");
            charts.quadrantSummary.on("click", function (params) {
                const quadrantId = params.data && params.data.qid ? params.data.qid : "";
                if (quadrantId) {
                    toggleArrayFilter(state.filters.quadrants, quadrantId);
                    syncFilterControls();
                    render();
                }
            });
        }

        if (charts.ltv2) {
            charts.ltv2.off("click");
            charts.ltv2.on("click", function (params) {
                if (params.data && params.data.bucketId) {
                    const bucketId = params.data.bucketId;
                    toggleArrayFilterValue("ltvBucket", bucketId);
                    syncFilterControls();
                    render();
                }
            });
        }

    }

    async function loadGeoJson() {
        if (window.TAIWAN_COUNTIES_GEOJSON) {
            state.geoJson = window.TAIWAN_COUNTIES_GEOJSON;
            registerMap();
            return;
        }

        try {
            const response = await fetch("geo/taiwan-counties.geojson");
            if (!response.ok) {
                throw new Error("GeoJSON not found");
            }
            const arrayBuffer = await response.arrayBuffer();
            const decoder = new TextDecoder("utf-8");
            const jsonText = decoder.decode(arrayBuffer);
            state.geoJson = JSON.parse(jsonText);
            registerMap();
        } catch (error) {
            showNotice("GeoJSON load failed. Please verify geo/taiwan-counties.js or geo/taiwan-counties.geojson exists.");
        }
    }

    function registerMap() {
        if (!state.geoJson || !charts.map) {
            return;
        }
        state.geoNameByCountyKey.clear();
        (state.geoJson.features || []).forEach(function (feature) {
            const props = feature.properties || {};
            const geoName = props.COUNTYNAME || props.name || "";
            const countyKey = normalizeCountyName(geoName);
            if (countyKey) {
                state.geoNameByCountyKey.set(countyKey, geoName);
            }
        });
        echarts.registerMap(TAIWAN_MAP_NAME, state.geoJson);
    }

    async function loadWorkbook() {
        state.timelinePeriods = loadTimelineManifest();
        state.periodDatasets.clear();
        const loadedPeriods = [];
        const failedFiles = [];
        for (const period of state.timelinePeriods) {
            try {
                const loaded = await loadWorkbookForPeriod(period);
                const dataset = buildPeriodDataset(period, loaded.rows, loaded.sourceText);
                dataset.districtRows = loaded.districtRows || [];
                prepareDistrictDataset(dataset);
                state.periodDatasets.set(period.periodId, dataset);
                loadedPeriods.push(period);
                if (!state.dataSourceText || loaded.sourceType === "sample") {
                    state.dataSourceText = loaded.sourceText;
                }
            } catch (error) {
                failedFiles.push(period.file);
            }
        }
        if (!loadedPeriods.length) {
            throw new Error("找不到時間軸資料清單");
        }
        state.timelinePeriods = loadedPeriods;
        const latestPeriodId = state.timelinePeriods[state.timelinePeriods.length - 1].periodId;
        if (!state.periodDatasets.has(state.selectedPeriodId)) {
            state.selectedPeriodId = latestPeriodId;
        }
        ensureComparePeriod();
        applySelectedPeriodDataset();
        populatePeriodControls();
        saveSelectedPeriodId();
        saveComparePeriodId();
        const failedHint = failedFiles.length ? `，略過 ${failedFiles.length} 個無法讀取檔案` : "";
        setStatus("已自動載入資料", `已載入 ${state.timelinePeriods.length} 期資料${failedHint}`);
    }

    function loadTimelineManifest() {
        const manifest = Array.isArray(window.TIMELINE_DATASETS) ? window.TIMELINE_DATASETS : [];
        const normalized = manifest.map(function (item) {
            const periodId = normalizeTimelinePeriodId(item.periodId || inferPeriodIdFromFile(item.file));
            return {
                periodId: periodId,
                label: text(item.label || formatTimelinePeriodLabel(periodId)),
                type: text(item.type || "month"),
                file: text(item.file || buildTimelineFilePath(periodId) || DEFAULT_EXCEL_PATH),
                inlineDefault: !!item.inlineDefault
            };
        }).filter(function (item) {
            return item.periodId && item.file;
        }).sort(function (a, b) {
            return a.periodId.localeCompare(b.periodId, "zh-Hant");
        });
        if (normalized.length) {
            return normalized;
        }
        return [{
            periodId: normalizeTimelinePeriodId(window.DEFAULT_EXCEL_PERIOD_ID) || "2026-05",
            label: formatTimelinePeriodLabel(window.DEFAULT_EXCEL_PERIOD_ID || "202605"),
            type: "month",
            file: DEFAULT_EXCEL_PATH,
            inlineDefault: true
        }];
    }

    async function loadWorkbookForPeriod(period) {
        let arrayBuffer = null;
        let sourceType = "excel";
        let sourceText = "資料來源：" + period.file;
        const sampleDataset = getTimelineSampleDataset(period.periodId);

        if (sampleDataset && sampleDataset.base64) {
            arrayBuffer = base64ToArrayBuffer(sampleDataset.base64);
            sourceType = "sample";
            sourceText = "資料來源：sample-xlsx.js 快取（" + (sampleDataset.fileName || period.file) + "）";
        }

        if (!arrayBuffer) {
            try {
                const fetchUrl = period.file + "?_ts=" + Date.now();
                const response = await fetch(fetchUrl, { cache: "no-store" });
                if (!response.ok) {
                    throw new Error("無法載入 Excel：" + period.file);
                }
                arrayBuffer = await response.arrayBuffer();
            } catch (error) {
                const samplePeriodId = normalizeTimelinePeriodId(window.DEFAULT_EXCEL_PERIOD_ID);
                const canUseSample = window.DEFAULT_EXCEL_BASE64 && (!samplePeriodId || samplePeriodId === period.periodId || period.inlineDefault);
                if (!canUseSample) {
                    throw error;
                }
                arrayBuffer = base64ToArrayBuffer(window.DEFAULT_EXCEL_BASE64);
                sourceType = "sample";
                sourceText = "資料來源：sample-xlsx.js 快取" + (window.DEFAULT_EXCEL_FILE_NAME ? "（" + window.DEFAULT_EXCEL_FILE_NAME + "）" : "");
            }
        }

        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const sheet = workbook.Sheets["工作表1"] || workbook.Sheets.Sheet1 || workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: 0 })
            .map(cleanRow)
            .map(scaleMoneyFieldsToK);
        const allCitiesSheet = workbook.Sheets.All_Cities || workbook.Sheets["All_Cities"];
        const districtRows = allCitiesSheet
            ? XLSX.utils.sheet_to_json(allCitiesSheet, { defval: 0 })
                .map(cleanRow)
                .map(scaleDistrictMoneyFieldsToK)
            : [];
        return {
            rows: rows,
            districtRows: districtRows,
            sourceType: sourceType,
            sourceText: sourceText
        };
    }

    async function loadFrontPageLoanRemainCapital() {
        state.frontPageLoanRemainCapitalRows = [];
        try {
            const response = await fetch(FRONT_PAGE_DATA_PATH + "?_ts=" + Date.now(), { cache: "no-store" });
            if (!response.ok) {
                return;
            }
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: "array" });
            const sheet = workbook.Sheets[LOAN_REMAIN_CAPITAL_SHEET];
            if (!sheet) {
                return;
            }
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
            state.frontPageLoanRemainCapitalRows = rows.slice(1).map(function (row) {
                return {
                    label: formatExcelDateAsYearMonth(row[0]),
                    loanRemainCapital: parseLooseNumber(row[1]),
                    loanRemainCapitalStructure: parseRatioValue(row[2])
                };
            }).filter(function (item) {
                return item.label && item.loanRemainCapital > 0;
            }).sort(function (a, b) {
                return a.label.localeCompare(b.label, "zh-Hant");
            });
        } catch (_error) {
            state.frontPageLoanRemainCapitalRows = [];
        }
    }

    function parseLooseNumber(value) {
        if (typeof value === "number") {
            return Number.isFinite(value) ? value : 0;
        }
        const normalized = text(value).replace(/,/g, "").replace(/%/g, "").trim();
        const parsed = Number(normalized);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function parseRatioValue(value) {
        if (typeof value === "number") {
            if (!Number.isFinite(value)) {
                return 0;
            }
            return value > 1 ? value / 100 : value;
        }
        const raw = text(value).trim();
        const parsed = parseLooseNumber(raw);
        if (!parsed) {
            return 0;
        }
        return raw.indexOf("%") !== -1 || parsed > 1 ? parsed / 100 : parsed;
    }

    function formatExcelDateAsYearMonth(value) {
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
            return formatYearMonthFromParts(value.getFullYear(), value.getMonth() + 1);
        }
        if (typeof value === "number" && Number.isFinite(value)) {
            const utcTime = Date.UTC(1899, 11, 30) + Math.round(value) * 86400000;
            const date = new Date(utcTime);
            return formatYearMonthFromParts(date.getUTCFullYear(), date.getUTCMonth() + 1);
        }
        const raw = text(value).trim();
        const match = raw.match(/^(\d{4})[-/](\d{1,2})[-/]\d{1,2}$/);
        if (match) {
            return formatYearMonthFromParts(Number(match[1]), Number(match[2]));
        }
        return "";
    }

    function formatYearMonthFromParts(year, month) {
        if (!year || !month) {
            return "";
        }
        return String(year) + "年" + String(month).padStart(2, "0") + "月";
    }

    function getTimelineSampleDataset(periodId) {
        const datasets = Array.isArray(window.TIMELINE_SAMPLE_DATASETS) ? window.TIMELINE_SAMPLE_DATASETS : [];
        const target = normalizeTimelinePeriodId(periodId);
        for (let i = 0; i < datasets.length; i += 1) {
            const item = datasets[i] || {};
            const itemPeriodId = normalizeTimelinePeriodId(item.periodId || inferPeriodIdFromFile(item.fileName || ""));
            if (itemPeriodId && itemPeriodId === target && item.base64) {
                return {
                    periodId: itemPeriodId,
                    fileName: text(item.fileName),
                    base64: text(item.base64)
                };
            }
        }
        return null;
    }

    function normalizeTimelinePeriodId(value) {
        const raw = text(value);
        const compactMatch = raw.match(/^(\d{4})(\d{2})$/);
        if (compactMatch) {
            return compactMatch[1] + "-" + compactMatch[2];
        }
        const dashedMatch = raw.match(/^(\d{4})-(\d{2})$/);
        return dashedMatch ? raw : "";
    }

    function inferPeriodIdFromFile(file) {
        const match = text(file).match(/總表_(\d{6})\.xlsx$/i);
        return match ? match[1] : "";
    }

    function buildTimelineFilePath(periodId) {
        const normalized = normalizeTimelinePeriodId(periodId);
        if (!normalized) {
            return "";
        }
        return "data/" + TIMELINE_FILE_PREFIX + normalized.replace("-", "") + ".xlsx";
    }

    function formatTimelinePeriodLabel(periodId) {
        const normalized = normalizeTimelinePeriodId(periodId);
        if (!normalized) {
            return text(periodId || "目前資料");
        }
        return normalized.slice(0, 4) + "年" + normalized.slice(5, 7) + "月";
    }

    function buildPeriodDataset(period, workbookRows, sourceText) {
        return {
            periodId: period.periodId,
            label: period.label,
            type: period.type,
            sourceText: sourceText || "資料來源：" + period.file,
            rawRows: workbookRows,
            districtRows: [],
            districtDetailsByCase: new Map(),
            districtOptionsByCounty: new Map(),
            districtValidation: null,
            baseRows: [],
            filteredRows: [],
            summary: null,
            countySummary: [],
            customerSummary: [],
            groupSummary: []
        };
    }

    function applySelectedPeriodDataset() {
        const dataset = state.periodDatasets.get(state.selectedPeriodId) || state.periodDatasets.get(state.timelinePeriods[state.timelinePeriods.length - 1].periodId);
        state.rows = dataset ? dataset.rawRows : [];
        state.dataSourceText = dataset ? dataset.sourceText : "";
        state.districtDetailsByCase = dataset ? dataset.districtDetailsByCase : new Map();
        state.districtOptionsByCounty = dataset ? dataset.districtOptionsByCounty : new Map();
        state.districtValidation = dataset ? dataset.districtValidation : null;
        pruneDistrictSelections();
        detectCityFields();
    }

    function base64ToArrayBuffer(base64) {
        const binary = atob(base64);
        const length = binary.length;
        const bytes = new Uint8Array(length);
        for (let i = 0; i < length; i += 1) {
            bytes[i] = binary.charCodeAt(i);
        }
        return bytes.buffer;
    }

    function cleanRow(row) {
        const cleaned = {};
        Object.keys(row).forEach(function (key) {
            const cleanKey = String(key).replace(/\ufeff|\u200b|\xa0/g, "").trim();
            const value = row[key];
            if (value === "" || value === "-" || value === null || typeof value === "undefined") {
                cleaned[cleanKey] = 0;
            } else if (typeof value === "string") {
                cleaned[cleanKey] = value.replace(/\ufeff|\u200b|\xa0/g, "").trim();
            } else {
                cleaned[cleanKey] = value;
            }
        });
        return cleaned;
    }

    function scaleMoneyFieldsToK(row) {
        const scaled = Object.assign({}, row);
        const moneyKeys = [FIELD.totalLimit, FIELD.limit, FIELD.activeLimit, FIELD.disbursed, FIELD.principal, FIELD.collateral];
        moneyKeys.forEach(function (key) {
            if (Object.prototype.hasOwnProperty.call(scaled, key)) {
                scaled[key] = toNumber(scaled[key]) / 1000;
            }
        });
        Object.keys(scaled).forEach(function (key) {
            if (/_remain$|_worth$|_capital$/u.test(key)) {
                scaled[key] = toNumber(scaled[key]) / 1000;
            }
        });
        return scaled;
    }

    function scaleDistrictMoneyFieldsToK(row) {
        const scaled = Object.assign({}, row);
        ["value_worth", "value_capital", "value_remain"].forEach(function (key) {
            if (Object.prototype.hasOwnProperty.call(scaled, key)) {
                scaled[key] = toNumber(scaled[key]) / 1000;
            }
        });
        return scaled;
    }

    function detectCityFields() {
        state.cityFields.clear();
        state.displayCountyByKey.clear();
        const sample = state.rows[0] || {};
        Object.keys(sample).forEach(function (key) {
            if (!key.endsWith("_remain") && !key.endsWith("_worth") && !key.endsWith("_capital")) {
                return;
            }
            const rawCounty = key.replace(/_(remain|worth|capital)$/u, "");
            const countyKey = normalizeCountyName(rawCounty);
            if (!state.cityFields.has(countyKey)) {
                state.cityFields.set(countyKey, {});
            }
            const field = state.cityFields.get(countyKey);
            if (key.endsWith("_remain")) {
                field.remain = key;
            } else if (key.endsWith("_worth")) {
                field.worth = key;
            } else {
                field.capital = key;
            }
            state.displayCountyByKey.set(countyKey, displayCountyName(rawCounty));
        });
    }

    function prepareDistrictDataset(dataset) {
        const districtRows = Array.isArray(dataset.districtRows) ? dataset.districtRows : [];
        const detailsByCase = new Map();
        const optionsByCounty = new Map();
        const mainKeys = new Set();
        const cityFields = new Map();
        const mainSums = new Map();
        const districtSums = new Map();
        const duplicateKeys = new Map();
        let emptyCityRows = 0;
        let emptyDistrictRows = 0;

        Object.keys(dataset.rawRows[0] || {}).forEach(function (key) {
            if (!/_worth$|_capital$|_remain$/u.test(key)) {
                return;
            }
            const rawCounty = key.replace(/_(worth|capital|remain)$/u, "");
            const metric = key.replace(/^.*_(worth|capital|remain)$/u, "$1");
            const county = normalizeCountyName(rawCounty);
            if (!cityFields.has(county)) {
                cityFields.set(county, {});
            }
            cityFields.get(county)[metric] = key;
        });

        dataset.rawRows.forEach(function (row) {
            const caseKey = getCaseKeyFromRow(row);
            if (!caseKey) {
                return;
            }
            mainKeys.add(caseKey);
            cityFields.forEach(function (fields, county) {
                addMetric(mainSums, buildMetricKey(caseKey, county, "worth"), toNumber(row[fields.worth]));
                addMetric(mainSums, buildMetricKey(caseKey, county, "capital"), toNumber(row[fields.capital]));
                addMetric(mainSums, buildMetricKey(caseKey, county, "remain"), toNumber(row[fields.remain]));
            });
        });

        districtRows.forEach(function (row) {
            const caseKey = getCaseKeyFromRow(row);
            const county = normalizeCountyName(row.city_name);
            const district = text(row.district_result);
            if (!county) {
                emptyCityRows += 1;
            }
            if (!district) {
                emptyDistrictRows += 1;
            }
            if (!caseKey || !county || !district) {
                return;
            }

            const duplicateKey = [caseKey, county, district].join("\u0001");
            duplicateKeys.set(duplicateKey, (duplicateKeys.get(duplicateKey) || 0) + 1);

            if (!detailsByCase.has(caseKey)) {
                detailsByCase.set(caseKey, new Map());
            }
            const cityMap = detailsByCase.get(caseKey);
            if (!cityMap.has(county)) {
                cityMap.set(county, new Map());
            }
            const districtMap = cityMap.get(county);
            const current = districtMap.get(district) || { worth: 0, capital: 0, remain: 0 };
            current.worth += toNumber(row.value_worth);
            current.capital += toNumber(row.value_capital);
            current.remain += toNumber(row.value_remain);
            districtMap.set(district, current);

            if (!optionsByCounty.has(county)) {
                optionsByCounty.set(county, new Set());
            }
            optionsByCounty.get(county).add(district);
            addMetric(districtSums, buildMetricKey(caseKey, county, "worth"), toNumber(row.value_worth));
            addMetric(districtSums, buildMetricKey(caseKey, county, "capital"), toNumber(row.value_capital));
            addMetric(districtSums, buildMetricKey(caseKey, county, "remain"), toNumber(row.value_remain));
        });

        const sortedOptionsByCounty = new Map();
        optionsByCounty.forEach(function (districts, county) {
            sortedOptionsByCounty.set(county, Array.from(districts).filter(Boolean).sort(localeCompare));
        });

        let amountMismatchCount = 0;
        let amountComparedCount = 0;
        districtSums.forEach(function (value, key) {
            amountComparedCount += 1;
            if (Math.abs(value - (mainSums.get(key) || 0)) > 1) {
                amountMismatchCount += 1;
            }
        });

        const allCityKeys = new Set(districtRows.map(getCaseKeyFromRow).filter(Boolean));
        const missingMainKeyCount = Array.from(allCityKeys).filter(function (key) {
            return !mainKeys.has(key);
        }).length;

        dataset.districtDetailsByCase = detailsByCase;
        dataset.districtOptionsByCounty = sortedOptionsByCounty;
        dataset.districtValidation = {
            rowCount: districtRows.length,
            countyCount: sortedOptionsByCounty.size,
            districtCount: Array.from(sortedOptionsByCounty.values()).reduce(function (sum, districts) {
                return sum + districts.length;
            }, 0),
            missingMainKeyCount: missingMainKeyCount,
            emptyCityRows: emptyCityRows,
            emptyDistrictRows: emptyDistrictRows,
            duplicateGroupCount: Array.from(duplicateKeys.values()).filter(function (count) { return count > 1; }).length,
            amountComparedCount: amountComparedCount,
            amountMismatchCount: amountMismatchCount,
            hasFatalIssue: missingMainKeyCount > 0 || emptyCityRows > 0 || emptyDistrictRows > 0 || amountMismatchCount > 0
        };
    }

    function buildMetricKey(caseKey, county, metric) {
        return [caseKey, normalizeCountyName(county), metric].join("\u0001");
    }

    function addMetric(map, key, value) {
        map.set(key, (map.get(key) || 0) + toNumber(value));
    }

    function getCaseKeyFromRow(row) {
        return text(row && row[FIELD.caseId]);
    }

    function buildDistrictVacancyKey(county, district) {
        return [normalizeCountyName(county), text(district)].join("|");
    }

    function normalizePercentLike(value) {
        const num = toNumber(value);
        return num > 1 ? num / 100 : num;
    }

    function prepareVacancyRows() {
        state.vacancyByCounty.clear();
        state.vacancyByDistrict.clear();
        const rows = Array.isArray(window.VACANCY_RATE_ROWS) ? window.VACANCY_RATE_ROWS : [];
        rows.forEach(function (row) {
            const county = normalizeCountyName(row["縣市"]);
            if (!county) {
                return;
            }
            const current = state.vacancyByCounty.get(county) || {
                vacantHomes: 0,
                unsoldHomes: 0,
                totalHomes: 0,
                quadrant: ""
            };
            current.vacantHomes += toNumber(row.vacantHomes);
            current.unsoldHomes += toNumber(row.unsoldHomes);
            current.totalHomes += toNumber(row.totalHomes);
            if (!current.quadrant && row.quadrant) {
                current.quadrant = String(row.quadrant).trim();
            }
            state.vacancyByCounty.set(county, current);
        });
        const districtRows = Array.isArray(window.VACANCY_DISTRICT_RATE_ROWS) ? window.VACANCY_DISTRICT_RATE_ROWS : [];
        districtRows.forEach(function (row) {
            const county = normalizeCountyName(row["縣市"]);
            const district = text(row["鄉鎮市區"]);
            if (!county || !district) {
                return;
            }
            state.vacancyByDistrict.set(buildDistrictVacancyKey(county, district), {
                vacantHomes: toNumber(row.vacantHomes),
                unsoldHomes: toNumber(row.unsoldHomes),
                totalHomes: toNumber(row.totalHomes),
                vacancyRate: normalizePercentLike(row.vacancyRate),
                unsoldRate: normalizePercentLike(row.unsoldRate),
                quadrant: text(row.quadrant) || "-"
            });
        });
    }

    function populateFilters() {
        const counties = sortCountiesByCommonOrder(Array.from(state.cityFields.keys()));
        state.filterOptions = {
            counties: counties.map(function (county) {
                return { value: county, label: displayCountyName(county) };
            }),
            industries: [
            { value: "純土融", label: "純土融" },
            { value: "純建融", label: "純建融" },
            { value: "土建融", label: "純土建融" },
            { value: "餘屋", label: "純餘屋" },
            { value: "其他", label: "其他" }
            ],
            ltvBuckets: LTV_BUCKETS.map(function (bucket) {
                return { value: bucket.id, label: bucket.label };
            })
        };
        renderFacetFilterControls();
    }

    function sortCountiesByCommonOrder(counties) {
        const orderKeys = COMMON_COUNTY_ORDER.map(normalizeCountyName);
        return counties.slice().sort(function (a, b) {
            const ai = orderKeys.indexOf(normalizeCountyName(a));
            const bi = orderKeys.indexOf(normalizeCountyName(b));
            const av = ai < 0 ? Number.MAX_SAFE_INTEGER : ai;
            const bv = bi < 0 ? Number.MAX_SAFE_INTEGER : bi;
            if (av !== bv) {
                return av - bv;
            }
            return localeCompare(a, b);
        });
    }

    function populateNewsCountyOptions() {
        if (!els.newsCountySelect) {
            return;
        }
        const options = [{ value: "", label: "請選擇區域" }].concat(COMMON_COUNTY_ORDER.map(function (county) {
            return { value: county, label: displayCountyName(county) };
        }));
        fillSelect(els.newsCountySelect, options);
    }

    function fillSelect(select, options) {
        select.innerHTML = options.map(function (option) {
            return `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`;
        }).join("");
    }

    function createEmptyFilters() {
        return { county: [], districts: {}, industry: [], industryLabel: "", ltvBucket: [], coverageBucket: "", anomalyMode: "", collateralLevels: [], quadrants: [], group: "", customer: "", riskLevel: "" };
    }

    function getArrayFilterValues(key) {
        const value = state.filters[key];
        if (Array.isArray(value)) {
            return value;
        }
        return value ? [value] : [];
    }

    function setArrayFilterValue(key, value, checked) {
        const values = getArrayFilterValues(key).filter(function (item) {
            return item !== value;
        });
        if (checked) {
            values.push(value);
        }
        state.filters[key] = values;
        if (key === "county") {
            syncDistrictFilterForCounty(value, checked);
        }
    }

    function toggleArrayFilterValue(key, value) {
        setArrayFilterValue(key, value, !getArrayFilterValues(key).includes(value));
    }

    function selectCountyFromMap(county) {
        const countyKey = normalizeCountyName(county);
        if (!countyKey) {
            return;
        }
        const selected = getArrayFilterValues("county").map(normalizeCountyName);
        if (selected.includes(countyKey)) {
            setArrayFilterValue("county", countyKey, false);
            if (state.activeDistrictCounty === countyKey) {
                state.activeDistrictCounty = "";
            }
        } else {
            setArrayFilterValue("county", countyKey, true);
            state.activeDistrictCounty = countyKey;
        }
        state.districtSearch = "";
        closeDistrictPopup();
    }

    function clearMapClickTimer() {
        if (state.mapClickTimer) {
            window.clearTimeout(state.mapClickTimer);
            state.mapClickTimer = null;
        }
    }

    function bindFilterDropdownFloatBehavior() {
        if (!els.filterPanel) {
            return;
        }
        document.addEventListener("toggle", function (event) {
            const current = event.target;
            if (!current || !current.matches || !current.matches("#filterPanel .matrix-dropdown")) {
                return;
            }
            if (!current.open) {
                return;
            }
            closeFilterDropdowns(current);
        }, true);
        document.addEventListener("pointerdown", function (event) {
            if (event.target && event.target.closest && event.target.closest("#filterPanel .matrix-dropdown")) {
                return;
            }
            closeFilterDropdowns();
        });
        document.addEventListener("keydown", function (event) {
            if (event.key === "Escape") {
                closeFilterDropdowns();
            }
        });
    }

    function closeFilterDropdowns(exceptDetails) {
        if (!els.filterPanel) {
            return;
        }
        els.filterPanel.querySelectorAll(".matrix-dropdown[open]").forEach(function (details) {
            if (details !== exceptDetails) {
                details.open = false;
            }
        });
    }

    function clearCountyRankClickTimer() {
        if (state.countyRankClickTimer) {
            window.clearTimeout(state.countyRankClickTimer);
            state.countyRankClickTimer = null;
        }
    }

    function handleCountyRankDrilldownDblclick(params) {
        clearCountyRankClickTimer();
        if (state.countyRankDrilldownCounty) {
            closeCountyRankDrilldown();
            return;
        }
        const county = normalizeCountyName((params && params.data && params.data.rawName) || (params && params.name) || "");
        if (!canDrilldownCountyRank(county) || !(state.districtOptionsByCounty.get(county) || []).length) {
            return;
        }
        state.countyRankDrilldownCounty = county;
        closeDistrictPopup();
        render();
    }

    function canDrilldownCountyRank(county) {
        return !!normalizeCountyName(county);
    }

    function closeCountyRankDrilldown() {
        if (!state.countyRankDrilldownCounty) {
            return;
        }
        clearCountyRankDrilldownFilter(state.countyRankDrilldownCounty);
        state.countyRankDrilldownCounty = "";
        syncFilterControls();
        render();
    }

    function toggleDistrictPopup(county, sourceEvent) {
        const countyKey = normalizeCountyName(county);
        if (!countyKey) {
            return;
        }
        if (state.districtPopupCounty === countyKey) {
            closeDistrictPopup();
            return;
        }
        openDistrictPopup(countyKey, sourceEvent);
    }

    function openDistrictPopup(county, sourceEvent) {
        const countyKey = normalizeCountyName(county);
        if (!countyKey) {
            return;
        }
        state.activeDistrictCounty = countyKey;
        state.districtPopupCounty = countyKey;
        state.districtSearch = "";
        state.districtPopupAnchor = getDistrictPopupAnchor(sourceEvent);
        renderDistrictPanel();
        window.requestAnimationFrame(positionDistrictPopup);
    }

    function closeDistrictPopup() {
        state.districtPopupCounty = "";
        state.districtPopupAnchor = null;
        if (els.districtPanel) {
            els.districtPanel.hidden = true;
            els.districtPanel.classList.remove("open");
            els.districtPanel.setAttribute("aria-hidden", "true");
        }
    }

    function closeTimelineRankDrilldown() {
        if (!state.timelineRankDrilldownCounty) {
            return;
        }
        clearCountyRankDrilldownFilter(state.timelineRankDrilldownCounty);
        state.timelineRankDrilldownCounty = "";
        syncFilterControls();
        renderTimelineDashboard();
    }

    function clearCountyRankDrilldownFilter(county) {
        const countyKey = normalizeCountyName(county);
        if (!countyKey) {
            return;
        }
        if (!state.filters.districts || typeof state.filters.districts !== "object") {
            state.filters.districts = {};
        }
        setArrayFilterValue("county", countyKey, false);
        delete state.filters.districts[countyKey];
        if (state.activeDistrictCounty === countyKey) {
            state.activeDistrictCounty = getArrayFilterValues("county").map(normalizeCountyName)[0] || "";
        }
    }

    function getDistrictPopupAnchor(sourceEvent) {
        const mapRect = els.mapChart ? els.mapChart.getBoundingClientRect() : null;
        const fallback = mapRect ? { x: mapRect.left + mapRect.width * 0.58, y: mapRect.top + mapRect.height * 0.42 } : { x: window.innerWidth / 2, y: window.innerHeight / 2 };
        if (!sourceEvent) {
            return fallback;
        }
        const x = Number.isFinite(sourceEvent.clientX) ? sourceEvent.clientX : fallback.x;
        const y = Number.isFinite(sourceEvent.clientY) ? sourceEvent.clientY : fallback.y;
        return { x: x, y: y };
    }

    function positionDistrictPopup() {
        if (!els.districtPanel || !state.districtPopupCounty) {
            return;
        }
        const anchor = state.districtPopupAnchor || getDistrictPopupAnchor();
        const rect = els.districtPanel.getBoundingClientRect();
        const padding = 12;
        const offset = 26;
        let left = anchor.x + offset;
        let top = anchor.y - Math.min(72, rect.height / 4);
        if (left + rect.width + padding > window.innerWidth) {
            left = anchor.x - rect.width - offset;
        }
        if (left < padding) {
            left = padding;
        }
        if (top + rect.height + padding > window.innerHeight) {
            top = window.innerHeight - rect.height - padding;
        }
        if (top < padding) {
            top = padding;
        }
        els.districtPanel.style.left = `${Math.round(left)}px`;
        els.districtPanel.style.top = `${Math.round(top)}px`;
    }

    function getSingleFilterValue(key) {
        const values = getArrayFilterValues(key);
        return values.length === 1 ? values[0] : "";
    }

    function syncDistrictFilterForCounty(county, checked) {
        const countyKey = normalizeCountyName(county);
        if (!countyKey) {
            return;
        }
        if (!state.filters.districts || typeof state.filters.districts !== "object") {
            state.filters.districts = {};
        }
        if (checked) {
            if (!Object.prototype.hasOwnProperty.call(state.filters.districts, countyKey)) {
                state.filters.districts[countyKey] = "*";
            }
            state.activeDistrictCounty = countyKey;
        } else {
            delete state.filters.districts[countyKey];
            if (state.activeDistrictCounty === countyKey) {
                state.activeDistrictCounty = getArrayFilterValues("county").map(normalizeCountyName)[0] || "";
            }
        }
    }

    function pruneDistrictSelections() {
        const selected = new Set(getArrayFilterValues("county").map(normalizeCountyName));
        const districts = state.filters.districts || {};
        Object.keys(districts).forEach(function (county) {
            if (!selected.has(county)) {
                delete districts[county];
            }
        });
        if (state.activeDistrictCounty && !selected.has(state.activeDistrictCounty)) {
            state.activeDistrictCounty = "";
        }
    }

    function getDistrictFilterValue(county) {
        const countyKey = normalizeCountyName(county);
        const districts = state.filters.districts || {};
        if (!Object.prototype.hasOwnProperty.call(districts, countyKey)) {
            return "*";
        }
        const value = districts[countyKey];
        if (value === "*") {
            return "*";
        }
        return Array.isArray(value) ? value.filter(Boolean) : [];
    }

    function isCountyPartiallySelected(county) {
        return Array.isArray(getDistrictFilterValue(county));
    }

    function setDistrictSelection(county, district, checked) {
        const countyKey = normalizeCountyName(county);
        const districtName = text(district);
        if (!countyKey || !districtName) {
            return;
        }
        if (!getArrayFilterValues("county").map(normalizeCountyName).includes(countyKey)) {
            setArrayFilterValue("county", countyKey, true);
        }
        const allDistricts = state.districtOptionsByCounty.get(countyKey) || [];
        let selected = getDistrictFilterValue(countyKey);
        selected = selected === "*" ? allDistricts.slice() : selected.slice();
        selected = selected.filter(function (item) {
            return item !== districtName;
        });
        if (checked) {
            selected.push(districtName);
        }
        selected = selected.filter(function (item, index, arr) {
            return item && arr.indexOf(item) === index;
        });
        if (!selected.length) {
            setArrayFilterValue("county", countyKey, false);
            if (!state.filters.districts || typeof state.filters.districts !== "object") {
                state.filters.districts = {};
            }
            state.filters.districts[countyKey] = [];
        } else {
            setArrayFilterValue("county", countyKey, true);
            state.filters.districts[countyKey] = selected.length === allDistricts.length ? "*" : selected;
        }
        state.activeDistrictCounty = countyKey;
    }

    function toggleDistrictFromCountyRank(county, district) {
        const countyKey = normalizeCountyName(county);
        const districtName = text(district);
        if (!countyKey || !districtName) {
            return;
        }
        if (!state.filters.districts || typeof state.filters.districts !== "object") {
            state.filters.districts = {};
        }
        setArrayFilterValue("county", countyKey, true);
        const current = getDistrictFilterValue(countyKey);
        const selected = current === "*" ? [] : current.slice();
        const alreadySelected = selected.length === 1 && selected[0] === districtName;
        state.filters.districts[countyKey] = alreadySelected ? "*" : [districtName];
        state.activeDistrictCounty = countyKey;
        syncFilterControls();
        render();
    }

    function selectAllActiveCountyDistricts() {
        const county = getActiveDistrictCounty();
        if (!county) {
            return;
        }
        setArrayFilterValue("county", county, true);
        state.filters.districts[county] = "*";
        state.activeDistrictCounty = county;
    }

    function clearActiveCountyDistricts() {
        const county = getActiveDistrictCounty();
        if (!county) {
            return;
        }
        setArrayFilterValue("county", county, false);
        if (!state.filters.districts || typeof state.filters.districts !== "object") {
            state.filters.districts = {};
        }
        state.filters.districts[county] = [];
        state.activeDistrictCounty = county;
    }

    function getActiveDistrictCounty() {
        const selected = getArrayFilterValues("county").map(normalizeCountyName);
        if (state.activeDistrictCounty) {
            return state.activeDistrictCounty;
        }
        if (!selected.length) {
            return "";
        }
        return selected[0];
    }

    function buildDistrictFilterQuery() {
        const query = {};
        getArrayFilterValues("county").map(normalizeCountyName).forEach(function (county) {
            const value = getDistrictFilterValue(county);
            query[county] = value === "*" ? "*" : value.slice();
        });
        return query;
    }

    function getMultiFilterSummaryLabel(key, emptyLabel, formatter) {
        const values = getArrayFilterValues(key);
        if (!values.length) {
            return emptyLabel;
        }
        if (values.length === 1) {
            return formatter(values[0]);
        }
        return `已選 ${formatNumber(values.length)} 項`;
    }

    function renderFacetFilterControls() {
        if (!state.filterOptions) {
            return;
        }
        const countyCounts = getFacetCounts("county", state.filterOptions.counties.map(function (item) { return item.value; }));
        renderCheckboxFilter(els.countyFilter, state.filterOptions.counties, getArrayFilterValues("county"), "county-filter", countyCounts);
        if (els.countyFilterSummary) {
            els.countyFilterSummary.textContent = getMultiFilterSummaryLabel("county", "全台", displayCountyName);
        }

        const industryCounts = getFacetCounts("industry", state.filterOptions.industries.map(function (item) { return item.value; }));
        renderCheckboxFilter(els.industryFilter, state.filterOptions.industries, getArrayFilterValues("industry"), "industry-filter", industryCounts);
        if (els.industryFilterSummary) {
            els.industryFilterSummary.textContent = getMultiFilterSummaryLabel("industry", "全部", getIndustryFilterDisplayName);
        }

        const ltvCounts = getFacetCounts("ltvBucket", state.filterOptions.ltvBuckets.map(function (item) { return item.value; }));
        renderCheckboxFilter(els.ltvFilter, state.filterOptions.ltvBuckets, getArrayFilterValues("ltvBucket"), "ltv-filter", ltvCounts);
        if (els.ltvFilterSummary) {
            els.ltvFilterSummary.textContent = getMultiFilterSummaryLabel("ltvBucket", "全部", function (value) {
                const bucket = LTV_BUCKETS.find(function (item) { return item.id === value; });
                return bucket ? bucket.label : value;
            });
        }

        updateAnomalyFilterOptions();
        renderDistrictPanel();
    }

    function renderDistrictPanel() {
        if (!els.districtPanel || !els.districtList) {
            return;
        }
        const county = state.districtPopupCounty;
        const validation = state.districtValidation;
        const query = buildDistrictFilterQuery();
        els.districtPanel.dataset.districtQuery = JSON.stringify(query);
        if (!county) {
            els.districtPanel.hidden = true;
            els.districtPanel.classList.remove("open");
            els.districtPanel.classList.add("empty");
            els.districtPanel.setAttribute("aria-hidden", "true");
            els.districtTitle.textContent = "請選取縣市";
            els.districtSummary.textContent = validation && validation.rowCount
                ? `已載入 ${formatNumber(validation.countyCount)} 縣市、${formatNumber(validation.districtCount)} 行政區`
                : "尚未選取行政區";
            els.districtList.innerHTML = "<p class=\"district-empty\">請先從地圖快速雙擊縣市。</p>";
            if (els.districtSelectAllBtn) {
                els.districtSelectAllBtn.disabled = true;
                els.districtSelectAllBtn.textContent = "全選";
            }
            if (els.districtClearBtn) {
                els.districtClearBtn.disabled = true;
            }
            if (els.districtSearch) {
                els.districtSearch.value = state.districtSearch;
                els.districtSearch.disabled = true;
            }
            return;
        }

        const districtStats = buildDistrictAmountRows(county);
        const allDistricts = districtStats.rows.map(function (item) { return item.district; });
        const keyword = text(state.districtSearch).trim();
        const current = getDistrictFilterValue(county);
        const selectedSet = current === "*" ? new Set(allDistricts) : new Set(current);
        const isCountySelected = getArrayFilterValues("county").map(normalizeCountyName).includes(county);
        const visibleRows = districtStats.rows.filter(function (item) {
            return !keyword || item.district.indexOf(keyword) !== -1;
        });

        els.districtPanel.hidden = false;
        els.districtPanel.classList.remove("empty");
        els.districtPanel.classList.add("open");
        els.districtPanel.setAttribute("aria-hidden", "false");
        els.districtTitle.textContent = displayCountyName(county);
        els.districtSummary.textContent = [
            `全市合計：${formatDistrictPopupMoney(districtStats.total)}`,
            `已列示：${formatNumber(allDistricts.length)}區`
        ].join("｜");
        if (els.districtSelectAllBtn) {
            els.districtSelectAllBtn.disabled = allDistricts.length === 0 || current === "*";
            els.districtSelectAllBtn.textContent = "全選";
        }
        if (els.districtClearBtn) {
            els.districtClearBtn.disabled = !isCountySelected;
        }
        if (els.districtSearch) {
            els.districtSearch.disabled = allDistricts.length === 0;
            els.districtSearch.value = state.districtSearch;
        }
        els.districtList.innerHTML = visibleRows.length ? visibleRows.map(function (item) {
            const checked = selectedSet.has(item.district) ? " checked" : "";
            const selectedClass = selectedSet.has(item.district) ? " selected" : "";
            return [
                `<label class="district-option${selectedClass}" data-district-hover-county="${escapeHtml(county)}">`,
                `<input type="checkbox" data-district-filter="1" data-county="${escapeHtml(county)}" value="${escapeHtml(item.district)}"${checked}>`,
                `<span class="district-option-name">${escapeHtml(item.district)}</span>`,
                `<span class="district-option-amount">${escapeHtml(formatDistrictPopupMoney(item.amount))}</span>`,
                `<span class="district-option-share">${escapeHtml(formatOneDecimalPercent(item.share))}</span>`,
                `</label>`
            ].join("");
        }).join("") : "<p class=\"district-empty\">沒有符合的行政區。</p>";
        window.requestAnimationFrame(positionDistrictPopup);
    }

    function buildDistrictAmountRows(county, sourceRows) {
        const countyKey = normalizeCountyName(county);
        const districts = new Map();
        (state.districtOptionsByCounty.get(countyKey) || []).forEach(function (district) {
            districts.set(district, { worth: 0, capital: 0, remain: 0 });
        });
        const rows = Array.isArray(sourceRows)
            ? sourceRows
            : applyMatrixFilters(getBaseFilteredRowsFromRawRows(state.rows, ["county"]));
        const counted = new Set();
        rows.forEach(function (item) {
            const districtSourceRows = Array.isArray(item.sourceRows) && item.sourceRows.length ? item.sourceRows : [item.row];
            districtSourceRows.forEach(function (row) {
                const caseKey = getCaseKeyFromRow(row);
                const cityMap = state.districtDetailsByCase.get(caseKey);
                const districtMap = cityMap && cityMap.get(countyKey);
                if (!caseKey || !districtMap) {
                    return;
                }
                districtMap.forEach(function (metrics, district) {
                    const districtName = text(district);
                    const dedupeKey = [caseKey, countyKey, districtName].join("\u0001");
                    if (!districtName || counted.has(dedupeKey)) {
                        return;
                    }
                    counted.add(dedupeKey);
                    const current = districts.get(districtName) || { worth: 0, capital: 0, remain: 0 };
                    current.worth += toNumber(metrics && metrics.worth);
                    current.capital += toNumber(metrics && metrics.capital);
                    current.remain += toNumber(metrics && metrics.remain);
                    districts.set(districtName, current);
                });
            });
        });
        const total = Array.from(districts.values()).reduce(function (sum, metrics) {
            return sum + toNumber(metrics && metrics.remain);
        }, 0);
        const rowsWithShare = Array.from(districts.entries()).map(function (entry) {
            const metrics = entry[1] || {};
            const principalValue = toNumber(metrics.remain);
            const approvedUnusedValue = toNumber(metrics.capital);
            const amount = principalValue;
            return {
                district: entry[0],
                amount: amount,
                principalValue: principalValue,
                approvedUnusedValue: approvedUnusedValue,
                collateralValue: toNumber(metrics.worth),
                value: state.exposureBasis === "household" ? principalValue + approvedUnusedValue : principalValue,
                share: safeDivide(amount, total)
            };
        }).sort(function (a, b) {
            if (b.value !== a.value) {
                return b.value - a.value;
            }
            return localeCompare(a.district, b.district);
        });
        return { rows: rowsWithShare, total: total };
    }

    function formatDistrictPopupMoney(value) {
        const amount = toNumber(value);
        if (state.amountUnit === "yi") {
            return `${(amount / 100000).toFixed(2)}億元`;
        }
        return `${formatNumber(amount)}仟元`;
    }

    function renderCheckboxFilter(container, options, selectedValues, dataName, counts) {
        if (!container) {
            return;
        }
        container.innerHTML = options.map(function (option) {
            const checked = selectedValues.includes(option.value) ? " checked" : "";
            const count = counts.get(option.value) || 0;
            return `<label class="matrix-checkbox-item"><input type="checkbox" data-${dataName} value="${escapeHtml(option.value)}"${checked}><span>${escapeHtml(option.label)}（共 ${escapeHtml(formatNumber(count))} 件）</span></label>`;
        }).join("");
    }

    function getFacetCounts(fieldKey, optionValues) {
        const ignored = [fieldKey];
        const rows = applyMatrixFilters(getBaseFilteredRowsFromRawRows(state.rows, ignored));
        const counts = new Map(optionValues.map(function (value) {
            return [value, 0];
        }));
        rows.forEach(function (item) {
            if (fieldKey === "county") {
                const county = normalizeCountyName(item.county);
                if (counts.has(county)) {
                    counts.set(county, counts.get(county) + 1);
                }
            } else if (fieldKey === "industry") {
                const industry = item.industryCategory || getIndustryCategory(item);
                if (counts.has(industry)) {
                    counts.set(industry, counts.get(industry) + 1);
                }
            } else if (fieldKey === "ltvBucket" && item.ltvBucket && counts.has(item.ltvBucket)) {
                counts.set(item.ltvBucket, counts.get(item.ltvBucket) + 1);
            }
        });
        return counts;
    }

    function updateAnomalyFilterOptions() {
        if (!els.anomalyFilter) {
            return;
        }
        const rows = applyMatrixFilters(getBaseFilteredRowsFromRawRows(state.rows, ["anomalyMode"]));
        const notStartedCount = rows.filter(isAnomalyItem).length;
        const startedCount = rows.length - notStartedCount;
        const current = state.filters.anomalyMode;
        fillSelect(els.anomalyFilter, [
            { value: "", label: `全部（共 ${formatNumber(rows.length)} 件）` },
            { value: "exclude", label: `僅顯示已起租（共 ${formatNumber(startedCount)} 件）` },
            { value: "only", label: `僅顯示未起租（共 ${formatNumber(notStartedCount)} 件）` }
        ]);
        els.anomalyFilter.value = current;
    }

    function populatePeriodControls() {
        const periodOptions = state.timelinePeriods.map(function (period) {
            return { value: period.periodId, label: period.label };
        });
        if (els.selectedPeriodSelect) {
            fillSelect(els.selectedPeriodSelect, periodOptions);
            els.selectedPeriodSelect.value = state.selectedPeriodId;
        }
        if (els.comparePeriodSelect) {
            const compareOptions = state.timelinePeriods.filter(function (period) {
                return period.periodId !== state.selectedPeriodId;
            }).map(function (period) {
                return { value: period.periodId, label: period.label };
            });
            fillSelect(els.comparePeriodSelect, [{ value: "", label: "無比較期" }].concat(compareOptions));
            els.comparePeriodSelect.value = state.comparePeriodId || "";
        }
        syncTimelineControls();
    }

    function ensureComparePeriod() {
        const selectedIndex = state.timelinePeriods.findIndex(function (period) {
            return period.periodId === state.selectedPeriodId;
        });
        const validCompare = state.comparePeriodId && state.periodDatasets.has(state.comparePeriodId) && state.comparePeriodId !== state.selectedPeriodId;
        if (validCompare) {
            return;
        }
        state.comparePeriodId = selectedIndex > 0 ? state.timelinePeriods[selectedIndex - 1].periodId : "";
    }

    function syncFilterControls() {
        renderFacetFilterControls();
        els.anomalyFilter.value = state.filters.anomalyMode;
        syncExposureBasisControls();
    }

    function render() {
        if (!state.rows.length) {
            return;
        }
        syncTimelineControls();

        const baseRows = getBaseFilteredRows();
        renderMatrixFilterButtons(baseRows);
        const filteredRows = applyMatrixFilters(baseRows);
        if (state.timelineMode === "timeline") {
            renderSummary(getTimelineIndicatorRows(), { forceRegionLabel: "全台" });
            setTimelineViewVisible(true);
            renderTimelineDashboard();
        } else {
            renderSummary(filteredRows);
            setTimelineViewVisible(false);
            renderMap(filteredRows);
            renderMainPanelCharts(filteredRows);
            if (charts.ltv) {
                renderLtvDistribution(getRowsForChart("ltv2"));
            }
            if (charts.coverage) {
                renderCoverageDistribution(getRowsForChart("coverage"));
            }
            if (els.concentrationList) {
                renderConcentration(filteredRows);
            }
        }
        renderCountySummaryTable(filteredRows);
        renderTable(filteredRows);
        renderNewsPanel();
        renderMainGridMode();
        renderMainPanelChartPicker();
        if (els.riskDefinitionNote) {
            els.riskDefinitionNote.textContent = RISK_NOTE_DEFINITION;
        }
        syncNewsFormVisibility();
        if (els.analysisDrawerPanel && !els.analysisDrawerPanel.hidden) {
            renderAnalysisCharts(filteredRows);
        }
        syncDeveloperMode();
        applyComponentDisplayNames();
        renderDistrictPanel();
        updateStickyOffsets();
    }

    function updateStickyOffsets() {
        const kpiBar = document.querySelector(".secondary-kpis.kpi-bar");
        const isVisible = kpiBar && kpiBar.offsetParent !== null && !kpiBar.hidden;
        const height = isVisible ? Math.ceil(kpiBar.getBoundingClientRect().height) : 0;
        document.documentElement.style.setProperty("--sticky-kpi-offset", `${height}px`);
    }

    function getFilteredRows() {
        return applyMatrixFilters(getBaseFilteredRows());
    }

    function getBaseFilteredRows() {
        return getBaseFilteredRowsFromRawRows(state.rows);
    }

    function getFilteredRowsFromRawRows(rawRows) {
        return applyMatrixFilters(getBaseFilteredRowsFromRawRows(rawRows));
    }

    function getTimelineRowsFromRawRows(rawRows) {
        if (!state.disconnectTimelineFilters) {
            return getFilteredRowsFromRawRows(rawRows);
        }
        return applyMatrixFilters(getBaseFilteredRowsFromRawRows(rawRows, [
            "county",
            "districts",
            "industry",
            "group",
            "customer",
            "riskLevel",
            "ltvBucket",
            "coverageBucket",
            "anomalyMode"
        ]), {
            ignoreCollateral: true,
            ignoreQuadrants: true
        });
    }

    function getTimelineIndicatorRows() {
        const dataset = state.periodDatasets.get(state.selectedPeriodId);
        const rawRows = dataset ? dataset.rawRows : state.rows;
        return applyMatrixFilters(getBaseFilteredRowsFromRawRows(rawRows, [
            "county",
            "districts",
            "industry",
            "group",
            "customer",
            "riskLevel",
            "ltvBucket",
            "coverageBucket",
            "anomalyMode"
        ]), {
            ignoreCollateral: true,
            ignoreQuadrants: true
        });
    }

    function getRowsForChart(chartId, rawRows) {
        const ignored = getIgnoredFiltersForChart(chartId);
        return applyMatrixFilters(getBaseFilteredRowsFromRawRows(rawRows || state.rows, ignored), {
            ignoreCollateral: ignored.indexOf("collateralLevels") !== -1,
            ignoreQuadrants: ignored.indexOf("quadrants") !== -1
        });
    }

    function getIgnoredFiltersForChart(chartId) {
        const map = {
            countyRank: ["county"],
            group: ["group"],
            topCustomer: ["customer"],
            industryPie: ["industry"],
            riskPie: ["riskLevel"],
            coverage: ["coverageBucket", "collateralLevels"],
            quadrantSummary: ["quadrants"],
            quadrantMatrix: ["customer"]
        };
        return map[chartId] || [];
    }

    function getActiveDistrictQuadrantCounty() {
        const selectedCounties = getArrayFilterValues("county").map(normalizeCountyName).filter(function (county) {
            return (state.districtOptionsByCounty.get(county) || []).length > 0;
        });
        const activeCounty = normalizeCountyName(state.activeDistrictCounty);
        if (activeCounty && selectedCounties.includes(activeCounty)) {
            return activeCounty;
        }
        return selectedCounties.length === 1 ? selectedCounties[0] : "";
    }

    function getDistrictsForQuadrantContext(county) {
        const countyKey = normalizeCountyName(county);
        const districts = getDistrictFilterValue(countyKey);
        if (districts === "*") {
            return (state.districtOptionsByCounty.get(countyKey) || []).slice();
        }
        return Array.isArray(districts) ? districts.slice() : [];
    }

    function getSelectedDistrictQuadrantIds(county) {
        const countyKey = normalizeCountyName(county);
        const selectedDistricts = getDistrictFilterValue(countyKey);
        if (!Array.isArray(selectedDistricts) || !selectedDistricts.length) {
            return new Set();
        }
        return new Set(selectedDistricts.map(function (district) {
            return getDistrictQuadrantId(countyKey, district);
        }).filter(Boolean));
    }

    function getDistrictQuadrantId(county, district) {
        const vacancy = state.vacancyByDistrict.get(buildDistrictVacancyKey(county, district));
        return vacancy ? getQuadrantIdBySupplyId(getSupplyPressure(vacancy).id) : "";
    }

    function getItemDistrictQuadrantId(item, county) {
        const countyKey = normalizeCountyName(county);
        const candidateDistricts = getDistrictsForQuadrantContext(countyKey);
        if (!countyKey || !candidateDistricts.length || !item) {
            return "";
        }
        const districtSet = new Set(candidateDistricts);
        const weights = new Map();
        const counted = new Set();
        const sourceRows = Array.isArray(item.sourceRows) && item.sourceRows.length ? item.sourceRows : [item.row];

        sourceRows.forEach(function (row) {
            const caseKey = getCaseKeyFromRow(row);
            const cityMap = state.districtDetailsByCase.get(caseKey);
            const districtMap = cityMap && cityMap.get(countyKey);
            if (!caseKey || !districtMap) {
                return;
            }
            districtMap.forEach(function (metrics, district) {
                const districtName = text(district);
                const dedupeKey = [caseKey, countyKey, districtName].join("\u0001");
                if (!districtSet.has(districtName) || counted.has(dedupeKey)) {
                    return;
                }
                counted.add(dedupeKey);
                const weight = toNumber(metrics && metrics.remain) || toNumber(metrics && metrics.capital) || toNumber(metrics && metrics.worth) || 0;
                weights.set(districtName, (weights.get(districtName) || 0) + weight);
            });
        });

        if (!weights.size) {
            return "";
        }
        const dominantDistrict = Array.from(weights.entries()).sort(function (a, b) {
            if (b[1] !== a[1]) {
                return b[1] - a[1];
            }
            return candidateDistricts.indexOf(a[0]) - candidateDistricts.indexOf(b[0]);
        })[0][0];
        return getDistrictQuadrantId(countyKey, dominantDistrict);
    }

    function getItemEffectiveQuadrantId(item) {
        const districtCounty = getActiveDistrictQuadrantCounty();
        if (districtCounty) {
            const districtQuadrantId = getItemDistrictQuadrantId(item, districtCounty);
            if (districtQuadrantId) {
                return districtQuadrantId;
            }
        }
        return item && item.matrix ? item.matrix.quadrantId : "";
    }

    function getBaseFilteredRowsFromRawRows(rawRows, ignoredFilters) {
        const ignored = new Set(ignoredFilters || []);
        const ignoreDistricts = ignored.has("districts");
        const selectedCounties = getArrayFilterValues("county").map(normalizeCountyName);
        const useCountyFilter = selectedCounties.length > 0 && !ignored.has("county");
        const singleCounty = useCountyFilter && selectedCounties.length === 1 ? selectedCounties[0] : "";
        const selectedCountyCaseKeys = singleCounty ? getSelectedCountyCaseKeys(rawRows, singleCounty) : new Set();
        const derivedRows = rawRows.flatMap(function (row, index) {
            return deriveRowsForMerge(row, index, singleCounty, ignoreDistricts);
        }).filter(function (item) {
            const row = item.row;
            if (useCountyFilter) {
                if (singleCounty) {
                    return hasCountyData(row, singleCounty) || selectedCountyCaseKeys.has(getBaseCaseKey(item));
                }
                return selectedCounties.includes(normalizeCountyName(item.county));
            }
            return true;
        });
        if (!ignoreDistricts) {
            dedupeSelectedDistrictMetrics(derivedRows);
        }

        const mergedRows = mergeCaseRows(derivedRows).map(function (item) {
            item.industryCategory = getIndustryCategory(item);
            return item;
        });
        const rowsForFilter = mergedRows.filter(function (item) {
            return item.principalBalance > 0 || item.approvedUnused > 0;
        });

        return rowsForFilter.filter(function (item) {
            if (!ignored.has("industry") && !passesIndustryFilter(item)) {
                return false;
            }
            if (!ignored.has("industry") && state.filters.industryLabel && item.industry !== state.filters.industryLabel) {
                return false;
            }
            if (!ignored.has("group") && state.filters.group && !item.groups.includes(state.filters.group)) {
                return false;
            }
            if (!ignored.has("customer") && state.filters.customer && item.applicant !== state.filters.customer) {
                return false;
            }
            if (!ignored.has("riskLevel") && state.filters.riskLevel && item.riskLevel !== state.filters.riskLevel) {
                return false;
            }
            const selectedLtvBuckets = getArrayFilterValues("ltvBucket");
            if (!ignored.has("ltvBucket") && selectedLtvBuckets.length && (!item.ltvBucket || !selectedLtvBuckets.includes(item.ltvBucket))) {
                return false;
            }
            if (!ignored.has("coverageBucket") && state.filters.coverageBucket && (!item.principalBalance || getCoverageBucketByRatio(item.coverageRatio || 0).id !== state.filters.coverageBucket)) {
                return false;
            }
            if (!ignored.has("anomalyMode") && state.filters.anomalyMode === "exclude" && isAnomalyItem(item)) {
                return false;
            }
            if (!ignored.has("anomalyMode") && state.filters.anomalyMode === "only" && !isAnomalyItem(item)) {
                return false;
            }
            return true;
        });
    }

    function applyMatrixFilters(rows, options) {
        const ignoreCollateral = options && options.ignoreCollateral;
        const ignoreQuadrants = options && options.ignoreQuadrants;
        return rows.filter(function (item) {
            if (!ignoreCollateral && state.filters.collateralLevels.length && !state.filters.collateralLevels.includes(item.matrix.collateral.id)) {
                return false;
            }
            if (!ignoreQuadrants && state.filters.quadrants.length && !state.filters.quadrants.includes(getItemEffectiveQuadrantId(item))) {
                return false;
            }
            return true;
        });
    }

    function dedupeSelectedDistrictMetrics(rows) {
        const seen = new Set();
        rows.forEach(function (item) {
            const county = normalizeCountyName(item.county);
            const districtFilter = getDistrictFilterValue(county);
            if (!Array.isArray(districtFilter) || !districtFilter.length) {
                return;
            }
            const key = [getBaseCaseKey(item), county].join("|");
            if (!seen.has(key)) {
                seen.add(key);
                return;
            }
            item.principalBalance = 0;
            item.approvedUnused = 0;
            item.householdAmount = 0;
            item.collateralValue = 0;
        });
    }

    function passesIndustryFilter(item) {
        const industries = getArrayFilterValues("industry");
        if (!industries.length) {
            return true;
        }
        return industries.includes(item.industryCategory || getIndustryCategory(item));
    }

    function getIndustryFilterDisplayName(value) {
        const map = {
            "土建融": "純土建融",
            "餘屋": "純餘屋"
        };
        return map[value] || value;
    }

    function normalizeIndustryLabel(value) {
        return text(value).replace(/\([0-9]+\)/g, "").replace(/\s/g, "");
    }

    function getIndustryCategory(item) {
        const normalizedIndustries = item.industries.map(function (entry) {
            return normalizeIndustryLabel(entry);
        });
        const hasLand = normalizedIndustries.some(function (entry) { return entry.includes("土融"); });
        const hasBuild = normalizedIndustries.some(function (entry) { return entry.includes("建融"); });
        const hasUnsold = normalizedIndustries.some(function (entry) { return entry.includes("餘屋"); });
        const hasOther = normalizedIndustries.some(function (entry) { return entry.includes("其他"); });
        const hasOnlyLandBuild = hasLand && hasBuild && !hasUnsold && !hasOther;
        const hasOnlyLand = hasLand && !hasBuild && !hasUnsold && !hasOther;
        const hasOnlyBuild = hasBuild && !hasLand && !hasUnsold && !hasOther;
        const hasOnlyUnsold = hasUnsold && !hasLand && !hasBuild && !hasOther;
        if (hasOnlyLandBuild) {
            return "土建融";
        }
        if (hasOnlyLand) {
            return "純土融";
        }
        if (hasOnlyBuild) {
            return "純建融";
        }
        if (hasOnlyUnsold) {
            return "餘屋";
        }
        return "其他";
    }

    function renderMatrixFilterButtons(baseRows) {
        const counts = getMatrixFacetCounts(baseRows);
        els.matrixCollateralFilters.innerHTML = COLLATERAL_LEVEL_FILTERS.map(function (item) {
            const count = counts.collateral.get(item.id) || 0;
            const checked = state.filters.collateralLevels.includes(item.id) ? " checked" : "";
            return `
                <label class="matrix-checkbox-item">
                    <input type="checkbox" data-collateral-level="${escapeHtml(item.id)}"${checked}>
                    <span>擔保覆蓋率${escapeHtml(item.label)}共${formatNumber(count)}件</span>
                </label>
            `;
        }).join("");

        els.matrixQuadrantFilters.innerHTML = QUADRANT_FILTERS.map(function (item) {
            const count = counts.quadrants.get(item.id) || 0;
            const checked = state.filters.quadrants.includes(item.id) ? " checked" : "";
            const label = `${item.label}-${item.description.replace("供需", "")}`;
            return `
                <label class="matrix-checkbox-item">
                    <input type="checkbox" data-quadrant-filter="${escapeHtml(item.id)}"${checked}>
                    <span>${escapeHtml(label)}共${formatNumber(count)}件</span>
                </label>
            `;
        }).join("");
    }

    function getCoverageBucketByRatio(coverageRatio) {
        return COVERAGE_BUCKETS.find(function (entry) {
            return coverageRatio >= entry.min && coverageRatio < entry.max;
        }) || COVERAGE_BUCKETS[0];
    }

    function getMatrixFacetCounts(baseRows) {
        const collateralRows = applyMatrixFilters(baseRows, { ignoreCollateral: true });
        const quadrantRows = applyMatrixFilters(baseRows, { ignoreQuadrants: true });
        return {
            collateral: countBy(collateralRows, function (item) {
                return item.matrix.collateral.id;
            }),
            quadrants: countBy(quadrantRows, function (item) {
                return getItemEffectiveQuadrantId(item);
            })
        };
    }

    function toggleArrayFilter(list, value) {
        const index = list.indexOf(value);
        if (index >= 0) {
            list.splice(index, 1);
        } else if (value) {
            list.push(value);
        }
    }

    function countBy(rows, getKey) {
        const result = new Map();
        rows.forEach(function (item) {
            const key = getKey(item);
            if (key) {
                result.set(key, (result.get(key) || 0) + 1);
            }
        });
        return result;
    }

    function deriveRow(row, county, ignoreDistricts) {
        const countyKey = county || "";
        const districtMetrics = countyKey ? getSelectedDistrictMetrics(row, countyKey, ignoreDistricts) : null;
        const principalBalance = districtMetrics ? districtMetrics.remain : (countyKey ? countyExposure(row, countyKey) : rowExposure(row));
        const collateralValue = districtMetrics ? districtMetrics.worth : (countyKey ? countyCollateral(row, countyKey) : rowCollateral(row));
        const approvedUnused = districtMetrics ? districtMetrics.capital : toNumber(row[FIELD.activeLimit]);
        return createDerivedItem(row, countyKey, principalBalance, collateralValue, approvedUnused);
    }

    function deriveRowsForMerge(row, sourceIndex, county, ignoreDistricts) {
        if (county) {
            const single = deriveRow(row, county, ignoreDistricts);
            single.sourceIndex = sourceIndex;
            return [single];
        }

        const countyRows = deriveCountySplitRows(row, ignoreDistricts);
        if (!countyRows.length) {
            const fallback = deriveRow(row, "", ignoreDistricts);
            fallback.sourceIndex = sourceIndex;
            return [fallback];
        }

        return countyRows.map(function (item) {
            item.sourceIndex = sourceIndex;
            return item;
        });
    }

    function deriveCountySplitRows(row, ignoreDistricts) {
        const counties = [];
        state.cityFields.forEach(function (_fields, county) {
            const districtMetrics = getSelectedDistrictMetrics(row, county, ignoreDistricts);
            const principal = districtMetrics ? districtMetrics.remain : countyExposure(row, county);
            const collateral = districtMetrics ? districtMetrics.worth : countyCollateral(row, county);
            const capital = districtMetrics ? districtMetrics.capital : null;
            if (principal > 0 || collateral > 0) {
                counties.push({
                    county: county,
                    principalBalance: principal,
                    collateralValue: collateral,
                    approvedUnused: capital
                });
            }
        });

        if (!counties.length) {
            return [];
        }

        const approvedUnused = toNumber(row[FIELD.activeLimit]);
        const totalPrincipal = counties.reduce(function (sum, item) {
            return sum + item.principalBalance;
        }, 0);
        const totalCollateral = counties.reduce(function (sum, item) {
            return sum + item.collateralValue;
        }, 0);

        return counties.map(function (entry, index) {
            if (entry.approvedUnused !== null && typeof entry.approvedUnused !== "undefined") {
                return createDerivedItem(
                    row,
                    entry.county,
                    entry.principalBalance,
                    entry.collateralValue,
                    entry.approvedUnused
                );
            }
            const shareBase = totalPrincipal > 0 ? totalPrincipal : totalCollateral;
            const shareValue = totalPrincipal > 0 ? entry.principalBalance : entry.collateralValue;
            const ratio = shareBase > 0 ? shareValue / shareBase : 1 / counties.length;
            const allocatedApprovedUnused = index === counties.length - 1
                ? approvedUnused - counties.slice(0, index).reduce(function (sum, previous) {
                    const previousShareValue = totalPrincipal > 0 ? previous.principalBalance : previous.collateralValue;
                    const previousRatio = shareBase > 0 ? previousShareValue / shareBase : 1 / counties.length;
                    return sum + (approvedUnused * previousRatio);
                }, 0)
                : approvedUnused * ratio;

            return createDerivedItem(
                row,
                entry.county,
                entry.principalBalance,
                entry.collateralValue,
                allocatedApprovedUnused
            );
        });
    }

    function getSelectedDistrictMetrics(row, county, ignoreDistricts) {
        if (ignoreDistricts) {
            return null;
        }
        const countyKey = normalizeCountyName(county);
        const selectedDistricts = getDistrictFilterValue(countyKey);
        if (selectedDistricts === "*") {
            return null;
        }
        const caseKey = getCaseKeyFromRow(row);
        const cityMap = state.districtDetailsByCase.get(caseKey);
        const districtMap = cityMap && cityMap.get(countyKey);
        if (!districtMap || !selectedDistricts.length) {
            return { worth: 0, capital: 0, remain: 0 };
        }
        return selectedDistricts.reduce(function (acc, district) {
            const item = districtMap.get(district);
            if (item) {
                acc.worth += item.worth;
                acc.capital += item.capital;
                acc.remain += item.remain;
            }
            return acc;
        }, { worth: 0, capital: 0, remain: 0 });
    }

    function createDerivedItem(row, county, principalBalance, collateralValue, approvedUnused) {
        const householdAmount = principalBalance + approvedUnused;
        const hasPrincipal = principalBalance > 0;
        const principalLtv = hasPrincipal ? safeDivide(principalBalance, collateralValue) : null;
        const householdLtv = safeDivide(householdAmount, collateralValue);
        const coverageRatio = hasPrincipal ? safeDivide(collateralValue, principalBalance) : null;
        const riskLevel = getRiskLevel(principalLtv || 0, collateralValue);
        const primaryCounty = county || getPrimaryCounty(row);
        const riskReasons = buildRiskReasons(principalLtv, coverageRatio, collateralValue, primaryCounty);

        return {
            row: row,
            county: primaryCounty,
            group: text(row[FIELD.group]) || "未命名集團",
            applicant: text(row[FIELD.applicant]) || text(row[FIELD.group]) || "未命名客戶",
            industry: text(row[FIELD.industry]) || "未分類",
            principalBalance: principalBalance,
            approvedUnused: approvedUnused,
            householdAmount: householdAmount,
            collateralValue: collateralValue,
            principalLtv: principalLtv,
            householdLtv: householdLtv,
            coverageRatio: coverageRatio,
            riskLevel: riskLevel,
            ltvBucket: hasPrincipal ? getLtvBucket(principalLtv) : "",
            riskReasons: riskReasons
        };
    }

    function mergeCaseRows(rows) {
        const groups = new Map();
        const countyByCase = resolveCountyByCase(rows);
        const caseIndustriesByCase = buildCaseIndustriesByCase(rows);
        rows.forEach(function (item) {
            const key = getCaseMergeKey(item, countyByCase);
            const resolvedCounty = getResolvedMergeCounty(item, countyByCase);
            if (!groups.has(key)) {
                groups.set(key, {
                    mergeKey: key,
                    row: item.row,
                    sourceRows: [],
                    sourceIndexes: [],
                    county: resolvedCounty,
                    group: item.group,
                    applicant: item.applicant,
                    caseId: text(item.row[FIELD.caseId]),
                    industries: [],
                    allIndustries: [],
                    groups: [],
                    principalBalance: 0,
                    approvedUnused: 0,
                    householdAmount: 0,
                    collateralValue: 0
                });
            }

            const current = groups.get(key);
            current.sourceRows.push(item.row);
            current.sourceIndexes.push(item.sourceIndex);
            current.principalBalance += item.principalBalance;
            current.approvedUnused += item.approvedUnused;
            current.collateralValue += item.collateralValue;
            pushUnique(current.allIndustries, item.industry);
            if (item.principalBalance > 0 || item.approvedUnused > 0) {
                pushUnique(current.industries, item.industry);
            }
            pushUnique(current.groups, item.group);
        });

        return Array.from(groups.values()).map(function (item) {
            item.householdAmount = item.principalBalance + item.approvedUnused;
            const caseIndustries = caseIndustriesByCase.get(getBaseCaseKey(item)) || item.allIndustries;
            item.industries = sortIndustriesByCode(caseIndustries);
            item.industry = item.industries.join(" 及 ");
            item.group = item.groups[0] || item.group;
            const hasPrincipal = item.principalBalance > 0;
            item.principalLtv = hasPrincipal ? safeDivide(item.principalBalance, item.collateralValue) : null;
            item.householdLtv = safeDivide(item.householdAmount, item.collateralValue);
            item.coverageRatio = hasPrincipal ? safeDivide(item.collateralValue, item.principalBalance) : null;
            item.riskLevel = getRiskLevel(item.principalLtv || 0, item.collateralValue);
            item.ltvBucket = hasPrincipal ? getLtvBucket(item.principalLtv) : "";
            item.riskReasons = buildRiskReasons(item.principalLtv || 0, item.coverageRatio || 0, item.collateralValue, item.county);
            item.matrix = getMatrixEvaluation(item);
            if (item.sourceRows.length > 1) {
                item.riskReasons.unshift(MERGED_CASE_REASON);
            }
            return item;
        });
    }

    function buildCaseIndustriesByCase(rows) {
        const caseRows = new Map();
        rows.forEach(function (item) {
            const caseKey = getBaseCaseKey(item);
            if (!caseKey) {
                return;
            }
            if (!caseRows.has(caseKey)) {
                caseRows.set(caseKey, new Map());
            }
            caseRows.get(caseKey).set(item.sourceIndex, item.industry);
        });

        const result = new Map();
        caseRows.forEach(function (rowMap, caseKey) {
            const industries = [];
            rowMap.forEach(function (industry) {
                pushUnique(industries, industry);
            });
            result.set(caseKey, sortIndustriesByCode(industries));
        });
        return result;
    }

    function sortIndustriesByCode(industries) {
        return industries.slice().sort(function (a, b) {
            const parsedA = parseIndustryCode(a);
            const parsedB = parseIndustryCode(b);
            if (parsedA.code !== parsedB.code) {
                return parsedA.code - parsedB.code;
            }
            return localeCompare(parsedA.label, parsedB.label);
        });
    }

    function parseIndustryCode(industry) {
        const label = text(industry);
        const match = /^\((\d+)\)\s*/.exec(label);
        return {
            code: match ? Number(match[1]) : Number.MAX_SAFE_INTEGER,
            label: label
        };
    }

    function getSelectedCountyCaseKeys(rawRows, county) {
        const keys = new Set();
        (rawRows || state.rows).forEach(function (row, index) {
            if (!hasCountyData(row, county)) {
                return;
            }
            keys.add(getBaseCaseKey({
                row: row,
                sourceIndex: index,
                applicant: text(row[FIELD.applicant]) || text(row[FIELD.group])
            }));
        });
        return keys;
    }

    function resolveCountyByCase(rows) {
        const countyByCase = new Map();
        rows.forEach(function (item) {
            const caseKey = getBaseCaseKey(item);
            if (!caseKey || isUnassignedCounty(item.county)) {
                return;
            }
            if (!countyByCase.has(caseKey)) {
                countyByCase.set(caseKey, item.county);
            }
        });
        return countyByCase;
    }

    function getCaseMergeKey(item, countyByCase) {
        return [
            getBaseCaseKey(item),
            normalizeMergeKey(getResolvedMergeCounty(item, countyByCase))
        ].join("|");
    }

    function getBaseCaseKey(item) {
        const caseId = text(item.row[FIELD.caseId]);
        const stableCaseId = caseId || `row-${item.sourceIndex}`;
        return [
            normalizeMergeKey(item.applicant),
            normalizeMergeKey(stableCaseId)
        ].join("|");
    }

    function getResolvedMergeCounty(item, countyByCase) {
        if (isUnassignedCounty(item.county)) {
            return countyByCase.get(getBaseCaseKey(item)) || item.county;
        }
        return item.county;
    }

    function isUnassignedCounty(county) {
        const normalized = normalizeMergeKey(county);
        return !normalized || normalized === normalizeMergeKey("全台");
    }

    function formatCountySummaryName(county) {
        return normalizeMergeKey(county) === "\u5168\u53f0" ? "\u5c1a\u672a\u8d77\u79df" : text(county);
    }

    function normalizeMergeKey(value) {
        return text(value).toLowerCase();
    }

    function pushUnique(list, value) {
        const normalized = text(value);
        if (normalized && !list.includes(normalized)) {
            list.push(normalized);
        }
    }

    function isHighRiskItem(item) {
        return item.riskLevel === "danger" || item.riskReasons.some(function (reason) {
            return reason !== MERGED_CASE_REASON;
        });
    }

    function isAnomalyItem(item) {
        return item.matrix && item.matrix.supply && item.matrix.supply.id === "unknown";
    }

    function getMatrixEvaluation(item) {
        const collateral = getCollateralMatrixLevel(item.principalLtv, item.coverageRatio, item.collateralValue);
        const supply = getSupplyPressure(getVacancyMetrics(item.county));
        return {
            collateral: collateral,
            supply: supply,
            quadrantId: getQuadrantIdBySupplyId(supply.id),
            label: `${collateral.label} / ${supply.label}`,
            className: `${collateral.className} ${supply.className}`
        };
    }

    function getQuadrantIdBySupplyId(supplyId) {
        const quadrant = QUADRANT_FILTERS.find(function (item) {
            return item.supplyId === supplyId;
        });
        return quadrant ? quadrant.id : "";
    }

    function getCollateralMatrixLevel(ltv, coverageRatio, collateralValue) {
        const bucket = LTV_BUCKETS.find(function (entry) {
            return coverageRatio >= entry.min && coverageRatio < entry.max;
        }) || LTV_BUCKETS[LTV_BUCKETS.length - 1];
        return MATRIX_META[bucket.id] || MATRIX_META["100+"];
    }

    function getSupplyPressure(vacancy) {
        const quadrant = text(vacancy.quadrant);
        if (/^I(\s|$)/u.test(quadrant)) {
            return SUPPLY_PRESSURE_META.both;
        }
        if (/^II(\s|$)/u.test(quadrant)) {
            return SUPPLY_PRESSURE_META.unsold;
        }
        if (/^IV(\s|$)/u.test(quadrant)) {
            return SUPPLY_PRESSURE_META.vacancy;
        }
        if (/^III(\s|$)/u.test(quadrant)) {
            return SUPPLY_PRESSURE_META.low;
        }
        return SUPPLY_PRESSURE_META.unknown;
    }

    function buildRiskReasons(ltv, coverageRatio, collateralValue, county) {
        const reasons = [];
        if (!collateralValue) {
            reasons.push("擔保值為 0");
        }
        if (ltv > 1) {
            reasons.push("LTV > 100%");
        }
        if (coverageRatio > 0 && coverageRatio < 1) {
            reasons.push("擔保覆蓋率 < 100%");
        }
        const vacancy = getVacancyMetrics(county);
        if (vacancy.vacancyRate >= 0.12) {
            reasons.push("高空屋率區域");
        }
        return reasons;
    }

    function renderSummary(rows, options) {
        const config = options || {};
        const selectedCounty = getSingleFilterValue("county");
        const regionCounty = config.forceRegionLabel ? "" : selectedCounty;
        const summary = summarizeCurrentKpiRows(rows);
        const vacancy = getVacancyMetrics(regionCounty);

        els.kpiPrincipal.textContent = formatK(summary.principalBalance);
        renderKpiSplitMoney(els.kpiApprovedUnused, summary.approvedUnusedStarted, summary.approvedUnusedNotStarted);
        els.kpiHousehold.textContent = formatK(summary.householdAmount);
        els.kpiCollateral.textContent = formatK(summary.collateralValue);
        if (els.kpiCases) {
            els.kpiCases.textContent = formatNumber(summary.count);
        }
        if (els.kpiPanelTitle) {
            els.kpiPanelTitle.setAttribute("data-kpi-count", String(summary.count));
            els.kpiPanelTitle.textContent = `整體概況(統計件數${formatNumber(summary.count)}件)`;
        }
        renderKpiDualMetric(els.kpiAvgLtv, summary.principalLtv, summary.householdLtv);
        renderKpiDualMetric(els.kpiAvgCoverage, summary.coverageRatio, summary.householdCoverageRatio);
        if (els.kpiVacancyHomes) {
            els.kpiVacancyHomes.textContent = `${formatNumber(vacancy.vacantHomes)}宅(${formatPercent(vacancy.vacancyRate)})`;
        }
        if (els.kpiUnsoldHomes) {
            els.kpiUnsoldHomes.textContent = `${formatNumber(vacancy.unsoldHomes)}宅(${formatPercent(vacancy.unsoldRate)})`;
        }

        const countyValues = getArrayFilterValues("county");
        const regionLabel = config.forceRegionLabel || (!countyValues.length ? "全台" : countyValues.map(displayCountyName).join("、"));
        els.currentRegionBadge.textContent = `目前區域：${regionLabel}`;
        if (els.dataSourceBadge) {
            els.dataSourceBadge.textContent = state.dataSourceText || "資料來源：-";
        }
        els.filterSummary.textContent = buildFilterSummary();
        applyKpiZeroState();
        syncFilterControls();
        renderActiveFiltersBar();
    }

    function applyKpiZeroState() {
        [
            els.kpiPrincipal, els.kpiApprovedUnused, els.kpiHousehold, els.kpiCollateral, els.kpiCases,
            els.kpiVacancyHomes, els.kpiUnsoldHomes, els.kpiAvgLtv, els.kpiAvgCoverage
        ].forEach(function (el) {
            if (!el) {
                return;
            }
            if (el.classList.contains("kpi-dual-value") || el.classList.contains("kpi-split-value")) {
                return;
            }
            const value = text(el.textContent).replace(/[,()\s宅K仟%]/g, "");
            const num = Number(value || "0");
            const isZero = !Number.isFinite(num) || num === 0;
            el.classList.toggle("is-zero", isZero);
        });
    }

    function renderKpiDualMetric(el, principalValue, householdValue) {
        if (!el) {
            return;
        }
        el.innerHTML = [
            `<span class="kpi-inline-row">` +
                `<em>本金</em><b>${escapeHtml(formatPercent(principalValue))}</b>` +
                ` <i aria-hidden="true">/</i> ` +
                `<em>總歸戶</em><b>${escapeHtml(formatPercent(householdValue))}</b>` +
            `</span>`,
            `<span class="kpi-note-row">(排除未起租)</span>`
        ].join("");
        el.classList.toggle("is-zero", !toNumber(principalValue) && !toNumber(householdValue));
    }

    function renderKpiSplitMoney(el, startedValue, notStartedValue) {
        if (!el) {
            return;
        }
        const totalValue = toNumber(startedValue) + toNumber(notStartedValue);
        const startedText = formatK(startedValue);
        const notStartedText = formatK(notStartedValue);
        const splitTextLength = ("已起租" + startedText + " / 未起租" + notStartedText).length;
        let fitScale = Math.max(0.78, Math.min(1, 31 / Math.max(31, splitTextLength)));
        if (state.amountUnit === "k" && splitTextLength >= 28) {
            fitScale = Math.min(fitScale, 0.84);
        }
        el.style.setProperty("--kpi-fit-scale", fitScale.toFixed(3));
        el.innerHTML = [
            `<span class="kpi-total-row"><b>${escapeHtml(formatK(totalValue))}</b><em>(總額)</em></span>`,
            `<span class="kpi-split-row">` +
                `<em>已起租</em><b>${escapeHtml(startedText)}</b>` +
                ` <i aria-hidden="true">/</i> ` +
                `<em>未起租</em><b>${escapeHtml(notStartedText)}</b>` +
            `</span>`
        ].join("");
        el.classList.toggle("is-zero", !toNumber(startedValue) && !toNumber(notStartedValue));
    }

    function summarizeCurrentKpiRows(rows) {
        const summary = summarizeRows(rows);
        const approvedUnusedStarted = rows.reduce(function (sum, item) {
            return sum + (isAnomalyItem(item) ? 0 : item.approvedUnused);
        }, 0);
        const approvedUnusedNotStarted = summary.approvedUnused - approvedUnusedStarted;
        const startedHouseholdAmount = summary.principalBalance + approvedUnusedStarted;
        summary.approvedUnusedStarted = approvedUnusedStarted;
        summary.approvedUnusedNotStarted = approvedUnusedNotStarted;
        summary.currentKpiHouseholdAmount = startedHouseholdAmount;
        summary.householdLtv = safeDivide(startedHouseholdAmount, summary.collateralValue);
        summary.householdCoverageRatio = safeDivide(summary.collateralValue, startedHouseholdAmount);
        return summary;
    }

    function summarizeRows(rows) {
        const summary = rows.reduce(function (acc, item) {
            acc.principalBalance += item.principalBalance;
            acc.approvedUnused += item.approvedUnused;
            acc.householdAmount += item.householdAmount;
            acc.collateralValue += item.collateralValue;
            acc.count += 1;
            if (isHighRiskItem(item)) {
                acc.highRiskCount += 1;
            }
            return acc;
        }, {
            principalBalance: 0,
            approvedUnused: 0,
            householdAmount: 0,
            collateralValue: 0,
            count: 0,
            highRiskCount: 0
        });
        summary.principalLtv = safeDivide(summary.principalBalance, summary.collateralValue);
        summary.householdLtv = safeDivide(summary.householdAmount, summary.collateralValue);
        summary.coverageRatio = safeDivide(summary.collateralValue, summary.principalBalance);
        summary.householdCoverageRatio = safeDivide(summary.collateralValue, summary.householdAmount);
        return summary;
    }

    function renderMap(rows) {
        if (!state.geoJson) {
            charts.map.setOption({
                title: { text: "尚未載入 GeoJSON", left: "center", top: "middle", textStyle: { color: getChartMutedColor(), fontSize: getMapFontSize(14) } }
            });
            return;
        }

        const countyData = aggregateByCounty(rows);
        const selectedCounties = getArrayFilterValues("county").map(normalizeCountyName);
        const activeDistrictCounty = normalizeCountyName(state.activeDistrictCounty);
        const tooltipCountyData = buildMapTooltipCountyData();
        const data = countyData.map(function (item) {
            const countyKey = normalizeCountyName(item.county);
            const isFocused = selectedCounties.length ? selectedCounties.includes(countyKey) : true;
            const isPartial = selectedCounties.includes(countyKey) && isCountyPartiallySelected(countyKey);
            const isDistrictHover = state.hoveredDistrictCounty && state.hoveredDistrictCounty === countyKey;
            const isActiveDistrictCounty = activeDistrictCounty && activeDistrictCounty === countyKey;
            const focusedStyle = {
                opacity: isFocused ? 1 : 0.35,
                borderColor: isDistrictHover ? "#F97316" : (isActiveDistrictCounty ? "#2563EB" : (isPartial ? "#F59E0B" : "#007AFF")),
                borderWidth: isDistrictHover || isActiveDistrictCounty ? 3 : (isFocused ? 2.2 : 1.3)
            };
            if (isPartial) {
                focusedStyle.areaColor = isDarkTheme() ? "#92400E" : "#FDE68A";
            }
            if (isActiveDistrictCounty) {
                focusedStyle.shadowBlur = 12;
                focusedStyle.shadowColor = "rgba(37, 99, 235, 0.42)";
            }
            return {
                name: state.geoNameByCountyKey.get(item.county) || item.displayName,
                displayName: item.displayName,
                value: item.principalBalance,
                countyKey: item.county,
                collateralValue: item.collateralValue,
                householdAmount: item.householdAmount,
                caseCount: item.count,
                highRiskCount: item.highRiskCount,
                ltv: safeDivide(item.principalBalance, item.collateralValue),
                itemStyle: selectedCounties.length || isDistrictHover || isActiveDistrictCounty ? focusedStyle : undefined
            };
        });
        const max = Math.max(1, ...data.map(function (item) { return item.value; }));
        const topCountyKeys = new Set(countyData.slice(0, 5).map(function (item) { return item.county; }));

        const mapHeatColors = ["#0F6F83", "#2D8C79", "#F2C84B", "#FF7A1F", "#EF4444"];
        charts.map.setOption({
            backgroundColor: "transparent",
            tooltip: {
                trigger: "item",
                textStyle: { fontSize: getMapFontSize(12) },
                formatter: function (params) {
                    const countyKey = normalizeCountyName(
                        params.name ||
                        (params.data && (params.data.countyKey || params.data.displayName || params.data.name)) ||
                        ""
                    );
                    const dataItem = tooltipCountyData.get(countyKey) || params.data;
                    if (!dataItem) {
                        return `${escapeHtml(params.name)}<br>無資料`;
                    }
                    return [
                        `<strong>${escapeHtml(dataItem.displayName || dataItem.name)}</strong>`,
                        `曝險：${formatK(dataItem.value)}`,
                        `LTV：${formatPercent(dataItem.ltv)}`,
                        `案件數：${formatNumber(dataItem.caseCount)}`,
                        `高風險：${formatNumber(dataItem.highRiskCount)}`
                    ].join("<br>");
                }
            },
            visualMap: {
                show: true,
                min: 0,
                max: max,
                right: 24,
                top: 76,
                calculable: true,
                itemHeight: 180,
                text: ["高風險（超額曝險）", "低風險（正常）"],
                inRange: { color: mapHeatColors },
                textStyle: { color: isDarkTheme() ? "#E2E8F0" : "#475569", fontSize: getMapFontSize(12) }
            },
            graphic: [],
            series: [{
                name: "縣市曝險",
                type: "map",
                map: TAIWAN_MAP_NAME,
                roam: false,
                zoom: 1.12,
                center: getCountyMapCenter("南投縣"),
                left: "4%",
                right: "4%",
                top: "4%",
                bottom: "4%",
                nameProperty: "COUNTYNAME",
                selectedMode: "single",
                emphasis: {
                    label: { show: true, color: "#f8fafc", fontSize: getMapFontSize(12), fontWeight: 700 },
                    itemStyle: { areaColor: isDarkTheme() ? "#F59E0B" : "#93C5FD" }
                },
                select: {
                    itemStyle: { borderColor: isDarkTheme() ? "#E2E8F0" : "#0F172A", borderWidth: 2 }
                },
                label: {
                    show: state.showTop5Labels,
                    color: isDarkTheme() ? "#f8fafc" : "#0F172A",
                    fontSize: getMapFontSize(12),
                    fontWeight: 700,
                    formatter: function (params) {
                        const dataItem = params.data;
                        if (!dataItem || !topCountyKeys.has(dataItem.countyKey)) {
                            return "";
                        }
                        return `${dataItem.displayName || dataItem.name}\n${formatShortMoney(dataItem.value)}`;
                    }
                },
                itemStyle: {
                    borderColor: "rgba(148,163,184,0.3)",
                    borderWidth: 1.3,
                    areaColor: isDarkTheme() ? "#1E2538" : "#EFF6FF"
                },
                data: data
            }]
        }, true);
        if (els.kinmenInsetBtn) {
            els.kinmenInsetBtn.classList.toggle("active", selectedCounties.includes(normalizeCountyName("金門縣")));
        }
    }

    function buildMapTooltipCountyData() {
        const rows = applyMatrixFilters(getBaseFilteredRowsFromRawRows(state.rows, ["county"]));
        const tooltipData = new Map();
        aggregateByCounty(rows).forEach(function (item) {
            const geoName = state.geoNameByCountyKey.get(item.county) || item.displayName;
            const dataItem = {
                name: state.geoNameByCountyKey.get(item.county) || item.displayName,
                displayName: item.displayName,
                value: item.principalBalance,
                countyKey: item.county,
                collateralValue: item.collateralValue,
                householdAmount: item.householdAmount,
                caseCount: item.count,
                highRiskCount: item.highRiskCount,
                ltv: safeDivide(item.principalBalance, item.collateralValue)
            };
            [
                item.county,
                item.displayName,
                geoName
            ].forEach(function (name) {
                const key = normalizeCountyName(name);
                if (key) {
                    tooltipData.set(key, dataItem);
                }
            });
        });
        return tooltipData;
    }

    function renderCountyRanking(rows, limit) {
        const topN = limit || Number.MAX_SAFE_INTEGER;
        const drilldownCounty = normalizeCountyName(state.countyRankDrilldownCounty);
        if (drilldownCounty) {
            const districtRows = buildDistrictAmountRows(drilldownCounty, rows).rows.filter(function (item) {
                return toNumber(item.value) > 0;
            }).reverse();
            if (!districtRows.length) {
                renderChartEmptyState(charts.countyRank, `${displayCountyName(drilldownCounty)}暫無行政區授信資料`);
                return;
            }
            const chartDom = charts.countyRank && charts.countyRank.getDom && charts.countyRank.getDom();
            if (chartDom) {
                chartDom.style.height = Math.max(320, districtRows.length * 32 + 88) + "px";
                charts.countyRank.resize();
            }
            const totalValue = districtRows.reduce(function (sum, item) {
                return sum + getExposureSortValue(item);
            }, 0);
            const selectedDistricts = getDistrictFilterValue(drilldownCounty);
            const selectedDistrictName = Array.isArray(selectedDistricts) && selectedDistricts.length === 1 ? selectedDistricts[0] : "";
            renderHorizontalBar(charts.countyRank, districtRows.map(function (item) {
                return {
                    name: item.district,
                    rawName: item.district,
                    principalValue: item.principalValue,
                    approvedUnusedValue: item.approvedUnusedValue,
                    value: getExposureSortValue(item)
                };
            }), getActiveBarGradient("countyRank"), {
                showPercentage: true,
                totalValue: totalValue,
                forceFourTicks: true,
                exposureBasis: state.exposureBasis,
                selectedRawName: selectedDistrictName,
                showTooltip: false
            });
            return;
        }
        const data = aggregateByCountyForExposure(rows).sort(function (a, b) {
            return getExposureSortValue(b) - getExposureSortValue(a);
        }).slice(0, topN).reverse();
        if (!data.length || !data.some(function (item) { return getExposureSortValue(item) > 0; })) {
            renderChartEmptyState(charts.countyRank, "暫無相關曝險數據");
            return;
        }
        const chartDom = charts.countyRank && charts.countyRank.getDom && charts.countyRank.getDom();
        if (chartDom) {
            chartDom.style.height = Math.max(320, data.length * 28 + 88) + "px";
            const analysisCard = chartDom.closest("[data-dev-hide-key='analysisCountyRank']");
            if (analysisCard) {
                analysisCard.classList.add("analysis-scroll-card");
            }
            charts.countyRank.resize();
        }
        const totalValue = rows.reduce(function (sum, item) {
            return sum + getExposureSortValue(item);
        }, 0);
        renderHorizontalBar(charts.countyRank, data.map(function (item) {
            return {
                name: item.displayName,
                rawName: item.displayName,
                principalValue: item.principalBalance,
                approvedUnusedValue: item.approvedUnused,
                value: getExposureSortValue(item)
            };
        }), getActiveBarGradient("countyRank"), {
            showPercentage: true,
            totalValue: totalValue,
            forceFourTicks: true,
            exposureBasis: state.exposureBasis,
            selectedRawName: getSingleFilterValue("county") ? displayCountyName(getSingleFilterValue("county")) : "",
            showTooltip: false
        });
    }

    function renderGroupChart(rows, limit) {
        const topN = limit || 10;
        const grouped = new Map();
        rows.forEach(function (item) {
            if (shouldHideTemporaryCustomerFromTopCreditRank(item)) {
                return;
            }
            const current = grouped.get(item.group) || { name: item.group, principalBalance: 0, approvedUnused: 0 };
            current.principalBalance += item.principalBalance;
            current.approvedUnused += item.approvedUnused;
            grouped.set(item.group, current);
        });
        const data = Array.from(grouped.values()).map(function (item) {
            return {
                name: item.name,
                principalValue: item.principalBalance,
                approvedUnusedValue: item.approvedUnused,
                value: getExposureSortValue(item)
            };
        }).sort(descValue).slice(0, topN).reverse();
        if (!data.length || !data.some(function (item) { return item.value > 0; })) {
            renderChartEmptyState(charts.group, "暫無相關曝險數據");
            return;
        }
        const chartDom = charts.group && charts.group.getDom && charts.group.getDom();
        if (chartDom) {
            chartDom.style.height = Math.max(320, data.length * 28 + 88) + "px";
            charts.group.resize();
        }
        const totalValue = rows.reduce(function (sum, item) {
            return sum + getExposureSortValue(item);
        }, 0);

        renderHorizontalBar(charts.group, data.map(function (item) {
            return {
                name: formatSensitiveEntityName(item.name),
                rawName: item.name,
                principalValue: item.principalValue,
                approvedUnusedValue: item.approvedUnusedValue,
                value: item.value
            };
        }), getActiveBarGradient("group"), {
            forceFourTicks: true,
            showPercentage: true,
            totalValue: totalValue,
            exposureBasis: state.exposureBasis,
            selectedRawName: state.filters.group,
            showTooltip: false
        });
    }

    function renderChartEmptyState(chart, message) {
        if (!chart) {
            return;
        }
        chart.setOption({
            xAxis: { show: false, type: "value" },
            yAxis: { show: false, type: "category", data: [] },
            series: [],
            grid: { left: 0, right: 0, top: 0, bottom: 0 },
            tooltip: { show: false },
            graphic: [
                {
                    type: "group",
                    left: "center",
                    top: "middle",
                    children: [
                        {
                            type: "text",
                            style: {
                                text: "⚠",
                                font: "600 22px sans-serif",
                                fill: getChartSubtleColor(),
                                textAlign: "center",
                                textVerticalAlign: "middle",
                                x: 0,
                                y: -14
                            }
                        },
                        {
                            type: "text",
                            style: {
                                text: message,
                                font: "500 14px sans-serif",
                                fill: getChartMutedColor(),
                                textAlign: "center",
                                textVerticalAlign: "middle",
                                x: 0,
                                y: 12
                            }
                        }
                    ]
                }
            ]
        }, true);
    }

    function isHexColor(value) {
        return typeof value === "string" && /^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(value.trim());
    }

    function renderHorizontalBar(chart, namesOrItems, valuesOrColor, maybeColor, extraOptions) {
        const isFirstItemObject = Array.isArray(namesOrItems) && namesOrItems.length > 0 && typeof namesOrItems[0] === "object" && namesOrItems[0] !== null;
        const isLegacySignature = !isFirstItemObject;
        const isItemsSignature = !isLegacySignature && Array.isArray(valuesOrColor) && valuesOrColor.length === 2 && isHexColor(valuesOrColor[0]) && isHexColor(valuesOrColor[1]);

        let options = {};
        if (isLegacySignature) {
            options = extraOptions || {};
        } else {
            options = maybeColor || {};
        }

        const items = isLegacySignature
            ? namesOrItems.map(function (name, index) {
                const value = toNumber(valuesOrColor[index]);
                return { name: text(name), rawName: text(name), value: value, principalValue: value, approvedUnusedValue: 0 };
            })
            : namesOrItems.map(function (item) {
                if (typeof item === "object" && item !== null) {
                    const principalValue = toNumber(item.principalValue !== undefined ? item.principalValue : item.value);
                    const approvedUnusedValue = toNumber(item.approvedUnusedValue);
                    const value = toNumber(item.value !== undefined ? item.value : principalValue + approvedUnusedValue);
                    return {
                        name: text(item.name),
                        rawName: text(item.rawName || item.name),
                        value: value,
                        principalValue: principalValue,
                        approvedUnusedValue: approvedUnusedValue
                    };
                }
                return { name: text(item), rawName: text(item), value: 0, principalValue: 0, approvedUnusedValue: 0 };
            });
        const names = items.map(function (item) { return item.name; });
        const values = items.map(function (item) { return item.value; });
        const color = isLegacySignature ? maybeColor : (isItemsSignature ? valuesOrColor : valuesOrColor);
        const isHouseholdMode = options.exposureBasis === "household";
        const chartWidth = chart && typeof chart.getWidth === "function" ? chart.getWidth() : 420;
        const compactLayout = chartWidth < 460;
        const maxValue = values.length ? Math.max.apply(null, values) : 0;

        let axisMax;
        let interval;
        if (options.forceFourTicks) {
            axisMax = maxValue > 0 ? maxValue * 1.35 : 12;
            interval = axisMax / 3;
        } else {
            axisMax = maxValue > 0 ? Math.ceil(maxValue * 1.2) : 10;
        }

        const gradient = Array.isArray(color)
            ? new echarts.graphic.LinearGradient(1, 0, 0, 0, [
                { offset: 0, color: color[0] },
                { offset: 1, color: color[1] }
            ])
            : color;
        const approvedUnusedGradient = getApprovedUnusedBarGradient(color);
        const axisColor = getChartTextColor();
        const splitLineColor = getChartSplitLineColor();
        const selectedRawName = text(options.selectedRawName);
        const hasSelectedBar = !!selectedRawName;
        function isSelectedBarItem(item) {
            if (!hasSelectedBar) {
                return false;
            }
            return text(item.rawName) === selectedRawName || text(item.name) === selectedRawName || normalizeCountyName(item.rawName) === normalizeCountyName(selectedRawName);
        }
        function getBarOpacity(item) {
            return hasSelectedBar && !isSelectedBarItem(item) ? 0.24 : 1;
        }
        function formatExposureBarLabel(value) {
            const valStr = formatShortMoney(value);
            if (options.showPercentage && options.totalValue) {
                return valStr + " (" + formatPercent(safeDivide(value, options.totalValue)) + ")";
            }
            return valStr;
        }

        const top3Colors = [
            ["#ffd269", "#ff9f1c"],
            ["#ffdf85", "#ffa62b"],
            ["#ffe8a3", "#ffb732"]
        ];

        if (chart && typeof chart.dispatchAction === "function") {
            chart.dispatchAction({ type: "hideTip" });
        }
        const chartZr = chart && typeof chart.getZr === "function" ? chart.getZr() : null;
        if (chartZr && options.showTooltip === false && !chart.__hideTooltipOnMove) {
            chart.__hideTooltipOnMove = function () {
                clearChartHoverAnnotations(chart);
            };
            chartZr.on("mousemove", chart.__hideTooltipOnMove);
        } else if (chartZr && options.showTooltip !== false && chart.__hideTooltipOnMove) {
            chartZr.off("mousemove", chart.__hideTooltipOnMove);
            chart.__hideTooltipOnMove = null;
        }
        chart.setOption({
            graphic: [],
            grid: { left: "5%", right: 92, top: 18, bottom: 28, containLabel: true },
            tooltip: options.showTooltip === false ? {
                show: false,
                trigger: "none",
                alwaysShowContent: false,
                axisPointer: { show: false, type: "none" }
            } : {
                trigger: "axis",
                axisPointer: { type: "shadow" },
                formatter: function (params) {
                    const item = params[0];
                    const data = item.data || {};
                    const rawName = data.rawName ? data.rawName : item.name;
                    const displayName = state.customerMasked ? "****" : rawName;
                    if (isHouseholdMode) {
                        const principalValue = toNumber(data.principalValue);
                        const approvedUnusedValue = toNumber(data.approvedUnusedValue);
                        const totalValue = principalValue + approvedUnusedValue;
                        const unusedRatio = safeDivide(approvedUnusedValue, totalValue);
                        const pctLine = options.showPercentage && options.totalValue
                            ? `<br>佔目前範圍：${formatPercent(safeDivide(totalValue, options.totalValue))}`
                            : "";
                        return "[" + escapeHtml(displayName) + "]<br>" +
                            "本金餘額: " + formatShortMoney(principalValue) + "<br>" +
                            "已准未用: " + formatShortMoney(approvedUnusedValue) + "<br>" +
                            "總歸戶金額: " + formatShortMoney(totalValue) + "<br>" +
                            "已准未用占比: " + formatPercent(unusedRatio) + pctLine;
                    }
                    const valStr = formatShortMoney(item.value);
                    if (options.showPercentage && options.totalValue) {
                        const pct = formatPercent(safeDivide(item.value, options.totalValue));
                        return "[" + escapeHtml(displayName) + "]<br>本金餘額: " + valStr + " (佔全台比率: " + pct + ")";
                    }
                    return "[" + escapeHtml(displayName) + "]<br>本金餘額: " + valStr;
                }
            },
            xAxis: {
                type: "value",
                min: 0,
                max: axisMax,
                interval: options.forceFourTicks ? interval : undefined,
                splitNumber: options.forceFourTicks ? 3 : (compactLayout ? 2 : 4),
                axisLine: { show: false },
                axisTick: { show: false },
                axisLabel: {
                    formatter: formatAxisMoney,
                    color: axisColor,
                    fontSize: getFSD(compactLayout ? 10 : 12),
                    hideOverlap: true,
                    showMaxLabel: true
                },
                splitLine: { lineStyle: { color: splitLineColor } }
            },
            yAxis: {
                type: "category",
                data: names,
                axisLine: { show: false },
                axisTick: { show: false },
                axisLabel: {
                    width: compactLayout ? 70 : 100,
                    overflow: "truncate",
                    color: axisColor,
                    fontSize: getFSD(12),
                    margin: 8
                }
            },
            series: isHouseholdMode ? [
                {
                    name: "本金餘額",
                    type: "bar",
                    tooltip: options.showTooltip === false ? { show: false } : undefined,
                    stack: "household",
                    data: items.map(function (item) {
                        const totalValue = item.principalValue + item.approvedUnusedValue;
                        return {
                            name: item.name,
                            rawName: item.rawName,
                            value: item.principalValue,
                            principalValue: item.principalValue,
                            approvedUnusedValue: item.approvedUnusedValue,
                            totalValue: totalValue,
                            itemStyle: {
                                color: gradient,
                                opacity: getBarOpacity(item),
                                borderRadius: item.approvedUnusedValue > 0 ? [0, 0, 0, 0] : [0, 5, 5, 0]
                            },
                            label: {
                                show: item.approvedUnusedValue <= 0 && totalValue > 0,
                                position: "right",
                                color: getChartTextColor(),
                                fontWeight: "bold",
                                fontSize: getFSD(compactLayout ? 10 : 11),
                                distance: 8,
                                formatter: function () {
                                    return formatExposureBarLabel(totalValue);
                                }
                            }
                        };
                    }),
                    barMaxWidth: 18
                },
                {
                    name: "已准未用",
                    type: "bar",
                    tooltip: options.showTooltip === false ? { show: false } : undefined,
                    stack: "household",
                    data: items.map(function (item) {
                        const totalValue = item.principalValue + item.approvedUnusedValue;
                        return {
                            name: item.name,
                            rawName: item.rawName,
                            value: item.approvedUnusedValue,
                            principalValue: item.principalValue,
                            approvedUnusedValue: item.approvedUnusedValue,
                            totalValue: totalValue,
                            itemStyle: {
                                color: approvedUnusedGradient,
                                opacity: getBarOpacity(item),
                                borderColor: isDarkTheme() ? "rgba(226, 232, 240, 0.22)" : "rgba(37, 99, 235, 0.18)",
                                borderWidth: item.approvedUnusedValue > 0 ? 1 : 0,
                                borderType: "dashed",
                                borderRadius: [0, 5, 5, 0]
                            },
                            label: {
                                show: item.approvedUnusedValue > 0 && totalValue > 0,
                                position: "right",
                                color: getChartTextColor(),
                                fontWeight: "bold",
                                fontSize: getFSD(compactLayout ? 10 : 11),
                                distance: 8,
                                formatter: function () {
                                    return formatExposureBarLabel(totalValue);
                                }
                            }
                        };
                    }),
                    barMaxWidth: 18
                }
            ] : [{
                type: "bar",
                tooltip: options.showTooltip === false ? { show: false } : undefined,
                data: items.map(function (item, idx) {
                    let itemColor = gradient;
                    if (options.highlightTop3) {
                        const rank = items.length - 1 - idx;
                        if (rank >= 0 && rank < 3) {
                            const c = top3Colors[rank];
                            itemColor = new echarts.graphic.LinearGradient(1, 0, 0, 0, [
                                { offset: 0, color: c[0] },
                                { offset: 1, color: c[1] }
                            ]);
                        }
                    }
                    return {
                        name: item.name,
                        rawName: item.rawName,
                        value: item.value,
                        itemStyle: { color: itemColor, opacity: getBarOpacity(item) }
                    };
                }),
                barMaxWidth: 18,
                itemStyle: { borderRadius: [0, 5, 5, 0] },
                label: {
                    show: true,
                    position: "right",
                    color: getChartTextColor(),
                    fontWeight: "bold",
                    fontSize: getFSD(compactLayout ? 10 : 11),
                    distance: 8,
                    formatter: function (params) {
                        return formatExposureBarLabel(params.value);
                    }
                }
            }]
        }, true);
        if (options.showTooltip === false) {
            clearChartHoverAnnotations(chart);
        }
    }

    function clearChartHoverAnnotations(chart) {
        if (!chart) {
            return;
        }
        if (typeof chart.dispatchAction === "function") {
            chart.dispatchAction({ type: "hideTip" });
        }
        const dom = typeof chart.getDom === "function" ? chart.getDom() : null;
        if (dom) {
            dom.removeAttribute("title");
            dom.querySelectorAll("[title]").forEach(function (node) {
                node.removeAttribute("title");
            });
            dom.querySelectorAll(".echarts-tooltip").forEach(function (node) {
                node.style.display = "none";
                node.style.pointerEvents = "none";
            });
        }
        document.querySelectorAll(".echarts-tooltip").forEach(function (node) {
            node.style.display = "none";
            node.style.pointerEvents = "none";
        });
    }

    function estimateChartTextWidth(value, fontSize) {
        return text(value).split("").reduce(function (total, char) {
            return total + (char.charCodeAt(0) > 255 ? fontSize : fontSize * 0.58);
        }, 0);
    }

    function renderLtvDistribution(rows) {
        const counts = new Map(LTV_BUCKETS.map(function (bucket) {
            return [bucket.id, 0];
        }));
        rows.forEach(function (item) {
            if (!item.principalBalance || !item.ltvBucket) {
                return;
            }
            counts.set(item.ltvBucket, (counts.get(item.ltvBucket) || 0) + 1);
        });
        charts.ltv.setOption({
            grid: { left: "5%", right: "5%", top: 18, bottom: 28, containLabel: true },
            tooltip: { trigger: "axis" },
            xAxis: {
                type: "category",
                data: LTV_BUCKETS.map(function (bucket) { return bucket.label; }),
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(10), interval: 0 }
            },
            yAxis: {
                type: "value",
                minInterval: 1,
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(10) },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            series: [{
                type: "bar",
                barMaxWidth: 42,
                data: LTV_BUCKETS.map(function (bucket) {
                    return {
                        value: counts.get(bucket.id) || 0,
                        bucketId: bucket.id,
                        itemStyle: {
                            color: getLtvBucketColor(bucket.id),
                            opacity: getArrayFilterValues("ltvBucket").length && !getArrayFilterValues("ltvBucket").includes(bucket.id) ? 0.24 : 1,
                            borderColor: getLtvBucketBorderColor(bucket.id),
                            borderWidth: 1,
                            borderRadius: [6, 6, 0, 0]
                        }
                    };
                })
            }]
        }, true);
    }

    function renderCoverageDistribution(rows) {
        const counts = new Map(COVERAGE_BUCKETS.map(function (bucket) {
            return [bucket.id, 0];
        }));

        const isCountyMode = state.coverageMode === "county";

        if (isCountyMode) {
            // 縣市模式：依縣市的平均覆蓋率來集計
            const countyGroups = aggregateByCounty(rows);
            countyGroups.forEach(function (countyItem) {
                if (countyItem.principalBalance <= 0) {
                    return;
                }
                const avgCoverage = safeDivide(countyItem.collateralValue, countyItem.principalBalance);
                const bucket = getCoverageBucketByRatio(avgCoverage);
                counts.set(bucket.id, (counts.get(bucket.id) || 0) + 1);
            });
        } else {
            // 案件模式：直接依案件的覆蓋率來集計
            rows.forEach(function (item) {
                if (item.principalBalance <= 0) {
                    return;
                }
                const bucket = getCoverageBucketByRatio(item.coverageRatio || 0);
                counts.set(bucket.id, (counts.get(bucket.id) || 0) + 1);
            });
        }

        const maxCount = COVERAGE_BUCKETS.reduce(function (max, bucket) {
            return Math.max(max, counts.get(bucket.id) || 0);
        }, 0);

        const bucketColors = {
            "under-100": ["#ff7b7b", "#d32f2f"], // 紅色
            "100-120": ["#ffb74d", "#f57c00"],   // 橘色
            "120-150": ["#4fc3f7", "#0288d1"],   // 藍色
            "150+": ["#81c784", "#388e3c"]        // 綠色
        };

        charts.coverage.setOption({
            grid: { left: "5%", right: "5%", top: 40, bottom: 30, containLabel: true },
            tooltip: { trigger: "axis" },
            xAxis: {
                type: "category",
                data: COVERAGE_BUCKETS.map(function (bucket) { return bucket.label; }),
                axisLabel: { color: getChartTextColor(), interval: 0, fontSize: getFSD(11) }
            },
            yAxis: {
                type: "value",
                max: maxCount > 0 ? Math.ceil(maxCount * 1.2) : 10,
                minInterval: 1,
                name: isCountyMode ? "縣市數" : "案件數",
                nameTextStyle: { color: getChartMutedColor(), padding: [0, 0, 0, -8], fontSize: getFSD(11) },
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(10) },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            series: [{
                type: "bar",
                barMaxWidth: 36,
                data: COVERAGE_BUCKETS.map(function (bucket) {
                    const value = counts.get(bucket.id) || 0;
                    const colors = bucketColors[bucket.id];
                    return {
                        bucketId: bucket.id,
                        value: value,
                        itemStyle: {
                            color: value > 0
                                ? new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                                    { offset: 0, color: colors[0] },
                                    { offset: 1, color: colors[1] }
                                ])
                                : getChartEmptyBarColor(),
                            opacity: state.filters.coverageBucket && state.filters.coverageBucket !== bucket.id ? 0.24 : 1,
                            borderRadius: [4, 4, 0, 0]
                        }
                    };
                }),
                label: {
                    show: true,
                    position: "inside",
                    color: "#ffffff",
                    fontWeight: "bold",
                    fontSize: getFSD(11),
                    textShadowColor: "rgba(0,0,0,0.5)",
                    textShadowBlur: 2,
                    formatter: function (params) {
                        return params.value > 0 ? formatNumber(params.value) : "";
                    }
                }
            }]
        }, true);
    }

    function renderConcentration(rows) {
        const counties = aggregateByCounty(rows).slice(0, 5);
        const total = counties.reduce(function (sum, item) {
            return sum + item.principalBalance;
        }, 0);
        els.concentrationList.innerHTML = counties.map(function (item) {
            const ratio = safeDivide(item.principalBalance, total);
            return `
                <div class="concentration-item">
                    <strong>${escapeHtml(item.displayName)}</strong>
                    <div class="meter"><span style="width:${Math.min(100, ratio * 100).toFixed(2)}%"></span></div>
                    <span>${formatPercent(ratio)}</span>
                </div>
            `;
        }).join("") || "<p class=\"table-summary\">目前沒有可顯示的縣市曝險。</p>";
    }

    function renderCountySummaryTable(rows) {
        if (!els.countySummaryRows) {
            return;
        }
        const summaries = buildCountySummaryRows(rows).filter(function (item) {
            return item.customerCount > 0;
        });
        if (!summaries.length) {
            els.countySummaryRows.innerHTML = "<tr><td colspan=\"12\">目前篩選條件下無縣市資料。</td></tr>";
            return;
        }
        const totalRemaining = summaries.reduce(function (sum, item) {
            return sum + item.remainingPrincipal;
        }, 0);

        const sortedSummaries = sortCountySummaryRows(summaries, totalRemaining);
        const bodyRows = sortedSummaries.map(function (item) {
            const topRatio = safeDivide(item.maxCustomerRemaining, item.remainingPrincipal);
            const regionShare = safeDivide(item.remainingPrincipal, totalRemaining);
            const ltv = safeDivide(item.remainingPrincipal, item.collateralTotal);
            const coverageRatio = safeDivide(item.collateralTotal, item.remainingPrincipal);
            const ltvText = item.remainingPrincipal > 0 ? formatPercent(ltv) : "-";
            const coverageText = item.remainingPrincipal > 0 ? formatPercent(coverageRatio) : "-";
            return `
                <tr>
                    <td>${escapeHtml(formatCountySummaryName(item.county))}</td>
                    <td>${formatNumber(item.customerCount)}</td>
                    <td><span class="customer-name-ellipsis" title="${escapeHtml(getCustomerTitle(item.maxCustomerName || "-"))}">${escapeHtml(formatCustomerDisplayName(item.maxCustomerName || "-"))}</span></td>
                    <td>${formatPercent(topRatio)}</td>
                    <td>${formatK(item.approvedLimitTotal)}</td>
                    <td>${formatK(item.disbursedTotal)}</td>
                    <td>${formatK(item.approvedUnusedTotal)}</td>
                    <td>${formatK(item.remainingPrincipal)}</td>
                    <td>${formatPercent(regionShare)}</td>
                    <td>${formatK(item.collateralTotal)}</td>
                    <td>${ltvText}</td>
                    <td>${coverageText}</td>
                </tr>
            `;
        }).join("");

        const totalRow = buildCountySummaryTotalRow(sortedSummaries, totalRemaining);
        els.countySummaryRows.innerHTML = bodyRows + totalRow;
    }

    function buildCountySummaryRows(rows) {
        const byCounty = new Map();
        COMMON_COUNTY_ORDER.forEach(function (countyName) {
            byCounty.set(countyName, {
                county: countyName,
                customerCount: 0,
                maxCustomerName: "",
                maxCustomerRemaining: 0,
                approvedLimitTotal: 0,
                disbursedTotal: 0,
                approvedUnusedTotal: 0,
                remainingPrincipal: 0,
                collateralTotal: 0
            });
        });

        rows.forEach(function (item) {
            const county = displayCountyName(item.county) || "Unassigned";
            if (!byCounty.has(county)) {
                byCounty.set(county, {
                    county: county,
                    customerCount: 0,
                    maxCustomerName: "",
                    maxCustomerRemaining: 0,
                    approvedLimitTotal: 0,
                    disbursedTotal: 0,
                    approvedUnusedTotal: 0,
                    remainingPrincipal: 0,
                    collateralTotal: 0
                });
            }
            const current = byCounty.get(county);
            current.customerCount += 1;
            current.approvedLimitTotal += item.householdAmount;
            current.disbursedTotal += item.principalBalance;
            current.approvedUnusedTotal += item.approvedUnused;
            current.remainingPrincipal += item.principalBalance;
            current.collateralTotal += item.collateralValue;
            if (item.principalBalance >= current.maxCustomerRemaining) {
                current.maxCustomerRemaining = item.principalBalance;
                current.maxCustomerName = item.applicant;
            }
        });

        const orderKeys = COMMON_COUNTY_ORDER.map(function (county) {
            return normalizeCountyName(county);
        });
        return Array.from(byCounty.values()).sort(function (a, b) {
            const ai = orderKeys.indexOf(normalizeCountyName(a.county));
            const bi = orderKeys.indexOf(normalizeCountyName(b.county));
            const av = ai < 0 ? 999 : ai;
            const bv = bi < 0 ? 999 : bi;
            if (av !== bv) {
                return av - bv;
            }
            return localeCompare(a.county, b.county);
        });
    }

    function buildCountySummaryTotalRow(rows, totalRemaining) {
        const total = rows.reduce(function (acc, item) {
            acc.customerCount += item.customerCount;
            acc.approvedLimitTotal += item.approvedLimitTotal;
            acc.disbursedTotal += item.disbursedTotal;
            acc.approvedUnusedTotal += item.approvedUnusedTotal;
            acc.remainingPrincipal += item.remainingPrincipal;
            acc.collateralTotal += item.collateralTotal;
            if (item.maxCustomerRemaining > acc.maxCustomerRemaining) {
                acc.maxCustomerRemaining = item.maxCustomerRemaining;
                acc.maxCustomerName = item.maxCustomerName;
            }
            return acc;
        }, {
            customerCount: 0,
            maxCustomerName: "",
            maxCustomerRemaining: 0,
            approvedLimitTotal: 0,
            disbursedTotal: 0,
            approvedUnusedTotal: 0,
            remainingPrincipal: 0,
            collateralTotal: 0
        });
        const topRatio = safeDivide(total.maxCustomerRemaining, total.remainingPrincipal);
        const regionShare = safeDivide(total.remainingPrincipal, totalRemaining);
        const ltv = safeDivide(total.remainingPrincipal, total.collateralTotal);
        const coverageRatio = safeDivide(total.collateralTotal, total.remainingPrincipal);
        const ltvText = total.remainingPrincipal > 0 ? formatPercent(ltv) : "-";
        const coverageText = total.remainingPrincipal > 0 ? formatPercent(coverageRatio) : "-";
        return `
            <tr>
                <td>Total</td>
                <td>${formatNumber(total.customerCount)}</td>
                <td><span class="customer-name-ellipsis" title="${escapeHtml(getCustomerTitle(total.maxCustomerName || "-"))}">${escapeHtml(formatCustomerDisplayName(total.maxCustomerName || "-"))}</span></td>
                <td>${formatPercent(topRatio)}</td>
                <td>${formatK(total.approvedLimitTotal)}</td>
                <td>${formatK(total.disbursedTotal)}</td>
                <td>${formatK(total.approvedUnusedTotal)}</td>
                <td>${formatK(total.remainingPrincipal)}</td>
                <td>${formatPercent(regionShare)}</td>
                <td>${formatK(total.collateralTotal)}</td>
                <td>${ltvText}</td>
                <td>${coverageText}</td>
            </tr>
        `;
    }

    function renderTable(rows) {
        const caseDetailRows = filterCaseDetailRows(rows);
        const searchedRows = filterTableRowsBySearch(caseDetailRows);
        const tableRows = sortCaseRows(searchedRows).slice(0, 300);

        const searchText = text(state.tableSearch);
        const searchNote = searchText ? `，搜尋「${searchText}」命中 ${formatNumber(searchedRows.length)} 筆` : "";
        els.tableSummary.textContent = `顯示 ${formatNumber(tableRows.length)} 筆案件${searchNote}，符合篩選共 ${formatNumber(caseDetailRows.length)} 筆。`;

        els.riskRows.innerHTML = tableRows.map(function (item) {
            const notes = getCaseNotes(item);
            const customerActiveClass = state.filters.customer === item.applicant ? " active" : "";
            const customerName = text(item.applicant);
            return `
                <tr>
                    <td><button class="customer-filter-button${customerActiveClass}" type="button" data-customer-filter="${escapeHtml(customerName)}" title="${escapeHtml(getCustomerTitle(customerName))}"><span class="customer-name-ellipsis">${escapeHtml(formatCustomerDisplayName(customerName))}</span></button></td>
                    <td>${escapeHtml(formatCountySummaryName(displayCountyName(item.county)))}</td>
                    <td>${escapeHtml(item.industry)}</td>
                    <td>${formatCaseNoDisplayForTable(item.caseId || "-")}</td>
                    <td>${formatK(item.principalBalance)}</td>
                    <td>${formatK(item.approvedUnused)}</td>
                    <td>${formatK(item.householdAmount)}</td>
                    <td>${formatK(item.collateralValue)}</td>
                    <td>${item.principalBalance > 0 ? formatPercent(item.principalLtv) : "-"}</td>
                    <td>${item.principalBalance > 0 ? formatPercent(item.coverageRatio) : "-"}</td>
                    <td><span class="quadrant-tag ${escapeHtml(item.matrix.supply.className)}">${escapeHtml(item.matrix.supply.label)}</span></td>
                    <td>${notes.map(function (note) {
                        return `<span class="note-tag">${escapeHtml(note)}</span>`;
                    }).join("") || "-"}</td>
                </tr>
            `;
        }).join("") || "<tr><td colspan=\"12\">目前沒有符合條件的案件。</td></tr>";
    }

    function exportCaseDetailExcel() {
        const rows = sortCaseRows(filterTableRowsBySearch(filterCaseDetailRows(getFilteredRows())));
        if (!rows.length) {
            showNotice("目前沒有可匯出的案件明細。");
            return;
        }
        const exportRows = rows.map(function (item) {
            const notes = getCaseNotes(item);
            return {
                "客戶": formatExportCustomerName(item.applicant),
                "縣市": formatCountySummaryName(displayCountyName(item.county)),
                "業別": text(item.industry),
                "初審單號": formatCaseNoDisplay(item.caseId || "-"),
                [`\u672c\u91d1\u9918\u984d(${getExportAmountUnitLabel()})`]: formatExportAmount(item.principalBalance),
                [`\u5df2\u51c6\u672a\u7528(${getExportAmountUnitLabel()})`]: formatExportAmount(item.approvedUnused),
                [`\u7e3d\u6b78\u6236\u91d1\u984d(${getExportAmountUnitLabel()})`]: formatExportAmount(item.householdAmount),
                [`\u64d4\u4fdd\u503c(${getExportAmountUnitLabel()})`]: formatExportAmount(item.collateralValue),
                "LTV": item.principalBalance > 0 ? formatPercent(item.principalLtv) : "-",
                "覆蓋率": item.principalBalance > 0 ? formatPercent(item.coverageRatio) : "-",
                "象限標籤": item.matrix && item.matrix.supply ? text(item.matrix.supply.label) : "",
                "備註": notes.join("、") || "-"
            };
        });
        if (downloadWorkbook(exportRows, "案件明細表", buildExportFileName("案件明細表"))) {
            showNotice(`已匯出案件明細表 ${formatNumber(exportRows.length)} 筆。`);
        }
    }

    function exportCountySummaryExcel() {
        const summaries = buildCountySummaryRows(getFilteredRows()).filter(function (item) {
            return item.customerCount > 0;
        });
        if (!summaries.length) {
            showNotice("目前沒有可匯出的縣市彙總資料。");
            return;
        }
        const totalRemaining = summaries.reduce(function (sum, item) {
            return sum + item.remainingPrincipal;
        }, 0);
        const sortedSummaries = sortCountySummaryRows(summaries, totalRemaining);
        const total = buildCountySummaryExportTotal(sortedSummaries);
        const exportRows = sortedSummaries.map(function (item) {
            return buildCountySummaryExportRow(item, totalRemaining);
        });
        exportRows.push(buildCountySummaryExportRow(total, totalRemaining));
        if (downloadWorkbook(exportRows, "縣市彙總表", buildExportFileName("縣市彙總表"))) {
            showNotice(`已匯出縣市彙總表 ${formatNumber(sortedSummaries.length)} 個縣市。`);
        }
    }

    function buildCountySummaryExportRow(item, totalRemaining) {
        const topRatio = safeDivide(item.maxCustomerRemaining, item.remainingPrincipal);
        const regionShare = safeDivide(item.remainingPrincipal, totalRemaining);
        const ltv = safeDivide(item.remainingPrincipal, item.collateralTotal);
        const coverageRatio = safeDivide(item.collateralTotal, item.remainingPrincipal);
        return {
            "縣市": text(item.county),
            "客戶計數": toNumber(item.customerCount),
            "最大客戶": formatExportCustomerName(item.maxCustomerName || "-"),
            "最大客戶本金佔比": formatPercent(topRatio),
            [`\u6838\u51c6\u984d\u5ea6(${getExportAmountUnitLabel()})`]: formatExportAmount(item.approvedLimitTotal),
            [`\u5df2\u64a5\u91d1\u984d(${getExportAmountUnitLabel()})`]: formatExportAmount(item.disbursedTotal),
            [`\u5df2\u51c6\u672a\u7528(${getExportAmountUnitLabel()})`]: formatExportAmount(item.approvedUnusedTotal),
            [`\u672c\u91d1\u9918\u984d(${getExportAmountUnitLabel()})`]: formatExportAmount(item.remainingPrincipal),
            "本金餘額結構比": formatPercent(regionShare),
            [`\u64d4\u4fdd\u503c(${getExportAmountUnitLabel()})`]: formatExportAmount(item.collateralTotal),
            "LTV": item.remainingPrincipal > 0 ? formatPercent(ltv) : "-",
            "擔保率": item.remainingPrincipal > 0 ? formatPercent(coverageRatio) : "-"
        };
    }

    function buildCountySummaryExportTotal(rows) {
        return rows.reduce(function (acc, item) {
            acc.customerCount += item.customerCount;
            acc.approvedLimitTotal += item.approvedLimitTotal;
            acc.disbursedTotal += item.disbursedTotal;
            acc.approvedUnusedTotal += item.approvedUnusedTotal;
            acc.remainingPrincipal += item.remainingPrincipal;
            acc.collateralTotal += item.collateralTotal;
            if (item.maxCustomerRemaining > acc.maxCustomerRemaining) {
                acc.maxCustomerRemaining = item.maxCustomerRemaining;
                acc.maxCustomerName = item.maxCustomerName;
            }
            return acc;
        }, {
            county: "Total",
            customerCount: 0,
            maxCustomerName: "",
            maxCustomerRemaining: 0,
            approvedLimitTotal: 0,
            disbursedTotal: 0,
            approvedUnusedTotal: 0,
            remainingPrincipal: 0,
            collateralTotal: 0
        });
    }

    function downloadWorkbook(rows, sheetName, fileName) {
        if (!window.XLSX || !window.XLSX.utils || typeof window.XLSX.writeFile !== "function") {
            showNotice("找不到 Excel 匯出套件，請確認 libs/xlsx.full.min.js 已載入。");
            return false;
        }
        try {
            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.json_to_sheet(rows);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
            XLSX.writeFile(workbook, fileName);
            return true;
        } catch (_error) {
            showNotice("Excel 匯出失敗，請確認瀏覽器下載權限。");
            return false;
        }
    }

    function buildExportFileName(label) {
        return `擔保地圖_${label}_${getExportTimestamp()}.xlsx`;
    }

    function formatExportCustomerName(value) {
        const normalized = text(value);
        if (!normalized || normalized === "-") {
            return normalized || "-";
        }
        return state.customerMasked ? "****" : normalized;
    }

    function getExportTimestamp() {
        const now = new Date();
        function pad(value) {
            return String(value).padStart(2, "0");
        }
        return [
            now.getFullYear(),
            pad(now.getMonth() + 1),
            pad(now.getDate()),
            "_",
            pad(now.getHours()),
            pad(now.getMinutes()),
            pad(now.getSeconds())
        ].join("");
    }

    function initTableSorting() {
        setupTableHeaderSortMeta("case", els.caseTableWrap);
        setupTableHeaderSortMeta("county", els.countySummaryWrap);
        refreshTableSortIndicators();
    }

    function setupTableHeaderSortMeta(tableType, wrapEl) {
        if (!wrapEl) {
            return;
        }
        const headers = wrapEl.querySelectorAll("thead th");
        headers.forEach(function (th, index) {
            th.setAttribute("data-sort-table", tableType);
            th.setAttribute("data-sort-index", String(index));
            th.setAttribute("title", "點擊排序");
            th.classList.add("sortable-th");
        });
    }

    function toggleTableSort(tableType, column) {
        const sortState = tableType === "case" ? state.caseSort : state.countySort;
        if (sortState.column === column) {
            sortState.direction = sortState.direction === "asc" ? "desc" : "asc";
        } else {
            sortState.column = column;
            sortState.direction = "asc";
        }
        refreshTableSortIndicators();
    }

    function refreshTableSortIndicators() {
        updateSortHeaderForTable("case", els.caseTableWrap, state.caseSort);
        updateSortHeaderForTable("county", els.countySummaryWrap, state.countySort);
    }

    function updateSortHeaderForTable(tableType, wrapEl, sortState) {
        if (!wrapEl) {
            return;
        }
        const headers = wrapEl.querySelectorAll(`thead th[data-sort-table='${tableType}']`);
        headers.forEach(function (th) {
            const index = toNumber(th.getAttribute("data-sort-index"));
            const isActive = index === sortState.column;
            th.classList.toggle("sort-active", isActive);
            th.setAttribute("data-sort-dir", isActive ? sortState.direction : "");
            th.setAttribute("title", isActive ? `目前：${sortState.direction === "asc" ? "升冪" : "降冪"}（點擊切換）` : "點擊排序");
        });
    }

    function sortCaseRows(rows) {
        const sortState = state.caseSort;
        const cloned = rows.slice();
        return cloned.sort(function (a, b) {
            const av = getCaseSortValue(a, sortState.column);
            const bv = getCaseSortValue(b, sortState.column);
            return compareSortValue(av, bv, sortState.direction);
        });
    }

    function getCaseSortValue(item, column) {
        switch (column) {
            case 0: return text(item.applicant);
            case 1: return text(displayCountyName(item.county));
            case 2: return text(item.industry);
            case 3: return text(item.caseId || "");
            case 4: return toNumber(item.principalBalance);
            case 5: return toNumber(item.approvedUnused);
            case 6: return toNumber(item.householdAmount);
            case 7: return toNumber(item.collateralValue);
            case 8: return toNumber(item.principalLtv);
            case 9: return toNumber(item.coverageRatio);
            case 10: return text(item.matrix && item.matrix.supply ? item.matrix.supply.label : "");
            case 11: return text(getCaseNotes(item).join(","));
            default: return toNumber(item.principalLtv);
        }
    }

    function sortCountySummaryRows(rows, totalRemaining) {
        const sortState = state.countySort;
        const cloned = rows.slice();
        return cloned.sort(function (a, b) {
            const av = getCountySortValue(a, sortState.column, totalRemaining);
            const bv = getCountySortValue(b, sortState.column, totalRemaining);
            return compareSortValue(av, bv, sortState.direction);
        });
    }

    function getCountySortValue(item, column, totalRemaining) {
        switch (column) {
            case 0: return text(item.county);
            case 1: return toNumber(item.customerCount);
            case 2: return text(item.maxCustomerName || "");
            case 3: return safeDivide(item.maxCustomerRemaining, item.remainingPrincipal);
            case 4: return toNumber(item.approvedLimitTotal);
            case 5: return toNumber(item.disbursedTotal);
            case 6: return toNumber(item.approvedUnusedTotal);
            case 7: return toNumber(item.remainingPrincipal);
            case 8: return safeDivide(item.remainingPrincipal, totalRemaining);
            case 9: return toNumber(item.collateralTotal);
            case 10: return safeDivide(item.remainingPrincipal, item.collateralTotal);
            case 11: return safeDivide(item.collateralTotal, item.remainingPrincipal);
            default: return text(item.county);
        }
    }

    function compareSortValue(a, b, direction) {
        const factor = direction === "asc" ? 1 : -1;
        const aIsNumber = typeof a === "number";
        const bIsNumber = typeof b === "number";
        if (aIsNumber && bIsNumber) {
            return (a - b) * factor;
        }
        return localeCompare(String(a || ""), String(b || "")) * factor;
    }
    function toggleSection(wrapId, buttonId) {
        const wrap = els[wrapId];
        const button = els[buttonId];
        if (!wrap || !button) {
            return;
        }
        const collapsed = wrap.style.display === "none";
        wrap.style.display = collapsed ? "" : "none";
        button.textContent = collapsed ? "摺疊" : "展開";
        button.classList.toggle("active", !collapsed);
    }

    function initCollapsedSections() {
        if (els.countyDrawerPanel && els.toggleCountyDrawer) {
            els.countyDrawerPanel.hidden = true;
            els.countyDrawerPanel.classList.remove("open");
            els.countyDrawerPanel.setAttribute("aria-hidden", "true");
            els.toggleCountyDrawer.classList.remove("active");
        }
        if (els.filterPanel) {
            if (isInlineFilterPanel()) {
                els.filterPanel.hidden = state.timelineMode === "timeline";
            } else {
                els.filterPanel.hidden = true;
                els.filterPanel.classList.remove("open");
            }
        }
        if (els.countySummaryWrap && els.toggleCountySummary) {
            els.countySummaryWrap.style.display = "";
            els.toggleCountySummary.textContent = "摺疊";
            els.toggleCountySummary.classList.add("active");
        }
        if (els.caseTableWrap && els.toggleCaseTable) {
            els.caseTableWrap.style.display = "";
            els.toggleCaseTable.textContent = "摺疊";
            els.toggleCaseTable.classList.add("active");
        }
    }

    function syncDrawerBackdrop() {
        var caseOpen = !!(els.caseDrawerPanel && els.caseDrawerPanel.classList.contains("open"));
        var filterOpen = !!(els.filterPanel && els.filterPanel.classList.contains("open"));
        var countyOpen = !!(els.countyDrawerPanel && els.countyDrawerPanel.classList.contains("open"));
        var analysisOpen = !!(els.analysisDrawerPanel && els.analysisDrawerPanel.classList.contains("open"));
        var anyOpen = caseOpen || filterOpen || countyOpen || analysisOpen;
        if (els.drawerBackdrop) {
            els.drawerBackdrop.hidden = !anyOpen;
        }
        document.body.classList.toggle("drawer-open", anyOpen);
    }

    function toggleCaseDrawer() {
        if (!els.caseDrawerPanel) {
            return;
        }
        setCaseDrawer(!els.caseDrawerPanel.classList.contains("open"));
    }

    function setCaseDrawer(open) {
        if (!els.caseDrawerPanel) {
            return;
        }
        if (open) {
            setFilterDrawer(false);
            setCountyDrawer(false);
            setAnalysisDrawer(false);
        }
        els.caseDrawerPanel.hidden = !open;
        els.caseDrawerPanel.classList.toggle("open", open);
        els.caseDrawerPanel.setAttribute("aria-hidden", open ? "false" : "true");
        if (els.openCaseDrawer) {
            els.openCaseDrawer.classList.toggle("active", open);
        }
        syncDrawerBackdrop();
    }

    function toggleFilterDrawer() {
        if (!els.filterPanel) {
            return;
        }
        if (isInlineFilterPanel()) {
            els.filterPanel.scrollIntoView({ behavior: "smooth", block: "start" });
            return;
        }
        setFilterDrawer(!els.filterPanel.classList.contains("open"));
    }

    function setFilterDrawer(open) {
        if (!els.filterPanel) {
            return;
        }
        if (isInlineFilterPanel()) {
            els.filterPanel.hidden = state.timelineMode === "timeline";
            syncDrawerBackdrop();
            return;
        }
        if (open) {
            setCaseDrawer(false);
            setCountyDrawer(false);
            setAnalysisDrawer(false);
        }
        els.filterPanel.hidden = !open;
        els.filterPanel.classList.toggle("open", open);
        syncDrawerBackdrop();
    }

    function isInlineFilterPanel() {
        return !!(els.filterPanel && els.filterPanel.classList.contains("inline-filter-panel"));
    }

    function toggleCountyDrawer() {
        if (!els.countyDrawerPanel) {
            return;
        }
        setCountyDrawer(!els.countyDrawerPanel.classList.contains("open"));
    }

    function setCountyDrawer(open) {
        if (!els.countyDrawerPanel) {
            return;
        }
        if (open) {
            setCaseDrawer(false);
            setFilterDrawer(false);
            setAnalysisDrawer(false);
        }
        els.countyDrawerPanel.hidden = !open;
        els.countyDrawerPanel.classList.toggle("open", open);
        els.countyDrawerPanel.setAttribute("aria-hidden", open ? "false" : "true");
        if (els.toggleCountyDrawer) {
            els.toggleCountyDrawer.classList.toggle("active", open);
        }
        syncDrawerBackdrop();
    }

    function toggleAnalysisDrawer() {
        if (!els.analysisDrawerPanel) {
            return;
        }
        setAnalysisDrawer(!els.analysisDrawerPanel.classList.contains("open"));
    }

    function setAnalysisDrawer(open) {
        if (!els.analysisDrawerPanel) {
            return;
        }
        if (open) {
            setCaseDrawer(false);
            setFilterDrawer(false);
            setCountyDrawer(false);
            initAnalysisCharts();
            renderAnalysisCharts(getFilteredRows());
        }
        els.analysisDrawerPanel.hidden = !open;
        els.analysisDrawerPanel.classList.toggle("open", open);
        els.analysisDrawerPanel.setAttribute("aria-hidden", open ? "false" : "true");
        if (els.toggleAnalysisDrawer) {
            els.toggleAnalysisDrawer.classList.toggle("active", open);
        }
        syncDeveloperMode();
        syncDrawerBackdrop();
        if (open) {
            setTimeout(function () {
                Object.values(charts).forEach(function (chart) {
                    if (chart && typeof chart.resize === "function") chart.resize();
                });
            }, 300);
        }
    }

    function getCountyMapCenter(countyName) {
        if (!state.geoJson || !state.geoJson.features || !state.geoJson.features.length) {
            return [121, 23.9];
        }
        const normalizedTarget = normalizeCountyName(countyName);
        const feature = state.geoJson.features.find(function (item) {
            const props = item && item.properties ? item.properties : {};
            const name = normalizeCountyName(props.COUNTYNAME || props.name || "");
            return name === normalizedTarget;
        });
        const cp = feature && feature.properties ? feature.properties.cp : null;
        if (Array.isArray(cp) && cp.length === 2) {
            return cp;
        }
        const bounds = getGeometryBounds(feature && feature.geometry);
        if (bounds) {
            return [(bounds.minX + bounds.maxX) / 2, (bounds.minY + bounds.maxY) / 2];
        }
        return [121, 23.9];
    }

    function getGeometryBounds(geometry) {
        if (!geometry || !Array.isArray(geometry.coordinates)) {
            return null;
        }
        const bounds = { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity };
        function walk(node) {
            if (!Array.isArray(node)) {
                return;
            }
            if (typeof node[0] === "number" && typeof node[1] === "number") {
                bounds.minX = Math.min(bounds.minX, node[0]);
                bounds.maxX = Math.max(bounds.maxX, node[0]);
                bounds.minY = Math.min(bounds.minY, node[1]);
                bounds.maxY = Math.max(bounds.maxY, node[1]);
                return;
            }
            node.forEach(walk);
        }
        walk(geometry.coordinates);
        return Number.isFinite(bounds.minX) ? bounds : null;
    }

    function initResizableTables() {
        const headers = document.querySelectorAll(".table-wrap table thead th");
        headers.forEach(function (header) {
            if (header.classList.contains("resizable-th")) {
                return;
            }
            header.classList.add("resizable-th");
            const handle = document.createElement("span");
            handle.className = "resize-handle";
            header.appendChild(handle);
            handle.addEventListener("mousedown", function (event) {
                const wrap = header.closest("#caseTableWrap, #countySummaryWrap");
                const tableType = wrap && wrap.id === "caseTableWrap" ? "case" : "county";
                if (state.tableWidthLocks[tableType]) {
                    return;
                }
                event.preventDefault();
                const startX = event.pageX;
                const startWidth = header.offsetWidth;
                function onMouseMove(moveEvent) {
                    const width = Math.max(90, startWidth + (moveEvent.pageX - startX));
                    header.style.width = `${width}px`;
                    header.style.minWidth = `${width}px`;
                }
                function onMouseUp() {
                    document.removeEventListener("mousemove", onMouseMove);
                    document.removeEventListener("mouseup", onMouseUp);
                    persistTableColumnWidths();
                }
                document.addEventListener("mousemove", onMouseMove);
                document.addEventListener("mouseup", onMouseUp);
            });
        });
    }

    function getCaseNotes(item) {
        return item.riskReasons.filter(function (reason) {
            return reason === MERGED_CASE_REASON;
        });
    }
    function filterTableRowsBySearch(rows) {
        const keyword = normalizeSearchText(state.tableSearch);
        if (!keyword) {
            return rows;
        }
        return rows.filter(function (item) {
            return normalizeSearchText(item.applicant).includes(keyword)
                || normalizeSearchText(item.caseId).includes(keyword);
        });
    }

    function filterCaseDetailRows(rows) {
        return rows.filter(function (item) {
            return normalizeSearchText(item.applicant) !== "0";
        });
    }

    function highlightSearchText(value) {
        const rawValue = text(value);
        const keyword = text(state.tableSearch);
        if (!keyword) {
            return escapeHtml(rawValue);
        }
        const index = rawValue.toLowerCase().indexOf(keyword.toLowerCase());
        if (index < 0) {
            return escapeHtml(rawValue);
        }
        return [
            escapeHtml(rawValue.slice(0, index)),
            `<mark>${escapeHtml(rawValue.slice(index, index + keyword.length))}</mark>`,
            escapeHtml(rawValue.slice(index + keyword.length))
        ].join("");
    }

    function normalizeSearchText(value) {
        return text(value).toLowerCase();
    }

    function formatCaseNoDisplayForTable(value) {
        if (state.customerMasked) {
            return escapeHtml(formatCaseNoDisplay(value));
        }
        return highlightSearchText(value);
    }
    function aggregateByCounty(rows) {
        const byCounty = new Map();
        rows.forEach(function (item) {
            const county = item.county;
            if (!byCounty.has(county)) {
                byCounty.set(county, {
                    county: county,
                    displayName: displayCountyName(county),
                    principalBalance: 0,
                    approvedUnused: 0,
                    householdAmount: 0,
                    collateralValue: 0,
                    count: 0,
                    highRiskCount: 0
                });
            }

            const current = byCounty.get(county);
            current.principalBalance += item.principalBalance;
            current.approvedUnused += item.approvedUnused;
            current.householdAmount += item.householdAmount;
            current.collateralValue += item.collateralValue;
            current.count += 1;
            if (isHighRiskItem(item)) {
                current.highRiskCount += 1;
            }
        });

        return Array.from(byCounty.values()).filter(function (item) {
            return item.principalBalance > 0 || item.collateralValue > 0;
        }).sort(function (a, b) {
            return b.principalBalance - a.principalBalance;
        });
    }

    function aggregateByCountyForExposure(rows) {
        const byCounty = new Map();
        rows.forEach(function (item) {
            const county = item.county;
            if (!byCounty.has(county)) {
                byCounty.set(county, {
                    county: county,
                    displayName: displayCountyName(county),
                    principalBalance: 0,
                    approvedUnused: 0,
                    householdAmount: 0,
                    count: 0
                });
            }

            const current = byCounty.get(county);
            current.principalBalance += item.principalBalance;
            current.approvedUnused += item.approvedUnused;
            current.householdAmount += item.householdAmount;
            current.count += 1;
        });

        return Array.from(byCounty.values()).filter(function (item) {
            return getExposureSortValue(item) > 0;
        });
    }

    function rowExposure(row) {
        let total = 0;
        state.cityFields.forEach(function (_fields, county) {
            total += countyExposure(row, county);
        });
        return total || toNumber(row[FIELD.principal]);
    }

    function rowCollateral(row) {
        let total = 0;
        state.cityFields.forEach(function (_fields, county) {
            total += countyCollateral(row, county);
        });
        return total || toNumber(row[FIELD.collateral]);
    }

    function countyExposure(row, county) {
        const field = state.cityFields.get(county);
        return field ? toNumber(row[field.remain]) : 0;
    }

    function countyCollateral(row, county) {
        const field = state.cityFields.get(county);
        return field ? toNumber(row[field.worth]) : 0;
    }

    function countyCapital(row, county) {
        const field = state.cityFields.get(county);
        return field ? toNumber(row[field.capital]) : 0;
    }

    function hasCountyData(row, county) {
        return countyExposure(row, county) > 0 || countyCollateral(row, county) > 0;
    }

    function getPrimaryCounty(row) {
        let selected = "";
        let maxValue = 0;
        state.cityFields.forEach(function (_fields, county) {
            const value = countyExposure(row, county);
            if (value > maxValue) {
                maxValue = value;
                selected = county;
            }
        });
        return selected || "全台";
    }

    function getVacancyMetrics(county) {
        if (county) {
            const data = state.vacancyByCounty.get(county);
            if (!data) {
                return { vacantHomes: 0, unsoldHomes: 0, totalHomes: 0, vacancyRate: 0, unsoldRate: 0, quadrant: "-" };
            }
            return {
                vacantHomes: data.vacantHomes,
                unsoldHomes: data.unsoldHomes,
                totalHomes: data.totalHomes,
                vacancyRate: safeDivide(data.vacantHomes, data.totalHomes),
                unsoldRate: safeDivide(data.unsoldHomes, data.totalHomes),
                quadrant: data.quadrant || "-"
            };
        }

        const total = Array.from(state.vacancyByCounty.values()).reduce(function (acc, item) {
            acc.vacantHomes += item.vacantHomes;
            acc.unsoldHomes += item.unsoldHomes;
            acc.totalHomes += item.totalHomes;
            return acc;
        }, { vacantHomes: 0, unsoldHomes: 0, totalHomes: 0 });

        return {
            vacantHomes: total.vacantHomes,
            unsoldHomes: total.unsoldHomes,
            totalHomes: total.totalHomes,
            vacancyRate: safeDivide(total.vacantHomes, total.totalHomes),
            unsoldRate: safeDivide(total.unsoldHomes, total.totalHomes),
            quadrant: "全台加總"
        };
    }

    function getSelectedDistrictVacancyMetrics(county) {
        const countyKey = normalizeCountyName(county);
        const selectedDistricts = getDistrictFilterValue(countyKey);
        if (!Array.isArray(selectedDistricts) || !selectedDistricts.length) {
            return [];
        }
        return selectedDistricts.map(function (district) {
            const data = state.vacancyByDistrict.get(buildDistrictVacancyKey(countyKey, district));
            if (!data) {
                return null;
            }
            return {
                district: district,
                vacancyRate: data.vacancyRate,
                unsoldRate: data.unsoldRate
            };
        }).filter(Boolean);
    }

    function getRiskLevel(ltv, collateralValue) {
        if (!collateralValue || ltv > 1) {
            return "danger";
        }
        if (ltv >= 0.85) {
            return "caution";
        }
        if (ltv >= 0.7) {
            return "normal";
        }
        return "safe";
    }

    function getLtvBucket(ltv) {
        const bucket = LTV_BUCKETS.find(function (entry) {
            return ltv >= entry.min && ltv < entry.max;
        });
        return bucket ? bucket.id : "100+";
    }

    function updateRiskClass(element, riskLevel) {
        element.classList.remove("risk-safe", "risk-normal", "risk-caution", "risk-danger");
        element.classList.add(RISK_META[riskLevel].className);
    }

    function buildFilterSummary() {
        const parts = [];
        const countyValues = getArrayFilterValues("county");
        parts.push(countyValues.length ? countyValues.map(formatCountyFilterLabel).join("、") : "全台");
        getArrayFilterValues("industry").forEach(function (industry) {
            parts.push(getIndustryFilterDisplayName(industry));
        });
        if (state.filters.industryLabel) {
            parts.push(`業別:${state.filters.industryLabel}`);
        }
        getArrayFilterValues("ltvBucket").forEach(function (ltvBucket) {
            const bucket = LTV_BUCKETS.find(function (item) { return item.id === ltvBucket; });
            parts.push(bucket ? bucket.label : ltvBucket);
        });
        if (state.filters.anomalyMode === "exclude") {
            parts.push("僅顯示已起租");
        }
        if (state.filters.anomalyMode === "only") {
            parts.push("僅顯示未起租");
        }
        getSelectedMatrixLabels().forEach(function (label) {
            parts.push(label);
        });
        if (state.filters.group) {
            parts.push(formatSensitiveEntityName(state.filters.group));
        }
        if (state.filters.customer) {
            parts.push(formatCustomerDisplayName(state.filters.customer));
        }
        return parts.join(" / ");
    }

    function formatCountyFilterLabel(county) {
        const districtValue = getDistrictFilterValue(county);
        if (districtValue === "*") {
            return displayCountyName(county);
        }
        return `${displayCountyName(county)}(${formatNumber(districtValue.length)}區)`;
    }

    function getMatrixSelectionSummary() {
        const labels = getSelectedMatrixLabels();
        return labels.length ? labels.join("、") : "未套用矩陣篩選";
    }

    function getSelectedMatrixLabels() {
        return COLLATERAL_LEVEL_FILTERS.filter(function (item) {
            return state.filters.collateralLevels.includes(item.id);
        }).map(function (item) {
            return `擔保覆蓋率 ${item.label}`;
        }).concat(QUADRANT_FILTERS.filter(function (item) {
            return state.filters.quadrants.includes(item.id);
        }).map(function (item) {
            return item.label;
        }));
    }

    function hasMatrixFilters() {
        return state.filters.collateralLevels.length > 0 || state.filters.quadrants.length > 0;
    }

    function setStatus(status, summary) {
        els.dataStatus.textContent = status;
        els.filterSummary.textContent = summary;
    }

    function renderActiveFiltersBar() {
        const chips = buildActiveFilterChips();
        renderFilterChipsTo("activeFilterChips", chips);
        renderFilterChipsTo("caseFilterChips", chips);
        renderFilterChipsTo("countyFilterChips", chips);
    }

    function renderFilterChipsTo(targetKey, chips) {
        if (!els[targetKey]) {
            return;
        }
        if (!chips.length) {
            els[targetKey].innerHTML = "<span class=\"filter-chip\">全部資料</span>";
            return;
        }
        els[targetKey].innerHTML = chips.map(function (chip) {
            const attrs = chip.attrs.map(function (item) {
                return `${item.key}=\"${escapeHtml(item.value)}\"`;
            }).join(" ");
            return `<span class="filter-chip">${escapeHtml(chip.label)} <button type="button" data-remove-filter="1" ${attrs}>×</button></span>`;
        }).join("");
    }

    function buildActiveFilterChips() {
        const chips = [];
        getArrayFilterValues("county").forEach(function (county) {
            const districtValue = getDistrictFilterValue(county);
            const districtLabel = districtValue === "*" ? "" : `（${formatNumber(districtValue.length)}區）`;
            chips.push({ label: `區域：${displayCountyName(county)}${districtLabel}`, attrs: [{ key: "data-filter-key", value: "county" }, { key: "data-filter-id", value: county }] });
        });
        getArrayFilterValues("industry").forEach(function (industry) {
            chips.push({ label: `產品/業別：${getIndustryFilterDisplayName(industry)}`, attrs: [{ key: "data-filter-key", value: "industry" }, { key: "data-filter-id", value: industry }] });
        });
        if (state.filters.industryLabel) {
            chips.push({ label: `業別標籤：${state.filters.industryLabel}`, attrs: [{ key: "data-filter-key", value: "industryLabel" }] });
        }
        if (state.filters.riskLevel) {
            const riskMap = { safe: "安全", normal: "正常", caution: "注意", danger: "危險" };
            chips.push({ label: `風險等級：${riskMap[state.filters.riskLevel] || state.filters.riskLevel}`, attrs: [{ key: "data-filter-key", value: "riskLevel" }] });
        }
        getArrayFilterValues("ltvBucket").forEach(function (ltvBucket) {
            const bucket = LTV_BUCKETS.find(function (item) { return item.id === ltvBucket; });
            chips.push({ label: `LTV：${bucket ? bucket.label : ltvBucket}`, attrs: [{ key: "data-filter-key", value: "ltvBucket" }, { key: "data-filter-id", value: ltvBucket }] });
        });
        if (state.filters.coverageBucket) {
            const bucket = COVERAGE_BUCKETS.find(function (item) { return item.id === state.filters.coverageBucket; });
            chips.push({ label: `覆蓋率：${bucket ? bucket.label : state.filters.coverageBucket}`, attrs: [{ key: "data-filter-key", value: "coverageBucket" }] });
        }
        if (state.filters.anomalyMode === "exclude") {
            chips.push({ label: "起租狀況：已起租", attrs: [{ key: "data-filter-key", value: "anomalyMode" }] });
        }
        if (state.filters.anomalyMode === "only") {
            chips.push({ label: "起租狀況：未起租", attrs: [{ key: "data-filter-key", value: "anomalyMode" }] });
        }
        state.filters.collateralLevels.forEach(function (id) {
            const item = COLLATERAL_LEVEL_FILTERS.find(function (entry) { return entry.id === id; });
            chips.push({
                label: `擔保覆蓋率：${item ? item.label : id}`,
                attrs: [{ key: "data-filter-key", value: "collateralLevels" }, { key: "data-filter-id", value: id }]
            });
        });
        state.filters.quadrants.forEach(function (id) {
            const item = QUADRANT_FILTERS.find(function (entry) { return entry.id === id; });
            chips.push({
                label: `象限：${item ? item.label : id}`,
                attrs: [{ key: "data-filter-key", value: "quadrants" }, { key: "data-filter-id", value: id }]
            });
        });
        if (state.filters.group) {
            chips.push({ label: `集團：${formatSensitiveEntityName(state.filters.group)}`, attrs: [{ key: "data-filter-key", value: "group" }] });
        }
        if (state.filters.customer) {
            chips.push({ label: `客戶：${formatCustomerDisplayName(state.filters.customer)}`, attrs: [{ key: "data-filter-key", value: "customer" }] });
        }
        return chips;
    }

    function removeFilterByChip(button) {
        const key = button.getAttribute("data-filter-key");
        const id = button.getAttribute("data-filter-id");
        if (!key) {
            return;
        }
        if (key === "collateralLevels") {
            state.filters.collateralLevels = state.filters.collateralLevels.filter(function (item) { return item !== id; });
            return;
        }
        if (key === "quadrants") {
            state.filters.quadrants = state.filters.quadrants.filter(function (item) { return item !== id; });
            return;
        }
        if (key === "county" || key === "industry" || key === "ltvBucket") {
            if (id) {
                state.filters[key] = getArrayFilterValues(key).filter(function (item) { return item !== id; });
                if (key === "county") {
                    syncDistrictFilterForCounty(id, false);
                }
            } else {
                state.filters[key] = [];
                if (key === "county") {
                    state.filters.districts = {};
                    state.activeDistrictCounty = "";
                }
            }
            return;
        }
        if (Object.prototype.hasOwnProperty.call(state.filters, key)) {
            state.filters[key] = "";
        }
    }

    function renderMainGridMode() {
        const mainGrid = document.querySelector(".main-grid");
        if (!mainGrid || !els.newsPanel || !els.toggleNewsPanel) {
            return;
        }
        const newsMode = state.mainGridMode === "news";
        mainGrid.classList.toggle("news-mode", newsMode);
        els.newsPanel.hidden = !newsMode;
        els.toggleNewsPanel.classList.toggle("active", newsMode);
        els.toggleNewsPanel.title = newsMode ? "新聞面板：開啟" : "新聞面板：關閉";
        els.toggleNewsPanel.setAttribute("aria-label", els.toggleNewsPanel.title);
        const label = els.toggleNewsPanel.querySelector(".tool-label");
        if (label) {
            label.textContent = "新聞";
        }
        scheduleChartReflow();
    }

    function renderNewsPanel() {
        if (!els.newsList || !els.newsRegionBadge) {
            return;
        }
        const selectedCounties = getArrayFilterValues("county").map(displayCountyName);
        els.newsRegionBadge.textContent = `目前區域：${selectedCounties.length ? selectedCounties.join("、") : "全台"}`;
        const list = selectedCounties.length
            ? state.newsItems.filter(function (item) { return selectedCounties.includes(item.county); })
            : state.newsItems.slice();
        if (!list.length) {
            els.newsList.innerHTML = "<li class=\"news-item\"><span>目前沒有符合區域的新聞</span></li>";
            return;
        }
        els.newsList.innerHTML = list.map(function (item) {
            const dateText = item.date ? `${escapeHtml(item.date)} ` : "";
            return `<li class="news-item"><a class="news-link" href="${escapeHtml(item.url)}" target="_blank" rel="noopener noreferrer">${dateText}[${escapeHtml(item.county)}] ${escapeHtml(item.title)}</a><span class="table-actions"><button class="ghost-button compact-button" type="button" data-edit-news-id="${escapeHtml(item.id)}">編輯</button><button class="ghost-button compact-button" type="button" data-delete-news-id="${escapeHtml(item.id)}">刪除</button></span></li>`;
        }).join("");
    }

    function syncNewsFormVisibility() {
        if (!els.newsForm || !els.toggleNewsFormBtn) {
            return;
        }
        const open = !!state.showNewsForm;
        els.newsForm.hidden = !open;
        if (!open) {
            els.toggleNewsFormBtn.textContent = "新增新聞";
            return;
        }
        els.toggleNewsFormBtn.textContent = state.editingNewsId ? "編輯中" : "收合輸入";
    }

    function resetNewsFormFields() {
        if (!els.newsCountySelect || !els.newsTitleInput || !els.newsDateInput || !els.newsUrlInput) {
            return;
        }
        els.newsCountySelect.value = "";
        els.newsTitleInput.value = "";
        els.newsDateInput.value = new Date().toISOString().slice(0, 10);
        els.newsUrlInput.value = "";
    }

    function initSizeHoverHints() {
        const targets = Array.from(document.querySelectorAll(".panel, .chart, .mini-kpi-grid article"));
        if (!targets.length) {
            return;
        }
        const badge = document.createElement("div");
        badge.className = "size-hover-badge";
        document.body.appendChild(badge);

        function showFor(el) {
            const rect = el.getBoundingClientRect();
            const width = Math.round(rect.width);
            const height = Math.round(rect.height);
            const titleEl = el.querySelector && el.querySelector("h2");
            const name = titleEl ? text(titleEl.textContent) : (el.id || "區塊");
            badge.textContent = `${name}：${width} x ${height}px`;
            badge.classList.add("visible");
        }

        targets.forEach(function (el) {
            el.addEventListener("mouseenter", function () {
                showFor(el);
            });
            el.addEventListener("mousemove", function () {
                showFor(el);
            });
            el.addEventListener("mouseleave", function () {
                badge.classList.remove("visible");
            });
        });
    }

    async function captureFullPagePng() {
        if (!els.capturePagePngBtn) {
            return;
        }
        const captureLibrary = getHtml2Canvas();
        if (!captureLibrary) {
            showNotice("完整頁面 PNG 截圖失敗，請確認 libs/html2canvas.min.js 已放在專案資料夾。");
            return;
        }
        const button = els.capturePagePngBtn;
        const originalText = button.innerHTML;
        button.disabled = true;
        button.innerHTML = '<span class="page-screenshot-glyph" aria-hidden="true">▣</span><span>產生中</span>';
        try {
            await waitForNextFrame();
            const canvas = await buildFullPagePngCanvas(captureLibrary);
            const selectedCounties = getArrayFilterValues("county").map(displayCountyName);
            const regionLabel = selectedCounties.length ? selectedCounties.join("、") : "全台";
            const filename = `擔保曝險地圖_${sanitizeFilename(regionLabel)}_${formatDownloadTimestamp()}.png`;
            downloadCanvasPng(canvas, filename);
        } catch (error) {
            console.error(error);
            showNotice("完整頁面 PNG 截圖失敗，請確認瀏覽器支援 html2canvas 與 Canvas 匯出。");
        } finally {
            button.disabled = false;
            button.innerHTML = originalText;
        }
    }

    async function buildFullPagePngCanvas(captureLibrary) {
        const hiddenNodes = Array.from(document.querySelectorAll("[data-screenshot-exclude]"));
        const previousDisplay = hiddenNodes.map(function (node) {
            return node.style.display;
        });
        hiddenNodes.forEach(function (node) {
            node.style.display = "none";
        });
        document.body.classList.add("screenshot-capture");
        try {
            await waitForNextFrame();
            const width = Math.ceil(Math.max(
                document.documentElement.scrollWidth,
                document.body ? document.body.scrollWidth : 0,
                document.documentElement.clientWidth
            ));
            const height = Math.ceil(Math.max(
                document.documentElement.scrollHeight,
                document.body ? document.body.scrollHeight : 0,
                document.documentElement.clientHeight
            ));
            const bgColor = isDarkTheme() ? "#1A2433" : "#F8FAFC";
            return await captureLibrary(document.body, {
                scale: 2,
                backgroundColor: bgColor,
                useCORS: true,
                allowTaint: false,
                logging: false,
                width: width,
                height: height,
                windowWidth: width,
                windowHeight: height,
                scrollX: 0,
                scrollY: 0,
                onclone: function (clonedDocument) {
                    if (clonedDocument.body) {
                        clonedDocument.body.classList.add("screenshot-capture");
                    }
                }
            });
        } finally {
            document.body.classList.remove("screenshot-capture");
            hiddenNodes.forEach(function (node, index) {
                node.style.display = previousDisplay[index];
            });
        }
    }

    function getHtml2Canvas() {
        if (window.html2canvas) {
            return window.html2canvas;
        }
        if (typeof html2canvas === "function") {
            return html2canvas;
        }
        return null;
    }

    function downloadCanvasPng(canvas, filename) {
        const link = document.createElement("a");
        link.download = filename;
        link.href = canvas.toDataURL("image/png");
        document.body.appendChild(link);
        link.click();
        link.remove();
    }

    function waitForNextFrame() {
        return new Promise(function (resolve) {
            requestAnimationFrame(function () {
                requestAnimationFrame(resolve);
            });
        });
    }

    function sanitizeFilename(value) {
        return text(value).replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_") || "全台";
    }

    function formatDownloadTimestamp() {
        const now = new Date();
        const pad = function (value) {
            return String(value).padStart(2, "0");
        };
        return [
            now.getFullYear(),
            pad(now.getMonth() + 1),
            pad(now.getDate()),
            "_",
            pad(now.getHours()),
            pad(now.getMinutes()),
            pad(now.getSeconds())
        ].join("");
    }

    function showNotice(message) {
        els.dependencyNotice.hidden = false;
        els.dependencyNotice.textContent = message;
    }

    function normalizeCountyName(name) {
        const value = text(name).replace(/\s/g, "");
        if (!value) {
            return "";
        }
        if (COUNTY_ALIAS[value]) {
            return COUNTY_ALIAS[value];
        }
        if (value.indexOf("臺") === 0) {
            return "台" + value.slice(1);
        }
        return value;
    }

    function displayCountyName(name) {
        const key = normalizeCountyName(name);
        return state.displayCountyByKey.get(key) || key || "全台";
    }

    function loadNewsItems() {
        try {
            const raw = localStorage.getItem(NEWS_STORAGE_KEY);
            if (!raw) {
                return [];
            }
            const parsed = JSON.parse(raw);
            if (!Array.isArray(parsed)) {
                return [];
            }
            return parsed.map(function (item) {
                return {
                    id: text(item.id || `news_${Date.now()}`),
                    county: text(item.county),
                    title: text(item.title),
                    date: text(item.date),
                    url: text(item.url)
                };
            }).filter(function (item) {
                return item.id && item.county && item.title && item.url;
            });
        } catch (_error) {
            return [];
        }
    }

    function saveNewsItems() {
        try {
            localStorage.setItem(NEWS_STORAGE_KEY, JSON.stringify(state.newsItems));
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadUiMode() {
        try {
            const raw = localStorage.getItem(UI_MODE_STORAGE_KEY);
            return raw === "ios" ? "ios" : "default";
        } catch (_error) {
            return "default";
        }
    }

    function loadThemeMode() {
        try {
            const raw = localStorage.getItem(THEME_MODE_STORAGE_KEY);
            return raw === "dark" ? "dark" : "light";
        } catch (_error) {
            return "light";
        }
    }

    function loadBarPalette() {
        try {
            const raw = localStorage.getItem(BAR_PALETTE_STORAGE_KEY);
            return BAR_PALETTE[raw] ? raw : "ocean";
        } catch (_error) {
            return "ocean";
        }
    }

    function saveBarPalette() {
        try {
            localStorage.setItem(BAR_PALETTE_STORAGE_KEY, state.selectedBarPalette);
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadExposureBasis() {
        try {
            const raw = localStorage.getItem(EXPOSURE_BASIS_STORAGE_KEY);
            return raw === "household" ? "household" : "principal";
        } catch (_error) {
            return "principal";
        }
    }

    function loadAmountUnit() {
        try {
            return localStorage.getItem(AMOUNT_UNIT_STORAGE_KEY) === "yi" ? "yi" : "k";
        } catch (error) {
            return "k";
        }
    }

    function saveAmountUnit() {
        try {
            localStorage.setItem(AMOUNT_UNIT_STORAGE_KEY, state.amountUnit === "yi" ? "yi" : "k");
        } catch (error) {}
    }

    function saveExposureBasis() {
        try {
            localStorage.setItem(EXPOSURE_BASIS_STORAGE_KEY, state.exposureBasis === "household" ? "household" : "principal");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadTimelineMode() {
        try {
            return localStorage.getItem(TIMELINE_MODE_STORAGE_KEY) === "timeline" ? "timeline" : "snapshot";
        } catch (_error) {
            return "snapshot";
        }
    }

    function saveTimelineMode() {
        try {
            localStorage.setItem(TIMELINE_MODE_STORAGE_KEY, state.timelineMode === "timeline" ? "timeline" : "snapshot");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadSelectedPeriodId() {
        try {
            return text(localStorage.getItem(SELECTED_PERIOD_STORAGE_KEY));
        } catch (_error) {
            return "";
        }
    }

    function saveSelectedPeriodId() {
        try {
            localStorage.setItem(SELECTED_PERIOD_STORAGE_KEY, state.selectedPeriodId || "");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadComparePeriodId() {
        try {
            return text(localStorage.getItem(COMPARE_PERIOD_STORAGE_KEY));
        } catch (_error) {
            return "";
        }
    }

    function saveComparePeriodId() {
        try {
            localStorage.setItem(COMPARE_PERIOD_STORAGE_KEY, state.comparePeriodId || "");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadTimelineMetric() {
        return FIXED_TIMELINE_METRIC;
    }

    function saveTimelineMetric() {
        try {
            localStorage.setItem(TIMELINE_METRIC_STORAGE_KEY, FIXED_TIMELINE_METRIC);
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadTimelineChangeMode() {
        try {
            return localStorage.getItem(TIMELINE_CHANGE_MODE_STORAGE_KEY) === "decrease" ? "decrease" : "increase";
        } catch (_error) {
            return "increase";
        }
    }

    function saveTimelineChangeMode() {
        try {
            localStorage.setItem(TIMELINE_CHANGE_MODE_STORAGE_KEY, state.timelineChangeMode === "decrease" ? "decrease" : "increase");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function clampFontSizeValue(key, value) {
        const limits = FONT_SIZE_LIMITS[key] || FONT_SIZE_LIMITS.body;
        const parsed = Number.parseInt(value, 10);
        if (!Number.isFinite(parsed)) {
            return limits.fallback;
        }
        return Math.min(limits.max, Math.max(limits.min, parsed));
    }

    function normalizeFontFamily(value) {
        const key = text(value);
        return FONT_FAMILY_OPTIONS[key] ? key : DEFAULT_FONT_SIZE_CONFIG.family;
    }

    function getAppFontFamily() {
        return FONT_FAMILY_OPTIONS[normalizeFontFamily(state.fontSizeConfig && state.fontSizeConfig.family)] || FONT_FAMILY_OPTIONS.system;
    }

    function getCanvasFontFamily() {
        return getAppFontFamily().split(",")[0].replace(/^"|"$/g, "");
    }

    function getCanvasFont(weight, size) {
        return weight + " " + size + "px " + getCanvasFontFamily();
    }

    function normalizeFontSizeConfig(input) {
        return {
            title: clampFontSizeValue("title", input && input.title),
            body: clampFontSizeValue("body", input && input.body),
            chart: clampFontSizeValue("chart", input && input.chart),
            map: clampFontSizeValue("map", input && input.map),
            family: normalizeFontFamily(input && input.family)
        };
    }

    function loadLegacyFontSizeConfig() {
        try {
            const raw = localStorage.getItem("fontSizePreference");
            const legacyMap = {
                small: { title: 14, body: 12, chart: 10, map: 11 },
                medium: DEFAULT_FONT_SIZE_CONFIG,
                large: { title: 18, body: 15, chart: 13, map: 14 },
                xlarge: { title: 21, body: 17, chart: 14, map: 15 }
            };
            return normalizeFontSizeConfig(legacyMap[raw] || DEFAULT_FONT_SIZE_CONFIG);
        } catch (_error) {
            return normalizeFontSizeConfig(DEFAULT_FONT_SIZE_CONFIG);
        }
    }

    function loadFontSizeConfig() {
        try {
            const raw = localStorage.getItem(FONT_SIZE_CONFIG_STORAGE_KEY);
            if (!raw) {
                return loadLegacyFontSizeConfig();
            }
            return normalizeFontSizeConfig(JSON.parse(raw));
        } catch (_error) {
            return loadLegacyFontSizeConfig();
        }
    }

    function getDefaultMainPanelChartSlots() {
        return ["countyRank", "group", "quadrantSummary", "ltv2"];
    }

    function normalizeMainPanelChartSlots(input) {
        const validSet = new Set(getMainChartCatalog().map(function (item) { return item.id; }));
        if (!Array.isArray(input)) {
            return getDefaultMainPanelChartSlots();
        }
        const normalized = input.filter(function (id) {
            return validSet.has(id);
        }).filter(function (id, index, arr) {
            return arr.indexOf(id) === index;
        });
        return normalized.slice(0, 4);
    }

    function loadMainPanelChartSlots() {
        try {
            const raw = localStorage.getItem(MAIN_PANEL_SLOTS_STORAGE_KEY);
            if (!raw) {
                return getDefaultMainPanelChartSlots();
            }
            return normalizeMainPanelChartSlots(JSON.parse(raw));
        } catch (_error) {
            return getDefaultMainPanelChartSlots();
        }
    }

    function saveMainPanelChartSlots() {
        try {
            localStorage.setItem(MAIN_PANEL_SLOTS_STORAGE_KEY, JSON.stringify(normalizeMainPanelChartSlots(state.mainPanelChartSlots)));
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadRankTopN() {
        try {
            const raw = JSON.parse(localStorage.getItem(RANK_TOPN_STORAGE_KEY) || "{}");
            return {
                group: raw.group === 10 ? 10 : 5,
                countyRank: raw.countyRank === 10 ? 10 : 5,
                topCustomer: raw.topCustomer === 5 ? 5 : 10
            };
        } catch (_error) {
            return { group: 5, countyRank: 5, topCustomer: 10 };
        }
    }

    function saveRankTopN() {
        try {
            localStorage.setItem(RANK_TOPN_STORAGE_KEY, JSON.stringify(state.rankTopN));
        } catch (_error) {
            // ignore storage failures
        }
    }

    function saveFontSizeConfig() {
        try {
            localStorage.setItem(FONT_SIZE_CONFIG_STORAGE_KEY, JSON.stringify(normalizeFontSizeConfig(state.fontSizeConfig)));
        } catch (_error) {
            // ignore
        }
    }

    function resetFontSizeConfig() {
        state.fontSizeConfig = normalizeFontSizeConfig(DEFAULT_FONT_SIZE_CONFIG);
        saveFontSizeConfig();
        applyFontSize();
        render();
        animateChartsResize();
    }

    function getFSD(size) {
        const chartBase = clampFontSizeValue("chart", state.fontSizeConfig && state.fontSizeConfig.chart);
        return clampFontSizeValue("chart", chartBase + (Number(size) - DEFAULT_FONT_SIZE_CONFIG.chart));
    }

    function getMapFontSize(size) {
        const mapBase = clampFontSizeValue("map", state.fontSizeConfig && state.fontSizeConfig.map);
        return clampFontSizeValue("map", mapBase + (Number(size) - DEFAULT_FONT_SIZE_CONFIG.map));
    }

    function getActiveBarGradient(type) {
        const key = BAR_PALETTE[state.selectedBarPalette] ? state.selectedBarPalette : "ocean";
        const palette = BAR_PALETTE[key];
        return palette[type] || BAR_PALETTE.ocean[type];
    }

    function hexToRgba(hex, alpha) {
        const normalized = text(hex).replace("#", "");
        if (!/^[0-9a-f]{6}$/i.test(normalized)) {
            return isDarkTheme() ? `rgba(96, 165, 250, ${alpha})` : `rgba(37, 99, 235, ${alpha})`;
        }
        const red = parseInt(normalized.slice(0, 2), 16);
        const green = parseInt(normalized.slice(2, 4), 16);
        const blue = parseInt(normalized.slice(4, 6), 16);
        return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
    }

    function getApprovedUnusedBarGradient(color) {
        const alphaStart = isDarkTheme() ? 0.52 : 0.38;
        const alphaEnd = isDarkTheme() ? 0.30 : 0.20;
        const colors = Array.isArray(color) ? color : ["#60A5FA", "#2563EB"];
        return new echarts.graphic.LinearGradient(1, 0, 0, 0, [
            { offset: 0, color: hexToRgba(colors[0], alphaStart) },
            { offset: 1, color: hexToRgba(colors[1], alphaEnd) }
        ]);
    }

    function isDarkTheme() {
        return state.themeMode === "dark";
    }

    function getChartTextColor() {
        return isDarkTheme() ? "#E2E8F0" : "#475569";
    }

    function getChartSplitLineColor() {
        return isDarkTheme() ? "#4A5D78" : "#E2E8F0";
    }

    function getChartMutedColor() {
        return isDarkTheme() ? "#C1CDDA" : "#5F6F82";
    }

    function getChartSubtleColor() {
        return isDarkTheme() ? "#CBD5E1" : "#6E6E73";
    }

    function getChartEmptyBarColor() {
        return isDarkTheme() ? "#354560" : "#E9ECEF";
    }

    function getChartSeriesBorderColor() {
        return isDarkTheme() ? "#1E2538" : "#FFFFFF";
    }

    function getChartGuideLineColor() {
        return isDarkTheme() ? "#475569" : "#B9C7D5";
    }

    function getRiskThemeColor(riskKey) {
        if (isDarkTheme()) {
            const darkColors = {
                safe: "#10B981",
                normal: "#0EA5E9",
                caution: "#F59E0B",
                danger: "#EF4444"
            };
            return darkColors[riskKey] || darkColors.safe;
        }
        return RISK_META[riskKey] ? RISK_META[riskKey].color : RISK_META.safe.color;
    }

    function getLtvBucketColor(bucketId) {
        const lightColors = {
            "0-60": ["#60A5FA", "#2563EB"],
            "60-80": ["#34D399", "#10B981"],
            "80-100": ["#FBBF24", "#F59E0B"],
            "100+": ["#F87171", "#DC2626"]
        };
        const darkColors = {
            "0-60": ["#60A5FA", "#3B82F6"],
            "60-80": ["#34D399", "#10B981"],
            "80-100": ["#FBBF24", "#F59E0B"],
            "100+": ["#F87171", "#EF4444"]
        };
        const colors = (isDarkTheme() ? darkColors : lightColors)[bucketId] || lightColors["0-60"];
        return new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: colors[0] },
            { offset: 1, color: colors[1] }
        ]);
    }

    function getLtvBucketBorderColor(bucketId) {
        const colors = {
            "0-60": "#2563EB",
            "60-80": "#10B981",
            "80-100": "#F59E0B",
            "100+": "#DC2626"
        };
        if (isDarkTheme() && bucketId === "100+") {
            return "#EF4444";
        }
        return colors[bucketId] || colors["0-60"];
    }

    function saveUiMode() {
        try {
            localStorage.setItem(UI_MODE_STORAGE_KEY, state.uiMode === "ios" ? "ios" : "default");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function saveThemeMode() {
        try {
            localStorage.setItem(THEME_MODE_STORAGE_KEY, state.themeMode === "dark" ? "dark" : "light");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadCustomerMask() {
        try {
            return localStorage.getItem(CUSTOMER_MASK_STORAGE_KEY) === "masked";
        } catch (_error) {
            return false;
        }
    }

    function saveCustomerMask() {
        try {
            localStorage.setItem(CUSTOMER_MASK_STORAGE_KEY, state.customerMasked ? "masked" : "visible");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadDeveloperMode() {
        return false;
    }

    function saveDeveloperMode() {
        try {
            localStorage.setItem(DEVELOPER_MODE_STORAGE_KEY, state.developerMode ? "enabled" : "disabled");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadDeveloperHiddenBlocks() {
        try {
            const stored = localStorage.getItem(DEV_HIDDEN_BLOCKS_STORAGE_KEY);
            if (stored === null) {
                return new Set(DEFAULT_DEV_HIDDEN_BLOCKS);
            }
            const raw = JSON.parse(stored || "[]");
            return new Set(Array.isArray(raw) ? raw.map(text).filter(Boolean) : []);
        } catch (_error) {
            return new Set(DEFAULT_DEV_HIDDEN_BLOCKS);
        }
    }

    function saveDeveloperHiddenBlocks() {
        try {
            localStorage.setItem(DEV_HIDDEN_BLOCKS_STORAGE_KEY, JSON.stringify(Array.from(state.devHiddenBlocks)));
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadComponentDisplayNames() {
        try {
            const raw = JSON.parse(localStorage.getItem(DISPLAY_NAME_STORAGE_KEY) || "{}");
            if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
                return {};
            }
            let migrated = false;
            const normalized = Object.keys(raw).reduce(function (acc, key) {
                const cleanKey = text(key);
                const cleanValue = text(raw[key]);
                if (cleanKey && cleanValue) {
                    if (isLegacyMainSlotDisplayName(cleanKey, cleanValue)) {
                        migrated = true;
                        return acc;
                    }
                    acc[cleanKey] = cleanValue;
                }
                return acc;
            }, {});
            if (migrated) {
                localStorage.setItem(DISPLAY_NAME_STORAGE_KEY, JSON.stringify(normalized));
            }
            return normalized;
        } catch (_error) {
            return {};
        }
    }

    function isLegacyMainSlotDisplayName(key, value) {
        const cleanValue = text(value);
        return (key === "mainSlot1" && ["前十大曝險集團", "前十大授信戶"].indexOf(cleanValue) !== -1)
            || (key === "mainSlot2" && ["區域曝險排行", "區域授信排行"].indexOf(cleanValue) !== -1);
    }

    function saveComponentDisplayNames() {
        try {
            localStorage.setItem(DISPLAY_NAME_STORAGE_KEY, JSON.stringify(state.componentDisplayNames));
        } catch (_error) {
            // ignore storage failures
        }
    }

    function loadDisconnectTimelineFilters() {
        try {
            const stored = localStorage.getItem(DISCONNECT_TIMELINE_FILTERS_STORAGE_KEY);
            return stored === null ? true : stored === "enabled";
        } catch (_error) {
            return true;
        }
    }

    function saveDisconnectTimelineFilters() {
        try {
            localStorage.setItem(DISCONNECT_TIMELINE_FILTERS_STORAGE_KEY, state.disconnectTimelineFilters ? "enabled" : "disabled");
        } catch (_error) {
            // ignore storage failures
        }
    }

    function getDeveloperHideBlocks() {
        const blocks = Array.from(document.querySelectorAll("[data-dev-hide-key]")).map(function (el) {
            const key = text(el.getAttribute("data-dev-hide-key"));
            return {
                key: key,
                label: text(el.getAttribute("data-dev-hide-label")) || key,
                element: el
            };
        }).filter(function (item, index, arr) {
            return item.key && arr.findIndex(function (entry) { return entry.key === item.key; }) === index;
        });
        const order = new Map();
        getDeveloperHideGroups().forEach(function (group, groupIndex) {
            group.items.forEach(function (item, itemIndex) {
                order.set(item.key, groupIndex * 100 + itemIndex);
            });
        });
        return blocks.sort(function (a, b) {
            const aOrder = order.has(a.key) ? order.get(a.key) : 9999;
            const bOrder = order.has(b.key) ? order.get(b.key) : 9999;
            if (aOrder !== bOrder) {
                return aOrder - bOrder;
            }
            return a.label.localeCompare(b.label, "zh-Hant");
        });
    }

    function getDeveloperHideGroups() {
        const blockMap = new Map(getDeveloperHideBlocksRaw().map(function (item) {
            return [item.key, item];
        }));
        const used = new Set();
        const groups = DEVELOPER_HIDE_GROUPS.map(function (group) {
            const items = group.items.map(function (item) {
                const block = blockMap.get(item.key);
                if (!block) {
                    return null;
                }
                used.add(item.key);
                return {
                    key: item.key,
                    label: item.label || block.label,
                    element: block.element
                };
            }).filter(Boolean);
            return { title: group.title, items: items };
        }).filter(function (group) {
            return group.items.length > 0;
        });
        const otherItems = Array.from(blockMap.values()).filter(function (item) {
            return !used.has(item.key);
        });
        if (otherItems.length) {
            groups.push({ title: "其他板塊", items: otherItems });
        }
        return groups;
    }

    function getDeveloperHideBlocksRaw() {
        return Array.from(document.querySelectorAll("[data-dev-hide-key]")).map(function (el) {
            const key = text(el.getAttribute("data-dev-hide-key"));
            return {
                key: key,
                label: text(el.getAttribute("data-dev-hide-label")) || key,
                element: el
            };
        }).filter(function (item, index, arr) {
            return item.key && arr.findIndex(function (entry) { return entry.key === item.key; }) === index;
        });
    }

    function syncDeveloperMode() {
        const isDeveloper = !!state.developerMode;
        document.body.classList.toggle("developer-mode", isDeveloper);
        if (els.developerModeToggle) {
            els.developerModeToggle.checked = isDeveloper;
        }
        if (els.developerPanel) {
            els.developerPanel.hidden = !isDeveloper;
        }
        if (els.timelineFilterLinkToggle) {
            els.timelineFilterLinkToggle.checked = !!state.disconnectTimelineFilters;
        }
        renderDeveloperHiddenList();
        renderDeveloperDisplayNameList();
        applyDeveloperHiddenBlocks();
    }

    function renderDeveloperHiddenList() {
        if (!els.developerHiddenList || !state.developerMode) {
            return;
        }
        const groups = getDeveloperHideGroups();
        els.developerHiddenList.innerHTML = groups.map(function (group) {
            const items = group.items.map(function (item) {
                const checked = state.devHiddenBlocks.has(item.key) ? " checked" : "";
                return `<label class="developer-hidden-item"><input type="checkbox" data-dev-hide-option="${escapeHtml(item.key)}"${checked}><span>${escapeHtml(item.label)}</span></label>`;
            }).join("");
            return `<section class="developer-hidden-group"><h3>${escapeHtml(group.title)}</h3><div class="developer-hidden-group-items">${items}</div></section>`;
        }).join("");
    }

    function renderDeveloperDisplayNameList() {
        if (!els.developerDisplayNameList || !state.developerMode) {
            return;
        }
        const groups = getDeveloperHideGroups();
        els.developerDisplayNameList.innerHTML = groups.map(function (group) {
            const rows = group.items.map(function (item) {
                const value = state.componentDisplayNames[item.key] || "";
                return [
                    `<div class="developer-display-row" data-display-name-row="${escapeHtml(item.key)}">`,
                    `<span class="developer-display-category" title="${escapeHtml(group.title)}">${escapeHtml(group.title)}</span>`,
                    `<span class="developer-display-original" title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</span>`,
                    `<input class="developer-display-input" type="text" data-display-name-key="${escapeHtml(item.key)}" value="${escapeHtml(value)}" placeholder="${escapeHtml(item.label)}" aria-label="${escapeHtml(item.label)}的新顯示名稱">`,
                    `<button class="ghost-button compact-button developer-display-reset" type="button" data-display-name-reset="${escapeHtml(item.key)}">重設</button>`,
                    "</div>"
                ].join("");
            }).join("");
            return [
                `<section class="developer-display-group">`,
                `<h3>${escapeHtml(group.title)}</h3>`,
                `<div class="developer-display-table">`,
                `<div class="developer-display-row header"><span>面板分類</span><span>原始名稱</span><span>新顯示名稱</span><span>操作</span></div>`,
                rows,
                `</div>`,
                `</section>`
            ].join("");
        }).join("");
    }

    function updateComponentDisplayName(key, value) {
        const cleanKey = text(key);
        if (!cleanKey) {
            return;
        }
        const cleanValue = text(value);
        if (cleanValue) {
            state.componentDisplayNames[cleanKey] = cleanValue;
        } else {
            delete state.componentDisplayNames[cleanKey];
        }
        saveComponentDisplayNames();
        applyComponentDisplayNames();
    }

    function resetComponentDisplayName(key) {
        const cleanKey = text(key);
        if (!cleanKey) {
            return;
        }
        delete state.componentDisplayNames[cleanKey];
        saveComponentDisplayNames();
        renderDeveloperDisplayNameList();
        applyComponentDisplayNames();
    }

    function resetAllComponentDisplayNames() {
        state.componentDisplayNames = {};
        saveComponentDisplayNames();
        renderDeveloperDisplayNameList();
        applyComponentDisplayNames();
        showNotice("元件顯示名稱已全部重設。");
    }

    function getDisplayNameForKey(key, fallback) {
        const cleanKey = text(key);
        return text(state.componentDisplayNames[cleanKey]) || text(fallback) || cleanKey;
    }

    function applyComponentDisplayNames() {
        getDeveloperHideGroups().forEach(function (group) {
            group.items.forEach(function (item) {
                const displayName = getDisplayNameForKey(item.key, item.label);
                document.querySelectorAll(`[data-dev-hide-key="${cssEscape(item.key)}"]`).forEach(function (target) {
                    applyComponentDisplayNameToElement(target, displayName);
                });
                applyMainSlotDisplayName(item.key, displayName);
            });
        });
    }

    function applyComponentDisplayNameToElement(target, displayName) {
        if (!target || !displayName) {
            return;
        }
        if (target.matches && target.matches("button")) {
            target.title = displayName;
            target.setAttribute("aria-label", displayName);
            const label = target.querySelector(".tool-label");
            if (label) {
                label.textContent = displayName;
            }
            return;
        }
        if (target.matches && target.matches("section, article")) {
            target.setAttribute("aria-label", displayName);
        }
        const heading = target.querySelector("h2, h3");
        if (heading) {
            if (heading.id === "kpiPanelTitle") {
                const count = toNumber(heading.getAttribute("data-kpi-count"));
                heading.textContent = `${displayName}(統計件數${formatNumber(count)}件)`;
                return;
            }
            heading.textContent = displayName;
        }
    }

    function applyMainSlotDisplayName(key, displayName) {
        const slotMap = {
            mainSlot1: els.mainSlotTitle1,
            mainSlot2: els.mainSlotTitle2,
            mainSlot3: els.mainSlotTitle3,
            mainSlot4: els.mainSlotTitle4
        };
        if (slotMap[key] && displayName) {
            slotMap[key].textContent = displayName;
        }
    }

    function cssEscape(value) {
        if (window.CSS && typeof window.CSS.escape === "function") {
            return window.CSS.escape(value);
        }
        return String(value).replace(/["\\]/g, "\\$&");
    }

    function initAltDisplacementMeasure() {
        if (!els.displacementMeasureOverlay) {
            return;
        }
        document.addEventListener("mousedown", function (event) {
            if (!event.altKey || event.button !== 0) {
                return;
            }
            startDisplacementMeasure(event);
        });
        document.addEventListener("mousemove", function (event) {
            if (!state.displacementMeasure.dragging) {
                return;
            }
            if (!event.altKey) {
                cancelDisplacementMeasure();
                return;
            }
            updateDisplacementMeasure(event.clientX, event.clientY, event.pageX, event.pageY);
        });
        document.addEventListener("mouseup", function (event) {
            if (!state.displacementMeasure.dragging || event.button !== 0) {
                return;
            }
            if (!event.altKey || state.displacementMeasure.cancelled) {
                cancelDisplacementMeasure();
                return;
            }
            finishDisplacementMeasure();
        });
        document.addEventListener("keyup", function (event) {
            if (event.key === "Alt" && state.displacementMeasure.dragging) {
                cancelDisplacementMeasure();
            }
        });
        document.addEventListener("keydown", function (event) {
            if (event.key === "Escape" && state.districtPopupCounty) {
                closeDistrictPopup();
            }
            if (event.key === "Escape" && state.displacementMeasure.dragging) {
                cancelDisplacementMeasure();
            }
        });
    }

    function startDisplacementMeasure(event) {
        event.preventDefault();
        state.displacementMeasure.active = true;
        state.displacementMeasure.dragging = true;
        state.displacementMeasure.cancelled = false;
        state.displacementMeasure.startX = event.clientX;
        state.displacementMeasure.startY = event.clientY;
        state.displacementMeasure.currentX = event.clientX;
        state.displacementMeasure.currentY = event.clientY;
        state.displacementMeasure.startPageX = Math.round(event.pageX);
        state.displacementMeasure.startPageY = Math.round(event.pageY);
        state.displacementMeasure.currentPageX = Math.round(event.pageX);
        state.displacementMeasure.currentPageY = Math.round(event.pageY);
        state.displacementMeasure.targetBox = getDisplacementTargetBox(event.clientX, event.clientY);
        state.displacementMeasure.endTargetBox = state.displacementMeasure.targetBox;
        if (els.displacementMeasureOverlay) {
            els.displacementMeasureOverlay.hidden = false;
        }
        updateDisplacementMeasure(event.clientX, event.clientY, event.pageX, event.pageY);
    }

    function updateDisplacementMeasure(currentX, currentY, pageX, pageY) {
        const measure = state.displacementMeasure;
        measure.currentX = currentX;
        measure.currentY = currentY;
        measure.currentPageX = Math.round(typeof pageX === "number" ? pageX : currentX + window.scrollX);
        measure.currentPageY = Math.round(typeof pageY === "number" ? pageY : currentY + window.scrollY);
        measure.endTargetBox = getDisplacementTargetBox(currentX, currentY);
        const dx = Math.round(currentX - measure.startX);
        const dy = Math.round(currentY - measure.startY);
        const direction = getDisplacementDirection(dx, dy);
        const transformText = `transform: translate(${dx}px, ${dy}px);`;
        if (els.displacementMeasureLine) {
            els.displacementMeasureLine.setAttribute("x1", String(measure.startX));
            els.displacementMeasureLine.setAttribute("y1", String(measure.startY));
            els.displacementMeasureLine.setAttribute("x2", String(currentX));
            els.displacementMeasureLine.setAttribute("y2", String(currentY));
        }
        positionDisplacementPoint(els.displacementMeasureStart, measure.startX, measure.startY);
        positionDisplacementPoint(els.displacementMeasureEnd, currentX, currentY);
        positionDisplacementHud(currentX, currentY);
        if (els.displacementMeasureDx) {
            els.displacementMeasureDx.textContent = `dx: ${dx}px`;
        }
        if (els.displacementMeasureDy) {
            els.displacementMeasureDy.textContent = `dy: ${dy}px`;
        }
        if (els.displacementMeasureDirection) {
            els.displacementMeasureDirection.textContent = `方向：${direction}`;
        }
        if (els.displacementMeasureStartCoord) {
            els.displacementMeasureStartCoord.textContent = `起點絕對座標：x ${measure.startPageX}px, y ${measure.startPageY}px`;
        }
        if (els.displacementMeasureCurrentCoord) {
            els.displacementMeasureCurrentCoord.textContent = `目前絕對座標：x ${measure.currentPageX}px, y ${measure.currentPageY}px`;
        }
        if (els.displacementMeasureTargetBox) {
            els.displacementMeasureTargetBox.textContent = formatDisplacementTargetBox(measure.targetBox, "起點");
        }
        if (els.displacementMeasureEndTargetBox) {
            els.displacementMeasureEndTargetBox.textContent = formatDisplacementTargetBox(measure.endTargetBox, "終點");
        }
        if (els.displacementMeasureTransform) {
            els.displacementMeasureTransform.textContent = transformText;
        }
    }

    function getDisplacementTargetBox(clientX, clientY) {
        const target = document.elementFromPoint(clientX, clientY);
        if (!target || target === document.documentElement || target === document.body || target.closest("[data-screenshot-exclude='1']")) {
            return null;
        }
        const rect = target.getBoundingClientRect();
        return {
            label: getElementDisplayLabel(target),
            panel: getElementPanelContext(target),
            left: Math.round(rect.left + window.scrollX),
            top: Math.round(rect.top + window.scrollY),
            width: Math.round(rect.width),
            height: Math.round(rect.height)
        };
    }

    function getElementDisplayLabel(target) {
        const devLabel = text(target.getAttribute("data-dev-hide-label"));
        if (devLabel) {
            return devLabel;
        }
        const block = target.closest("[data-dev-hide-label]");
        if (block && text(block.getAttribute("data-dev-hide-label"))) {
            return text(block.getAttribute("data-dev-hide-label"));
        }
        const chart = target.closest(".chart[id]");
        if (chart) {
            return `#${chart.id}`;
        }
        const control = target.closest("button,input,select,textarea,[role='button']");
        if (control) {
            const controlLabel = text(control.getAttribute("aria-label")) || text(control.getAttribute("title")) || text(control.textContent);
            if (controlLabel) {
                return controlLabel;
            }
        }
        const id = text(target.id);
        if (id) {
            return `#${id}`;
        }
        const className = typeof target.className === "string" ? text(target.className).split(/\s+/u).filter(Boolean).slice(0, 3).join(".") : "";
        return target.tagName ? `${target.tagName.toLowerCase()}${className ? "." + className : ""}` : "element";
    }

    function getElementPanelContext(target) {
        if (target.closest("#timelinePanel")) {
            return "趨勢面板";
        }
        if (target.closest("#analysisDrawerPanel")) {
            return "分析圖表";
        }
        if (target.closest("#caseDrawerPanel")) {
            return "案件明細表";
        }
        if (target.closest("#countyDrawerPanel")) {
            return "縣市彙總表";
        }
        if (target.closest("#developerPanel")) {
            return "開發者設定";
        }
        if (target.closest(".main-grid,.secondary-kpis,#filterPanel,#activeFiltersBar")) {
            return "當期面板";
        }
        return state.timelineMode === "timeline" ? "趨勢面板" : "當期面板";
    }

    function formatDisplacementTargetBox(box, labelPrefix) {
        const prefix = labelPrefix || "起點";
        if (!box) {
            return `${prefix}物件：未偵測`;
        }
        return `${prefix}物件：${box.label}；所在面板：${box.panel}；絕對位置：left ${box.left}px, top ${box.top}px, w ${box.width}px, h ${box.height}px`;
    }

    function positionDisplacementPoint(point, x, y) {
        if (!point) {
            return;
        }
        point.style.left = `${x}px`;
        point.style.top = `${y}px`;
    }

    function positionDisplacementHud(x, y) {
        if (!els.displacementMeasureHud) {
            return;
        }
        const offset = 14;
        const rect = els.displacementMeasureHud.getBoundingClientRect();
        const nextLeft = Math.min(x + offset, window.innerWidth - rect.width - offset);
        const nextTop = Math.min(y + offset, window.innerHeight - rect.height - offset);
        els.displacementMeasureHud.style.left = `${Math.max(offset, nextLeft)}px`;
        els.displacementMeasureHud.style.top = `${Math.max(offset, nextTop)}px`;
    }

    function getDisplacementDirection(dx, dy) {
        const parts = [];
        if (dx > 0) {
            parts.push(`往右 ${dx}px`);
        } else if (dx < 0) {
            parts.push(`往左 ${Math.abs(dx)}px`);
        }
        if (dy > 0) {
            parts.push(`往下 ${dy}px`);
        } else if (dy < 0) {
            parts.push(`往上 ${Math.abs(dy)}px`);
        }
        return parts.length ? parts.join("，") : "無位移";
    }

    function getDisplacementClipboardText() {
        const measure = state.displacementMeasure;
        const dx = Math.round(measure.currentX - measure.startX);
        const dy = Math.round(measure.currentY - measure.startY);
        const targetBox = measure.targetBox ? `; ${formatDisplacementTargetBox(measure.targetBox, "起點")}` : "";
        const endTargetBox = measure.endTargetBox ? `; ${formatDisplacementTargetBox(measure.endTargetBox, "終點")}` : "";
        return `dx: ${dx}px; dy: ${dy}px; 方向：${getDisplacementDirection(dx, dy)}; 起點絕對座標：x ${measure.startPageX}px, y ${measure.startPageY}px; 目前絕對座標：x ${measure.currentPageX}px, y ${measure.currentPageY}px${targetBox}${endTargetBox}; transform: translate(${dx}px, ${dy}px);`;
    }

    function finishDisplacementMeasure() {
        const textToCopy = getDisplacementClipboardText();
        state.displacementMeasure.dragging = false;
        state.displacementMeasure.active = false;
        copyTextToClipboard(textToCopy).then(function () {
            hideDisplacementMeasure();
            showNotice("位移參數已複製");
        }).catch(function () {
            hideDisplacementMeasure();
            showNotice("位移參數複製失敗，請確認瀏覽器剪貼簿權限。");
        });
    }

    function cancelDisplacementMeasure() {
        state.displacementMeasure.cancelled = true;
        state.displacementMeasure.dragging = false;
        state.displacementMeasure.active = false;
        hideDisplacementMeasure();
    }

    function hideDisplacementMeasure() {
        if (els.displacementMeasureOverlay) {
            els.displacementMeasureOverlay.hidden = true;
        }
    }

    function copyTextToClipboard(value) {
        if (navigator.clipboard && window.isSecureContext) {
            return navigator.clipboard.writeText(value);
        }
        return new Promise(function (resolve, reject) {
            const textarea = document.createElement("textarea");
            textarea.value = value;
            textarea.setAttribute("readonly", "readonly");
            textarea.style.position = "fixed";
            textarea.style.left = "-9999px";
            document.body.appendChild(textarea);
            textarea.select();
            try {
                const copied = document.execCommand("copy");
                document.body.removeChild(textarea);
                copied ? resolve() : reject(new Error("copy failed"));
            } catch (error) {
                document.body.removeChild(textarea);
                reject(error);
            }
        });
    }

    function applyDeveloperHiddenBlocks() {
        const isDeveloper = !!state.developerMode;
        getDeveloperHideBlocks().forEach(function (item) {
            const hiddenForUser = state.devHiddenBlocks.has(item.key);
            document.querySelectorAll("[data-dev-hide-key]").forEach(function (target) {
                if (text(target.getAttribute("data-dev-hide-key")) !== item.key) {
                    return;
                }
                target.classList.toggle("user-hidden", hiddenForUser && !isDeveloper);
                target.classList.toggle("developer-hidden-preview", hiddenForUser && isDeveloper);
                target.setAttribute("data-dev-hidden-active", hiddenForUser ? "true" : "false");
            });
        });
    }

    function syncCustomerMaskToggle() {
        if (!els.customerMaskToggle) {
            return;
        }
        els.customerMaskToggle.checked = !!state.customerMasked;
        const label = els.customerMaskToggle.closest(".privacy-switch");
        if (!label) {
            return;
        }
        const textEl = label.querySelector(".privacy-switch-text");
        if (textEl) {
            textEl.textContent = state.customerMasked ? "加密" : "顯示";
        }
    }

    function syncExposureBasisControls() {
        const isHousehold = state.exposureBasis === "household";
        if (els.exposurePrincipalBtn) {
            els.exposurePrincipalBtn.classList.toggle("active", !isHousehold);
            els.exposurePrincipalBtn.setAttribute("aria-pressed", String(!isHousehold));
        }
        if (els.exposureHouseholdBtn) {
            els.exposureHouseholdBtn.classList.toggle("active", isHousehold);
            els.exposureHouseholdBtn.setAttribute("aria-pressed", String(isHousehold));
        }
        if (els.timelineExposurePrincipalBtn) {
            els.timelineExposurePrincipalBtn.classList.toggle("active", !isHousehold);
            els.timelineExposurePrincipalBtn.setAttribute("aria-pressed", String(!isHousehold));
        }
        if (els.timelineExposureHouseholdBtn) {
            els.timelineExposureHouseholdBtn.classList.toggle("active", isHousehold);
            els.timelineExposureHouseholdBtn.setAttribute("aria-pressed", String(isHousehold));
        }
        if (els.rankExposurePrincipalBtn) {
            els.rankExposurePrincipalBtn.classList.toggle("active", !isHousehold);
            els.rankExposurePrincipalBtn.setAttribute("aria-pressed", String(!isHousehold));
        }
        if (els.rankExposureHouseholdBtn) {
            els.rankExposureHouseholdBtn.classList.toggle("active", isHousehold);
            els.rankExposureHouseholdBtn.setAttribute("aria-pressed", String(isHousehold));
        }
    }

    function syncTimelineControls() {
        const isTimeline = state.timelineMode === "timeline";
        document.body.classList.toggle("timeline-active", isTimeline);
        if (els.viewModeToggleBtn) {
            const label = els.viewModeToggleBtn.querySelector(".mode-switch-text") || els.viewModeToggleBtn.querySelector(".tool-label");
            if (label) {
                label.textContent = isTimeline ? "趨勢分析" : "當期分析";
            }
            els.viewModeToggleBtn.querySelectorAll("[data-view-mode]").forEach(function (button) {
                const active = button.getAttribute("data-view-mode") === state.timelineMode;
                button.classList.toggle("active", active);
                button.setAttribute("aria-pressed", String(active));
            });
            els.viewModeToggleBtn.classList.toggle("active", isTimeline);
            els.viewModeToggleBtn.title = isTimeline ? "切換為當期分析" : "切換為趨勢分析";
            els.viewModeToggleBtn.setAttribute("aria-label", els.viewModeToggleBtn.title);
        }
        if (els.selectedPeriodSelect && els.selectedPeriodSelect.value !== state.selectedPeriodId) {
            els.selectedPeriodSelect.value = state.selectedPeriodId;
        }
        if (els.comparePeriodSelect && els.comparePeriodSelect.value !== (state.comparePeriodId || "")) {
            els.comparePeriodSelect.value = state.comparePeriodId || "";
        }
        document.querySelectorAll(".timeline-only-control").forEach(function (el) {
            el.hidden = !isTimeline;
        });
        document.querySelectorAll(".snapshot-exposure-basis-control").forEach(function (el) {
            el.hidden = isTimeline;
        });
        document.querySelectorAll(".snapshot-only-tool").forEach(function (el) {
            el.hidden = isTimeline;
        });
        document.querySelectorAll(".snapshot-only-panel").forEach(function (el) {
            el.hidden = isTimeline;
        });
        if (els.refreshFiltersBtn) {
            els.refreshFiltersBtn.hidden = isTimeline;
        }
        if (els.clearFilters) {
            els.clearFilters.hidden = isTimeline;
        }
        if (els.timelineChangeModeSwitch) {
            els.timelineChangeModeSwitch.querySelectorAll("[data-value]").forEach(function (button) {
                button.classList.toggle("active", button.getAttribute("data-value") === state.timelineChangeMode);
            });
        }
    }

    function syncAmountUnitControl() {
        if (!els.amountUnitToggleBtn) {
            return;
        }
        const isYi = state.amountUnit === "yi";
        const label = els.amountUnitToggleBtn.querySelector(".tool-label");
        if (label) {
            label.textContent = isYi ? "億元" : "仟元";
        }
        els.amountUnitToggleBtn.querySelectorAll("[data-amount-unit]").forEach(function (item) {
            const active = item.getAttribute("data-amount-unit") === state.amountUnit;
            item.classList.toggle("active", active);
            item.setAttribute("aria-pressed", String(active));
        });
        els.amountUnitToggleBtn.classList.toggle("active", isYi);
        els.amountUnitToggleBtn.title = isYi ? "切換為仟元" : "切換為億元";
        els.amountUnitToggleBtn.setAttribute("aria-label", els.amountUnitToggleBtn.title);
    }

    function setTimelineViewVisible(visible) {
        if (els.timelinePanel) {
            els.timelinePanel.hidden = !visible;
        }
        document.querySelectorAll(".main-grid, .analysis-grid").forEach(function (el) {
            el.hidden = visible;
        });
        if (visible) {
            setAnalysisDrawer(false);
            setCaseDrawer(false);
            setCountyDrawer(false);
            state.mainGridMode = "default";
        }
    }

    function applyUiMode() {
        const isIos = state.uiMode === "ios";
        document.body.classList.toggle("ui-ios", isIos);
        if (els.uiModeDefaultBtn) {
            els.uiModeDefaultBtn.classList.toggle("active", !isIos);
        }
        if (els.uiModeIosBtn) {
            els.uiModeIosBtn.classList.toggle("active", isIos);
        }
    }

    function applyThemeMode() {
        const isDark = state.themeMode === "dark";
        document.body.classList.toggle("theme-dark", isDark);
        document.body.classList.toggle("theme-light", !isDark);
        if (els.themeLightBtn) {
            els.themeLightBtn.classList.toggle("active", !isDark);
            els.themeLightBtn.setAttribute("aria-pressed", String(!isDark));
        }
        if (els.themeDarkBtn) {
            els.themeDarkBtn.classList.toggle("active", isDark);
            els.themeDarkBtn.setAttribute("aria-pressed", String(isDark));
        }
        if (els.themeToggleBtn) {
            const label = els.themeToggleBtn.querySelector(".tool-label");
            if (label) {
                label.textContent = isDark ? "深色" : "淺色";
            }
            els.themeToggleBtn.classList.toggle("active", isDark);
            els.themeToggleBtn.setAttribute("aria-pressed", String(isDark));
            els.themeToggleBtn.title = isDark ? "切換為淺色模式" : "切換為深色模式";
            els.themeToggleBtn.setAttribute("aria-label", els.themeToggleBtn.title);
        }
    }

    function applyFontSize() {
        state.fontSizeConfig = normalizeFontSizeConfig(state.fontSizeConfig);
        const rootStyle = document.documentElement.style;
        rootStyle.setProperty("--font-size-title", state.fontSizeConfig.title + "px");
        rootStyle.setProperty("--font-size-body", state.fontSizeConfig.body + "px");
        rootStyle.setProperty("--font-size-chart", state.fontSizeConfig.chart + "px");
        rootStyle.setProperty("--font-size-map", state.fontSizeConfig.map + "px");
        rootStyle.setProperty("--app-font-family", getAppFontFamily());
        if (els.fontFamilySelect) {
            els.fontFamilySelect.value = state.fontSizeConfig.family;
        }
        if (els.titleFontSizeInput) {
            els.titleFontSizeInput.value = String(state.fontSizeConfig.title);
        }
        if (els.bodyFontSizeInput) {
            els.bodyFontSizeInput.value = String(state.fontSizeConfig.body);
        }
        if (els.chartFontSizeInput) {
            els.chartFontSizeInput.value = String(state.fontSizeConfig.chart);
        }
        if (els.mapFontSizeInput) {
            els.mapFontSizeInput.value = String(state.fontSizeConfig.map);
        }
    }

    function applyToolbarState() {
        var toolbarState = localStorage.getItem("toolbarState");
        if (toolbarState === "collapsed") {
            document.body.classList.remove("toolbar-open");
        } else {
            document.body.classList.add("toolbar-open");
        }
        syncToolbarControl();
    }

    function syncToolbarControl() {
        if (!els.toggleToolbarBtn) {
            return;
        }
        const isOpen = document.body.classList.contains("toolbar-open");
        els.toggleToolbarBtn.classList.toggle("active", isOpen);
        els.toggleToolbarBtn.setAttribute("aria-pressed", String(isOpen));
        els.toggleToolbarBtn.title = isOpen ? "收合快捷工具列" : "展開快捷工具列";
        els.toggleToolbarBtn.setAttribute("aria-label", els.toggleToolbarBtn.title);
    }

    function animateChartsResize() {
        var startTime = performance.now();
        var duration = 400; // transition 時間為 400ms
        var isAnalysisOpen = !!(els.analysisDrawerPanel && els.analysisDrawerPanel.classList.contains("open"));

        function step(timestamp) {
            var progress = timestamp - startTime;
            
            // 主畫面圖表 + 時間軸圖表
            var targetCharts = [charts.map, charts.ltv, charts.coverage]
                .concat(charts.mainSlots || [])
                .concat([charts.timelineTrend, charts.loanRemainCapital, charts.timelineCountyRank, charts.timelineCountyChange, charts.timelineCustomerChange]);
            fitMainSlotChartsToContainers();
            
            // 如果分析抽屜開啟，也包含分析圖表
            if (isAnalysisOpen) {
                targetCharts.push(
                    charts.group, charts.countyRank,
                    charts.industryPie, charts.riskPie,
                    charts.topCustomer, charts.quadrantMatrix, charts.quadrantSummary, charts.ltv2
                );
            }

            targetCharts.forEach(function (chart) {
                if (chart) {
                    chart.resize();
                }
            });
            
            if (progress < duration) {
                requestAnimationFrame(step);
            } else {
                targetCharts.forEach(function (chart) {
                    if (chart) {
                        chart.resize();
                    }
                });
            }
        }
        requestAnimationFrame(step);
    }

    function resizeVisibleCharts() {
        const targetCharts = [charts.map, charts.ltv, charts.coverage]
            .concat(charts.mainSlots || [])
            .concat([charts.timelineTrend, charts.loanRemainCapital, charts.timelineCountyRank, charts.timelineCountyChange, charts.timelineCustomerChange]);
        fitMainSlotChartsToContainers();
        const isAnalysisOpen = !!(els.analysisDrawerPanel && els.analysisDrawerPanel.classList.contains("open"));
        if (isAnalysisOpen) {
            targetCharts.push(
                charts.group, charts.countyRank,
                charts.industryPie, charts.riskPie,
                charts.topCustomer, charts.quadrantMatrix, charts.quadrantSummary, charts.ltv2
            );
        }
        targetCharts.forEach(function (chart) {
            if (chart && typeof chart.resize === "function") {
                chart.resize();
            }
        });
    }

    function fitMainSlotChartsToContainers() {
        (charts.mainSlots || []).forEach(function (slotChart, index) {
            fitMainSlotChartToContainer(index, slotChart);
        });
    }

    function scheduleChartReflow() {
        requestAnimationFrame(function () {
            resizeVisibleCharts();
            requestAnimationFrame(resizeVisibleCharts);
        });
        [80, 220, 460].forEach(function (delay) {
            setTimeout(resizeVisibleCharts, delay);
        });
    }

    function initDashboardResizeObserver() {
        if (typeof ResizeObserver !== "function") {
            return;
        }
        const targets = [
            document.querySelector(".dashboard-shell"),
            document.querySelector(".snapshot-sticky-header"),
            document.querySelector(".secondary-kpis.kpi-bar"),
            els.activeFiltersBar,
            els.timelinePanel,
            els.timelineTrendChart,
            els.loanRemainCapitalChart,
            els.timelineCountyRankChart
        ].filter(Boolean);
        const observer = new ResizeObserver(debounce(function () {
            updateStickyOffsets();
            scheduleChartReflow();
        }, 80));
        targets.forEach(function (target) {
            observer.observe(target);
        });
    }

    function loadTableColumnWidths() {
        try {
            const raw = localStorage.getItem(TABLE_WIDTH_STORAGE_KEY);
            if (!raw) {
                return { case: [], county: [] };
            }
            const parsed = JSON.parse(raw);
            return {
                case: Array.isArray(parsed.case) ? parsed.case : [],
                county: Array.isArray(parsed.county) ? parsed.county : []
            };
        } catch (_error) {
            return { case: [], county: [] };
        }
    }

    function loadTableWidthLocks() {
        try {
            const raw = localStorage.getItem(TABLE_WIDTH_LOCK_STORAGE_KEY);
            if (!raw) {
                return { case: false, county: false };
            }
            const parsed = JSON.parse(raw);
            return {
                case: !!parsed.case,
                county: !!parsed.county
            };
        } catch (_error) {
            return { case: false, county: false };
        }
    }

    function getTableHeaderWidths(tableType) {
        const wrap = tableType === "case" ? els.caseTableWrap : els.countySummaryWrap;
        if (!wrap) {
            return [];
        }
        return Array.from(wrap.querySelectorAll("thead th")).map(function (th) {
            return Math.round(th.getBoundingClientRect().width);
        });
    }

    function applyStoredTableWidths() {
        if (state.tableWidthLocks.case) {
            applyTableWidths("case", state.tableColumnWidths.case);
        }
        if (state.tableWidthLocks.county) {
            applyTableWidths("county", state.tableColumnWidths.county);
        }
    }

    function applyTableWidths(tableType, widths) {
        if (!Array.isArray(widths) || !widths.length) {
            return;
        }
        const wrap = tableType === "case" ? els.caseTableWrap : els.countySummaryWrap;
        if (!wrap) {
            return;
        }
        const headers = wrap.querySelectorAll("thead th");
        headers.forEach(function (th, idx) {
            const width = toNumber(widths[idx]);
            if (width > 0) {
                th.style.width = `${width}px`;
                th.style.minWidth = `${width}px`;
            }
        });
    }

    function persistTableColumnWidths() {
        state.tableColumnWidths.case = getTableHeaderWidths("case");
        state.tableColumnWidths.county = getTableHeaderWidths("county");
        if (!state.tableWidthLocks.case && !state.tableWidthLocks.county) {
            return;
        }
        try {
            localStorage.setItem(TABLE_WIDTH_STORAGE_KEY, JSON.stringify({
                case: state.tableWidthLocks.case ? state.tableColumnWidths.case : [],
                county: state.tableWidthLocks.county ? state.tableColumnWidths.county : []
            }));
            localStorage.setItem(TABLE_WIDTH_LOCK_STORAGE_KEY, JSON.stringify({
                case: !!state.tableWidthLocks.case,
                county: !!state.tableWidthLocks.county
            }));
        } catch (_error) {
            // ignore storage failures
        }
    }

    function syncTableWidthLockButtons() {
        if (els.lockCaseColumns) {
            const on = !!state.tableWidthLocks.case;
            els.lockCaseColumns.textContent = on ? "欄寬鎖定：開" : "欄寬鎖定：關";
            els.lockCaseColumns.classList.toggle("map-lock-on", on);
        }
        if (els.lockCountyColumns) {
            const on = !!state.tableWidthLocks.county;
            els.lockCountyColumns.textContent = on ? "欄寬鎖定：開" : "欄寬鎖定：關";
            els.lockCountyColumns.classList.toggle("map-lock-on", on);
        }
    }

    function toNumber(value) {
        if (typeof value === "number" && Number.isFinite(value)) {
            return value;
        }
        const cleaned = String(value || "0").replace(/,/g, "").replace(/-/g, "0").trim();
        const parsed = Number(cleaned);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function text(value) {
        if (value === null || typeof value === "undefined") {
            return "";
        }
        return String(value).replace(/\ufeff|\u200b|\xa0/g, "").trim();
    }

    function truncateCustomerName(value, maxChars) {
        const normalized = text(value);
        const limit = Number.isFinite(maxChars) ? Math.max(1, Math.floor(maxChars)) : 6;
        if (normalized.length <= limit) {
            return normalized;
        }
        return `${normalized.slice(0, limit)}...`;
    }

    function formatCustomerDisplayName(value) {
        const normalized = text(value);
        if (!normalized || normalized === "-") {
            return normalized || "-";
        }
        return state.customerMasked ? "****" : truncateCustomerName(normalized);
    }

    function formatSensitiveEntityName(value) {
        const normalized = text(value);
        if (!normalized || normalized === "-") {
            return normalized || "-";
        }
        return state.customerMasked ? "****" : normalized;
    }

    function formatCaseNoDisplay(value) {
        const normalized = text(value);
        if (!normalized || normalized === "-") {
            return normalized || "-";
        }
        return state.customerMasked ? "****" : normalized;
    }

    function getCustomerTitle(value) {
        const normalized = text(value);
        if (!normalized || normalized === "-") {
            return normalized || "-";
        }
        return state.customerMasked ? "客戶名稱已加密" : normalized;
    }

    function safeDivide(numerator, denominator) {
        if (!denominator) {
            return 0;
        }
        return numerator / denominator;
    }

    function formatNumber(value) {
        return Math.round(value || 0).toLocaleString("zh-TW");
    }

    function formatK(value) {
        const amount = toNumber(value);
        if (state.amountUnit === "yi") {
            const yi = amount / 100000;
            const text = yi >= 100 ? yi.toFixed(0) : yi.toFixed(1);
            return `${text.replace(/\\.0$/, "")}億元`;
        }
        return `${formatNumber(amount)}仟`;
    }

    function getExportAmountUnitLabel() {
        return state.amountUnit === "yi" ? "\u5104\u5143" : "\u4edf\u5143";
    }

    function formatExportAmount(value) {
        const amount = toNumber(value);
        if (state.amountUnit !== "yi") {
            return amount;
        }
        return Number((amount / 100000).toFixed(2));
    }

    function formatPercent(value) {
        return `${((value || 0) * 100).toFixed(2)}%`;
    }

    function formatWholePercent(value) {
        return `${Math.round((value || 0) * 100)}%`;
    }

    function formatOneDecimalPercent(value) {
        return `${((value || 0) * 100).toFixed(1)}%`;
    }

    function formatShortMoney(value) {
        const amount = toNumber(value);
        if (state.amountUnit === "yi") {
            const yi = amount / 100000; // 1億 = 100,000K
            const text = yi >= 100 ? yi.toFixed(0) : yi.toFixed(1);
            return `${text.replace(/\\.0$/, "")}億元`;
        }
        return `${formatNumber(amount)}仟`;
    }

    function formatFrontPageMoney(value) {
        const amount = toNumber(value);
        const yi = amount / 100000000;
        const text = yi >= 100 ? yi.toFixed(0) : yi.toFixed(1);
        return `${text.replace(/\\.0$/, "")}億元`;
    }

    function formatFrontPageMoneyAxis(value) {
        const amount = toNumber(value);
        return `${(amount / 100000000).toFixed(0)}億`;
    }

    function formatAxisMoney(value) {
        const amount = toNumber(value);
        if (state.amountUnit === "yi") {
            return `${(amount / 100000).toFixed(1).replace(/\.0$/, "")}億`;
        }
        return `${formatNumber(amount)}仟`;
    }

    function escapeHtml(value) {
        return String(value).replace(/[&<>"']/g, function (char) {
            return {
                "&": "&amp;",
                "<": "&lt;",
                ">": "&gt;",
                "\"": "&quot;",
                "'": "&#039;"
            }[char];
        });
    }

    function localeCompare(a, b) {
        return String(a).localeCompare(String(b), "zh-Hant");
    }

    function descValue(a, b) {
        return b.value - a.value;
    }

    function getExposureSortValue(item) {
        const principal = toNumber(item.principalBalance !== undefined ? item.principalBalance : item.principalValue);
        const approvedUnused = toNumber(item.approvedUnused !== undefined ? item.approvedUnused : item.approvedUnusedValue);
        return state.exposureBasis === "household" ? principal + approvedUnused : principal;
    }

    function shouldHideTemporaryCustomerFromTopCreditRank(item) {
        if (!item || getExposureSortValue(item) <= 0) {
            return false;
        }
        const caseId = item.caseId !== undefined ? item.caseId : item.row && item.row[FIELD.caseId];
        const applicant = item.applicant !== undefined ? item.applicant : item.row && item.row[FIELD.applicant];
        const group = item.group !== undefined ? item.group : item.row && item.row[FIELD.group];
        return isMissingCaseIdValue(caseId) && (isTemporaryZeroCustomerName(applicant) || isTemporaryZeroCustomerName(group));
    }

    function isMissingCaseIdValue(value) {
        if (value === null || typeof value === "undefined") {
            return true;
        }
        const normalized = text(value).toLowerCase();
        return normalized === "" || normalized === "0" || normalized === "null" || normalized === "undefined";
    }

    function isTemporaryZeroCustomerName(value) {
        const normalized = text(value).toLowerCase();
        return normalized === "0" || normalized === "0.0" || normalized === "null";
    }

    function getTimelineMetricMeta(metric) {
        const map = {
            principalBalance: { label: "本金餘額", unit: "money" },
            approvedUnused: { label: "已准未用", unit: "money" },
            householdAmount: { label: "總歸戶", unit: "money" },
            caseCount: { label: "案件數", unit: "count" }
        };
        return map[metric] || null;
    }

    function getTimelineMetricValue(summary, metric) {
        if (!summary) {
            return 0;
        }
        if (metric === "caseCount") {
            return toNumber(summary.caseCount !== undefined ? summary.caseCount : summary.count);
        }
        return toNumber(summary[metric]);
    }

    function formatTimelineMetricValue(value, metric) {
        const meta = getTimelineMetricMeta(metric);
        return meta && meta.unit === "count" ? `${formatNumber(value)}件` : formatShortMoney(value);
    }

    function getTimelineMetricLabel(metric) {
        const meta = getTimelineMetricMeta(metric);
        return meta ? meta.label : "本金餘額";
    }

    function getTimelinePeriodLabel(periodId) {
        const period = state.timelinePeriods.find(function (item) {
            return item.periodId === periodId;
        });
        return period ? period.label : "-";
    }

    function buildTimelineSummary() {
        return state.timelinePeriods.map(function (period) {
            const dataset = state.periodDatasets.get(period.periodId);
            const filteredRows = dataset ? getTimelineRowsFromRawRows(dataset.rawRows) : [];
            const summary = summarizeRows(filteredRows);
            const countySummary = aggregateByCountyForTimeline(filteredRows);
            const customerSummary = aggregateByEntityForTimeline(filteredRows, function (item) {
                return item.applicant || "未命名客戶";
            });
            const groupSummary = aggregateByEntityForTimeline(filteredRows, function (item) {
                return item.group || "未命名集團";
            });
            if (dataset) {
                dataset.baseRows = getBaseFilteredRowsFromRawRows(dataset.rawRows);
                dataset.filteredRows = filteredRows;
                dataset.timelineFiltersDisconnected = !!state.disconnectTimelineFilters;
                dataset.summary = summary;
                dataset.countySummary = countySummary;
                dataset.customerSummary = customerSummary;
                dataset.groupSummary = groupSummary;
            }
            return {
                periodId: period.periodId,
                label: period.label,
                rows: filteredRows,
                summary: summary,
                countySummary: countySummary,
                customerSummary: customerSummary,
                groupSummary: groupSummary
            };
        });
    }

    function aggregateByCountyForTimeline(rows) {
        return aggregateByCountyForExposure(rows).map(function (item) {
            return {
                key: normalizeCountyName(item.county),
                name: item.displayName,
                principalBalance: item.principalBalance,
                approvedUnused: item.approvedUnused,
                householdAmount: item.principalBalance + item.approvedUnused,
                caseCount: item.count || 0
            };
        });
    }

    function aggregateByEntityForTimeline(rows, getKey) {
        const groups = new Map();
        rows.forEach(function (item) {
            const key = text(getKey(item));
            const current = groups.get(key) || {
                key: key,
                name: key,
                principalBalance: 0,
                approvedUnused: 0,
                householdAmount: 0,
                caseCount: 0
            };
            current.principalBalance += item.principalBalance;
            current.approvedUnused += item.approvedUnused;
            current.householdAmount += item.householdAmount;
            current.caseCount += 1;
            groups.set(key, current);
        });
        return Array.from(groups.values());
    }

    function buildTimelineDeltas(currentRows, compareRows, metric) {
        const compareMap = new Map(compareRows.map(function (item) {
            return [item.key, item];
        }));
        const currentMap = new Map(currentRows.map(function (item) {
            return [item.key, item];
        }));
        const keys = new Set(Array.from(compareMap.keys()).concat(Array.from(currentMap.keys())));
        return Array.from(keys).map(function (key) {
            const current = currentMap.get(key) || { key: key, name: key };
            const compare = compareMap.get(key) || { key: key, name: current.name || key };
            const currentValue = getTimelineMetricValue(current, metric);
            const compareValue = getTimelineMetricValue(compare, metric);
            const delta = currentValue - compareValue;
            return {
                key: key,
                name: current.name || compare.name || key,
                currentValue: currentValue,
                compareValue: compareValue,
                value: delta
            };
        });
    }

    function getTimelineChangeStatus(currentValue, compareValue) {
        if (compareValue === 0 && currentValue > 0) {
            return "新增";
        }
        if (currentValue === 0 && compareValue > 0) {
            return "歸零";
        }
        if (currentValue === 0 && compareValue === 0) {
            return "0.00%";
        }
        return formatPercent(safeDivide(currentValue - compareValue, compareValue));
    }

    function getTimelineDeltaClass(delta) {
        if (delta > 0) {
            return "up";
        }
        if (delta < 0) {
            return "down";
        }
        return "flat";
    }

    function debounce(fn, wait) {
        let timer = null;
        return function () {
            const args = arguments;
            clearTimeout(timer);
            timer = setTimeout(function () {
                fn.apply(null, args);
            }, wait);
        };
    }

    function getMainChartCatalog() {
        return [
            { id: "group", title: "前十大授信戶", badge: "" },
            { id: "countyRank", title: "區域授信排行", badge: "" },
            { id: "industryPie", title: "產品/業別佔比", badge: "" },
            { id: "riskPie", title: "風險等級分布", badge: "" },
            { id: "topCustomer", title: "前十大客戶曝險排行", badge: "" },
            { id: "quadrantMatrix", title: "象限矩陣散佈圖", badge: "" },
            { id: "quadrantSummary", title: "供需四象限摘要", badge: "" },
            { id: "ltv2", title: "LTV 區間分布", badge: "" },
            { id: "exposureMixPie", title: "本金餘額 / 已准未用", badge: "" }
        ];
    }

    function renderTimelineDashboard() {
        const timeline = buildTimelineSummary();
        const current = timeline.find(function (item) {
            return item.periodId === state.selectedPeriodId;
        }) || timeline[timeline.length - 1];
        const compare = timeline.find(function (item) {
            return item.periodId === state.comparePeriodId;
        }) || null;

        renderTimelineTrendCards(current, compare);
        renderTimelineTrendChart(timeline);
        renderLoanRemainCapitalChart();
        renderTimelineCountyRankChart(current);
        renderTimelineDeltaCharts(current, compare);
        renderTimelineFlow(current, compare);
        if (els.timelineCompareBadge) {
            els.timelineCompareBadge.textContent = compare ? `比較期：${compare.label}` : "比較期：無";
        }
        if (els.timelineMetricBadge) {
            els.timelineMetricBadge.textContent = "本金餘額 + 已准未用";
        }
    }

    function renderTimelineTrendCards(current, compare) {
        if (!els.timelineTrendCards || !current) {
            return;
        }
        const metrics = ["principalBalance", "approvedUnused", "householdAmount", "caseCount"];
        els.timelineTrendCards.innerHTML = metrics.map(function (metric) {
            const currentValue = getTimelineMetricValue(current.summary, metric);
            const compareValue = compare ? getTimelineMetricValue(compare.summary, metric) : 0;
            const delta = currentValue - compareValue;
            const deltaClass = compare ? getTimelineDeltaClass(delta) : "flat";
            const deltaText = compare
                ? `${delta >= 0 ? "+" : ""}${formatTimelineMetricValue(delta, metric)} / ${getTimelineChangeStatus(currentValue, compareValue)}`
                : "無比較期";
            return `
                <article class="timeline-card">
                    <span>${escapeHtml(getTimelineMetricLabel(metric))}</span>
                    <strong>${escapeHtml(formatTimelineMetricValue(currentValue, metric))}<em class="timeline-delta ${deltaClass}">(${escapeHtml(deltaText)})</em></strong>
                </article>
            `;
        }).join("");
    }

    function renderTimelineTrendChart(timeline) {
        if (!charts.timelineTrend) {
            return;
        }
        const displayTimeline = timeline.slice(-12);
        const barGradient = getActiveBarGradient("countyRank");
        const barLabelColor = isDarkTheme() ? "#ffffff" : "#111827";
        const principalGrowth = displayTimeline.map(function (item, index) {
            if (index === 0) {
                return 0;
            }
            const current = getTimelineMetricValue(item.summary, "principalBalance");
            const previous = getTimelineMetricValue(displayTimeline[index - 1].summary, "principalBalance");
            return safeDivide(current - previous, previous);
        });
        const approvedUnusedGrowth = displayTimeline.map(function (item, index) {
            if (index === 0) {
                return 0;
            }
            const current = getTimelineMetricValue(item.summary, "approvedUnused");
            const previous = getTimelineMetricValue(displayTimeline[index - 1].summary, "approvedUnused");
            return safeDivide(current - previous, previous);
        });
        const zoomEnd = 100;
        charts.timelineTrend.setOption({
            grid: { left: 54, right: 72, top: 58, bottom: 76, containLabel: true },
            legend: {
                top: 0,
                right: 8,
                textStyle: { color: getChartTextColor(), fontSize: getFSD(11) }
            },
            tooltip: {
                trigger: "axis",
                confine: true,
                appendToBody: false,
                formatter: function (params) {
                    const principal = params.find(function (item) { return item.seriesName === "本金餘額"; });
                    const approvedUnused = params.find(function (item) { return item.seriesName === "已准未用"; });
                    const principalGrowthItem = params.find(function (item) { return item.seriesName === "本金成長率"; });
                    const approvedGrowthItem = params.find(function (item) { return item.seriesName === "已准未用成長率"; });
                    const principalValue = principal ? toNumber(principal.value) : 0;
                    const approvedUnusedValue = approvedUnused ? toNumber(approvedUnused.value) : 0;
                    return `${escapeHtml(params[0].name)}<br>` +
                        `本金餘額：${escapeHtml(formatTimelineMetricValue(principalValue, "principalBalance"))}<br>` +
                        `已准未用：${escapeHtml(formatTimelineMetricValue(approvedUnusedValue, "approvedUnused"))}<br>` +
                        `總歸戶：${escapeHtml(formatTimelineMetricValue(principalValue + approvedUnusedValue, "householdAmount"))}<br>` +
                        `本金成長率：${escapeHtml(formatPercent(principalGrowthItem ? principalGrowthItem.value : 0))}<br>` +
                        `已准未用成長率：${escapeHtml(formatPercent(approvedGrowthItem ? approvedGrowthItem.value : 0))}`;
                }
            },
            xAxis: {
                type: "category",
                data: displayTimeline.map(function (item) { return item.label; }),
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(11), hideOverlap: true },
                axisLine: { lineStyle: { color: getChartSplitLineColor() } },
                axisTick: { show: false }
            },
            dataZoom: getTimelineDataZoom(displayTimeline.length, zoomEnd),
            yAxis: [
                {
                    type: "value",
                    name: "金額",
                    axisLabel: {
                        color: getChartTextColor(),
                        fontSize: getFSD(10),
                        formatter: function (value) {
                            return formatAxisMoney(value);
                        }
                    },
                    splitLine: { lineStyle: { color: getChartSplitLineColor() } }
                },
                {
                    type: "value",
                    name: "成長率",
                    axisLabel: {
                        color: getChartTextColor(),
                        fontSize: getFSD(10),
                        formatter: function (value) {
                            return formatPercent(value);
                        }
                    },
                    splitLine: { show: false }
                }
            ],
            series: [
                {
                    name: "本金餘額",
                    type: "bar",
                    stack: "exposure",
                    yAxisIndex: 0,
                    barWidth: displayTimeline.length >= 10 ? 18 : 26,
                    barMaxWidth: 30,
                    data: displayTimeline.map(function (item) {
                        return getTimelineMetricValue(item.summary, "principalBalance");
                    }),
                    itemStyle: {
                        color: barGradient[1],
                        borderRadius: [0, 0, 5, 5]
                    },
                    label: {
                        show: true,
                        position: "inside",
                        color: barLabelColor,
                        fontSize: getFSD(10),
                        formatter: function (params) {
                            return formatTimelineMetricValue(params.value, "principalBalance");
                        }
                    }
                },
                {
                    name: "已准未用",
                    type: "bar",
                    stack: "exposure",
                    yAxisIndex: 0,
                    barWidth: displayTimeline.length >= 10 ? 18 : 26,
                    barMaxWidth: 30,
                    data: displayTimeline.map(function (item) {
                        return getTimelineMetricValue(item.summary, "approvedUnused");
                    }),
                    itemStyle: {
                        color: barGradient[0],
                        opacity: 0.34,
                        borderRadius: [5, 5, 0, 0]
                    },
                    label: {
                        show: true,
                        position: "top",
                        color: barLabelColor,
                        fontSize: getFSD(10),
                        formatter: function (params) {
                            return formatTimelineMetricValue(params.value, "approvedUnused");
                        }
                    }
                },
                {
                    name: "本金成長率",
                    type: "line",
                    yAxisIndex: 1,
                    smooth: true,
                    symbolSize: 7,
                    data: principalGrowth,
                    lineStyle: { width: 2.5, color: "#DC2626" },
                    itemStyle: { color: "#DC2626" }
                },
                {
                    name: "已准未用成長率",
                    type: "line",
                    yAxisIndex: 1,
                    smooth: true,
                    symbolSize: 7,
                    data: approvedUnusedGrowth,
                    lineStyle: { width: 2.5, color: "#7C3AED" },
                    itemStyle: { color: "#7C3AED" }
                }
            ]
        }, true);
    }

    function renderLoanRemainCapitalChart() {
        if (!charts.loanRemainCapital) {
            return;
        }
        const rows = state.frontPageLoanRemainCapitalRows || [];
        if (!rows.length) {
            renderChartEmptyState(charts.loanRemainCapital, "無法載入本金餘額結構趨勢資料");
            return;
        }
        const barGradient = getActiveBarGradient("countyRank");
        charts.loanRemainCapital.setOption({
            grid: { left: 58, right: 76, top: 58, bottom: 76, containLabel: true },
            legend: {
                top: 0,
                right: 8,
                textStyle: { color: getChartTextColor(), fontSize: getFSD(11) }
            },
            tooltip: {
                trigger: "axis",
                axisPointer: { type: "shadow" },
                confine: true,
                appendToBody: false,
                formatter: function (params) {
                    const capital = params.find(function (item) { return item.seriesName === "本金餘額"; });
                    const structure = params.find(function (item) { return item.seriesName === "本金餘額結構比"; });
                    return `${escapeHtml(params[0].name)}<br>` +
                        `本金餘額：${escapeHtml(formatFrontPageMoney(capital ? capital.value : 0))}<br>` +
                        `本金餘額結構比：${escapeHtml(formatPercent(structure ? structure.value : 0))}`;
                }
            },
            xAxis: {
                type: "category",
                data: rows.map(function (item) { return item.label; }),
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(11), hideOverlap: true },
                axisLine: { lineStyle: { color: getChartSplitLineColor() } },
                axisTick: { show: false }
            },
            dataZoom: getInsideTimelineDataZoom(),
            yAxis: [
                {
                    type: "value",
                    name: "金額",
                    min: 0,
                    max: LOAN_REMAIN_CAPITAL_AXIS_MAX,
                    interval: 10000000000,
                    axisLabel: {
                        color: getChartTextColor(),
                        fontSize: getFSD(10),
                        formatter: function (value) {
                            return formatFrontPageMoneyAxis(value);
                        }
                    },
                    splitLine: { lineStyle: { color: getChartSplitLineColor() } }
                },
                {
                    type: "value",
                    name: "結構比",
                    min: 0,
                    max: LOAN_REMAIN_CAPITAL_STRUCTURE_AXIS_MAX,
                    interval: 0.05,
                    axisLabel: {
                        color: getChartTextColor(),
                        fontSize: getFSD(10),
                        formatter: function (value) {
                            return formatPercent(value);
                        }
                    },
                    splitLine: { show: false }
                }
            ],
            series: [
                {
                    name: "本金餘額",
                    type: "bar",
                    yAxisIndex: 0,
                    barMaxWidth: 30,
                    data: rows.map(function (item) { return item.loanRemainCapital; }),
                    itemStyle: {
                        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                            { offset: 0, color: barGradient[0] },
                            { offset: 1, color: barGradient[1] }
                        ]),
                        borderRadius: [5, 5, 0, 0]
                    },
                    label: {
                        show: true,
                        position: "top",
                        color: isDarkTheme() ? "#ffffff" : "#111827",
                        fontSize: getFSD(10),
                        formatter: function (params) {
                            return formatFrontPageMoneyAxis(params.value);
                        }
                    }
                },
                {
                    name: "本金餘額結構比",
                    type: "line",
                    yAxisIndex: 1,
                    smooth: true,
                    symbolSize: 7,
                    data: rows.map(function (item) { return item.loanRemainCapitalStructure; }),
                    lineStyle: { width: 2.5, color: "#F97316" },
                    itemStyle: { color: "#F97316" },
                    label: {
                        show: true,
                        position: "top",
                        distance: 12,
                        offset: [14, -6],
                        color: getChartTextColor(),
                        fontSize: getFSD(10),
                        formatter: function (params) {
                            return formatPercent(params.value);
                        }
                    },
                    markLine: {
                        silent: true,
                        symbol: "none",
                        lineStyle: {
                            color: "#DC2626",
                            width: 2,
                            type: "dashed"
                        },
                        label: {
                            show: true,
                            formatter: "20% 示警線",
                            position: "insideEndTop",
                            color: "#DC2626",
                            fontSize: getFSD(10),
                            fontWeight: 800
                        },
                        data: [{ yAxis: 0.2 }]
                    }
                }
            ]
        }, true);
    }

    function getTimelineDataZoom(itemCount, endValue) {
        return [
            {
                type: "inside",
                xAxisIndex: 0,
                zoomLock: false,
                filterMode: "none"
            },
            {
                type: "slider",
                xAxisIndex: 0,
                height: 18,
                bottom: 18,
                start: 0,
                end: endValue,
                filterMode: "none",
                brushSelect: false,
                borderColor: getChartSplitLineColor(),
                fillerColor: isDarkTheme() ? "rgba(96, 165, 250, 0.22)" : "rgba(37, 99, 235, 0.18)",
                handleStyle: {
                    color: isDarkTheme() ? "#93C5FD" : "#2563EB",
                    borderColor: isDarkTheme() ? "#BFDBFE" : "#1D4ED8"
                },
                moveHandleStyle: {
                    color: isDarkTheme() ? "#93C5FD" : "#2563EB"
                },
                textStyle: {
                    color: getChartTextColor(),
                    fontSize: getFSD(10)
                },
                showDetail: false
            }
        ];
    }

    function getInsideTimelineDataZoom() {
        return [{
            type: "inside",
            xAxisIndex: 0,
            zoomLock: false,
            filterMode: "none"
        }];
    }

    function getTimelineDeltaAxisLimit(maxAbs) {
        const rawLimit = Math.max(toNumber(maxAbs), 1);
        const step = Math.max(100000, Math.ceil(rawLimit / 3 / 100000) * 100000);
        return {
            max: step * 3,
            interval: step
        };
    }

    function formatAxisHundredMillion(value) {
        return formatAxisMoney(value);
    }

    function renderTimelineCountyRankChart(current) {
        if (!charts.timelineCountyRank) {
            return;
        }
        const drilldownCounty = normalizeCountyName(state.timelineRankDrilldownCounty);
        if (els.timelineCountyRankTitle) {
            els.timelineCountyRankTitle.textContent = drilldownCounty ? `${displayCountyName(drilldownCounty)}行政區授信排行` : "區域授信排行";
        }
        if (els.timelineCountyRankBackBtn) {
            els.timelineCountyRankBackBtn.hidden = !drilldownCounty;
        }
        if (!current || !Array.isArray(current.countySummary)) {
            renderChartEmptyState(charts.timelineCountyRank, "暫無當期區域授信資料");
            return;
        }
        if (drilldownCounty) {
            const districtRows = buildDistrictAmountRows(drilldownCounty, current.rows || []).rows.filter(function (item) {
                return toNumber(item.value) > 0;
            }).reverse();
            if (!districtRows.length) {
                renderChartEmptyState(charts.timelineCountyRank, `${displayCountyName(drilldownCounty)}暫無行政區授信資料`);
                return;
            }
            const chartDom = charts.timelineCountyRank.getDom && charts.timelineCountyRank.getDom();
            if (chartDom) {
                chartDom.style.height = Math.max(420, districtRows.length * 32 + 96) + "px";
                charts.timelineCountyRank.resize();
            }
            const districtTotalValue = districtRows.reduce(function (sum, item) {
                return sum + getExposureSortValue(item);
            }, 0);
            renderHorizontalBar(charts.timelineCountyRank, districtRows.map(function (item) {
                return {
                    name: item.district,
                    rawName: item.district,
                    principalValue: item.principalValue,
                    approvedUnusedValue: item.approvedUnusedValue,
                    value: getExposureSortValue(item)
                };
            }), getActiveBarGradient("countyRank"), {
                forceFourTicks: true,
                showPercentage: true,
                totalValue: districtTotalValue,
                exposureBasis: state.exposureBasis,
                showTooltip: false
            });
            return;
        }
        const rows = current.countySummary.filter(function (item) {
            return getExposureSortValue(item) > 0;
        }).sort(function (a, b) {
            return getExposureSortValue(b) - getExposureSortValue(a);
        }).reverse();
        if (!rows.length) {
            renderChartEmptyState(charts.timelineCountyRank, "暫無當期區域授信資料");
            return;
        }
        const chartDom = charts.timelineCountyRank.getDom && charts.timelineCountyRank.getDom();
        if (chartDom) {
            chartDom.style.height = Math.max(420, rows.length * 28 + 86) + "px";
            charts.timelineCountyRank.resize();
        }
        const totalValue = rows.reduce(function (sum, item) {
            return sum + getExposureSortValue(item);
        }, 0);
        renderHorizontalBar(charts.timelineCountyRank, rows.map(function (item) {
            return {
                name: item.name,
                rawName: item.name,
                principalValue: item.principalBalance,
                approvedUnusedValue: item.approvedUnused,
                value: getExposureSortValue(item)
            };
        }), getActiveBarGradient("countyRank"), {
            forceFourTicks: true,
            showPercentage: true,
            totalValue: totalValue,
            exposureBasis: state.exposureBasis,
            showTooltip: false
        });
    }

    function renderTimelineDeltaCharts(current, compare) {
        const countyMetric = state.exposureBasis === "household" ? "householdAmount" : FIXED_TIMELINE_METRIC;
        const customerMetric = FIXED_TIMELINE_METRIC;
        const emptyMessage = compare ? "暫無變化資料" : "需要至少兩期資料才能比較";
        if (!current || !compare) {
            renderChartEmptyState(charts.timelineCountyChange, emptyMessage);
            renderChartEmptyState(charts.timelineCustomerChange, emptyMessage);
            return;
        }
        const countyDeltas = buildTimelineDeltas(current.countySummary, compare.countySummary, countyMetric);
        const customerDeltas = buildTimelineDeltas(current.customerSummary, compare.customerSummary, customerMetric);
        renderTimelineDeltaBar(charts.timelineCountyChange, countyDeltas, countyMetric, {
            minHeight: 420,
            includeZero: true
        });
        renderTimelineDeltaBar(charts.timelineCustomerChange, customerDeltas, customerMetric, {
            limit: 10,
            formatName: formatSensitiveEntityName
        });
    }

    function renderTimelineDeltaBar(chart, rows, metric, options) {
        if (!chart) {
            return;
        }
        const config = options || {};
        const formatName = typeof config.formatName === "function" ? config.formatName : function (value) {
            return value;
        };
        const sortedRows = rows.filter(function (item) {
            return config.includeZero || item.value !== 0;
        }).sort(function (a, b) {
            return state.timelineChangeMode === "decrease" ? a.value - b.value : b.value - a.value;
        });
        const sorted = (config.limit ? sortedRows.slice(0, config.limit) : sortedRows).reverse();

        if (!sorted.length) {
            renderChartEmptyState(chart, "暫無變化資料");
            return;
        }
        const chartDom = chart.getDom && chart.getDom();
        if (chartDom) {
            chartDom.style.height = Math.max(config.minHeight || 280, sorted.length * 28 + 86) + "px";
            chart.resize();
        }

        const maxAbs = Math.max.apply(null, sorted.map(function (item) {
            return Math.abs(item.value);
        })) || 1;
        const axisLimit = getTimelineDeltaAxisLimit(maxAbs);
        chart.setOption({
            grid: { left: "5%", right: "5%", top: 22, bottom: 28, containLabel: true },
            tooltip: {
                trigger: "axis",
                axisPointer: { type: "shadow" },
                formatter: function (params) {
                    const item = params[0].data;
                    return `${escapeHtml(item.displayName)}<br>` +
                        `本期：${escapeHtml(formatTimelineMetricValue(item.currentValue, metric))}<br>` +
                        `比較期：${escapeHtml(formatTimelineMetricValue(item.compareValue, metric))}<br>` +
                        `差額：${escapeHtml(formatTimelineMetricValue(item.value, metric))}<br>` +
                        `變化：${escapeHtml(getTimelineChangeStatus(item.currentValue, item.compareValue))}`;
                }
            },
            xAxis: {
                type: "value",
                min: -axisLimit.max,
                max: axisLimit.max,
                interval: axisLimit.interval,
                splitNumber: 6,
                axisLabel: {
                    color: getChartTextColor(),
                    fontSize: getFSD(12),
                    formatter: function (value) {
                        return metric === "caseCount" ? formatNumber(value) : formatAxisHundredMillion(value);
                    },
                    hideOverlap: true
                },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            yAxis: {
                type: "category",
                data: sorted.map(function (item) { return formatName(item.name); }),
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(12), width: 92, overflow: "truncate" },
                axisTick: { show: false },
                axisLine: { show: false }
            },
            series: [{
                type: "bar",
                data: sorted.map(function (item) {
                    const displayName = formatName(item.name);
                    return {
                        name: displayName,
                        displayName: displayName,
                        value: item.value,
                        currentValue: item.currentValue,
                        compareValue: item.compareValue,
                        itemStyle: {
                            color: item.value >= 0 ? "#EF4444" : "#10B981",
                            borderRadius: item.value >= 0 ? [0, 5, 5, 0] : [5, 0, 0, 5]
                        }
                    };
                }),
                barMaxWidth: 18,
                label: {
                    show: true,
                    position: "right",
                    color: getChartTextColor(),
                    fontSize: getFSD(10),
                    formatter: function (params) {
                        const prefix = params.value > 0 ? "+" : "";
                        return prefix + formatTimelineMetricValue(params.value, metric);
                    }
                }
            }]
        }, true);
    }

    function renderTimelineFlow(current, compare) {
        if (!els.timelineFlowCards || !els.timelineFlowList) {
            return;
        }
        if (!current || !compare) {
            els.timelineFlowCards.innerHTML = "";
            els.timelineFlowList.innerHTML = "<p class=\"table-summary\">需要至少兩期資料才能比較淨增加 / 淨減少。</p>";
            return;
        }
        const metric = FIXED_TIMELINE_METRIC;
        const compareMap = new Map(compare.customerSummary.map(function (item) {
            return [item.key, item];
        }));
        const currentMap = new Map(current.customerSummary.map(function (item) {
            return [item.key, item];
        }));
        const netIncrease = [];
        const netDecrease = [];
        currentMap.forEach(function (item, key) {
            const currentValue = getTimelineMetricValue(item, metric);
            const compareValue = getTimelineMetricValue(compareMap.get(key), metric);
            const delta = currentValue - compareValue;
            if (delta > 0) {
                netIncrease.push({ name: item.name, value: delta });
            }
        });
        compareMap.forEach(function (item, key) {
            if (currentMap.has(key)) {
                return;
            }
            const compareValue = getTimelineMetricValue(item, metric);
            if (compareValue > 0) {
                netDecrease.push({ name: item.name, value: compareValue });
            }
        });
        currentMap.forEach(function (item, key) {
            if (!compareMap.has(key)) {
                return;
            }
            const currentValue = getTimelineMetricValue(item, metric);
            const compareValue = getTimelineMetricValue(compareMap.get(key), metric);
            const delta = currentValue - compareValue;
            if (delta < 0) {
                netDecrease.push({ name: item.name, value: Math.abs(delta) });
            }
        });
        netIncrease.sort(descValue);
        netDecrease.sort(descValue);
        const netIncreaseTotal = netIncrease.reduce(function (sum, item) { return sum + item.value; }, 0);
        const netDecreaseTotal = netDecrease.reduce(function (sum, item) { return sum + item.value; }, 0);
        els.timelineFlowCards.innerHTML = [
            { label: "淨增加戶數", value: formatNumber(netIncrease.length) + "戶" },
            { label: "淨減少戶數", value: formatNumber(netDecrease.length) + "戶" },
            { label: "淨增加金額", value: formatTimelineMetricValue(netIncreaseTotal, metric) },
            { label: "淨減少金額", value: formatTimelineMetricValue(netDecreaseTotal, metric) }
        ].map(function (item) {
            return `<article class="timeline-flow-card"><span>${escapeHtml(item.label)}</span><strong>${escapeHtml(item.value)}</strong></article>`;
        }).join("");
        function renderList(title, rows) {
            const list = rows.map(function (item) {
                return `<li><span>${escapeHtml(formatSensitiveEntityName(item.name))}</span><strong>${escapeHtml(formatTimelineMetricValue(item.value, metric))}</strong></li>`;
            }).join("") || "<li><span>無資料</span><strong>-</strong></li>";
            return `<article><span>${escapeHtml(title)}</span><ul>${list}</ul></article>`;
        }
        els.timelineFlowList.innerHTML = renderList("淨增加", netIncrease) + renderList("淨減少", netDecrease);
    }

    function renderMainPanelChartPicker() {
        if (!els.mainPanelChartPicker) {
            return;
        }
        const selected = new Set(state.mainPanelChartSlots);
        els.mainPanelChartPicker.innerHTML = getMainChartCatalog().map(function (item) {
            const checked = selected.has(item.id) ? " checked" : "";
            return `<label class="chart-picker-item"><input type="checkbox" data-main-chart-id="${escapeHtml(item.id)}"${checked}>${escapeHtml(item.title)}</label>`;
        }).join("");
    }

    function withChartBinding(chartKey, targetChart, renderFn) {
        const previous = charts[chartKey];
        charts[chartKey] = targetChart;
        try {
            renderFn();
        } finally {
            charts[chartKey] = previous;
        }
    }

    function fitMainSlotChartToContainer(index, slotChart) {
        if (index !== 2 || !slotChart || !slotChart.getDom) {
            return;
        }
        const chartDom = slotChart.getDom();
        const slot = chartDom && chartDom.closest ? chartDom.closest(".main-slot") : null;
        if (!slot) {
            return;
        }
        const slotHeight = slot.getBoundingClientRect().height;
        if (!slotHeight) {
            return;
        }
        const slotStyle = window.getComputedStyle(slot);
        const titleRow = slot.querySelector(".section-title-row");
        const titleStyle = titleRow ? window.getComputedStyle(titleRow) : null;
        const occupiedHeight = [
            parseFloat(slotStyle.paddingTop) || 0,
            parseFloat(slotStyle.paddingBottom) || 0,
            titleRow ? titleRow.getBoundingClientRect().height : 0,
            titleStyle ? (parseFloat(titleStyle.marginBottom) || 0) : 0
        ].reduce(function (sum, value) {
            return sum + value;
        }, 0);
        const availableHeight = Math.floor(slotHeight - occupiedHeight);
        if (availableHeight <= 0) {
            return;
        }
        chartDom.style.height = `${Math.max(80, availableHeight)}px`;
        slotChart.resize();
    }

    function renderMainPanelCharts(rows) {
        state.mainPanelChartSlots = normalizeMainPanelChartSlots(state.mainPanelChartSlots);
        const activeSlots = getDefaultMainPanelChartSlots();
        const slotTitles = [els.mainSlotTitle1, els.mainSlotTitle2, els.mainSlotTitle3, els.mainSlotTitle4];
        const slotBadges = [els.mainSlotBadge1, els.mainSlotBadge2, els.mainSlotBadge3, els.mainSlotBadge4];
        const slotToggles = [els.mainSlotCoverageToggle1, els.mainSlotCoverageToggle2, els.mainSlotCoverageToggle3, els.mainSlotCoverageToggle4];
        const slotLtvBasisToggles = [null, null, null, els.mainSlotLtvBasisToggle4];
        const catalogMap = new Map(getMainChartCatalog().map(function (item) { return [item.id, item]; }));
        document.querySelectorAll(".slot-coverage-toggle .segmented-btn").forEach(function (btn) {
            btn.classList.toggle("active", btn.getAttribute("data-value") === state.coverageMode);
        });

        (charts.mainSlots || []).forEach(function (slotChart, index) {
            const slotId = activeSlots[index];
            const slotMeta = catalogMap.get(slotId);
            if (!slotChart) {
                return;
            }
            if (!slotMeta) {
                if (slotTitles[index]) {
                    slotTitles[index].textContent = "未配置圖表";
                }
                if (slotBadges[index]) {
                    slotBadges[index].hidden = true;
                    slotBadges[index].textContent = "";
                }
                if (slotToggles[index]) {
                    slotToggles[index].hidden = true;
                }
                if (slotLtvBasisToggles[index]) {
                    slotLtvBasisToggles[index].hidden = true;
                }
                renderChartEmptyState(slotChart, "未配置圖表");
                return;
            }
            if (slotTitles[index]) {
                const drilldownCounty = slotId === "countyRank" ? normalizeCountyName(state.countyRankDrilldownCounty) : "";
                slotTitles[index].textContent = drilldownCounty ? `${displayCountyName(drilldownCounty)}行政區授信排行` : slotMeta.title;
            }
            if (index === 0 && els.countyRankBackBtn) {
                els.countyRankBackBtn.hidden = slotId !== "countyRank" || !state.countyRankDrilldownCounty;
            }
            renderMainSlotMarketMeta(index, slotId);
            if (slotBadges[index]) {
                slotBadges[index].hidden = !slotMeta.badge;
                slotBadges[index].textContent = slotMeta.badge || "";
                slotBadges[index].removeAttribute("data-topn-target");
                slotBadges[index].removeAttribute("title");
            }
            if (slotToggles[index]) {
                slotToggles[index].hidden = slotId !== "coverage";
            }
            if (slotLtvBasisToggles[index]) {
                slotLtvBasisToggles[index].hidden = slotId !== "ltv2";
            }

            fitMainSlotChartToContainer(index, slotChart);
            const chartRows = getRowsForChart(slotId);
            if (slotId === "group") withChartBinding("group", slotChart, function () { renderGroupChart(chartRows, 10); });
            else if (slotId === "countyRank") withChartBinding("countyRank", slotChart, function () { renderCountyRanking(chartRows); });
            else if (slotId === "industryPie") withChartBinding("industryPie", slotChart, function () { renderIndustryPie(chartRows); });
            else if (slotId === "riskPie") withChartBinding("riskPie", slotChart, function () { renderRiskPie(chartRows); });
            else if (slotId === "topCustomer") withChartBinding("topCustomer", slotChart, function () { renderTopCustomerChart(chartRows); });
            else if (slotId === "quadrantMatrix") withChartBinding("quadrantMatrix", slotChart, function () { renderQuadrantMatrix(chartRows); });
            else if (slotId === "quadrantSummary") withChartBinding("quadrantSummary", slotChart, function () { renderQuadrantSummary(chartRows); });
            else if (slotId === "ltv2") withChartBinding("ltv2", slotChart, function () { renderLtv2(chartRows); });
            else if (slotId === "coverage") withChartBinding("coverage", slotChart, function () { renderCoverageDistribution(chartRows); });
            else if (slotId === "exposureMixPie") withChartBinding("exposureMixPie", slotChart, function () { renderExposureMixPie(chartRows); });

            bindMainSlotChartEvents(slotChart, slotId);
        });
    }

    function renderMainSlotMarketMeta(index, slotId) {
        if (index !== 2 || !els.mainSlotMarketMeta3) {
            return;
        }
        if (slotId !== "quadrantSummary") {
            els.mainSlotMarketMeta3.hidden = true;
            els.mainSlotMarketMeta3.textContent = "";
            return;
        }
        const singleCounty = getSingleFilterValue("county");
        const vacancy = getVacancyMetrics(singleCounty);
        const selectedCounties = getArrayFilterValues("county").map(displayCountyName);
        const areaLabel = selectedCounties.length ? selectedCounties.join("、") : "全台";
        const districtLines = singleCounty ? getSelectedDistrictVacancyMetrics(singleCounty).map(function (item) {
            return `<span class="market-meta-line">${escapeHtml(item.district)} 空屋率${escapeHtml(formatOneDecimalPercent(item.vacancyRate))} / 待售新成屋率${escapeHtml(formatOneDecimalPercent(item.unsoldRate))}</span>`;
        }) : [];
        els.mainSlotMarketMeta3.hidden = false;
        els.mainSlotMarketMeta3.innerHTML = [
            `<span class="market-meta-line">${escapeHtml(areaLabel)} 空屋率${escapeHtml(formatOneDecimalPercent(vacancy.vacancyRate))} / 待售新成屋率${escapeHtml(formatOneDecimalPercent(vacancy.unsoldRate))}</span>`
        ].concat(districtLines).concat([
            `<span class="market-meta-line">資料日期${escapeHtml(VACANCY_DATA_DATE_LABEL)}</span>`
        ]).join("");
    }

    function bindMainSlotChartEvents(chart, slotId) {
        if (!chart) {
            return;
        }
        chart.off("click");
        chart.off("dblclick");
        chart.on("click", function (params) {
            if (slotId === "group") {
                const group = params.data && params.data.rawName ? params.data.rawName : (params.name || "");
                state.filters.group = state.filters.group === group ? "" : group;
                render();
                return;
            }
            if (slotId === "countyRank") {
                const rawName = params.data && params.data.rawName ? params.data.rawName : (params.name || "");
                if (state.countyRankDrilldownCounty) {
                    if (params.event && params.event.event && params.event.event.detail > 1) {
                        return;
                    }
                    toggleDistrictFromCountyRank(state.countyRankDrilldownCounty, rawName);
                    return;
                }
                const county = normalizeCountyName(rawName);
                if (county && !state.countyRankDrilldownCounty && !(params.event && params.event.event && params.event.event.detail > 1)) {
                    clearCountyRankClickTimer();
                    state.countyRankClickTimer = window.setTimeout(function () {
                        toggleArrayFilterValue("county", county);
                        syncFilterControls();
                        render();
                        state.countyRankClickTimer = null;
                    }, 220);
                }
                return;
            }
            if (slotId === "coverage") {
                const bucketId = params.data && params.data.bucketId ? params.data.bucketId : "";
                if (bucketId) {
                    state.filters.coverageBucket = state.filters.coverageBucket === bucketId ? "" : bucketId;
                    syncFilterControls();
                    render();
                }
                return;
            }
            if (slotId === "industryPie") {
                const name = params.name || "";
                state.filters.industry = [];
                state.filters.industryLabel = state.filters.industryLabel === name ? "" : name;
                syncFilterControls();
                render();
                return;
            }
            if (slotId === "riskPie" && params.data && params.data.key) {
                const key = params.data.key;
                state.filters.riskLevel = state.filters.riskLevel === key ? "" : key;
                render();
                return;
            }
            if (slotId === "quadrantMatrix" && params.data && params.data.raw && params.data.raw.applicant) {
                const customer = params.data.raw.applicant;
                state.filters.customer = state.filters.customer === customer ? "" : customer;
                render();
                return;
            }
            if (slotId === "quadrantSummary" && params.data && params.data.qid) {
                toggleArrayFilter(state.filters.quadrants, params.data.qid);
                syncFilterControls();
                render();
                return;
            }
            if (slotId === "topCustomer") {
                const rawCustomer = params.data && params.data.rawName ? params.data.rawName : (params.name || "");
                if (rawCustomer) {
                    state.filters.customer = state.filters.customer === rawCustomer ? "" : rawCustomer;
                    render();
                }
                return;
            }
            if (slotId === "ltv2" && params.data && params.data.bucketId) {
                const ltvBucketId = params.data.bucketId;
                toggleArrayFilterValue("ltvBucket", ltvBucketId);
                syncFilterControls();
                render();
            }
        });
        chart.on("dblclick", function (params) {
            if (slotId === "countyRank") {
                handleCountyRankDrilldownDblclick(params);
            }
        });
        if (slotId === "countyRank") {
            const zr = chart.getZr && chart.getZr();
            if (zr && !chart.__countyRankBlankDblclickBound) {
                zr.on("dblclick", function (event) {
                    if (state.countyRankDrilldownCounty && !event.target) {
                        closeCountyRankDrilldown();
                    }
                });
                chart.__countyRankBlankDblclickBound = true;
            }
        }
    }

    // ==========================================================================
    // 進階分析圖表渲染函數群 (Advanced Analytics Chart Renderers)
    // ==========================================================================

    function renderAnalysisCharts(rows) {
        if (!charts.analysisInited) {
            initAnalysisCharts();
        }
        if (els.analysisGroupChart) renderGroupChart(getRowsForChart("group"));
        if (els.analysisCountyRankChart) renderCountyRanking(getRowsForChart("countyRank"));
        if (els.industryPieChart) renderIndustryPie(getRowsForChart("industryPie"));
        if (els.riskPieChart) renderRiskPie(getRowsForChart("riskPie"));
        if (els.topCustomerChart) renderTopCustomerChart(getRowsForChart("topCustomer"));
        if (els.quadrantMatrixChart) renderQuadrantMatrix(getRowsForChart("quadrantMatrix"));
        if (els.quadrantSummaryChart) renderQuadrantSummary(getRowsForChart("quadrantSummary"));
        if (els.ltvBucketChart2) renderLtv2(getRowsForChart("ltv2"));
    }

    function renderExposureMixPie(rows) {
        const summary = summarizeRows(rows);
        const data = [
            { name: "本金餘額", value: summary.principalBalance, itemStyle: { color: "#2563EB" } },
            { name: "已准未用", value: summary.approvedUnused, itemStyle: { color: "#60A5FA" } }
        ];
        if (!data.some(function (item) { return item.value > 0; })) {
            renderChartEmptyState(charts.exposureMixPie, "暫無本金或已准未用數據");
            return;
        }
        charts.exposureMixPie.setOption({
            tooltip: {
                trigger: "item",
                formatter: function (params) {
                    return `${escapeHtml(params.name)}<br>${escapeHtml(formatShortMoney(params.value))}<br>${escapeHtml(formatPercent(params.percent / 100))}`;
                }
            },
            legend: {
                bottom: 0,
                selectedMode: false,
                textStyle: { color: getChartTextColor(), fontSize: getFSD(11) }
            },
            series: [{
                type: "pie",
                radius: ["44%", "68%"],
                center: ["50%", "44%"],
                avoidLabelOverlap: true,
                label: {
                    color: getChartTextColor(),
                    formatter: function (params) {
                        return `${params.name}\n${formatPercent(params.percent / 100)}`;
                    }
                },
                data: data
            }]
        }, true);
    }

    function renderIndustryPie(rows) {
        const grouped = new Map();
        rows.forEach(function (item) {
            grouped.set(item.industry, (grouped.get(item.industry) || 0) + item.principalBalance);
        });
        const data = Array.from(grouped.entries()).map(function ([name, value]) {
            return { name: name, value: value };
        }).sort(descValue);

        if (!data.length || !data.some(function (item) { return item.value > 0; })) {
            renderChartEmptyState(charts.industryPie, "暫無相關業別數據");
            return;
        }

        charts.industryPie.setOption({
            tooltip: {
                trigger: "item",
                formatter: function (params) {
                    return `${escapeHtml(params.name)}<br>本金餘額: ${formatShortMoney(params.value)} (${params.percent.toFixed(2)}%)`;
                }
            },
            legend: {
                orient: "vertical",
                left: "left",
                type: "scroll",
                selectedMode: false,
                textStyle: { color: getChartTextColor(), fontSize: getFSD(11) },
                pageIconColor: getChartMutedColor()
            },
            series: [{
                name: "產品/業別佔比",
                type: "pie",
                radius: ["40%", "70%"],
                center: ["60%", "50%"],
                avoidLabelOverlap: true,
                itemStyle: {
                    borderRadius: 8,
                    borderColor: getChartSeriesBorderColor(),
                    borderWidth: 2
                },
                label: {
                    show: false
                },
                emphasis: {
                    label: {
                        show: true,
                        fontSize: getFSD(14),
                        fontWeight: "bold",
                        formatter: "{b}\n{d}%"
                    }
                },
                data: data.map(function (item) {
                    const selected = state.filters.industryLabel && state.filters.industryLabel === item.name;
                    return {
                        name: item.name,
                        value: item.value,
                        itemStyle: { opacity: state.filters.industryLabel && !selected ? 0.24 : 1 }
                    };
                })
            }]
        }, true);
    }

    function renderRiskPie(rows) {
        const counts = { safe: 0, normal: 0, caution: 0, danger: 0 };
        const balances = { safe: 0, normal: 0, caution: 0, danger: 0 };
        rows.forEach(function (item) {
            const r = item.riskLevel || "safe";
            counts[r] = (counts[r] || 0) + 1;
            balances[r] = (balances[r] || 0) + item.principalBalance;
        });

        const data = [
            { name: "安全", value: balances.safe, key: "safe" },
            { name: "正常", value: balances.normal, key: "normal" },
            { name: "注意", value: balances.caution, key: "caution" },
            { name: "危險", value: balances.danger, key: "danger" }
        ].filter(function (item) {
            return item.value > 0;
        });

        if (!data.length) {
            renderChartEmptyState(charts.riskPie, "暫無相關風險數據");
            return;
        }

        charts.riskPie.setOption({
            tooltip: {
                trigger: "item",
                formatter: function (params) {
                    const count = counts[params.data.key] || 0;
                    return `${params.name}<br>本金餘額: ${formatShortMoney(params.value)} (${params.percent.toFixed(2)}%)<br>案件數: ${count} 筆`;
                }
            },
            legend: {
                orient: "vertical",
                left: "left",
                selectedMode: false,
                textStyle: { color: getChartTextColor(), fontSize: getFSD(11) }
            },
            series: [{
                name: "風險等級分布",
                type: "pie",
                radius: ["40%", "70%"],
                center: ["60%", "50%"],
                avoidLabelOverlap: true,
                itemStyle: {
                    borderRadius: 8,
                    borderColor: getChartSeriesBorderColor(),
                    borderWidth: 2
                },
                label: {
                    show: false
                },
                emphasis: {
                    label: {
                        show: true,
                        fontSize: getFSD(14),
                        fontWeight: "bold",
                        formatter: "{b}\n{d}%"
                    }
                },
                data: data.map(function (item) {
                    return {
                        name: item.name,
                        value: item.value,
                        key: item.key,
                        itemStyle: {
                            color: getRiskThemeColor(item.key),
                            opacity: state.filters.riskLevel && state.filters.riskLevel !== item.key ? 0.24 : 1
                        }
                    };
                })
            }]
        }, true);
    }

    function renderLtvCoverageScatter(rows) {
        const data = rows.filter(function (item) {
            return item.principalLtv > 0 && item.coverageRatio > 0;
        });

        if (!data.length) {
            renderChartEmptyState(charts.ltvScatter, "暫無相關曝險數據");
            return;
        }

        const maxBalance = Math.max.apply(null, data.map(function (item) { return item.principalBalance; })) || 1;

        charts.ltvScatter.setOption({
            grid: { left: 48, right: 28, top: 40, bottom: 44, containLabel: true },
            tooltip: {
                trigger: "item",
                formatter: function (params) {
                    const item = params.data.raw;
                    const name = formatSensitiveEntityName(item.applicant);
                    return `[${escapeHtml(name)}]<br>` +
                           `縣市: ${item.county}<br>` +
                           `業別: ${item.industry}<br>` +
                           `LTV: ${formatPercent(item.principalLtv)}<br>` +
                           `擔保覆蓋率: ${formatPercent(item.coverageRatio)}<br>` +
                           `本金餘額: ${formatShortMoney(item.principalBalance)}`;
                }
            },
            xAxis: {
                type: "value",
                name: "LTV",
                nameTextStyle: { color: getChartMutedColor(), fontSize: getFSD(11) },
                axisLabel: {
                    formatter: function (value) {
                        return `${(value * 100).toFixed(0)}%`;
                    },
                    color: getChartTextColor(),
                    fontSize: getFSD(10)
                },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            yAxis: {
                type: "value",
                name: "覆蓋率",
                nameTextStyle: { color: getChartMutedColor(), fontSize: getFSD(11) },
                axisLabel: {
                    formatter: function (value) {
                        return `${(value * 100).toFixed(0)}%`;
                    },
                    color: getChartTextColor(),
                    fontSize: getFSD(10)
                },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            series: [{
                type: "scatter",
                data: data.map(function (item) {
                    const symbolSize = Math.max(8, Math.min(36, 8 + (item.principalBalance / maxBalance) * 28));
                    return {
                        value: [item.principalLtv, item.coverageRatio],
                        symbolSize: symbolSize,
                        itemStyle: {
                            color: getRiskThemeColor(item.riskLevel),
                            opacity: 0.8,
                            borderColor: getChartSeriesBorderColor(),
                            borderWidth: 1
                        },
                        raw: item
                    };
                })
            }]
        }, true);
    }

    function renderTopCustomerChart(rows, limit) {
        const topN = limit || state.rankTopN.topCustomer || 10;
        const grouped = new Map();
        rows.forEach(function (item) {
            if (shouldHideTemporaryCustomerFromTopCreditRank(item)) {
                return;
            }
            const key = item.applicant || "未命名客戶";
            const current = grouped.get(key) || { name: key, principalBalance: 0, approvedUnused: 0 };
            current.principalBalance += item.principalBalance;
            current.approvedUnused += item.approvedUnused;
            grouped.set(key, current);
        });
        const data = Array.from(grouped.values()).map(function (item) {
            return {
                name: item.name,
                principalValue: item.principalBalance,
                approvedUnusedValue: item.approvedUnused,
                value: getExposureSortValue(item)
            };
        }).sort(descValue).slice(0, topN).reverse();

        if (!data.length || !data.some(function (item) { return item.value > 0; })) {
            renderChartEmptyState(charts.topCustomer, "暫無相關客戶曝險數據");
            return;
        }

        const totalValue = rows.reduce(function (sum, item) {
            return sum + getExposureSortValue(item);
        }, 0);

        renderHorizontalBar(charts.topCustomer, data.map(function (item) {
            return {
                name: formatSensitiveEntityName(item.name),
                rawName: item.name,
                principalValue: item.principalValue,
                approvedUnusedValue: item.approvedUnusedValue,
                value: item.value
            };
        }), getActiveBarGradient("group"), {
            showPercentage: true,
            totalValue: totalValue,
            highlightTop3: true,
            exposureBasis: state.exposureBasis,
            selectedRawName: state.filters.customer,
            showTooltip: false
        });
    }

    function renderQuadrantMatrix(rows) {
        const data = rows.filter(function (item) {
            return item.principalLtv > 0 && item.coverageRatio > 0;
        });

        if (!data.length) {
            renderChartEmptyState(charts.quadrantMatrix, "暫無相關象限數據");
            return;
        }

        const maxBalance = Math.max.apply(null, data.map(function (item) { return item.principalBalance; })) || 1;

        charts.quadrantMatrix.setOption({
            grid: { left: 48, right: 28, top: 40, bottom: 44, containLabel: true },
            tooltip: {
                trigger: "item",
                formatter: function (params) {
                    const item = params.data.raw;
                    const name = formatSensitiveEntityName(item.applicant);
                    return `[${escapeHtml(name)}]<br>` +
                           `縣市: ${item.county}<br>` +
                           `業別: ${item.industry}<br>` +
                           `LTV: ${formatPercent(item.principalLtv)}<br>` +
                           `擔保覆蓋率: ${formatPercent(item.coverageRatio)}<br>` +
                           `本金餘額: ${formatShortMoney(item.principalBalance)}`;
                }
            },
            xAxis: {
                type: "value",
                name: "LTV",
                nameTextStyle: { color: getChartMutedColor(), fontSize: getFSD(11) },
                min: 0,
                max: 1.5,
                axisLabel: {
                    formatter: function (value) {
                        return `${(value * 100).toFixed(0)}%`;
                    },
                    color: getChartTextColor(),
                    fontSize: getFSD(10)
                },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            yAxis: {
                type: "value",
                name: "覆蓋率",
                nameTextStyle: { color: getChartMutedColor(), fontSize: getFSD(11) },
                min: 0,
                max: 3.0,
                axisLabel: {
                    formatter: function (value) {
                        return `${(value * 100).toFixed(0)}%`;
                    },
                    color: getChartTextColor(),
                    fontSize: getFSD(10)
                },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            graphic: [
                {
                    type: "text",
                    left: "12%",
                    top: "15%",
                    style: {
                        text: "II: 健全區 (低LTV/高覆蓋)",
                        font: getCanvasFont("bold", getFSD(11)),
                        fill: "#16835f",
                        opacity: 0.65
                    }
                },
                {
                    type: "text",
                    left: "12%",
                    bottom: "15%",
                    style: {
                        text: "III: 正常區 (低LTV/低覆蓋)",
                        font: getCanvasFont("bold", getFSD(11)),
                        fill: "#1769aa",
                        opacity: 0.65
                    }
                },
                {
                    type: "text",
                    right: "12%",
                    top: "15%",
                    style: {
                        text: "I: 注意區 (高LTV/高覆蓋)",
                        font: getCanvasFont("bold", getFSD(11)),
                        fill: "#c46a12",
                        opacity: 0.65
                    }
                },
                {
                    type: "text",
                    right: "12%",
                    bottom: "15%",
                    style: {
                        text: "IV: 風險區 (高LTV/低覆蓋)",
                        font: getCanvasFont("bold", getFSD(11)),
                        fill: "#b92d2b",
                        opacity: 0.65
                    }
                }
            ],
            series: [{
                type: "scatter",
                data: data.map(function (item) {
                    const symbolSize = Math.max(8, Math.min(36, 8 + (item.principalBalance / maxBalance) * 28));
                    return {
                        value: [item.principalLtv, item.coverageRatio],
                        symbolSize: symbolSize,
                        itemStyle: {
                            color: getRiskThemeColor(item.riskLevel),
                            opacity: 0.8,
                            borderColor: getChartSeriesBorderColor(),
                            borderWidth: 1
                        },
                        raw: item
                    };
                }),
                markLine: {
                    silent: true,
                    lineStyle: { type: "dashed", color: getChartGuideLineColor(), width: 1.5 },
                    symbol: "none",
                    data: [
                        { xAxis: 0.8 },
                        { yAxis: 1.0 }
                    ]
                }
            }]
        }, true);
    }

    function renderQuadrantSummary(rows) {
        if (!charts.quadrantSummary) {
            return;
        }
        const summary = {
            q1: { label: "第一象限", description: "供需雙高警戒區", count: 0, balance: 0, color: "#EF4444", bg: "#FEE2E2", darkColor: "#FF6B6B", darkBg: "#4A2325", value: [1, 1, 2, 2] },
            q2: { label: "第二象限", description: "建商激戰區", count: 0, balance: 0, color: "#F59E0B", bg: "#FEF3C7", darkColor: "#FFD166", darkBg: "#44331A", value: [0, 1, 1, 2] },
            q3: { label: "第三象限", description: "剛需熱區", count: 0, balance: 0, color: "#3B82F6", bg: "#DBEAFE", darkColor: "#4EA8DE", darkBg: "#1A2E40", value: [0, 0, 1, 1] },
            q4: { label: "第四象限", description: "存量去化區", count: 0, balance: 0, color: "#10B981", bg: "#D1FAE5", darkColor: "#06D6A0", darkBg: "#1B362A", value: [1, 0, 2, 1] }
        };

        const districtQuadrantCounty = getActiveDistrictQuadrantCounty();
        const activeDistrictQuadrants = districtQuadrantCounty ? getSelectedDistrictQuadrantIds(districtQuadrantCounty) : new Set();
        rows.forEach(function (item) {
            const qid = getItemEffectiveQuadrantId(item);
            if (!summary[qid]) {
                return;
            }
            summary[qid].count += 1;
            summary[qid].balance += item.principalBalance;
        });

        const items = Object.keys(summary).map(function (key) {
            return summary[key];
        });
        const anyData = items.some(function (entry) { return entry.count > 0 || entry.balance > 0; });
        if (!anyData) {
            renderChartEmptyState(charts.quadrantSummary, "暫無象限摘要數據");
            return;
        }

        const data = ["q2", "q1", "q3", "q4"].map(function (key) {
            const entry = summary[key];
            const selected = state.filters.quadrants.includes(key);
            const districtSelected = activeDistrictQuadrants.has(key);
            return {
                qid: key,
                value: entry.value,
                label: entry.label,
                description: entry.description,
                count: entry.count,
                balance: entry.balance,
                color: entry.color,
                bg: entry.bg,
                darkColor: entry.darkColor,
                darkBg: entry.darkBg,
                selected: selected || districtSelected,
                filterSelected: selected,
                districtSelected: districtSelected
            };
        });
        const hasActiveQuadrant = data.some(function (item) {
            return item.selected;
        });

        charts.quadrantSummary.setOption({
            grid: { left: 12, right: 12, top: 12, bottom: 12, containLabel: false },
            tooltip: {
                trigger: "item",
                formatter: function (params) {
                    const item = params.data || {};
                    return `${escapeHtml(item.label || "")}<br>${escapeHtml(item.description || "")}<br>${formatNumber(item.count || 0)} 件 / ${formatShortMoney(item.balance || 0)}`;
                }
            },
            xAxis: { type: "value", min: 0, max: 2, show: false },
            yAxis: { type: "value", min: 0, max: 2, show: false },
            series: [{
                type: "custom",
                data: data,
                renderItem: function (params, api) {
                    const x0 = api.value(0);
                    const y0 = api.value(1);
                    const x1 = api.value(2);
                    const y1 = api.value(3);
                    const start = api.coord([x0, y1]);
                    const end = api.coord([x1, y0]);
                    const width = Math.max(0, end[0] - start[0]);
                    const height = Math.max(0, end[1] - start[1]);
                    const item = data[params.dataIndex] || {};
                    const centerX = start[0] + width / 2;
                    const centerY = start[1] + height / 2;
                    const strokeColor = item.selected ? "#2563EB" : "rgba(255,255,255,0.95)";
                    const lineWidth = item.selected ? 3 : 2;
                    const darkMode = isDarkTheme();
                    const textColor = darkMode ? item.darkColor : "#1E293B";
                    const subTextColor = darkMode ? item.darkColor : "#1E293B";
                    const fillColor = darkMode ? item.darkBg : item.bg;
                    const children = [{
                        type: "rect",
                        shape: {
                            x: start[0],
                            y: start[1],
                            width: width,
                            height: height
                        },
                        style: {
                            fill: fillColor,
                            opacity: hasActiveQuadrant && !item.selected ? 0.35 : 1,
                            stroke: strokeColor,
                            lineWidth: lineWidth,
                            shadowBlur: item.selected ? 10 : 3,
                            shadowColor: item.selected ? "rgba(37,99,235,0.28)" : "rgba(20,40,70,0.08)"
                        }
                    }];

                    children.push({
                        type: "text",
                        x: centerX,
                        y: centerY - 25,
                        style: {
                            text: item.label,
                            fill: textColor,
                            font: getCanvasFont("800", getFSD(16)),
                            align: "center",
                            verticalAlign: "middle"
                        }
                    }, {
                        type: "text",
                        x: centerX,
                        y: centerY - 3,
                        style: {
                            text: item.description,
                            fill: subTextColor,
                            font: getCanvasFont("700", getFSD(13)),
                            align: "center",
                            verticalAlign: "middle"
                        }
                    }, {
                        type: "text",
                        x: centerX,
                        y: centerY + 23,
                        style: {
                            text: `${formatNumber(item.count)} 件 / ${formatShortMoney(item.balance)}`,
                            fill: subTextColor,
                            font: getCanvasFont("700", getFSD(12)),
                            align: "center",
                            verticalAlign: "middle"
                        }
                    });

                    return {
                        type: "group",
                        children: children
                    };
                }
            }]
        }, true);
    }

    function renderLtv2(rows) {
        const isAmountBasis = state.analysisLtvBasis === "amount";
        const values = new Map(LTV_BUCKETS.map(function (bucket) {
            return [bucket.id, 0];
        }));
        rows.forEach(function (item) {
            if (!item.principalBalance || !item.ltvBucket) {
                return;
            }
            const currentValue = values.get(item.ltvBucket) || 0;
            values.set(item.ltvBucket, currentValue + (isAmountBasis ? item.principalBalance : 1));
        });
        document.querySelectorAll(".ltv-basis-toggle [data-ltv-basis]").forEach(function (btn) {
                const active = btn.getAttribute("data-ltv-basis") === (isAmountBasis ? "amount" : "count");
                btn.classList.toggle("active", active);
                btn.setAttribute("aria-pressed", String(active));
        });
        charts.ltv2.setOption({
            grid: { left: "5%", right: "5%", top: 18, bottom: 28, containLabel: true },
            tooltip: {
                trigger: "axis",
                formatter: function (params) {
                    const item = params && params[0];
                    if (!item) {
                        return "";
                    }
                    const valueText = isAmountBasis ? formatK(item.value) : `${formatNumber(item.value)}件`;
                    return `${item.axisValue}<br>${isAmountBasis ? "本金餘額" : "件數"}：${valueText}`;
                }
            },
            xAxis: {
                type: "category",
                data: LTV_BUCKETS.map(function (bucket) { return bucket.label; }),
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(10), interval: 0 }
            },
            yAxis: {
                type: "value",
                minInterval: isAmountBasis ? undefined : 1,
                axisLabel: {
                    color: getChartTextColor(),
                    fontSize: getFSD(10),
                    formatter: function (value) {
                        return isAmountBasis ? formatAxisMoney(value) : formatNumber(value);
                    }
                },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            series: [{
                name: isAmountBasis ? "本金餘額" : "件數",
                type: "bar",
                barMaxWidth: 42,
                data: LTV_BUCKETS.map(function (bucket) {
                    return {
                        value: values.get(bucket.id) || 0,
                        bucketId: bucket.id,
                        itemStyle: {
                            color: getLtvBucketColor(bucket.id),
                            borderColor: getLtvBucketBorderColor(bucket.id),
                            borderWidth: 1,
                            borderRadius: [6, 6, 0, 0]
                        }
                    };
                })
            }]
        }, true);
    }

    function renderCoverage2(rows) {
        const counts = new Map(COVERAGE_BUCKETS.map(function (bucket) {
            return [bucket.id, 0];
        }));

        const isCaseMode = state.coverageMode !== "county";

        if (isCaseMode) {
            rows.forEach(function (item) {
                if (item.principalBalance <= 0) {
                    return;
                }
                const bucket = getCoverageBucketByRatio(item.coverageRatio || 0);
                counts.set(bucket.id, (counts.get(bucket.id) || 0) + 1);
            });
        } else {
            const counties = aggregateByCounty(rows);
            counties.forEach(function (item) {
                if (item.principalBalance <= 0) {
                    return;
                }
                const avgRatio = safeDivide(item.collateralValue, item.principalBalance);
                const bucket = getCoverageBucketByRatio(avgRatio);
                counts.set(bucket.id, (counts.get(bucket.id) || 0) + 1);
            });
        }

        const maxCount = COVERAGE_BUCKETS.reduce(function (max, bucket) {
            return Math.max(max, counts.get(bucket.id) || 0);
        }, 0);

        const bucketColors = {
            "under-100": ["#ff6b6b", "#b92d2b"],
            "100-120": ["#ffb366", "#c46a12"],
            "120-150": ["#4fc3f7", "#1769aa"],
            "150+": ["#66bb6a", "#16835f"]
        };

        charts.coverage2.setOption({
            grid: { left: 46, right: 18, top: 40, bottom: 28, containLabel: false },
            tooltip: { 
                trigger: "axis",
                formatter: function (params) {
                    const item = params[0];
                    const unit = isCaseMode ? "筆" : "個縣市";
                    return `${item.name}<br>${isCaseMode ? "案件量" : "縣市數"}: ${formatNumber(item.value)} ${unit}`;
                }
            },
            xAxis: {
                type: "category",
                data: COVERAGE_BUCKETS.map(function (bucket) { return bucket.label; }),
                axisLabel: { color: getChartTextColor(), interval: 0, fontSize: getFSD(11) }
            },
            yAxis: {
                type: "value",
                max: maxCount > 0 ? Math.ceil(maxCount * 1.25) : 10,
                minInterval: 1,
                axisLabel: { color: getChartTextColor(), fontSize: getFSD(10) },
                splitLine: { lineStyle: { color: getChartSplitLineColor() } }
            },
            series: [{
                type: "bar",
                barWidth: 32,
                data: COVERAGE_BUCKETS.map(function (bucket) {
                    const value = counts.get(bucket.id) || 0;
                    const colors = bucketColors[bucket.id];
                    return {
                        bucketId: bucket.id,
                        value: value,
                        itemStyle: {
                            color: value > 0
                                ? new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                                    { offset: 0, color: colors[0] },
                                    { offset: 1, color: colors[1] }
                                    ])
                                : getChartEmptyBarColor(),
                            borderRadius: [6, 6, 0, 0]
                        }
                    };
                }),
                label: {
                    show: true,
                    position: "top",
                    color: getChartSubtleColor(),
                    fontSize: getFSD(11),
                    formatter: function (params) {
                        return formatNumber(params.value);
                    }
                }
            }]
        }, true);
    }

    function renderConcentrationChart(rows) {
        const data = aggregateByCounty(rows).slice(0, 5).reverse();
        if (!data.length || !data.some(function (item) { return item.principalBalance > 0; })) {
            renderChartEmptyState(charts.concentration, "暫無相關曝險數據");
            return;
        }
        const totalPrincipal = rows.reduce(function (sum, item) {
            return sum + item.principalBalance;
        }, 0);

        renderHorizontalBar(charts.concentration, data.map(function (item) {
            return item.displayName;
        }), data.map(function (item) {
            return item.principalBalance;
        }), getActiveBarGradient("countyRank"), {
            showPercentage: true,
            totalValue: totalPrincipal,
            highlightTop3: true
        });
    }
})();
