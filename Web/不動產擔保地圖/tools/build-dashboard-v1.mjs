import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, "..");
const CASE_CSV = path.join(ROOT_DIR, "data", "案件明細表V1.csv");
const COUNTY_CSV = path.join(ROOT_DIR, "data", "縣市彙總表V1.csv");
const OUTPUT_HTML = path.join(ROOT_DIR, "dashboard-v1.html");

const cases = readCsv(CASE_CSV);
const counties = readCsv(COUNTY_CSV);

fs.writeFileSync(OUTPUT_HTML, buildHtml(cases, counties), "utf8");
console.log(`Dashboard created: ${OUTPUT_HTML}`);
console.log(`Cases: ${cases.length}`);
console.log(`Counties: ${counties.length}`);

function readCsv(filePath) {
    const text = fs.readFileSync(filePath, "utf8").replace(/^\ufeff/, "");
    const rows = parseCsv(text);
    const headers = rows.shift() || [];
    return rows.filter(function (row) {
        return row.some(function (cell) {
            return String(cell || "").trim();
        });
    }).map(function (row) {
        const item = {};
        headers.forEach(function (header, index) {
            item[header] = row[index] ?? "";
        });
        return item;
    });
}

function parseCsv(text) {
    const rows = [];
    let row = [];
    let cell = "";
    let quote = false;
    for (let index = 0; index < text.length; index += 1) {
        const char = text[index];
        const next = text[index + 1];
        if (quote) {
            if (char === "\"" && next === "\"") {
                cell += "\"";
                index += 1;
            } else if (char === "\"") {
                quote = false;
            } else {
                cell += char;
            }
            continue;
        }
        if (char === "\"") {
            quote = true;
        } else if (char === ",") {
            row.push(cell);
            cell = "";
        } else if (char === "\n") {
            row.push(cell);
            rows.push(row);
            row = [];
            cell = "";
        } else if (char !== "\r") {
            cell += char;
        }
    }
    if (cell || row.length) {
        row.push(cell);
        rows.push(row);
    }
    return rows;
}

function buildHtml(caseRows, countyRows) {
    return `<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>擔保曝險 Dashboard V1</title>
    <script src="libs/echarts.min.js"></script>
    <style>
        :root {
            --ink: #152420;
            --muted: #62726d;
            --paper: #f7f1e4;
            --paper-strong: #fffaf0;
            --panel: rgba(255, 250, 240, 0.86);
            --line: rgba(21, 36, 32, 0.14);
            --forest: #164b3f;
            --moss: #6e7f48;
            --gold: #d79a25;
            --clay: #b95f35;
            --red: #a6332b;
            --blue: #276c8f;
            --shadow: 0 24px 70px rgba(61, 50, 30, 0.18);
        }

        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            color: var(--ink);
            font-family: "Noto Serif TC", "Songti TC", "Microsoft JhengHei", serif;
            background:
                radial-gradient(circle at 12% 10%, rgba(215, 154, 37, 0.24), transparent 26rem),
                radial-gradient(circle at 82% 8%, rgba(39, 108, 143, 0.18), transparent 24rem),
                linear-gradient(135deg, #f8efe0 0%, #e9dfc8 48%, #f6ead7 100%);
            min-height: 100vh;
        }

        body::before {
            content: "";
            position: fixed;
            inset: 0;
            pointer-events: none;
            opacity: 0.12;
            background-image:
                linear-gradient(90deg, rgba(21, 36, 32, 0.08) 1px, transparent 1px),
                linear-gradient(rgba(21, 36, 32, 0.08) 1px, transparent 1px);
            background-size: 36px 36px;
            mask-image: linear-gradient(to bottom, #000, transparent 90%);
        }

        .shell {
            position: relative;
            width: min(1540px, calc(100% - 32px));
            margin: 0 auto;
            padding: 28px 0 42px;
        }

        .hero {
            display: grid;
            grid-template-columns: minmax(280px, 1fr) auto;
            gap: 24px;
            align-items: end;
            margin-bottom: 18px;
            animation: rise 0.55s ease both;
        }

        .eyebrow {
            display: inline-flex;
            gap: 10px;
            align-items: center;
            padding: 8px 12px;
            border: 1px solid rgba(22, 75, 63, 0.22);
            border-radius: 999px;
            color: var(--forest);
            background: rgba(255, 250, 240, 0.68);
            font-size: 13px;
            letter-spacing: 0.08em;
        }

        h1 {
            margin: 14px 0 8px;
            font-size: clamp(34px, 5vw, 76px);
            line-height: 0.94;
            letter-spacing: -0.08em;
            font-weight: 900;
        }

        .hero p {
            max-width: 820px;
            margin: 0;
            color: var(--muted);
            font-size: 16px;
            line-height: 1.75;
        }

        .source-card {
            min-width: 280px;
            padding: 18px;
            border: 1px solid var(--line);
            border-radius: 24px;
            background: rgba(255, 250, 240, 0.72);
            box-shadow: var(--shadow);
        }

        .source-card strong {
            display: block;
            font-size: 28px;
            line-height: 1;
            color: var(--forest);
        }

        .source-card span {
            display: block;
            margin-top: 8px;
            color: var(--muted);
            font-size: 13px;
        }

        .filter-bar {
            position: sticky;
            top: 10px;
            z-index: 10;
            display: grid;
            grid-template-columns: repeat(5, minmax(160px, 1fr));
            gap: 12px;
            margin: 20px 0;
            padding: 14px;
            border: 1px solid var(--line);
            border-radius: 28px;
            background: rgba(255, 250, 240, 0.82);
            box-shadow: 0 18px 44px rgba(61, 50, 30, 0.14);
            backdrop-filter: blur(16px);
        }

        label {
            display: grid;
            gap: 6px;
            color: var(--muted);
            font-size: 12px;
            letter-spacing: 0.08em;
        }

        select,
        button {
            font: inherit;
        }

        select {
            width: 100%;
            min-height: 42px;
            border: 1px solid rgba(21, 36, 32, 0.18);
            border-radius: 16px;
            padding: 0 12px;
            color: var(--ink);
            background: #fffaf0;
            outline: none;
        }

        .reset-button {
            align-self: end;
            min-height: 42px;
            border: 0;
            border-radius: 16px;
            color: #fffaf0;
            background: linear-gradient(135deg, var(--forest), #0d2f28);
            cursor: pointer;
            box-shadow: 0 14px 28px rgba(22, 75, 63, 0.22);
        }

        .grid {
            display: grid;
            grid-template-columns: 1.1fr 1.35fr;
            gap: 18px;
        }

        .panel {
            border: 1px solid var(--line);
            border-radius: 30px;
            background: var(--panel);
            box-shadow: var(--shadow);
            overflow: hidden;
            animation: rise 0.55s ease both;
        }

        .panel-header {
            display: flex;
            justify-content: space-between;
            gap: 16px;
            padding: 18px 20px 12px;
            border-bottom: 1px solid var(--line);
        }

        .panel-title {
            margin: 0;
            font-size: 18px;
            font-weight: 900;
            letter-spacing: -0.04em;
        }

        .panel-subtitle {
            margin: 4px 0 0;
            color: var(--muted);
            font-size: 12px;
        }

        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 12px;
            margin-bottom: 18px;
        }

        .kpi {
            min-height: 126px;
            padding: 18px;
            border: 1px solid var(--line);
            border-radius: 26px;
            background: rgba(255, 250, 240, 0.76);
            box-shadow: 0 16px 38px rgba(61, 50, 30, 0.12);
            animation: rise 0.55s ease both;
        }

        .kpi:nth-child(2) { animation-delay: 0.05s; }
        .kpi:nth-child(3) { animation-delay: 0.10s; }
        .kpi:nth-child(4) { animation-delay: 0.15s; }

        .kpi span {
            display: block;
            color: var(--muted);
            font-size: 12px;
            letter-spacing: 0.12em;
        }

        .kpi strong {
            display: block;
            margin-top: 16px;
            color: var(--forest);
            font-size: clamp(25px, 3vw, 42px);
            line-height: 1;
            letter-spacing: -0.08em;
        }

        .kpi em {
            display: block;
            margin-top: 9px;
            color: var(--clay);
            font-size: 12px;
            font-style: normal;
        }

        .map-panel {
            min-height: 760px;
        }

        .taiwan-map {
            position: relative;
            min-height: 680px;
            padding: 20px;
        }

        .map-island {
            position: absolute;
            inset: 42px 120px 48px 118px;
            border-radius: 48% 52% 42% 58% / 16% 18% 82% 84%;
            background: linear-gradient(180deg, rgba(22, 75, 63, 0.08), rgba(215, 154, 37, 0.10));
            transform: rotate(-8deg);
            border: 1px solid rgba(22, 75, 63, 0.13);
        }

        .county-chip {
            position: absolute;
            min-width: 96px;
            border: 1px solid rgba(21, 36, 32, 0.16);
            border-radius: 18px;
            padding: 9px 10px;
            color: var(--ink);
            background: rgba(255, 250, 240, 0.92);
            cursor: pointer;
            box-shadow: 0 10px 26px rgba(61, 50, 30, 0.12);
            transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
        }

        .county-chip:hover,
        .county-chip.active {
            transform: translateY(-3px) scale(1.03);
            border-color: var(--forest);
            box-shadow: 0 18px 40px rgba(22, 75, 63, 0.20);
        }

        .county-chip b,
        .county-chip span {
            display: block;
        }

        .county-chip b {
            font-size: 14px;
            letter-spacing: -0.03em;
        }

        .county-chip span {
            margin-top: 3px;
            color: var(--muted);
            font-size: 11px;
        }

        .county-chip i {
            position: absolute;
            right: 8px;
            top: 8px;
            width: var(--dot);
            height: var(--dot);
            min-width: 8px;
            min-height: 8px;
            max-width: 26px;
            max-height: 26px;
            border-radius: 999px;
            background: var(--tone);
            box-shadow: 0 0 0 6px color-mix(in srgb, var(--tone), transparent 78%);
        }

        .right-stack {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 18px;
        }

        .chart {
            height: 320px;
        }

        .wide {
            grid-column: 1 / -1;
        }

        .tables {
            display: grid;
            grid-template-columns: 0.9fr 1.3fr;
            gap: 18px;
            margin-top: 18px;
        }

        .table-wrap {
            max-height: 430px;
            overflow: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }

        th {
            position: sticky;
            top: 0;
            z-index: 1;
            background: #efe3cd;
            color: #41504a;
            text-align: left;
            white-space: nowrap;
        }

        th,
        td {
            padding: 10px 12px;
            border-bottom: 1px solid rgba(21, 36, 32, 0.10);
        }

        td.numeric {
            text-align: right;
            font-variant-numeric: tabular-nums;
        }

        .pill {
            display: inline-flex;
            align-items: center;
            min-height: 24px;
            border-radius: 999px;
            padding: 2px 9px;
            color: #fffaf0;
            background: var(--forest);
            font-size: 12px;
            white-space: nowrap;
        }

        .empty {
            padding: 20px;
            color: var(--muted);
        }

        @keyframes rise {
            from {
                opacity: 0;
                transform: translateY(16px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        @media (max-width: 1180px) {
            .hero,
            .grid,
            .tables {
                grid-template-columns: 1fr;
            }

            .filter-bar,
            .kpi-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
        }

        @media (max-width: 760px) {
            .shell {
                width: min(100% - 20px, 1540px);
                padding-top: 18px;
            }

            .filter-bar,
            .kpi-grid,
            .right-stack {
                grid-template-columns: 1fr;
            }

            .taiwan-map {
                min-height: 900px;
            }
        }
    </style>
</head>
<body>
    <div class="shell">
        <header class="hero">
            <div>
                <span class="eyebrow">FMG Collateral Atlas · CSV V1</span>
                <h1>擔保曝險<br>風險儀表板</h1>
                <p>整合縣市彙總表與案件明細表，提供地圖、象限、LTV、擔保覆蓋率與縣市條件篩選。點擊地圖縣市或調整篩選器，即時重算 KPI、圖表與明細表。</p>
            </div>
            <div class="source-card">
                <strong id="heroCaseCount">0</strong>
                <span>篩選後案件數 / 資料來源：案件明細表V1、縣市彙總表V1</span>
            </div>
        </header>

        <section class="filter-bar" aria-label="dashboard filters">
            <label>縣市
                <select id="countyFilter"></select>
            </label>
            <label>象限
                <select id="quadrantFilter"></select>
            </label>
            <label>LTV
                <select id="ltvFilter">
                    <option value="">全部</option>
                    <option value="0-60">0-60%</option>
                    <option value="60-80">60-80%</option>
                    <option value="80-100">80-100%</option>
                    <option value="100+">100%+</option>
                </select>
            </label>
            <label>擔保覆蓋率
                <select id="coverageFilter">
                    <option value="">全部</option>
                    <option value="under-100">100%以下</option>
                    <option value="100-120">100-120%</option>
                    <option value="120-150">120-150%</option>
                    <option value="150+">150%以上</option>
                </select>
            </label>
            <button id="resetFilters" class="reset-button" type="button">清除篩選</button>
        </section>

        <section class="kpi-grid">
            <article class="kpi"><span>本金餘額</span><strong id="kpiPrincipal">0</strong><em>篩選後案件加總</em></article>
            <article class="kpi"><span>授信戶歸戶金額</span><strong id="kpiHousehold">0</strong><em>本金 + 已核未動撥</em></article>
            <article class="kpi"><span>平均 LTV</span><strong id="kpiLtv">0%</strong><em>本金餘額 / 擔保品估值</em></article>
            <article class="kpi"><span>擔保覆蓋率</span><strong id="kpiCoverage">0%</strong><em>擔保品估值 / 本金餘額</em></article>
        </section>

        <main class="grid">
            <section class="panel map-panel">
                <div class="panel-header">
                    <div>
                        <h2 class="panel-title">地圖曝險篩選</h2>
                        <p class="panel-subtitle">縣市泡泡大小代表本金餘額，顏色代表曝險強度。</p>
                    </div>
                    <span class="pill" id="activeCounty">全台</span>
                </div>
                <div class="taiwan-map" id="countyMap">
                    <div class="map-island"></div>
                </div>
            </section>

            <section class="right-stack">
                <article class="panel">
                    <div class="panel-header">
                        <div>
                            <h2 class="panel-title">象限分布</h2>
                            <p class="panel-subtitle">以案件數觀察供給壓力區位。</p>
                        </div>
                    </div>
                    <div class="chart" id="quadrantChart"></div>
                </article>
                <article class="panel">
                    <div class="panel-header">
                        <div>
                            <h2 class="panel-title">LTV 分布</h2>
                            <p class="panel-subtitle">本金 LTV 梯度。</p>
                        </div>
                    </div>
                    <div class="chart" id="ltvChart"></div>
                </article>
                <article class="panel wide">
                    <div class="panel-header">
                        <div>
                            <h2 class="panel-title">擔保覆蓋率與縣市曝險</h2>
                            <p class="panel-subtitle">左：覆蓋率區間；右：本金餘額 Top 10。</p>
                        </div>
                    </div>
                    <div class="chart" id="coverageChart"></div>
                </article>
                <article class="panel wide">
                    <div class="panel-header">
                        <div>
                            <h2 class="panel-title">案件散點矩陣</h2>
                            <p class="panel-subtitle">X 軸 LTV，Y 軸擔保覆蓋率，點大小代表本金餘額。</p>
                        </div>
                    </div>
                    <div class="chart" id="scatterChart"></div>
                </article>
            </section>
        </main>

        <section class="tables">
            <article class="panel">
                <div class="panel-header">
                    <div>
                        <h2 class="panel-title">縣市彙總表</h2>
                        <p class="panel-subtitle" id="countySummaryNote">依目前條件重算。</p>
                    </div>
                </div>
                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>縣市</th>
                                <th>案件數</th>
                                <th>本金餘額</th>
                                <th>區域佔比</th>
                            </tr>
                        </thead>
                        <tbody id="countyRows"></tbody>
                    </table>
                </div>
            </article>
            <article class="panel">
                <div class="panel-header">
                    <div>
                        <h2 class="panel-title">案件明細表</h2>
                        <p class="panel-subtitle" id="caseSummaryNote">顯示篩選後前 80 筆。</p>
                    </div>
                </div>
                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>客戶</th>
                                <th>縣市</th>
                                <th>用途</th>
                                <th>案號</th>
                                <th>本金</th>
                                <th>LTV</th>
                                <th>覆蓋率</th>
                                <th>象限</th>
                            </tr>
                        </thead>
                        <tbody id="caseRows"></tbody>
                    </table>
                </div>
            </article>
        </section>
    </div>

    <script>
        const CASE_ROWS = ${JSON.stringify(caseRows)};
        const COUNTY_ROWS = ${JSON.stringify(countyRows)};
    </script>
    <script>
        const countyPositions = {
            "基隆市": [70, 11], "台北市": [58, 14], "新北市": [68, 19], "桃園市": [49, 23],
            "新竹縣": [44, 30], "新竹市": [32, 31], "苗栗縣": [41, 38], "台中市": [38, 47],
            "彰化縣": [28, 54], "南投縣": [48, 56], "雲林縣": [30, 63], "嘉義縣": [36, 70],
            "嘉義市": [23, 70], "台南市": [34, 78], "高雄市": [44, 85], "屏東縣": [54, 92],
            "宜蘭縣": [77, 31], "花蓮縣": [76, 55], "台東縣": [69, 78], "澎湖縣": [10, 68],
            "金門縣": [8, 34], "連江縣": [19, 10], "未分縣市": [15, 84], "全台": [50, 8]
        };

        const colors = {
            forest: "#164b3f",
            gold: "#d79a25",
            clay: "#b95f35",
            red: "#a6332b",
            blue: "#276c8f",
            moss: "#6e7f48",
            ink: "#152420"
        };

        const state = {
            county: "",
            quadrant: "",
            ltv: "",
            coverage: ""
        };

        const charts = {
            quadrant: echarts.init(document.getElementById("quadrantChart")),
            ltv: echarts.init(document.getElementById("ltvChart")),
            coverage: echarts.init(document.getElementById("coverageChart")),
            scatter: echarts.init(document.getElementById("scatterChart"))
        };

        const cases = CASE_ROWS.map(function (item) {
            return {
                customer: item["客戶名稱"],
                county: cleanCounty(item["縣市"]),
                industry: item["用途"],
                caseId: item["案號"],
                principal: number(item["本金餘額(K)"]),
                unused: number(item["已核未動撥(K)"]),
                household: number(item["授信戶歸戶金額(K)"]),
                collateral: number(item["擔保品估值(K)"]),
                ltv: number(item["本金LTV(%)"]),
                coverage: number(item["擔保覆蓋率(%)"]),
                quadrant: item["空餘屋供給壓力"] || "未分類",
                note: item["風險註記"] || "-"
            };
        });

        const countyBase = COUNTY_ROWS.filter(function (item) {
            return item["縣市"] !== "Total";
        }).map(function (item) {
            return {
                county: cleanCounty(item["縣市"]),
                customerCount: number(item["客戶數"]),
                principal: number(item["本金餘額(K)"]),
                share: number(item["區域佔比(%)"]),
                collateral: number(item["擔保品估值(K)"])
            };
        });

        initFilters();
        render();
        window.addEventListener("resize", function () {
            Object.values(charts).forEach(function (chart) {
                chart.resize();
            });
        });

        function initFilters() {
            fillSelect("countyFilter", ["", ...unique(cases.map(function (item) { return item.county; })).sort(localeSort)], "全台");
            fillSelect("quadrantFilter", ["", ...unique(cases.map(function (item) { return item.quadrant; })).sort(localeSort)], "全部");
            ["countyFilter", "quadrantFilter", "ltvFilter", "coverageFilter"].forEach(function (id) {
                document.getElementById(id).addEventListener("change", function (event) {
                    const key = id.replace("Filter", "");
                    state[key] = event.target.value;
                    render();
                });
            });
            document.getElementById("resetFilters").addEventListener("click", function () {
                state.county = "";
                state.quadrant = "";
                state.ltv = "";
                state.coverage = "";
                syncFilters();
                render();
            });
        }

        function fillSelect(id, values, emptyLabel) {
            const select = document.getElementById(id);
            select.innerHTML = values.map(function (value) {
                const label = value || emptyLabel;
                return "<option value=\\"" + escapeHtml(value) + "\\">" + escapeHtml(label) + "</option>";
            }).join("");
        }

        function syncFilters() {
            document.getElementById("countyFilter").value = state.county;
            document.getElementById("quadrantFilter").value = state.quadrant;
            document.getElementById("ltvFilter").value = state.ltv;
            document.getElementById("coverageFilter").value = state.coverage;
        }

        function render() {
            syncFilters();
            const filtered = getFilteredCases();
            const countyRows = aggregateCounties(filtered);
            renderKpis(filtered);
            renderMap(countyRows);
            renderQuadrant(filtered);
            renderLtv(filtered);
            renderCoverageAndCounty(countyRows, filtered);
            renderScatter(filtered);
            renderTables(countyRows, filtered);
        }

        function getFilteredCases() {
            return cases.filter(function (item) {
                if (state.county && item.county !== state.county) {
                    return false;
                }
                if (state.quadrant && item.quadrant !== state.quadrant) {
                    return false;
                }
                if (state.ltv && bucketLtv(item.ltv) !== state.ltv) {
                    return false;
                }
                if (state.coverage && bucketCoverage(item.coverage) !== state.coverage) {
                    return false;
                }
                return true;
            });
        }

        function aggregateCounties(rows) {
            const map = new Map();
            rows.forEach(function (item) {
                if (!map.has(item.county)) {
                    map.set(item.county, { county: item.county, cases: 0, principal: 0, household: 0, collateral: 0 });
                }
                const current = map.get(item.county);
                current.cases += 1;
                current.principal += item.principal;
                current.household += item.household;
                current.collateral += item.collateral;
            });
            const total = Array.from(map.values()).reduce(function (sum, item) {
                return sum + item.principal;
            }, 0);
            return Array.from(map.values()).map(function (item) {
                item.share = total ? item.principal / total * 100 : 0;
                return item;
            }).sort(function (a, b) {
                return b.principal - a.principal;
            });
        }

        function renderKpis(rows) {
            const principal = sum(rows, "principal");
            const household = sum(rows, "household");
            const collateral = sum(rows, "collateral");
            document.getElementById("heroCaseCount").textContent = rows.length.toLocaleString("zh-TW");
            document.getElementById("kpiPrincipal").textContent = money(principal);
            document.getElementById("kpiHousehold").textContent = money(household);
            document.getElementById("kpiLtv").textContent = percent(principal && collateral ? principal / collateral * 100 : 0);
            document.getElementById("kpiCoverage").textContent = percent(principal ? collateral / principal * 100 : 0);
            document.getElementById("activeCounty").textContent = state.county || "全台";
        }

        function renderMap(rows) {
            const wrap = document.getElementById("countyMap");
            wrap.querySelectorAll(".county-chip").forEach(function (node) {
                node.remove();
            });
            const max = Math.max(1, ...countyBase.map(function (item) { return item.principal; }));
            const currentByCounty = new Map(rows.map(function (item) {
                return [item.county, item];
            }));
            countyBase.forEach(function (base) {
                const current = currentByCounty.get(base.county) || { county: base.county, cases: 0, principal: 0 };
                const pos = countyPositions[base.county] || [50, 50];
                const intensity = base.principal / max;
                const chip = document.createElement("button");
                chip.type = "button";
                chip.className = "county-chip" + (state.county === base.county ? " active" : "");
                chip.style.left = pos[0] + "%";
                chip.style.top = pos[1] + "%";
                chip.style.setProperty("--dot", (8 + intensity * 28) + "px");
                chip.style.setProperty("--tone", toneByShare(intensity));
                chip.innerHTML = "<i></i><b>" + escapeHtml(base.county) + "</b><span>" + money(current.principal) + " · " + current.cases + "件</span>";
                chip.addEventListener("click", function () {
                    state.county = state.county === base.county ? "" : base.county;
                    render();
                });
                wrap.appendChild(chip);
            });
        }

        function renderQuadrant(rows) {
            const data = groupCount(rows, "quadrant").map(function (item) {
                return { name: item.name, value: item.value };
            });
            charts.quadrant.setOption({
                color: [colors.red, colors.gold, colors.blue, colors.moss, colors.forest],
                tooltip: { trigger: "item" },
                legend: { bottom: 0, textStyle: { color: colors.ink } },
                series: [{
                    type: "pie",
                    radius: ["42%", "70%"],
                    center: ["50%", "43%"],
                    itemStyle: { borderRadius: 10, borderColor: "#fffaf0", borderWidth: 2 },
                    label: { formatter: "{b}\\n{c}件", color: colors.ink },
                    data: data
                }]
            });
        }

        function renderLtv(rows) {
            const labels = ["0-60", "60-80", "80-100", "100+"];
            const data = labels.map(function (label) {
                return rows.filter(function (item) { return bucketLtv(item.ltv) === label; }).length;
            });
            charts.ltv.setOption(barOption(labels.map(labelLtv), data, [colors.forest, colors.blue]));
        }

        function renderCoverageAndCounty(countyRows, rows) {
            const coverageLabels = ["under-100", "100-120", "120-150", "150+"];
            const coverageData = coverageLabels.map(function (label) {
                return rows.filter(function (item) { return bucketCoverage(item.coverage) === label; }).length;
            });
            const top = countyRows.slice(0, 10).reverse();
            charts.coverage.setOption({
                color: [colors.clay, colors.forest],
                tooltip: { trigger: "axis" },
                grid: [
                    { left: 42, right: "56%", top: 36, bottom: 42 },
                    { left: "54%", right: 22, top: 36, bottom: 42 }
                ],
                xAxis: [
                    { type: "category", data: coverageLabels.map(labelCoverage), axisLabel: { color: colors.muted } },
                    { type: "value", gridIndex: 1, axisLabel: { color: colors.muted } }
                ],
                yAxis: [
                    { type: "value", axisLabel: { color: colors.muted } },
                    { type: "category", gridIndex: 1, data: top.map(function (item) { return item.county; }), axisLabel: { color: colors.muted } }
                ],
                series: [
                    { type: "bar", data: coverageData, barWidth: 28, itemStyle: { borderRadius: [10, 10, 0, 0] } },
                    { type: "bar", xAxisIndex: 1, yAxisIndex: 1, data: top.map(function (item) { return item.principal; }), itemStyle: { borderRadius: [0, 10, 10, 0] } }
                ]
            });
        }

        function renderScatter(rows) {
            charts.scatter.setOption({
                color: [colors.forest],
                tooltip: {
                    formatter: function (params) {
                        const item = params.data[3];
                        return escapeHtml(item.customer) + "<br>" + escapeHtml(item.county) + " · " + escapeHtml(item.caseId) + "<br>LTV " + item.ltv.toFixed(2) + "% / 覆蓋率 " + item.coverage.toFixed(2) + "%";
                    }
                },
                grid: { left: 62, right: 24, top: 34, bottom: 52 },
                xAxis: { name: "LTV %", axisLabel: { color: colors.muted }, splitLine: { lineStyle: { color: "rgba(21,36,32,.08)" } } },
                yAxis: { name: "覆蓋率 %", axisLabel: { color: colors.muted }, splitLine: { lineStyle: { color: "rgba(21,36,32,.08)" } } },
                series: [{
                    type: "scatter",
                    symbolSize: function (data) {
                        return Math.max(8, Math.min(38, Math.sqrt(data[2]) / 10));
                    },
                    data: rows.map(function (item) {
                        return [item.ltv, item.coverage, item.principal, item];
                    }),
                    itemStyle: { opacity: 0.76 }
                }]
            });
        }

        function renderTables(countyRows, rows) {
            const total = sum(countyRows, "principal");
            document.getElementById("countyRows").innerHTML = countyRows.map(function (item) {
                return "<tr><td>" + escapeHtml(item.county) + "</td><td class=\\"numeric\\">" + item.cases + "</td><td class=\\"numeric\\">" + money(item.principal) + "</td><td class=\\"numeric\\">" + percent(total ? item.principal / total * 100 : 0) + "</td></tr>";
            }).join("") || "<tr><td class=\\"empty\\" colspan=\\"4\\">沒有符合條件的縣市資料</td></tr>";
            document.getElementById("caseRows").innerHTML = rows.slice().sort(function (a, b) {
                return b.principal - a.principal;
            }).slice(0, 80).map(function (item) {
                return "<tr><td>" + escapeHtml(item.customer) + "</td><td>" + escapeHtml(item.county) + "</td><td>" + escapeHtml(item.industry) + "</td><td>" + escapeHtml(item.caseId) + "</td><td class=\\"numeric\\">" + money(item.principal) + "</td><td class=\\"numeric\\">" + percent(item.ltv) + "</td><td class=\\"numeric\\">" + percent(item.coverage) + "</td><td><span class=\\"pill\\">" + escapeHtml(item.quadrant) + "</span></td></tr>";
            }).join("") || "<tr><td class=\\"empty\\" colspan=\\"8\\">沒有符合條件的案件資料</td></tr>";
            document.getElementById("countySummaryNote").textContent = "目前 " + countyRows.length + " 個縣市，合計 " + money(total) + "。";
            document.getElementById("caseSummaryNote").textContent = "符合條件 " + rows.length + " 筆，顯示前 80 筆。";
        }

        function barOption(labels, data, colorPair) {
            return {
                color: colorPair,
                tooltip: { trigger: "axis" },
                grid: { left: 42, right: 22, top: 34, bottom: 50 },
                xAxis: { type: "category", data: labels, axisLabel: { color: colors.muted } },
                yAxis: { type: "value", axisLabel: { color: colors.muted }, splitLine: { lineStyle: { color: "rgba(21,36,32,.08)" } } },
                series: [{ type: "bar", data: data, barWidth: 34, itemStyle: { borderRadius: [12, 12, 0, 0] } }]
            };
        }

        function bucketLtv(value) {
            if (value < 60) return "0-60";
            if (value < 80) return "60-80";
            if (value < 100) return "80-100";
            return "100+";
        }

        function cleanCounty(value) {
            return value === "city_none" ? "未分縣市" : value;
        }

        function bucketCoverage(value) {
            if (value < 100) return "under-100";
            if (value < 120) return "100-120";
            if (value < 150) return "120-150";
            return "150+";
        }

        function labelLtv(value) {
            return ({ "0-60": "0-60%", "60-80": "60-80%", "80-100": "80-100%", "100+": "100%+" })[value] || value;
        }

        function labelCoverage(value) {
            return ({ "under-100": "100%以下", "100-120": "100-120%", "120-150": "120-150%", "150+": "150%以上" })[value] || value;
        }

        function toneByShare(value) {
            if (value >= 0.75) return colors.red;
            if (value >= 0.45) return colors.clay;
            if (value >= 0.22) return colors.gold;
            return colors.forest;
        }

        function groupCount(rows, key) {
            const map = new Map();
            rows.forEach(function (item) {
                map.set(item[key], (map.get(item[key]) || 0) + 1);
            });
            return Array.from(map.entries()).map(function (entry) {
                return { name: entry[0], value: entry[1] };
            }).sort(function (a, b) {
                return b.value - a.value;
            });
        }

        function unique(values) {
            return Array.from(new Set(values.filter(Boolean)));
        }

        function sum(rows, key) {
            return rows.reduce(function (total, item) {
                return total + number(item[key]);
            }, 0);
        }

        function number(value) {
            const parsed = Number(String(value || "0").replace(/,/g, ""));
            return Number.isFinite(parsed) ? parsed : 0;
        }

        function money(value) {
            const n = number(value);
            if (Math.abs(n) >= 100000) {
                return (n / 100000).toFixed(1).replace(/\\.0$/, "") + "億";
            }
            if (Math.abs(n) >= 10000) {
                return (n / 10000).toFixed(1).replace(/\\.0$/, "") + "千萬";
            }
            return Math.round(n).toLocaleString("zh-TW") + "K";
        }

        function percent(value) {
            return number(value).toFixed(2).replace(/\\.00$/, "") + "%";
        }

        function localeSort(a, b) {
            return String(a).localeCompare(String(b), "zh-Hant");
        }

        function escapeHtml(value) {
            return String(value ?? "").replace(/[&<>"']/g, function (char) {
                return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\\"": "&quot;", "'": "&#039;" }[char];
            });
        }
    </script>
</body>
</html>`;
}
