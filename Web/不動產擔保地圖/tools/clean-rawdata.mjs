import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, "..");
const DATA_DIR = path.join(ROOT_DIR, "data");
const BACKUP_DIR = path.join(ROOT_DIR, "backups");
const DEFAULT_RAWDATA_PATH = path.join(DATA_DIR, "範例.xlsx");
const DETAIL_OUTPUT_NAME = "案件明細表.csv";
const COUNTY_OUTPUT_NAME = "縣市彙總表.csv";
const UTF8_BOM = "\ufeff";
const require = createRequire(import.meta.url);
const XLSX = require(path.join(ROOT_DIR, "libs", "xlsx.full.min.js"));

const COUNTY_ALIAS = {
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

const LTV_BUCKETS = [
    { id: "0-60", min: 0, max: 0.6 },
    { id: "60-80", min: 0.6, max: 0.8 },
    { id: "80-100", min: 0.8, max: 1 },
    { id: "100+", min: 1, max: Infinity }
];

const COVERAGE_BUCKETS = [
    { id: "under-100", min: 0, max: 1 },
    { id: "100-120", min: 1, max: 1.2 },
    { id: "120-150", min: 1.2, max: 1.5 },
    { id: "150+", min: 1.5, max: Infinity }
];

const MATRIX_META = {
    "0-60": { id: "0-60", label: "0-60%", className: "matrix-weak" },
    "60-80": { id: "60-80", label: "60-80%", className: "matrix-tight" },
    "80-100": { id: "80-100", label: "80-100%", className: "matrix-tight" },
    "100+": { id: "100+", label: "100%+", className: "matrix-strong" }
};

const SUPPLY_PRESSURE_META = {
    unknown: { id: "unknown", label: "供給資料不足", className: "supply-unknown" }
};

const MERGED_CASE_REASON = "合併案件";

const state = {
    rows: [],
    cityFields: new Map(),
    displayCountyByKey: new Map()
};

const CASE_HEADERS = [
    "客戶名稱",
    "縣市",
    "用途",
    "案號",
    "本金餘額(K)",
    "已核未動撥(K)",
    "授信戶歸戶金額(K)",
    "擔保品估值(K)",
    "本金LTV(%)",
    "擔保覆蓋率(%)",
    "空餘屋供給壓力",
    "風險註記"
];

const COUNTY_HEADERS = [
    "縣市",
    "客戶數",
    "最大客戶",
    "最大客戶佔比(%)",
    "授信戶歸戶金額(K)",
    "動撥金額(K)",
    "已核未動撥(K)",
    "本金餘額(K)",
    "區域佔比(%)",
    "擔保品估值(K)"
];

function main() {
    const args = process.argv.slice(2);
    const inputPath = args[0] ? path.resolve(process.cwd(), args[0]) : DEFAULT_RAWDATA_PATH;
    const outputDir = args[1] ? path.resolve(process.cwd(), args[1]) : DATA_DIR;

    if (!fs.existsSync(inputPath)) {
        throw new Error(`找不到 rawdata 檔案：${inputPath}`);
    }

    const rawRows = readWorkbookRows(inputPath);
    state.rows = rawRows.map(cleanRow).map(scaleMoneyFieldsToK);
    detectCityFields();

    const caseRows = buildCaseDetailRows();
    const countyRows = buildCountyRows(caseRows);

    fs.mkdirSync(outputDir, { recursive: true });
    const detailPath = path.join(outputDir, DETAIL_OUTPUT_NAME);
    const countyPath = path.join(outputDir, COUNTY_OUTPUT_NAME);
    backupIfExists(detailPath);
    backupIfExists(countyPath);
    fs.writeFileSync(detailPath, UTF8_BOM + rowsToCsv(caseRows.map(toCaseCsvRow), CASE_HEADERS), "utf8");
    fs.writeFileSync(countyPath, UTF8_BOM + rowsToCsv(countyRows.map(toCountyCsvRow), COUNTY_HEADERS), "utf8");

    console.log(`rawdata: ${inputPath}`);
    console.log(`案件明細表: ${detailPath}`);
    console.log(`縣市彙總表: ${countyPath}`);
    console.log(`案件筆數: ${caseRows.length}`);
    console.log(`縣市彙總列數: ${countyRows.length}`);
}

function readWorkbookRows(inputPath) {
    const workbookBuffer = fs.readFileSync(inputPath);
    const workbook = XLSX.read(workbookBuffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
        return [];
    }
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: 0 });
}

function backupIfExists(filePath) {
    if (!fs.existsSync(filePath)) {
        return;
    }
    fs.mkdirSync(BACKUP_DIR, { recursive: true });
    const backupName = `${path.basename(filePath)}.${timestamp()}.bak`;
    fs.copyFileSync(filePath, path.join(BACKUP_DIR, backupName));
}

function timestamp() {
    const now = new Date();
    const pad = function (value) {
        return String(value).padStart(2, "0");
    };
    return [
        now.getFullYear(),
        pad(now.getMonth() + 1),
        pad(now.getDate())
    ].join("") + "_" + [
        pad(now.getHours()),
        pad(now.getMinutes()),
        pad(now.getSeconds())
    ].join("");
}

function buildCaseDetailRows() {
    const derivedRows = state.rows.flatMap(function (row, index) {
        return deriveRowsForMerge(row, index);
    });
    const mergedRows = mergeCaseRows(derivedRows).map(function (item) {
        item.industryCategory = getIndustryCategory(item);
        return item;
    });
    return sortCaseRows(mergedRows.filter(function (item) {
        return item.principalBalance > 0 || item.approvedUnused > 0;
    }));
}

function buildCountyRows(caseRows) {
    const summaries = buildCountySummaryRows(caseRows).filter(function (item) {
        return item.customerCount > 0;
    });
    const totalRemaining = summaries.reduce(function (sum, item) {
        return sum + item.remainingPrincipal;
    }, 0);
    state._countyTotalRemaining = totalRemaining;
    const sortedRows = sortCountySummaryRows(summaries, totalRemaining);
    return sortedRows.concat([buildCountySummaryTotal(sortedRows, totalRemaining)]);
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
        if (/_remain$|_worth$/u.test(key)) {
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
        if (!key.endsWith("_remain") && !key.endsWith("_worth")) {
            return;
        }
        const rawCounty = key.replace(/_(remain|worth)$/u, "");
        const countyKey = normalizeCountyName(rawCounty);
        if (!state.cityFields.has(countyKey)) {
            state.cityFields.set(countyKey, {});
        }
        const field = state.cityFields.get(countyKey);
        if (key.endsWith("_remain")) {
            field.remain = key;
        } else {
            field.worth = key;
        }
        state.displayCountyByKey.set(countyKey, displayCountyName(rawCounty));
    });
}

function deriveRowsForMerge(row, sourceIndex) {
    const countyRows = deriveCountySplitRows(row);
    if (!countyRows.length) {
        const fallback = deriveRow(row, "");
        fallback.sourceIndex = sourceIndex;
        return [fallback];
    }

    return countyRows.map(function (item) {
        item.sourceIndex = sourceIndex;
        return item;
    });
}

function deriveRow(row, county) {
    const countyKey = county || "";
    const principalBalance = countyKey ? countyExposure(row, countyKey) : rowExposure(row);
    const collateralValue = countyKey ? countyCollateral(row, countyKey) : rowCollateral(row);
    const approvedUnused = toNumber(row[FIELD.activeLimit]);
    return createDerivedItem(row, countyKey, principalBalance, collateralValue, approvedUnused);
}

function deriveCountySplitRows(row) {
    const counties = [];
    state.cityFields.forEach(function (_fields, county) {
        const principal = countyExposure(row, county);
        const collateral = countyCollateral(row, county);
        if (principal > 0 || collateral > 0) {
            counties.push({
                county: county,
                principalBalance: principal,
                collateralValue: collateral
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

        return createDerivedItem(row, entry.county, entry.principalBalance, entry.collateralValue, allocatedApprovedUnused);
    });
}

function createDerivedItem(row, county, principalBalance, collateralValue, approvedUnused) {
    const householdAmount = principalBalance + approvedUnused;
    const principalLtv = safeDivide(principalBalance, collateralValue);
    const householdLtv = safeDivide(householdAmount, collateralValue);
    const coverageRatio = safeDivide(collateralValue, principalBalance);
    const primaryCounty = county || getPrimaryCounty(row);

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
        riskLevel: getRiskLevel(principalLtv, collateralValue),
        ltvBucket: getLtvBucket(principalLtv),
        riskReasons: buildRiskReasons(principalLtv, coverageRatio, collateralValue)
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
        item.industries = caseIndustries.slice();
        item.industry = item.industries.join(" 及 ");
        item.group = item.groups[0] || item.group;
        item.principalLtv = safeDivide(item.principalBalance, item.collateralValue);
        item.householdLtv = safeDivide(item.householdAmount, item.collateralValue);
        item.coverageRatio = safeDivide(item.collateralValue, item.principalBalance);
        item.riskLevel = getRiskLevel(item.principalLtv, item.collateralValue);
        item.ltvBucket = getLtvBucket(item.principalLtv);
        item.riskReasons = buildRiskReasons(item.principalLtv, item.coverageRatio, item.collateralValue);
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
        result.set(caseKey, industries);
    });
    return result;
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

function normalizeMergeKey(value) {
    return text(value).toLowerCase();
}

function pushUnique(list, value) {
    const normalized = text(value);
    if (normalized && !list.includes(normalized)) {
        list.push(normalized);
    }
}

function getMatrixEvaluation(item) {
    const collateral = getCollateralMatrixLevel(item.coverageRatio);
    const supply = SUPPLY_PRESSURE_META.unknown;
    return {
        collateral: collateral,
        supply: supply,
        quadrantId: "",
        label: `${collateral.label} / ${supply.label}`,
        className: `${collateral.className} ${supply.className}`
    };
}

function getCollateralMatrixLevel(coverageRatio) {
    const bucket = LTV_BUCKETS.find(function (entry) {
        return coverageRatio >= entry.min && coverageRatio < entry.max;
    }) || LTV_BUCKETS[LTV_BUCKETS.length - 1];
    return MATRIX_META[bucket.id] || MATRIX_META["100+"];
}

function buildRiskReasons(ltv, coverageRatio, collateralValue) {
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
    return reasons;
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

function normalizeIndustryLabel(value) {
    return text(value).replace(/\([0-9]+\)/g, "").replace(/\s/g, "");
}

function buildCountySummaryRows(rows) {
    const byCounty = new Map();
    COMMON_COUNTY_ORDER.forEach(function (countyName) {
        byCounty.set(countyName, createCountySummary(countyName));
    });

    rows.forEach(function (item) {
        const county = displayCountyName(item.county) || "Unassigned";
        if (!byCounty.has(county)) {
            byCounty.set(county, createCountySummary(county));
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

function buildCountySummaryTotal(rows, totalRemaining) {
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
        acc.totalRemaining = totalRemaining;
        return acc;
    }, createCountySummary("Total"));
}

function createCountySummary(county) {
    return {
        county: county,
        customerCount: 0,
        maxCustomerName: "",
        maxCustomerRemaining: 0,
        approvedLimitTotal: 0,
        disbursedTotal: 0,
        approvedUnusedTotal: 0,
        remainingPrincipal: 0,
        collateralTotal: 0,
        totalRemaining: 0
    };
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

function getCoverageBucketByRatio(coverageRatio) {
    return COVERAGE_BUCKETS.find(function (entry) {
        return coverageRatio >= entry.min && coverageRatio < entry.max;
    }) || COVERAGE_BUCKETS[0];
}

function sortCaseRows(rows) {
    return rows.slice().sort(function (a, b) {
        return compareSortValue(getCaseSortValue(a, 8), getCaseSortValue(b, 8), "desc");
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
    return rows.slice().sort(function (a, b) {
        return compareSortValue(getCountySortValue(a, 0, totalRemaining), getCountySortValue(b, 0, totalRemaining), "asc");
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

function toCaseCsvRow(item) {
    return {
        "客戶名稱": item.applicant,
        "縣市": displayCountyName(item.county),
        "用途": item.industry,
        "案號": item.caseId || "-",
        "本金餘額(K)": roundMoney(item.principalBalance),
        "已核未動撥(K)": roundMoney(item.approvedUnused),
        "授信戶歸戶金額(K)": roundMoney(item.householdAmount),
        "擔保品估值(K)": roundMoney(item.collateralValue),
        "本金LTV(%)": percentNumber(item.principalLtv),
        "擔保覆蓋率(%)": percentNumber(item.coverageRatio),
        "空餘屋供給壓力": item.matrix && item.matrix.supply ? item.matrix.supply.label : "供給資料不足",
        "風險註記": getCaseNotes(item).join("、") || "-"
    };
}

function toCountyCsvRow(item) {
    const totalRemaining = item.county === "Total" ? item.remainingPrincipal : getCountyTotalRemaining();
    return {
        "縣市": item.county,
        "客戶數": item.customerCount,
        "最大客戶": item.maxCustomerName || "-",
        "最大客戶佔比(%)": percentNumber(safeDivide(item.maxCustomerRemaining, item.remainingPrincipal)),
        "授信戶歸戶金額(K)": roundMoney(item.approvedLimitTotal),
        "動撥金額(K)": roundMoney(item.disbursedTotal),
        "已核未動撥(K)": roundMoney(item.approvedUnusedTotal),
        "本金餘額(K)": roundMoney(item.remainingPrincipal),
        "區域佔比(%)": percentNumber(safeDivide(item.remainingPrincipal, totalRemaining)),
        "擔保品估值(K)": roundMoney(item.collateralTotal)
    };
}

function getCountyTotalRemaining() {
    return state._countyTotalRemaining || 0;
}

function rowsToCsv(rows, headers) {
    return [
        headers.map(csvCell).join(","),
        ...rows.map(function (row) {
            return headers.map(function (header) {
                return csvCell(row[header]);
            }).join(",");
        })
    ].join("\r\n") + "\r\n";
}

function csvCell(value) {
    const raw = value === null || typeof value === "undefined" ? "" : String(value);
    const clean = raw.replace(/\ufeff|\u200b|\xa0/g, "").trim();
    const safe = /^[=+@]/u.test(clean) ? `'${clean}` : clean;
    if (/[",\r\n]/u.test(safe)) {
        return `"${safe.replace(/"/g, "\"\"")}"`;
    }
    return safe;
}

function getCaseNotes(item) {
    return item.riskReasons.filter(function (reason) {
        return reason === MERGED_CASE_REASON;
    });
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

function safeDivide(numerator, denominator) {
    if (!denominator) {
        return 0;
    }
    return numerator / denominator;
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

function roundMoney(value) {
    return Math.round(toNumber(value));
}

function percentNumber(value) {
    return Number((toNumber(value) * 100).toFixed(2));
}

function localeCompare(a, b) {
    return String(a).localeCompare(String(b), "zh-Hant");
}

try {
    main();
} catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
}
