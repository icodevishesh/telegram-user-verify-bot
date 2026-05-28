"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.processOldOnly = processOldOnly;
exports.processNonDepositedReport = processNonDepositedReport;
exports.processInactiveUsersReport = processInactiveUsersReport;
exports.processDuplicates = processDuplicates;
exports.processVerification = processVerification;
const XLSX = __importStar(require("xlsx"));
const path = __importStar(require("path"));
const fs = __importStar(require("fs"));
// ── Sheet readers ─────────────────────────────────────────────────────────────
function readStandardSheet(filePath) {
    const wb = XLSX.readFile(filePath, { raw: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
}
/**
 * Xilion CSVs have a "sep=," sentinel on row 0.
 * Real headers are on row 1, data starts at row 2.
 */
function readXilionSheet(filePath) {
    const wb = XLSX.readFile(filePath, { raw: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const headerRowIdx = raw.findIndex((row) => String(row[0]).trim().toLowerCase() === "wallet");
    if (headerRowIdx === -1) {
        const headers = raw[1];
        return raw.slice(2).map((row) => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = String(row[i] ?? ""); });
            return obj;
        });
    }
    const headers = raw[headerRowIdx];
    return raw.slice(headerRowIdx + 1).map((row) => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = String(row[i] ?? ""); });
        return obj;
    });
}
// ── ID normalisation ──────────────────────────────────────────────────────────
function norm(val) {
    if (val === null || val === undefined)
        return "";
    return String(val).trim().replace(/^#/, "").toLowerCase();
}
// ── Build broker ID sets ──────────────────────────────────────────────────────
function buildXmIds(rows) {
    const ids = new Set();
    for (const row of rows) {
        const mt = norm(row["MT4/MT5 ID"]);
        const client = norm(row["Client ID"]);
        if (mt)
            ids.add(mt);
        if (client)
            ids.add(client);
    }
    return ids;
}
function buildXilionIds(rows) {
    const ids = new Set();
    for (const row of rows) {
        const wallet = norm(row["Wallet"]);
        if (wallet)
            ids.add(wallet);
    }
    return ids;
}
// ── Extract user IDs from MTF record row ──────────────────────────────────────
const MTF_ID_COL = "Your User Id for XM and Xellion (NA - if not applicable)";
const SKIP_VALUES = new Set(["na", "n/a", "nil", "none", "", "xm", "xilion", "xellion"]);
function extractUserIds(row) {
    const raw = row[MTF_ID_COL];
    if (raw === null || raw === undefined || raw === "")
        return [];
    return String(raw)
        .split(/[,\s\/]+/)
        .map(norm)
        .filter((v) => v && !SKIP_VALUES.has(v));
}
// ── Normalize values for comparison ───────────────────────────────────────────
function normalizeForCompare(val) {
    if (val === null || val === undefined)
        return "";
    return String(val)
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ")
        .replace(/^\+/, "");
}
function normalizeName(val) {
    if (val === null || val === undefined)
        return "";
    return String(val).trim().toLowerCase().replace(/\s+/g, " ");
}
function normalizePhone(val) {
    if (val === null || val === undefined)
        return "";
    const digits = String(val).replace(/\D/g, "");
    if (digits.length > 10 && digits.startsWith("91"))
        return digits.slice(-10);
    return digits;
}
function autoSizeColumns(ws, rows, fallbackKeys) {
    const keys = Object.keys(rows[0] ?? {}).length > 0 ? Object.keys(rows[0]) : fallbackKeys;
    ws["!cols"] = keys.map((k) => ({
        wch: Math.max(k.length, ...rows.map((r) => String(r[k] ?? "").length)),
    }));
}
function writeRowsToWorkbook(rows, sheetName, filePrefix, fallbackKeys) {
    const outputDir = path.join(__dirname, "outputs");
    if (!fs.existsSync(outputDir))
        fs.mkdirSync(outputDir, { recursive: true });
    const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const outputPath = path.join(outputDir, `${filePrefix}_${ts}.xlsx`);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows, { header: fallbackKeys });
    autoSizeColumns(ws, rows, fallbackKeys);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, outputPath);
    return outputPath;
}
function getVerifiedMasterRows(masterRows) {
    return masterRows.filter((row) => row["Verification"] === "Yes");
}
function resolveBundledFile(fileName) {
    const candidates = [
        path.join(process.cwd(), "files", fileName),
        path.join(__dirname, "files", fileName),
        path.join(__dirname, "..", "files", fileName),
    ];
    const found = candidates.find((candidate) => fs.existsSync(candidate));
    if (!found)
        throw new Error(`Could not find ${fileName} in the files directory.`);
    return found;
}
// ── Column-agnostic field extractors ─────────────────────────────────────────
// Handles column name differences between Old ("Email Id", "Contact Number")
// and New ("Email ID", "Phone Number") sheets.
function getEmail(row) {
    return normalizeForCompare(row["Email Id"] ||
        row["Email ID"] ||
        row["Email"] ||
        "");
}
function getPhone(row) {
    return normalizeForCompare(row["Contact Number"] ||
        row["Phone Number"] ||
        row["Phone"] ||
        row["Contact"] ||
        "");
}
// ── Shared: build phone set from New rows (phone only matching) ──────────────
function buildNewPhoneSet(newRows) {
    const newByPhone = new Set();
    for (const row of newRows) {
        const phone = getPhone(row);
        if (phone)
            newByPhone.add(phone);
    }
    return newByPhone;
}
// ── Old Only export ───────────────────────────────────────────────────────────
// Returns only rows from Old that have NO phone match in New.
// Output sheet contains all original Old fields — no extra columns added.
async function processOldOnly(opts) {
    const { oldPath, newPath } = opts;
    const oldRows = readStandardSheet(oldPath);
    const newRows = readStandardSheet(newPath);
    if (oldRows.length === 0)
        throw new Error("old.csv appears to be empty or has no data rows.");
    if (newRows.length === 0)
        throw new Error("new.csv appears to be empty or has no data rows.");
    // Build phone lookup set from New (phone only)
    const newByPhone = buildNewPhoneSet(newRows);
    // Keep only Old rows that have NO phone match in New
    const oldOnlyRows = [];
    for (const row of oldRows) {
        const phone = getPhone(row);
        const isInNew = phone && newByPhone.has(phone);
        if (!isInNew) {
            oldOnlyRows.push({ ...row });
        }
    }
    // Write output XLSX
    const outputDir = path.join(__dirname, "outputs");
    if (!fs.existsSync(outputDir))
        fs.mkdirSync(outputDir, { recursive: true });
    const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const outputPath = path.join(outputDir, `old_only_${ts}.xlsx`);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(oldOnlyRows);
    // Auto-size columns
    if (oldOnlyRows.length > 0) {
        const keys = Object.keys(oldOnlyRows[0]);
        ws["!cols"] = keys.map((k) => ({
            wch: Math.max(k.length, ...oldOnlyRows.map((r) => String(r[k] ?? "").length)),
        }));
    }
    XLSX.utils.book_append_sheet(wb, ws, "Old Only");
    XLSX.writeFile(wb, outputPath);
    return {
        outputPath,
        oldCount: oldRows.length,
        newCount: newRows.length,
        oldOnlyCount: oldOnlyRows.length,
    };
}
// ── Static reports from bundled workbook files ────────────────────────────────
async function processNonDepositedReport(opts = {}) {
    const masterPath = opts.masterPath ?? resolveBundledFile("MTF Indicator Access (Responses).xlsx");
    const nonDepositedPath = opts.nonDepositedPath ?? resolveBundledFile("NON DEPOSITED LIST.xlsx");
    const masterRows = readStandardSheet(masterPath);
    const nonDepositedRows = readStandardSheet(nonDepositedPath);
    if (masterRows.length === 0)
        throw new Error("Master sheet appears to be empty or has no data rows.");
    if (nonDepositedRows.length === 0)
        throw new Error("Non-deposited sheet appears to be empty or has no data rows.");
    const verifiedRows = getVerifiedMasterRows(masterRows);
    const nonDepositedKeys = new Map();
    for (const row of nonDepositedRows) {
        const phone = normalizePhone(row["Phone"]);
        const fullName = normalizeName(`${row["Firstname"] ?? ""} ${row["Lastname"] ?? ""}`);
        if (phone && fullName)
            nonDepositedKeys.set(`${phone}|${fullName}`, row);
    }
    const matchedRows = [];
    for (const row of verifiedRows) {
        const phone = normalizePhone(row["Phone Number"]);
        const name = normalizeName(row["Name"]);
        const match = nonDepositedKeys.get(`${phone}|${name}`);
        if (match) {
            matchedRows.push({
                Name: row["Name"],
                "Phone Number": row["Phone Number"],
                Verification: row["Verification"],
                "Matched Firstname": match["Firstname"],
                "Matched Lastname": match["Lastname"],
                "Matched Phone": match["Phone"],
                Email: match["Email"],
                Login: match["Login"],
                Agent: match["Agent"],
            });
        }
    }
    const fallbackKeys = [
        "Name",
        "Phone Number",
        "Verification",
        "Matched Firstname",
        "Matched Lastname",
        "Matched Phone",
        "Email",
        "Login",
        "Agent",
    ];
    return {
        outputPath: writeRowsToWorkbook(matchedRows, "Non Deposited", "non_deposited_verified", fallbackKeys),
        verifiedCount: verifiedRows.length,
        sourceCount: nonDepositedRows.length,
        matchedCount: matchedRows.length,
    };
}
async function processInactiveUsersReport(opts = {}) {
    const masterPath = opts.masterPath ?? resolveBundledFile("MTF Indicator Access (Responses).xlsx");
    const inactivePath = opts.inactivePath ?? resolveBundledFile("DP BUT INACTIVE.xlsx");
    const masterRows = readStandardSheet(masterPath);
    const inactiveRows = readStandardSheet(inactivePath);
    if (masterRows.length === 0)
        throw new Error("Master sheet appears to be empty or has no data rows.");
    if (inactiveRows.length === 0)
        throw new Error("Inactive sheet appears to be empty or has no data rows.");
    const verifiedRows = getVerifiedMasterRows(masterRows);
    const inactiveKeys = new Map();
    for (const row of inactiveRows) {
        const phone = normalizePhone(row["Phone number"]);
        const name = normalizeName(row["Customer"]);
        if (phone && name)
            inactiveKeys.set(`${phone}|${name}`, row);
    }
    const matchedRows = [];
    for (const row of verifiedRows) {
        const phone = normalizePhone(row["Phone Number"]);
        const name = normalizeName(row["Name"]);
        const match = inactiveKeys.get(`${phone}|${name}`);
        if (match) {
            matchedRows.push({
                Name: row["Name"],
                "Phone Number": row["Phone Number"],
                Verification: row["Verification"],
                "Matched Customer": match["Customer"],
                "Matched Phone": match["Phone number"],
                Login: match["Login"],
            });
        }
    }
    const fallbackKeys = [
        "Name",
        "Phone Number",
        "Verification",
        "Matched Customer",
        "Matched Phone",
        "Login",
    ];
    return {
        outputPath: writeRowsToWorkbook(matchedRows, "Inactive Users", "inactive_verified_users", fallbackKeys),
        verifiedCount: verifiedRows.length,
        sourceCount: inactiveRows.length,
        matchedCount: matchedRows.length,
    };
}
// ── Duplicate detection export ────────────────────────────────────────────────
async function processDuplicates(opts) {
    const { oldPath, newPath } = opts;
    const oldRows = readStandardSheet(oldPath);
    const newRows = readStandardSheet(newPath);
    if (oldRows.length === 0)
        throw new Error("old.csv appears to be empty or has no data rows.");
    if (newRows.length === 0)
        throw new Error("new.csv appears to be empty or has no data rows.");
    // Build lookup maps from Old rows (by email and phone separately)
    const oldByEmail = new Map();
    const oldByPhone = new Map();
    for (const row of oldRows) {
        const email = getEmail(row);
        const phone = getPhone(row);
        if (email)
            oldByEmail.set(email, row);
        if (phone)
            oldByPhone.set(phone, row);
    }
    let oldOnlyCount = 0;
    let newOnlyCount = 0;
    let bothCount = 0;
    const matchedOldEmails = new Set();
    const matchedOldPhones = new Set();
    const outputRows = [];
    for (const row of newRows) {
        const email = getEmail(row);
        const phone = getPhone(row);
        const matchedByEmail = !!(email && oldByEmail.has(email));
        const matchedByPhone = !!(phone && oldByPhone.has(phone));
        const isInOld = matchedByEmail || matchedByPhone;
        if (isInOld) {
            bothCount++;
            if (matchedByEmail)
                matchedOldEmails.add(email);
            if (matchedByPhone)
                matchedOldPhones.add(phone);
        }
        else {
            newOnlyCount++;
        }
        let matchType = "-";
        if (matchedByEmail && matchedByPhone)
            matchType = "Email + Phone";
        else if (matchedByEmail)
            matchType = "Email";
        else if (matchedByPhone)
            matchType = "Phone";
        outputRows.push({
            ...row,
            "In Old": isInOld ? "Yes" : "No",
            "In New": "Yes",
            "Match Type": matchType,
            "Presence": isInOld ? "Present in Both" : "Present in New Only",
        });
    }
    // Add Old-only rows
    for (const row of oldRows) {
        const email = getEmail(row);
        const phone = getPhone(row);
        const wasMatched = (email && matchedOldEmails.has(email)) ||
            (phone && matchedOldPhones.has(phone));
        if (!wasMatched) {
            oldOnlyCount++;
            outputRows.push({
                ...row,
                "In Old": "Yes",
                "In New": "No",
                "Match Type": "-",
                "Presence": "Present in Old Only",
            });
        }
    }
    // Sort: Old Only → New Only → Both
    outputRows.sort((a, b) => {
        const order = {
            "Present in Old Only": 0,
            "Present in New Only": 1,
            "Present in Both": 2,
        };
        return (order[a["Presence"]] ?? 0) - (order[b["Presence"]] ?? 0);
    });
    // Write output XLSX
    const outputDir = path.join(__dirname, "outputs");
    if (!fs.existsSync(outputDir))
        fs.mkdirSync(outputDir, { recursive: true });
    const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const outputPath = path.join(outputDir, `duplicates_${ts}.xlsx`);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(outputRows);
    const keys = Object.keys(outputRows[0] ?? {});
    ws["!cols"] = keys.map((k) => ({
        wch: Math.max(k.length, ...outputRows.map((r) => String(r[k] ?? "").length)),
    }));
    XLSX.utils.book_append_sheet(wb, ws, "Comparison");
    XLSX.writeFile(wb, outputPath);
    return {
        outputPath,
        oldCount: oldRows.length,
        newCount: newRows.length,
        duplicatesFound: bothCount,
        oldOnly: oldOnlyCount,
        newOnly: newOnlyCount,
        both: bothCount,
    };
}
// ── Verification export ───────────────────────────────────────────────────────
async function processVerification(opts) {
    const { xmPath, xilionPath, mtfPath } = opts;
    const xmRows = readStandardSheet(xmPath);
    const xilionRows = readXilionSheet(xilionPath);
    const mtfRows = readStandardSheet(mtfPath);
    if (xmRows.length === 0)
        throw new Error("xm.csv appears to be empty or has no data rows.");
    if (xilionRows.length === 0)
        throw new Error("xilion.csv appears to be empty or has no data rows.");
    if (mtfRows.length === 0)
        throw new Error("mtf.csv appears to be empty or has no data rows.");
    const xmIds = buildXmIds(xmRows);
    const xilionIds = buildXilionIds(xilionRows);
    if (xmIds.size === 0)
        throw new Error("No IDs found in xm.csv — expected 'MT4/MT5 ID' and 'Client ID' columns.");
    if (xilionIds.size === 0)
        throw new Error("No IDs found in xilion.csv — expected a 'Wallet' column.");
    let verified = 0;
    let notVerified = 0;
    const outputRows = mtfRows.map((row) => {
        const userIds = extractUserIds(row);
        const isVerified = userIds.some((id) => xmIds.has(id)) ||
            userIds.some((id) => xilionIds.has(id));
        if (isVerified)
            verified++;
        else
            notVerified++;
        return { ...row, Verified: isVerified ? "Yes" : "No" };
    });
    const outputDir = path.join(__dirname, "outputs");
    if (!fs.existsSync(outputDir))
        fs.mkdirSync(outputDir, { recursive: true });
    const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const outputPath = path.join(outputDir, `verified_${ts}.xlsx`);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(outputRows);
    const keys = Object.keys(outputRows[0] ?? {});
    ws["!cols"] = keys.map((k) => ({
        wch: Math.max(k.length, ...outputRows.map((r) => String(r[k] ?? "").length)),
    }));
    XLSX.utils.book_append_sheet(wb, ws, "Verified Records");
    XLSX.writeFile(wb, outputPath);
    return { outputPath, total: outputRows.length, verified, notVerified };
}
