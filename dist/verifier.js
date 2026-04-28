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
    // Find the actual header row (first row whose first cell is "Wallet")
    const headerRowIdx = raw.findIndex((row) => String(row[0]).trim().toLowerCase() === "wallet");
    if (headerRowIdx === -1) {
        // Fallback: treat row 1 as header (skip row 0 sep= line)
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
/** Trim, strip leading #, lowercase */
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
// ── Main export ───────────────────────────────────────────────────────────────
async function processVerification(opts) {
    const { xmPath, xilionPath, mtfPath } = opts;
    // 1. Read all 3 sheets
    const xmRows = readStandardSheet(xmPath);
    const xilionRows = readXilionSheet(xilionPath);
    const mtfRows = readStandardSheet(mtfPath);
    if (xmRows.length === 0)
        throw new Error("xm.csv appears to be empty or has no data rows.");
    if (xilionRows.length === 0)
        throw new Error("xilion.csv appears to be empty or has no data rows.");
    if (mtfRows.length === 0)
        throw new Error("mtf.csv appears to be empty or has no data rows.");
    // 2. Build combined broker ID lookup (XM + Xilion merged)
    const xmIds = buildXmIds(xmRows);
    const xilionIds = buildXilionIds(xilionRows);
    if (xmIds.size === 0)
        throw new Error("No IDs found in xm.csv — expected 'MT4/MT5 ID' and 'Client ID' columns.");
    if (xilionIds.size === 0)
        throw new Error("No IDs found in xilion.csv — expected a 'Wallet' column.");
    // 3. Tag each MTF record
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
    // 4. Write output XLSX
    const outputDir = path.join(__dirname, "outputs");
    if (!fs.existsSync(outputDir))
        fs.mkdirSync(outputDir, { recursive: true });
    const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const outputPath = path.join(outputDir, `verified_${ts}.xlsx`);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(outputRows);
    // Auto-size columns
    const keys = Object.keys(outputRows[0] ?? {});
    ws["!cols"] = keys.map((k) => ({
        wch: Math.max(k.length, ...outputRows.map((r) => String(r[k] ?? "").length)),
    }));
    XLSX.utils.book_append_sheet(wb, ws, "Verified Records");
    XLSX.writeFile(wb, outputPath);
    return { outputPath, total: outputRows.length, verified, notVerified };
}
