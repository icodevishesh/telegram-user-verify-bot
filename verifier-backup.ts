import * as XLSX from "xlsx";
import * as path from "path";
import * as fs from "fs";

// ── Public interface ──────────────────────────────────────────────────────────
export interface VerifyOptions {
  xmPath: string;
  xilionPath: string;
  mtfPath: string;
}

export interface VerifyResult {
  outputPath: string;
  total: number;
  verified: number;
  notVerified: number;
}

// ── Sheet readers ─────────────────────────────────────────────────────────────

function readStandardSheet(filePath: string): Record<string, string>[] {
  const wb = XLSX.readFile(filePath, { raw: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json<Record<string, string>>(ws, { defval: "" });
}

/**
 * Xilion CSVs have a "sep=," sentinel on row 0.
 * Real headers are on row 1, data starts at row 2.
 */
function readXilionSheet(filePath: string): Record<string, string>[] {
  const wb = XLSX.readFile(filePath, { raw: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" });

  // Find the actual header row (first row whose first cell is "Wallet")
  const headerRowIdx = raw.findIndex(
    (row) => String(row[0]).trim().toLowerCase() === "wallet"
  );

  if (headerRowIdx === -1) {
    // Fallback: treat row 1 as header (skip row 0 sep= line)
    const headers = raw[1] as string[];
    return (raw.slice(2) as string[][]).map((row) => {
      const obj: Record<string, string> = {};
      headers.forEach((h, i) => { obj[h] = String(row[i] ?? ""); });
      return obj;
    });
  }

  const headers = raw[headerRowIdx] as string[];
  return (raw.slice(headerRowIdx + 1) as string[][]).map((row) => {
    const obj: Record<string, string> = {};
    headers.forEach((h, i) => { obj[h] = String(row[i] ?? ""); });
    return obj;
  });
}

// ── ID normalisation ──────────────────────────────────────────────────────────

/** Trim, strip leading #, lowercase */
function norm(val: unknown): string {
  if (val === null || val === undefined) return "";
  return String(val).trim().replace(/^#/, "").toLowerCase();
}

// ── Build broker ID sets ──────────────────────────────────────────────────────

function buildXmIds(rows: Record<string, string>[]): Set<string> {
  const ids = new Set<string>();
  for (const row of rows) {
    const mt = norm(row["MT4/MT5 ID"]);
    const client = norm(row["Client ID"]);
    if (mt) ids.add(mt);
    if (client) ids.add(client);
  }
  return ids;
}

function buildXilionIds(rows: Record<string, string>[]): Set<string> {
  const ids = new Set<string>();
  for (const row of rows) {
    const wallet = norm(row["Wallet"]);
    if (wallet) ids.add(wallet);
  }
  return ids;
}

// ── Extract user IDs from MTF record row ──────────────────────────────────────

const MTF_ID_COL = "Your User Id for XM and Xellion (NA - if not applicable)";
const SKIP_VALUES = new Set(["na", "n/a", "nil", "none", "", "xm", "xilion", "xellion"]);

function extractUserIds(row: Record<string, unknown>): string[] {
  const raw = row[MTF_ID_COL];
  if (raw === null || raw === undefined || raw === "") return [];
  return String(raw)
    .split(/[,\s\/]+/)
    .map(norm)
    .filter((v) => v && !SKIP_VALUES.has(v));
}

// ── Duplicate detection types ─────────────────────────────────────────────────
export interface DuplicateOptions {
  oldPath: string;
  newPath: string;
}

export interface DuplicateResult {
  outputPath: string;
  oldCount: number;
  newCount: number;
  duplicatesFound: number;
  oldOnly: number;
  newOnly: number;
  both: number;
}

// ── Normalize values for comparison ───────────────────────────────────────────
function normalizeForCompare(val: unknown): string {
  if (val === null || val === undefined) return "";
  return String(val)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/^\+/, "");
}

// ── Column-agnostic field extractors ─────────────────────────────────────────
// Handles column name differences between Old ("Email Id", "Contact Number")
// and New ("Email ID", "Phone Number") sheets.

function getEmail(row: Record<string, string>): string {
  return normalizeForCompare(
    row["Email Id"] ||
    row["Email ID"] ||
    row["Email"] ||
    ""
  );
}

function getPhone(row: Record<string, string>): string {
  return normalizeForCompare(
    row["Contact Number"] ||
    row["Phone Number"] ||
    row["Phone"] ||
    row["Contact"] ||
    ""
  );
}

// ── Duplicate detection export ────────────────────────────────────────────────
export async function processDuplicates(opts: DuplicateOptions): Promise<DuplicateResult> {
  const { oldPath, newPath } = opts;

  // Read both sheets
  const oldRows = readStandardSheet(oldPath);
  const newRows = readStandardSheet(newPath);

  if (oldRows.length === 0) throw new Error("old.csv appears to be empty or has no data rows.");
  if (newRows.length === 0) throw new Error("new.csv appears to be empty or has no data rows.");

  // ── Build lookup maps from Old rows (by email and phone separately) ──────
  // Using OR logic: a person is a duplicate if email OR phone matches.
  const oldByEmail = new Map<string, Record<string, string>>();
  const oldByPhone = new Map<string, Record<string, string>>();

  for (const row of oldRows) {
    const email = getEmail(row);
    const phone = getPhone(row);
    if (email) oldByEmail.set(email, row);
    if (phone) oldByPhone.set(phone, row);
  }

  // ── Process New rows ──────────────────────────────────────────────────────
  let oldOnlyCount = 0;
  let newOnlyCount = 0;
  let bothCount = 0;

  // Track which old emails/phones were matched (to find Old-only rows later)
  const matchedOldEmails = new Set<string>();
  const matchedOldPhones = new Set<string>();

  const outputRows: Record<string, string>[] = [];

  for (const row of newRows) {
    const email = getEmail(row);
    const phone = getPhone(row);

    const matchedByEmail = !!(email && oldByEmail.has(email));
    const matchedByPhone = !!(phone && oldByPhone.has(phone));
    const isInOld = matchedByEmail || matchedByPhone;

    if (isInOld) {
      bothCount++;
      if (matchedByEmail) matchedOldEmails.add(email);
      if (matchedByPhone) matchedOldPhones.add(phone);
    } else {
      newOnlyCount++;
    }

    let matchType = "-";
    if (matchedByEmail && matchedByPhone) matchType = "Email + Phone";
    else if (matchedByEmail) matchType = "Email";
    else if (matchedByPhone) matchType = "Phone";

    outputRows.push({
      ...row,
      "In Old": isInOld ? "Yes" : "No",
      "In New": "Yes",
      "Match Type": matchType,
      "Presence": isInOld ? "Present in Both" : "Present in New Only",
    });
  }

  // ── Add Old-only rows (those not matched by any New row) ──────────────────
  for (const row of oldRows) {
    const email = getEmail(row);
    const phone = getPhone(row);

    const wasMatched =
      (email && matchedOldEmails.has(email)) ||
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

  // ── Sort: Old Only → New Only → Both ─────────────────────────────────────
  outputRows.sort((a, b) => {
    const order: Record<string, number> = {
      "Present in Old Only": 0,
      "Present in New Only": 1,
      "Present in Both": 2,
    };
    return (order[a["Presence"]] ?? 0) - (order[b["Presence"]] ?? 0);
  });

  // ── Write output XLSX ─────────────────────────────────────────────────────
  const outputDir = path.join(__dirname, "outputs");
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const outputPath = path.join(outputDir, `duplicates_${ts}.xlsx`);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(outputRows);

  // Auto-size columns
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
export async function processVerification(opts: VerifyOptions): Promise<VerifyResult> {
  const { xmPath, xilionPath, mtfPath } = opts;

  // 1. Read all 3 sheets
  const xmRows     = readStandardSheet(xmPath);
  const xilionRows = readXilionSheet(xilionPath);
  const mtfRows    = readStandardSheet(mtfPath);

  if (xmRows.length === 0)     throw new Error("xm.csv appears to be empty or has no data rows.");
  if (xilionRows.length === 0) throw new Error("xilion.csv appears to be empty or has no data rows.");
  if (mtfRows.length === 0)    throw new Error("mtf.csv appears to be empty or has no data rows.");

  // 2. Build combined broker ID lookup (XM + Xilion merged)
  const xmIds     = buildXmIds(xmRows);
  const xilionIds = buildXilionIds(xilionRows);

  if (xmIds.size === 0)     throw new Error("No IDs found in xm.csv — expected 'MT4/MT5 ID' and 'Client ID' columns.");
  if (xilionIds.size === 0) throw new Error("No IDs found in xilion.csv — expected a 'Wallet' column.");

  // 3. Tag each MTF record
  let verified = 0;
  let notVerified = 0;

  const outputRows = mtfRows.map((row) => {
    const userIds = extractUserIds(row);
    const isVerified =
      userIds.some((id) => xmIds.has(id)) ||
      userIds.some((id) => xilionIds.has(id));

    if (isVerified) verified++; else notVerified++;
    return { ...row, Verified: isVerified ? "Yes" : "No" };
  });

  // 4. Write output XLSX
  const outputDir = path.join(__dirname, "outputs");
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const outputPath = path.join(outputDir, `verified_${ts}.xlsx`);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(outputRows);

  // Auto-size columns
  const keys = Object.keys(outputRows[0] ?? {});
  ws["!cols"] = keys.map((k) => ({
    wch: Math.max(k.length, ...outputRows.map((r) => String(r[k as keyof typeof r] ?? "").length)),
  }));

  XLSX.utils.book_append_sheet(wb, ws, "Verified Records");
  XLSX.writeFile(wb, outputPath);

  return { outputPath, total: outputRows.length, verified, notVerified };
}