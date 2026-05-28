import "dotenv/config";
import TelegramBot from "node-telegram-bot-api";
import * as fs from "fs";
import * as path from "path";
import {
  processVerification,
  processDuplicates,
  processOldOnly,
  processNonDepositedReport,
  processInactiveUsersReport,
} from "./verifier";

const TOKEN = process.env.TELEGRAM_BOT_TOKEN!;
if (!TOKEN) {
  console.error("❌ TELEGRAM_BOT_TOKEN is not set in .env");
  process.exit(1);
}

const bot = new TelegramBot(TOKEN, { polling: true });

// ── Session type ──────────────────────────────────────────────────────────────
interface UploadedFile {
  originalName: string;
  savedPath: string;
}

type SessionType = "verify" | "duplicates" | "oldonly" | "nodeposited" | "inactiveusers";

interface Session {
  type: SessionType;
  xm?: UploadedFile;
  xilion?: UploadedFile;
  mtf?: UploadedFile;
  old?: UploadedFile;
  new?: UploadedFile;
  master?: UploadedFile;
  nonDeposited?: UploadedFile;
  inactive?: UploadedFile;
  processing: boolean;
}

const sessions: Record<number, Session> = {};

// ── Accepted filenames ────────────────────────────────────────────────────────
const VERIFY_NAMES = ["xm", "xilion", "mtf"] as const;
type VerifyName = (typeof VERIFY_NAMES)[number];

const DUPLICATE_NAMES = ["old", "new"] as const;
type DuplicateName = (typeof DUPLICATE_NAMES)[number];

const NON_DEPOSITED_NAMES = ["master", "nonDeposited"] as const;
type NonDepositedName = (typeof NON_DEPOSITED_NAMES)[number];

const INACTIVE_USERS_NAMES = ["master", "inactive"] as const;
type InactiveUsersName = (typeof INACTIVE_USERS_NAMES)[number];

const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// ── Helpers ───────────────────────────────────────────────────────────────────
function resolveVerifySlot(fileName: string): VerifyName | null {
  const base = path.basename(fileName, path.extname(fileName)).toLowerCase().trim();
  if (VERIFY_NAMES.includes(base as VerifyName)) return base as VerifyName;
  return null;
}

function resolveDuplicateSlot(fileName: string): DuplicateName | null {
  const base = path.basename(fileName, path.extname(fileName)).toLowerCase().trim();
  if (DUPLICATE_NAMES.includes(base as DuplicateName)) return base as DuplicateName;
  return null;
}

function resolveNonDepositedSlot(fileName: string): NonDepositedName | null {
  const base = path.basename(fileName, path.extname(fileName)).toLowerCase().trim();
  if (base === "mtf indicator access (responses)") return "master";
  if (base === "non deposited list") return "nonDeposited";
  return null;
}

function resolveInactiveUsersSlot(fileName: string): InactiveUsersName | null {
  const base = path.basename(fileName, path.extname(fileName)).toLowerCase().trim();
  if (base === "mtf indicator access (responses)") return "master";
  if (base === "dp but inactive") return "inactive";
  return null;
}

function verifySessionSummary(s: Session): string {
  return [
    s.xm     ? `✅ xm.csv`     : `⬜ xm.csv`,
    s.xilion ? `✅ xilion.csv` : `⬜ xilion.csv`,
    s.mtf    ? `✅ mtf.csv`    : `⬜ mtf.csv`,
  ].join("\n");
}

function duplicateSessionSummary(s: Session): string {
  return [
    s.old ? `✅ old.csv` : `⬜ old.csv`,
    s.new ? `✅ new.csv` : `⬜ new.csv`,
  ].join("\n");
}

function nonDepositedSessionSummary(s: Session): string {
  return [
    s.master ? `✅ MTF Indicator Access (Responses).xlsx` : `⬜ MTF Indicator Access (Responses).xlsx`,
    s.nonDeposited ? `✅ NON DEPOSITED LIST.xlsx` : `⬜ NON DEPOSITED LIST.xlsx`,
  ].join("\n");
}

function inactiveUsersSessionSummary(s: Session): string {
  return [
    s.master ? `✅ MTF Indicator Access (Responses).xlsx` : `⬜ MTF Indicator Access (Responses).xlsx`,
    s.inactive ? `✅ DP BUT INACTIVE.xlsx` : `⬜ DP BUT INACTIVE.xlsx`,
  ].join("\n");
}

function isVerifyComplete(s: Session): boolean {
  return !!(s.xm && s.xilion && s.mtf);
}

function isDuplicateComplete(s: Session): boolean {
  return !!(s.old && s.new);
}

function isNonDepositedComplete(s: Session): boolean {
  return !!(s.master && s.nonDeposited);
}

function isInactiveUsersComplete(s: Session): boolean {
  return !!(s.master && s.inactive);
}

function cleanupSession(chatId: number) {
  const s = sessions[chatId];
  if (!s) return;
  const keys =
    s.type === "verify"
      ? VERIFY_NAMES
      : s.type === "nodeposited"
        ? NON_DEPOSITED_NAMES
        : s.type === "inactiveusers"
          ? INACTIVE_USERS_NAMES
        : DUPLICATE_NAMES;
  for (const key of keys) {
    const f = s[key as keyof Session] as UploadedFile | undefined;
    if (f?.savedPath && fs.existsSync(f.savedPath)) {
      try { fs.unlinkSync(f.savedPath); } catch {}
    }
  }
  delete sessions[chatId];
}

// ── Commands ──────────────────────────────────────────────────────────────────
bot.onText(/\/start/, (msg) => {
  bot.sendMessage(
    msg.chat.id,
    `👋 *Welcome to the MTF Verification Bot!*\n\n` +
    `Matches user IDs from your records against XM and Xilion broker sheets.\n\n` +
    `*Commands:*\n` +
    `/verify — Start a new verification session\n` +
    `/duplicates — Compare old & new CSV sheets for duplicates\n` +
    `/oldonly — Extract records present in Old but not in New\n` +
    `/nodeposited — Upload master + non-deposited files and list verified users\n` +
    `/nondeposited — Alias for /nodeposited\n` +
    `/inactiveusers — Upload master + inactive files and list verified users\n` +
    `/status — Check which files have been uploaded\n` +
    `/reset  — Clear the current session\n` +
    `/help   — Show usage instructions`,
    { parse_mode: "Markdown" }
  );
});

bot.onText(/\/help/, (msg) => {
  bot.sendMessage(
    msg.chat.id,
    `*How to use:*\n\n` +
    `*Verification:*\n` +
    `1️⃣ Send /verify to start\n` +
    `2️⃣ Upload all *3 files* (in any order):\n` +
    `   • \`xm.csv\` — XM broker sheet\n` +
    `   • \`xilion.csv\` — Xilion broker sheet\n` +
    `   • \`mtf.csv\` — Your user records\n` +
    `3️⃣ Processing starts automatically once all 3 are received\n` +
    `4️⃣ Bot sends back an XLSX with a *Verified* column\n\n` +
    `*Duplicate Detection:*\n` +
    `1️⃣ Send /duplicates to start\n` +
    `2️⃣ Upload *2 files* (in any order):\n` +
    `   • \`old.csv\` — Old records\n` +
    `   • \`new.csv\` — New records\n` +
    `3️⃣ Processing starts automatically\n` +
    `4️⃣ Bot sends back an XLSX with full comparison\n\n` +
    `*Old Only (New Leads):*\n` +
    `1️⃣ Send /oldonly to start\n` +
    `2️⃣ Upload *2 files* (in any order):\n` +
    `   • \`old.csv\` — Old records\n` +
    `   • \`new.csv\` — New records\n` +
    `3️⃣ Processing starts automatically\n` +
    `4️⃣ Bot sends back an XLSX with *only* records from Old that are NOT in New\n\n` +
    `*Non Deposited:*\n` +
    `1️⃣ Send /nodeposited\n` +
    `2️⃣ Upload *2 files* (in any order):\n` +
    `   • \`MTF Indicator Access (Responses).xlsx\`\n` +
    `   • \`NON DEPOSITED LIST.xlsx\`\n` +
    `3️⃣ Bot sends back an XLSX with verified users found in the non-deposited sheet\n\n` +
    `*Inactive Users:*\n` +
    `1️⃣ Send /inactiveusers\n` +
    `2️⃣ Upload *2 files* (in any order):\n` +
    `   • \`MTF Indicator Access (Responses).xlsx\`\n` +
    `   • \`DP BUT INACTIVE.xlsx\`\n` +
    `3️⃣ Bot sends back an XLSX with verified users found in the inactive sheet\n\n` +
    `⚠️ Files must be named *exactly* as shown above\n\n` +
    `*Matching logic:*\n` +
    `• XM → \`MT4/MT5 ID\` and \`Client ID\` columns\n` +
    `• Xilion → \`Wallet\` column (e.g. #316393)\n` +
    `• Duplicates → Matches on \`Email\` OR \`Phone\` (either field is enough)\n` +
    `• Old Only → Matches on \`Phone\` only\n` +
    `• Non Deposited → Verified users matched by \`Phone Number\` + \`Name\`\n` +
    `• Inactive Users → Verified users matched by \`Phone Number\` + \`Name\``,
    { parse_mode: "Markdown" }
  );
});

bot.onText(/\/reset/, (msg) => {
  cleanupSession(msg.chat.id);
  bot.sendMessage(msg.chat.id, "🔄 Session cleared. Send /verify, /duplicates, /oldonly, /nodeposited, or /inactiveusers to start again.");
});

bot.onText(/\/status/, (msg) => {
  const chatId = msg.chat.id;
  const s = sessions[chatId];
  if (!s) {
    bot.sendMessage(chatId, "No active session. Send /verify, /duplicates, /oldonly, /nodeposited, or /inactiveusers to begin.");
    return;
  }

  if (s.type === "verify") {
    const extra = s.processing
      ? `⚙️ Verification is currently running…`
      : isVerifyComplete(s) ? `✅ All files received.` : `Upload the remaining files to continue.`;
    bot.sendMessage(
      chatId,
      `*Upload Status:*\n\n${verifySessionSummary(s)}\n\n${extra}`,
      { parse_mode: "Markdown" }
    );
  } else if (s.type === "nodeposited") {
    const extra = s.processing
      ? `⚙️ Non-deposited report is currently running…`
      : isNonDepositedComplete(s) ? `✅ All files received.` : `Upload the remaining files to continue.`;
    bot.sendMessage(
      chatId,
      `*Upload Status:*\n\n${nonDepositedSessionSummary(s)}\n\n${extra}`,
      { parse_mode: "Markdown" }
    );
  } else if (s.type === "inactiveusers") {
    const extra = s.processing
      ? `⚙️ Inactive users report is currently running…`
      : isInactiveUsersComplete(s) ? `✅ All files received.` : `Upload the remaining files to continue.`;
    bot.sendMessage(
      chatId,
      `*Upload Status:*\n\n${inactiveUsersSessionSummary(s)}\n\n${extra}`,
      { parse_mode: "Markdown" }
    );
  } else {
    // duplicates or oldonly — both use old + new
    const extra = s.processing
      ? `⚙️ Processing is currently running…`
      : isDuplicateComplete(s) ? `✅ All files received.` : `Upload the remaining files to continue.`;
    bot.sendMessage(
      chatId,
      `*Upload Status:*\n\n${duplicateSessionSummary(s)}\n\n${extra}`,
      { parse_mode: "Markdown" }
    );
  }
});

bot.onText(/\/verify/, (msg) => {
  const chatId = msg.chat.id;
  cleanupSession(chatId);
  sessions[chatId] = { type: "verify", processing: false };
  bot.sendMessage(
    chatId,
    `✅ *Verification session started!*\n\n` +
    `Upload these *3 files* in any order:\n\n` +
    `⬜ \`xm.csv\`\n` +
    `⬜ \`xilion.csv\`\n` +
    `⬜ \`mtf.csv\`\n\n` +
    `⚠️ Files must be named exactly as shown above.`,
    { parse_mode: "Markdown" }
  );
});

bot.onText(/\/duplicates/, (msg) => {
  const chatId = msg.chat.id;
  cleanupSession(chatId);
  sessions[chatId] = { type: "duplicates", processing: false };
  bot.sendMessage(
    chatId,
    `✅ *Duplicate check session started!*\n\n` +
    `Upload these *2 files* in any order:\n\n` +
    `⬜ \`old.csv\` — Old records\n` +
    `⬜ \`new.csv\` — New records\n\n` +
    `⚠️ Files must be named exactly as shown above.\n\n` +
    `📋 Matches on: *Email* OR *Phone* (either field is enough)`,
    { parse_mode: "Markdown" }
  );
});

bot.onText(/\/oldonly/, (msg) => {
  const chatId = msg.chat.id;
  cleanupSession(chatId);
  sessions[chatId] = { type: "oldonly", processing: false };
  bot.sendMessage(
    chatId,
    `✅ *Old Only session started!*\n\n` +
    `This will extract records from *old.csv* that are *NOT* present in *new.csv*.\n\n` +
    `Upload these *2 files* in any order:\n\n` +
    `⬜ \`old.csv\` — Old records\n` +
    `⬜ \`new.csv\` — New records\n\n` +
    `⚠️ Files must be named exactly as shown above.\n\n` +
    `📋 Matches on: *Phone Number* only`,
    { parse_mode: "Markdown" }
  );
});

bot.onText(/\/(?:nodeposited|nondeposited)/, (msg) => {
  const chatId = msg.chat.id;
  cleanupSession(chatId);
  sessions[chatId] = { type: "nodeposited", processing: false };
  bot.sendMessage(
    chatId,
    `✅ *Non-Deposited report session started!*\n\n` +
    `Upload these *2 files* in any order:\n\n` +
    `⬜ \`MTF Indicator Access (Responses).xlsx\`\n` +
    `⬜ \`NON DEPOSITED LIST.xlsx\`\n\n` +
    `⚠️ Files must be named exactly as shown above.`,
    { parse_mode: "Markdown" }
  );
});

bot.onText(/\/inactiveusers/, (msg) => {
  const chatId = msg.chat.id;
  cleanupSession(chatId);
  sessions[chatId] = { type: "inactiveusers", processing: false };
  bot.sendMessage(
    chatId,
    `✅ *Inactive Users report session started!*\n\n` +
    `Upload these *2 files* in any order:\n\n` +
    `⬜ \`MTF Indicator Access (Responses).xlsx\`\n` +
    `⬜ \`DP BUT INACTIVE.xlsx\`\n\n` +
    `⚠️ Files must be named exactly as shown above.`,
    { parse_mode: "Markdown" }
  );
});

// ── Document handler ──────────────────────────────────────────────────────────
bot.on("document", async (msg) => {
  const chatId = msg.chat.id;

  const s = sessions[chatId];
  if (!s) {
    bot.sendMessage(chatId, "⚠️ No active session. Send /verify, /duplicates, /oldonly, /nodeposited, or /inactiveusers first.");
    return;
  }

  if (s.processing) return;

  const doc = msg.document!;
  const fileName = (doc.file_name || "").trim();
  const ext = path.extname(fileName).toLowerCase();

  if (![".csv", ".xlsx", ".xls"].includes(ext)) {
    bot.sendMessage(chatId,
      `❌ *${fileName}* — only CSV or XLSX files are accepted.`,
      { parse_mode: "Markdown" }
    );
    return;
  }

  // ── Validate filename based on session type ───────────────────────────────
  let slot: string | null = null;

  if (s.type === "verify") {
    slot = resolveVerifySlot(fileName);
  } else if (s.type === "nodeposited") {
    slot = resolveNonDepositedSlot(fileName);
  } else if (s.type === "inactiveusers") {
    slot = resolveInactiveUsersSlot(fileName);
  } else {
    // duplicates and oldonly both use old + new
    slot = resolveDuplicateSlot(fileName);
  }

  if (!slot) {
    const expectedFiles =
      s.type === "verify"
        ? "`xm.csv` · `xilion.csv` · `mtf.csv`"
        : s.type === "nodeposited"
          ? "`MTF Indicator Access (Responses).xlsx` · `NON DEPOSITED LIST.xlsx`"
          : s.type === "inactiveusers"
            ? "`MTF Indicator Access (Responses).xlsx` · `DP BUT INACTIVE.xlsx`"
          : "`old.csv` · `new.csv`";
    bot.sendMessage(
      chatId,
      `❌ *${fileName}* is not a recognised file name.\n\n` +
      `Please rename your files to exactly:\n${expectedFiles}`,
      { parse_mode: "Markdown" }
    );
    return;
  }

  // ── Reject duplicate slot ─────────────────────────────────────────────────
  const slotKey = slot as keyof Session;
  if (s[slotKey]) {
    bot.sendMessage(
      chatId,
      `⚠️ *${fileName}* was already uploaded.\n` +
      `Use /reset to start over if you want to replace it.`,
      { parse_mode: "Markdown" }
    );
    return;
  }

  // ── Download file ─────────────────────────────────────────────────────────
  const statusMsg = await bot.sendMessage(chatId, `⏳ Receiving *${fileName}*…`, { parse_mode: "Markdown" });

  try {
    const fileLink = await bot.getFileLink(doc.file_id);
    const savedPath = path.join(UPLOAD_DIR, `${chatId}_${slot}_${Date.now()}${ext}`);
    const response = await fetch(fileLink);
    fs.writeFileSync(savedPath, Buffer.from(await response.arrayBuffer()));

    if (
      slot === "old" ||
      slot === "new" ||
      slot === "xm" ||
      slot === "xilion" ||
      slot === "mtf" ||
      slot === "master" ||
      slot === "nonDeposited" ||
      slot === "inactive"
    ) {
      (s as unknown as Record<string, UploadedFile>)[slot] = { originalName: fileName, savedPath };
    }

    if (s.type === "verify") {
      const remaining = VERIFY_NAMES.filter((k) => !s[k as keyof Session]);
      const remainingList = remaining.map((r) => `\`${r}.csv\``).join(" · ");

      await bot.editMessageText(
        `✅ *${fileName}* received!\n\n` +
        `*Progress:*\n${verifySessionSummary(s)}\n\n` +
        (remaining.length > 0
          ? `Still needed: ${remainingList}`
          : `🚀 All files received! Starting verification…`),
        { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
      );

      if (isVerifyComplete(s) && !s.processing) {
        s.processing = true;
        await runVerification(chatId, s);
      }

    } else if (s.type === "nodeposited") {
      const remaining = NON_DEPOSITED_NAMES.filter((k) => !s[k as keyof Session]);
      const remainingList = remaining
        .map((r) => r === "master" ? `\`MTF Indicator Access (Responses).xlsx\`` : `\`NON DEPOSITED LIST.xlsx\``)
        .join(" · ");

      await bot.editMessageText(
        `✅ *${fileName}* received!\n\n` +
        `*Progress:*\n${nonDepositedSessionSummary(s)}\n\n` +
        (remaining.length > 0
          ? `Still needed: ${remainingList}`
          : `🚀 All files received! Starting non-deposited report…`),
        { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
      );

      if (isNonDepositedComplete(s) && !s.processing) {
        s.processing = true;
        await runNonDeposited(chatId, s);
      }

    } else if (s.type === "inactiveusers") {
      const remaining = INACTIVE_USERS_NAMES.filter((k) => !s[k as keyof Session]);
      const remainingList = remaining
        .map((r) => r === "master" ? `\`MTF Indicator Access (Responses).xlsx\`` : `\`DP BUT INACTIVE.xlsx\``)
        .join(" · ");

      await bot.editMessageText(
        `✅ *${fileName}* received!\n\n` +
        `*Progress:*\n${inactiveUsersSessionSummary(s)}\n\n` +
        (remaining.length > 0
          ? `Still needed: ${remainingList}`
          : `🚀 All files received! Starting inactive users report…`),
        { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
      );

      if (isInactiveUsersComplete(s) && !s.processing) {
        s.processing = true;
        await runInactiveUsers(chatId, s);
      }

    } else {
      // duplicates or oldonly
      const remaining = DUPLICATE_NAMES.filter((k) => !s[k as keyof Session]);
      const remainingList = remaining.map((r) => `\`${r}.csv\``).join(" · ");
      const actionLabel = s.type === "oldonly" ? "old only extraction" : "duplicate check";

      await bot.editMessageText(
        `✅ *${fileName}* received!\n\n` +
        `*Progress:*\n${duplicateSessionSummary(s)}\n\n` +
        (remaining.length > 0
          ? `Still needed: ${remainingList}`
          : `🚀 All files received! Starting ${actionLabel}…`),
        { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
      );

      if (isDuplicateComplete(s) && !s.processing) {
        s.processing = true;
        if (s.type === "oldonly") {
          await runOldOnly(chatId, s);
        } else {
          await runDuplicates(chatId, s);
        }
      }
    }

  } catch (err: any) {
    bot.editMessageText(
      `❌ Failed to receive *${fileName}*: ${err.message}`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
    );
  }
});

// ── Verification runner ───────────────────────────────────────────────────────
async function runVerification(chatId: number, s: Session) {
  try {
    const result = await processVerification({
      xmPath: s.xm!.savedPath,
      xilionPath: s.xilion!.savedPath,
      mtfPath: s.mtf!.savedPath,
    });

    await bot.sendMessage(
      chatId,
      `✅ *Verification Complete!*\n\n` +
      `📊 Total records: *${result.total}*\n` +
      `✅ Verified (Yes): *${result.verified}*\n` +
      `❌ Not Verified (No): *${result.notVerified}*\n\n` +
      `Sending your output file…`,
      { parse_mode: "Markdown" }
    );

    await bot.sendDocument(chatId, result.outputPath, {
      caption: `Verified output — ${new Date().toLocaleDateString("en-IN")}`,
    });

    try { fs.unlinkSync(result.outputPath); } catch {}
    cleanupSession(chatId);

  } catch (err: any) {
    s.processing = false;
    bot.sendMessage(
      chatId,
      `❌ *Verification failed:* ${err.message}\n\nCheck your files and use /reset to try again.`,
      { parse_mode: "Markdown" }
    );
  }
}

// ── Duplicate check runner ────────────────────────────────────────────────────
async function runDuplicates(chatId: number, s: Session) {
  try {
    const result = await processDuplicates({
      oldPath: s.old!.savedPath,
      newPath: s.new!.savedPath,
    });

    await bot.sendMessage(
      chatId,
      `✅ *Comparison Complete!*\n\n` +
      `📊 Old records: *${result.oldCount}*\n` +
      `📊 New records: *${result.newCount}*\n\n` +
      `*Results:*\n` +
      `🔶 Old Only: *${result.oldOnly}*\n` +
      `🟢 New Only: *${result.newOnly}*\n` +
      `🔵 Both (Duplicates): *${result.both}*\n\n` +
      `Sending your output file…`,
      { parse_mode: "Markdown" }
    );

    await bot.sendDocument(chatId, result.outputPath, {
      caption: `Comparison output — ${new Date().toLocaleDateString("en-IN")}`,
    });

    try { fs.unlinkSync(result.outputPath); } catch {}
    cleanupSession(chatId);

  } catch (err: any) {
    s.processing = false;
    bot.sendMessage(
      chatId,
      `❌ *Comparison failed:* ${err.message}\n\nCheck your files and use /reset to try again.`,
      { parse_mode: "Markdown" }
    );
  }
}

// ── Old Only runner ───────────────────────────────────────────────────────────
async function runOldOnly(chatId: number, s: Session) {
  try {
    const result = await processOldOnly({
      oldPath: s.old!.savedPath,
      newPath: s.new!.savedPath,
    });

    await bot.sendMessage(
      chatId,
      `✅ *Old Only Extraction Complete!*\n\n` +
      `📊 Old records: *${result.oldCount}*\n` +
      `📊 New records: *${result.newCount}*\n\n` +
      `🔶 Records in Old but NOT in New: *${result.oldOnlyCount}*\n\n` +
      `Sending your output file…`,
      { parse_mode: "Markdown" }
    );

    await bot.sendDocument(chatId, result.outputPath, {
      caption: `Old Only output — ${new Date().toLocaleDateString("en-IN")}`,
    });

    try { fs.unlinkSync(result.outputPath); } catch {}
    cleanupSession(chatId);

  } catch (err: any) {
    s.processing = false;
    bot.sendMessage(
      chatId,
      `❌ *Old Only extraction failed:* ${err.message}\n\nCheck your files and use /reset to try again.`,
      { parse_mode: "Markdown" }
    );
  }
}

// ── Static report runners ────────────────────────────────────────────────────
async function runNonDeposited(chatId: number, s: Session) {
  const statusMsg = await bot.sendMessage(chatId, "⏳ Building non-deposited verified users report…");

  try {
    const result = await processNonDepositedReport({
      masterPath: s.master!.savedPath,
      nonDepositedPath: s.nonDeposited!.savedPath,
    });

    await bot.editMessageText(
      `✅ *Non-Deposited Report Complete!*\n\n` +
      `📊 Verified users: *${result.verifiedCount}*\n` +
      `📊 Non-deposited records: *${result.sourceCount}*\n` +
      `✅ Matches found: *${result.matchedCount}*\n\n` +
      `Sending your output file…`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
    );

    await bot.sendDocument(chatId, result.outputPath, {
      caption: `Non-deposited verified users — ${new Date().toLocaleDateString("en-IN")}`,
    });

    try { fs.unlinkSync(result.outputPath); } catch {}
    cleanupSession(chatId);

  } catch (err: any) {
    s.processing = false;
    await bot.editMessageText(
      `❌ *Non-deposited report failed:* ${err.message}`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
    );
  }
}

async function runInactiveUsers(chatId: number, s: Session) {
  const statusMsg = await bot.sendMessage(chatId, "⏳ Building inactive verified users report…");

  try {
    const result = await processInactiveUsersReport({
      masterPath: s.master!.savedPath,
      inactivePath: s.inactive!.savedPath,
    });

    await bot.editMessageText(
      `✅ *Inactive Users Report Complete!*\n\n` +
      `📊 Verified users: *${result.verifiedCount}*\n` +
      `📊 Inactive records: *${result.sourceCount}*\n` +
      `✅ Matches found: *${result.matchedCount}*\n\n` +
      `Sending your output file…`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
    );

    await bot.sendDocument(chatId, result.outputPath, {
      caption: `Inactive verified users — ${new Date().toLocaleDateString("en-IN")}`,
    });

    try { fs.unlinkSync(result.outputPath); } catch {}
    cleanupSession(chatId);

  } catch (err: any) {
    s.processing = false;
    await bot.editMessageText(
      `❌ *Inactive users report failed:* ${err.message}`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
    );
  }
}

console.log("🤖 Verification bot is running…");
