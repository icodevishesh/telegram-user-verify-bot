import "dotenv/config";
import TelegramBot from "node-telegram-bot-api";
import * as fs from "fs";
import * as path from "path";
import { processVerification, processDuplicates } from "./verifier";

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

type SessionType = "verify" | "duplicates";

interface Session {
  type: SessionType;
  xm?: UploadedFile;
  xilion?: UploadedFile;
  mtf?: UploadedFile;
  old?: UploadedFile;
  new?: UploadedFile;
  processing: boolean;   // ← lock: prevents duplicate verification runs
}

const sessions: Record<number, Session> = {};

// ── Accepted filenames ────────────────────────────────────────────────────────
const VERIFY_NAMES = ["xm", "xilion", "mtf"] as const;
type VerifyName = (typeof VERIFY_NAMES)[number];

const DUPLICATE_NAMES = ["old", "new"] as const;
type DuplicateName = (typeof DUPLICATE_NAMES)[number];

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

function isVerifyComplete(s: Session): boolean {
  return !!(s.xm && s.xilion && s.mtf);
}

function isDuplicateComplete(s: Session): boolean {
  return !!(s.old && s.new);
}

function cleanupSession(chatId: number) {
  const s = sessions[chatId];
  if (!s) return;
  const keys = s.type === "verify" ? VERIFY_NAMES : DUPLICATE_NAMES;
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
    `4️⃣ Bot sends back an XLSX with a *Duplicate* column\n\n` +
    `⚠️ Files must be named *exactly* as shown above\n\n` +
    `*Matching logic:*\n` +
    `• XM → \`MT4/MT5 ID\` and \`Client ID\` columns\n` +
    `• Xilion → \`Wallet\` column (e.g. #316393)\n` +
    `• Duplicates → Matches on \`Name\`, \`Email Id\`, \`Contact Number\``,
    { parse_mode: "Markdown" }
  );
});

bot.onText(/\/reset/, (msg) => {
  cleanupSession(msg.chat.id);
  bot.sendMessage(msg.chat.id, "🔄 Session cleared. Send /verify to start again.");
});

bot.onText(/\/status/, (msg) => {
  const chatId = msg.chat.id;
  const s = sessions[chatId];
  if (!s) {
    bot.sendMessage(chatId, "No active session. Send /verify or /duplicates to begin.");
    return;
  }

  if (s.type === "verify") {
    const extra = s.processing
      ? `⚙️ Verification is currently running…`
      : isVerifyComplete(s)
      ? `✅ All files received.`
      : `Upload the remaining files to continue.`;
    bot.sendMessage(
      chatId,
      `*Upload Status:*\n\n${verifySessionSummary(s)}\n\n${extra}`,
      { parse_mode: "Markdown" }
    );
  } else {
    const extra = s.processing
      ? `⚙️ Duplicate check is currently running…`
      : isDuplicateComplete(s)
      ? `✅ All files received.`
      : `Upload the remaining files to continue.`;
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
    `📋 Columns checked: *Name*, *Email Id*, *Contact Number*`,
    { parse_mode: "Markdown" }
  );
});

// ── Document handler ──────────────────────────────────────────────────────────
bot.on("document", async (msg) => {
  const chatId = msg.chat.id;

  const s = sessions[chatId];
  if (!s) {
    bot.sendMessage(chatId, "⚠️ No active session. Send /verify or /duplicates first.");
    return;
  }

  // If already processing, ignore any stray events
  if (s.processing) return;

  const doc = msg.document!;
  const fileName = (doc.file_name || "").trim();
  const ext = path.extname(fileName).toLowerCase();

  // ── Validate extension ────────────────────────────────────────────────────
  if (![".csv", ".xlsx", ".xls"].includes(ext)) {
    bot.sendMessage(chatId,
      `❌ *${fileName}* — only CSV or XLSX files are accepted.`,
      { parse_mode: "Markdown" }
    );
    return;
  }

  // ── Validate filename based on session type ────────────────────────────────
  let slot: string | null = null;
  let acceptedNames: readonly string[] = [];

  if (s.type === "verify") {
    slot = resolveVerifySlot(fileName);
    acceptedNames = VERIFY_NAMES;
  } else {
    slot = resolveDuplicateSlot(fileName);
    acceptedNames = DUPLICATE_NAMES;
  }

  if (!slot) {
    const expectedFiles = s.type === "verify"
      ? "`xm.csv` · `xilion.csv` · `mtf.csv`"
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

    if (slot === "old" || slot === "new" || slot === "xm" || slot === "xilion" || slot === "mtf") {
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

      // ── All 3 ready → lock and run (only once) ────────────────────────────
      if (isVerifyComplete(s) && !s.processing) {
        s.processing = true;
        await runVerification(chatId, s);
      }
    } else {
      const remaining = DUPLICATE_NAMES.filter((k) => !s[k as keyof Session]);
      const remainingList = remaining.map((r) => `\`${r}.csv\``).join(" · ");

      await bot.editMessageText(
        `✅ *${fileName}* received!\n\n` +
        `*Progress:*\n${duplicateSessionSummary(s)}\n\n` +
        (remaining.length > 0
          ? `Still needed: ${remainingList}`
          : `🚀 All files received! Starting duplicate check…`),
        { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" }
      );

      // ── All 2 ready → lock and run (only once) ────────────────────────────
      if (isDuplicateComplete(s) && !s.processing) {
        s.processing = true;
        await runDuplicates(chatId, s);
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
    s.processing = false;   // release lock on error so user can /reset and retry
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
      `*Total Records:*\n` +
      `📊 Old records: *${result.oldCount}*\n` +
      `📊 New records: *${result.newCount}*\n\n` +
      `*Results:*\n` +
      `� Old Only: *${result.oldOnly}*\n` +
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
    s.processing = false;   // release lock on error so user can /reset and retry
    bot.sendMessage(
      chatId,
      `❌ *Comparison failed:* ${err.message}\n\nCheck your files and use /reset to try again.`,
      { parse_mode: "Markdown" }
    );
  }
}

console.log("🤖 Verification bot is running…");
