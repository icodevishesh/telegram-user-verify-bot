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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
require("dotenv/config");
const node_telegram_bot_api_1 = __importDefault(require("node-telegram-bot-api"));
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const verifier_1 = require("./verifier");
const TOKEN = process.env.TELEGRAM_BOT_TOKEN;
if (!TOKEN) {
    console.error("❌ TELEGRAM_BOT_TOKEN is not set in .env");
    process.exit(1);
}
const bot = new node_telegram_bot_api_1.default(TOKEN, { polling: true });
const sessions = {};
// ── Accepted filenames ────────────────────────────────────────────────────────
const VERIFY_NAMES = ["xm", "xilion", "mtf"];
const DUPLICATE_NAMES = ["old", "new"];
const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR))
    fs.mkdirSync(UPLOAD_DIR, { recursive: true });
// ── Helpers ───────────────────────────────────────────────────────────────────
function resolveVerifySlot(fileName) {
    const base = path.basename(fileName, path.extname(fileName)).toLowerCase().trim();
    if (VERIFY_NAMES.includes(base))
        return base;
    return null;
}
function resolveDuplicateSlot(fileName) {
    const base = path.basename(fileName, path.extname(fileName)).toLowerCase().trim();
    if (DUPLICATE_NAMES.includes(base))
        return base;
    return null;
}
function verifySessionSummary(s) {
    return [
        s.xm ? `✅ xm.csv` : `⬜ xm.csv`,
        s.xilion ? `✅ xilion.csv` : `⬜ xilion.csv`,
        s.mtf ? `✅ mtf.csv` : `⬜ mtf.csv`,
    ].join("\n");
}
function duplicateSessionSummary(s) {
    return [
        s.old ? `✅ old.csv` : `⬜ old.csv`,
        s.new ? `✅ new.csv` : `⬜ new.csv`,
    ].join("\n");
}
function isVerifyComplete(s) {
    return !!(s.xm && s.xilion && s.mtf);
}
function isDuplicateComplete(s) {
    return !!(s.old && s.new);
}
function cleanupSession(chatId) {
    const s = sessions[chatId];
    if (!s)
        return;
    const keys = s.type === "verify" ? VERIFY_NAMES : DUPLICATE_NAMES;
    for (const key of keys) {
        const f = s[key];
        if (f?.savedPath && fs.existsSync(f.savedPath)) {
            try {
                fs.unlinkSync(f.savedPath);
            }
            catch { }
        }
    }
    delete sessions[chatId];
}
// ── Commands ──────────────────────────────────────────────────────────────────
bot.onText(/\/start/, (msg) => {
    bot.sendMessage(msg.chat.id, `👋 *Welcome to the MTF Verification Bot!*\n\n` +
        `Matches user IDs from your records against XM and Xilion broker sheets.\n\n` +
        `*Commands:*\n` +
        `/verify — Start a new verification session\n` +
        `/duplicates — Compare old & new CSV sheets for duplicates\n` +
        `/oldonly — Extract records present in Old but not in New\n` +
        `/status — Check which files have been uploaded\n` +
        `/reset  — Clear the current session\n` +
        `/help   — Show usage instructions`, { parse_mode: "Markdown" });
});
bot.onText(/\/help/, (msg) => {
    bot.sendMessage(msg.chat.id, `*How to use:*\n\n` +
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
        `⚠️ Files must be named *exactly* as shown above\n\n` +
        `*Matching logic:*\n` +
        `• XM → \`MT4/MT5 ID\` and \`Client ID\` columns\n` +
        `• Xilion → \`Wallet\` column (e.g. #316393)\n` +
        `• Duplicates → Matches on \`Email\` OR \`Phone\` (either field is enough)\n` +
        `• Old Only → Matches on \`Phone\` only`, { parse_mode: "Markdown" });
});
bot.onText(/\/reset/, (msg) => {
    cleanupSession(msg.chat.id);
    bot.sendMessage(msg.chat.id, "🔄 Session cleared. Send /verify, /duplicates, or /oldonly to start again.");
});
bot.onText(/\/status/, (msg) => {
    const chatId = msg.chat.id;
    const s = sessions[chatId];
    if (!s) {
        bot.sendMessage(chatId, "No active session. Send /verify, /duplicates, or /oldonly to begin.");
        return;
    }
    if (s.type === "verify") {
        const extra = s.processing
            ? `⚙️ Verification is currently running…`
            : isVerifyComplete(s) ? `✅ All files received.` : `Upload the remaining files to continue.`;
        bot.sendMessage(chatId, `*Upload Status:*\n\n${verifySessionSummary(s)}\n\n${extra}`, { parse_mode: "Markdown" });
    }
    else {
        // duplicates or oldonly — both use old + new
        const extra = s.processing
            ? `⚙️ Processing is currently running…`
            : isDuplicateComplete(s) ? `✅ All files received.` : `Upload the remaining files to continue.`;
        bot.sendMessage(chatId, `*Upload Status:*\n\n${duplicateSessionSummary(s)}\n\n${extra}`, { parse_mode: "Markdown" });
    }
});
bot.onText(/\/verify/, (msg) => {
    const chatId = msg.chat.id;
    cleanupSession(chatId);
    sessions[chatId] = { type: "verify", processing: false };
    bot.sendMessage(chatId, `✅ *Verification session started!*\n\n` +
        `Upload these *3 files* in any order:\n\n` +
        `⬜ \`xm.csv\`\n` +
        `⬜ \`xilion.csv\`\n` +
        `⬜ \`mtf.csv\`\n\n` +
        `⚠️ Files must be named exactly as shown above.`, { parse_mode: "Markdown" });
});
bot.onText(/\/duplicates/, (msg) => {
    const chatId = msg.chat.id;
    cleanupSession(chatId);
    sessions[chatId] = { type: "duplicates", processing: false };
    bot.sendMessage(chatId, `✅ *Duplicate check session started!*\n\n` +
        `Upload these *2 files* in any order:\n\n` +
        `⬜ \`old.csv\` — Old records\n` +
        `⬜ \`new.csv\` — New records\n\n` +
        `⚠️ Files must be named exactly as shown above.\n\n` +
        `📋 Matches on: *Email* OR *Phone* (either field is enough)`, { parse_mode: "Markdown" });
});
bot.onText(/\/oldonly/, (msg) => {
    const chatId = msg.chat.id;
    cleanupSession(chatId);
    sessions[chatId] = { type: "oldonly", processing: false };
    bot.sendMessage(chatId, `✅ *Old Only session started!*\n\n` +
        `This will extract records from *old.csv* that are *NOT* present in *new.csv*.\n\n` +
        `Upload these *2 files* in any order:\n\n` +
        `⬜ \`old.csv\` — Old records\n` +
        `⬜ \`new.csv\` — New records\n\n` +
        `⚠️ Files must be named exactly as shown above.\n\n` +
        `📋 Matches on: *Phone Number* only`, { parse_mode: "Markdown" });
});
// ── Document handler ──────────────────────────────────────────────────────────
bot.on("document", async (msg) => {
    const chatId = msg.chat.id;
    const s = sessions[chatId];
    if (!s) {
        bot.sendMessage(chatId, "⚠️ No active session. Send /verify, /duplicates, or /oldonly first.");
        return;
    }
    if (s.processing)
        return;
    const doc = msg.document;
    const fileName = (doc.file_name || "").trim();
    const ext = path.extname(fileName).toLowerCase();
    if (![".csv", ".xlsx", ".xls"].includes(ext)) {
        bot.sendMessage(chatId, `❌ *${fileName}* — only CSV or XLSX files are accepted.`, { parse_mode: "Markdown" });
        return;
    }
    // ── Validate filename based on session type ───────────────────────────────
    let slot = null;
    if (s.type === "verify") {
        slot = resolveVerifySlot(fileName);
    }
    else {
        // duplicates and oldonly both use old + new
        slot = resolveDuplicateSlot(fileName);
    }
    if (!slot) {
        const expectedFiles = s.type === "verify"
            ? "`xm.csv` · `xilion.csv` · `mtf.csv`"
            : "`old.csv` · `new.csv`";
        bot.sendMessage(chatId, `❌ *${fileName}* is not a recognised file name.\n\n` +
            `Please rename your files to exactly:\n${expectedFiles}`, { parse_mode: "Markdown" });
        return;
    }
    // ── Reject duplicate slot ─────────────────────────────────────────────────
    const slotKey = slot;
    if (s[slotKey]) {
        bot.sendMessage(chatId, `⚠️ *${fileName}* was already uploaded.\n` +
            `Use /reset to start over if you want to replace it.`, { parse_mode: "Markdown" });
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
            s[slot] = { originalName: fileName, savedPath };
        }
        if (s.type === "verify") {
            const remaining = VERIFY_NAMES.filter((k) => !s[k]);
            const remainingList = remaining.map((r) => `\`${r}.csv\``).join(" · ");
            await bot.editMessageText(`✅ *${fileName}* received!\n\n` +
                `*Progress:*\n${verifySessionSummary(s)}\n\n` +
                (remaining.length > 0
                    ? `Still needed: ${remainingList}`
                    : `🚀 All files received! Starting verification…`), { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" });
            if (isVerifyComplete(s) && !s.processing) {
                s.processing = true;
                await runVerification(chatId, s);
            }
        }
        else {
            // duplicates or oldonly
            const remaining = DUPLICATE_NAMES.filter((k) => !s[k]);
            const remainingList = remaining.map((r) => `\`${r}.csv\``).join(" · ");
            const actionLabel = s.type === "oldonly" ? "old only extraction" : "duplicate check";
            await bot.editMessageText(`✅ *${fileName}* received!\n\n` +
                `*Progress:*\n${duplicateSessionSummary(s)}\n\n` +
                (remaining.length > 0
                    ? `Still needed: ${remainingList}`
                    : `🚀 All files received! Starting ${actionLabel}…`), { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" });
            if (isDuplicateComplete(s) && !s.processing) {
                s.processing = true;
                if (s.type === "oldonly") {
                    await runOldOnly(chatId, s);
                }
                else {
                    await runDuplicates(chatId, s);
                }
            }
        }
    }
    catch (err) {
        bot.editMessageText(`❌ Failed to receive *${fileName}*: ${err.message}`, { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: "Markdown" });
    }
});
// ── Verification runner ───────────────────────────────────────────────────────
async function runVerification(chatId, s) {
    try {
        const result = await (0, verifier_1.processVerification)({
            xmPath: s.xm.savedPath,
            xilionPath: s.xilion.savedPath,
            mtfPath: s.mtf.savedPath,
        });
        await bot.sendMessage(chatId, `✅ *Verification Complete!*\n\n` +
            `📊 Total records: *${result.total}*\n` +
            `✅ Verified (Yes): *${result.verified}*\n` +
            `❌ Not Verified (No): *${result.notVerified}*\n\n` +
            `Sending your output file…`, { parse_mode: "Markdown" });
        await bot.sendDocument(chatId, result.outputPath, {
            caption: `Verified output — ${new Date().toLocaleDateString("en-IN")}`,
        });
        try {
            fs.unlinkSync(result.outputPath);
        }
        catch { }
        cleanupSession(chatId);
    }
    catch (err) {
        s.processing = false;
        bot.sendMessage(chatId, `❌ *Verification failed:* ${err.message}\n\nCheck your files and use /reset to try again.`, { parse_mode: "Markdown" });
    }
}
// ── Duplicate check runner ────────────────────────────────────────────────────
async function runDuplicates(chatId, s) {
    try {
        const result = await (0, verifier_1.processDuplicates)({
            oldPath: s.old.savedPath,
            newPath: s.new.savedPath,
        });
        await bot.sendMessage(chatId, `✅ *Comparison Complete!*\n\n` +
            `📊 Old records: *${result.oldCount}*\n` +
            `📊 New records: *${result.newCount}*\n\n` +
            `*Results:*\n` +
            `🔶 Old Only: *${result.oldOnly}*\n` +
            `🟢 New Only: *${result.newOnly}*\n` +
            `🔵 Both (Duplicates): *${result.both}*\n\n` +
            `Sending your output file…`, { parse_mode: "Markdown" });
        await bot.sendDocument(chatId, result.outputPath, {
            caption: `Comparison output — ${new Date().toLocaleDateString("en-IN")}`,
        });
        try {
            fs.unlinkSync(result.outputPath);
        }
        catch { }
        cleanupSession(chatId);
    }
    catch (err) {
        s.processing = false;
        bot.sendMessage(chatId, `❌ *Comparison failed:* ${err.message}\n\nCheck your files and use /reset to try again.`, { parse_mode: "Markdown" });
    }
}
// ── Old Only runner ───────────────────────────────────────────────────────────
async function runOldOnly(chatId, s) {
    try {
        const result = await (0, verifier_1.processOldOnly)({
            oldPath: s.old.savedPath,
            newPath: s.new.savedPath,
        });
        await bot.sendMessage(chatId, `✅ *Old Only Extraction Complete!*\n\n` +
            `📊 Old records: *${result.oldCount}*\n` +
            `📊 New records: *${result.newCount}*\n\n` +
            `🔶 Records in Old but NOT in New: *${result.oldOnlyCount}*\n\n` +
            `Sending your output file…`, { parse_mode: "Markdown" });
        await bot.sendDocument(chatId, result.outputPath, {
            caption: `Old Only output — ${new Date().toLocaleDateString("en-IN")}`,
        });
        try {
            fs.unlinkSync(result.outputPath);
        }
        catch { }
        cleanupSession(chatId);
    }
    catch (err) {
        s.processing = false;
        bot.sendMessage(chatId, `❌ *Old Only extraction failed:* ${err.message}\n\nCheck your files and use /reset to try again.`, { parse_mode: "Markdown" });
    }
}
console.log("🤖 Verification bot is running…");
