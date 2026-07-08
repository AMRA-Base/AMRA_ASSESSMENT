// ─── AMRA Assessment — Apps Script Backend ───────────────────────────────────
// Deploy: Extensions → Apps Script → Deploy → New Deployment
//   Type: Web App | Execute as: Me | Who has access: Anyone
//
// One-time setup: Extensions → Apps Script → Project Settings → Script Properties
//   Key: ANSWER_KEY_FILE_ID   Value: Drive file ID of level1_answers.json
//
// Sheets written: "Assessments" (summary) and "Assessment_Details" (full record)
// Both tabs are auto-created with headers on first run.

const ASSESSMENTS_TAB = "Assessments";
const DETAILS_TAB     = "Assessment_Details";
const PASS_THRESHOLD  = 0.70;

// ── Schema definitions ────────────────────────────────────────────────────────

const ASSESSMENTS_HEADERS = [
  "ref_id", "full_name", "email", "age", "trading_experience",
  "level", "score", "percentage", "time_taken_minutes",
  "mode", "session_id", "submitted_at"
];

const DETAILS_HEADERS = [
  "ref_id", "full_name", "email", "age", "trading_experience",
  "firm_name", "number_of_dependents", "level", "score", "percentage",
  "beginner_correct", "intermediate_correct", "session_id",
  "time_taken_minutes", "mode", "submitted_at", "recorded_at",
  "answers_json", "observations", "tab_switch_count", "fullscreen_exit_count"
];

const SUSPICIOUS_HEADERS = [
  "timestamp", "ref_id", "full_name", "email", "session_id",
  "activity_type", "details", "reviewed", "review_notes"
];

const TOKENS_HEADERS = [
  "token", "full_name", "email", "level", "created_at",
  "expires_at", "used", "used_at", "session_id", "candidate_url", "created_by", "remarks"
];

const CONFIG_HEADERS = ["key", "value", "description", "last_updated"];

const CONFIG_SEED = [
  ["PASS_THRESHOLD",     "70",    "Minimum percentage to pass"],
  ["LEVEL1_QUESTIONS",   "25",    "Total questions per L1 session"],
  ["TOKEN_REQUIRED",     "false", "Whether token gate is active"],
  ["TOKEN_EXPIRY_HOURS", "48",    "Token validity in hours"],
  ["MAX_ATTEMPTS_EMAIL", "1",     "Max attempts per email address"],
  ["CURRENT_LEVEL",      "1",     "Active assessment level"],
];

// ── Entry point ───────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // Token validation action (called by frontend before loading session JSON)
    if (payload.action === "validate_token") {
      return validateTokenAction(payload.token);
    }

    if (!payload.session_id || !Array.isArray(payload.user_answers)) {
      return respond({ error: "Invalid payload: missing session_id or user_answers" });
    }

    const answerKey = loadAnswerKey();
    const result    = scoreAnswers(payload.user_answers, answerKey);
    const scorePct  = result.score / 25;
    const passFail  = scorePct >= PASS_THRESHOLD ? "pass" : "fail";

    writeAssessments(payload, result, scorePct, passFail);
    writeDetails(payload, result, scorePct, passFail);
    checkMonitoringThresholds(payload);

    if (payload.token) {
      markTokenUsed(payload.token);
    }

    return respond({
      referenceId:    payload.referenceId,
      score:          result.score,
      totalQuestions: 25,
      score_pct:      scorePct,
      pass_fail:      passFail
    });

  } catch (err) {
    console.error("doPost error:", err.message, err.stack);
    return respond({ error: err.message });
  }
}

// ── Token validation ──────────────────────────────────────────────────────────
// Tokens tab columns (0-indexed): token(0), full_name(1), email(2), level(3),
//   created_at(4), expires_at(5), used(6), used_at(7), session_id(8), ...

function validateTokenAction(token) {
  if (!token) return respond({ error: "Token is required" });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tokens");
  if (!sheet) return respond({ error: "Tokens tab not found" });

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).trim() !== String(token).trim()) continue;

    const rowSessionId = String(row[8]).trim();

    if (row[6] === true || String(row[6]).toLowerCase() === "true") {
      logSuspiciousActivity("used_token_attempt", `Token reuse attempted: ${token}`, rowSessionId);
      return respond({ error: "This token has already been used" });
    }

    const expiresAt = new Date(row[5]);
    if (!row[5] || isNaN(expiresAt.getTime()) || expiresAt < new Date()) {
      logSuspiciousActivity("expired_token", `Expired token presented: ${token}`, rowSessionId);
      return respond({ error: "This token has expired" });
    }

    if (!rowSessionId) return respond({ error: "Token has no session assigned" });

    return respond({ valid: true, session_id: rowSessionId });
  }

  logSuspiciousActivity("invalid_token", `Unknown token presented: ${token}`, "");
  return respond({ error: "Invalid token" });
}

function markTokenUsed(token) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tokens");
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(token).trim()) continue;
    sheet.getRange(i + 1, 7).setValue(true);                     // used
    sheet.getRange(i + 1, 8).setValue(new Date().toISOString()); // used_at
    return;
  }
}

// context = { refId, fullName, email } — optional; omit for token-validation events
//            where candidate identity is not yet known.
function logSuspiciousActivity(activityType, details, sessionId, context) {
  try {
    const sheet = getOrCreateSheet("SuspiciousActivity", SUSPICIOUS_HEADERS);
    const ctx = context || {};
    sheet.appendRow([
      new Date().toISOString(),
      ctx.refId    || "",
      ctx.fullName || "",
      ctx.email    || "",
      sessionId    || "",
      activityType,
      details,
      false,
      "",
    ]);
  } catch (err) {
    console.error("logSuspiciousActivity error:", err.message);
  }
}

function checkMonitoringThresholds(payload) {
  const tabSwitches = payload.tabSwitchCount      || 0;
  const fsExits     = payload.fullscreenExitCount || 0;
  if (tabSwitches === 0 && fsExits === 0) return;

  const ctx = { refId: payload.referenceId, fullName: payload.name, email: payload.email };

  if (tabSwitches >= 3) {
    logSuspiciousActivity(
      "excessive_tab_switching",
      `Tab switches during exam: ${tabSwitches}`,
      payload.session_id,
      ctx
    );
  }
  if (fsExits >= 2) {
    logSuspiciousActivity(
      "excessive_fullscreen_exit",
      `Fullscreen exits during exam: ${fsExits}`,
      payload.session_id,
      ctx
    );
  }
}

// ── Token admin helper ────────────────────────────────────────────────────────
// Run from Apps Script editor: select generateToken in the dropdown, then click Run.
// Or call from a one-line wrapper: function run() { generateToken("Name","email@x.com",1); }
//
// Columns written (matches TOKENS_HEADERS):
//   token | full_name | email | level | created_at | expires_at | used | used_at
//   session_id | candidate_url | created_by | remarks

// Stable session IDs assigned per level.
// Update when a new session file is generated for a level.
const SESSION_POOL = {
  "1": "69969ee5-bdb5-4ce0-856b-b1f65b29c88f",
  "2": "",
  "3": "",
};

function generateToken(fullName, email, level) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const tokenSheet = getOrCreateSheet("Tokens", TOKENS_HEADERS);
  const levelStr   = String(level);
  const prefix     = `AMRA-L${levelStr}-`;

  // ── Find highest sequence number for this level ───────────────────────────
  let maxSeq = 0;
  if (tokenSheet.getLastRow() > 1) {
    const col1 = tokenSheet.getRange(2, 1, tokenSheet.getLastRow() - 1, 1).getValues();
    for (const [tok] of col1) {
      const s = String(tok);
      if (!s.startsWith(prefix)) continue;
      const n = parseInt(s.slice(prefix.length), 10);
      if (!isNaN(n) && n > maxSeq) maxSeq = n;
    }
  }

  const token = `${prefix}${String(maxSeq + 1).padStart(3, "0")}`;

  // ── Timestamps ────────────────────────────────────────────────────────────
  let expiryHours = 2;
  const configSheet = ss.getSheetByName("Config");
  if (configSheet && configSheet.getLastRow() > 1) {
    const rows = configSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === "TOKEN_EXPIRY_HOURS") {
        expiryHours = parseInt(rows[i][1]) || 2;
        break;
      }
    }
  }

  const now        = new Date();
  const createdAt  = now.toISOString();
  const expiresAt  = new Date(now.getTime() + expiryHours * 3600 * 1000).toISOString();

  const candidateUrl = `https://amra-base.github.io/AMRA_ASSESSMENT/?token=${token}`;

  let createdBy = "admin";
  try { createdBy = Session.getEffectiveUser().getEmail() || "admin"; } catch (_) {}

  tokenSheet.appendRow([
    token,        // token
    fullName,     // full_name
    email,        // email
    levelStr,     // level
    createdAt,    // created_at
    expiresAt,    // expires_at
    false,        // used
    "",           // used_at
    SESSION_POOL[levelStr] || "",  // session_id
    candidateUrl, // candidate_url
    createdBy,    // created_by
    "",           // remarks
  ]);

  Logger.log(`Token : ${token}`);
  Logger.log(`URL   : ${candidateUrl}`);
  Logger.log(`Expiry: ${expiresAt}`);

  return { token, candidateUrl, expiresAt };
}

// ── Answer key ────────────────────────────────────────────────────────────────

function loadAnswerKey() {
  const fileId = PropertiesService.getScriptProperties()
                   .getProperty("ANSWER_KEY_FILE_ID");
  if (!fileId) throw new Error("Script Property ANSWER_KEY_FILE_ID is not set");
  const raw = DriveApp.getFileById(fileId).getBlob().getDataAsString("UTF-8");
  return JSON.parse(raw).keys;   // { "L1-S1-Q001": "correct answer text", ... }
}

// ── Scoring ───────────────────────────────────────────────────────────────────

function scoreAnswers(userAnswers, answerKey) {
  let score = 0, answered = 0, beginnerCorrect = 0, intermediateCorrect = 0;

  for (const item of userAnswers) {
    const ua       = item.user_answer;
    const expected = answerKey[item.stable_id];

    if (ua === null || ua === undefined) continue;
    answered++;

    if (expected === undefined) continue;   // stable_id not in key — skip
    if (ua.trim() !== expected.trim())  continue;

    score++;
    if      (item.difficulty === "beginner")     beginnerCorrect++;
    else if (item.difficulty === "intermediate")  intermediateCorrect++;
  }

  return { score, answered, beginnerCorrect, intermediateCorrect };
}

// ── Sheet: Assessments (summary row) ─────────────────────────────────────────

function writeAssessments(payload, result, scorePct, passFail) {
  const sheet = getOrCreateSheet(ASSESSMENTS_TAB, [
    "ref_id", "full_name", "email", "age", "trading_experience",
    "level", "score", "percentage", "time_taken_minutes",
    "mode", "session_id", "submitted_at"
  ]);

  // mode: "online" = token-generated controlled link; "offline" = shared/common link
  const mode           = payload.mode || (payload.token ? "online" : "offline");
  const timeTakenMin   = payload.timeTaken !== undefined
                           ? parseFloat((payload.timeTaken / 60).toFixed(1))
                           : "";

  sheet.appendRow([
    payload.referenceId,
    payload.name,
    payload.email,
    payload.age,
    payload.tradingExperience,
    1,
    result.score,
    parseFloat((scorePct * 100).toFixed(1)),   // number, not string
    timeTakenMin,
    mode,
    payload.session_id,
    payload.submittedAt
  ]);
}

// ── Sheet: Assessment_Details (full record) ───────────────────────────────────

function writeDetails(payload, result, scorePct, passFail) {
  const sheet = getOrCreateSheet(DETAILS_TAB, DETAILS_HEADERS);

  const mode         = payload.mode || (payload.token ? "online" : "offline");
  const timeTakenMin = payload.timeTaken !== undefined
                         ? parseFloat((payload.timeTaken / 60).toFixed(1))
                         : "";

  sheet.appendRow([
    payload.referenceId,
    payload.name,
    payload.email,
    payload.age,
    payload.tradingExperience,
    payload.firmName,
    payload.dependents,
    1,
    result.score,
    parseFloat((scorePct * 100).toFixed(1)),   // number, not string
    result.beginnerCorrect,
    result.intermediateCorrect,
    payload.session_id,
    timeTakenMin,
    mode,
    payload.submittedAt,
    new Date().toISOString(),                       // recorded_at: server write time
    JSON.stringify(payload.user_answers),           // answers_json: full audit, no correct answers
    "",                                             // observations: admin-populated post-review
    payload.tabSwitchCount      !== undefined ? payload.tabSwitchCount      : "",
    payload.fullscreenExitCount !== undefined ? payload.fullscreenExitCount : ""
  ]);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function getOrCreateSheet(name, headers) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Schema initialization ─────────────────────────────────────────────────────
// Run initializeSheets() ONCE from the Apps Script editor (not via web request).
// Safe to re-run: replaces header row only, preserves all data rows.
// Never touches Legacy_Responses.

function initializeSheets() {
  ensureOperationalTab_(ASSESSMENTS_TAB,    ASSESSMENTS_HEADERS);
  ensureOperationalTab_(DETAILS_TAB,        DETAILS_HEADERS);
  ensureOperationalTab_("SuspiciousActivity", SUSPICIOUS_HEADERS);
  ensureOperationalTab_("Tokens",           TOKENS_HEADERS);
  ensureConfigTab_();
  Logger.log("initializeSheets() complete.");
}

// Creates tab if missing or updates header row if mismatched.
// Data rows (row 2+) are never modified. Freezes row 1. Applies filter.
function ensureOperationalTab_(name, headers) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  // Write/overwrite header row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Clear stale header cells to the right of the new schema
  const lastCol = sheet.getLastColumn();
  if (lastCol > headers.length) {
    sheet.getRange(1, headers.length + 1, 1, lastCol - headers.length).clearContent();
  }

  sheet.setFrozenRows(1);

  // Replace any existing filter with one spanning the full schema width
  const existing = sheet.getFilter();
  if (existing) existing.remove();
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).createFilter();
}

// Creates Config tab or updates headers. Seeds missing keys only — never overwrites existing values.
function ensureConfigTab_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName("Config");

  if (!sheet) {
    sheet = ss.insertSheet("Config");
  }

  sheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS]);
  sheet.setFrozenRows(1);
  // No filter on Config — it is a settings store, not an operational data tab

  // Collect existing keys (column A, rows 2+)
  const existingKeys = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
    : [];

  // Append only seed rows whose key does not yet exist
  const now = new Date().toISOString();
  for (const [key, value, description] of CONFIG_SEED) {
    if (!existingKeys.includes(key)) {
      sheet.appendRow([key, value, description, now]);
    }
  }
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
// Run setupDashboard() once from the Apps Script editor to build or rebuild
// the Dashboard tab. Safe to re-run: clears and rewrites every time.

function setupDashboard() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName("Dashboard");
  if (!sheet) {
    sheet = ss.insertSheet("Dashboard");
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  // ── Color palette ──────────────────────────────────────────────────────────
  const C = {
    bg:       "#0F172A",
    card:     "#1E293B",
    cardAlt:  "#162032",
    border:   "#334155",
    accent:   "#F97316",
    positive: "#22C55E",
    warning:  "#EAB308",
    negative: "#EF4444",
    text:     "#F8FAFC",
    textSec:  "#CBD5E1",
  };

  const BS = SpreadsheetApp.BorderStyle;

  // ── Sheet background ───────────────────────────────────────────────────────
  sheet.getRange(1, 1, 60, 10).setBackground(C.bg).setFontColor(C.text);

  // ── Row heights ────────────────────────────────────────────────────────────
  sheet.setRowHeight(1,  30);
  sheet.setRowHeight(2,  30);
  sheet.setRowHeight(3,  22);
  sheet.setRowHeight(4,  38);
  sheet.setRowHeight(5,  38);
  sheet.setRowHeight(6,  22);
  sheet.setRowHeight(7,  38);
  sheet.setRowHeight(8,  38);
  sheet.setRowHeight(9,  22);
  sheet.setRowHeight(10, 40);
  sheet.setRowHeight(11, 14);
  sheet.setRowHeight(12, 24);
  sheet.setRowHeight(13, 26);

  // ── Column widths ──────────────────────────────────────────────────────────
  [1,2,3,4,5,6,7,8,9].forEach((col, i) =>
    sheet.setColumnWidth(col, i === 0 ? 90 : 120)
  );

  // ── Title (A1:I2) ──────────────────────────────────────────────────────────
  sheet.getRange("A1:I2").merge()
    .setValue("AMRA Assessment Dashboard")
    .setBackground(C.card)
    .setFontColor(C.accent)
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(false, false, true, false, false, false, C.border, BS.SOLID_MEDIUM);

  // ── KPI Section 1 — labels row 3, cards rows 4–5 ──────────────────────────
  [["B3","TOTAL ASSESSMENTS"],["D3","PASS RATE"],
   ["F3","AVG PERCENTAGE"],   ["H3","AVG TIME TAKEN"]].forEach(([cell, lbl]) => {
    sheet.getRange(cell).setValue(lbl)
      .setFontColor(C.textSec).setFontSize(8).setFontWeight("bold")
      .setHorizontalAlignment("center");
  });

  ["B4:C5","D4:E5","F4:G5","H4:I5"].forEach(r => {
    sheet.getRange(r).merge()
      .setBackground(C.card).setFontColor(C.text)
      .setFontSize(22).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false, C.border, BS.SOLID_MEDIUM);
  });

  sheet.getRange("B4").setFormula('=COUNTA(Assessments!A2:A)');
  sheet.getRange("D4").setFormula('=IFERROR(TEXT(COUNTIF(Assessments!H2:H,">=70")/COUNTA(Assessments!A2:A),"0%"),"-")');
  sheet.getRange("F4").setFormula('=IFERROR(TEXT(AVERAGE(Assessments!H2:H),"0.0")&"%","-")');
  sheet.getRange("H4").setFormula('=IFERROR(TEXT(AVERAGE(Assessments!I2:I),"0.0")&" min","-")');

  // ── KPI Section 2 — labels row 6, cards rows 7–8 ──────────────────────────
  [["B6","AVG BEGINNER CORRECT"],["D6","AVG INTERMEDIATE CORRECT"],
   ["F6","OFFLINE ATTEMPTS"],    ["H6","ONLINE ATTEMPTS"]].forEach(([cell, lbl]) => {
    sheet.getRange(cell).setValue(lbl)
      .setFontColor(C.textSec).setFontSize(8).setFontWeight("bold")
      .setHorizontalAlignment("center");
  });

  ["B7:C8","D7:E8","F7:G8","H7:I8"].forEach(r => {
    sheet.getRange(r).merge()
      .setBackground(C.card).setFontColor(C.text)
      .setFontSize(22).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false, C.border, BS.SOLID_MEDIUM);
  });

  sheet.getRange("B7").setFormula('=IFERROR(TEXT(AVERAGE(Assessment_Details!K2:K),"0.0"),"-")');
  sheet.getRange("D7").setFormula('=IFERROR(TEXT(AVERAGE(Assessment_Details!L2:L),"0.0"),"-")');
  sheet.getRange("F7").setFormula('=COUNTIF(Assessments!J2:J,"offline")');
  sheet.getRange("H7").setFormula('=COUNTIF(Assessments!J2:J,"online")');

  // ── Section 3: Quick Flags — labels row 9, values row 10 ──────────────────
  [["B9","SUSPICIOUS ACTIVITY"],["D9","HIGHEST SCORE"],["F9","LOWEST SCORE"]].forEach(([cell, lbl]) => {
    sheet.getRange(cell).setValue(lbl)
      .setFontColor(C.textSec).setFontSize(8).setFontWeight("bold")
      .setHorizontalAlignment("center");
  });

  ["B10","D10","F10"].forEach(cell => {
    sheet.getRange(cell)
      .setBackground(C.card).setFontColor(C.text)
      .setFontSize(18).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false, C.border, BS.SOLID);
  });

  sheet.getRange("B10").setFormula('=COUNTA(SuspiciousActivity!A2:A)');
  sheet.getRange("D10").setFormula('=IFERROR(MAX(Assessments!G2:G),"-")');
  sheet.getRange("F10").setFormula('=IFERROR(MIN(Assessments!G2:G),"-")');

  // ── Section 4: Latest Submissions — header row 13, data row 14 ────────────
  sheet.getRange("A12").setValue("LATEST SUBMISSIONS  (10 most recent)")
    .setFontColor(C.accent).setFontSize(9).setFontWeight("bold");

  const tableHeaders = ["ref_id", "full_name", "score", "percentage", "mode", "submitted_at"];
  sheet.getRange(13, 1, 1, tableHeaders.length).setValues([tableHeaders])
    .setBackground(C.card).setFontColor(C.accent)
    .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, C.border, BS.SOLID);

  sheet.getRange("A14:F23")
    .setFontColor(C.text).setFontSize(9);
  sheet.getRange("D14:D23").setHorizontalAlignment("center");

  sheet.getRange("A14").setFormula(
    '=IFERROR(QUERY(Assessments!A2:L,"SELECT A,B,G,H,J,L WHERE A IS NOT NULL ORDER BY L DESC LIMIT 10",0),"No submissions yet")'
  );

  // ── Helper cell for Pass Rate CF (CF rules cannot reference other sheets) ──
  // J4: invisible numeric ratio used only by conditional formatting rules below
  sheet.getRange("J4")
    .setFormula('=IFERROR(COUNTIF(Assessments!H2:H,">=70")/COUNTA(Assessments!A2:A),0)')
    .setBackground(C.bg).setFontColor(C.bg);

  // ── Conditional formatting ─────────────────────────────────────────────────
  // Alternating rows — table data area A14:F23
  const cfOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=ISODD(ROW())")
    .setBackground(C.bg)
    .setRanges([sheet.getRange("A14:F23")])
    .build();
  const cfEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=ISEVEN(ROW())")
    .setBackground(C.card)
    .setRanges([sheet.getRange("A14:F23")])
    .build();

  // Pass Rate card D4:E5 — references local helper cell $J$4 (cross-sheet CF not allowed)
  const cfPRGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$J$4>0.75")
    .setBackground(C.positive).setFontColor("#0F172A")
    .setRanges([sheet.getRange("D4:E5")])
    .build();
  const cfPRYellow = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($J$4>=0.5,$J$4<=0.75)")
    .setBackground(C.warning).setFontColor("#0F172A")
    .setRanges([sheet.getRange("D4:E5")])
    .build();
  const cfPRRed = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$J$4<0.5")
    .setBackground(C.negative).setFontColor(C.text)
    .setRanges([sheet.getRange("D4:E5")])
    .build();

  sheet.setConditionalFormatRules([cfOdd, cfEven, cfPRGreen, cfPRYellow, cfPRRed]);

  // ── Freeze title block ─────────────────────────────────────────────────────
  sheet.setFrozenRows(2);

  Logger.log("setupDashboard() complete — dark theme applied.");
}

// ─────────────────────────────────────────────────────────────────────────────

// ── Google Form → auto token ──────────────────────────────────────────────────
// Triggered automatically when the linked Google Form is submitted.
// Setup: Apps Script → Triggers (clock icon) → Add trigger
//   Function: onFormSubmit | Event source: From spreadsheet | Event type: On form submit
//
// Form must have exactly these question titles (case-sensitive):
//   "Full Name"  and  "Email Address"

function onFormSubmit(e) {
  try {
    // e.values = [timestamp, google-email(collected), answer1, answer2, ...]
    // Position 2 = Full Name, Position 3 = Email Address
    const name  = (e.values[2] || "").trim();
    const email = (e.values[3] || "").trim().toLowerCase();

    console.log("Form response — name: " + name + " | email: " + email);

    if (!name || !email) {
      console.error("onFormSubmit: missing name or email in form response");
      console.error("All values: " + JSON.stringify(e.values));
      return;
    }

    const result = generateToken(name, email, 1);

    MailApp.sendEmail({
      to: email,
      subject: "AMRA Capital — Level 1 Assessment Link",
      body:
        "Dear " + name + ",\n\n" +
        "Thank you for your interest in AMRA Capital.\n\n" +
        "Your Level 1 Assessment link is ready:\n\n" +
        result.candidateUrl + "\n\n" +
        "This link is valid for 48 hours and is single-use only.\n\n" +
        "Before you begin:\n" +
        "• Ensure a stable internet connection\n" +
        "• You have 30 minutes to complete 25 questions\n" +
        "• Do not switch tabs or exit fullscreen during the exam\n\n" +
        "Best regards,\n" +
        "AMRA Capital"
    });

    console.log("Token generated and emailed to: " + email);
  } catch (err) {
    console.error("onFormSubmit error: " + err.message);
  }
}

// ── Admin runners — select these from the dropdown in the Apps Script editor ──
// Edit name/email below, then select runGenerateToken → Run.
// Never run generateToken directly from the dropdown — it will get undefined args.

function runGenerateToken() {
  const name  = "Test User";          // ← change this
  const email = "test@amra.capital";  // ← change this
  const level = 1;                    // ← 1 for Level 1
  const result = generateToken(name, email, level);
  Logger.log("Generated: " + JSON.stringify(result));
}

// Apps Script adds Access-Control-Allow-Origin: * automatically for web apps
// deployed to "Anyone". doOptions() handles preflight for non-simple requests.
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
