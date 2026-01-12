// ===========================================
// CHEKHOV ORAL EXAMINER - Google Apps Script
// ===========================================
// This script handles:
// 1. Serving the student submission portal
// 2. Processing paper submissions
// 3. Providing essay lookup for 11Labs agent
// 4. Providing randomized questions for 11Labs agent
// 5. Receiving transcripts via webhook
// 6. Grading via Claude API
// ===========================================

// CONFIGURATION
const SPREADSHEET_ID = "1Az9LPFedJ6c5qMthY4fKHOSbUyCmpNiEU6gRPIB1dKk";
const SUBMISSIONS_SHEET = "Database";  // Renamed from Sheet1
const CONFIG_SHEET = "Config";
const PROMPTS_SHEET = "Prompts";
const QUESTIONS_SHEET = "Questions";
const LOGS_SHEET = "Logs";

// ===========================================
// SPREADSHEET LOGGING (visible in Logs tab)
// ===========================================

/**
 * Writes a log entry to the Logs sheet for easy debugging
 * @param {string} source - The function/context name
 * @param {string} message - The log message
 * @param {Object|string} data - Optional additional data
 */
function sheetLog(source, message, data = "") {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logsSheet = ss.getSheetByName(LOGS_SHEET);

    // Create Logs sheet if it doesn't exist
    if (!logsSheet) {
      logsSheet = ss.insertSheet(LOGS_SHEET);
      logsSheet.appendRow(["Timestamp", "Source", "Message", "Data"]);
      logsSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    }

    // Format data as string if it's an object
    const dataStr = (typeof data === "object") ? JSON.stringify(data) : data;

    // Add log entry
    logsSheet.appendRow([new Date(), source, message, dataStr]);

    // Also log to console for Apps Script logs
    console.log(`[${source}] ${message}`, dataStr);

  } catch (e) {
    // Don't let logging errors break the main flow
    console.log("Logging error:", e.toString());
  }
}

/**
 * Clears all log entries (keeps header row)
 * Run this manually from script editor to clear logs
 */
function clearLogs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logsSheet = ss.getSheetByName(LOGS_SHEET);
  if (logsSheet && logsSheet.getLastRow() > 1) {
    logsSheet.deleteRows(2, logsSheet.getLastRow() - 1);
  }
}

// Column indices for Submissions sheet (1-based)
const COL = {
  TIMESTAMP: 1,
  STUDENT_NAME: 2,
  CODE: 3,
  PAPER: 4,
  STATUS: 5,
  DEFENSE_STARTED: 6,
  DEFENSE_ENDED: 7,
  TRANSCRIPT: 8,
  CLAUDE_GRADE: 9,
  CLAUDE_COMMENTS: 10,
  INSTRUCTOR_NOTES: 11,
  FINAL_GRADE: 12
};

// Status values
const STATUS = {
  SUBMITTED: "Submitted",
  DEFENSE_STARTED: "Defense Started",
  DEFENSE_COMPLETE: "Defense Complete",
  GRADED: "Graded",
  REVIEWED: "Reviewed"
};

// ===========================================
// DEFAULT VALUES (used when Config sheet doesn't exist)
// ===========================================
const DEFAULTS = {
  claude_api_key: "",
  claude_model: "claude-sonnet-4-20250514",
  max_paper_length: "15000",
  webhook_secret: "default_secret_change_me",
  content_questions_count: "2",
  process_questions_count: "1"
};

// ===========================================
// CONFIGURATION HELPERS
// ===========================================

/**
 * Retrieves a configuration value from the Config sheet
 * Falls back to DEFAULTS if Config sheet doesn't exist
 * @param {string} key - The config key to look up
 * @returns {string} The config value
 */
function getConfig(key) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName(CONFIG_SHEET);

    // If Config sheet doesn't exist, use defaults
    if (!configSheet) {
      if (DEFAULTS.hasOwnProperty(key)) {
        return DEFAULTS[key];
      }
      throw new Error("Config key not found and no default: " + key);
    }

    const data = configSheet.getDataRange().getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }

    // Key not in sheet, try defaults
    if (DEFAULTS.hasOwnProperty(key)) {
      return DEFAULTS[key];
    }
    throw new Error("Config key not found: " + key);

  } catch (e) {
    // If any error, try defaults
    if (DEFAULTS.hasOwnProperty(key)) {
      return DEFAULTS[key];
    }
    throw e;
  }
}

/**
 * Retrieves a prompt from the Prompts sheet
 * @param {string} promptName - The prompt name to look up
 * @returns {string} The prompt text
 */
function getPrompt(promptName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const promptsSheet = ss.getSheetByName(PROMPTS_SHEET);

    if (!promptsSheet) {
      throw new Error("Prompts sheet not found. Please create a 'Prompts' tab.");
    }

    const data = promptsSheet.getDataRange().getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === promptName) {
        return data[i][1];
      }
    }
    throw new Error("Prompt not found: " + promptName);

  } catch (e) {
    throw e;
  }
}

/**
 * Retrieves randomized questions from the Questions sheet
 * @param {number} contentCount - Number of content questions to return (default from config)
 * @param {number} processCount - Number of process questions to return (default from config)
 * @returns {Object} Object with contentQuestions and processQuestions arrays
 */
function getRandomizedQuestions(contentCount, processCount) {
  // Use config defaults if not specified
  if (contentCount === undefined) {
    contentCount = parseInt(getConfig("content_questions_count"));
  }
  if (processCount === undefined) {
    processCount = parseInt(getConfig("process_questions_count"));
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const questionsSheet = ss.getSheetByName(QUESTIONS_SHEET);

  if (!questionsSheet) {
    throw new Error("Questions sheet not found. Please create a 'Questions' tab with columns: category, question");
  }

  const data = questionsSheet.getDataRange().getValues();

  // Separate questions by category (no header row expected)
  const contentQuestions = [];
  const processQuestions = [];

  for (let i = 0; i < data.length; i++) {
    const category = data[i][0]?.toString().toLowerCase().trim();
    const question = data[i][1]?.toString().trim();

    if (!question) continue; // Skip empty rows

    if (category === "content") {
      contentQuestions.push(question);
    } else if (category === "process") {
      processQuestions.push(question);
    }
  }

  // Shuffle and select the requested number of questions
  const selectedContent = shuffleArray(contentQuestions).slice(0, contentCount);
  const selectedProcess = shuffleArray(processQuestions).slice(0, processCount);

  return {
    contentQuestions: selectedContent,
    processQuestions: selectedProcess,
    totalSelected: selectedContent.length + selectedProcess.length
  };
}

/**
 * Fisher-Yates shuffle algorithm for randomizing arrays
 * @param {Array} array - The array to shuffle
 * @returns {Array} A new shuffled array (does not modify original)
 */
function shuffleArray(array) {
  // Create a copy to avoid modifying the original
  const shuffled = [...array];

  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }

  return shuffled;
}

// ===========================================
// WEB APP ENTRY POINTS
// ===========================================

/**
 * Handles GET requests - serves the portal or handles API calls
 */
function doGet(e) {
  const action = e?.parameter?.action;

  console.log("=== doGet called ===");
  console.log("Action:", action || "none (serving portal)");

  // API endpoint for 11Labs to fetch essays
  if (action === "getEssay") {
    return handleGetEssay(e);
  }

  // API endpoint for 11Labs to fetch randomized questions
  if (action === "getQuestions") {
    return handleGetQuestions(e);
  }

  // Default: serve the HTML portal
  console.log("Serving HTML portal");
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Oral Defense Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Handles POST requests - receives webhooks from 11Labs
 */
function doPost(e) {
  try {
    console.log("=== doPost called ===");
    console.log("Content length:", e.postData?.length);

    const payload = JSON.parse(e.postData.contents);
    console.log("Payload type:", payload.type);

    // Verify webhook secret if provided
    const providedSecret = e?.parameter?.secret;
    const expectedSecret = getConfig("webhook_secret");

    if (providedSecret !== expectedSecret) {
      console.log("POST: Secret validation FAILED");
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Invalid webhook secret"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    console.log("POST: Secret validation passed");

    // Handle transcript webhook from 11Labs
    return handleTranscriptWebhook(payload);

  } catch (error) {
    console.log("EXCEPTION in doPost:", error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===========================================
// PAPER SUBMISSION (Called from frontend)
// ===========================================

/**
 * Processes a paper submission from the portal
 * @param {Object} formObject - Contains name and essay fields
 * @returns {Object} Status and code or error message
 */
function processSubmission(formObject) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);

    // Validate paper length
    const maxLength = parseInt(getConfig("max_paper_length"));
    if (formObject.essay.length > maxLength) {
      return {
        status: "error",
        message: `Paper exceeds maximum length of ${maxLength} characters. Your paper has ${formObject.essay.length} characters.`
      };
    }

    // Generate a unique code
    const code = generateUniqueCode(sheet);

    // Create row with all columns (empty strings for unused columns)
    const newRow = new Array(12).fill("");
    newRow[COL.TIMESTAMP - 1] = new Date();
    newRow[COL.STUDENT_NAME - 1] = formObject.name;
    newRow[COL.CODE - 1] = code;
    newRow[COL.PAPER - 1] = formObject.essay;
    newRow[COL.STATUS - 1] = STATUS.SUBMITTED;

    sheet.appendRow(newRow);

    return { status: "success", code: code };

  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * Generates a unique 4-digit code not already in use
 * @param {Sheet} sheet - The submissions sheet
 * @returns {string} A unique 4-digit code
 */
function generateUniqueCode(sheet) {
  const lastRow = sheet.getLastRow();
  const existingCodes = new Set();

  if (lastRow > 1) { // Skip header row
    const codeColumn = sheet.getRange(2, COL.CODE, lastRow - 1, 1).getValues();
    codeColumn.forEach(row => {
      if (row[0]) existingCodes.add(row[0].toString());
    });
  }

  let code;
  let attempts = 0;
  const maxAttempts = 100;

  do {
    code = Math.floor(1000 + Math.random() * 9000).toString();
    attempts++;
    if (attempts > maxAttempts) {
      throw new Error("Unable to generate unique code after " + maxAttempts + " attempts");
    }
  } while (existingCodes.has(code));

  return code;
}

// ===========================================
// 11LABS ESSAY LOOKUP (GET endpoint)
// ===========================================

/**
 * Handles essay lookup requests from 11Labs agent
 * GET ?action=getEssay&code=1234&secret=xxx
 */
function handleGetEssay(e) {
  try {
    const code = e?.parameter?.code;
    const providedSecret = e?.parameter?.secret;
    const expectedSecret = getConfig("webhook_secret");

    // Log all incoming parameters to spreadsheet for debugging
    sheetLog("handleGetEssay", "Request received", {
      allParams: e?.parameter,
      code: code,
      codeType: typeof code,
      codeLength: code ? code.length : "N/A",
      codeCharCodes: code ? code.split('').map(c => c.charCodeAt(0)).join(',') : "N/A"
    });

    // Validate secret
    if (providedSecret !== expectedSecret) {
      sheetLog("handleGetEssay", "SECRET FAILED", {
        provided: providedSecret ? providedSecret.substring(0, 4) + "..." : "none",
        expected: expectedSecret ? expectedSecret.substring(0, 4) + "..." : "none"
      });
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Invalid secret"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    sheetLog("handleGetEssay", "Secret OK", "");

    // Validate code provided
    if (!code) {
      sheetLog("handleGetEssay", "ERROR: No code provided", "");
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "No code provided"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Look up the essay
    sheetLog("handleGetEssay", "Looking up code", code);
    const result = getEssayByCode(code);

    if (!result) {
      sheetLog("handleGetEssay", "CODE NOT FOUND", { searchedFor: code });
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Code not found. Please check the code and try again."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    sheetLog("handleGetEssay", "Found student", {
      name: result.studentName,
      status: result.status,
      row: result.row
    });

    // Allow essay retrieval during active defense (for retries/reconnections)
    // But reject if defense is already complete or graded
    if (result.status !== STATUS.SUBMITTED && result.status !== STATUS.DEFENSE_STARTED) {
      sheetLog("handleGetEssay", "INVALID STATUS", {
        code: code,
        status: result.status
      });
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "This code has already been used for a defense."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Only update status if this is the first call (status is still Submitted)
    if (result.status === STATUS.SUBMITTED) {
      sheetLog("handleGetEssay", "Updating to Defense Started", code);
      updateStudentStatus(code, STATUS.DEFENSE_STARTED, { defenseStarted: new Date() });
    }

    const wordCount = result.essay.split(/\s+/).length;
    sheetLog("handleGetEssay", "SUCCESS - returning essay", {
      code: code,
      student: result.studentName,
      wordCount: wordCount
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      studentName: result.studentName,
      essay: result.essay,
      wordCount: wordCount
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    sheetLog("handleGetEssay", "EXCEPTION", error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===========================================
// 11LABS QUESTIONS LOOKUP (GET endpoint)
// ===========================================

/**
 * Handles randomized questions requests from 11Labs agent
 * GET ?action=getQuestions&secret=xxx
 * Optional: &contentCount=4&processCount=2
 */
function handleGetQuestions(e) {
  try {
    console.log("=== getQuestions Request ===");
    console.log("All parameters:", JSON.stringify(e?.parameter));

    const providedSecret = e?.parameter?.secret;
    const expectedSecret = getConfig("webhook_secret");

    // Validate secret
    if (providedSecret !== expectedSecret) {
      console.log("Secret validation FAILED");
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Invalid secret"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    console.log("Secret validation passed");

    // Get optional count parameters
    const contentCount = e?.parameter?.contentCount
      ? parseInt(e.parameter.contentCount)
      : undefined;
    const processCount = e?.parameter?.processCount
      ? parseInt(e.parameter.processCount)
      : undefined;

    // Get randomized questions
    const questions = getRandomizedQuestions(contentCount, processCount);

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      contentQuestions: questions.contentQuestions,
      processQuestions: questions.processQuestions,
      totalQuestions: questions.totalSelected
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Looks up an essay by its defense code
 * @param {string} code - The 4-digit defense code
 * @returns {Object|null} Student data or null if not found
 */
function getEssayByCode(code) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  // Normalize the search code (trim whitespace, convert to string)
  const searchCode = code.toString().trim();

  // Collect all existing codes for debugging
  const existingCodes = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const rowCode = data[i][COL.CODE - 1]?.toString().trim() || "";
    existingCodes.push(rowCode);

    if (rowCode === searchCode) {
      sheetLog("getEssayByCode", "MATCH FOUND", {
        row: i + 1,
        code: rowCode,
        student: data[i][COL.STUDENT_NAME - 1]
      });
      return {
        row: i + 1, // 1-based row number
        studentName: data[i][COL.STUDENT_NAME - 1],
        essay: data[i][COL.PAPER - 1],
        status: data[i][COL.STATUS - 1]
      };
    }
  }

  // No match found - log all existing codes to spreadsheet for debugging
  const nearMatches = existingCodes.filter(ec =>
    ec.includes(searchCode) || searchCode.includes(ec)
  );

  sheetLog("getEssayByCode", "NO MATCH FOUND", {
    searchedFor: searchCode,
    searchLength: searchCode.length,
    searchCharCodes: searchCode.split('').map(c => c.charCodeAt(0)).join(','),
    totalCodesInDB: existingCodes.length,
    existingCodes: existingCodes.join(", "),
    nearMatches: nearMatches.length > 0 ? nearMatches.join(", ") : "none"
  });

  return null;
}

/**
 * Updates a student's status and optional fields
 * @param {string} code - The defense code
 * @param {string} newStatus - The new status value
 * @param {Object} additionalFields - Optional fields to update
 */
function updateStudentStatus(code, newStatus, additionalFields = {}) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.CODE - 1].toString() === code.toString()) {
      const row = i + 1;

      // Update status
      sheet.getRange(row, COL.STATUS).setValue(newStatus);

      // Update additional fields
      if (additionalFields.defenseStarted) {
        sheet.getRange(row, COL.DEFENSE_STARTED).setValue(additionalFields.defenseStarted);
      }
      if (additionalFields.defenseEnded) {
        sheet.getRange(row, COL.DEFENSE_ENDED).setValue(additionalFields.defenseEnded);
      }
      if (additionalFields.transcript) {
        sheet.getRange(row, COL.TRANSCRIPT).setValue(additionalFields.transcript);
      }
      if (additionalFields.grade) {
        sheet.getRange(row, COL.CLAUDE_GRADE).setValue(additionalFields.grade);
      }
      if (additionalFields.comments) {
        sheet.getRange(row, COL.CLAUDE_COMMENTS).setValue(additionalFields.comments);
      }

      return true;
    }
  }
  return false;
}

// ===========================================
// TRANSCRIPT WEBHOOK (POST endpoint)
// ===========================================

/**
 * Handles incoming transcript webhooks from 11Labs
 * Expected payload format:
 * {
 *   "type": "post_call_transcription",
 *   "event_timestamp": 1739537297,
 *   "data": {
 *     "agent_id": "xyz",
 *     "conversation_id": "abc",
 *     "status": "done",
 *     "transcript": [
 *       { "role": "agent", "message": "Hello..." },
 *       { "role": "user", "message": "My code is 1234" }
 *     ]
 *   }
 * }
 */
function handleTranscriptWebhook(payload) {
  try {
    console.log("=== Transcript Webhook Received ===");
    console.log("Payload type:", payload.type);

    // Extract data from 11Labs webhook payload
    const data = payload.data || payload;
    const transcriptArray = data.transcript || [];
    const conversationId = data.conversation_id || "";

    console.log("Conversation ID:", conversationId);
    console.log("Transcript entries:", transcriptArray.length);

    // Convert transcript array to readable string
    const transcriptText = formatTranscript(transcriptArray);

    // Extract the defense code from the transcript
    const code = extractCodeFromTranscript(transcriptText);
    console.log("Extracted code from transcript:", code);

    if (!code) {
      // Log the payload for debugging
      console.log("ERROR: No code found in transcript");
      console.log("Transcript text:", transcriptText.substring(0, 500) + "...");
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Could not determine student code from transcript",
        conversation_id: conversationId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Update the student record
    const updated = updateStudentStatus(code, STATUS.DEFENSE_COMPLETE, {
      defenseEnded: new Date(),
      transcript: transcriptText
    });

    if (!updated) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Could not find student with code: " + code
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Optionally trigger grading immediately
    // gradeDefense(code);

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Transcript saved for code: " + code
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Formats a transcript array into readable text
 * @param {Array} transcriptArray - Array of {role, message} objects
 * @returns {string} Formatted transcript text
 */
function formatTranscript(transcriptArray) {
  if (!Array.isArray(transcriptArray)) {
    return String(transcriptArray);
  }

  return transcriptArray.map(entry => {
    const role = entry.role === "agent" ? "EXAMINER" : "STUDENT";
    return `${role}: ${entry.message}`;
  }).join("\n\n");
}

/**
 * Attempts to extract the defense code from transcript text
 * Looks for 4-digit codes, prioritizing those near "code" mentions
 * @param {string} transcript - The conversation transcript
 * @returns {string|null} The extracted code or null
 */
function extractCodeFromTranscript(transcript) {
  // First, try to find a code mentioned near the word "code"
  const codeContextMatch = transcript.match(/code[^\d]*(\d{4})\b/i);
  if (codeContextMatch) {
    return codeContextMatch[1];
  }

  // Fallback: find any 4-digit number in student responses
  const studentLines = transcript.split('\n')
    .filter(line => line.startsWith('STUDENT:'));

  for (const line of studentLines) {
    const match = line.match(/\b(\d{4})\b/);
    if (match) {
      return match[0];
    }
  }

  // Last resort: any 4-digit number in the transcript
  const matches = transcript.match(/\b(\d{4})\b/g);
  if (matches && matches.length > 0) {
    return matches[0];
  }

  return null;
}

// ===========================================
// CLAUDE GRADING (Phase 4 - placeholder)
// ===========================================

/**
 * Grades a defense using Claude API
 * @param {string} code - The student's defense code
 */
function gradeDefense(code) {
  // TODO: Implement in Phase 4
  // 1. Get paper and transcript
  // 2. Build prompt from Prompts sheet
  // 3. Call Claude API
  // 4. Parse response
  // 5. Update sheet with grade and comments
}

// ===========================================
// UTILITY FUNCTIONS
// ===========================================

/**
 * Includes HTML files in other HTML files (standard Apps Script pattern)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * Manual trigger to grade all completed defenses
 * Can be run from script editor or triggered by menu
 */
function gradeAllPending() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.STATUS - 1] === STATUS.DEFENSE_COMPLETE) {
      const code = data[i][COL.CODE - 1].toString();
      gradeDefense(code);
    }
  }
}

/**
 * Creates a custom menu in the spreadsheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Oral Defense')
    .addItem('Grade All Pending', 'gradeAllPending')
    .addItem('Refresh Status Counts', 'showStatusCounts')
    .addToUi();
}

/**
 * Shows a summary of submission statuses
 */
function showStatusCounts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1] || "Unknown";
    counts[status] = (counts[status] || 0) + 1;
  }

  let message = "Status Summary:\n";
  for (const [status, count] of Object.entries(counts)) {
    message += `${status}: ${count}\n`;
  }

  SpreadsheetApp.getUi().alert(message);
}
