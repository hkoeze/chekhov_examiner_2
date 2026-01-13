// ===========================================
// CHEKHOV ORAL EXAMINER - Google Apps Script
// ===========================================
// This script handles:
// 1. Serving the student submission portal
// 2. Processing paper submissions (generates UUID session_id)
// 3. Providing randomized questions for 11Labs agent
// 4. Receiving transcripts via webhook (matched by session_id)
// 5. Grading via Claude API (Phase 4)
//
// Workflow:
// - Student submits essay via portal -> gets session_id
// - Portal configures 11Labs widget with session_id
// - Student has voice defense, pastes essay in chat
// - 11Labs sends transcript webhook with session_id
// - System matches transcript to submission, stores for grading
// ===========================================

// CONFIGURATION
const SPREADSHEET_ID = "1UE2yNBbZmTKe4g1C_lVoAvE3RcZAuHGbjEYstDm2xzI";
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
  SESSION_ID: 3,  // Changed from CODE - now stores UUID for webhook correlation
  PAPER: 4,
  STATUS: 5,
  DEFENSE_STARTED: 6,
  DEFENSE_ENDED: 7,
  TRANSCRIPT: 8,
  AI_MULTIPLIER: 9,   // Renamed from CLAUDE_GRADE
  AI_COMMENT: 10,     // Renamed from CLAUDE_COMMENTS
  INSTRUCTOR_NOTES: 11,
  FINAL_GRADE: 12,
  CONVERSATION_ID: 13  // New: stores 11Labs conversation_id as backup
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
  gemini_api_key: "",
  gemini_model: "gemini-3-flash-preview",
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
 * @returns {Object} Status and session_id or error message
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

    // Generate a unique session ID (UUID)
    const sessionId = generateSessionId();

    // Create row with all columns (empty strings for unused columns)
    const newRow = new Array(13).fill("");
    newRow[COL.TIMESTAMP - 1] = new Date();
    newRow[COL.STUDENT_NAME - 1] = formObject.name;
    newRow[COL.SESSION_ID - 1] = sessionId;
    newRow[COL.PAPER - 1] = formObject.essay;
    newRow[COL.STATUS - 1] = STATUS.SUBMITTED;

    sheet.appendRow(newRow);

    // Format the new row: clip text and set compact height (2 lines max)
    const newRowNum = sheet.getLastRow();
    sheet.getRange(newRowNum, 1, 1, 13).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.setRowHeightsForced(newRowNum, 1, 42);

    sheetLog("processSubmission", "Essay submitted", {
      studentName: formObject.name,
      sessionId: sessionId,
      essayLength: formObject.essay.length
    });

    return { status: "success", sessionId: sessionId };

  } catch (e) {
    sheetLog("processSubmission", "ERROR", e.toString());
    return { status: "error", message: e.toString() };
  }
}

/**
 * Generates a unique session ID (UUID v4 format)
 * Used to correlate portal submissions with 11Labs webhook callbacks
 * @returns {string} A UUID string
 */
function generateSessionId() {
  // Generate UUID v4 format: xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx
  const template = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
  return template.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
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
 * Looks up a submission by its session ID
 * @param {string} sessionId - The UUID session ID
 * @returns {Object|null} Student data or null if not found
 */
function getSubmissionBySessionId(sessionId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  const searchId = sessionId.toString().trim();

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const rowSessionId = data[i][COL.SESSION_ID - 1]?.toString().trim() || "";

    if (rowSessionId === searchId) {
      sheetLog("getSubmissionBySessionId", "MATCH FOUND", {
        row: i + 1,
        sessionId: rowSessionId,
        student: data[i][COL.STUDENT_NAME - 1]
      });
      return {
        row: i + 1,
        studentName: data[i][COL.STUDENT_NAME - 1],
        essay: data[i][COL.PAPER - 1],
        status: data[i][COL.STATUS - 1],
        transcript: data[i][COL.TRANSCRIPT - 1] || ""
      };
    }
  }

  sheetLog("getSubmissionBySessionId", "NO MATCH FOUND", { searchedFor: searchId });
  return null;
}

/**
 * Looks up a submission by student name (fallback method)
 * Returns the most recent submission with matching name and valid status
 * @param {string} name - The student's name
 * @returns {Object|null} Student data or null if not found
 */
function getSubmissionByName(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  const searchName = name.toString().trim().toLowerCase();
  let bestMatch = null;

  // Skip header row, find most recent matching submission
  for (let i = 1; i < data.length; i++) {
    const rowName = data[i][COL.STUDENT_NAME - 1]?.toString().trim().toLowerCase() || "";
    const status = data[i][COL.STATUS - 1];

    // Only match submissions that are awaiting defense or in progress
    if (rowName === searchName &&
        (status === STATUS.SUBMITTED || status === STATUS.DEFENSE_STARTED)) {
      bestMatch = {
        row: i + 1,
        sessionId: data[i][COL.SESSION_ID - 1],
        studentName: data[i][COL.STUDENT_NAME - 1],
        essay: data[i][COL.PAPER - 1],
        status: status
      };
      // Don't break - continue to find most recent (last in sheet)
    }
  }

  if (bestMatch) {
    sheetLog("getSubmissionByName", "MATCH FOUND", {
      row: bestMatch.row,
      student: bestMatch.studentName,
      sessionId: bestMatch.sessionId
    });
  } else {
    sheetLog("getSubmissionByName", "NO MATCH FOUND", { searchedFor: name });
  }

  return bestMatch;
}

/**
 * Updates a student's status and optional fields
 * @param {string} sessionId - The session ID (UUID)
 * @param {string} newStatus - The new status value
 * @param {Object} additionalFields - Optional fields to update
 */
function updateStudentStatus(sessionId, newStatus, additionalFields = {}) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.SESSION_ID - 1]?.toString() === sessionId.toString()) {
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
        sheet.getRange(row, COL.AI_MULTIPLIER).setValue(additionalFields.grade);
      }
      if (additionalFields.comments) {
        sheet.getRange(row, COL.AI_COMMENT).setValue(additionalFields.comments);
      }
      if (additionalFields.conversationId) {
        sheet.getRange(row, COL.CONVERSATION_ID).setValue(additionalFields.conversationId);
      }

      sheetLog("updateStudentStatus", "Updated", {
        sessionId: sessionId,
        newStatus: newStatus,
        row: row
      });

      return true;
    }
  }

  sheetLog("updateStudentStatus", "NOT FOUND", { sessionId: sessionId });
  return false;
}

// ===========================================
// TRANSCRIPT WEBHOOK (POST endpoint)
// ===========================================

/**
 * Handles incoming transcript webhooks from 11Labs
 * Expected payload format (matches Get Conversation API response):
 * {
 *   "type": "post_call_transcription",
 *   "event_timestamp": 1739537297,
 *   "data": {
 *     "agent_id": "xyz",
 *     "conversation_id": "abc",
 *     "status": "done",
 *     "transcript": [
 *       { "role": "agent", "message": "Hello..." },
 *       { "role": "user", "message": "..." }
 *     ],
 *     "conversation_initiation_client_data": {
 *       "dynamic_variables": {
 *         "session_id": "uuid-here"
 *       }
 *     }
 *   }
 * }
 */
function handleTranscriptWebhook(payload) {
  try {
    sheetLog("handleTranscriptWebhook", "Webhook received", {
      type: payload.type,
      hasData: !!payload.data
    });

    // Extract data from 11Labs webhook payload
    const data = payload.data || payload;
    const transcriptArray = data.transcript || [];
    const conversationId = data.conversation_id || "";

    // Extract session_id from dynamic_variables (primary method)
    const clientData = data.conversation_initiation_client_data || {};
    const dynamicVars = clientData.dynamic_variables || {};
    let sessionId = dynamicVars.session_id || null;

    sheetLog("handleTranscriptWebhook", "Extracted data", {
      conversationId: conversationId,
      transcriptEntries: transcriptArray.length,
      sessionId: sessionId,
      hasDynamicVars: Object.keys(dynamicVars).length > 0
    });

    // Convert transcript array to readable string
    const transcriptText = formatTranscript(transcriptArray);

    // Try to find the submission record
    let submission = null;
    let matchMethod = "";

    // Method 1: Match by session_id from dynamic_variables
    if (sessionId) {
      submission = getSubmissionBySessionId(sessionId);
      if (submission) matchMethod = "session_id";
    }

    // Method 2: Fallback - try to extract student name from transcript and match
    if (!submission) {
      const studentName = extractStudentNameFromTranscript(transcriptText);
      if (studentName) {
        submission = getSubmissionByName(studentName);
        if (submission) {
          matchMethod = "name_fallback";
          sessionId = submission.sessionId;
        }
      }
    }

    // Method 3: Last resort - find most recent "SUBMITTED" or "DEFENSE_STARTED" record
    if (!submission) {
      submission = getMostRecentPendingSubmission();
      if (submission) {
        matchMethod = "most_recent_fallback";
        sessionId = submission.sessionId;
      }
    }

    if (!submission) {
      sheetLog("handleTranscriptWebhook", "NO MATCH FOUND", {
        conversationId: conversationId,
        attemptedSessionId: dynamicVars.session_id
      });
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Could not find matching student submission",
        conversation_id: conversationId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    sheetLog("handleTranscriptWebhook", "MATCH FOUND", {
      matchMethod: matchMethod,
      sessionId: sessionId,
      studentName: submission.studentName
    });

    // Update the student record
    const updated = updateStudentStatus(sessionId, STATUS.DEFENSE_COMPLETE, {
      defenseStarted: submission.status === STATUS.SUBMITTED ? new Date() : null,
      defenseEnded: new Date(),
      transcript: transcriptText,
      conversationId: conversationId
    });

    if (!updated) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Could not update student record for session: " + sessionId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Transcript saved",
      session_id: sessionId,
      match_method: matchMethod
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    sheetLog("handleTranscriptWebhook", "EXCEPTION", error.toString());
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
 * Attempts to extract the student's name from transcript
 * Looks for common patterns like "my name is X" or "I'm X"
 * @param {string} transcript - The conversation transcript
 * @returns {string|null} The extracted name or null
 */
function extractStudentNameFromTranscript(transcript) {
  // Look for common name introduction patterns in student lines
  const studentLines = transcript.split('\n')
    .filter(line => line.startsWith('STUDENT:'))
    .join(' ');

  // Pattern: "my name is [Name]" or "I'm [Name]" or "I am [Name]"
  const patterns = [
    /my name is ([A-Z][a-z]+ [A-Z][a-z]+)/i,
    /my name is ([A-Z][a-z]+)/i,
    /I'm ([A-Z][a-z]+ [A-Z][a-z]+)/i,
    /I am ([A-Z][a-z]+ [A-Z][a-z]+)/i,
    /this is ([A-Z][a-z]+ [A-Z][a-z]+)/i
  ];

  for (const pattern of patterns) {
    const match = studentLines.match(pattern);
    if (match) {
      return match[1].trim();
    }
  }

  return null;
}

/**
 * Gets the most recent submission that's awaiting defense
 * Used as a last-resort fallback when session_id matching fails
 * @returns {Object|null} The most recent pending submission or null
 */
function getMostRecentPendingSubmission() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  const data = sheet.getDataRange().getValues();

  let mostRecent = null;
  let mostRecentTime = null;

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    const timestamp = data[i][COL.TIMESTAMP - 1];

    if (status === STATUS.SUBMITTED || status === STATUS.DEFENSE_STARTED) {
      const rowTime = new Date(timestamp).getTime();
      if (!mostRecentTime || rowTime > mostRecentTime) {
        mostRecentTime = rowTime;
        mostRecent = {
          row: i + 1,
          sessionId: data[i][COL.SESSION_ID - 1],
          studentName: data[i][COL.STUDENT_NAME - 1],
          essay: data[i][COL.PAPER - 1],
          status: status
        };
      }
    }
  }

  if (mostRecent) {
    sheetLog("getMostRecentPendingSubmission", "Found", {
      sessionId: mostRecent.sessionId,
      studentName: mostRecent.studentName
    });
  }

  return mostRecent;
}

// ===========================================
// GEMINI GRADING
// ===========================================

/**
 * Calls the Gemini API with the given prompt
 * @param {string} prompt - The full prompt to send
 * @returns {Object} Parsed response with grade and comments
 */
function callGemini(prompt) {
  const apiKey = getConfig("gemini_api_key");
  const model = getConfig("gemini_model") || "gemini-3-flash-preview";

  if (!apiKey) {
    throw new Error("Gemini API key not configured. Add 'gemini_api_key' to Config sheet.");
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.3,  // Lower temperature for more consistent grading
      maxOutputTokens: 2048,
      thinkingConfig: {
        thinkingLevel: "medium"  // Options: minimal, low, medium, high
      }
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  sheetLog("callGemini", "Calling Gemini API", { model: model, promptLength: prompt.length });

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    sheetLog("callGemini", "API Error", { code: responseCode, response: responseText });
    throw new Error(`Gemini API error (${responseCode}): ${responseText}`);
  }

  const result = JSON.parse(responseText);

  // Extract the text from Gemini's response
  const generatedText = result.candidates?.[0]?.content?.parts?.[0]?.text || "";

  sheetLog("callGemini", "API Success", { responseLength: generatedText.length });

  return generatedText;
}

/**
 * Parses Gemini's grading response to extract grade and comments
 * Expects format with scores for each rubric element and a final average
 * @param {string} response - The raw response from Gemini
 * @returns {Object} Object with grade (number) and comments (string)
 */
function parseGradingResponse(response) {
  // Try to extract the final grade multiplier
  // Look for patterns like "Final Score: 0.95" or "final grade multiplier: 0.95"
  const gradePatterns = [
    /final\s*(?:score|grade|multiplier)[:\s]*([0-9]+\.?[0-9]*)/i,
    /average[:\s]*([0-9]+\.?[0-9]*)/i,
    /([0-9]+\.[0-9]+)\s*(?:final|average|overall)/i
  ];

  let grade = null;
  for (const pattern of gradePatterns) {
    const match = response.match(pattern);
    if (match) {
      grade = parseFloat(match[1]);
      break;
    }
  }

  // If no grade found, try to calculate from individual scores
  if (!grade) {
    const scoreMatches = response.match(/(?:score|rating)[:\s]*([0-9]+\.?[0-9]*)/gi);
    if (scoreMatches && scoreMatches.length >= 4) {
      const scores = scoreMatches.map(m => parseFloat(m.match(/([0-9]+\.?[0-9]*)/)[1]));
      grade = scores.reduce((a, b) => a + b, 0) / scores.length;
    }
  }

  // Default to 1.0 if parsing fails
  if (!grade || isNaN(grade)) {
    grade = 1.0;
    sheetLog("parseGradingResponse", "Could not parse grade, defaulting to 1.0", { response: response.substring(0, 500) });
  }

  return {
    grade: Math.round(grade * 100) / 100,  // Round to 2 decimal places
    comments: response
  };
}

/**
 * Grades a defense using Gemini API
 * @param {string} sessionId - The student's session ID
 * @returns {Object} Result with success status and grade info
 */
function gradeDefense(sessionId) {
  try {
    sheetLog("gradeDefense", "Starting grading", { sessionId: sessionId });

    // 1. Get the submission data
    const submission = getSubmissionBySessionId(sessionId);
    if (!submission) {
      throw new Error("Submission not found for session: " + sessionId);
    }

    if (!submission.transcript) {
      throw new Error("No transcript found for session: " + sessionId);
    }

    // 2. Get prompts from Prompts sheet
    const systemPrompt = getPrompt("grading_system_prompt");
    const rubric = getPrompt("grading_rubric");

    // 3. Build the full prompt
    const fullPrompt = `${systemPrompt}

${rubric}

---

STUDENT NAME: ${submission.studentName}

---

STUDENT ESSAY:
${submission.essay}

---

ORAL DEFENSE TRANSCRIPT:
${submission.transcript}

---

Please analyze this oral defense using the rubric above. For each of the 4 criteria, provide:
1. The score you assign
2. Brief justification for that score

Then calculate the final average grade multiplier and provide the ~200 word rationale as specified in the rubric.

Format your response with clear headings for each rubric element.`;

    // 4. Call Gemini API
    const response = callGemini(fullPrompt);

    // 5. Parse the response
    const parsed = parseGradingResponse(response);

    // 6. Update the sheet
    const updated = updateStudentStatus(sessionId, STATUS.GRADED, {
      grade: parsed.grade,
      comments: parsed.comments
    });

    if (!updated) {
      throw new Error("Failed to update student record");
    }

    sheetLog("gradeDefense", "Grading complete", {
      sessionId: sessionId,
      grade: parsed.grade
    });

    return {
      success: true,
      sessionId: sessionId,
      grade: parsed.grade,
      comments: parsed.comments
    };

  } catch (error) {
    sheetLog("gradeDefense", "ERROR", { sessionId: sessionId, error: error.toString() });
    return {
      success: false,
      sessionId: sessionId,
      error: error.toString()
    };
  }
}

// ===========================================
// UTILITY FUNCTIONS
// ===========================================

/**
 * Formats the Database sheet for better readability:
 * - Sets all cells to clip overflow text (no wrapping or overflow)
 * - Sets compact column widths appropriate for each data type
 * - Sets compact row heights to show more entries
 */
function formatDatabaseSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);

  if (!sheet) {
    sheetLog("formatDatabaseSheet", "Database sheet not found", {});
    return;
  }

  // Format a large range to cover future submissions (1000 rows)
  const maxRows = 1000;
  // Get actual number of columns in sheet (or use a reasonable max)
  const lastCol = Math.max(sheet.getLastColumn(), 13, 26); // At least 26 columns (A-Z)

  // Set all cells to CLIP wrap strategy (no wrapping, no overflow)
  const fullRange = sheet.getRange(1, 1, maxRows, lastCol);
  fullRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Set compact column widths (in pixels) for known columns
  const columnWidths = {
    [COL.TIMESTAMP]: 130,        // Date/time
    [COL.STUDENT_NAME]: 120,     // Names
    [COL.SESSION_ID]: 100,       // UUID (clipped)
    [COL.PAPER]: 150,            // Long text - keep narrow
    [COL.STATUS]: 110,           // Status values
    [COL.DEFENSE_STARTED]: 130,  // Date/time
    [COL.DEFENSE_ENDED]: 130,    // Date/time
    [COL.TRANSCRIPT]: 150,       // Long text - keep narrow
    [COL.AI_MULTIPLIER]: 80,     // Numeric grade
    [COL.AI_COMMENT]: 150,       // Long text - keep narrow
    [COL.INSTRUCTOR_NOTES]: 120, // Notes
    [COL.FINAL_GRADE]: 80,       // Grade
    [COL.CONVERSATION_ID]: 100   // ID (clipped)
  };

  for (const [col, width] of Object.entries(columnWidths)) {
    sheet.setColumnWidth(parseInt(col), width);
  }

  // Set compact row height for all rows (42 pixels = 2 lines max)
  sheet.setRowHeightsForced(1, maxRows, 42);

  sheetLog("formatDatabaseSheet", "Formatting applied", {
    rows: maxRows,
    columns: lastCol
  });
}

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
      const sessionId = data[i][COL.SESSION_ID - 1].toString();
      gradeDefense(sessionId);
    }
  }
}

/**
 * Creates a custom menu in the spreadsheet and applies formatting
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Oral Defense')
    .addItem('Grade All Pending', 'gradeAllPending')
    .addItem('Refresh Status Counts', 'showStatusCounts')
    .addSeparator()
    .addItem('Format Database Sheet', 'formatDatabaseSheet')
    .addToUi();

  // Auto-format the database sheet on open
  formatDatabaseSheet();
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
