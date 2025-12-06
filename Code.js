/**
 * Math Marking System - Backend Code
 * Features: Chunking Support, Strict 1M/1A Grading, Traditional Chinese
 */

// CONFIGURATION
const APP_NAME = "Math Marking System";
const DRIVE_FOLDER_NAME = "Math_Marking_System_Uploads";
const SHEET_NAME = "Math_Scores";
const DEFAULT_BOT = "GPT-5.1"; // Default model

/**
 * Serves the Web App
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Initial Setup: Creates Drive Folder and Spreadsheet
 */
function setup() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
  
  const files = DriveApp.getFilesByName(SHEET_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
  }
  
  let sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Student ID", "File URL", "Score", "Comments", "Graded By"]);
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('FOLDER_ID', folder.getId());
  scriptProperties.setProperty('SHEET_ID', ss.getId());
  
  return "Setup Complete! Folder ID: " + folder.getId();
}

/**
 * Helper to get or create the folder
 */
function getOrCreateFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('FOLDER_ID');
  
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {}
  }
  
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    const folder = folders.next();
    props.setProperty('FOLDER_ID', folder.getId());
    return folder;
  }
  
  const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

/**
 * Uploads a file to Drive
 */
function uploadFile(data) {
  try {
    const folder = getOrCreateFolder();
    const blob = Utilities.newBlob(Utilities.base64Decode(data.data), data.mimeType, data.fileName);
    const file = folder.createFile(blob);
    
    return {
      success: true,
      fileId: file.getId(),
      url: file.getUrl(),
      name: file.getName()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Gets list of PDF files
 */
function getDriveFiles() {
  const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) return [];
  
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.PDF);
  const fileList = [];
  
  while (files.hasNext()) {
    const file = files.next();
    fileList.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      created: file.getDateCreated().toLocaleDateString()
    });
  }
  return fileList;
}

/**
 * Gets file content as Base64
 */
function getFileContent(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return {
      success: true,
      data: Utilities.base64Encode(blob.getBytes()),
      name: file.getName(),
      mimeType: file.getMimeType()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Saves grade to Google Sheet
 */
function saveGrade(studentId, score, comment) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];
  
  sheet.appendRow([new Date(), studentId, "", score, comment, "Teacher"]);
  return "Saved successfully!";
}

/**
 * Calls Poe API
 * - Supports Chunking (Accepts partial image lists)
 * - Enforces Strict 1M/1A Grading
 * - Enforces Traditional Chinese
 */
function callPoeAPI(studentImages, solutionImages, studentIndex, modelName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('POE_API_KEY');
  if (!apiKey) {
    return { error: "API Key not found. Please add POE_API_KEY to Script Properties." };
  }
  const apiUrl = "https://api.poe.com/v1/chat/completions";
  
  // --- STRICT MARKING SCHEME PROMPT ---
  const systemPrompt = `
    You are a STRICT Math Teacher grading a specific part of a student's exam.
    
    **MANDATORY SCORING RULES (1M + 1A):**
    For every question, you must identify two specific marks:
    
    1. **M Mark (Method):** - Award **1 mark** if the student shows the correct formula, substitution, or logical step matching the Solution Key.
       - Award **0 marks** if the method is missing, wrong, or skipped.
       
    2. **A Mark (Answer):** - Award **1 mark** ONLY if the final answer is exactly correct (including units/signs).
       - **CRITICAL:** If the Method (M) is wrong, the Answer (A) is automatically **0**. (No "lucky guesses" allowed).
    
    **EXAMPLE SCORING:**
    - Correct Method + Correct Answer = **2/2** (M1 A1)
    - Correct Method + Wrong Answer = **1/2** (M1 A0)
    - Wrong Method + Correct Answer = **0/2** (M0 A0)
    
    **OUTPUT REQUIREMENTS:**
    1. **Language:** All comments must be in **Traditional Chinese (繁體中文)**.
    2. **Format:** Return ONLY valid JSON.
    3. **Content:** Specifically mention "M1" or "A0" in your comments so the student knows where they lost points.
    4. **Scope:** Grade ONLY the questions visible in this image chunk.
    
    **JSON STRUCTURE:**
    {
      "student_name": "Name (if visible)",
      "total_score": "ignored",
      "overall_comment": "Summary in Traditional Chinese.",
      "questions": [
        { "id": "Q1", "score": "X/Y", "comment": "M1 (步驟正確), A0 (計算錯誤)..." },
        { "id": "Q2", "score": "X/Y", "comment": "M1 A1 (全對)" }
      ]
    }
  `;
  
  const userContent = [
    { "type": "text", "text": `Grade this exam chunk (Student ${studentIndex}) using the 1M/1A method.` }
  ];
  
  // 1. Add Solution Images (Limit to 5 to save context window)
  if (solutionImages && solutionImages.length > 0) {
    solutionImages.slice(0, 5).forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
  }
  
  // 2. Add Student Images (Accept ALL images sent by client - Client handles chunking)
  if (Array.isArray(studentImages)) {
    studentImages.forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
  }
  
  const payload = {
    "model": modelName || DEFAULT_BOT,
    "messages": [
      { "role": "system", "content": systemPrompt },
      { "role": "user", "content": userContent }
    ],
    "temperature": 0.1 // Keep it low for strict rule-following
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "Authorization": `Bearer ${apiKey}` },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      return { error: `Poe API Error (${responseCode}): ${responseText}` };
    }
    
    let json;
    try {
      json = JSON.parse(responseText);
    } catch (e) {
      return { error: "Invalid JSON response from Poe." };
    }
    
    if (!json.choices || !json.choices[0].message) {
      return { error: "No content generated." };
    }
    
    // Robust Parsing (Extract JSON from text)
    const textResponse = json.choices[0].message.content;
    const jsonMatch = textResponse.match(/\{[\s\S]*\}/);
    
    if (!jsonMatch) {
      return { error: "No JSON found in response: " + textResponse };
    }
    
    let gradeData;
    try {
      gradeData = JSON.parse(jsonMatch[0]);
    } catch (e) {
      return { error: "JSON Parse Failed: " + e.toString() };
    }

    // --- MATH CORRECTION LOGIC (Re-calculate sum for this chunk) ---
    if (gradeData.questions && Array.isArray(gradeData.questions)) {
      let calculatedObtained = 0;
      let calculatedTotal = 0;
      let hasDenominator = false;

      gradeData.questions.forEach(q => {
        if (q.score) {
          const parts = q.score.toString().split('/');
          const obtained = parseFloat(parts[0]);
          if (!isNaN(obtained)) {
            calculatedObtained += obtained;
          }
          if (parts.length > 1) {
            const possible = parseFloat(parts[1]);
            if (!isNaN(possible)) {
              calculatedTotal += possible;
              hasDenominator = true;
            }
          }
        }
      });

      // Force the total score to be the sum of questions
      if (hasDenominator && calculatedTotal > 0) {
        gradeData.total_score = `${calculatedObtained}/${calculatedTotal}`;
      } else {
        gradeData.total_score = `${calculatedObtained}`;
      }
    }
    // -------------------------------------------------------------

    return gradeData;
    
  } catch (e) {
    console.error("Critical Error in callPoeAPI:", e);
    return { error: "Connection Failed. Details: " + e.toString() };
  }
}

/**
 * Generates a PDF Report from the grading data
 */
function createPdfReport(gradingData) {
  let html = `
    <html>
      <head>
        <style>
          body { font-family: 'Microsoft JhengHei', sans-serif; padding: 40px; line-height: 1.6; }
          .student-section { margin-bottom: 40px; page-break-inside: avoid; border-bottom: 1px dashed #ccc; padding-bottom: 20px; }
          .header { font-size: 1.2em; font-weight: bold; margin-bottom: 10px; }
          .score { font-weight: bold; color: #d93025; }
          .label { font-weight: bold; }
          .question-item { margin-left: 20px; }
        </style>
      </head>
      <body>
  `;
  
  gradingData.forEach((item, index) => {
    const safeName = escapeHtml(item.student_name || "Student " + (index + 1));
    const safeScore = escapeHtml(item.total_score);
    const safeComment = escapeHtml(item.overall_comment);
    
    let qHtml = "";
    if (item.questions && item.questions.length > 0) {
      item.questions.forEach(q => {
        const safeId = escapeHtml(q.id);
        const safeQScore = escapeHtml(q.score);
        const safeQComment = escapeHtml(q.comment);
        qHtml += `<div class="question-item"><strong>${safeId}:</strong> (分數: ${safeQScore}) (評語: ${safeQComment})</div>`;
      });
    } else {
      qHtml = "<div class='question-item'>No specific questions found.</div>";
    }
    html += `
      <div class="student-section">
        <div class="header">學生 (${index + 1}): ${safeName}</div>
        <div><span class="label">分數:</span> <span class="score">${safeScore}</span></div>
        <div><span class="label">整體評語:</span> ${safeComment}</div>
        <div><span class="label">細項評語:</span></div>
        ${qHtml}
      </div>
    `;
  });
  
  html += `</body></html>`;
  
  const blob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);
  blob.setName(`Math_Exam_Report_${new Date().toLocaleDateString()}.pdf`);
  
  const folder = getOrCreateFolder();
  const file = folder.createFile(blob);
  
  return file.getUrl();
}

/**
 * Helper to escape HTML to prevent XSS
 */
function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
