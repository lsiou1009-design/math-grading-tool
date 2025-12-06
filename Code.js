/**
 * Math Marking System - Backend Code
 * 
 * INSTRUCTIONS:
 * 1. Create a new Google Apps Script project at https://script.google.com
 * 2. Paste this code into 'Code.gs'
 * 3. Create 'index.html', 'JavaScript.html', and 'Stylesheet.html' and paste their respective content.
 * 4. Run the 'setup()' function once to initialize the Drive Folder and Sheet.
 * 5. IMPORTANT: Go to Project Settings (Gear Icon) > Script Properties > Add new property:
 *    Property: GEMINI_API_KEY
 *    Value: Your-Actual-API-Key
 * 6. Deploy as Web App (Execute as: Me, Who has access: Anyone).
 */
// CONFIGURATION
const APP_NAME = "Math Marking System";
const DRIVE_FOLDER_NAME = "Math_Marking_System_Uploads";
const SHEET_NAME = "Math_Scores";
const POE_BOT_NAME = "Claude-Opus-4.5"; // Using Claude Opus 4.5 on Poe
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
 * Initial Setup: Creates Drive Folder and Spreadsheet if they don't exist.
 */
function setup() {
  // 1. Create Drive Folder
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
  
  // 2. Create Spreadsheet
  const files = DriveApp.getFilesByName(SHEET_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
  }
  
  // Setup Sheet Headers
  let sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Student ID", "File URL", "Score", "Comments", "Graded By"]);
  }
  
  // Save IDs to Properties for easy access later
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('FOLDER_ID', folder.getId());
  scriptProperties.setProperty('SHEET_ID', ss.getId());
  
  return "Setup Complete! Folder ID: " + folder.getId();
}
/**
 * Helper to get or create the folder if ID is missing
 */
function getOrCreateFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('FOLDER_ID');
  
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // ID might be invalid/deleted, fall through to search/create
    }
  }
  
  // Search by name
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    const folder = folders.next();
    props.setProperty('FOLDER_ID', folder.getId());
    return folder;
  }
  
  // Create new
  const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}
/**
 * Uploads a file to the specific Drive folder
 */
function uploadFile(data) {
  try {
    const folder = getOrCreateFolder();
    
    const blob = Utilities.newBlob(Utilities.base64Decode(data.data), data.mimeType, data.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
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
 * Gets list of PDF files from the folder
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
 * Gets the base64 content of a file from Drive
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
 * Saves the grade to the Google Sheet
 */
function saveGrade(studentId, score, comment) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];
  
  sheet.appendRow([new Date(), studentId, "", score, comment, "Teacher"]);
  return "Saved successfully!";
}
/**
 * Calls Poe API (OpenAI Compatible) to grade the work
 * Receives a base64 IMAGE of the student's work.
 */
/**
 * Calls Poe API (OpenAI Compatible) to grade the work
 * Receives an ARRAY of base64 IMAGES of the student's work AND the solution.
 */
function callPoeAPI(studentImages, solutionImages, studentIndex) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('POE_API_KEY');
  if (!apiKey) {
    return { error: "API Key not found. Please add POE_API_KEY to Script Properties." };
  }
  const apiUrl = "https://api.poe.com/v1/chat/completions";
  
  const systemPrompt = `
    You are a STRICT math teacher grading a student's exam.
    
    **INPUTS:**
    1. **STUDENT WORK**: Images of the student's handwritten answers.
    2. **SOLUTION KEY**: Images of the official marking scheme.
    
    **MARKING SCHEME INSTRUCTIONS (M/A Marks):**
    - **M (Method) Marks**: Award these if the student shows the correct method or step.
    - **A (Answer) Marks**: Award these ONLY if the final answer is EXACTLY correct as per the solution.
    - **ZERO TOLERANCE for A Marks**: If the student's final answer does not match the solution (and is not a mathematically equivalent form like 0.5 vs 1/2), you MUST NOT award the A mark.
    - **Follow the Solution Key STRICTLY**. If the solution says "M1 A1", look for that specific step and answer.
    
    **CRITICAL GRADING RULES:**
    1. **Do not be generous.** If a step is missing or wrong, deduct the mark.
    2. **Check every question.** Do not skip questions.
    3. **If the answer is wrong, the A mark is 0.** No exceptions.
    
    **YOUR TASK:**
    1. **TRANSCRIBE**: Read the student's work.
    2. **COMPARE**: Check against the provided Solution Key.
    3. **GRADE**: Assign a score based on the M/A marks.
    4. **COMMENT**: Write a short comment in **Traditional Chinese (繁體中文)** explaining where marks were lost (e.g., "Missing M1 for substitution").
    
    **OUTPUT REQUIREMENTS:**
    - **Format**: You MUST return the exact JSON structure below.
    
    **JSON STRUCTURE:**
    {
      "student_name": "Name found on paper (or Student ${studentIndex})",
      "total_score": "X/100",
      "overall_comment": "Summary of performance (max 50 words)",
      "questions": [
        { "id": "Q1", "score": "X/Y", "comment": "Feedback on M/A marks (max 20 words)" },
        { "id": "Q2", "score": "X/Y", "comment": "Feedback on M/A marks (max 20 words)" }
        // ... add more questions as found
      ]
    }
    
    If the page is blank, score 0 and return empty questions array.
    RETURN ONLY JSON. NO MARKDOWN.
  `;
  // Construct User Message
  const userContent = [
    {
      "type": "text",
      "text": `Grade this student's work (Student ${studentIndex}).`
    }
  ];
  // 1. Add Solution Images (Context)
  if (solutionImages && solutionImages.length > 0) {
    console.log(`[Poe API] Adding ${solutionImages.length} solution images to prompt.`);
    userContent.push({ "type": "text", "text": "--- OFFICIAL SOLUTION KEY START ---" });
    solutionImages.forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
    userContent.push({ "type": "text", "text": "--- OFFICIAL SOLUTION KEY END ---" });
  } else {
    console.warn("[Poe API] NO SOLUTION IMAGES PROVIDED!");
    userContent.push({ "type": "text", "text": "WARNING: No Solution Key provided. Grade based on general mathematical correctness." });
  }
  // 2. Add Student Images
  userContent.push({ "type": "text", "text": "--- STUDENT WORK START ---" });
  if (Array.isArray(studentImages)) {
    studentImages.forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
  } else {
    userContent.push({
      "type": "image_url",
      "image_url": { "url": `data:image/jpeg;base64,${studentImages}` }
    });
  }
  userContent.push({ "type": "text", "text": "--- STUDENT WORK END ---" });
  const payload = {
    "model": POE_BOT_NAME,
    "messages": [
      {
        "role": "system",
        "content": systemPrompt
      },
      {
        "role": "user",
        "content": userContent
      }
    ],
    "temperature": 0.3
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": `Bearer ${apiKey}`
    },
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
    // Extract text from Poe response
    const textResponse = json.choices[0].message.content;
    
    // Attempt to parse JSON from the text response
    const cleanJson = textResponse.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(cleanJson);
    
  } catch (e) {
    return { error: "API Error: " + e.toString() };
  }
}
/**
 * Helper to escape HTML to prevent XSS in PDF generation
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
/**
 * Generates a PDF Report from the grading data
 * Format:
 * 學生 (1): Name
 * 分數:
 * 整體評語:
 * 細項評語:
 * Q1: (分數) (評語)
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
    // Sanitize all inputs before putting them into HTML
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
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}
