/**
 * Math Marking System - Backend Code
 */

// CONFIGURATION
const APP_NAME = "Math Marking System";
const DRIVE_FOLDER_NAME = "Math_Marking_System_Uploads";
const SHEET_NAME = "Math_Scores";
const DEFAULT_BOT = "Claude-Opus-4.5"; 

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
 * Initial Setup
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
 * Helper to get/create folder
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
 * Uploads file
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
 * Get files list
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
 * Get file content
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
 * Save Grade
 */
function saveGrade(studentId, score, comment) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];
  
  sheet.appendRow([new Date(), studentId, "", score, comment, "Teacher"]);
  return "Saved successfully!";
}

/**
 * Calls Poe API - WITH IMPROVED SYSTEM PROMPT & MATH CORRECTION
 */
function callPoeAPI(studentImages, solutionImages, studentIndex, modelName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('POE_API_KEY');
  if (!apiKey) {
    return { error: "API Key not found. Please add POE_API_KEY to Script Properties." };
  }
  const apiUrl = "https://api.poe.com/v1/chat/completions";
  
  // --- UPDATED SYSTEM PROMPT: FORCE "THINKING" ---
  const systemPrompt = `
    You are a STRICT math teacher grading a student's exam.
    
    **INPUTS:**
    1. **STUDENT WORK**: Images of handwritten math.
    2. **SOLUTION KEY**: Images of the marking scheme.
    
    **YOUR PROCESS (CRITICAL):**
    Do not output JSON immediately. You must "think" before you grade.
    For each question, follow these steps internally:
    1. **TRANSCRIBE**: Write out what the student actually wrote (e.g., "Student wrote x^2 + 5").
    2. **COMPARE**: Check that specific step against the Solution Key image.
    3. **VERIFY**: Did they simply copy the final answer? Did they show the "M" (Method) steps?
    4. **DECIDE**: Assign marks based strictly on the scheme.
    
    **GRADING RULES:**
    - **M Marks**: Award ONLY if the method is visible and correct.
    - **A Marks**: Award ONLY if the final answer matches the solution exactly.
    - If the student skips steps required by the Solution Key, deduct the M mark.
    
    **OUTPUT FORMAT:**
    After your analysis, output the results in this **EXACT JSON** format inside a code block:
    
    \`\`\`json
    {
      "student_name": "Student Name",
      "overall_comment": "Summary...",
      "questions": [
        { "id": "Q1", "score": "X/Y", "comment": "Missing step 2..." },
        { "id": "Q2", "score": "X/Y", "comment": "Correct." }
      ]
    }
    \`\`\`
  `;
  // ------------------------------------------------
  
  const userContent = [
    { "type": "text", "text": `Grade this student's work (Student ${studentIndex}).` }
  ];
  
  // 1. Add Solution Images
  if (solutionImages && solutionImages.length > 0) {
    userContent.push({ "type": "text", "text": "--- OFFICIAL SOLUTION KEY START ---" });
    solutionImages.forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
    userContent.push({ "type": "text", "text": "--- OFFICIAL SOLUTION KEY END ---" });
  } else {
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
    "model": modelName || DEFAULT_BOT,
    "messages": [
      { "role": "system", "content": systemPrompt },
      { "role": "user", "content": userContent }
    ],
    "temperature": 0.3
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
    
    // Robust Parsing
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

    // MATH CORRECTION LOGIC
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

      if (hasDenominator && calculatedTotal > 0) {
        gradeData.total_score = `${calculatedObtained}/${calculatedTotal}`;
      } else {
        gradeData.total_score = `${calculatedObtained}`;
      }
    }

    return gradeData;
    
  } catch (e) {
    console.error("Critical Error in callPoeAPI:", e);
    return { error: "Connection Failed. Details: " + e.toString() };
  }
}

/**
 * PDF Report Generator
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

function escapeHtml(text) {
  if (!text) return "";
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}
