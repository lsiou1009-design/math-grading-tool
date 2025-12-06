/**
 * Math Marking System - Backend Code
 * Optimized for Strict 1M/1A Grading + Chunking Support
 */

const APP_NAME = "Math Marking System";
const DRIVE_FOLDER_NAME = "Math_Marking_System_Uploads";
const SHEET_NAME = "Math_Scores";
const DEFAULT_BOT = "GPT-5.1"; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate().setTitle(APP_NAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setup() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  const files = DriveApp.getFilesByName(SHEET_NAME);
  let ss = files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.create(SHEET_NAME);
  let sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(["Timestamp", "Student ID", "File URL", "Score", "Comments", "Graded By"]);
  PropertiesService.getScriptProperties().setProperty('FOLDER_ID', folder.getId());
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  return "Setup Complete!";
}

function getOrCreateFolder() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function uploadFile(data) {
  try {
    const folder = getOrCreateFolder();
    const blob = Utilities.newBlob(Utilities.base64Decode(data.data), data.mimeType, data.fileName);
    const file = folder.createFile(blob);
    return { success: true, fileId: file.getId(), url: file.getUrl(), name: file.getName() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getDriveFiles() {
  const folder = getOrCreateFolder();
  const files = folder.getFilesByType(MimeType.PDF);
  const list = [];
  while (files.hasNext()) {
    const f = files.next();
    list.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), created: f.getDateCreated().toLocaleDateString() });
  }
  return list;
}

function getFileContent(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return { success: true, data: Utilities.base64Encode(file.getBlob().getBytes()), name: file.getName() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveGrade(studentId, score, comment) {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  ss.getSheets()[0].appendRow([new Date(), studentId, "", score, comment, "Teacher"]);
  return "Saved!";
}

/**
 * Calls Poe API - STRICT 1M / 1A GRADING
 */
function callPoeAPI(studentImages, solutionImages, studentIndex, modelName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('POE_API_KEY');
  if (!apiKey) return { error: "API Key missing." };
  
  const apiUrl = "https://api.poe.com/v1/chat/completions";
  
  // --- STRICT MARKING SCHEME PROMPT ---
  const systemPrompt = `
    You are a STRICT Math Teacher. You are grading a student's exam chunk.
    
    **MANDATORY SCORING RULES (1M + 1A):**
    For every question, you must identify two specific marks:
    
    1. **M Mark (Method):** - Award **1 mark** if the student shows the correct formula, substitution, or logical step matching the Solution Key.
       - Award **0 marks** if the method is missing, wrong, or skipped.
       
    2. **A Mark (Answer):** - Award **1 mark** ONLY if the final answer is exactly correct (including units/signs).
       - **CRITICAL:** If the Method (M) is wrong, the Answer (A) is automatically **0**. (No "lucky guesses" allowed).
    
    **EXAMPLE:**
    - Correct Method + Correct Answer = **2/2**
    - Correct Method + Wrong Answer = **1/2** (Award M1, A0)
    - Wrong Method + Correct Answer = **0/2** (Award M0, A0)
    
    **OUTPUT REQUIREMENTS:**
    1. **Language:** All comments must be in **Traditional Chinese (繁體中文)**.
    2. **Format:** Return ONLY valid JSON.
    3. **Content:** Specifically mention "M1" or "A0" in your comments so the student knows where they lost points.
    
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
  
  const userContent = [{ "type": "text", "text": `Grade this exam chunk (Student ${studentIndex}) using the 1M/1A method.` }];
  
  // Add Solution Images (Limit to 5)
  if (solutionImages && solutionImages.length > 0) {
    solutionImages.slice(0, 5).forEach(img => {
      userContent.push({ "type": "image_url", "image_url": { "url": `data:image/jpeg;base64,${img}` } });
    });
  }
  
  // Add Student Images (All chunks)
  if (Array.isArray(studentImages)) {
    studentImages.forEach(img => {
      userContent.push({ "type": "image_url", "image_url": { "url": `data:image/jpeg;base64,${img}` } });
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
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    if (code !== 200) return { error: `Poe API Error (${code}): ${text}` };
    
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return { error: "No JSON found in response." };
    
    return JSON.parse(jsonMatch[0]);
    
  } catch (e) {
    console.error("API Error:", e);
    return { error: "Connection Failed: " + e.toString() };
  }
}

function createPdfReport(gradingData) {
  let html = `<html><head><style>body{font-family:'Microsoft JhengHei',sans-serif;padding:40px;}.header{font-weight:bold;font-size:1.2em;margin-bottom:10px;}.score{color:#d93025;font-weight:bold;}.question-item{margin-left:20px;}</style></head><body>`;
  
  gradingData.forEach((item, index) => {
    html += `<div class="student-section">
      <div class="header">學生: ${item.student_name}</div>
      <div><span class="label">分數:</span> <span class="score">${item.total_score}</span></div>
      <div><span class="label">總評:</span> ${item.overall_comment}</div>
      ${(item.questions || []).map(q => 
        `<div class="question-item"><strong>${q.id}:</strong> ${q.score} (${q.comment})</div>`
      ).join('')}
    </div><hr>`;
  });
  
  html += `</body></html>`;
  const blob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);
  blob.setName(`Math_Report_${new Date().getTime()}.pdf`);
  const file = getOrCreateFolder().createFile(blob);
  return file.getUrl();
}
