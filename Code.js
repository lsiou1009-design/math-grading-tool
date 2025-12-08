/**
 * Math Marking System - Backend Code
 * Features: Strict Folder Linking, Chunking Support, POE API Integration
 * Fixed: Dashboard "No files found" sync issue
 */

// ==========================================
// CONFIGURATION
// ==========================================
const APP_NAME = "Math Marking System";
const SHEET_NAME = "Math_Scores";
const DEFAULT_BOT = "GPT-5.1"; 

// ==========================================
// WEB APP SERVING
// ==========================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// INITIAL SETUP
// ==========================================
function setup() {
  // 測試連線
  const folder = getOrCreateFolder();
  
  // 設定或建立試算表
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
  scriptProperties.setProperty('SHEET_ID', ss.getId());
  
  return "Setup Complete! Connected to Folder: " + folder.getName();
}

// ==========================================
// DRIVE & FILE HANDLING (STRICT MODE)
// ==========================================

// [核心修正] 嚴格讀取屬性中的 ID，不再自動建立新資料夾
function getOrCreateFolder() {
  const props = PropertiesService.getScriptProperties();
  const targetId = props.getProperty('FOLDER_ID');

  // 安全檢查 1: 屬性是否存在
  if (!targetId) {
    throw new Error("❌ 系統錯誤：找不到 'FOLDER_ID' 屬性。請先確認您已執行設定程序 (Step 1)。");
  }

  // 安全檢查 2: 資料夾是否有效
  try {
    const folder = DriveApp.getFolderById(targetId);
    return folder;
  } catch (e) {
    throw new Error("❌ 資料夾存取失敗：您設定的 FOLDER_ID (" + targetId + ") 無效、被刪除或無權限。");
  }
}

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
 * Gets list of Files grouped by Student Paper
 * FIX: Diagnosis mode to trace why files are disappearing
 */
function getDriveFiles() {
  try {
    const folder = getOrCreateFolder(); 
    console.log("📂 [Step 1] Accessing Folder: " + folder.getName());

    // 使用 searchFiles 確保能找到最新檔案
    const files = folder.searchFiles("trashed = false");
    const allFiles = [];

    // 1. 收集檔案
    while (files.hasNext()) {
      const file = files.next();
      // 轉換成簡單物件，避免 Date 物件造成序列化問題
      allFiles.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        mimeType: file.getMimeType(), // 確保是字串
        created: file.getDateCreated().getTime() // 轉成 timestamp 數字，避免傳輸錯誤
      });
    }

    console.log(`📂 [Step 2] Found ${allFiles.length} raw files.`);
    
    // 如果這裡就是 0，那 searchFiles 有問題 (但根據你的 log，這裡應該不是 0)
    if (allFiles.length === 0) return [];

    // 2. 排序
    allFiles.sort((a, b) => b.created - a.created);

    // 3. 分組邏輯 (重點檢查區)
    const fileMap = {};
    
    allFiles.forEach(file => {
      const name = file.name;
      const type = file.mimeType;
      
      // 判定是否為報告或成績單
      const isReport = name.includes("_Report_") && type === "application/pdf";
      const isCsv = name.includes("_Grades_") && (type === "text/csv" || type === "application/vnd.ms-excel");

      let baseName;

      if (isReport || isCsv) {
        // 是子檔案 (Child)
        const separator = isReport ? "_Report_" : "_Grades_";
        baseName = name.split(separator)[0];
        
        if (!fileMap[baseName]) fileMap[baseName] = { children: [] };
        
        fileMap[baseName].children.push({
          ...file,
          type: isReport ? 'PDF Report' : 'CSV Grades',
          displayDate: formatDate(new Date(file.created))
        });
        
        console.log(`   ➡️ Classified [${name}] as CHILD of [${baseName}]`);
      } else {
        // 是主檔案 (Parent)
        // 移除副檔名邏輯
        baseName = name.replace(/\.[^/.]+$/, ""); 
        
        if (!fileMap[baseName]) fileMap[baseName] = { children: [] };
        
        if (!fileMap[baseName].parent) {
          fileMap[baseName].parent = {
            ...file,
            displayDate: formatDate(new Date(file.created))
          };
           console.log(`   ➡️ Classified [${name}] as PARENT [${baseName}]`);
        } else {
           console.log(`   ⚠️ Duplicate Parent ignored: [${name}]`);
        }
      }
    });

    // 4. 轉換為列表
    const result = [];
    Object.keys(fileMap).forEach(key => {
      const item = fileMap[key];
      
      if (item.parent) {
        // 正常情況：有主檔案
        result.push({
          ...item.parent,
          generatedFiles: item.children
        });
      } else if (item.children.length > 0) {
        // 孤兒檔案：主檔案不見了，但有報告
        result.push({
          id: item.children[0].id,
          name: key + " [Source File Missing]",
          url: "#",
          mimeType: "application/pdf", 
          displayDate: item.children[0].displayDate,
          generatedFiles: item.children,
          isOrphan: true 
        });
      }
    });

    console.log(`📂 [Step 3] Grouping Complete. Final count: ${result.length}`);
    return result;

  } catch (e) {
    console.error("❌ Critical Error in getDriveFiles: " + e.toString());
    // 發生錯誤時傳回空陣列，避免前端卡死
    throw new Error("Backend Error: " + e.message); 
  }
}

function formatDate(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
}

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

// ==========================================
// GRADING LOGIC (POE API)
// ==========================================
function saveGrade(studentId, score, comment) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  let ss;
  if (sheetId) {
      try { ss = SpreadsheetApp.openById(sheetId); } catch(e) {}
  }
  if (!ss) {
     const files = DriveApp.getFilesByName(SHEET_NAME);
     if (files.hasNext()) ss = SpreadsheetApp.open(files.next());
     else ss = SpreadsheetApp.create(SHEET_NAME);
  }

  const sheet = ss.getSheets()[0];
  sheet.appendRow([new Date(), studentId, "", score, comment, "Teacher"]);
  return "Saved successfully!";
}

function callPoeAPI(studentImages, solutionImages, studentIndex, modelName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('POE_API_KEY');
  if (!apiKey) {
    return { error: "API Key not found. Please add POE_API_KEY to Script Properties." };
  }
  const apiUrl = "https://api.poe.com/v1/chat/completions";
  
  // --- STRICT MARKING SCHEME PROMPT ---
  const systemPrompt = `
    You are a STRICT Math Teacher.
    
    **TASK:** Compare the Student's Work (chunk) against the Solution Key.
    
    **RULE 0: NON-CHRONOLOGICAL ORDER (CRITICAL)**
    - The Student's Work may be in **ANY ORDER** (e.g., Q5 may appear before Q1).
    - The Student's Page Count may NOT match the Solution Key Page Count.
    - You must **SEARCH the entire provided Solution Key** to find the specific Question ID that corresponds to the student's work.
    - Do NOT assume the first page of Student Work matches the first page of the Solution.

    **RULE 1: SEPARATE ALL SUB-QUESTIONS**
    - You MUST output every sub-question as a separate JSON entry (e.g., "Q1a", "Q1b", "Q5a(i)").
    - **Do NOT group them.**
    
    **RULE 2: IDENTIFY QUESTIONS STRICTLY**
    - Only grade questions that appear in the provided Solution Key.
    - If you cannot find the question in the Solution Key images, ignore it (do not hallucinate).

    **RULE 3: HANDLING BLANKS**
    - If a question EXISTS in the Solution Key but is blank/skipped by student:
      - Score: "0/Total"
      - Comment: "未作答 (Blank)"
    
    **RULE 4: STRICT 1M / 1A GRADING**
    - **M Mark (Method):** 1 mark if method is correct.
    - **A Mark (Answer):** 1 mark if FINAL ANSWER matches EXACTLY.
    - **NEGATIVE LOGIC:** If Answer != Solution, A mark is 0.

    **OUTPUT FORMAT:**
    - Language: **Traditional Chinese (繁體中文)** ONLY.
    - Format: Valid JSON ONLY.
    
    **JSON STRUCTURE:**
    {
      "student_name": "Name",
      "total_score": "ignored",
      "overall_comment": "Summary.",
      "questions": [
        { "id": "Q1a", "score": "2/2", "comment": "1M 1A (全對)" },
        { "id": "Q5", "score": "0/3", "comment": "未作答 (Blank)" }
      ]
    }
  `;
  
  const userContent = [
    { "type": "text", "text": `Grade this exam chunk (Student ${studentIndex}). Search Solution Key for matching questions. Separate sub-questions.` }
  ];
  
  // 1. Add Solution Images 
  if (solutionImages && solutionImages.length > 0) {
    solutionImages.slice(0, 15).forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
  }
  
  // 2. Add Student Images
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
    "temperature": 0.1 
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

    // --- MATH CORRECTION LOGIC ---
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

// ==========================================
// REPORT GENERATION (PDF & CSV)
// ==========================================
function createPdfReport(gradingData, sourceFileName) {
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
  
  const cleanSourceName = sourceFileName ? sourceFileName.replace(/\.pdf$/i, "") : "Exam";
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmm");
  blob.setName(`${cleanSourceName}_Report_${timestamp}.pdf`);
  
  const folder = getOrCreateFolder();
  const file = folder.createFile(blob);
  
  return file.getUrl();
}

function saveCsvReport(csvContent, sourceFileName) {
  try {
    const folder = getOrCreateFolder();
    const cleanSourceName = sourceFileName ? sourceFileName.replace(/\.pdf$/i, "") : "Exam";
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmm");
    const fileName = `${cleanSourceName}_Grades_${timestamp}.csv`;
    
    const blob = Utilities.newBlob(csvContent, MimeType.CSV, fileName);
    const file = folder.createFile(blob);
    
    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
