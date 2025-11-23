// Code.gs
const USE_GROQ_FOR_TEXT  = true;
const TARGET_LANGUAGE = 'th'; // Target language for multilingual search, e.g., 'th' for Thai
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const RESPONSE_SHEET_NAME = "House-Expense-Form";
const HELPER_SHEET_NAME = "Helper_Lists";
const DRIVE_FOLDER_NAME = "House-Expense-AI"; // For uploaded receipts
const DRIVE_FOLDER_EMAIL_ATTACHMENT = "Attachments from Email";
const DRIVE_FOLDER_GENERATED_PDF = "Generated PDF Reports"
const EXPENSE_SUMMARY_SHEET_NAME = "Expense_Summary";
const ORCHESTRATOR_LOG_SHEET_NAME = "AI_Orchestrator_Tasks";


// modal name
const GROQ_LLAMA_MODEL_NAME= 'llama-3.3-70b-versatile';
const GROQ_QWEN_MODEL_NAME= 'qwen/qwen3-32b';
const GROQ_MODEL_NAME = GROQ_QWEN_MODEL_NAME;
// const GEMINI_MODEL_NAME= "gemini-2.5-flash".trim(); // For V1
const GEMINI_MODEL_NAME= "gemini-2.5-flash-preview-09-2025".trim(); // for v1Beta  "gemini-2.5-flash"; "gemini-1.5-pro-latest"; 

//api endpoints
const GROQ_TEX_ENDPOINT = "https://api.groq.com/openai/v1/chat/completions";
const GROQ_AUDIO_ENDPOINT = "https://api.groq.com/openai/v1/audio/transcriptions";
const GEMINI_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL_NAME}:generateContent`;
// const GEMINI_ENDPOINT = `https://generativelanguage.googleapis.com/v1/models/${GEMINI_MODEL_NAME}:generateContent`; 
const ASSEMBLY_AUDIO_UPLOAD_ENDPOINT = "https://api.assemblyai.com/v2/upload";
const ASSEMBLY_AUDIO_TRANSCRIPT_ENDPOINT = "https://api.assemblyai.com/v2/transcript";

// Column mapping for Expense_Summary (1-indexed)
const PRIMARY_CATEGORY_COL = 14; // N
const SUBCATEGORY_COL = 15;      // O
const OCR_TEXT_COL = 21;         // U
const AI_CAT_COL = 26;           // Z
const AI_SUBCAT_COL = 27;        // AA
const AI_AMOUNT_COL = 28;        // AB
const AI_VENDOR_COL = 29;        // AC
const AI_DATE_COL = 30;          // AD

const DEFAULT_CURRENCY = 'THB';
const DEFAULT_CURRENCY_SYMBOL = '฿';

// Audio recording config
const MAX_FILE_SIZE_MB = 10;

// Email Query for Expenses days
const EMAIL_QUERY_DAYS = 15;
const PROCESSED_LABEL_NAME = "AI-Expense-Logged";

// Column mapping for RESPONSE_SHEET_NAME (0-indexed for .getValues() array)
const RESPONSE_COLS = {
    TIMESTAMP: 0,       // A
    DATE: 1,            // B
    COST_CENTER: 2,     // C
    PRIMARY_CAT: 3,     // D
    SUB_CAT: 4,         // E
    VENDOR: 5,          // F
    DESCRIPTION: 6,     // G
    AMOUNT: 7,          // H
    PAYMENT_METHOD: 8,  // I
    RECEIPT_URL: 9,     // J
    OCR_TEXT: 10,       // K
    EMAIL: 11,          // L
    NOTES: 12           // M
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('WebApp.html');
  template.serverSpreadsheetId = SPREADSHEET_ID; // Add this line
  template.currencySymbol = DEFAULT_CURRENCY_SYMBOL;
  template.currencyCode = DEFAULT_CURRENCY;
  return template
    .evaluate()
    .setTitle('AI Expense Logger')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUserProfileData() {
  try {
    const user = Session.getActiveUser();
    const userEmail = user.getEmail();
    const summaryStats = getUserExpenseSummaryStats(userEmail); // This already works

    // --- Fetch first/last entry dates for the SPECIFIC user ---
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSE_SHEET_NAME);
    const allData = sheet.getRange("B2:L" + sheet.getLastRow()).getValues();
    const userDateRows = allData.filter(row => row[10] === userEmail && row[0] instanceof Date);
    
    if (userDateRows.length > 0) {
      userDateRows.sort((a, b) => a[0] - b[0]); // Sort by date
      summaryStats.firstEntryDate = Utilities.formatDate(userDateRows[0][0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      summaryStats.lastEntryDate = Utilities.formatDate(userDateRows[userDateRows.length - 1][0], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      summaryStats.firstEntryDate = 'N/A';
      summaryStats.lastEntryDate = 'N/A';
    }
    
    const profileData = getInitialProfileData(); // Get the static data
    profileData.stats = summaryStats; // Combine them
    
    return profileData;
  } catch (e) {
    Logger.log(`Error in getUserProfileData: ${e.message}`);
    return { error: e.message };
  }
}

function getInitialProfileData() {
  try {
    const user = Session.getActiveUser();
    const helperSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HELPER_SHEET_NAME);
    const mastersRange = helperSheet.getDataRange().getValues();
    const headers = mastersRange.shift();
    const mastersData = {};
    
    const userEmail = user.getEmail();
    const helperDataForKeys = getHelperListsData();
    const globalApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    let apiKeyInUse = { 
      key: "Not Set", 
      isUserSpecific: false, 
      platform: "N/A",
      purpose: "AI Expense Analysis & Q&A",
      extractionModel: "Gemini 1.5 Flash (Primary), Groq Llama 3.1 (Fallback)",
      qaModel: "Groq Llama 3.1 (Primary), Gemini 1.5 Flash (Fallback)"
    };
    if (helperDataForKeys.userApiKeys && helperDataForKeys.userApiKeys[userEmail.toLowerCase()]) {
      apiKeyInUse.key = helperDataForKeys.userApiKeys[userEmail.toLowerCase()];
      apiKeyInUse.isUserSpecific = true;
      apiKeyInUse.platform = "User Provided (Gemini/Groq)";
    } else if (globalApiKey) {
      apiKeyInUse.key = globalApiKey;
      apiKeyInUse.isUserSpecific = false;
      apiKeyInUse.platform = "App Default (Google Gemini)";
    }

    // --- CHANGED LOGIC FOR CATEGORY MAPPINGS ---
    const primaryCatHeader = "Primary Category";
    const subCatHeader = "Subcategory";
    const primaryCatIndex = headers.indexOf(primaryCatHeader);
    const subCatIndex = headers.indexOf(subCatHeader);
    const categoryMappings = [];

    if (primaryCatIndex > -1 && subCatIndex > -1) {
      mastersRange.forEach(row => {
        if (row[primaryCatIndex]) { // Only need to check if primary category exists
          categoryMappings.push({
            primary: row[primaryCatIndex],
            subs: row[subCatIndex] || '' // The comma-separated string
          });
        }
      });
      mastersData["Category Mappings"] = categoryMappings;
    }
    // --- END CHANGED LOGIC ---
    
    headers.forEach((header, index) => {
      if (header && header !== primaryCatHeader && header !== subCatHeader) {
        mastersData[header] = mastersRange.map(row => row[index]).filter(String);
      }
    });

    return {
      userName: user.getUsername(),
      userEmail: userEmail,
      apiKeyInUse: apiKeyInUse,
      masters: mastersData
    };
  } catch (e) {
    Logger.log(`Error in getInitialProfileData: ${e.message}`);
    return { error: e.message };
  }
}

function addMasterRecord(masterName, newValue) {
  try {
    const helperSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HELPER_SHEET_NAME);
    const headers = helperSheet.getRange(1, 1, 1, helperSheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(masterName);

    if (colIndex === -1) {
      return { success: false, error: `Master list "${masterName}" not found.` };
    }

    const col = colIndex + 1;
    const lastRow = helperSheet.getRange(helperSheet.getMaxRows(), col).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    helperSheet.getRange(lastRow + 1, col).setValue(newValue);
    
    CacheService.getScriptCache().remove('HELPER_DATA_CACHE'); // IMPORTANT: Invalidate cache
    return { success: true };
  } catch (e) {
    Logger.log(`Error adding master record: ${e.message}`);
    return { success: false, error: e.message };
  }
}

function updateCategoryMapping(originalPrimary, newPrimary, newSubsString) {
  try {
    const helperSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HELPER_SHEET_NAME);
    const primaryCatColumn = 4; // Column D
    const subCatColumn = 5;     // Column E

    const textFinder = helperSheet.getRange(2, primaryCatColumn, helperSheet.getLastRow()).createTextFinder(originalPrimary).matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (!foundCell) {
      return { success: false, error: `Could not find the original category "${originalPrimary}" to update.` };
    }

    const row = foundCell.getRow();
    helperSheet.getRange(row, primaryCatColumn).setValue(newPrimary);
    helperSheet.getRange(row, subCatColumn).setValue(newSubsString);

    CacheService.getScriptCache().remove('HELPER_DATA_CACHE'); // IMPORTANT: Invalidate cache
    return { success: true };
  } catch (e) {
    Logger.log(`Error in updateCategoryMapping: ${e.message}`);
    return { success: false, error: e.message };
  }
}

function updateMasterItem(masterName, originalValue, newValue) {
  try {
    const helperSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HELPER_SHEET_NAME);
    const textFinder = helperSheet.createTextFinder(originalValue).matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (foundCell) {
      const headerRow = helperSheet.getRange(1, 1, 1, helperSheet.getLastColumn()).getValues()[0];
      const colIndex = headerRow.indexOf(masterName);
      if (foundCell.getColumn() === colIndex + 1) {
        foundCell.setValue(newValue);
        CacheService.getScriptCache().remove('HELPER_DATA_CACHE'); // IMPORTANT: Invalidate cache
        return { success: true };
      }
    }
    return { success: false, error: "Item not found in the correct master list." };
  } catch (e) {
    Logger.log(`Error updating master item: ${e.message}`);
    return { success: false, error: e.message };
  }
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Sets up initial data validations for primary categories and other dropdowns.
 * Run this function manually once after setting up the Helper_Lists sheet.
 */
function setupInitialValidations() {
  const expenseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPENSE_SUMMARY_SHEET_NAME);
  const helperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HELPER_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!expenseSheet || !helperSheet) {
    ui.alert("Error", "Ensure 'Expense_Summary' and 'Helper_Lists' sheets exist.", ui.ButtonSet.OK);
    return;
  }

  // Primary Category Validation (Column N)
  const primaryCategoriesRange = helperSheet.getRange("D2:D" + helperSheet.getLastRow()).getValues();
  const uniquePrimaryCategories = [...new Set(primaryCategoriesRange.filter(String).map(row => row[0]))];
  if (uniquePrimaryCategories.length > 0) {
    const primaryCatRule = SpreadsheetApp.newDataValidation().requireValueInList(uniquePrimaryCategories).setAllowInvalid(false).build();
    expenseSheet.getRange(`N2:N${expenseSheet.getMaxRows()}`).setDataValidation(primaryCatRule);
    Logger.log("Primary Category validation applied to Column N.");
  }

  // Cost Center Validation (Column M)
  const costCentersRange = helperSheet.getRange("A2:A" + helperSheet.getLastRow()).getValues();
  const uniqueCostCenters = [...new Set(costCentersRange.filter(String).map(row => row[0]))];
   if (uniqueCostCenters.length > 0) {
    const costCenterRule = SpreadsheetApp.newDataValidation().requireValueInList(uniqueCostCenters).setAllowInvalid(false).build();
    expenseSheet.getRange(`M2:M${expenseSheet.getMaxRows()}`).setDataValidation(costCenterRule);
    Logger.log("Cost Center validation applied to Column M.");
  }

  // Payment Method Validation (Column S)
  const paymentMethodsRange = helperSheet.getRange("B2:B" + helperSheet.getLastRow()).getValues();
  const uniquePaymentMethods = [...new Set(paymentMethodsRange.filter(String).map(row => row[0]))];
  if (uniquePaymentMethods.length > 0) {
    const paymentMethodRule = SpreadsheetApp.newDataValidation().requireValueInList(uniquePaymentMethods).setAllowInvalid(false).build();
    expenseSheet.getRange(`S2:S${expenseSheet.getMaxRows()}`).setDataValidation(paymentMethodRule);
    Logger.log("Payment Method validation applied to Column S.");
  }
  ui.alert("Setup Complete", "Initial data validations have been applied.", ui.ButtonSet.OK);
}

/**
 * Prepares lists of categories and subcategories for Gemini prompts.
 */
function getCategoryListsForPrompt() {
  const helperData = getHelperListsData();
  if (!helperData || helperData.error) {
      Logger.log("Could not get category lists: helperData is not loaded or contains an error.");
      return { primaryCategories: "", subcategories: "" };
  }

  const primaryCategories = (helperData.primaryCategories || []).join(", ");
  const subcategories = (helperData.allUniqueSubcategories || []).join(", ");
  
  return { primaryCategories, subcategories };
}


/**
 * Processes OCR text from the currently selected row using Gemini API.
 */
function processSelectedRowOCR() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPENSE_SUMMARY_SHEET_NAME);
  if (!sheet) {
      ui.alert("Sheet 'Expense_Summary' not found.");
      return;
  }
  const activeRow = sheet.getActiveRange().getRow();
  if (activeRow <= 1) { // Assuming row 1 is header
    ui.alert("Please select a data row (not the header).");
    return;
  }

  const ocrText = sheet.getRange(activeRow, OCR_TEXT_COL).getValue();
  if (!ocrText || ocrText.toString().trim() === "") {
    ui.alert("No text found in OCR Input column (U) for the selected row.");
    return;
  }

  const apiKey = getGeminiApiKey();
  if (!apiKey) return;

  const { primaryCategories, subcategories } = getCategoryListsForPrompt();
  if (!primaryCategories) {
      ui.alert("Could not load category lists from Helper_Lists sheet.");
      return;
  }
  
  const currentYear = new Date().getFullYear();

  const prompt = `
    Analyze the following receipt text. Extract the information and provide it strictly in JSON format.
    The current year is ${currentYear}. If the year is not mentioned in the date, assume it's ${currentYear}.
    If the receipt contains multiple items, try to determine the overall vendor and total amount.

    Receipt Text:
    ---
    ${ocrText}
    ---

    Desired JSON output format:
    {
      "date": "YYYY-MM-DD", // Extracted date
      "vendor": "Vendor Name", // Extracted vendor
      "amount": 123.45, // Extracted total amount as a number
      "suggested_category": "Suggested Primary Category from list", // Suggest from: ${primaryCategories}
      "suggested_subcategory": "Suggested Subcategory from list" // Suggest from: ${subcategories} (must be relevant to suggested_category)
    }

    If a field cannot be extracted or confidently suggested, use null for its value in the JSON.
    For example, if "7-11" is the vendor, "Convenience Store - Food/Drinks" or "Convenience Store - Non-Food" could be categories based on items.
    For "Makro", "Lotus", "BigC", "CJ", "Household Supplies" is a likely category.
    For "Grab", "Transportation" or "Food" (Dining Out - Takeout/Delivery) are likely.
    For "Indian Grocery", the category is likely "Food" and subcategory "Groceries - Indian Specific".
    Prioritize accuracy.
  `;

  const requestBody = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": {
      "responseMimeType": "application/json", // Request JSON output
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestBody),
    'muteHttpExceptions': true,
    'headers': { 'X-Goog-Api-Key': apiKey } 
  };

  ui.showModalDialog(HtmlService.createHtmlOutput("<p>Processing with Gemini AI... Please wait.</p>").setWidth(300).setHeight(100), "AI Processing");
  const response = robustUrlFetch(`${GEMINI_ENDPOINT}?key=${apiKey}`, options);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput("<p> </p>").setWidth(1).setHeight(1), " "); // Close dialog

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode === 200) {
    try {
      const jsonResponse = JSON.parse(responseText);
      // Accessing the generated text according to Gemini's current API structure
      const extractedDataString = jsonResponse.candidates[0].content.parts[0].text;
      const extractedData = JSON.parse(extractedDataString); // This should now be the JSON object

      sheet.getRange(activeRow, AI_DATE_COL).setValue(extractedData.date || null);
      sheet.getRange(activeRow, AI_VENDOR_COL).setValue(extractedData.vendor || null);
      sheet.getRange(activeRow, AI_AMOUNT_COL).setValue(extractedData.amount || null);
      sheet.getRange(activeRow, AI_CAT_COL).setValue(extractedData.suggested_category || null);
      sheet.getRange(activeRow, AI_SUBCAT_COL).setValue(extractedData.suggested_subcategory || null);

      // Optionally, try to auto-fill the main columns if AI suggestions are good
      if (extractedData.date) sheet.getRange(activeRow, 12).setValue(extractedData.date); // Date (L)
      if (extractedData.vendor) sheet.getRange(activeRow, 16).setValue(extractedData.vendor); // Vendor (P)
      if (extractedData.amount) sheet.getRange(activeRow, 18).setValue(extractedData.amount); // Amount (R)
      if (extractedData.suggested_category) sheet.getRange(activeRow, PRIMARY_CATEGORY_COL).setValue(extractedData.suggested_category); // This will trigger onEdit for subcategory
      // After primary category is set, onEdit might run. We might need a small delay or set subcategory carefully.
      if (extractedData.suggested_subcategory) {
          Utilities.sleep(500); // Small delay for onEdit to potentially finish for primary category
          sheet.getRange(activeRow, SUBCATEGORY_COL).setValue(extractedData.suggested_subcategory);
      }

      ui.alert("Success", "AI processing complete. Check columns Z-AD for suggestions, and L, P, R, N, O for auto-filled values. Please verify.", ui.ButtonSet.OK);
    } catch (e) {
      Logger.log(`Error parsing Gemini JSON: ${e.message}\nResponse: ${responseText}\nStack: ${e.stack}`);
      ui.alert("Parsing Error", `Could not parse AI response. Error: ${e.message}. Raw response: ${responseText.substring(0,300)}`, ui.ButtonSet.OK);
    }
  } else {
    Logger.log(`Gemini API Error ${responseCode}: ${responseText}`);
    ui.alert("AI Error", `Gemini API returned an error (${responseCode}): ${responseText.substring(0,300)}`, ui.ButtonSet.OK);
  }
}


/**
 * Dialog to ask for number of months for insights.
 */
function getSpendingInsightsDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
      'Spending Insights',
      'Enter number of past full months to analyze (e.g., 3 for the last 3 completed months):',
      ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    const numMonths = parseInt(result.getResponseText());
    if (isNaN(numMonths) || numMonths <= 0) {
      ui.alert('Invalid input. Please enter a positive number.');
      return;
    }
    generateSpendingInsights(numMonths);
  }
}

/**
 * Generates spending insights using Gemini AI for the last N full months.
 */
function generateSpendingInsights(numMonths) {
  const ui = SpreadsheetApp.getUi();
  // The sheet is 'House-Expense-Form', which corresponds to RESPONSE_SHEET_NAME
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    ui.alert("Sheet '" + RESPONSE_SHEET_NAME + "' not found.");
    return;
  }

  const apiKey = getGeminiApiKey();
  if (!apiKey) return;

  // Determine date range for the last N full months
  const today = new Date();
  const lastDayOfPreviousMonth = new Date(today.getFullYear(), today.getMonth(), 0);
  const firstDayOfAnalysisPeriod = new Date(lastDayOfPreviousMonth.getFullYear(), lastDayOfPreviousMonth.getMonth() - numMonths + 1, 1);

  Logger.log(`Analyzing data from ${firstDayOfAnalysisPeriod.toDateString()} to ${lastDayOfPreviousMonth.toDateString()}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
      ui.alert("No expense data found to analyze.");
      return;
  }

  // Fetch all data from the relevant columns once for efficiency
  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let expenseDataForPrompt = "Date,Cost_Center,Primary_Category,Subcategory,Vendor,Amount\n";
  let count = 0;

  for (const row of allData) {
    // Use the robust RESPONSE_COLS mapping object for clarity and maintainability
    const expenseDate = row[RESPONSE_COLS.DATE];
    const primaryCat = row[RESPONSE_COLS.PRIMARY_CAT];
    const amount = row[RESPONSE_COLS.AMOUNT];

    // Ensure essential data exists and the date is a valid Date object
    if (!expenseDate || !(expenseDate instanceof Date) || !primaryCat || !amount) {
        continue; // Skip invalid rows
    }
    
    if (expenseDate >= firstDayOfAnalysisPeriod && expenseDate <= lastDayOfPreviousMonth) {
      const costCenter = row[RESPONSE_COLS.COST_CENTER] || "N/A";
      const subCat = row[RESPONSE_COLS.SUB_CAT] || "N/A";
      const vendor = row[RESPONSE_COLS.VENDOR] || "N/A";
      
      expenseDataForPrompt += `${Utilities.formatDate(expenseDate, Session.getScriptTimeZone(), "yyyy-MM-dd")},${costCenter},${primaryCat},${subCat},"${vendor.toString().replace(/"/g, '""')}",${amount}\n`;
      count++;
    }
  }
  
  if (count === 0) {
      ui.alert("No Data", `No expense data found for the last ${numMonths} full months.`, ui.ButtonSet.OK);
      return;
  }
  
  // Truncate if too long (Gemini has input token limits)
  const MAX_PROMPT_LENGTH = 28000; 
  if (expenseDataForPrompt.length > MAX_PROMPT_LENGTH) {
      expenseDataForPrompt = expenseDataForPrompt.substring(0, MAX_PROMPT_LENGTH) + "\n...[Data Truncated]...";
      Logger.log("Data for Gemini prompt was truncated due to length.");
  }

  Logger.log(`Sending ${count} records to Gemini.`);

  const prompt = `
    You are a financial analyst reviewing home expense data.
    The data covers the last ${numMonths} full months. Cost centers are House-Flat10, House-Britannia25, and Common.
    The data format is CSV: Date,Cost_Center,Primary_Category,Subcategory,Vendor,Amount

    Expense Data:
    ---
    ${expenseDataForPrompt}
    ---

    Please provide the following insights:
    1.  Overall Spending Trend: What is the general trend of total spending over these ${numMonths} months? Increasing, decreasing, stable?
    2.  Top 3-5 Spending Categories: List the primary categories with the highest total spending during this period.
    3.  Cost Center Comparison (if possible): Briefly compare spending habits or significant categories for House-Flat10 and House-Britannia25. If data is insufficient for a detailed comparison, state so.
    4.  Vendor Insights: Are there any specific vendors like 7-11, Grab, Makro, Lotus, BigC, CJ, or Indian Groceries that show significant or frequent spending? Mention any patterns.
    5.  Potential Savings: Identify 2-3 specific subcategories or vendors where spending seems high or variable, suggesting potential areas for optimization or savings. Provide brief reasoning.
    6.  Noteworthy Observations: Any other unusual patterns, spikes in specific categories/subcategories, or interesting findings?

    Format your response clearly using markdown for readability (headings, bullet points). Be concise and actionable.
  `;
  
  // ... The rest of your function (UrlFetchApp call and response handling) remains exactly the same.
  const requestBody = { "contents": [{ "parts": [{ "text": prompt }] }] };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestBody),
    'muteHttpExceptions': true,
    'headers': { 'X-Goog-Api-Key': apiKey } 
  };

  ui.showModalDialog(HtmlService.createHtmlOutput("<p>Generating insights with Gemini AI... This may take a moment.</p>").setWidth(400).setHeight(150), "AI Insights");
  const response = robustUrlFetch(`${GEMINI_ENDPOINT}?key=${apiKey}`, options);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput("<p> </p>").setWidth(1).setHeight(1), " "); // Close dialog

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode === 200) {
    try {
      const jsonResponse = JSON.parse(responseText);
      const insights = jsonResponse.candidates[0].content.parts[0].text;
      const htmlOutput = HtmlService.createHtmlOutput(`<pre style="white-space: pre-wrap; word-wrap: break-word;">${insights}</pre>`).setWidth(700).setHeight(500);
      ui.showDialog(htmlOutput);
    } catch (e) {
      Logger.log(`Error parsing Gemini insights: ${e.message}\nResponse: ${responseText}\nStack: ${e.stack}`);
      ui.alert("Parsing Error", `Could not parse AI insights. Error: ${e.message}. Raw response: ${responseText.substring(0,300)}`, ui.ButtonSet.OK);
    }
  } else {
    Logger.log(`Gemini Insights API Error ${responseCode}: ${responseText}`);
    ui.alert("AI Error", `Gemini API returned an error for insights (${responseCode}): ${responseText.substring(0,300)}`, ui.ButtonSet.OK);
  }
}

function currentDateTime(){
  console.log(new Date());
}

// --- AI Processing ---
function processInputWithAI(userInputText, fileData) {
  return extractExpenseDetails(userInputText, fileData);
}

function extractExpenseDetails(userInputText, fileData) {

  const groqApiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
    
  // If it's a text-only expense extraction, try Groq first for speed.
  if (USE_GROQ_FOR_TEXT && groqApiKey && userInputText && !fileData) {
    Logger.log("Attempting fast extraction with Groq...");
    const groqResult = callGroqForExtraction(userInputText);
    if (groqResult && !groqResult.error) {
        Logger.log("✅ Groq extraction successful.");
        return { ...groqResult, isAnswer: false };
    }
    Logger.log(`⚠️ Groq extraction failed or returned error. Falling back to Gemini. Error: ${groqResult.error}`);
  }

  const geminiApiKey = getGeminiApiKey();
  if (!geminiApiKey) return { error: "No AI API Key found." };

   // --- GEMINI (for multi-modal capability) ---
  // Gemini here is the only one that can handle images/PDFs for this project as 28 september 2025.
  if (geminiApiKey) {
    try {
      Logger.log("Attempting extraction with Gemini...");
      const geminiResult = callGeminiForExtraction(userInputText, fileData, geminiApiKey);
      if (geminiResult && !geminiResult.error) {
        Logger.log("✅ Gemini extraction successful.");
        return { ...geminiResult, isAnswer: false };
      }
      // If Gemini returns a structured error, we'll fall through to the catch block.
      throw new Error(geminiResult.error || "Gemini returned an empty or invalid response.");
    } catch (e) {
      Logger.log(`⚠️ Gemini extraction failed. Error: ${e.message}. Attempting fallback to Groq...`);
      // The catch block will now naturally lead to the Groq fallback below.
    }
  }
  
}

function callGeminiForExtraction(userInputText, fileData, apiKey) {
  const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(); // Get user's email
  
  const helperData = getHelperListsData();
  if (helperData.error) {
    return { error: "Failed to load necessary helper data for AI: " + helperData.error };
  }

  let suggestedCostCenter = 'Others (Common)'; // Default to the fallback
  const userAliases = helperData.userAliases || {};
  const costCenters = helperData.costCenters || [];
  
  // Find the alias (e.g., "amon") that maps to the current user's email
  const userAlias = Object.keys(userAliases).find(alias => userAliases[alias] === userEmail.toLowerCase());

  if (userAlias) {
    // Find a cost center that contains the user's alias (e.g., "Personal-Amon")
    const foundCostCenter = costCenters.find(cc => cc.toLowerCase().includes(userAlias.toLowerCase()));
    if (foundCostCenter) {
      suggestedCostCenter = foundCostCenter;
    }
  }
  Logger.log(`Suggested default cost center for ${userEmail} is: ${suggestedCostCenter}`);

  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd (EEEE)");
  const parts = [];

  const extractionPrompt = `
      **ROLE:**
      You are an expert expense tracking assistant.
      
      **GOAL:**
      Your goal is to extract expense details with high accuracy. 
      Do NOT include markdown, backticks, or explanations.
      If you add anything else, it will break the system.

      **Analysis Instructions & Rules:**:
      The user's email is ${userEmail} and The current Date is ${currentDate}.
      Analyze the provided text and/or image to extract expense details. If an image is provided, also extract all visible text from it.
      1. **Prioritize:** Information from a receipt (image/PDF) is the highest authority. Use user's text for context or missing details.
      2. **Date:** Find the transaction/payment date. If no year is specified, assume the current year: ${new Date().getFullYear()}. Format as YYYY-MM-DD.
      3. **Cost Center (CRITICAL REASONING):**
          - The user's personal default cost center is **"${suggestedCostCenter}"**.
          - **RULE A:** If the expense is clearly personal (e.g., "my lunch", "clothing for myself", "gift for a friend"), you MUST use the user's personal default: **"${suggestedCostCenter}"**.
          - **RULE B:** If the expense is clearly for the household (e.g., "groceries for the kitchen", "electricity bill", "rent", "home repair", "vegetables"), you MUST choose one of the house cost centers: **"House-Flat10"** or **"House-Britannia25"**. If you cannot determine which house, default to "House-Flat10".
          - **RULE C:** Only use **"Others (Common)"** if the expense is explicitly shared between multiple people or if you absolutely cannot decide between personal and household.
          - The full list of available cost centers is: ${helperData.costCenters.join(", ")}.
      4. **Vendor:**
          - First, check if the vendor in the text is a close match to any in this list of KNOWN VENDORS: ${helperData.vendors.join(", ")}. If yes, use the name from the list.
          - If no match, accurately identify the vendor from the receipt/text (e.g., "7-ELEVEN", "Tops Market", "GrabFood").
          - If it's truly ambiguous, use "Others".
      5. **Description:** Create a concise summary.
          - If specific items are listed (e.g., "Milk, Bread, Eggs"), include them.
          - If a person's name is mentioned (e.g., "lunch with John", "payment to Jane"), include it.
          - If it's a utility bill, mention the service (e.g., "Electricity Bill - May").
      6. **Amount:** Extract the final total amount paid. It must be a number. If a currency is mentioned other than ${DEFAULT_CURRENCY}, attempt to provide the value with your best estimate in ${DEFAULT_CURRENCY} or the raw numeric value if conversion is not possible.
      7. **Payment Method:** Suggest a method from this list: ${helperData.paymentMethods.join(", ")}. Infer from text (e.g., "paid by card", "visa", "mastercard" -> "Credit Card"; "scan" -> "QR/PromptPay"; "cash" -> "Cash"). If unknown, default to "Cash".
      8. **Categorization (2-step process):**
          - Step A: First, choose the best **Primary Category** from this strict list: ${helperData.primaryCategories.join(", ")}.
          - Step B: Then, based on the chosen Primary Category, select the most relevant **Subcategory** from this strict list: ${helperData.allUniqueSubcategories.join(", ")}. The subcategory MUST logically belong to the primary category.
          - If you cannot determine a specific category, default Primary Category to "Miscellaneous" and Subcategory to "Other".
      9. **Extracted Text:** Provide ALL plain text extracted from the image/PDF. If no file was provided, echo the user's input text here.
      
      **Output Format:**
       Return a single, valid JSON object with the following keys. Do not add any explanation before or after the JSON.
      {
        "date": "YYYY-MM-DD",
        "costCenter": "string",
        "vendor": "string",
        "description": "string",
        "amount": 123.45,
        "paymentMethod": "string",
        "primaryCategory": "string",
        "subcategory": "string",
        "extractedPlainText": "string"
      }
    `;
  
  parts.push({ text: extractionPrompt });
  if (userInputText) parts.push({ text: `\n\nUser's text input: "${userInputText}"` });
  
  let fileType;
  if (fileData && fileData.base64Data && fileData.mimeType) {
    if (fileData.mimeType.startsWith("image/"))
      fileType = 'image';
    else if (fileData.mimeType === "application/pdf")
      fileType = 'pdf';

    parts.push({ text: `\n\nAnalyze this ${fileType}:` });
    parts.push({
      inlineData: { mimeType: fileData.mimeType, data: fileData.base64Data }
    });
    parts.push({ text: `\n\nExtract all plain text from the ${fileType} as well.` });
  }

const requestBody = {
  contents: [{ parts: parts }]
  // generationConfig: { responseMimeType: "application/json" }
};


  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true,
    headers: { 'X-Goog-Api-Key': apiKey } 
  };
  
  const url = `${GEMINI_ENDPOINT}?key=${apiKey}`;
  try {
    
      const response = robustUrlFetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
    
      if (responseCode === 200) {
        
        const jsonResponse = JSON.parse(responseText);

        let textResponse = jsonResponse.candidates[0].content.parts[0].text;
        
        if (jsonResponse.candidates && jsonResponse.candidates[0].content.parts[0].text) {

          // Clean up common formatting issues
          textResponse = textResponse.replace(/```json/gi, "").replace(/```/g, "").trim();
          const extractedData = JSON.parse(textResponse);
          Logger.log("Gemini Raw Parsed Data (Extraction): " + JSON.stringify(extractedData));
          return { ...extractedData, isAnswer: false }; // Add flag
        
        } else {

            const errorDetails = jsonResponse.candidates ? JSON.stringify(jsonResponse.candidates[0]) : "No candidates in response.";
            Logger.log(`Could not parse Gemini response. Details: ${errorDetails}`);
            return { error: `Could not parse Gemini response. Details: ${errorDetails}`, isAnswer: false };
        }
      } else {
        throw new Error(`Gemini API Error (${responseCode}): ${responseText.substring(0, 500)}`);
      } 
    
    } catch (e) { 
      Logger.log(`⚠️ Gemini extraction failed after all retries. Error: ${e.message}. Attempting fallback to Groq...`);
     // --- Fallback Model: Groq (only for text input) ---
      const groqApiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
      if (groqApiKey && userInputText) {
        try {
          const groqResult = callGroqForExtraction(userInputText);
          if (groqResult && !groqResult.error) {
            Logger.log("✅ Fallback to Groq successful.");
            return { ...groqResult, isAnswer: false };
          }
          // If Groq also fails, we'll fall through to the final error return.
          Logger.log(`❌ Groq fallback also failed. Error: ${groqResult.error}`);
        } catch (groqError) {
          Logger.log(`❌ Groq fallback threw an exception. Error: ${groqError.message}`);
        }
      }

      // If all primary and fallback attempts fail, return the original error to the user.
      return { error: "AI processing failed. The primary model may be overloaded, and no fallback was available. Please try again later. Details: " + e.message + "\n May be try text only input and not image or pdf", isAnswer: false };
  }
}

function callGroqForExtraction(userInputText) {
  const groqApiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  const groqModel = GROQ_MODEL_NAME;
  
  const helperData = getHelperListsData();
  const userEmail = Session.getActiveUser().getEmail();
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd (EEEE)");

  // --- NEW, SMARTER DEFAULT COST CENTER LOGIC (Mirrors Gemini's) ---
  let suggestedCostCenter = 'Others (Common)'; // Default to the fallback
  const userAliases = helperData.userAliases || {};
  const costCenters = helperData.costCenters || [];
  
  const userAlias = Object.keys(userAliases).find(alias => userAliases[alias] === userEmail.toLowerCase());

  if (userAlias) {
    const foundCostCenter = costCenters.find(cc => cc.toLowerCase().includes(userAlias.toLowerCase()));
    if (foundCostCenter) {
      suggestedCostCenter = foundCostCenter;
    }
  }
  Logger.log(`[Groq] Suggested default cost center for ${userEmail} is: ${suggestedCostCenter}`);
  // --- END NEW LOGIC ---

  const prompt = `
    You are a meticulous financial data extraction assistant. Your task is to analyze the user's text and convert it into a structured JSON object with perfect accuracy.

    **Context:**
    - User's Text: "${userInputText}"
    - Current Date: ${currentDate}
    - User's Personal Default Cost Center: "${suggestedCostCenter}"

    **Reference Data (Strict Lists):**
    - Primary Categories: ${helperData.primaryCategories.join(", ")}
    - All Subcategories: ${helperData.allUniqueSubcategories.join(", ")}
    - Payment Methods: ${helperData.paymentMethods.join(", ")}
    - Known Vendors: ${helperData.vendors.join(", ")}
    - Cost Centers: ${helperData.costCenters.join(", ")}

    **Extraction Rules:**
    1.  **Date:** Extract the date. If no year is mentioned, assume the current year. Format MUST be "YYYY-MM-DD". For "yesterday", calculate the date based on the current date.
    2.  **Amount:** Extract only the numerical value. It MUST be a number (e.g., 1110.00), not a string with currency.
    3.  **Vendor:** First, try to match a vendor from the "Known Vendors" list. If no close match, use the exact name from the user's text (e.g., "PTT").
    4.  **Payment Method:** Choose the best match from the "Payment Methods" list. Infer intelligently (e.g., "KTC credit card" -> "Credit Card").
    
    5.  **Cost Center (CRITICAL REASONING):**
        - The user's personal default is **"${suggestedCostCenter}"**.
        - **RULE A:** If the expense is clearly personal (e.g., "my lunch", "clothing for myself", "gift for a friend"), you MUST use the user's personal default: **"${suggestedCostCenter}"**.
        - **RULE B:** If the expense is clearly for the household (e.g., "groceries for the kitchen", "electricity bill", "rent", "home repair", "vegetables"), you MUST choose one of the house cost centers: **"House-Flat10"** or **"House-Britannia25"**. If you cannot determine which house, default to "House-Flat10".
        - **RULE C:** Only use **"Others (Common)"** if the expense is explicitly shared between multiple people or if you absolutely cannot decide between personal and household.

    6.  **Categorization:** First, choose the most logical **Primary Category**. Then, choose a related **Subcategory**. Both MUST exist in the reference lists.
    7.  **Description:** Create a brief, useful summary of the expense.
    8.  **extractedPlainText:** This field MUST contain the user's original, unmodified input text.

    **Output Format:**
    Respond ONLY with a single, valid JSON object. Do not include any introductory text, explanations, or markdown formatting like \`\`\`json.

    {
      "date": "YYYY-MM-DD",
      "costCenter": "string",
      "vendor": "string",
      "description": "string",
      "amount": number,
      "paymentMethod": "string",
      "primaryCategory": "string",
      "subcategory": "string",
      "extractedPlainText": "string"
    }
  `;

  try {
    const response = robustUrlFetch(GROQ_TEX_ENDPOINT, {
      method: "post",
      headers: { "Authorization": "Bearer " + groqApiKey, "Content-Type": "application/json" },
      payload: JSON.stringify({
        "messages": [{ "role": "user", "content": prompt }],
        "model": groqModel,
        "temperature": 0.1,
        "response_format": { "type": "json_object" }
      }),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      const json = JSON.parse(responseText);
      const content = json.choices[0].message.content;
      Logger.log("Groq processed successfully: " + content); // Log the actual content
      return JSON.parse(content);
    } else {
      Logger.log(`Groq API Error ${responseCode}: ${responseText}`);
      // Safely parse error message
      try {
        return { error: `Groq API Error (${responseCode}): ${JSON.parse(responseText).error.message}` };
      } catch (e) {
        return { error: `Groq API Error (${responseCode}): ${responseText}` };
      }
    }
  } catch (e) {
    Logger.log(`Error calling Groq: ${e.message}`);
    return { error: e.message };
  }
}

// --- File Handling ---
function uploadFileToDrive(fileName, mimeType, base64Data) {
  try {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(); // For folder structure or just logging
    let folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
    
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, fileName);
    
    // Sanitize filename
    const safeFileName = `${new Date().toISOString()}_${fileName.replace(/[^a-zA-Z0-9._-]/g, '_')}`;
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Or more restrictive if needed
    
    Logger.log(`File uploaded: ${file.getName()}, URL: ${file.getUrl()}, User: ${userEmail}`);
    return { fileUrl: file.getUrl(), fileName: file.getName() };
  } catch (e) {
    Logger.log(`Error uploading file to Drive: ${e.message}\nStack: ${e.stack}`);
    return { error: 'Failed to upload file: ' + e.message };
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    Logger.log(`Folder "${folderName}" not found, creating it.`);
    return DriveApp.createFolder(folderName);
  }
}

// --- Save Expense Data ---
function saveExpenseEntry(data) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSE_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${RESPONSE_SHEET_NAME}" not found.`);

    // Data object from client should match sheet columns B-M, J, K
    const newRow = createExpenseRowArray(data);
    
    sheet.appendRow(newRow);
    Logger.log("Expense entry saved: " + JSON.stringify(data));

    if (data.reminderDateTime) {
      try {
        const reminderTitle = `Expense Reminder: ${data.itemDescription || 'Review Expense'}`;
        const reminderDescription = `Reminder for expense of ${data.amount} ${DEFAULT_CURRENCY} for Vendor: ${data.vendorMerchant}.\nNotes: ${data.notes || 'N/A'}`;
        createCalendarReminderTool({'title':reminderTitle,'dateTime':new Date(data.reminderDateTime),'description':reminderDescription});
        Logger.log(`Successfully created calendar reminder for ${data.reminderDateTime}`);
      } catch (calError) {
        Logger.log(`Could not create calendar reminder. Error: ${calError.toString()}`);
        // Don't fail the whole operation, just log the error.
      }
    }
    
    // Get summary stats for the success message
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    const summaryStats = getUserExpenseSummaryStats(userEmail);
    if (summaryStats.error) {
      CacheService.getScriptCache().remove('HISTORICAL_DATA_CACHE');
      Logger.log("Invalidated historical data cache due to new expense entry.");
      // Log error but still return success for saving, message won't have stats
      Logger.log("Could not fetch summary stats for user " + userEmail + ": " + summaryStats.error);
      return { success: true, message: "Expense saved successfully! (Summary stats unavailable)", email: Session.getActiveUser().getEmail() };
    }

    CacheService.getScriptCache().remove('HISTORICAL_DATA_CACHE');
    Logger.log("Invalidated historical data cache due to new expense entry.");

    return { 
      success: true, 
      message: "Expense saved successfully!",
      summaryStats: summaryStats, // Pass stats to client
      email: Session.getActiveUser().getEmail()
    };
    
  } catch (e) {
    Logger.log(`Error saving expense: ${e.message}\nStack: ${e.stack}`);
    return { success: false, error: 'Failed to save expense: ' + e.message };
  }
}



/**
 * Creates a custom menu in the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Expense AI Tools')
      .addItem('0. Show Web App URL', 'showWebAppUrl')
      // .addItem('1. Setup Initial Validations (Run Once)', 'setupInitialValidations')
      .addSeparator()
      .addItem('2. Process OCR Text for Selected Row', 'processSelectedRowOCR')
      .addItem('3. Get Spending Insights (Gemini)', 'getSpendingInsightsDialog') // Keep this from previous version
      .addToUi();
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  if (url) {
    const htmlOutput = HtmlService.createHtmlOutput(`<p>Your Web App URL is: <a href="${url}" target="_blank">${url}</a></p><p>Ensure you have deployed your script as a Web App (Deploy > New Deployment, select Web App, execute as "Me", access "Anyone" or "Anyone with Google Account").</p>`)
      .setWidth(600)
      .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Web App URL');
  } else {
    SpreadsheetApp.getUi().alert('Web App URL not available. Make sure you have deployed the script as a web app.');
  }
}


function getUserExpenseSummaryStats(userEmail) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSE_SHEET_NAME);
    if (!sheet) throw new Error("Response sheet not found for summary.");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { N_thisMonth: 0, Amt_thisMonth: 0, N_total: 0, Amt_total: 0 };

    const dataRange = sheet.getRange("B2:L" + lastRow);
    const values = dataRange.getValues();

    let N_thisMonth = 0;
    let Amt_thisMonth = 0;
    let N_total = 0;
    let Amt_total = 0;

    const scriptTimeZone = Session.getScriptTimeZone();
    const currentMonthStr = Utilities.formatDate(new Date(), scriptTimeZone, "yyyy-MM");
    
    // --- ADDED FOR DEBUGGING ---
    // Logger.log("--- Starting Stats Calculation ---");
    // Logger.log("Target User: " + userEmail);
    // Logger.log("Current Month String for Comparison: " + currentMonthStr);
    // --- END DEBUGGING ADDITION ---

    values.forEach((row, index) => {
      const emailInSheet = row[10];
      if (emailInSheet === userEmail) {
        N_total++;
        const amount = parseFloat(row[6]);
        if (!isNaN(amount)) {
          Amt_total += amount;
        }

        const expenseDateRaw = row[0];
        let expenseDate = null;

        // --- ADDED FOR DEBUGGING ---
        // This is the most important log. It tells us what the script is actually seeing.
        // Logger.log(`Row ${index + 2}: Raw Date Value = "${expenseDateRaw}" | Type = ${typeof expenseDateRaw}`);
        // --- END DEBUGGING ADDITION ---

        if (expenseDateRaw instanceof Date) {
          expenseDate = expenseDateRaw;
        } else if (typeof expenseDateRaw === 'string' && expenseDateRaw.includes('/')) {
          const parts = expenseDateRaw.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
               const parsedDate = new Date(Date.UTC(year, month, day));
               if (!isNaN(parsedDate.getTime())) {
                 expenseDate = parsedDate;
               }
            }
          }
        }

        if (expenseDate) {
          const expenseMonthStr = Utilities.formatDate(expenseDate, scriptTimeZone, "yyyy-MM");
          
          // --- ADDED FOR DEBUGGING ---
          // This log shows us the final values being compared.
          if (expenseMonthStr === currentMonthStr) {
            Logger.log(`SUCCESS on Row ${index + 2}: Match found! Comparing "${expenseMonthStr}" === "${currentMonthStr}"`);
          }
          // --- END DEBUGGING ADDITION ---

          if (expenseMonthStr === currentMonthStr) {
            N_thisMonth++;
            if (!isNaN(amount)) {
              Amt_thisMonth += amount;
            }
          }
        } else {
            // --- ADDED FOR DEBUGGING ---
            Logger.log(`FAIL on Row ${index + 2}: Could not parse the raw date value into a valid Date object.`);
            // --- END DEBUGGING ADDITION ---
        }
      }
    });
    
    const finalResult = {
      N_thisMonth: N_thisMonth,
      Amt_thisMonth: parseFloat(Amt_thisMonth.toFixed(2)),
      N_total: N_total,
      Amt_total: parseFloat(Amt_total.toFixed(2))
    };

    // --- ADDED FOR DEBUGGING ---
    Logger.log("--- Calculation Finished ---");
    Logger.log("Final Result: " + JSON.stringify(finalResult));
    // --- END DEBUGGING ADDITION ---
    
    return finalResult;

  } catch (e) {
    Logger.log("FATAL Error in getUserExpenseSummaryStats: " + e.toString());
    return { error: "Could not retrieve summary stats." };
  }
}


function robustSanitize(html) {
  if (!html) return '';
  
  // Decode any HTML entities (like &#3456;) into their actual characters.
  let decodedText = Utilities.parseHtml(html).getText();
  
  // Whitelist common characters: English, Thai, Numbers, and common punctuation/symbols.
  // - \u0020-\u007E : Basic Latin (ASCII)
  // - \u0E00-\u0E7F : Thai
  // - \u0900-\u097F : Devanagari (Hindi)
  // - \d\s          : Digits and whitespace
  // - .,_@#&+-/()   : Common punctuation and symbols
  // - $€£¥₹฿        : Currency symbols (explicitly include ₹)
  const whitelistRegex = /[^\u0020-\u007E\u0E00-\u0E7F\u0900-\u097F\d\s.,_@#&+\-\/()$€£¥₹฿]/g;
  
  // Remove unwanted characters and normalize whitespace.
  return decodedText.replace(whitelistRegex, ' ').replace(/\s+/g, ' ').trim();
}

function sanitizeHtml(html) {
  return html
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<img[^>]*(1x1|awstrack|facebook|linkedin|tracking)[^>]*>/gi, '') // Remove tracking images
    .replace(/\son\w+="[^"]*"/gi, '') // Strip JS events (onclick etc.)
    .replace(/<\/?(?!div|p|br|h[1-6]|strong|em|b|i|table|tr|td|th|img|span|ul|ol|li)[^>]*>/gi, '') // Allow safe tags only
    .replace(/<[^>]+>/g, tag => tag.replace(/javascript:/gi, '')); // Neutralize any JS attempts
}

function sanitizeHtmlForPdf(rawHtml) {
  if (!rawHtml) return '';

  let cleanText = robustSanitize(rawHtml);
  const htmlContent = `
    <div style="font-family:sans-serif;padding:15px;">
      ${cleanText.replace(/\n/g, '<br/>')}
    </div>
  `;
  return htmlContent;
}


function extractExpenseData(body) {
  return {
    amount: (body.match(/(?:฿|USD|THB|RM|INR)?\s?\d{1,3}(,\d{3})*(\.\d{2})?/gi) || []).slice(0, 3),
    invoiceNumber: (body.match(/invoice[\s#:]*([A-Z0-9\-]+)/i) || [])[1] || null,
    taxId: (body.match(/tax\s*id[:\s]*([\d\-]+)/i) || [])[1] || null,
    bookingId: (body.match(/booking id[:\s]*([A-Z0-9]+)/i) || [])[1] || null,
    vendor: (body.match(/(?:From|Vendor|Merchant)[:\s]*([A-Za-z0-9 ,.&]+)/i) || [])[1] || null
  };
}

function getAttachmentAsBase64(messageId) {
  try {
    const message = GmailApp.getMessageById(messageId);
    if (!message) throw new Error("Message not found.");
    
    const attachments = message.getAttachments();
    if (attachments.length === 0) throw new Error("No attachments found for this message.");
    
    const file = attachments[0];
    const base64Data = Utilities.base64Encode(file.getBytes());
    
    return {
      fileName: file.getName(),
      mimeType: file.getContentType(),
      base64Data: base64Data
    };

  } catch (e) {
    Logger.log(`Error in getAttachmentAsBase64: ${e.toString()}`);
    return { error: `An error occurred: ${e.message}` };
  }
}

function getOrCreateNestedFolder(parentFolderName, childFolderName) {
  // Find or create the parent folder at the root of Drive
  let parentFolder;
  const parentFolders = DriveApp.getFoldersByName(parentFolderName);

  if (parentFolders.hasNext()) {
    parentFolder = parentFolders.next();
  } else {
    Logger.log(`Parent folder "${parentFolderName}" not found, creating it.`);
    parentFolder = DriveApp.createFolder(parentFolderName);
  }

  // Find or create the child folder inside the parent folder
  let childFolder;
  const childFolders = parentFolder.getFoldersByName(childFolderName);

  if (childFolders.hasNext()) {
    childFolder = childFolders.next();
  } else {
    Logger.log(`Child folder "${childFolderName}" not found inside "${parentFolderName}", creating it.`);
    childFolder = parentFolder.createFolder(childFolderName);
  }
  
  return childFolder;
}

function getEmailDetails(messageId, selectedAttachmentName) {
  try {
    const message = GmailApp.getMessageById(messageId);
    if (!message) throw new Error("Message not found.");
    
    // Always get the full HTML body
    const details = {
      fullBody: message.getBody() 
    };

    // If the user selected an attachment, fetch its data
    if (selectedAttachmentName) {
      const attachments = message.getAttachments();
      const realAttachment = attachments.find(att => att.getName() === selectedAttachmentName);

      if (realAttachment) {
        // Case 1: A real attachment was found and selected
        details.attachmentData = {
          fileName: realAttachment.getName(),
          mimeType: realAttachment.getContentType(),
          base64Data: Utilities.base64Encode(realAttachment.getBytes())
        };
      } else if (selectedAttachmentName.startsWith('email_screenshot_')) {
        // Case 2: The "fake" PDF was selected. We need to generate it again.
        const body = message.getPlainBody() || message.getBody();
        const htmlContent = `<div style="font-family:sans-serif;padding:15px;border:1px solid #ccc;background:#f9f9f9;">
          <h3>Email Content</h3>
          <p><strong>From:</strong> ${message.getFrom()}</p>
          <p><strong>Subject:</strong> ${message.getSubject()}</p>
          <p><strong>Date:</strong> ${Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")}</p>
          <hr>
          <div>${body.replace(/\n/g, "<br>")}</div>
        </div>`;
        
        const pdfBlob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
        pdfBlob.setName(selectedAttachmentName);

        details.attachmentData = {
          fileName: pdfBlob.getName(),
          mimeType: 'application/pdf',
          base64Data: Utilities.base64Encode(pdfBlob.getBytes())
        };
      } else {
        Logger.log(`Warning: Attachment "${selectedAttachmentName}" was not found in message ${messageId} and is not a generated PDF.`);
      }
    }
    
    return details;

  } catch (e) {
    Logger.log(`Error in getEmailDetails: ${e.toString()}`);
    return { error: `An error occurred: ${e.message}` };
  }
}

function processSingleEmailForMulti(messageId, selectedAttachmentName) {
  try {
    const helperData = getHelperListsData();
    const apiKey = getGeminiApiKey(helperData);
    if (!apiKey) return { error: "API Key not configured." };
    
    // Step 1: Get email details including the direct link
    const message = GmailApp.getMessageById(messageId);
    if (!message) return { error: "Message not found." };
    const threadId = message.getThread().getId();
    const emailLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
    
    const details = getEmailDetails(messageId, selectedAttachmentName);
    if (details.error) {
      return { error: `Could not get details: ${details.error}` };
    }

    // Step 2: Prepare data for the AI
    // We now use the ROBUST sanitize function on the full body.
    const plainTextBody = robustSanitize(details.fullBody);
    const fileData = details.attachmentData || null;

    // Step 3: Call the AI processor
    const aiResult = processInputWithAI(plainTextBody, fileData);
    
    // Step 4: Add the email link to the description for reference
    if (aiResult && !aiResult.error) {
      aiResult.description = (aiResult.description || '') + `\n\n(Source: ${emailLink})`;
      aiResult.threadId = threadId; // Pass threadId back for labeling
    }
    
    return aiResult;

  } catch (e) {
    Logger.log(`Error in processSingleEmailForMulti for ${messageId}: ${e.toString()}`);
    return { error: e.message };
  }
}

// Function to save multiple entries
function saveMultipleExpenses(expenseDataArray) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSE_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${RESPONSE_SHEET_NAME}" not found.`);
    const rowsToAppend = [];
    const threadIdsToLabel = new Set();

    expenseDataArray.forEach(data => {
      const newRow = createExpenseRowArray(data);
      rowsToAppend.push(newRow);
      if (data.threadId) {
        threadIdsToLabel.add(data.threadId);
      }
    });

    if (rowsToAppend.length > 0) {
      // Append all rows at once for performance
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      Logger.log(`Saved ${rowsToAppend.length} expense entries.`);

      // Apply labels to all processed threads
      threadIdsToLabel.forEach(threadId => {
        applyProcessedLabelToThread(threadId);
      });

      CacheService.getScriptCache().remove('HISTORICAL_DATA_CACHE');
      Logger.log("Invalidated historical data cache due to new expense entry.");
      return { success: true, count: rowsToAppend.length };
    }

    CacheService.getScriptCache().remove('HISTORICAL_DATA_CACHE');
    Logger.log("Invalidated historical data cache due to new expense entry.");
    return { success: false, error: "No valid data to save." };

  } catch(e) {
    Logger.log(`Error in saveMultipleExpenses: ${e.toString()}`);
    return { success: false, error: e.message };
  }
}


function createExpenseRowArray(data) {
  const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return [
    new Date(),                                     // A: Timestamp
    data.dateOfExpense || null,                     // B: Date of Expense
    data.costCenter || null,                        // C: Cost Center
    data.primaryCategory || null,                   // D: Primary Category
    data.subcategory || null,                       // E: Subcategory
    data.vendorMerchant || null,                    // F: Vendor/Merchant
    data.itemDescription || null,                   // G: Item/Description
    data.amount ? parseFloat(data.amount) : null,   // H: Amount
    data.paymentMethod || null,                     // I: Payment Method
    data.receiptUrl || null,                        // J: Upload Receipt URL
    data.ocrText || null,                           // K: Pasted/Extracted OCR Text
    userEmail,                                      // L: Email address
    data.notes || null,                             // M: Notes             
    data.aiSuggestedCategory || null,               // Q: AI Category
    data.aiSuggestedSubCat || null,                 // R: AI Sub-Category
    data.aiExtractedAmount ? parseFloat(data.aiExtractedAmount) : null, // S: AI Amount
    data.aiExtractedVendor || null,                 // T: AI Vendor
    data.aiExtractedDate || null                    // U: AI Date
  ];
}


/**
 * =================================================================================
 * --- AGENTIC ORCHESTRATOR SYSTEM ---
 * This section contains the complete agentic system.
 * =================================================================================
 */
/**
 * PRIMARY ENTRY POINT: Orchestrates the entire agentic workflow.
 * It now prioritizes Gemini and falls back to Groq, with an improved final error handler.
 * @param {string} question The user's natural language question.
 * @returns {Object} A structured response for the client (text, table, or chart).
 */
function answerQuestionAgentic(question = 'List all of my chocolate and bread expenses') {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'AGENT_ANSWER_' + question.toLowerCase().trim();

    const cachedResult = cache.get(cacheKey);
    if (cachedResult) {
        Logger.log("✅ CACHE HIT: Returning cached answer instantly.");
        return { status: 'complete', result: JSON.parse(cachedResult) };
    }

    // SMART PATH SELECTION - classify query complexity first
    const queryType = classifyQueryComplexity(question);
    Logger.log(`Query classification: ${queryType.type} - ${queryType.reason}`);


    // Simple queries can be handled synchronously without background task
    if (queryType.type === 'SIMPLE') {
      Logger.log("CACHE MISS: Starting background agent task for simple query.");
      try {
        const startTime = new Date();
        const simpleResult = processSimpleQueryDirectly(question, queryType);
        const processingTime = new Date() - startTime;
        
        if (simpleResult && !simpleResult.error) {
          // Only cache successful simple results
          cache.put(cacheKey, JSON.stringify(simpleResult), 3600);
          // Still log for monitoring
          logAgentTask({
              taskId: 'simple_' + Utilities.getUuid(),
              userEmail: Session.getActiveUser().getEmail(),
              startTime: startTime,
              endTime: new Date(),
              status: 'COMPLETE',
              provider: 'direct',
              turns: 1,
              question: question,
              summary: simpleResult.summary || JSON.stringify(simpleResult).substring(0, 200)
          });
          return { status: 'complete', result: simpleResult };
        }
      } catch (e) {
          Logger.log(`Simple query processing failed (${e.message}). Falling back to full agent.`);
      }
    }
    
    Logger.log("CACHE MISS: Starting background agent task for complex query.");
    const taskId = 'task_' + Utilities.getUuid();
    
    // Immediately start the orchestrator. The client will poll the taskId.
    runOrchestratorTask(question, cacheKey, taskId); 
    
    return { status: 'running', taskId: taskId, message: 'Processing in background...'  };
}

function translateToEnglishIfNeeded(text) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    return '';
  }
  try {
    // Use empty string for source language to trigger auto-detection
    const translatedText = LanguageApp.translate(text, '', 'en');
    return translatedText;
  } catch (e) {
    Logger.log(`Warning: Translation failed. Using original text. Error: ${e.message}`);
    return text; 
  }
}

/**
 * Determine if a question requires historical context data
 */

function requiresHistoricalContext(question) {
    // First, get the English version of the question
    const englishQuestion = translateToEnglishIfNeeded(question);
    
    const normalized = englishQuestion.toLowerCase();
    const contextKeywords = ['vendor', 'category', 'categor', 'spent at', 'bought at', 'who', 'where', 
                            'which', 'what kind', 'type of', 'sort of', 'kind of'];
    
    return contextKeywords.some(keyword => normalized.includes(keyword));
}

/**
 * Classifies query complexity to determine execution path
 */
function classifyQueryComplexity(question) {
  // First, get the English version of the question
  const englishQuestion = translateToEnglishIfNeeded(question);

  // Normalize the translated question
  const normalized = englishQuestion.toLowerCase().trim();
  
  // Simple query patterns (these can now reliably work on the English text)
  const simplePatterns = [
      /list (?:my|our) expenses (?:for|in) (january|february|march|april|may|june|july|august|september|october|november|december)/i,
      /show (?:my|our) expenses (?:for|in) (january|february|march|april|may|june|july|august|september|october|november|december)/i,
      /how much did i spend on (food|transportation|shopping|utilities|entertainment|housing)/i,
      /total expenses for (january|february|march|april|may|june|july|august|september|october|november|december)/i,
      /what did i spend on (?:groceries|food) (?:last|this) month/i,
      /list expenses (?:from|for) ([a-z]+)(?:\s|$)/i,
      /^show (?:me|us) (?:my|our) (?:recent|latest|last) (\d+)? expenses?$/i
  ];
  
  // Complex query patterns
  const complexPatterns = [
      /(create|generate|make).*\b(report|pdf|document|summary)\b/i,
      /\b(chart|graph|visualization|breakdown)\b/i,
      /(email|send).*expense.*(?:to|for)/i,
      /(reminder|alert|notify|notify me).*when/i,
      /compare.*spending.*between.*months/i,
      /analyze.*trends/i,
      /forecast.*spending/i
  ];
  
  // Check for simple patterns first
  for (const pattern of simplePatterns) {
      if (pattern.test(normalized)) {
          return {
              type: 'SIMPLE',
              reason: 'Matches simple data retrieval pattern',
              pattern: pattern.toString()
          };
      }
  }
  
  // Check for complex patterns
  for (const pattern of complexPatterns) {
      if (pattern.test(normalized)) {
          return {
              type: 'COMPLEX',
              reason: 'Requires multi-step processing or document generation',
              pattern: pattern.toString()
          };
      }
  }
  
  // Default to simple if under 8 words, otherwise complex
  const wordCount = normalized.split(/\s+/).filter(word => word.length > 0).length;
  return {
      type: wordCount <= 8 ? 'SIMPLE' : 'COMPLEX',
      reason: `Word count heuristic (${wordCount} words)`,
      wordCount: wordCount
  };
}


/**
 * Process simple queries directly without the full agent workflow
 */
function processSimpleQueryDirectly(question, queryType) {
    const userDetails = getUserDetails();
    const normalized = question.toLowerCase();
    
    // 1. INTELLIGENT DATE LOGIC
    // Default: Last 90 days
    let startDate = new Date();
    startDate.setDate(startDate.getDate() - 90);
    let endDate = new Date();
    
    // Check for "Forever" keywords
    const foreverKeywords = /\b(ever|all time|history|total|entire)\b/i;
    if (foreverKeywords.test(normalized)) {
        startDate = new Date("2000-01-01"); // Search from the beginning
    } else {
        // Month detection (Existing logic)
        const monthNames = ["january", "february", "march", "april", "may", "june", 
                            "july", "august", "september", "october", "november", "december"];
        const currentYear = new Date().getFullYear();
        
        for (let i = 0; i < monthNames.length; i++) {
            if (normalized.includes(monthNames[i])) {
                const monthIndex = i;
                startDate = new Date(currentYear, monthIndex, 1);
                endDate = new Date(currentYear, monthIndex + 1, 0); // Last day of month
                break;
            }
        }
    }
    
    let filterParams = {
        userEmail: userDetails.email,
        costCenter: userDetails.userCostCenter,
        startDate: Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        endDate: Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
    };
    
    const helperData = getHelperListsData(); 

    // 2. Category & Vendor detection (Existing Logic)
    const categories = helperData.primaryCategories || [];
    for (const category of categories) {
        if (normalized.includes(category.toLowerCase())) {
            filterParams.primaryCategory = category;
            break;
        }
    }

    if (!filterParams.primaryCategory) {
        const subcats = helperData.allUniqueSubcategories || [];
        for (const subcat of subcats) {
            if (normalized.includes(subcat.toLowerCase())) {
                filterParams.subcategory = subcat;
                break;
            }
        }
    }
    
    const vendors = helperData.vendors || [];
    for (const vendor of vendors) {
        if (normalized.includes(vendor.toLowerCase())) {
            filterParams.vendor = vendor;
            break;
        }
    }
    
    // 3. Amount filters
    const amountMatches = normalized.match(/(?:over|above|more than|greater than|exceeds) (\d+(?:\.\d+)?)/i);
    if (amountMatches && amountMatches[1]) {
        filterParams.minAmount = parseFloat(amountMatches[1]);
    }

    // 4. ROBUST TEXT SEARCH FALLBACK
    if (!filterParams.primaryCategory && !filterParams.subcategory && !filterParams.vendor) {
        // Expanded list of conversational filler words to ignore
        const fillerWords = [
            "show", "list", "find", "get", "search", "fetch", "check",
            "me", "my", "our", "us", "all", "i", "we", "you",
            "expenses", "purchases", "spent", "bought", "buy", "transactions", "records",
            "for", "in", "on", "at", "from",
            "have", "has", "had", "did", "do", "does", "can", "could",
            "ever", "a", "an", "the", "any"
        ];

        // Create regex to match whole words only (\b)
        const fillerRegex = new RegExp(`\\b(${fillerWords.join('|')})\\b`, 'gi');
        
        // Remove fillers AND punctuation (like ?)
        let cleanQuery = normalized
            .replace(fillerRegex, '')       // Remove words
            .replace(/[?.,!]/g, '')         // Remove punctuation
            .replace(/\s+/g, ' ')           // Collapse spaces
            .trim();
        
        if (cleanQuery.length > 0) {
            filterParams.textSearch = cleanQuery;
        }
    }
    
    // Get the data
    const expenseData = getExpenseDataForAI(filterParams);
    
    if (expenseData.error) {
        return { type: 'error', summary: 'Error retrieving expense data', content: expenseData.error };
    }
    
    if (expenseData.count === 0) {
        return {
            type: 'text',
            summary: 'No matching expenses found',
            content: `I couldn't find any expenses matching "${filterParams.textSearch || question}".`
        };
    }
    
    // Format response
    let summaryText = `Here are your expenses matching "${filterParams.textSearch || question}"`;
    if (filterParams.primaryCategory) summaryText = `Here are your ${filterParams.primaryCategory} expenses`;
    else if (filterParams.subcategory) summaryText = `Here are your ${filterParams.subcategory} expenses`;
    else if (filterParams.vendor) summaryText = `Here are your expenses at ${filterParams.vendor}`;
    
    // Add date context to summary if it was "ever"
    if (foreverKeywords.test(normalized)) {
        summaryText += " (All Time)";
    }

    return {
        type: 'table',
        summary: `${summaryText} (Total: ${DEFAULT_CURRENCY_SYMBOL}${expenseData.totalAmount})`,
        data: {
            data: expenseData.data,
            columns: ["date", "description", "primaryCategory", "vendor", "amount", "receiptUrl"]
        }
    };
}

/**
 * THE POLLING FUNCTION: Now simply reads the latest status from the cache.
 */
function getAgentTaskUpdate(taskId) {
  const cache = CacheService.getScriptCache();
  const state = cache.get(taskId);

  if (!state) {
    return { status: 'running', message: 'Initializing agent...' };
  }

  try {
    return JSON.parse(state);
  } catch (e) {
    return { status: 'error', message: 'Failed to parse task state: ' + e.message };
  }
}

function getDegradedModeResponse(fallbackAttempts, question, userEmail) {
    Logger.log('Activating degraded mode response');
    
    // Try to extract any useful partial data from fallback attempts
    let partialData = null;
    let partialSummary = "I encountered processing errors, but here's what I could retrieve:";
    
    // Find the best partial result
    for (const attempt of fallbackAttempts) {
        if (attempt.lastResult && attempt.lastResult.data && attempt.lastResult.data.length > 0) {
            partialData = attempt.lastResult.data;
            if (attempt.lastResult.summary) {
                partialSummary = attempt.lastResult.summary;
            }
            break;
        }
    }
    
    // If we have partial data, format it as a table response
    if (partialData && partialData.length > 0) {
        return {
            type: 'table',
            summary: `⚠️ Partial Results: ${partialSummary}`,
            data: {
                data: partialData,
                columns: ["date", "description", "primaryCategory", "vendor", "amount", "receiptUrl"]
            },
            degradedMode: true,
            errorContext: `Processing failed after ${fallbackAttempts.length} provider attempts.`
        };
    }
    
    // Last resort: Get recent expenses directly without AI processing
    try {
        const ninetyDaysAgo = new Date();
        ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);
        
        const fallbackData = getExpenseDataForAI({
            userEmail: userEmail,
            startDate: ninetyDaysAgo.toISOString().slice(0, 10),
            endDate: new Date().toISOString().slice(0, 10)
        });
        
        if (fallbackData.data && fallbackData.data.length > 0) {
            return {
                type: 'table',
                summary: '⚠️ System Error: Showing your recent expenses as fallback',
                data: {
                    data: fallbackData.data.slice(0, 20), // Limit to 20 items
                    columns: ["date", "description", "primaryCategory", "vendor", "amount", "receiptUrl"]
                },
                degradedMode: true,
                errorContext: 'AI processing unavailable. Showing raw expense data.'
            };
        }
    } catch (e) {
        Logger.log(`Fallback data retrieval failed: ${e.message}`);
    }
    
    // Ultimate fallback - simple error message with guidance
    return {
        type: 'text',
        summary: '❌ System Error',
        content: `I'm sorry, but I couldn't process your request due to technical issues. Please try:\n1. Simplifying your question\n2. Checking your API keys in settings\n3. Trying again in a few minutes\n\nOriginal question: "${question.substring(0, 100)}${question.length > 100 ? '...' : ''}"`,
        degradedMode: true
    };
}

/**
 * TRIGGERED FUNCTION: This runs in the background to execute the full orchestrator.
 */
function runOrchestratorTask(question, cacheKey, taskId) {
  const cache = CacheService.getScriptCache();
  const startTime = new Date();
  const userEmail = Session.getActiveUser().getEmail();

  try {
    Logger.log("🚀 Starting stateful agentic workflow...");
    updateTaskStatus(taskId, { stage: 'start', message: '🚀 Agent activated...'});
    const executionResult  = agentOrchestrator(question, taskId);
    
    saveUserSessionContext(
        userEmail, 
        question, 
        executionResult.result.summary // The text summary of the answer
    );
    
    if (executionResult.status === 'ERROR') {
        // This handles errors that happened inside the agent loop
        throw executionResult.error;
    }

    const finalState = {
        status: 'complete',
        message: '✅ Agent completed successfully',
        result: executionResult.result,
        endTime: new Date().toISOString(),
        taskId: taskId,
    };

    cache.put(taskId, JSON.stringify(finalState), 3600); // Cache final state for 1 hour
    cache.put(cacheKey, JSON.stringify(executionResult.result), 3600); // Cache final answer for reuse

    // ✅ SUCCESS LOGGING
    logAgentTask({
      taskId: taskId, userEmail: userEmail, startTime: startTime, endTime: new Date(),
      status: executionResult.status, provider: executionResult.provider,
      turns: executionResult.turns, question: question,
      summary: executionResult.result.summary || (typeof executionResult.result.content === 'string' ? executionResult.result.content.substring(0, 500) : null)
    });

  } catch (e) {

    Logger.log(`❌ runOrchestratorTask Task Failed: ${e.message}\nStack: ${e.stack}`);
    
    // Use the degraded mode system to get a response
    const degradedResponse = getDegradedModeResponse(
        e.fallbackAttempts || [{ error: e.message }],
        question,
        userEmail
    );

    // Format the final error state
    const errorState = {
        status: 'complete', // Still "complete" but with degraded content
        result: degradedResponse,
        message: '⚠️ Completed with limitations',
        taskId: taskId,
        processingTimeMs: new Date() - startTime,
        errorDetails: {
            originalError: e.message,
            fallbackAttempts: e.fallbackAttempts ? e.fallbackAttempts.length : 1
        }
    };
    
    // Cache the degraded result for a short time
    cache.put(taskId, JSON.stringify(errorState), 300); // 5 minutes
    cache.put(cacheKey, JSON.stringify(degradedResponse), 300);

     // Log the error for monitoring
    logAgentTask({
        taskId: taskId,
        userEmail: userEmail,
        startTime: startTime,
        endTime: new Date(),
        status: 'DEGRADED',
        provider: e.provider || 'unknown',
        turns: e.turns || 0,
        question: question,
        summary: degradedResponse.summary || 'Degraded mode response',
        error: e.message,
        fallbackAttempts: e.fallbackAttempts ? e.fallbackAttempts.length : 1
    });
  }
}

/**
 * HELPER: Updates the status message for an ongoing agent task in the cache.
 */
let _lastUpdateTime = 0;
function updateTaskStatus(taskId, message) {
  const now = Date.now();
  if (now - _lastUpdateTime < 1000) return; // throttle: 1 per second
  _lastUpdateTime = now;

  try {
    const cache = CacheService.getScriptCache();
    const state = {
      status: 'running',
      message: message, // message can be a string or an object
      lastUpdate: new Date().toISOString()
    };
    cache.put(taskId, JSON.stringify(state), 600); // Cache state for 10 minutes
  } catch (e) {
    Logger.log('updateTaskStatus error: ' + e.message);
  }
}


/**
 * Returns the unified set of available tools for the agent.
 * Format is OpenAI-compatible for seamless use with Groq and Gemini.
 */
function getAgentTools() {
  return [{
    "type": "function",
    "function": {
      "name": "get_expense_data",
      "description": "Fetches, filters, and summarizes expense records from the spreadsheet. Use this to answer any question about spending history.",
      "parameters": {
        "type": "object",
        "properties": {
          "personSearch": { "type": "string", "description": "A name or alias (e.g., 'Priya', 'papa') to search for across multiple fields (Cost Center, Description, Email, Notes, OCR Text). Use this for questions about a specific person's activities." },
          "userEmail": { "type": "string", "description": "The email of the person who owns the data. Use this ONLY for self-referential questions ('my expenses', 'I spent') if personSearch is not applicable." },
          "startDate": { "type": "string", "description": "The start date for the filter in YYYY-MM-DD format." },
          "endDate": { "type": "string", "description": "The end date for the filter in YYYY-MM-DD format." },
          "primaryCategory": { "type": "string", "description": "Filter by a primary category like 'Food' or 'Transportation'." },
          "subcategory": { "type": "string", "description": "Filter by a specific subcategory like 'Groceries' or 'Taxi'." },
          "vendor": { "type": "string", "description": "Filter by a specific vendor like 'Grab' or '7-11'." },
          "costCenter": { "type": "string", "description": "Filter by a cost center like 'House-Flat10'." },
          "paymentMethod": { "type": "string", "description": "Filter by a payment method like 'Credit Card' or 'Cash'." },
          "minAmount": { "type": "number", "description": "The minimum expense amount to include." },
          "maxAmount": { "type": "number", "description": "The maximum expense amount to include." },
          "textSearch": { "type": "string", "description": "A keyword to search for within the expense description, notes, or OCR text." },
          "monthYear": { "type": "string", "description": "Filter by derived Month-Year (e.g., '2024-05')." },
          "monthName": { "type": "string", "description": "Filter by full month name (e.g., 'May')." },
          "year": { "type": "string", "description": "Filter by 4-digit year (e.g., '2024')." }
        },
        "required": []
      }
    }
  }, {
    "type": "function",
    "function": {
      "name": "create_calendar_reminder",
      "description": "Sets a reminder or creates an event in the user's Google Calendar.",
      "parameters": {
        "type": "object",
        "properties": {
          "title": { "type": "string", "description": "The title of the calendar event or reminder." },
          "dateTime": { "type": "string", "description": "The date and time for the event in ISO 8601 format (e.g., '2024-08-15T10:00:00')." },
          "description": { "type": "string", "description": "A brief description for the event." }
        }, "required": ["title", "dateTime"]
      }
    }
  }, {
    "type": "function",
    "function": {
      "name": "send_email",
      "description": "Sends an email on behalf of the user. Can include data as the body or a file as an attachment.",
      "parameters": {
        "type": "object",
        "properties": {
          "recipient": { "type": "string", "description": "The email address of the recipient. If the user says 'me' or 'myself', this MUST be the logged-in user's email." },
          "subject": { "type": "string", "description": "The subject line of the email." },
          "body": { "type": "string", "description": "The HTML or plain text content of the email body. This can be a summary of data from another tool." },
          "attachmentBase64": { "type": "string", "description": "Optional. A base64 encoded string of the file to attach." },
          "attachmentMimeType": { "type": "string", "description": "Optional. The MIME type of the attachment (e.g., 'application/pdf')." },
          "attachmentFileName": { "type": "string", "description": "Optional. The desired file name for the attachment." }
        }, "required": ["recipient", "subject", "body"]
      }
    }
  },{
      "type": "function",
      "function": {
        "name": "create_pdf_report",
        "description": "Generates a professional PDF document from an array of expense data. Use this when the user asks to create a 'report', 'summary PDF', or 'document' of their expenses. This tool takes the data from `get_expense_data` as input.",
        "parameters": {
          "type": "object",
          "properties": {
            "title": { "type": "string", "description": "The title of the PDF report, e.g., 'May 2024 Expenses Report'." },
            "data": {
              "type": "array",
              "description": "An array of expense objects, typically the output from the `get_expense_data` tool.",
              "items": { "type": "object" }
            }
          },
          "required": ["title", "data"]
        }
      }
    },{
      "type": "function",
      "function": {
        "name": "search_expense_emails",
        "description": "Searches the user's Gmail inbox for unlogged, expense-related emails like receipts or invoices based on predefined keywords from the helper sheet.",
        "parameters": {
          "type": "object",
          "properties": {
            "userQuery": { "type": "string", "description": "Optional keywords from the user to add to the search, e.g., 'from Amazon'." }
          },
          "required": []
        }
      }
    },{
    "type": "function",
    "function": {
      "name": "upload_file_to_drive",
      "description": "Saves a file (like a PDF) to Google Drive and provides a shareable link. Use this after generating a file that the user wants to access.",
      "parameters": {
        "type": "object",
        "properties": {
          "fileName": { "type": "string", "description": "The desired file name, e.g., 'Expense Report.pdf'." },
          "mimeType": { "type": "string", "description": "The MIME type of the file, e.g., 'application/pdf'." },
          "base64Data": { "type": "string", "description": "The base64 encoded content of the file, usually from another tool like `create_pdf_report`." }
        },
        "required": ["fileName", "mimeType", "base64Data"]
      }
    }
  }];
}

/**
 * The "BRAIN" PROMPT: Merges your planner and summarizer logic into a single, stateful prompt.
 * This allows the agent to decide whether to call another tool OR provide a final answer.
 * @returns {string} The deliberation prompt for the AI.
 */
function getDeliberationPrompt(originalQuestion, history, context, taskId, turnNumber = 1, maxTurns = 5) {
    updateTaskStatus(taskId, { stage: 'thinking', message: 'Analyzing historical data...' });
    
    // Only load historical data when needed (not on first turn for simple queries)
    let historicalData = { categories: "Not loaded yet", subcategories: "Not loaded yet", vendors: "Not loaded yet" };
    if (turnNumber > 1 || history.length > 0 || requiresHistoricalContext(originalQuestion)) {
        historicalData = getHistoricalDataForAgent();
    }

    // We are reverting to the more detailed history string from your original prompt.
    // It provides better context for the AI's next step, which improves accuracy.
    const historyString = history.length > 0 ?
        history.map(h => {
            if (h.role === 'tool') return `🔍 OBSERVATION (Tool Result): ${h.content}`;
            if (h.role === 'assistant' && h.tool_calls) return `⚙️ ACTION: Calling tool '${h.tool_calls[0].function.name}'`;
            if (h.role === 'assistant' && h.content) return `🧠 THOUGHT: ${h.content}`;
            return ''; // Skip user messages in this specific summary view to save tokens
        }).filter(Boolean).join('\n') :
        "No actions have been taken yet.";

    // Your smart turn-counter logic is integrated into the main rules.
   const turnWarning = (turnNumber >= maxTurns) 
        ? `\n⚠️ CRITICAL WARNING: This is your FINAL TURN (${turnNumber}/${maxTurns}). You CANNOT call more tools. You MUST summarize the data you have now.`
        : '';

    const sessionContext = getUserSessionContext(context.userEmail);
    let contextString = "";
        if (sessionContext.lastQuery) {
            contextString = `
        **PREVIOUS USER QUERY:** "${sessionContext.lastQuery}"
        **PREVIOUS RESULT SUMMARY:** "${sessionContext.lastDataSummary}"
        (Use this if the user asks follow-up questions like "filter that" or "send this")
            `;
        }

    // This prompt now combines the best of both your versions.
    return `
    **SYSTEM ROLE:**
    You are the "Fin" Agent, an expert financial orchestrator. Your goal is to answer the user's question by retrieving accurate data.

    **SEMANTIC CONTEXT:**
    - **Current Date:** "${context.currentDate}"
    - **Logged-in User Email:** "${context.userEmail}"
    - **Personal Cost Center:** "${context.userCostCenter}"
    - **Aliases (name -> email):** ${context.aliasMapString}
    - **Goal:** "${originalQuestion}"
    ${contextString}

    ---
    **AVAILABLE DATA INDICES:**
    - **Primary Categories:** ${historicalData.categories}
    - **Subcategories:** ${historicalData.subcategories} 
    - **Vendors:** ${historicalData.vendors}
    ---

    **CONVERSATION HISTORY (Actions Taken So Far):**
    ---
    ${historyString}
    ---

    **DECISION PROTOCOL (Follow Strictly):**
    1. **ANALYZE:** Look at the "Observation" in the history. Does it contain the answer?
    2. **DECIDE:**
       - IF data is missing -> Call a Tool (e.g., 'get_expense_data').
       - IF data is sufficient -> Provide Final Answer.
       - IF you failed previously -> Try a different search term or broader date range.

    **TOOL CALL INSTRUCTIONS (When the goal is INCOMPLETE):**
        1. Identify the single next tool needed to advance the goal (e.g., get_expense_data, create_pdf_report).
        2. Construct the tool call with all necessary parameters. For personal queries ("me","mine", "my", "I"), you MUST use BOTH 'userEmail' and 'costCenter'.
        3. CRITICAL: Your response MUST be ONLY the raw JSON object for the tool call. Do not add any other text or markdown.
      
    **FINAL ANSWER INSTRUCTIONS (ONLY when the goal is COMPLETE):**
        1.  **ALWAYS provide a concise, one-sentence text summary** of the findings first.
        2.  **ANALYZE THE USER'S REQUEST & THE DATA:**
            - If the user asked for a "chart", "graph", or "breakdown", generate a **CHART**.
            - If the user asked to "list", "show", or if the Observation contains transactions, generate a **TABLE**.
            - **CRITICAL RULE:** Even if you generate a chart, you MUST ALSO include the full table data.
        3.  **FORMAT YOUR RESPONSE:** After the summary, provide a single, valid JSON block inside \`\`\`json ... \`\`\` containing the chart and/or table data.

    **--- REQUIRED FINAL ANSWER JSON FORMAT ---**
    You MUST respond with a JSON object inside a \`\`\`json block that looks like this:
    \`\`\`json
    {
      "type": "chart_and_table",
      "summary": "This is the text summary that appears above the data.",
      "chartData": {
        "type": "pie",
        "data": {
          "labels": ["Food", "Transport", "Shopping"],
          "datasets": [{ "label": "Spending by Category", "data": [5000, 1500, 2500] }]
        },
        "options": { "responsive": true }
      },
      "tableData": {
        "data": [ ... the FULL array of data objects from the tool's observation ... ],
        "columns": ["date", "description", "primaryCategory", "vendor", "amount", "receiptUrl"]
      }
    }
    \`\`\`
    **NOTES ON FORMAT:**
    - If no chart is needed, omit the "chartData" key and set "type" to "table".
    - **NEVER omit the "tableData" key if the observation contained a list of expenses.**
    `;
}

/**
 * The CORE ORCHESTRATOR: Manages the conversation loop for complex, multi-step tasks.
 * @param {string} question The user's original question.
 * @param {string} provider The AI provider to use ('groq' or 'gemini').
 * @returns {Object} The final, structured response for the client.
 */
function agentOrchestrator(question, taskId) {
  updateTaskStatus(taskId, { stage: 'start', message: '🚀 Starting stateful agentic workflow...' });
  let executionResult;
  let lastToolResultFromFailure = null; // <-- State variable to hold partial results

  let fallbackAttempts = [];
  const MAX_TURNS_ADAPTIVE = getAdaptiveMaxTurns(question);
  const startTime = new Date();
  let currentHistory = []; // Store history here to pass to fallback

  try {
    updateTaskStatus(taskId, { stage: 'planning', message: 'Attempting with primary provider (GEMINI)...' });
    executionResult = runAgentLoop(question, 'gemini', taskId, MAX_TURNS_ADAPTIVE, []);
    executionResult.provider = 'gemini';
      executionResult.turns = executionResult.turns || 0;
    executionResult.status = 'COMPLETE';
  } catch (e) {

      fallbackAttempts.push({
          provider: 'gemini',
          error: e.message,
          turns: e.turns || 0,
          lastResult: e.lastToolResult || null
      });
      const failedHistory = e.history || [];
      Logger.log(`⚠️ Gemini orchestrator failed after ${e.turns || 0} turns: ${e.message}. Switching to fallback with ${failedHistory.length} history items.`);
      
      // Intelligent fallback selection based on error type
      const fallbackProvider = selectFallbackProvider(e, question);
      if (fallbackProvider) {
        updateTaskStatus(taskId, { 
          stage: 'fallback', 
          message: `⚠️ Gemini failed. Storing partial result and Switching to ${fallbackProvider.toUpperCase()} for better results...`,
          details: e.message.substring(0, 100)
        });
        
        try {
          executionResult = runAgentLoop(question, fallbackProvider, taskId, MAX_TURNS_ADAPTIVE - (e.turns || 0), failedHistory);
          executionResult.provider = fallbackProvider;
          executionResult.turns = (e.turns || 0) + (executionResult.turns || 0);
          executionResult.status = 'FALLBACK_SUCCESS';
        } catch (fallbackError) {
          fallbackAttempts.push({
              provider: fallbackProvider,
              error: fallbackError.message,
              turns: fallbackError.turns || 0,
              lastResult: fallbackError.lastToolResult || null
          });
          
          Logger.log(`❌ ${fallbackProvider.toUpperCase()} fallback also failed: ${fallbackError.message}`);
          throw createEnhancedError(fallbackError, fallbackAttempts, question, startTime);
        }
      } else {
        // No suitable fallback provider
        throw createEnhancedError(e, fallbackAttempts, question, startTime);
      }
  }
  executionResult.status = 'COMPLETE';
  executionResult.processingTimeMs = new Date() - startTime;
  return executionResult;
}

/**
 * Select the most appropriate fallback provider based on error type
 */
function selectFallbackProvider(error, question) {
    const errorMessage = (error.message || '').toLowerCase();
    const questionLower = question.toLowerCase();
    
    // Pattern-based fallback selection
    if (errorMessage.includes('quota') || errorMessage.includes('rate limit') || errorMessage.includes('429')) {
        return 'groq'; // Groq typically has higher rate limits
    }
    
    if (errorMessage.includes('timeout') || errorMessage.includes('504')) {
        return 'groq'; // Groq often faster for simple tasks
    }
    
    if (errorMessage.includes('invalid') || errorMessage.includes('schema') || errorMessage.includes('format')) {
        // For format errors, try the other provider
        return error.provider === 'gemini' ? 'groq' : 'gemini';
    }
    
    // Question-based fallback selection
    if (questionLower.includes('pdf') || questionLower.includes('report') || questionLower.includes('document')) {
        return 'gemini'; // Gemini better with document generation
    }
    
    if (questionLower.includes('quick') || questionLower.includes('fast') || questionLower.includes('summary')) {
        return 'groq'; // Groq typically faster for simple summaries
    }
    
    // Default fallback strategy
    return error.provider === 'gemini' ? 'groq' : 'gemini';
}


/**
 * Create an enhanced error with fallback data and degradation strategy
 */
function createEnhancedError(originalError, fallbackAttempts, question, startTime) {
    const enhancedError = new Error(`All AI providers failed to process: "${question.substring(0, 50)}..."`);
    enhancedError.originalError = originalError;
    enhancedError.fallbackAttempts = fallbackAttempts;
    enhancedError.processingTimeMs = new Date() - startTime;
    
    // Find the best partial result from fallback attempts
    let bestPartialResult = null;
    let maxDataCount = -1;
    
    fallbackAttempts.forEach(attempt => {
        if (attempt.lastResult && attempt.lastResult.data && attempt.lastResult.data.length > maxDataCount) {
            maxDataCount = attempt.lastResult.data.length;
            bestPartialResult = attempt.lastResult;
        }
    });
    
    // Attach the best partial result if available
    if (bestPartialResult) {
        enhancedError.lastToolResult = bestPartialResult;
    }
    
    return enhancedError;
}

/**
 * Get adaptive maximum turns based on query complexity
 */
function getAdaptiveMaxTurns(question) {
    const normalized = question.toLowerCase();
    
    // Short queries typically need fewer turns
    const wordCount = normalized.split(/\s+/).filter(word => word.length > 0).length;
    if (wordCount <= 5) return 2;
    
    // Complex operations need more turns
    if (/(create|generate|make).*report|pdf|document/i.test(normalized) ||
        /(compare|analyze|forecast).*spending/i.test(normalized) ||
        /(email|send).*with.*attachment/i.test(normalized)) {
        return 6;
    }
    
    // Default adaptive turns
    return wordCount < 10 ? 3 : 4;
}

function runAgentLoop(question, provider, taskId, maxTurns, initialHistory = []) {
    const apiKey = provider === 'groq' ? PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY') : getGeminiApiKey();
    if (!apiKey) throw new Error(`API key for ${provider} is not configured.`);

    const userDetails = getUserDetails();
    const context = {
        userEmail: userDetails.email,
        userAlias: userDetails.alias,
        userCostCenter: userDetails.userCostCenter,
        aliasMapString: JSON.stringify(getHelperListsData().userAliases),
        currentDate: new Date().toISOString().slice(0, 10)
    };
    const availableTools = getAgentTools();
    let conversationHistory = [...initialHistory];
    let lastToolResult = null;
    const MAX_TURNS = 5;

    for (let i = 0; i < MAX_TURNS; i++) {
      updateTaskStatus(taskId, { stage: 'thinking', message: `Thinking (Turn ${i + 1}, using ${provider.toUpperCase()})...` });
      const deliberationPrompt = getDeliberationPrompt(question, conversationHistory, context, taskId);
      let modelResponse;

      updateTaskStatus(taskId, { stage: 'thinking', message: `Contacting ${provider.toUpperCase()}...` });
        
      if (provider === 'gemini') {
          const geminiContents = convertOaiToGeminiMessages([...conversationHistory, { role: 'user', content: deliberationPrompt }]);
          const response = callGeminiWithTools(geminiContents, availableTools, apiKey);
          if (!response || !response.candidates || !response.candidates[0]?.content?.parts?.[0]) {
              throw new Error("Gemini Planner returned an invalid response: " + JSON.stringify(response));
          }
          const geminiPart = response.candidates[0].content.parts[0];
          modelResponse = {
              role: 'assistant', content: geminiPart.text || null,
              tool_calls: geminiPart.functionCall ? [{
                  id: geminiPart.functionCall.name,
                  name: geminiPart.functionCall.name,
                  type: 'function',
                  function: { name: geminiPart.functionCall.name, arguments: JSON.stringify(geminiPart.functionCall.args || {}) }
              }] : null
          };
      } else { // Groq provider logic
          const messagesForAPI = [...conversationHistory, { role: 'user', content: deliberationPrompt }];
          const response = callGroqWithTools(messagesForAPI, availableTools, apiKey, GROQ_MODEL_NAME);
          if (response.error) {
              // Specific logic to try the larger Groq model on a size limit error
              if (response.message && (response.message.includes("413") || response.message.includes("Request too large"))) {
                  updateTaskStatus(taskId, {stage: 'fallback', message: `Groq model too small, retrying with Llama...`});
                  const llamaResponse = callGroqWithTools([...conversationHistory, { role: 'user', content: deliberationPrompt }], availableTools, apiKey, GROQ_LLAMA_MODEL_NAME);
                  if(llamaResponse.error) throw new Error(`Groq (Llama) API call failed: ${llamaResponse.message}`);
                  modelResponse = llamaResponse.choices[0].message;
              } else {
                  throw new Error(`Groq API call failed: ${response.message}`);
              }
          } else {
                modelResponse = response.choices[0].message;
          }
        }
        
        // --- Centralized Rescue Logic ---
        if (modelResponse.content && !modelResponse.tool_calls) {
          const contentStr = modelResponse.content.trim();
          if (contentStr.startsWith('{') && contentStr.endsWith('}')) {
              try {
                  const parsedContent = JSON.parse(contentStr);
                  if (parsedContent.name && parsedContent.arguments) {
                      const functionName = parsedContent.name.split(':').pop().trim();
                      modelResponse.tool_calls = [{
                          id: "rescued_call_" + i, type: "function",
                          function: { name: functionName, arguments: JSON.stringify(parsedContent.arguments) }
                      }];
                      modelResponse.content = null;
                  }
              } catch (e) { /* Ignore parsing errors */ }
          }
        }
        
        conversationHistory.push(modelResponse);

        if (modelResponse.content) {
          updateTaskStatus(taskId, { stage: 'finalizing', message: 'Formatting final answer...' }); 
          const finalAnswer = parseFinalAnswer(modelResponse.content, lastToolResult);
          return { result: finalAnswer, provider: provider, turns: i + 1 };
          // return parseFinalAnswer(modelResponse.content, lastToolResult);
        }

        if (modelResponse.tool_calls && modelResponse.tool_calls.length > 0) {
            const toolResults = [];
            for (const [toolIndex, toolCall] of modelResponse.tool_calls.entries()) {
              const functionName = toolCall.function.name;
              const functionArgs = JSON.parse(toolCall.function.arguments);
              
              updateTaskStatus(taskId, {
                  stage: 'tool_start',
                  message: `✅ Executing tool: ${functionName}`,
                  details: functionArgs
              });
              // Pass the result of the previous tool in the same turn as context
              const toolResult = executeTool(functionName, functionArgs, context, question, lastToolResult);
              lastToolResult = toolResult; // Update lastToolResult for the next tool in the sequence

              updateTaskStatus(taskId, {
                  stage: 'tool_end',
                  message: `Tool ${functionName} finished.`,
                  details: { summary: toolResult.summary || `Count: ${toolResult.count || 0}` }
              });

              const resultSummaryForHistory = {
                  summary: toolResult.summary,
                  count: toolResult.count,
                  error: toolResult.error || null,
                  data_is_available_for_next_step: (toolResult.data && toolResult.data.length > 0) || (toolResult.base64Pdf)
              };
              if (provider === 'gemini') {
                toolResults.push({
                    role: 'function', // Gemini uses 'function' role
                    parts: [{
                        functionResponse: {
                            name: functionName,
                            response: { content: resultSummaryForHistory }
                        }
                    }]
                });
              } else { // For Groq/OpenAI format
                toolResults.push({
                    role: 'tool',
                    tool_call_id: toolCall.id,
                    name: functionName,
                    content: JSON.stringify(resultSummaryForHistory)
                });
              }
              // Check if this is the first turn (i === 0) AND the first tool call in this turn (toolIndex === 0)
              if (i === 0 && toolIndex === 0) {
                // Functions that can provide a direct, final answer.
                const directAnswerFunctions = ['get_expense_data', 'search_expense_emails'];
                
                if (directAnswerFunctions.includes(functionName)) {
                  const simpleQueryWords = ['list', 'show', 'find', 'what', 'how many', 'who', 'search'];
                  const isSimpleQuery = simpleQueryWords.some(word => question.toLowerCase().includes(word));
                  const complexActionWords = ['pdf', 'report', 'reminder', 'calendar', 'send'];
                  const isComplexTask = complexActionWords.some(word => question.toLowerCase().includes(word));

                  if (isSimpleQuery && !isComplexTask) {
                      Logger.log(`✅ Optimization triggered for '${functionName}': Returning data directly.`);
                      updateTaskStatus(taskId, { stage: 'finalizing', message: 'Data found! Formatting final response...' });
                      
                      // Create a final answer structure from the tool result
                      const finalAnswer = {
                          type: 'table',
                          summary: toolResult.message || toolResult.summary,
                          data: {
                              data: toolResult.emails || toolResult.data, // Use 'emails' or 'data' key
                              columns: toolResult.emails 
                                  ? ["date", "sender", "subject", "bodySnippet", "hasAttachment"] 
                                  : ["date", "description", "primaryCategory", "vendor", "amount", "receiptUrl"]
                          }
                      };

                      return { 
                          result: finalAnswer, 
                          provider: provider, 
                          turns: i + 1 
                      };
                  }
                }
              }

            }
            // Add all tool results from this turn to the history
            conversationHistory.push(...toolResults);

        } else {
            throw new Error("Model failed to either respond with text or call a tool.");
        }
    }

    const finalError = new Error("Agent exceeded maximum number of turns.");
    finalError.lastToolResult = lastToolResult;
    finalError.history = conversationHistory;
    finalError.turns = MAX_TURNS; // Add turns to the error object
    finalError.provider = provider;
    throw finalError;
}

/**
 * Helper for making tool-based calls to the Groq API.
 * This function is specifically designed for the OpenAI-compatible endpoint.
 */
function callGroqWithTools(messages, tools, apiKey, groqModel) {

  const payload = {
    "model": groqModel,
    "messages": messages,
    "temperature": 0.2, // Lower temp for better planning
  };

  // Only include tools if they are provided, for the summarizer step we don't need them
  if (tools && tools.length > 0) {
    payload.tools = tools;
    payload.tool_choice = "auto";
  }

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'muteHttpExceptions': true
  };

  const response = robustUrlFetch(GROQ_TEX_ENDPOINT, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log(`[DEBUG] Groq Response Code:${responseCode}, Response Text: ${JSON.stringify(responseText, null, 2)}`);

  if (responseCode !== 200) {
    Logger.log(`Groq API Error: Code ${responseCode}, Response: ${responseText}`);
    // Return a structured error to be handled by the calling function
    return { error: true, message: `Groq API failed with code ${responseCode}: ${responseText}` };
  }
  
  return JSON.parse(responseText);
}

function getUserSessionContext(userEmail) {
  const cache = CacheService.getUserCache();
  const sessionKey = `SESSION_${userEmail}`;
  const cached = cache.get(sessionKey);
  return cached ? JSON.parse(cached) : { lastQuery: null, lastDataSummary: null };
}

function saveUserSessionContext(userEmail, query, resultSummary) {
  const cache = CacheService.getUserCache();
  const sessionKey = `SESSION_${userEmail}`;
  const context = {
    lastQuery: query,
    lastDataSummary: resultSummary ? resultSummary.substring(0, 200) : "No data found"
  };
  cache.put(sessionKey, JSON.stringify(context), 600); // Keep context for 10 minutes
}

// Helper for making tool-based calls
function callGeminiWithTools(contents, tools, apiKey) {
  const url = `${GEMINI_ENDPOINT}?key=${apiKey}`;
  const geminiTools = convertOaiToGeminiTools(tools);
  
  const payload = { contents, generationConfig: { temperature: 0.2 } };
  if (geminiTools) {
    payload.tools = [geminiTools];
  }

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true,
    'headers': { 'X-Goog-Api-Key': apiKey } 
  };
  const response = robustUrlFetch(url, options);
  
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  Logger.log(`[DEBUG] Gemini Response Code: ${responseCode}, Response Text: ${responseText}`);

  if (responseCode !== 200) {
    // If the response is not OK, throw a detailed error.
    throw new Error(`Gemini API Error (${responseCode}): ${responseText}`);
  }
  
  return JSON.parse(responseText);
}



/**
 * Converts an array of OpenAI-formatted tools into the structure required by the Gemini API.
 * @param {Array} oaiTools - An array of tools in OpenAI's format.
 * @returns {Object|null} A Gemini-compatible tool configuration object or null.
 */
function convertOaiToGeminiTools(oaiTools) {
    if (!oaiTools || oaiTools.length === 0) {
        return null;
    }
    const functionDeclarations = oaiTools.map(tool => tool.function);
    return { functionDeclarations };
}

/**
 * Converts OpenAI-style message history to Gemini-native format for the API call.
 */
function convertOaiToGeminiMessages(history) {
    const geminiMessages = [];
    let modelTurnParts = [];

    history.forEach(msg => {
        if (msg.role === 'user') {
            geminiMessages.push({ role: 'user', parts: [{ text: msg.content }] });
        } else if (msg.role === 'assistant') {
            if (msg.tool_calls) {
                const toolCall = msg.tool_calls[0];
                modelTurnParts.push({
                    functionCall: {
                        name: toolCall.function.name,
                        args: JSON.parse(toolCall.function.arguments || '{}')
                    }
                });
            }
            if (msg.content) {
                modelTurnParts.push({ text: msg.content });
            }
            geminiMessages.push({ role: 'model', parts: modelTurnParts });
            modelTurnParts = []; // Reset for the next model turn
        } else if (msg.role === 'tool') {
            geminiMessages.push({
                role: 'function',
                parts: [{
                    functionResponse: {
                        name: msg.name,
                        response: { content: JSON.parse(msg.content) }
                    }
                }]
            });
        }
    });
    return geminiMessages;
}

/**
 * A unified, efficient function to search for expense-related emails.
 * @param {string} userQuery - Optional keywords from the user to refine the search.
 * @param {string} detailLevel - The level of detail to return: 'ids' or 'summary'.
 * @returns {object} An object containing the search results or an error.
 */
function searchExpenseEmails(userQuery, detailLevel = 'summary') {
  try {
    const helperData = getHelperListsData();
    if (helperData.error) throw new Error(helperData.error);
    if (!helperData.email_query_keywords || helperData.email_query_keywords.length === 0) {
      return { error: `No "email_query_keywords" found in your helper sheet.` };
    }

    // --- 1. Build the Core Search Query ---
    const keywordsQueryPart = `(${helperData.email_query_keywords.join(" OR ")})`;
    const timeFramePart = `newer_than:${helperData.email_query_days || EMAIL_QUERY_DAYS}d`;
    // CRITICAL: Exclude emails that have already been processed.
    let baseQuery = `${keywordsQueryPart} ${timeFramePart} -label:${PROCESSED_LABEL_NAME} in:anywhere`; 
    if (userQuery && userQuery.trim() !== "") {
      baseQuery += ` ${userQuery.trim()}`;
    }
    Logger.log(`Executing Gmail search (detailLevel: ${detailLevel}): ${baseQuery}`);

    const threads = GmailApp.search(baseQuery, 0, 50); // Limit to 50 results for performance

    // --- 2. Return Data Based on Requested Detail Level ---
    if (detailLevel === 'ids') {
      const threadIds = threads.map(thread => thread.getId());
      return { success: true, threadIds: threadIds, count: threadIds.length };
    }

    if (detailLevel === 'summary') {
      const results = threads.map(thread => {
        const message = thread.getMessages()[0]; // Get first message for summary
        const hasAttachment = message.getAttachments().length > 0;

        return {
          threadId: thread.getId(),
          date: Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          subject: message.getSubject(),
          sender: message.getFrom(),
          bodySnippet: message.getPlainBody().substring(0, 150) + '...',
          hasAttachment: hasAttachment
        };
      });
      return { success: true, emails: results, count: results.length };
    }
    
    return { error: "Invalid detail level specified." };

  } catch (e) {
    Logger.log(`Error in searchExpenseEmails: ${e.toString()}`);
    return { error: `An error occurred: ${e.message}` };
  }
}

function searchExpenseEmailsTool(args) {
  try {
    // Call the new unified function with 'summary' detail level
    const searchResult = searchExpenseEmails(args.userQuery || "", 'summary'); 
    
    if (searchResult.error) {
      return { success: false, error: searchResult.error };
    }

    if (searchResult.count === 0) {
      return { success: true, message: "No new expense-related emails were found matching the criteria." };
    }

    // The result is already a perfect summary for the agent to process.
    return { 
      success: true, 
      message: `Found ${searchResult.count} potential expense emails.`,
      emails: searchResult.emails // 'emails' is the array of summary objects
    };

  } catch (e) {
    Logger.log(`Error in searchExpenseEmailsTool: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * A central tool executor with caching and error resilience. It runs the function chosen by the AI.
 * The 'send_email' case resolves aliases as well.
 */
function executeTool(name, args, context, originalQuestion, previousToolResult = null) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `TOOL_${name}_${Utilities.computeDigest(JSON.stringify(args))}`;
    
    // Try to get cached result first
    const cachedResult = cache.get(cacheKey);
    if (cachedResult) {
        Logger.log(`✅ CACHE HIT for tool ${name}`);
        return JSON.parse(cachedResult);
    }
    
    try {
        let result;
        let cacheExpirySeconds = 3600; // Default 1 hour
        
        switch (name) {
            case 'get_expense_data':
                result = getExpenseDataForAI(args);
                // Shorter cache for recent data
                if (args.startDate && new Date(args.startDate) > new Date(Date.now() - 7 * 24 * 60 * 60 * 1000)) {
                    cacheExpirySeconds = 300; // 5 minutes for recent data
                }
                result.originalQuestion = originalQuestion;
                break;
                
            case 'create_calendar_reminder':
                result = createCalendarReminderTool(args);
                cacheExpirySeconds = 60; // Very short cache for actions
                break;
                
            case 'send_email':
                result = executeSendEmailTool(args, context, previousToolResult);
                cacheExpirySeconds = 60; // Very short cache for actions
                break;
                
            case 'create_pdf_report':
                result = createPdfReportTool(args);
                cacheExpirySeconds = 1800; // 30 minutes for generated documents
                break;
                
            case 'search_expense_emails':
                result = searchExpenseEmailsTool(args);
                cacheExpirySeconds = 600; // 10 minutes for email searches
                break;
                
            case 'upload_file_to_drive':
                result = uploadFileToDriveTool(args);
                cacheExpirySeconds = 60; // Very short cache for actions
                break;
                
            default:
                return { error: `Unknown tool: ${name}` };
        }
        
        // Only cache successful results
        if (result && !result.error) {
            try {
                // Deep clone to avoid reference issues
                const cacheableResult = JSON.parse(JSON.stringify(result));
                cache.put(cacheKey, JSON.stringify(cacheableResult), cacheExpirySeconds);
                Logger.log(`Cached result for ${name} with expiry ${cacheExpirySeconds}s`);
            } catch (e) {
                Logger.log(`Failed to cache result for ${name}: ${e.message}`);
            }
        }
        
        return result;
    } catch (e) {
        Logger.log(`Tool ${name} execution error: ${e.message}`);
        
        // Degraded mode - try to return partial data
        if (name === 'get_expense_data' && previousToolResult && previousToolResult.data) {
            Logger.log('⚠️ Returning previous data as fallback for get_expense_data');
            return {
                success: true,
                summary: `Showing previous results due to processing error: ${e.message.substring(0, 100)}`,
                data: previousToolResult.data,
                error: e.message,
                degradedMode: true
            };
        }
        
        return { 
            error: `Tool execution failed: ${e.message}`,
            tool: name,
            arguments: args
        };
    }
}

/**
 * A helper function to parse the final text from the LLM, extracting JSON if present.
 * @param {string} finalAnswerText - The raw text response from the summarizer LLM.
 * @returns {Object} A structured response object for the client.
 */
function parseFinalAnswer(finalAnswerText, lastToolResult = null) {
  let finalResponse = { type: 'text', content: finalAnswerText };
  // CATCH unhelpful/empty responses from the AI before parsing
  if (!finalAnswerText || finalAnswerText.trim().length < 10 || finalAnswerText.toLowerCase().includes("i cannot") || finalAnswerText.toLowerCase().includes("i am unable")) {
      return { 
          type: 'error', 
          content: 'The AI agent did not provide a specific answer or data. This can happen if the request is ambiguous or if the agent is unable to find a suitable tool.' 
      };
  }

  const jsonMatch = finalAnswerText.match(/```json\n([\s\S]*?)\n```/);

  if (jsonMatch && jsonMatch[1]) {
    try {
      const jsonData = JSON.parse(jsonMatch[1]);
      const summaryText = finalAnswerText.split('```json')[0].trim();
      finalResponse = {
        type: jsonData.type || 'table',
        summary: summaryText,
        data: jsonData
      };
    } catch (e) {
      Logger.log(`Error parsing final JSON from AI: ${e.message}`);
      // Fallback to text if JSON is malformed
      finalResponse = { type: 'text', content: finalAnswerText };
    }
  }

  if ((!finalResponse.data || !finalResponse.data.tableData) && lastToolResult && lastToolResult.data && lastToolResult.data.length > 0) {
    Logger.log("✅ AI omitted table data. Injecting last tool result to ensure table is shown.");

    // The AI gave a text summary, which we want to keep.
    const summary = finalResponse.summary || finalResponse.content;

    finalResponse = {
      type: 'table',
      summary: summary,
      data: {
        // tableData: {
        // }
          data: lastToolResult.data,
          columns: ["date", "description", "primaryCategory", "vendor", "amount", "receiptUrl"]
      }
    };
  }

  // Check if the original question asked for a report and add the trigger.
  if (lastToolResult && lastToolResult.originalQuestion && (lastToolResult.originalQuestion.toLowerCase().includes('pdf') || lastToolResult.originalQuestion.toLowerCase().includes('report'))) {
      if (finalResponse.data) {
          finalResponse.data.triggerClientReport = true;
          Logger.log("✅ Added trigger for client-side report generation.");
      }
  }
  return finalResponse;
}

// The "Tool" function the AI can call
function getExpenseDataForAI(filters) {
  try {
    Logger.log("Executing unified query with filters: " + JSON.stringify(filters));
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSE_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { summary: "No expense data found.", count: 0, data: [] };

    const allData = sheet.getRange("B2:M" + lastRow).getValues();
    let filteredData = allData;
    // 0=Date(B), 1=CostCenter(C), 2=PrimaryCat(D), 3=SubCat(E), 4=Vendor(F), 5=Desc(G),
    // 6=Amount(H), 7=PaymentMethod(I), 8=ReceiptURL(J), 9=OCRText(K), 10=Email(L), 11=Notes(M)
    
    // --- ALIAS & TEXT FILTERING (No changes here) ---
    if (filters.personSearch || filters.textSearch) {
      // Update this line
      const stopWords = new Set(['and', 'or', 'the', 'a', 'in', 'for', 'of', '&', 'me', 'my', 'all', 'show', 'list']);
      const initialTerms = (filters.personSearch || filters.textSearch).toLowerCase().split(/[\s,\|]+/).filter(term => term && !stopWords.has(term));
      const allSearchTerms = new Set();
      initialTerms.forEach(term => {
        allSearchTerms.add(term);
        try {
            if (/[a-zA-Z]/.test(term)) {
              allSearchTerms.add(LanguageApp.translate(term, 'en', TARGET_LANGUAGE).toLowerCase());
            }
        } catch (e) {
          Logger.log(`Could not translate search term '${term}': ${e.message}`);
        }
      });
      const searchTermsArray = [...allSearchTerms];
      filteredData = filteredData.filter(row => {
          const combinedText = `${row[5] || ''} ${row[9] || ''} ${row[11] || ''}`.toLowerCase();
          const email = (row[10] || '').toLowerCase();
          const textMatch = searchTermsArray.some(term => combinedText.includes(term));
          let personMatch = false;
          if (filters.personSearch) {
              const aliasMap = getHelperListsData().userAliases;
              personMatch = initialTerms.some(term => aliasMap[term] && aliasMap[term] === email);
          }
          return textMatch || personMatch;
      });
    }

    if (filters.userEmail && filters.costCenter) {
      // This is a personal query from the agent. We need (email OR cost center).
      // We apply this special filter first.
      filteredData = filteredData.filter(row => 
        row[10] === filters.userEmail || (row[1] && row[1].toLowerCase() === filters.costCenter.toLowerCase())
      );
      // Now, we apply the rest of the filters to this narrowed-down set.
      // We don't need to filter by userEmail or costCenter again below.
    } else {
      // Fallback to original behavior if only one is present
      if (filters.userEmail) {
        filteredData = filteredData.filter(row => row[10] === filters.userEmail);
      }
      if (filters.costCenter) {
        filteredData = filteredData.filter(row => row[1] && row[1].toLowerCase() === filters.costCenter.toLowerCase());
      }
    }


    // --- CATEGORICAL & NUMERIC FILTERS ---
    if (filters.primaryCategory) filteredData = filteredData.filter(row => row[2] && row[2].toLowerCase() === filters.primaryCategory.toLowerCase());
    if (filters.subcategory) filteredData = filteredData.filter(row => row[3] && row[3].toLowerCase() === filters.subcategory.toLowerCase());
    if (filters.vendor) filteredData = filteredData.filter(row => row[4] && row[4].toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.costCenter) filteredData = filteredData.filter(row => row[1] && row[1].toLowerCase() === filters.costCenter.toLowerCase());
    if (filters.paymentMethod) filteredData = filteredData.filter(row => row[7] && row[7].toLowerCase() === filters.paymentMethod.toLowerCase());
    if (filters.minAmount) filteredData = filteredData.filter(row => parseFloat(row[6]) >= filters.minAmount);
    if (filters.maxAmount) filteredData = filteredData.filter(row => parseFloat(row[6]) <= filters.maxAmount);
    if (filters.monthYear) filteredData = filteredData.filter(row => row[13] === filters.monthYear); // Column N
    if (filters.monthName) filteredData = filteredData.filter(row => row[14] === filters.monthName); // Column O
    if (filters.year) filteredData = filteredData.filter(row => row[15] === filters.year); // Column P
    
     // --- Date Filters ---
    if (filters.startDate && filters.endDate) {
        const start = new Date(filters.startDate);
        const end = new Date(filters.endDate);
        end.setHours(23, 59, 59);
        filteredData = filteredData.filter(row => {
            const expenseDate = row[0] instanceof Date ? row[0] : new Date(row[0]);
            return !isNaN(expenseDate.getTime()) && expenseDate >= start && expenseDate <= end;
        });
    }

    // --- DATA MAPPING & RETURN ---
    const totalAmount = filteredData.reduce((sum, row) => sum + (parseFloat(row[6]) || 0), 0);
    const fullData = filteredData.map(row => ({
      date: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[0],
      description: row[5], amount: parseFloat(row[6]) || 0, primaryCategory: row[2],
      subcategory: row[3], vendor: row[4], receiptUrl: row[8]
    })).sort((a, b) => new Date(b.date) - new Date(a.date));

    
    // Logger.log({ summary: `Found ${filteredData.length} records totaling ${DEFAULT_CURRENCY_SYMBOL}${totalAmount.toFixed(2)}.`, count: filteredData.length,totalAmount: totalAmount.toFixed(2), data: fullData });
    
    return { summary: `Found ${filteredData.length} records totaling ${DEFAULT_CURRENCY_SYMBOL}${totalAmount.toFixed(2)}.`, count: filteredData.length,totalAmount: totalAmount.toFixed(2), data: fullData };
  } catch (e) {
    Logger.log(`Error in getExpenseDataForAI: ${e.message}`);
    return { error: `Error fetching data: ${e.message}` };
  }
}

/**
 * TOOL: Creates a professional PDF expense report from a data array.
 * It fetches receipt images and embeds them directly into the PDF.
 * @param {object} args The arguments for the tool.
 * @param {Array<object>} args.data The array of expense objects.
 * @param {string} args.title The title for the report.
 * @returns {object} An object containing the base64 encoded PDF and a success message.
 */
function createPdfReportTool(args) {
  try {
    const { data, title } = args;
    if (!data || data.length === 0) {
      return { success: false, error: "No data provided to generate the report." };
    }

    const user = Session.getActiveUser();
    let totalAmount = 0;
    
    // This loop builds the HTML for each row, including the image fetching logic.
    const rowsHtml = data.map(item => {
      totalAmount += parseFloat(item.amount) || 0;
      let imageHtml = '';

      // Fetch and embed image if a valid URL is present
      if (item.receiptUrl && (item.receiptUrl.includes('googleusercontent.com') || item.receiptUrl.match(/\.(jpeg|jpg|gif|png)$/i))) {
        try {
          // Fetch the image from the URL and convert it to a Base64 string for embedding
          const imageBlob = UrlFetchApp.fetch(item.receiptUrl).getBlob();
          const base64Image = Utilities.base64Encode(imageBlob.getBytes());
          const mimeType = imageBlob.getContentType();
          imageHtml = `<img src="data:${mimeType};base64,${base64Image}" style="max-width:80px; max-height:50px; object-fit: cover;" />`;
        } catch (e) {
          Logger.log(`Could not fetch image for PDF from URL ${item.receiptUrl}: ${e.message}`);
          imageHtml = '<span>(No Image)</span>';
        }
      }

      // Return the complete HTML for the table row
      return `
        <tr>
          <td>${escapeHtml(item.date)}</td>
          <td>${escapeHtml(item.description)}</td>
          <td>${escapeHtml(item.vendor)}</td>
          <td>${escapeHtml(item.primaryCategory)}</td>
          <td style="text-align:right;">${DEFAULT_CURRENCY_SYMBOL}${parseFloat(item.amount || 0).toFixed(2)}</td>
          <td style="text-align:center;">${imageHtml}</td>
        </tr>
      `;
    }).join('');

    // This is the full HTML structure for the PDF document
    const reportHtml = `
      <html>
        <head><style>
          body { font-family: 'Helvetica', 'Arial', sans-serif; font-size: 10pt; color: #333; }
          h1 { color: #1a73e8; font-size: 18pt; text-align: center; border-bottom: 2px solid #1a73e8; padding-bottom: 10px; }
          h2 { font-size: 12pt; color: #555; text-align: right; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; page-break-inside: auto; }
          tr { page-break-inside: avoid; page-break-after: auto; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; word-break: break-word; }
          th { background-color: #f2f2f2; font-weight: bold; }
        </style></head>
        <body>
          <h1>${escapeHtml(title)}</h1>
          <h2>Report for: ${escapeHtml(user.getEmail())}</h2>
          <h2>Total Amount: ${DEFAULT_CURRENCY_SYMBOL}${totalAmount.toFixed(2)}</h2>
          <table>
            <thead><tr>
              <th>Date</th><th>Description</th><th>Vendor</th><th>Category</th><th style="text-align:right;">Amount</th><th style="text-align:center;">Receipt</th>
            </tr></thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </body>
      </html>
    `;

    // Use the reliable DriveApp method for PDF conversion
    const tempFile = DriveApp.createFile('temp_report.html', reportHtml, MimeType.HTML);
    const pdfBlob = tempFile.getAs(MimeType.PDF);
    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
    tempFile.setTrashed(true); // Clean up the temporary HTML file

    const finalFileName = `${title.replace(/[^a-z0-9]/gi, '_')}.pdf`;

    return { 
        success: true, 
        message: `Successfully generated PDF report titled "${title}".`, 
        base64Pdf: base64Pdf, 
        fileName: finalFileName 
    };

  } catch (e) {
    Logger.log(`Error in createPdfReportTool: ${e.message}\nStack: ${e.stack}`);
    return { success: false, error: e.message };
  }
}
// Helper to escape HTML for the PDF generator
function escapeHtml(str) {
    if (!str) return '';
    return str.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}


function createCalendarReminderTool(args) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.createEvent(args.title, new Date(args.dateTime), new Date(args.dateTime), {
      description: args.description || ''
    });
    event.addPopupReminder(10);
    Logger.log('Event ID: ' + event.getId());
    return { success: true, message: `Reminder "${args.title}" set for ${new Date(args.dateTime).toLocaleString()}.` };
  } catch (e) {
    Logger.log(`Error creating calendar event: ${e.message}`);
    return { success: false, error: e.message };
  }
}

function executeSendEmailTool(args, context, previousToolResult = null) {
    // Enhanced attachment handling
    let hasValidAttachment = false;
    
    if (previousToolResult && previousToolResult.base64Pdf) {
        args.attachmentBase64 = previousToolResult.base64Pdf;
        args.attachmentMimeType = 'application/pdf';
        args.attachmentFileName = previousToolResult.fileName || 'Expense_Report.pdf';
        hasValidAttachment = true;
        Logger.log("Injected PDF from previous tool call into send_email.");
    }
    
    // Resolve recipient aliases
    let finalRecipient = args.recipient;
    if (finalRecipient) {
        finalRecipient = finalRecipient.toLowerCase();
        if (['me', 'myself', 'mine'].includes(finalRecipient)) {
            finalRecipient = context.userEmail;
        } else if (!finalRecipient.includes('@')) {
            const aliasMap = getHelperListsData().userAliases || {};
            const resolvedEmail = aliasMap[finalRecipient];
            if (resolvedEmail) {
                finalRecipient = resolvedEmail;
                Logger.log(`Resolved alias "${args.recipient}" to email "${finalRecipient}"`);
            } else {
                return { 
                    error: `Could not resolve recipient "${args.recipient}". Please use a valid email or known alias.`,
                    suggestion: `Known aliases: ${Object.keys(aliasMap).join(', ')}`
                };
            }
        }
    }
    
    args.recipient = finalRecipient;
    
    // Add safety check for missing required fields
    if (!args.subject || !args.subject.trim()) {
        args.subject = `Expense Report - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}`;
    }
    
    if (!args.body || !args.body.trim()) {
        args.body = "Please find the attached expense report.";
        if (previousToolResult && previousToolResult.summary) {
            args.body = previousToolResult.summary + "<br><br>Please find the detailed report attached.";
        }
    }
    
    // Execute the send
    const result = sendEmailTool(args, context);
    
    // Add context-aware success message
    if (result.success) {
        result.message = `✅ Email sent to ${finalRecipient}` + (hasValidAttachment ? ' with report attachment' : '');
    }
    
    return result;
}

function sendEmailTool(args, context) {
  try {
    let finalRecipient = args.recipient;
    if (args.recipient && (args.recipient.toLowerCase() === 'me' || args.recipient.toLowerCase() === 'myself')) {
      finalRecipient = context.userEmail;
    }

    let finalHtmlBody = args.body;

    // Check if the agent passed table data along with the body text.
    if (args.tableData && args.tableData.length > 0) {
        const totalAmount = args.tableData.reduce((sum, item) => sum + (parseFloat(item.amount) || 0), 0);
        const tableRows = args.tableData.map(item => `
            <tr>
                <td style="border: 1px solid #ddd; padding: 8px;">${escapeHtml(item.date)}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">${escapeHtml(item.description)}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">${escapeHtml(item.vendor)}</td>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${DEFAULT_CURRENCY_SYMBOL}${parseFloat(item.amount || 0).toFixed(2)}</td>
            </tr>
        `).join('');

        finalHtmlBody += `
            <br><hr><br>
            <h3>Supporting Data (${args.tableData.length} items, Total: ${DEFAULT_CURRENCY_SYMBOL}${totalAmount.toFixed(2)})</h3>
            <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 12px;">
                <thead>
                    <tr>
                        <th style="background-color: #f2f2f2; border: 1px solid #ddd; padding: 8px; text-align: left;">Date</th>
                        <th style="background-color: #f2f2f2; border: 1px solid #ddd; padding: 8px; text-align: left;">Description</th>
                        <th style="background-color: #f2f2f2; border: 1px solid #ddd; padding: 8px; text-align: left;">Vendor</th>
                        <th style="background-color: #f2f2f2; border: 1px solid #ddd; padding: 8px; text-align: right;">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${tableRows}
                </tbody>
            </table>
        `;
    }

    const options = {
      to: finalRecipient,
      subject: args.subject,
      htmlBody: finalHtmlBody
    };

    if (args.attachmentBase64 && args.attachmentMimeType && args.attachmentFileName) {
      const decodedData = Utilities.base64Decode(args.attachmentBase64);
      const blob = Utilities.newBlob(decodedData, args.attachmentMimeType, args.attachmentFileName);
      options.attachments = [blob];
    }

    MailApp.sendEmail(options);
    return { success: true, message: `Email sent successfully to ${finalRecipient}.` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function uploadFileToDriveTool(args) {
  try {
    const { base64Data, mimeType, fileName } = args;
    // let folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
    const folder = getOrCreateNestedFolder(DRIVE_FOLDER_NAME, DRIVE_FOLDER_GENERATED_PDF);
    
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, fileName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileUrl = file.getUrl();
    Logger.log(`File uploaded via tool: ${file.getName()}, URL: ${fileUrl}`);
    return { success: true, fileUrl: fileUrl };
  } catch (e) {
    Logger.log(`Error in uploadFileToDriveTool: ${e.message}`);
    return { success: false, error: 'Failed to upload file: ' + e.message };
  }
}

function getOrCreateGmailLabel() {
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_NAME);
  if (!label) {
    Logger.log(`Label "${PROCESSED_LABEL_NAME}" not found. Creating it.`);
    label = GmailApp.createLabel(PROCESSED_LABEL_NAME);
  }
  return label;
}

function applyProcessedLabelToThread(threadId) {
  try {
    if (!threadId) return { success: false, error: "Thread ID is required." };
    const thread = GmailApp.getThreadById(threadId);
    const label = getOrCreateGmailLabel();
    thread.addLabel(label);
    Logger.log(`Applied label "${PROCESSED_LABEL_NAME}" to thread ID: ${threadId}`);
    return { success: true };
  } catch (e) {
    Logger.log(`Error applying label to thread ${threadId}: ${e.toString()}`);
    return { error: e.message };
  }
}

// --- Get Historical Data for context ---
function getHistoricalDataForAgent() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'HISTORICAL_DATA_CACHE';
  
  const cachedData = cache.get(CACHE_KEY);
  if (cachedData) {
    Logger.log("CACHE HIT: Loaded historical data for agent from cache.");
    return JSON.parse(cachedData);
  }

  Logger.log("CACHE MISS: Fetching historical data for agent from spreadsheet.");
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSE_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { vendors: "N/A", categories: "N/A", subcategories: "N/A" };

    // Fetch columns D (Primary Cat), E (Subcat), and F (Vendor)
    const data = sheet.getRange("D2:F" + lastRow).getValues();
    
    const uniqueVendors = new Set();
    const uniqueCategories = new Set();
    const uniqueSubcategories = new Set();

    data.forEach(row => {
      if (row[0]) uniqueCategories.add(row[0]); // Primary Category is in the 1st column (index 0)
      if (row[1]) uniqueSubcategories.add(row[1]); // Subcategory is in the 2nd column (index 1)
      if (row[2]) uniqueVendors.add(row[2]); // Vendor is in the 3rd column (index 2)
    });

    const finalData = {
      vendors: [...uniqueVendors].slice(0, 75).join(", "), // use .slice(0, 50).join(", ") to Limit to 50 to keep prompt size reasonable
      categories: [...uniqueCategories].join(", "),
      subcategories: [...uniqueSubcategories].slice(0, 100).join(", ")
    };

    cache.put(CACHE_KEY, JSON.stringify(finalData), 3600); // Cache for 1 hour
    return finalData;

  } catch (e) {
    Logger.log(`Error in getHistoricalDataForAgent: ${e.message}`);
    return { vendors: "", categories: "", subcategories: "" };
  }
}

function getUserDetails() {
  const cache = CacheService.getUserCache();
  const CACHE_KEY = 'USER_DETAILS';
  
  let cachedDetails = cache.get(CACHE_KEY);
  if (cachedDetails) {
    return JSON.parse(cachedDetails);
  }

  const userEmail = Session.getActiveUser().getEmail();
  const helperData = getHelperListsData(); // This is already cached
  const userAliases = helperData.userAliases || {};
  const costCenters = helperData.costCenters || [];
  let suggestedCostCenter = 'Others (Common)'; // Default

  const userAlias = Object.keys(userAliases).find(alias => userAliases[alias] === userEmail.toLowerCase());

  if (userAlias) {
    const foundCostCenter = costCenters.find(cc => cc.toLowerCase().includes(userAlias.toLowerCase()));
    if (foundCostCenter) {
      suggestedCostCenter = foundCostCenter;
    }
  }

  const userDetails = {
    email: userEmail,
    alias: userAlias || userEmail.split('@')[0],
    userCostCenter: suggestedCostCenter // NEW: Make this available globally
  };
  
  cache.put(CACHE_KEY, JSON.stringify(userDetails), 21600); // Cache for 6 hours
  return userDetails;
}

// --- Get Data for Web App Dropdowns and agentic context ---
function getHelperListsData() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'HELPER_DATA_CACHE';
  
  const cachedData = cache.get(CACHE_KEY);
  if (cachedData) {
    Logger.log("CACHE HIT: Loaded helper data from cache.");
    return JSON.parse(cachedData);
  }

  Logger.log("CACHE MISS: Fetching helper data from spreadsheet.");
  
  try {
    const helperSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HELPER_SHEET_NAME);
    if (!helperSheet) throw new Error(`Sheet "${HELPER_SHEET_NAME}" not found.`);

    const dataRange = helperSheet.getRange("A2:J" + helperSheet.getLastRow()).getValues();

    const userAliases = {};
    const costCenters = new Set();
    const paymentMethods = new Set();
    const vendors = new Set();
    const email_query_keywords = new Set();
    const categories = {};
    const primaryCategoriesList = new Set();
    const allSubcategoriesList = new Set();
    const email_query_days = dataRange[0][7];
    const userApiKeys = {};
    
    dataRange.forEach(row => {
      if (row[0]) costCenters.add(row[0]);
      if (row[1]) paymentMethods.add(row[1]);
      if (row[2]) vendors.add(row[2]);
      if (row[6]) email_query_keywords.add(row[6]);

      // --- NEW, ROBUST LOGIC FOR CATEGORIES ---
      const primaryCategory = row[3];
      const subcategoriesString = row[4];

      if (primaryCategory) {
        primaryCategoriesList.add(primaryCategory);
        if (!categories[primaryCategory]) {
          categories[primaryCategory] = [];
        }
        if (subcategoriesString) {
          // This regex splits by comma, but ignores commas inside parentheses.
          const subCats = subcategoriesString.split(/,(?![^\(]*\))/g).map(s => s.trim()).filter(Boolean);
          
          subCats.forEach(subCat => {
            if (!categories[primaryCategory].includes(subCat)) {
              categories[primaryCategory].push(subCat);
            }
            allSubcategoriesList.add(subCat);
          });
        }
      }
      // --- END NEW LOGIC ---

      const apiKeyEntry = row[8];
      if (apiKeyEntry && apiKeyEntry.includes('=')) {
        const [email, key] = apiKeyEntry.split('=').map(s => s.trim());
        if (email && key) { userApiKeys[email.toLowerCase()] = key; }
      }
      const aliasEntry = row[9];
      if (aliasEntry && aliasEntry.includes('->')) {
        const [aliasesPart, emailPart] = aliasEntry.split('->').map(s => s.trim());
        if (aliasesPart && emailPart) {
          const email = emailPart.toLowerCase();
          const aliases = aliasesPart.split(',').map(a => a.trim().toLowerCase());
          aliases.forEach(alias => {
            userAliases[alias] = email;
          });
        }
      }
    });
    
    const finalData = {
      costCenters: [...costCenters],
      paymentMethods: [...paymentMethods],
      vendors: [...vendors],
      email_query_keywords: [...email_query_keywords],
      email_query_days: email_query_days,
      primaryCategories: [...primaryCategoriesList],
      subcategoriesMap: categories,
      allUniqueSubcategories: [...allSubcategoriesList].sort(),
      userApiKeys: userApiKeys,
      userAliases: userAliases
    };

    cache.put(CACHE_KEY, JSON.stringify(finalData), 3600);
    
    return finalData;

  } catch (error) {
    Logger.log(`Error in getHelperListsData: ${error.message} \nStack: ${error.stack}`);
    return { error: error.message };
  }
}


/**
 * Gets the Gemini API Key from Script Properties.
 * REMEMBER TO SET THIS UP: File > Project properties > Script properties
 */
function getGeminiApiKey() {
  const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const helperData = getHelperListsData();

  // 1. Check for user-specific key from the sheet
  if (helperData && helperData.userApiKeys && helperData.userApiKeys[userEmail.toLowerCase()]) {
    Logger.log(`Using API key for user: ${userEmail}`);
    return helperData.userApiKeys[userEmail.toLowerCase()];
  }

  // 2. Fallback to global Script Property key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    // This alert should ideally be handled client-side for a better UX,
    // but a server-side log and error is critical.
    Logger.log("API Key Error: No user-specific key found and global GEMINI_API_KEY is not set in Script Properties.");
    return null;
  }
  Logger.log("Using global API key from Script Properties.");
  return apiKey;
}



function robustUrlFetch(url, options, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      // If the response is a rate limit error (429)
      if (responseCode === 429) {
          Logger.log(`Attempt ${i + 1}: Rate limit hit (429). Retrying...`);
          // Exponential backoff: 2s, 4s, 8s... plus random jitter up to 1s.
          const waitTime = Math.pow(2, i + 1) * 1000 + Math.random() * 1000;
          Utilities.sleep(waitTime);
          continue; // Go to the next iteration of the loop to retry
      }

      // If successful or a non-retriable client error, return immediately.
      if (responseCode < 500) {
        return response;
      }
      
      Logger.log(`Attempt ${i + 1}: Received server error ${responseCode}. Retrying...`);

    } catch (e) {
      Logger.log(`Attempt ${i + 1}: Caught exception during fetch. Retrying... Error: ${e.message}`);
    }

    // Wait before the next general retry (for 5xx errors or network issues)
    if (i < maxRetries - 1) {
      const waitTime = Math.pow(2, i) * 1000 + Math.random() * 1000;
      Utilities.sleep(waitTime);
    }
  }

  // If all retries fail, throw an error.
  throw new Error(`Failed to fetch URL after ${maxRetries} attempts.`);
}


/**
 * Logs the result of an agentic task to a dedicated sheet for monitoring.
 * @param {object} logData - An object containing all the data to be logged.
 */
function logAgentTask(logData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(ORCHESTRATOR_LOG_SHEET_NAME);

    // If the sheet doesn't exist, create it and add headers
    if (!sheet) {
      sheet = ss.insertSheet(ORCHESTRATOR_LOG_SHEET_NAME);
      const headers = [
        "Timestamp", "Task ID", "User Email", "Start Time", "End Time", 
        "Duration (ms)", "Status", "Final Provider", "Turns", 
        "Question", "Final Summary", "Error Message"
      ];
      sheet.appendRow(headers);
      sheet.getRange("A1:L1").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const duration = logData.endTime.getTime() - logData.startTime.getTime();

    const newRow = [
      new Date(),
      logData.taskId || null,
      logData.userEmail || null,
      logData.startTime,
      logData.endTime,
      duration,
      logData.status || 'UNKNOWN',
      logData.provider || 'N/A',
      logData.turns || 0,
      logData.question || null,
      logData.summary || '',
      logData.error || ''
    ];

    sheet.appendRow(newRow);

  } catch (e) {
    Logger.log(`FATAL: Could not write to orchestrator log sheet. Error: ${e.message}`);
  }
}


/**
 * The AI Financial Expert Friend.
 * This function uses a specific persona to provide creative financial advice.
 */
function getFinancialAdvice() {
  const apiKey = getGeminiApiKey();
  if (!apiKey) return { type: 'text', content: "AI Financial Friend is not configured." };

  const userEmail = Session.getActiveUser().getEmail();
  const ninetyDaysAgo = new Date();
  ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);

  const expenseDataResponse = getExpenseDataForAI({
    userEmail: userEmail,
    startDate: ninetyDaysAgo.toISOString().slice(0, 10),
    endDate: new Date().toISOString().slice(0, 10)
  });

  if (expenseDataResponse.error || expenseDataResponse.count === 0) {
    return { type: 'text', summary: "Not enough data", content: "I'd love to help, but I couldn't find enough recent expense data to analyze. Please log some more expenses first!" };
  }

  // Summarize data to keep the prompt concise
  const dataSummary = expenseDataResponse.data // .slice(0, 100) Limit to 100 recent items  
    .map(e => `${e.date}, ${e.primaryCategory}, ${e.vendor}, ${e.amount.toFixed(2)}`)
    .join('\n');
  // const categories = (getHelperListsData().primaryCategories || []).join(", ");

   const expertPrompt = `
    **Role:** You are "Fin," an exceptionally creative and insightful world-class behavioral economist and frugal-living expert. Your personality is friendly, encouraging, and known for your "out-of-the-box" thinking. Your goal is to help the user find practical and innovative ways to save money.
    
    Your goal is to help the user find practical and innovative ways to save money based on their spending habits.

    **Rules:**
    1. **Diagnose First**: Identify top 3 spending categories from the data.
    2. **Dual-Strategy per Category**:
      - **Practical**: A proven, low-effort method (e.g., "Use cash envelopes for groceries").
      - **Creative**: A novel, fun, or community-based hack (e.g., "Start a 'No-Spend Weekend Challenge' with friends where you only use what's in your pantry").
    3. **Quantify Impact**: Estimate potential monthly savings for each idea (e.g., "This could save you ~฿1,200/month").
    4. **Personalize**: Reference actual vendors/categories from the user's data (e.g., "Since you spend often at 7-11...").
    5. **Challenge**: End with a 7-day "Fin's Experiment" tied to their biggest leak.
    **Tone**: Warm, empowering, slightly playful—never judgmental.

    **Analysis Context:**
    - The user's total spending in the last 90 days is: ${DEFAULT_CURRENCY_SYMBOL}${expenseDataResponse.totalAmount}
    - Here is a sample of their recent spending (CSV: Date, Category, Vendor, Amount):
    ---
    ${dataSummary}
    ---

    **Your Task:** Generate the advice following all rules and the specified structure perfectly.
    1.  **Greeting & Analysis:** Start with a friendly greeting. Briefly analyze the user's spending, identifying the top 2-3 categories where they spend the most.
    2.  **Brainstorm Creative Savings:** For each top category, provide exactly TWO saving strategies in a list: **Practical Idea** and **Creative Idea**.
    3.  **"Fin's Weekly Experiment":** Suggest a fun, actionable one-week challenge.
    4.  **Closing Motivation:** End with a short, motivational sign-off.
    5.  **Formatting:** Use markdown for clear headings, bold text for emphasis, and bullet points.
  `;


  const response = callGeminiWithTools(
    convertOaiToGeminiMessages([{ "role": "user", "content": expertPrompt }]),
    [], 
    apiKey
  );

  const advice = response.candidates[0].content.parts[0].text;
  return { summary: "Here's some personalized advice from your financial friend, Fin:", type: 'text', content: advice };
}

// --- Audio & Advanced Voice ---
function processAudioInput(base64Audio, clientMimeType) {
  
  // Check base64 size (~75% of original)
  const byteLength = Utilities.base64Decode(base64Audio).length;
  const fileSizeMB = (byteLength / (1024 * 1024)).toFixed(2);
  if (fileSizeMB > MAX_FILE_SIZE_MB) {
    return { error: `Audio file too large (${fileSizeMB} MB). Max allowed is ${MAX_FILE_SIZE_MB} MB.` };
  }

  // // Decode base64 string
  // const blob = Utilities.newBlob(Utilities.base64Decode(base64Audio), clientMimeType, 'audio_input');

  // // Upload to temporary Drive file
  // let folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  // const tempFile = folder.createFile(blob);
  // const fileId = tempFile.getId();
  // const fileUrl = `https://drive.google.com/uc?export=download&id= ${fileId}`;

  let transcript = null;

  // Strip data URL prefix if present
  if (base64Audio.startsWith('data:')) {
    base64Audio = base64Audio.split(',')[1];
  }


  // Try Groq Whisper API
  const groqApiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  if (groqApiKey) {
    try {
      transcript = callGroqWhisper(base64Audio, clientMimeType, groqApiKey);
      if (transcript) {
        Logger.log("✅ Transcribed with Groq Whisper: " + transcript);
        // DriveApp.getFileById(fileId).setTrashed(true); // Clean up
        return processInputWithAI(transcript, null);
      }
    } catch (e) {
      Logger.log(`❌ Groq Whisper failed: ${e.message}\nTranscript: ${transcript}\nStack: ${e.stack}`);
    }
  }

  // Try AssemblyAI
  const assemblyAiKey = PropertiesService.getScriptProperties().getProperty('ASSEMBLYAI_API_KEY');
  if (assemblyAiKey) {
    try {
      // const audioUrl = getPublicAudioUrl(base64Audio, clientMimeType);
      transcript = callAssemblyAi(base64Audio, assemblyAiKey);
      if (transcript) {
        Logger.log("✅ Transcribed with AssemblyAI: " + transcript);
        // DriveApp.getFileById(fileId).setTrashed(true); // Clean up
        return processInputWithAI(transcript, null);
      }
    } catch (e) {
      Logger.log(`❌ AssemblyAI failed: ${e.message}\nTranscript: ${transcript}\nStack: ${e.stack}`);
    }
  }

  // Final fallback: Gemini Vision API (supports audio interpretation)
  const geminiApiKey = getGeminiApiKey();
  if (geminiApiKey) {
    try {
      transcript = callGeminiModelforAudio(base64Audio, geminiApiKey);
      if (transcript) {
        Logger.log("✅ Transcribed with Gemini Vision: " + transcript);
        // DriveApp.getFileById(fileId).setTrashed(true); // Clean up
        return processInputWithAI(transcript, null);
      }
    } catch (e) {
      Logger.log(`❌ Gemini Vision failed: ${e.message}\nTranscript: ${transcript}\nStack: ${e.stack}`);
    }
  }

  // All methods failed
  // DriveApp.getFileById(fileId).setTrashed(true);
  return { error: "All transcription services failed." };
}

// Helper Functions for Audio transcription
// 1. Call Groq Whisper API
function callGroqWhisper(base64Audio, mimeType, apiKey) {
  const boundary = '----WebKitFormBoundary' + Math.random().toString(36).substring(2);
  const delimiter = `--${boundary}\r\n`;
  const closeDelimiter = `--${boundary}--\r\n`;

  const audioBytes = Utilities.base64Decode(base64Audio);

  const bodyParts = [];

  // Add model field
  bodyParts.push(Utilities.newBlob(
    delimiter +
    'Content-Disposition: form-data; name="model"\r\n\r\n' +
    'whisper-large-v3\r\n',
    'text/plain'
  ).getBytes());

  // Add audio file field
  bodyParts.push(Utilities.newBlob(
    delimiter +
    'Content-Disposition: form-data; name="file"; filename="audio.webm"\r\n' +
    `Content-Type: ${mimeType}\r\n\r\n`,
    'text/plain'
  ).getBytes());

  // Add actual audio content
  bodyParts.push(audioBytes);

  // Close boundary
  bodyParts.push(Utilities.newBlob('\r\n' + closeDelimiter, 'text/plain').getBytes());

  // Flatten and create final payload
  const payload = [].concat.apply([], bodyParts);

  const options = {
    method: 'post',
    contentType: `multipart/form-data; boundary=${boundary}`,
    headers: {
      'Authorization': `Bearer ${apiKey}`
    },
    payload: payload,
    muteHttpExceptions: true
  };

  const response = robustUrlFetch(GROQ_AUDIO_ENDPOINT, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    throw new Error(`Groq Whisper failed with ${code}: ${text}`);
  }

  const json = JSON.parse(text);
  return json.text || null;
}


function callAssemblyAi(base64Audio, apiKey) {
  const audioBytes = Utilities.base64Decode(base64Audio);
  
  // Upload raw bytes
  const uploadResponse = robustUrlFetch(ASSEMBLY_AUDIO_UPLOAD_ENDPOINT, {
    method: 'post',
    contentType: 'application/octet-stream',
    headers: { 'authorization': apiKey },
    payload: audioBytes,
    muteHttpExceptions: true
  });

  const uploadCode = uploadResponse.getResponseCode();
  const uploadText = uploadResponse.getContentText();

  if (uploadCode !== 200) {
    throw new Error(`AssemblyAI upload failed: ${uploadCode} - ${uploadText}`);
  }

  const uploadJson = JSON.parse(uploadText);
  const uploadUrl = uploadJson.upload_url;

  // Start transcription
  const transcribeResponse = robustUrlFetch(ASSEMBLY_AUDIO_TRANSCRIPT_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'authorization': apiKey },
    payload: JSON.stringify({ audio_url: uploadUrl }),
    muteHttpExceptions: true
  });

  const transcribeId = JSON.parse(transcribeResponse.getContentText()).id;

  // Poll until completion
  let result;
  for (let i = 0; i < 30; i++) {
    const statusRes = robustUrlFetch(`${ASSEMBLY_AUDIO_TRANSCRIPT_ENDPOINT}/${transcribeId}`, {
      method: 'get',
      headers: { 'authorization': apiKey },
      muteHttpExceptions: true
    });
    result = JSON.parse(statusRes.getContentText());
    if (result.status === 'completed') break;
    if (result.status === 'failed') throw new Error(`Transcription failed: ${JSON.stringify(result)}`);
    Utilities.sleep(2000);
  }

  return result.text.trim();
}

function callGeminiModelforAudio(base64Audio, apiKey) {
  const url = `${GEMINI_ENDPOINT}?key=${apiKey}`;//?key=${apiKey}

  const payload = {
    contents: [{
      parts: [
        { text: "Transcribe this audio." },
        { inline_data: { mime_type: "audio/webm", data: base64Audio } }
      ]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { 'X-Goog-Api-Key': apiKey } 
  };

  try {
    const response = robustUrlFetch(url, options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code !== 200) {
      throw new Error(`Gemini failed with ${code}: ${text}`);
    }

    const json = JSON.parse(text);
    return json.candidates?.[0]?.content?.parts?.[0]?.text || null;
  } catch (e) {
    Logger.log("Gemini Vision Error: " + e.toString());
    return null;
  }
}

// Utility: Download URL and encode as base64
function base64EncodeUrl(url) {
  const response = UrlFetchApp.fetch(url);
  return Utilities.base64Encode(response.getContent());
}

function getPublicAudioUrl(base64Audio, mimeType) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Audio), mimeType, 'audio_input.wav');
  let folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  const file = folder.createFile(blob).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  return `https://drive.google.com/uc?export=download&id= ${fileId}`;
}

function getPublicAudioStream(base64Audio, mimeType) {
  const bytes = Utilities.base64Decode(base64Audio);
  return bytes; // returns raw binary data as byte array
}

// --- Fetch Expense Data for Viewing ---
function getExpenseDataForView(selectionCriteria) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSE_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${RESPONSE_SHEET_NAME}" not found.`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { // Assuming row 1 is header
      return { success: true, data: [], message: "No expenses found." };
    }

    // Fetch relevant columns for display: A to M
    // A=Timestamp, B=Date, C=CostCenter, D=PrimaryCat, E=SubCat, F=Vendor, 
    // G=Description, H=Amount, I=PaymentMethod, J=ReceiptLink, K=OCRText, L=Email, M=Notes
    const range = sheet.getRange("A2:M" + lastRow); 
    const allValues = range.getValues();

    let filteredValues = [];
    let filterStartDate = null;
    let filterEndDate = null;
    const today = new Date();
    today.setHours(0,0,0,0); // Normalize today to start of day

    if (selectionCriteria && selectionCriteria.type) {
      if (selectionCriteria.type === 'dateRange') {
        filterStartDate = new Date(selectionCriteria.startDate);
        filterEndDate = new Date(selectionCriteria.endDate);
        // Adjust endDate to include the whole day
        filterEndDate.setHours(23,59,59,999);
      } else if (selectionCriteria.type === 'period') {
        const period = selectionCriteria.value;
        const currentYear = today.getFullYear();
        const currentMonth = today.getMonth(); // 0-indexed

        switch (period) {
          case 'thisMonth':
            filterStartDate = new Date(currentYear, currentMonth, 1);
            filterEndDate = new Date(currentYear, currentMonth + 1, 0, 23,59,59,999); // Last day of current month
            break;
          case 'previousMonth':
            filterStartDate = new Date(currentYear, currentMonth - 1, 1);
            filterEndDate = new Date(currentYear, currentMonth, 0, 23,59,59,999); // Last day of previous month
            break;
          case 'past3Months': // Current month + 2 preceding full months
            filterStartDate = new Date(currentYear, currentMonth - 2, 1); // Start of 2 months ago
            filterEndDate = new Date(currentYear, currentMonth + 1, 0, 23,59,59,999); // End of current month
            break;
          case 'past6Months': // Current month + 5 preceding full months
            filterStartDate = new Date(currentYear, currentMonth - 5, 1);
            filterEndDate = new Date(currentYear, currentMonth + 1, 0, 23,59,59,999);
            break;
          case 'thisYear': // Year to Date
            filterStartDate = new Date(currentYear, 0, 1); // Jan 1st of current year
            filterEndDate = new Date(currentYear, currentMonth, today.getDate(), 23,59,59,999); // Up to today
            break;
          case 'lastYear':
            filterStartDate = new Date(currentYear - 1, 0, 1); // Jan 1st of last year
            filterEndDate = new Date(currentYear - 1, 11, 31, 23,59,59,999); // Dec 31st of last year
            break;
          case 'ALL':
            // No date filtering needed, allValues will be used
            break;
          default:
            Logger.log("Unknown period: " + period + ". Fetching all.");
            // Default to ALL if period is unrecognized
        }
      }
    } else {
      // Default behavior if no criteria passed (e.g., fetch ALL or This Month)
      // For safety, let's default to ALL if no criteria, or you can choose another default.
      Logger.log("No selectionCriteria provided, fetching all data.");
    }

    if (selectionCriteria && selectionCriteria.value !== 'ALL' && filterStartDate && filterEndDate) {
      Logger.log(`Filtering from ${filterStartDate.toDateString()} to ${filterEndDate.toDateString()}`);
      filteredValues = allValues.filter(row => {
        const expenseDate = row[1]; // Column B - Date of Expense
        // Ensure expenseDate is a valid Date object for comparison
        if (expenseDate instanceof Date) {
          const normalizedExpenseDate = new Date(expenseDate);
          normalizedExpenseDate.setHours(0,0,0,0); // Normalize for consistent comparison
          return normalizedExpenseDate >= filterStartDate && normalizedExpenseDate <= filterEndDate;
        }
        // Try to parse if it's a string (though server-side it should be Date object from sheet)
        if (typeof expenseDate === 'string') {
            try {
                const parsedDate = new Date(expenseDate);
                parsedDate.setHours(0,0,0,0);
                return parsedDate >= filterStartDate && parsedDate <= filterEndDate;
            } catch (e) { return false; } // Invalid date string
        }
        return false; // Skip if not a date
      });
    } else {
      filteredValues = allValues; // Use all values if 'ALL' or no valid filter dates
    }

    const expenses = filteredValues.map(row => {
      let formattedDate = row[1]; 
      if (formattedDate instanceof Date) {
        formattedDate = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return {
        timestamp: row[0] ? (row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : row[0]) : null,
        date: formattedDate,
        costCenter: row[2] || '',
        primaryCategory: row[3] || '',
        subcategory: row[4] || '',
        vendor: row[5] || '',
        description: row[6] || '',
        amount: row[7] ? parseFloat(row[7]) : 0, // Ensure amount is number
        paymentMethod: row[8] || '',
        receiptLink: row[9] || '',
        email: row[11] || '',
      };
    });
    
    Logger.log(`Workspaceed ${expenses.length} expenses for viewing.`);
    return { success: true, data: expenses.reverse(), message: `Found ${expenses.length} expenses.`}; // Show newest first
  } catch (error) {
    Logger.log(`Error in getExpenseDataForView: ${error.message} \nStack: ${error.stack}`);
    return { success: false, error: error.message, data: [] };
  }
}
