// Three Points Hospitality - Training Accountability System
// COMPLETE FIXED VERSION - No More Duplicates
// 
// FIXES APPLIED (January 16, 2026):
// ✅ Fixed "This Month" formulas - now uses DATE(YEAR(TODAY()),MONTH(TODAY()),1) instead of EOMONTH(TODAY(),-1)+1
// ✅ Fixed "Last Month" formulas - now correctly calculates first and last day of previous month
// ✅ Fixed "Active Trainers" count - now uses HARDCODED values from TRAINER_COUNTS constant
//    - White Buffalo: Shows 2 trainers (Holden, Dani)
//    - Original American Kitchen: Shows 8 trainers (Emma, Desiree, Natalia, Rachel, Kai, Val, Carson, Tanner)
//    - Cantina Añejo: Shows 11 trainers (Shaylee, Macy, Lilah, Ella, Adeniza, Suzy, Gabriella, Christian, Evan, Axaielle, Valeria)
//    - "This Month" and "Last Month" show how many of those trainers actually conducted assessments
// ✅ All Location Summary metrics now show accurate data for current and previous months

// Configuration
const CONFIG = {
  EMAIL_NOTIFICATIONS: {
    ENABLED: true,
    MANAGERS: ['devin@threepointshospitality.com'], // Update with your actual emails
    SUBJECT_PREFIX: '[Three Points] Training Assessment:'
  }
};

// Trainer counts by location (hardcoded from current roster)
const TRAINER_COUNTS = {
  'Cantina Añejo': 11,           // Shaylee, Macy, Lilah, Ella, Adeniza, Suzy, Gabriella, Christian, Evan, Axaielle, Valeria
  'Original American Kitchen': 8, // Emma, Desiree, Natalia, Rachel, Kai, Val, Carson, Tanner
  'White Buffalo': 2              // Holden, Dani
};

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🎯 Training System')
    .addItem('➕ Add New Trainer', 'addTrainerToHTML')
    .addSeparator()
    .addItem('👥 View Current Trainers', 'viewCurrentTrainers')
    .addItem('🔄 Refresh Analytics', 'refreshAnalytics')
    .addItem('📊 Populate Trainer Performance', 'populateTrainerPerformanceFromExistingData')
    .addItem('📍 Populate Location Performance', 'populateLocationPerformanceFromExistingData')
    .addItem('🎯 Fix Analytics Dashboard', 'populateAnalyticsDashboardFromExistingData')
    .addItem('🔧 Recreate Location Summary (FIXED)', 'recreateLocationSummarySheet')
    .addItem('🎨 Apply Location Colors', 'applyLocationColorCoding')
    .addItem('⚙️ System Setup', 'initializeSystem')
    .addToUi();
}

// Location aliases handled in HTML frontend

/**
 * Add new trainer to HTML file
 */
function addTrainerToHTML() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Get current trainer data from HTML
    const currentData = getTrainerDataFromHTML();
    const locations = Object.keys(currentData).sort();
    
    if (locations.length === 0) {
      ui.alert('Error', 'No locations found in HTML file.', ui.ButtonSet.OK);
      return;
    }
    
    // Get trainer name
    const trainerResponse = ui.prompt(
      'Add New Trainer',
      'Enter trainer name (First Last):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (trainerResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const trainerName = trainerResponse.getResponseText().trim();
    if (!trainerName) {
      ui.alert('Error', 'Trainer name cannot be empty.', ui.ButtonSet.OK);
      return;
    }
    
    // Get location with alias support
    const locationOptions = locations.join(', ');
    const aliasInfo = '\n\nShortcuts: CA = Cantina Añejo, OAK = Original American Kitchen, WB = White Buffalo';
    
    const locationResponse = ui.prompt(
      'Select Location',
      `Available locations: ${locationOptions}${aliasInfo}\n\nEnter location for ${trainerName}:`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (locationResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const locationInput = locationResponse.getResponseText().trim();
    if (!locationInput) {
      ui.alert('Error', 'Location cannot be empty.', ui.ButtonSet.OK);
      return;
    }
    
    // Resolve location alias
    const resolvedLocation = resolveLocationAlias(locationInput, locations);
    
    if (!resolvedLocation) {
      ui.alert('Error', `Location "${locationInput}" not found.\n\nAvailable: ${locationOptions}${aliasInfo}`, ui.ButtonSet.OK);
      return;
    }
    
    // Check if trainer already exists
    if (currentData[resolvedLocation].includes(trainerName)) {
      ui.alert('Error', `Trainer "${trainerName}" already exists at "${resolvedLocation}".`, ui.ButtonSet.OK);
      return;
    }
    
    // Add trainer and update HTML
    currentData[resolvedLocation].push(trainerName);
    currentData[resolvedLocation].sort(); // Keep alphabetical
    
    const success = updateTrainerDataInHTML(currentData);
    
    if (success) {
      ui.alert('Success', `✓ Added "${trainerName}" to "${resolvedLocation}"`, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Failed to update HTML file.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('Error', `Failed to add trainer: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Add new location to HTML file
 */
// addLocationToHTML removed - locations are fixed (Cantina Añejo, Original American Kitchen, White Buffalo)

/**
 * View current trainers from HTML file
 */
function viewCurrentTrainers() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const currentData = getTrainerDataFromHTML();
    
    let message = 'Current Trainers by Location:\n\n';
    
    Object.keys(currentData).sort().forEach(location => {
      message += `📍 ${location}:\n`;
      if (currentData[location].length === 0) {
        message += '   No trainers assigned\n';
      } else {
        currentData[location].forEach(trainer => {
          message += `   • ${trainer}\n`;
        });
      }
      message += '\n';
    });
    
    ui.alert('Current Trainers', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Failed to load trainers: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Extract trainer data from HTML file
 */
function getTrainerDataFromHTML() {
  try {
    // Get the HTML file from the current project
    const files = DriveApp.getFiles();
    let htmlFile = null;
    
    // Try to find index.html in the current project
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName() === 'index.html') {
        htmlFile = file;
        break;
      }
    }
    
    if (!htmlFile) {
      throw new Error('index.html file not found');
    }
    
    const htmlContent = htmlFile.getBlob().getDataAsString();
    
    // Find the trainersByLocation object
    const startPattern = /const\s+trainersByLocation\s*=\s*{/;
    const startMatch = htmlContent.match(startPattern);
    
    if (!startMatch) {
      throw new Error('trainersByLocation object not found in HTML file');
    }
    
    const startIndex = startMatch.index + startMatch[0].length - 1;
    
    // Find the matching closing brace
    let braceCount = 0;
    let endIndex = startIndex;
    
    for (let i = startIndex; i < htmlContent.length; i++) {
      if (htmlContent[i] === '{') braceCount++;
      if (htmlContent[i] === '}') braceCount--;
      if (braceCount === 0) {
        endIndex = i;
        break;
      }
    }
    
    // Extract the object string
    const objectString = htmlContent.substring(startIndex, endIndex + 1);
    
    // Parse the object (this is a simplified parser - assumes clean format)
    const cleanedString = objectString
      .replace(/'/g, '"')  // Replace single quotes with double quotes
      .replace(/,\s*}/g, '}')  // Remove trailing commas
      .replace(/,\s*]/g, ']'); // Remove trailing commas in arrays
    
    return JSON.parse(cleanedString);
    
  } catch (error) {
    console.log('Error parsing HTML:', error.toString());
    
    // Fallback: return current known structure
    return {
      'Original American Kitchen': [
        'Emma Thomas', 'Desiree Edwards', 'Natalia Martinez',
        'Rachel Donly', 'Kai Nishikawa', 'Val Revilla', 'Carson Fontana', 'Tanner Griffin'
      ],
      'White Buffalo': ['Holden Fernandez', 'Dani Mizrachi'],
      'Cantina Añejo': [
        'Shaylee Estes', 'Macy Williams', 'Lilah Bowers', 'Ella Agustin',
        'Adeniza Fenne', 'Suzy Takla', 'Gabriella McMillan', 'Christian Lucas',
        'Evan Amato', 'Axaielle Cazeau-Quinn', 'Valeria Cvjetkovic', , 'Emma Yang', 'Davia Geders', 'Selina Ayup'
      ]
    };
  }
}

/**
 * Update trainer data in HTML file
 */
function updateTrainerDataInHTML(newData) {
  try {
    // Get the HTML file from the current project
    const files = DriveApp.getFiles();
    let htmlFile = null;
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName() === 'index.html') {
        htmlFile = file;
        break;
      }
    }
    
    if (!htmlFile) {
      throw new Error('index.html file not found');
    }
    
    const htmlContent = htmlFile.getBlob().getDataAsString();
    
    // Generate new trainersByLocation object string
    const newObjectString = generateTrainersByLocationString(newData);
    
    // Find and replace the existing object
    const pattern = /(const\s+trainersByLocation\s*=\s*){[\s\S]*?};/;
    const newHtmlContent = htmlContent.replace(pattern, `$1${newObjectString};`);
    
    // Save the updated HTML file
    htmlFile.setContent(newHtmlContent);
    
    console.log('✓ HTML file updated successfully');
    return true;
    
  } catch (error) {
    console.log('Error updating HTML:', error.toString());
    return false;
  }
}

/**
 * Generate formatted trainersByLocation object string
 */
function generateTrainersByLocationString(data) {
  let result = '{\n';
  
  const locations = Object.keys(data).sort();
  
  locations.forEach((location, index) => {
    result += `        '${location}': [\n`;
    
    const trainers = data[location].sort();
    trainers.forEach((trainer, trainerIndex) => {
      const comma = trainerIndex < trainers.length - 1 ? ',' : '';
      result += `            '${trainer}'${comma}\n`;
    });
    
    const comma = index < locations.length - 1 ? ',' : '';
    result += `        ]${comma}\n`;
  });
  
  result += '    }';
  return result;
}

/**
 * Serve the HTML web app
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Training Accountability - Three Points Hospitality')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Include external files (CSS/JS) in HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Enhanced submission handler with duplicate prevention
 */
function submitTrainingData(data) {
  console.log('=== FORM SUBMISSION ===');
  console.log('Data received:', JSON.stringify(data));
  
  try {
    // Generate submission hash for duplicate detection
    const submissionHash = generateSubmissionHash(data);
    
    // Check for recent duplicate submissions (within last 5 minutes)
    if (isDuplicateSubmission(submissionHash)) {
      console.log('❌ Duplicate submission detected, blocking...');
      return {
        success: false,
        error: 'Duplicate submission detected. Please wait before submitting again.',
        timestamp: new Date().toISOString()
      };
    }
    
    // Validate required fields
    const requiredFields = ['location', 'trainer', 'trainee', 'position', 'trainingDay', 
                            'knowledgeScore', 'technicalScore', 'serviceScore', 
                            'teamworkScore', 'professionalismScore', 'totalScore', 'percentage', 'performanceLevel'];
    
    for (const field of requiredFields) {
      if (!data[field] && data[field] !== 0) {
        throw new Error(`Missing required field: ${field}`);
      }
    }
    
    // Get the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Initialize all sheets if they don't exist
    initializeAllSheets(spreadsheet);
    
    // Insert main record with transaction-like behavior
    const recordId = insertMainRecordSafe(spreadsheet, data, submissionHash);
    
    // Only proceed with notifications if main record was successful
    if (recordId) {
      // Analytics are now handled by the populate functions via menu
      
      // Send email notification (with rate limiting)
      if (CONFIG.EMAIL_NOTIFICATIONS.ENABLED) {
        sendNotificationEmailSafe(data, recordId);
      }
    }
    
    console.log('✓ Record saved successfully:', recordId);
    
    return {
      success: true,
      recordId: recordId,
      message: `Training assessment submitted successfully with ID: ${recordId}`,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.log('❌ Error:', error.toString());
    return {
      success: false,
      error: error.toString(),
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Generate unique hash for submission to detect duplicates
 */
function generateSubmissionHash(data) {
  const hashString = `${data.location}-${data.trainer}-${data.trainee}-${data.position}-${data.trainingDay}`;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, hashString)
    .map(byte => (byte + 256).toString(16).slice(-2))
    .join('');
}

/**
 * Check if this is a duplicate submission within last 5 minutes
 */
function isDuplicateSubmission(submissionHash) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Training Records');
    if (!sheet) return false;
    
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return false; // Only headers
    
    const values = dataRange.getValues();
    const fiveMinutesAgo = new Date(Date.now() - 5 * 60 * 1000);
    
    // Check last 20 rows for recent duplicates (performance optimization)
    const startRow = Math.max(1, values.length - 20);
    
    for (let i = startRow; i < values.length; i++) {
      const rowTimestamp = new Date(values[i][1]); // Column B - Timestamp
      
      if (rowTimestamp > fiveMinutesAgo) {
        // Generate hash for this row
        const rowData = {
          location: values[i][2],
          trainer: values[i][3],
          trainee: values[i][4],
          position: values[i][5],
          trainingDay: values[i][6]
        };
        
        const rowHash = generateSubmissionHash(rowData);
        
        if (rowHash === submissionHash) {
          console.log('Duplicate found in row:', i + 1);
          return true;
        }
      }
    }
    
    return false;
    
  } catch (error) {
    console.log('Error checking duplicates:', error.toString());
    return false; // Allow submission if we can't check
  }
}

/**
 * Safe record insertion with additional checks
 */
function insertMainRecordSafe(spreadsheet, data, submissionHash) {
  try {
    const sheet = spreadsheet.getSheetByName('Training Records');
    if (!sheet) {
      throw new Error('Training Records sheet not found');
    }
    
    // Lock to prevent concurrent modifications
    const lock = LockService.getScriptLock();
    
    try {
      lock.waitLock(10000); // Wait up to 10 seconds
      
      // Double-check for duplicates after acquiring lock
      if (isDuplicateSubmission(submissionHash)) {
        throw new Error('Duplicate submission detected after lock');
      }
      
      // Generate unique record ID
      const timestamp = new Date();
      const recordId = 'TR-' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + 
                       Math.random().toString(36).substr(2, 6).toUpperCase();
      
      // Prepare row data
      const rowData = [
        recordId,                     // A: Record ID
        timestamp,                    // B: Timestamp
        data.location,                // C: Location
        data.trainer,                 // D: Trainer
        data.trainee,                 // E: Trainee
        data.position,                // F: Position
        data.trainingDay,             // G: Training Day
        data.shift || '',             // H: Shift
        data.performanceLevel,        // I: Performance Level
        data.overallNotes || '',      // J: Overall Notes
        data.knowledgeScore,          // K: Knowledge Score
        data.technicalScore,          // L: Technical Score
        data.serviceScore,            // M: Service Score
        data.teamworkScore,           // N: Teamwork Score
        data.professionalismScore,    // O: Professionalism Score
        data.totalScore,              // P: Total Score
        data.percentage               // Q: Percentage
      ];
      
      console.log('Inserting row data:', rowData);
      
      // Insert the row
      sheet.appendRow(rowData);
      
      // Apply conditional formatting
      const lastRow = sheet.getLastRow();
      const performanceLevelCell = sheet.getRange(lastRow, 9);
      
      if (data.performanceLevel === 'Excellent') {
        performanceLevelCell.setBackground('#d5f4e6');
      } else if (data.performanceLevel === 'Good') {
        performanceLevelCell.setBackground('#fff3cd');
      } else {
        performanceLevelCell.setBackground('#f8d7da');
      }

      console.log('✓ Record inserted successfully:', recordId);

      // Force Google Sheets to recalculate all dependent sheets
      SpreadsheetApp.flush();

      return recordId;
      
    } finally {
      lock.releaseLock();
    }
    
  } catch (error) {
    console.log('❌ Error inserting record:', error.toString());
    throw error;
  }
}

/**
 * Rate-limited email sending
 */
function sendNotificationEmailSafe(data, recordId) {
  try {
    // Check if we've sent too many emails recently (max 10 per hour)
    const now = new Date();
    const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
    
    // Simple rate limiting using PropertiesService
    const emailCount = PropertiesService.getScriptProperties().getProperty('emailCount_' + 
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHH')) || '0';
    
    if (parseInt(emailCount) >= 10) {
      console.log('Email rate limit reached, skipping notification');
      return;
    }
    
    // Send the email
    sendNotificationEmail(data, recordId);
    
    // Update rate limit counter
    PropertiesService.getScriptProperties().setProperty('emailCount_' + 
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHH'), 
      (parseInt(emailCount) + 1).toString());
    
  } catch (error) {
    console.log('Error in safe email sending:', error.toString());
  }
}

/**
 * Initialize all required sheets
 */
function initializeAllSheets(spreadsheet) {
  // Training Records Sheet
  let mainSheet = spreadsheet.getSheetByName('Training Records');
  if (!mainSheet) {
    mainSheet = createTrainingRecordsSheet(spreadsheet);
  }
  
  // Analytics Dashboard Sheet
  let analyticsSheet = spreadsheet.getSheetByName('Analytics Dashboard');
  if (!analyticsSheet) {
    analyticsSheet = createAnalyticsSheet(spreadsheet);
  }
  
  // Location Summary Sheet (for detailed view like screenshot 4)
  let locationSummarySheet = spreadsheet.getSheetByName('Location Summary');
  if (!locationSummarySheet) {
    locationSummarySheet = createLocationSummarySheet(spreadsheet);
  }
  
  // Trainer Performance Sheet (Monthly)
  let trainersSheet = spreadsheet.getSheetByName('Trainer Performance');
  if (!trainersSheet) {
    trainersSheet = createTrainerPerformanceSheet(spreadsheet);
  }
}

/**
 * Create the main Training Records sheet
 */
function createTrainingRecordsSheet(spreadsheet) {
  console.log('Creating Training Records sheet...');
  
  const sheet = spreadsheet.insertSheet('Training Records');
  
  // Updated headers to match your actual data structure
  const headers = [
    'Record ID',           // A
    'Timestamp',           // B
    'Location',            // C
    'Trainer',             // D
    'Trainee',             // E
    'Position',            // F
    'Training Day',        // G
    'Shift',               // H
    'Performance Level',   // I
    'Overall Notes',       // J
    'Knowledge Score',     // K
    'Technical Score',     // L
    'Service Score',       // M
    'Teamwork Score',      // N
    'Professionalism Score', // O
    'Total Score',         // P
    'Percentage'           // Q
  ];
  
  // Set headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Format headers
  headerRange.setBackground('#2C5AA0')
           .setFontColor('#FFFFFF')
           .setFontWeight('bold')
           .setFontSize(11);
  
  // Set column widths
  sheet.setColumnWidth(1, 120);  // Record ID
  sheet.setColumnWidth(2, 150);  // Timestamp
  sheet.setColumnWidth(3, 150);  // Location
  sheet.setColumnWidth(4, 120);  // Trainer
  sheet.setColumnWidth(5, 120);  // Trainee
  sheet.setColumnWidth(6, 100);  // Position
  sheet.setColumnWidth(7, 80);   // Training Day
  sheet.setColumnWidth(8, 80);   // Shift
  sheet.setColumnWidth(9, 120);  // Performance Level
  sheet.setColumnWidth(10, 300); // Overall Notes
  sheet.setColumnWidth(11, 80);  // Knowledge Score
  sheet.setColumnWidth(12, 80);  // Technical Score
  sheet.setColumnWidth(13, 80);  // Service Score
  sheet.setColumnWidth(14, 80);  // Teamwork Score
  sheet.setColumnWidth(15, 80);  // Professionalism Score
  sheet.setColumnWidth(16, 80);  // Total Score
  sheet.setColumnWidth(17, 80);  // Percentage
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  console.log('✓ Training Records sheet created');
  return sheet;
}

/**
 * Create Enhanced Analytics Dashboard sheet
 */
function createAnalyticsSheet(spreadsheet) {
  console.log('Creating Enhanced Analytics Dashboard...');
  
  const sheet = spreadsheet.insertSheet('Analytics Dashboard');
  
  // Title
  sheet.getRange('A1').setValue('Three Points Hospitality - Training Analytics Dashboard')
    .setFontSize(18).setFontWeight('bold').setBackground('#2C5AA0').setFontColor('#FFFFFF');
  sheet.getRange('A1:H1').merge();
  
  // Last Updated
  sheet.getRange('A3').setValue('Last Updated:').setFontWeight('bold');
  sheet.getRange('B3').setValue(new Date()).setNumberFormat('MM/dd/yyyy hh:mm:ss AM/PM');
  
  // Overall Performance Section
  sheet.getRange('A5').setValue('📊 OVERALL PERFORMANCE METRICS').setFontSize(14).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
  sheet.getRange('A5:H5').merge();
  
  const summaryHeaders = ['Metric', 'This Month', 'Last Month', 'All Time'];
  sheet.getRange('A6:D6').setValues([summaryHeaders])
    .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
  
  // Overall metrics with formulas
  const metricsData = [
    ['Total Assessments', '', '', `=COUNTA('Training Records'!A:A)-1`],
    ['Average Score (%)', '', '', `=IF(COUNTA('Training Records'!Q:Q)-1>0,ROUND(AVERAGE('Training Records'!Q2:Q),1),"No Data")`],
    ['Excellent Rate (%)', '', '', `=IF(COUNTA('Training Records'!I:I)-1>0,ROUND(COUNTIF('Training Records'!I2:I,"Excellent")/COUNTA('Training Records'!I2:I)*100,1),"No Data")`],
    ['Good Rate (%)', '', '', `=IF(COUNTA('Training Records'!I:I)-1>0,ROUND(COUNTIF('Training Records'!I2:I,"Good")/COUNTA('Training Records'!I2:I)*100,1),"No Data")`],
    ['Needs Improvement (%)', '', '', `=IF(COUNTA('Training Records'!I:I)-1>0,ROUND(COUNTIF('Training Records'!I2:I,"Needs Improvement")/COUNTA('Training Records'!I2:I)*100,1),"No Data")`],
    ['Active Trainees', '', '', `=COUNTA(UNIQUE('Training Records'!E2:E))`],
    ['Active Trainers', '', '', `=COUNTA(UNIQUE('Training Records'!D2:D))`]
  ];
  
  sheet.getRange(7, 1, metricsData.length, metricsData[0].length).setValues(metricsData);
  
  // Performance by Location Section
  sheet.getRange('A15').setValue('🏢 PERFORMANCE BY LOCATION').setFontSize(14).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
  sheet.getRange('A15:H15').merge();
  
  const locationHeaders = ['Location', 'Total Assessments', 'Avg Score (%)', 'Excellence Rate (%)', 'Success Rate (≥80%)', 'Most Active Trainer'];
  sheet.getRange('A16:F16').setValues([locationHeaders])
    .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
  
  // Location performance data with working formulas
  const locations = ['Original American Kitchen', 'White Buffalo', 'Cantina Añejo'];
  for (let i = 0; i < locations.length; i++) {
    const row = 17 + i;
    sheet.getRange(row, 1).setValue(locations[i]);
    sheet.getRange(row, 2).setFormula(`=COUNTIF('Training Records'!C:C,"${locations[i]}")`);
    sheet.getRange(row, 3).setFormula(`=IF(COUNTIF('Training Records'!C:C,"${locations[i]}")>0,ROUND(AVERAGEIF('Training Records'!C:C,"${locations[i]}",'Training Records'!Q:Q),1),0)`);
    sheet.getRange(row, 4).setFormula(`=IF(COUNTIF('Training Records'!C:C,"${locations[i]}")>0,ROUND(COUNTIFS('Training Records'!C:C,"${locations[i]}",'Training Records'!I:I,"Excellent")/COUNTIF('Training Records'!C:C,"${locations[i]}")*100,1),0)`);
    sheet.getRange(row, 5).setFormula(`=IF(COUNTIF('Training Records'!C:C,"${locations[i]}")>0,ROUND(COUNTIFS('Training Records'!C:C,"${locations[i]}",'Training Records'!Q:Q,">=80")/COUNTIF('Training Records'!C:C,"${locations[i]}")*100,1),0)`);
    sheet.getRange(row, 6).setValue('See Trainer Performance');
  }
  
  // Recent Activity Section
  sheet.getRange('A21').setValue('📅 RECENT ACTIVITY (Last 30 Days)').setFontSize(14).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
  sheet.getRange('A21:H21').merge();
  
  sheet.getRange('A22').setValue('Assessments This Month:').setFontWeight('bold');
  sheet.getRange('B22').setFormula('=COUNTIFS(\'Training Records\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Training Records\'!B:B,"<="&EOMONTH(TODAY(),0))');
  
  sheet.getRange('A23').setValue('Average Score This Month:').setFontWeight('bold');
  sheet.getRange('B23').setFormula('=IF(COUNTIFS(\'Training Records\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Training Records\'!B:B,"<="&EOMONTH(TODAY(),0))>0,ROUND(AVERAGEIFS(\'Training Records\'!Q:Q,\'Training Records\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Training Records\'!B:B,"<="&EOMONTH(TODAY(),0)),1),"No Data")');
  
  // Set column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 150);
  
  console.log('✓ Enhanced Analytics Dashboard created');
  return sheet;
}

/**
 * Remove the createLocationSummarySheet function as it's no longer needed
 */

/**
 * Create Location Summary sheet (detailed view by location)
 */
function createLocationSummarySheet(spreadsheet) {
  console.log('Creating Location Summary sheet...');
  
  // Safety check: if no spreadsheet passed, get the active one
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  const sheet = spreadsheet.insertSheet('Location Summary');
  
  // Title
  sheet.getRange('A1').setValue('Location Performance Summary')
    .setFontSize(16).setFontWeight('bold').setBackground('#2C5AA0').setFontColor('#FFFFFF');
  sheet.getRange('A1:H1').merge();
  
  // Last Updated
  sheet.getRange('A3').setValue('Last Updated:').setFontWeight('bold');
  sheet.getRange('B3').setValue(new Date()).setNumberFormat('MM/dd/yyyy hh:mm:ss AM/PM');
  
  const locations = ['Original American Kitchen', 'White Buffalo', 'Cantina Añejo'];
  let currentRow = 5;
  
  locations.forEach(location => {
    // Location header
    sheet.getRange(currentRow, 1).setValue(`📍 ${location.toUpperCase()}`)
      .setFontSize(12).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
    sheet.getRange(currentRow, 1, 1, 8).merge();
    currentRow++;
    
    // Metrics headers
    const headers = ['Metric', 'Value', 'This Month', 'Last Month', 'Trend', 'Details'];
    sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers])
      .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
    currentRow++;
    
    // FIXED: Location metrics with corrected formulas
    const metricsData = [
      [
        'Total Assessments', 
        `=COUNTIF('Training Records'!C:C,"${location}")`,
        // FIXED: This Month - Count records from first day of current month to today
        `=COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())`,
        // FIXED: Last Month - Count records from first day to last day of previous month
        `=COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))`,
        `=IF(AND(D${currentRow}>0,C${currentRow}>0),IF(C${currentRow}>D${currentRow},"↗️ +"&C${currentRow}-D${currentRow},IF(C${currentRow}<D${currentRow},"↘️ -"&D${currentRow}-C${currentRow},"→ Stable")),"New")`,
        'Assessment count'
      ],
      [
        'Average Score (%)', 
        `=IF(COUNTIF('Training Records'!C:C,"${location}")>0,ROUND(AVERAGEIF('Training Records'!C:C,"${location}",'Training Records'!Q:Q),1),0)`,
        // FIXED: This Month average
        `=IF(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())>0,ROUND(AVERAGEIFS('Training Records'!Q:Q,'Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY()),1),0)`,
        // FIXED: Last Month average
        `=IF(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))>0,ROUND(AVERAGEIFS('Training Records'!Q:Q,'Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1)),1),0)`,
        `=IF(AND(D${currentRow+1}>0,C${currentRow+1}>0),IF(C${currentRow+1}>D${currentRow+1},"↗️ +"&ROUND(C${currentRow+1}-D${currentRow+1},1),IF(C${currentRow+1}<D${currentRow+1},"↘️ -"&ROUND(D${currentRow+1}-C${currentRow+1},1),"→ Stable")),"New")`,
        'Performance average'
      ],
      [
        'Excellence Rate (%)', 
        `=IF(COUNTIF('Training Records'!C:C,"${location}")>0,ROUND(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!I:I,"Excellent")/COUNTIF('Training Records'!C:C,"${location}")*100,1),0)`,
        // FIXED: This Month excellence rate
        `=IF(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())>0,ROUND(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!I:I,"Excellent",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())/COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())*100,1),0)`,
        // FIXED: Last Month excellence rate
        `=IF(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))>0,ROUND(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!I:I,"Excellent",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))/COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))*100,1),0)`,
        `=IF(AND(D${currentRow+2}>0,C${currentRow+2}>0),IF(C${currentRow+2}>D${currentRow+2},"↗️ +"&ROUND(C${currentRow+2}-D${currentRow+2},1),IF(C${currentRow+2}<D${currentRow+2},"↘️ -"&ROUND(D${currentRow+2}-C${currentRow+2},1),"→ Stable")),"New")`,
        'Excellent ratings %'
      ],
      [
        'Success Rate (≥80%)', 
        `=IF(COUNTIF('Training Records'!C:C,"${location}")>0,ROUND(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!Q:Q,">=80")/COUNTIF('Training Records'!C:C,"${location}")*100,1),0)`,
        // FIXED: This Month success rate
        `=IF(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())>0,ROUND(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!Q:Q,">=80",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())/COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Training Records'!B:B,"<="&TODAY())*100,1),0)`,
        // FIXED: Last Month success rate
        `=IF(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))>0,ROUND(COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!Q:Q,">=80",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))/COUNTIFS('Training Records'!C:C,"${location}",'Training Records'!B:B,">="&DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1),'Training Records'!B:B,"<="&EOMONTH(TODAY(),-1))*100,1),0)`,
        `=IF(AND(D${currentRow+3}>0,C${currentRow+3}>0),IF(C${currentRow+3}>D${currentRow+3},"↗️ +"&ROUND(C${currentRow+3}-D${currentRow+3},1),IF(C${currentRow+3}<D${currentRow+3},"↘️ -"&ROUND(D${currentRow+3}-C${currentRow+3},1),"→ Stable")),"New")`,
        'Scores 80% or higher'
      ],
      [
        'Active Trainers', 
        // Using hardcoded trainer count from TRAINER_COUNTS constant (Total available trainers at this location)
        `=${TRAINER_COUNTS[location]}`,
        // FIXED: This Month - Count unique trainers who actually conducted assessments
        `=IFERROR(SUMPRODUCT(1/COUNTIF('Training Records'!D:D,IF(('Training Records'!C:C="${location}")*('Training Records'!B:B>=DATE(YEAR(TODAY()),MONTH(TODAY()),1))*('Training Records'!B:B<=TODAY())*('Training Records'!D:D<>""),'Training Records'!D:D))),0)`,
        // FIXED: Last Month - Count unique trainers who actually conducted assessments
        `=IFERROR(SUMPRODUCT(1/COUNTIF('Training Records'!D:D,IF(('Training Records'!C:C="${location}")*('Training Records'!B:B>=DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1))*('Training Records'!B:B<=EOMONTH(TODAY(),-1))*('Training Records'!D:D<>""),'Training Records'!D:D))),0)`,
        `=IF(AND(D${currentRow+4}>0,C${currentRow+4}>0),IF(C${currentRow+4}>D${currentRow+4},"↗️ +"&C${currentRow+4}-D${currentRow+4},IF(C${currentRow+4}<D${currentRow+4},"↘️ -"&D${currentRow+4}-C${currentRow+4},"→ Stable")),"New")`,
        'Total trainers on roster / Active this period'
      ]
    ];
    
    sheet.getRange(currentRow, 1, metricsData.length, metricsData[0].length).setValues(metricsData);
    currentRow += metricsData.length + 2; // Add space between locations
  });
  
  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 150);
  
  console.log('✓ Location Summary sheet created with FIXED formulas');
  return sheet;
}

/**
 * Create Monthly Trainer Performance sheet (dynamic - only shows data when it exists)
 */
function createTrainerPerformanceSheet(spreadsheet) {
  console.log('Creating Dynamic Monthly Trainer Performance sheet...');
  
  const sheet = spreadsheet.insertSheet('Trainer Performance');
  
  let currentColumn = 1;
  
  // Create headers for each location
  const locations = ['Cantina Añejo', 'Original American Kitchen', 'White Buffalo'];
  const locationColors = ['#FCE5CD', '#CFE2F3', '#D9D2E9']; // Light orange 3, Light blue 3, Light purple 3
  
  locations.forEach((location, locationIndex) => {
    // Location header with color coding
    sheet.getRange(1, currentColumn).setValue(location)
      .setFontSize(12).setFontWeight('bold').setBackground(locationColors[locationIndex]).setFontColor('#000000');
    sheet.getRange(1, currentColumn, 1, 5).merge();
    

    // Column headers
    const headers = ['Month', 'Trainer', 'Assessments Given', 'Avg Score Given (%)', 'Success Rate (%)'];
    sheet.getRange(2, currentColumn, 1, headers.length).setValues([headers])
      .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(currentColumn, 80);     // Month
    sheet.setColumnWidth(currentColumn + 1, 150); // Trainer
    sheet.setColumnWidth(currentColumn + 2, 120); // Assessments
    sheet.setColumnWidth(currentColumn + 3, 120); // Avg Score
    sheet.setColumnWidth(currentColumn + 4, 120); // Success Rate
    // Force the data columns to be Numbers so they don't turn into Dates
    sheet.getRange(3, currentColumn + 2, 100, 3).setNumberFormat("0");
    
    // Move to next location
    currentColumn += 6; // 5 columns + 1 spacing
  });
  
  // Add note about dynamic data
  sheet.getRange('A20').setValue('📝 Note: Data will automatically populate here when training assessments are submitted')
    .setFontStyle('italic').setFontColor('#666666');
  sheet.getRange('A20:R20').merge();
  
  console.log('✓ Dynamic Monthly Trainer Performance sheet created');
  return sheet;
}

/**
 * Helper function to get month number
 */
function getMonthNumber(monthName) {
  const months = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
  };
  return months[monthName] || 1;
}

/**
 * Update analytics data
 */
// updateAnalytics function removed - replaced by optimized populate functions

/**
 * Update location summary with fresh data
 */
function updateLocationSummary(spreadsheet) {
  try {
    if (!spreadsheet) {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    const sheet = spreadsheet.getSheetByName('Location Summary');
    if (!sheet) {
      console.log('Location Summary sheet not found');
      return;
    }
    
    // Update timestamp
    sheet.getRange('B3').setValue(new Date());
    
    // Force recalculation by refreshing formulas
    SpreadsheetApp.flush();
    
    console.log('✓ Location Summary updated');
  } catch (error) {
    console.log('Error in updateLocationSummary:', error.toString());
  }
}

/**
 * Update monthly trainer performance data (dynamic - only adds rows when data exists)
 */
function updateMonthlyTrainerPerformance(spreadsheet, data) {
  try {
    // Ensure we have valid parameters
    if (!data) {
      console.log('No data provided to updateMonthlyTrainerPerformance');
      return;
    }
    
    if (!spreadsheet) {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    const sheet = spreadsheet.getSheetByName('Trainer Performance');
    if (!sheet) {
      console.log('Trainer Performance sheet not found');
      return;
    }
    
    // Validate required data fields
    if (!data.trainer || !data.location) {
      console.log('Missing required data fields for trainer performance');
      return;
    }
    
    // Get current month/year from the submission
    const submissionDate = new Date();
    const monthName = Utilities.formatDate(submissionDate, Session.getScriptTimeZone(), 'MMM yyyy');
    
    // Determine which column section to use based on location
    let columnOffset = 1;
    if (data.location === 'Original American Kitchen') {
      columnOffset = 7; // Second section
    } else if (data.location === 'White Buffalo') {
      columnOffset = 13; // Third section
    }
    
    // Find if this trainer/month combination already exists
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let targetRow = -1;
    let monthRowFound = false;
    let lastDataRow = 2; // Start after headers
    
    // Look for existing data in this location's section
    for (let i = 2; i < values.length; i++) {
      const currentLocationCol = columnOffset - 1; // Convert to 0-based index
      
      // Check if this is our location's section and if there's data
      if (values[i][currentLocationCol + 1] === data.trainer && 
          values[i][currentLocationCol] === monthName) {
        targetRow = i + 1; // Convert to 1-based
        break;
      }
      
      // Track the last row with data in this section
      if (values[i][currentLocationCol] || values[i][currentLocationCol + 1]) {
        lastDataRow = i + 1;
      }
    }
    
    if (targetRow === -1) {
      // Need to add new row - find the right place
      let insertRow = lastDataRow + 1;
      
      // Check if we need to add a month header first
      let needsMonthHeader = true;
      for (let i = 2; i < values.length; i++) {
        if (values[i][columnOffset - 1] === monthName) {
          needsMonthHeader = false;
          break;
        }
      }
      
      if (needsMonthHeader) {
        // Add month header row
        sheet.getRange(insertRow, columnOffset).setValue(monthName)
          .setFontWeight('bold').setBackground('#E8F4FD');
        insertRow++;
      }
      
      // Add trainer data row
      const assessmentCount = 1;
      const avgScore = data.percentage || 0;
      const successRate = (data.percentage >= 80) ? 100 : 0;
      
      sheet.getRange(insertRow, columnOffset).setValue(''); // Month (empty for trainer rows)
      sheet.getRange(insertRow, columnOffset + 1).setValue(data.trainer);
      sheet.getRange(insertRow, columnOffset + 2).setValue(assessmentCount).setNumberFormat("0");
      sheet.getRange(insertRow, columnOffset + 3).setValue(avgScore).setNumberFormat("0");
      sheet.getRange(insertRow, columnOffset + 4).setValue(successRate).setNumberFormat("0");
      
    } else {
      // Update existing row
      const currentAssessments = values[targetRow - 1][columnOffset + 1] || 0;
      const currentAvg = values[targetRow - 1][columnOffset + 2] || 0;
      const currentSuccessRate = values[targetRow - 1][columnOffset + 3] || 0;
      
      const newAssessments = currentAssessments + 1;
      const newAvg = ((currentAvg * currentAssessments) + data.percentage) / newAssessments;
      
      const currentSuccesses = Math.round((currentSuccessRate / 100) * currentAssessments);
      const newSuccesses = currentSuccesses + (data.percentage >= 80 ? 1 : 0);
      const newSuccessRate = (newSuccesses / newAssessments) * 100;
      
      sheet.getRange(targetRow, columnOffset + 2).setValue(newAssessments).setNumberFormat("0");
      sheet.getRange(targetRow, columnOffset + 3).setValue(Math.round(newAvg)).setNumberFormat("0");
      sheet.getRange(targetRow, columnOffset + 4).setValue(Math.round(newSuccessRate)).setNumberFormat("0");
    }
    
    console.log(`✓ Updated trainer performance for ${data.trainer} at ${data.location} for ${monthName}`);
    
  } catch (error) {
    console.log('Error in updateMonthlyTrainerPerformance:', error.toString());
  }
}

/**
 * Populate Analytics Dashboard with real data from Training Records
 */
function populateAnalyticsDashboardFromExistingData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the Training Records sheet
    const trainingRecordsSheet = spreadsheet.getSheetByName('Training Records');
    if (!trainingRecordsSheet) {
      ui.alert('Error', 'Training Records sheet not found!', ui.ButtonSet.OK);
      return;
    }
    
    // Get the Analytics Dashboard sheet
    let analyticsSheet = spreadsheet.getSheetByName('Analytics Dashboard');
    if (!analyticsSheet) {
      analyticsSheet = createAnalyticsSheet(spreadsheet);
    }
    
    // Get all data from Training Records
    const data = trainingRecordsSheet.getDataRange().getValues();
    const headers = data[0];
    const records = data.slice(1);
    
    // Use fixed column positions
    const timestampCol = 1;  // Column B - Timestamp
    const locationCol = 2;   // Column C - Location  
    const trainerCol = 3;    // Column D - Trainer
    const traineeCol = 4;    // Column E - Trainee
    const performanceLevelCol = 8; // Column I - Performance Level
    const percentageCol = 16; // Column Q - Percentage
    
    // Calculate time periods
    const now = new Date();
    const thisWeekStart = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
    const lastWeekStart = new Date(thisWeekStart.getTime() - 7 * 24 * 60 * 60 * 1000);
    const thisMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const lastMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const lastMonthEnd = new Date(thisMonthStart.getTime() - 1);
    
    // Initialize metrics
    let metrics = {
      thisWeek: { total: 0, totalScore: 0, excellent: 0, good: 0, needsImprovement: 0, trainers: new Set(), trainees: new Set() },
      lastWeek: { total: 0, totalScore: 0, excellent: 0, good: 0, needsImprovement: 0, trainers: new Set(), trainees: new Set() },
      thisMonth: { total: 0, totalScore: 0, excellent: 0, good: 0, needsImprovement: 0, trainers: new Set(), trainees: new Set() },
      allTime: { total: 0, totalScore: 0, excellent: 0, good: 0, needsImprovement: 0, trainers: new Set(), trainees: new Set() },
      locations: {}
    };
    
    // Process each record
    records.forEach(record => {
      if (!record[timestampCol] || !record[locationCol]) return;
      
      const timestamp = new Date(record[timestampCol]);
      const location = record[locationCol];
      const trainer = record[trainerCol];
      const trainee = record[traineeCol];
      const performanceLevel = record[performanceLevelCol] || '';
      const percentage = parseFloat(record[percentageCol]) || 0;
      
      if (isNaN(timestamp.getTime()) || percentage === 0) return;
      
      // Initialize location if not exists
      if (!metrics.locations[location]) {
        metrics.locations[location] = { total: 0, totalScore: 0, excellent: 0, good: 0, needsImprovement: 0 };
      }
      
      // Update location metrics
      metrics.locations[location].total++;
      metrics.locations[location].totalScore += percentage;
      if (performanceLevel === 'Excellent') metrics.locations[location].excellent++;
      if (performanceLevel === 'Good') metrics.locations[location].good++;
      if (performanceLevel === 'Needs Improvement') metrics.locations[location].needsImprovement++;
      
      // Update time-based metrics
      if (timestamp >= thisWeekStart) {
        updateMetrics(metrics.thisWeek, percentage, performanceLevel, trainer, trainee);
      }
      if (timestamp >= lastWeekStart && timestamp < thisWeekStart) {
        updateMetrics(metrics.lastWeek, percentage, performanceLevel, trainer, trainee);
      }
      if (timestamp >= thisMonthStart) {
        updateMetrics(metrics.thisMonth, percentage, performanceLevel, trainer, trainee);
      }
      updateMetrics(metrics.allTime, percentage, performanceLevel, trainer, trainee);
    });
    
    // Update the dashboard with real data
    updateAnalyticsDashboardValues(analyticsSheet, metrics);
    
    ui.alert('Success', `✓ Analytics Dashboard updated with real data!\n\n• Fixed calculation errors\n• Added meaningful time periods\n• Real performance metrics\n• Processed ${records.length} training records`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Failed to update analytics: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Helper function to update metrics
 */
function updateMetrics(period, percentage, performanceLevel, trainer, trainee) {
  period.total++;
  period.totalScore += percentage;
  if (performanceLevel === 'Excellent') period.excellent++;
  if (performanceLevel === 'Good') period.good++;
  if (performanceLevel === 'Needs Improvement') period.needsImprovement++;
  if (trainer) period.trainers.add(trainer);
  if (trainee) period.trainees.add(trainee);
}

/**
 * Update Analytics Dashboard with calculated values
 */
function updateAnalyticsDashboardValues(sheet, metrics) {
  // Update timestamp
  sheet.getRange('B3').setValue(new Date());
  
  // Clear existing data in metrics section
  sheet.getRange('B7:E13').clear();
  
  // Overview metrics table
  const overviewData = [
    ['Total Assessments', 
     metrics.thisWeek.total, 
     metrics.lastWeek.total, 
     metrics.thisMonth.total, 
     metrics.allTime.total],
    ['Average Score (%)', 
     metrics.thisWeek.total > 0 ? Math.round(metrics.thisWeek.totalScore / metrics.thisWeek.total) : 0,
     metrics.lastWeek.total > 0 ? Math.round(metrics.lastWeek.totalScore / metrics.lastWeek.total) : 0,
     metrics.thisMonth.total > 0 ? Math.round(metrics.thisMonth.totalScore / metrics.thisMonth.total) : 0,
     metrics.allTime.total > 0 ? Math.round(metrics.allTime.totalScore / metrics.allTime.total) : 0],
    ['Excellent Performance %', 
     metrics.thisWeek.total > 0 ? Math.round((metrics.thisWeek.excellent / metrics.thisWeek.total) * 100) : 0,
     metrics.lastWeek.total > 0 ? Math.round((metrics.lastWeek.excellent / metrics.lastWeek.total) * 100) : 0,
     metrics.thisMonth.total > 0 ? Math.round((metrics.thisMonth.excellent / metrics.thisMonth.total) * 100) : 0,
     metrics.allTime.total > 0 ? Math.round((metrics.allTime.excellent / metrics.allTime.total) * 100) : 0],
    ['Good Performance %', 
     metrics.thisWeek.total > 0 ? Math.round((metrics.thisWeek.good / metrics.thisWeek.total) * 100) : 0,
     metrics.lastWeek.total > 0 ? Math.round((metrics.lastWeek.good / metrics.lastWeek.total) * 100) : 0,
     metrics.thisMonth.total > 0 ? Math.round((metrics.thisMonth.good / metrics.thisMonth.total) * 100) : 0,
     metrics.allTime.total > 0 ? Math.round((metrics.allTime.good / metrics.allTime.total) * 100) : 0],
    ['Needs Improvement %', 
     metrics.thisWeek.total > 0 ? Math.round((metrics.thisWeek.needsImprovement / metrics.thisWeek.total) * 100) : 0,
     metrics.lastWeek.total > 0 ? Math.round((metrics.lastWeek.needsImprovement / metrics.lastWeek.total) * 100) : 0,
     metrics.thisMonth.total > 0 ? Math.round((metrics.thisMonth.needsImprovement / metrics.thisMonth.total) * 100) : 0,
     metrics.allTime.total > 0 ? Math.round((metrics.allTime.needsImprovement / metrics.allTime.total) * 100) : 0],
    ['Unique Trainees', 
     metrics.thisWeek.trainees.size,
     metrics.lastWeek.trainees.size,
     metrics.thisMonth.trainees.size,
     metrics.allTime.trainees.size],
    ['Active Trainers', 
     metrics.thisWeek.trainers.size,
     metrics.lastWeek.trainers.size,
     metrics.thisMonth.trainers.size,
     metrics.allTime.trainers.size]
  ];
  
  sheet.getRange(7, 1, overviewData.length, overviewData[0].length).setValues(overviewData);
  
  // Location performance section (starting around row 17)
  const locations = ['Cantina Añejo', 'Original American Kitchen', 'White Buffalo'];
  for (let i = 0; i < locations.length; i++) {
    const location = locations[i];
    const row = 17 + i;
    const locationData = metrics.locations[location] || { total: 0, totalScore: 0, excellent: 0, needsImprovement: 0 };
    
    sheet.getRange(row, 1).setValue(location);
    sheet.getRange(row, 2).setValue(locationData.total);
    sheet.getRange(row, 3).setValue(locationData.total > 0 ? Math.round(locationData.totalScore / locationData.total) : 0);
    sheet.getRange(row, 4).setValue(locationData.total > 0 ? Math.round((locationData.excellent / locationData.total) * 100) : 0);
    sheet.getRange(row, 5).setValue(locationData.total > 0 ? Math.round((locationData.needsImprovement / locationData.total) * 100) : 0);
  }
}

/**
 * Send email notification to managers
 */
// Redundant email function removed - using sendNotificationEmailSafe instead

/**
 * Enhanced manual analytics refresh
 */
function refreshAnalytics() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Update all analytics with real data
    populateAnalyticsDashboardFromExistingData();
    populateTrainerPerformanceFromExistingData();
    populateLocationPerformanceFromExistingData();
    updateLocationSummary(spreadsheet);
    
    // --- START OF FORMATTING FIX ---
    // This forces the "Trainer Performance" sheet to stop showing dates
    const trainerSheet = spreadsheet.getSheetByName('Trainer Performance');
    if (trainerSheet) {
      // Column C,D,E (Cantina), I,J,K (OAK), O,P,Q (White Buffalo)
      const dataColumns = [3, 4, 5, 9, 10, 11, 15, 16, 17];
      dataColumns.forEach(col => {
        // Formats rows 3 through 100 as plain numbers
        trainerSheet.getRange(3, col, 97, 1).setNumberFormat("0");
      });
      console.log('✓ Forced Number Formatting on Trainer Performance');
    }
    // --- END OF FORMATTING FIX ---

    // Force recalculation of analytics sheets
    const sheets = ['Analytics Dashboard', 'Location Summary', 'Trainer Performance', 'Monthly Location Performance'];
    sheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        if (sheetName === 'Analytics Dashboard') {
          try {
            sheet.getRange('B3').setValue(new Date());
          } catch (e) {}
        }
        SpreadsheetApp.flush();
      }
    });
    
    ui.alert('Success', '✓ All Analytics Refreshed and Formatted!', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Failed to refresh analytics: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Populate Monthly Location Performance sheet with existing data
 */
function populateLocationPerformanceFromExistingData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the Training Records sheet
    const trainingRecordsSheet = spreadsheet.getSheetByName('Training Records');
    if (!trainingRecordsSheet) {
      ui.alert('Error', 'Training Records sheet not found!', ui.ButtonSet.OK);
      return;
    }
    
    // Check if Monthly Location Performance sheet exists, if not create it
    let locationSheet = spreadsheet.getSheetByName('Monthly Location Performance');
    if (!locationSheet) {
      locationSheet = createMonthlyLocationPerformanceSheet(spreadsheet);
    }
    
    // Get all data from Training Records (skip header row)
    const data = trainingRecordsSheet.getDataRange().getValues();
    const headers = data[0];
    const records = data.slice(1);
    
    // Use fixed column positions
    const timestampCol = 1;  // Column B - Timestamp
    const locationCol = 2;   // Column C - Location  
    const performanceLevelCol = 8; // Column I - Performance Level
    const percentageCol = 16; // Column Q - Percentage
    
    // Group data by location and month
    const groupedData = {};
    
    records.forEach(record => {
      if (!record[locationCol] || !record[timestampCol]) return;
      
      const location = record[locationCol];
      const timestamp = new Date(record[timestampCol]);
      const performanceLevel = record[performanceLevelCol] || '';
      const percentage = parseFloat(record[percentageCol]) || 0;
      
      // Skip invalid data
      if (isNaN(timestamp.getTime())) return;
      
      const monthKey = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MMM yyyy');
      const key = `${location}|${monthKey}`;
      
      if (!groupedData[key]) {
        groupedData[key] = {
          location: location,
          month: monthKey,
          totalAssessments: 0,
          totalScore: 0,
          excellentCount: 0,
          goodCount: 0,
          needsImprovementCount: 0,
          successCount: 0
        };
      }
      
      groupedData[key].totalAssessments++;
      groupedData[key].totalScore += percentage;
      
      if (performanceLevel === 'Excellent') {
        groupedData[key].excellentCount++;
      } else if (performanceLevel === 'Good') {
        groupedData[key].goodCount++;
      } else if (performanceLevel === 'Needs Improvement') {
        groupedData[key].needsImprovementCount++;
      }
      
      if (percentage >= 80) {
        groupedData[key].successCount++;
      }
    });
    
    // Clear existing data (keep title row only)
    const lastRow = locationSheet.getLastRow();
    if (lastRow > 1) {
      // Use clear() instead of clearContent() to remove formatting AND content
      locationSheet.getRange(2, 1, lastRow - 1, locationSheet.getLastColumn()).clear();
    }
    
    // Sort data by location and month
    const sortedData = Object.values(groupedData).sort((a, b) => {
      if (a.location !== b.location) {
        return a.location.localeCompare(b.location);
      }
      return new Date(a.month).getTime() - new Date(b.month).getTime();
    });
    
    // Populate data
    let currentRow = 2;
    const locations = ['Cantina Añejo', 'Original American Kitchen', 'White Buffalo'];
    const locationColors = ['#FCE5CD', '#CFE2F3', '#D9D2E9']; // Light orange 3, Light blue 3, Light purple 3
    
    locations.forEach((location, locationIndex) => {
      // Add location header with color coding (NO MERGE - use background color across cells instead)
      locationSheet.getRange(currentRow, 1).setValue(location)
        .setFontSize(12).setFontWeight('bold').setBackground(locationColors[locationIndex]).setFontColor('#000000');
      // Apply background color to all cells in the row instead of merging
      locationSheet.getRange(currentRow, 1, 1, 7).setBackground(locationColors[locationIndex]);
      currentRow++;
      
      // Add column headers for this location
      const headers = ['Month', 'Total Assessments', 'Average Score', 'Excellent Count', 'Good Count', 'Needs Improvement Count', 'Success Rate (%)'];
      locationSheet.getRange(currentRow, 1, 1, headers.length).setValues([headers])
        .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
      currentRow++;
      
      // Add data for this location
      const locationData = sortedData.filter(item => item.location === location);
      
      locationData.forEach(item => {
        const avgScore = item.totalAssessments > 0 ? Math.round(item.totalScore / item.totalAssessments) : 0;
        const successRate = item.totalAssessments > 0 ? Math.round((item.successCount / item.totalAssessments) * 100) : 0;
        
        locationSheet.getRange(currentRow, 1).setValue(item.month);
        locationSheet.getRange(currentRow, 2).setValue(item.totalAssessments);
        locationSheet.getRange(currentRow, 3).setValue(avgScore);
        locationSheet.getRange(currentRow, 4).setValue(item.excellentCount);
        locationSheet.getRange(currentRow, 5).setValue(item.goodCount);
        locationSheet.getRange(currentRow, 6).setValue(item.needsImprovementCount);
        locationSheet.getRange(currentRow, 7).setValue(successRate);
        
        currentRow++;
      });
      
      // Add spacing between locations
      currentRow += 2;
    });

    // Force recalculation after all data is populated
    SpreadsheetApp.flush();

    ui.alert('Success', `✓ Monthly Location Performance populated with data!\n\n• Processed ${records.length} training records\n• Grouped by location and month\n• Shows actual performance metrics`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Failed to populate location performance: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Create Monthly Location Performance sheet
 */
function createMonthlyLocationPerformanceSheet(spreadsheet) {
  console.log('Creating Monthly Location Performance sheet...');
  
  // Ensure we have a valid spreadsheet object
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  const sheet = spreadsheet.insertSheet('Monthly Location Performance');
  
  // Title
  sheet.getRange('A1').setValue('Monthly Location Performance Analysis')
    .setFontSize(16).setFontWeight('bold').setBackground('#2C5AA0').setFontColor('#FFFFFF');
  sheet.getRange('A1:G1').merge();
  
  // Set column widths
  sheet.setColumnWidth(1, 120); // Month
  sheet.setColumnWidth(2, 140); // Total Assessments
  sheet.setColumnWidth(3, 130); // Average Score
  sheet.setColumnWidth(4, 120); // Excellent Count
  sheet.setColumnWidth(5, 100);  // Good Count
  sheet.setColumnWidth(6, 160); // Needs Improvement Count
  sheet.setColumnWidth(7, 120); // Success Rate
  
  console.log('✓ Monthly Location Performance sheet created');
  return sheet;
}

/**
 * Populate Trainer Performance sheet with existing data from Training Records
 */
function populateTrainerPerformanceFromExistingData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the Training Records sheet
    const trainingRecordsSheet = spreadsheet.getSheetByName('Training Records');
    if (!trainingRecordsSheet) {
      ui.alert('Error', 'Training Records sheet not found!', ui.ButtonSet.OK);
      return;
    }
    
    // Get the Trainer Performance sheet
    const trainerSheet = spreadsheet.getSheetByName('Trainer Performance');
    if (!trainerSheet) {
      ui.alert('Error', 'Trainer Performance sheet not found!', ui.ButtonSet.OK);
      return;
    }
    
    // Get all data from Training Records (skip header row)
    const data = trainingRecordsSheet.getDataRange().getValues();
    const headers = data[0];
    const records = data.slice(1);
    
    // Find column indices using exact positions from your headers
    // Record ID=0, Timestamp=1, Location=2, Trainer=3, Trainee=4, Position=5, Training Day=6, 
    // Shift=7, Performance Level=8, Overall Notes=9, Knowledge Score=10, Technical Score=11, 
    // Service Score=12, Teamwork Score=13, Professionalism Score=14, Total Score=15, Percentage=16
    const timestampCol = 1;  // Column B - Timestamp
    const locationCol = 2;   // Column C - Location  
    const trainerCol = 3;    // Column D - Trainer
    const percentageCol = 16; // Column Q - Percentage
    
    console.log('Using fixed column positions:', {
      timestamp: timestampCol,
      location: locationCol,
      trainer: trainerCol,
      percentage: percentageCol
    });
    console.log('Sample headers:', headers.slice(0, 5), '...', headers.slice(15, 17));
    
    // No need to check for missing columns since we're using fixed positions
    // Group data by location, trainer, and month
    const groupedData = {};
    
    records.forEach(record => {
      if (!record[trainerCol] || !record[locationCol] || !record[timestampCol]) return;
      
      const location = record[locationCol];
      const trainer = record[trainerCol];
      const timestamp = new Date(record[timestampCol]);
      const percentage = parseFloat(record[percentageCol]) || 0;
      
      // Skip invalid data
      if (isNaN(timestamp.getTime())) return;
      
      const monthKey = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MMM yyyy');
      const key = `${location}|${trainer}|${monthKey}`;
      
      if (!groupedData[key]) {
        groupedData[key] = {
          location: location,
          trainer: trainer,
          month: monthKey,
          assessments: 0,
          totalScore: 0,
          successCount: 0
        };
      }
      
      groupedData[key].assessments++;
      groupedData[key].totalScore += percentage;
      if (percentage >= 80) {
        groupedData[key].successCount++;
      }
    });
    
    // Clear existing data (keep headers and location titles)
    const lastRow = trainerSheet.getLastRow();
    if (lastRow > 2) {
      // Use clear() instead of clearContent() to remove formatting AND content
      // This prevents orphaned formatting from accumulating
      trainerSheet.getRange(3, 1, lastRow - 2, trainerSheet.getLastColumn()).clear();
    }
    
    // Column positions for each location
    const locationColumns = {
      'Cantina Añejo': 1,
      'Original American Kitchen': 7,
      'White Buffalo': 13
    };
    
    // Sort data by location and month
    const sortedData = Object.values(groupedData).sort((a, b) => {
      if (a.location !== b.location) {
        return a.location.localeCompare(b.location);
      }
      return new Date(a.month).getTime() - new Date(b.month).getTime();
    });
    
    // Populate data by location
    Object.keys(locationColumns).forEach(location => {
      const startCol = locationColumns[location];
      let currentRow = 3;
      let currentMonth = '';
      
      // Filter data for this location
      const locationData = sortedData.filter(item => item.location === location);
      
      locationData.forEach(item => {
        // Add month header if new month (only in month column)
        if (item.month !== currentMonth) {
          // Place month in the Month column only
          trainerSheet.getRange(currentRow, startCol).setValue(item.month)
            .setFontWeight('bold').setBackground('#E8F4FD');
          // Clear other columns in this row
          trainerSheet.getRange(currentRow, startCol + 1, 1, 4).clearContent();
          currentMonth = item.month;
          currentRow++;
        }
        
        // Add trainer data in separate row
        const avgScore = Math.round(item.totalScore / item.assessments);
        const successRate = Math.round((item.successCount / item.assessments) * 100);
        
        // Clear month cell for trainer rows, then add trainer data
        trainerSheet.getRange(currentRow, startCol).setValue('');
        trainerSheet.getRange(currentRow, startCol + 1).setValue(item.trainer);
        trainerSheet.getRange(currentRow, startCol + 2).setValue(item.assessments);
        trainerSheet.getRange(currentRow, startCol + 3).setValue(avgScore);
        trainerSheet.getRange(currentRow, startCol + 4).setValue(successRate);
        
        currentRow++;
      });
    });

    // Force recalculation after all data is populated
    SpreadsheetApp.flush();

    ui.alert('Success', `✓ Trainer Performance populated with existing data!\n\n• Processed ${records.length} training records\n• Grouped by location, trainer, and month\n• Data now visible in all three location sections`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Failed to populate data: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Apply color coding to existing location headers
 */
function applyLocationColorCoding() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const locationColors = {
      'Cantina Añejo': '#FCE5CD',      // Light orange 3
      'Original American Kitchen': '#CFE2F3', // Light blue 3  
      'White Buffalo': '#D9D2E9'       // Light purple 3
    };
    
    // Update Trainer Performance sheet
    const trainerSheet = spreadsheet.getSheetByName('Trainer Performance');
    if (trainerSheet) {
      const locations = ['Cantina Añejo', 'Original American Kitchen', 'White Buffalo'];
      let currentColumn = 1;
      
      locations.forEach((location, locationIndex) => {
        trainerSheet.getRange(1, currentColumn).setBackground(locationColors[location]).setFontColor('#000000');
        currentColumn += 6; // Move to next location section
      });
    }
    
    // Update Monthly Location Performance sheet
    const locationSheet = spreadsheet.getSheetByName('Monthly Location Performance');
    if (locationSheet) {
      const data = locationSheet.getDataRange().getValues();
      
      // Find location headers and apply colors
      for (let row = 1; row <= data.length; row++) {
        const cellValue = data[row - 1][0]; // Column A
        if (locationColors[cellValue]) {
          locationSheet.getRange(row, 1, 1, 7).setBackground(locationColors[cellValue]).setFontColor('#000000');
        }
      }
    }
    
    ui.alert('Success', '✓ Location color coding applied!\n\n🍊 Cantina Añejo - Light Orange\n🔵 Original American Kitchen - Light Blue\n🟣 White Buffalo - Light Purple', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Failed to apply color coding: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Initialize the complete system
 */
function initializeSystem() {
  console.log('=== INITIALIZING COMPLETE SYSTEM ===');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Initialize all sheets
    initializeAllSheets(spreadsheet);
    
    // Delete default sheet if it exists
    const defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (defaultSheet && spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(defaultSheet);
      console.log('Deleted default sheet "Sheet1"');
    }

    // Dashboard metrics will be populated via menu functions
    console.log('✓ Complete system initialized successfully!');
    console.log('Spreadsheet URL:', spreadsheet.getUrl());
    
    return {
      success: true,
      message: 'Complete system ready!',
      spreadsheetUrl: spreadsheet.getUrl()
    };
    
  } catch (error) {
    console.log('❌ Initialization failed:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Recreate Location Summary sheet with FIXED formulas
 * Use this after updating the code to fix the Active Trainers count
 */
function recreateLocationSummarySheet() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Recreate Location Summary Sheet',
    'This will DELETE the existing Location Summary sheet and create a new one with FIXED formulas.\n\n' +
    'The new sheet will have:\n' +
    '✅ Correct "This Month" calculations\n' +
    '✅ Correct "Last Month" calculations\n' +
    '✅ Accurate Active Trainer counts\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Cancelled', 'Location Summary sheet was not modified.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete the existing Location Summary sheet if it exists
    const existingSheet = spreadsheet.getSheetByName('Location Summary');
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
      console.log('Deleted old Location Summary sheet');
    }
    
    // Create new Location Summary sheet with fixed formulas
    createLocationSummarySheet(spreadsheet);
    
    ui.alert(
      'Success!',
      '✅ Location Summary sheet recreated with FIXED formulas!\n\n' +
      'The sheet now shows:\n' +
      '• Accurate "This Month" data (January 2026)\n' +
      '• Accurate "Last Month" data (December 2025)\n' +
      '• Correct Active Trainer counts:\n' +
      '  - White Buffalo: Should show 2 trainers\n' +
      '  - Original American Kitchen: Should show ~8 trainers\n' +
      '  - Cantina Añejo: Should show ~12 trainers',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('Error', 'Failed to recreate Location Summary sheet:\n\n' + error.toString(), ui.ButtonSet.OK);
    console.error('Error recreating Location Summary:', error);
  }
}
/**
 * Formats the PAID VALIDATION sheet
 */
function formatPaidValidationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PAID VALIDATION');
  
  // If sheet doesn't exist, create it
  if (!sheet) {
    sheet = ss.insertSheet('PAID VALIDATION');
  }

  const lastRow = 100; // Adjust based on your needs
  
  // 1. Clear existing weird formatting but keep data
  sheet.getRange(1, 1, lastRow, 12).clearFormat();

  // 2. Define the three sections (Locations)
  // Section 1: A-C (Cantina), Section 2: E-G (OAK), Section 3: I-K (White Buffalo)
  const sections = [1, 5, 9]; 
  const colors = ['#FCE5CD', '#CFE2F3', '#D9D2E9']; // Orange, Blue, Purple
  
  sections.forEach((col, index) => {
    // Format Location Header (Row 1)
    sheet.getRange(1, col, 1, 3).merge()
      .setBackground(colors[index])
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Format Column Headers (Row 2: Date, Trainer, Paid)
    const headerRange = sheet.getRange(2, col, 1, 3);
    headerRange.setBackground('#8FA4A7')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Force "Date" column format (Jan 2026 style)
    sheet.getRange(3, col, lastRow, 1).setNumberFormat('MMM yyyy');

    // Force "Trainer" column to Plain Text
    sheet.getRange(3, col + 1, lastRow, 1).setNumberFormat('@');

    // Force "PAID" column to be Checkboxes
    const checkboxRange = sheet.getRange(3, col + 2, lastRow, 1);
    checkboxRange.insertCheckboxes();
    checkboxRange.setHorizontalAlignment('center');
  });

  // 3. Set Column Widths for readability
  const widths = [100, 150, 60, 30, 100, 150, 60, 30, 100, 150, 60];
  widths.forEach((width, i) => {
    sheet.setColumnWidth(i + 1, width);
  });

  // Freeze the top 2 rows so headers stay put
  sheet.setFrozenRows(2);
  
  SpreadsheetApp.getUi().alert('PAID VALIDATION formatting complete!');
}
/**
 * Resolves location aliases (CA, OAK, WB) to full names
 * @param {string} input The user's input (alias or full name)
 * @param {string[]} availableLocations The list of valid location keys
 * @return {string|null} The resolved location name or null if not found
 */
function resolveLocationAlias(input, availableLocations) {
  const normalizedInput = input.trim().toUpperCase();
  
  // Define mapping for your specific shortcuts
  const aliasMap = {
    'CA': 'Cantina Añejo',
    'OAK': 'Original American Kitchen',
    'WB': 'White Buffalo'
  };

  // 1. Check if it's a known alias
  let resolved = aliasMap[normalizedInput] || input;

  // 2. Find case-insensitive match in the actual available locations
  const finalLocation = availableLocations.find(loc => 
    loc.toLowerCase() === resolved.toLowerCase()
  );

  return finalLocation || null;
}