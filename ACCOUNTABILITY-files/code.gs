// ============================================================================
// THREE POINTS HOSPITALITY — TRAINING ACCOUNTABILITY SYSTEM
// C.O.R.E. (Consistent, Operational, Repeatable, Excellence)
// ============================================================================
//
// This Google Apps Script powers the Trainer Accountability Form — a web app
// submitted by trainers at the end of every training shift. It captures:
//   1. Session info (location, trainer, trainee, position, day, shift)
//   2. Trainee performance scores (3 categories, named labels: Poor→Excellent)
//   3. Three accountability questions (coverage, gaps, plan forward)
//   4. End-of-shift recap confirmation (yes/no)
//   5. Photo upload of signed daily floor checklist
//
// Data flows into a Google Sheets backend with auto-generated analytics.
// Enforcement: No form + no photo = no trainer incentive for that shift.
//
// ============================================================================

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------

const CONFIG = {
  EMAIL: {
    ENABLED: true,
    RECIPIENTS: ['devin@threepointshospitality.com'],
    SUBJECT_PREFIX: '[TPH Training]',
    HOURLY_LIMIT: 10
  },
  LOCATIONS: ['Cantina Añejo', 'Original American Kitchen', 'White Buffalo'],
  LOCATION_ALIASES: {
    'CA': 'Cantina Añejo',
    'CANTINA': 'Cantina Añejo',
    'OAK': 'Original American Kitchen',
    'WB': 'White Buffalo'
  },
  SCORING: {
    MAX_PER_CATEGORY: 5,
    CATEGORIES: 3,
    MAX_TOTAL: 15,
    THRESHOLDS: { EXCELLENT: 90, GOOD: 75 },
    // Named label → numeric value
    LABEL_MAP: { 'Poor': 1, 'Developing': 2, 'Average': 3, 'Strong': 4, 'Excellent': 5 }
  },
  PHOTO_UPLOAD: {
    FOLDER_NAME: 'TPH Training Checklists',
    MAX_SIZE_MB: 10,
    ALLOWED_TYPES: ['image/jpeg', 'image/png', 'image/heic']
  },
  DUPLICATE_WINDOW_MS: 5 * 60 * 1000 // 5 minutes
};

// Trainer roster by location. Single source of truth — used by both
// the backend (validation, analytics) and the frontend (dropdown population).
const TRAINERS = {
  'Cantina Añejo': [
    'Adeniza Fenne', 'Axaielle Cazeau-Quinn', 'Christian Lucas',
    'Davia Geders', 'Ella Agustin', 'Emma Yang', 'Evan Amato',
    'Gabriella McMillan', 'Lilah Bowers', 'Lilly Denny',
    'Macy Williams', 'Selina Ayup', 'Shaylee Estes',
    'Suzy Takla', 'Valeria Cvjetkovic'
  ],
  'Original American Kitchen': [
    'Carson Fontana', 'Desiree Edwards', 'Emma Thomas',
    'Kai Nishikawa', 'Natalia Martinez', 'Rachel Donly',
    'Tanner Griffin', 'Val Revilla'
  ],
  'White Buffalo': [
    'Dani Mizrachi', 'Holden Fernandez'
  ]
};

// Derived count for analytics sheets
const TRAINER_COUNTS = {};
for (const loc of CONFIG.LOCATIONS) {
  TRAINER_COUNTS[loc] = TRAINERS[loc].length;
}

// Training Records column layout — A=0 through S=18 (19 columns, no Record ID)
const COLUMNS = {
  TIMESTAMP: 0,          // A
  LOCATION: 1,           // B
  TRAINER: 2,            // C
  TRAINEE: 3,            // D
  POSITION: 4,           // E
  TRAINING_DAY: 5,       // F
  SHIFT: 6,              // G
  PERFORMANCE_LEVEL: 7,  // H
  OVERALL_NOTES: 8,      // I
  PERFORMANCE_SCORE: 9,  // J
  KNOWLEDGE_SCORE: 10,   // K
  ATTITUDE_SCORE: 11,    // L
  TOTAL_SCORE: 12,       // M
  PERCENTAGE: 13,        // N
  WHAT_COVERED: 14,      // O
  WHERE_STRUGGLING: 15,  // P
  PLAN_FORWARD: 16,      // Q
  RECAP: 17,             // R
  PHOTO_URL: 18          // S
};

const HEADERS = [
  'Timestamp', 'Location', 'Trainer', 'Trainee', 'Position',
  'Training Day', 'Shift', 'Performance Level', 'Overall Notes',
  'Performance Score', 'Knowledge Score', 'Attitude Score',
  'Total Score', 'Percentage',
  'What Was Covered', 'Where Struggling', 'Plan for Next Shift',
  'Recap', 'Checklist Photo URL'
];


// ===========================================================================
// WEB APP ENTRY POINTS
// ===========================================================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Training Accountability — Three Points Hospitality')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Called by the frontend to populate trainer dropdowns from a single source. */
function getTrainerRoster() {
  return TRAINERS;
}


// ===========================================================================
// CUSTOM MENU
// ===========================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎯 Training System')
    .addItem('🔄 Refresh All Analytics', 'refreshAnalytics')
    .addSeparator()
    .addItem('📊 Rebuild Analytics Dashboard', 'rebuildAnalyticsDashboard')
    .addItem('📍 Rebuild Location Summary', 'rebuildLocationSummary')
    .addItem('👥 Rebuild Trainer Performance', 'rebuildTrainerPerformance')
    .addItem('📅 Rebuild Monthly Location Performance', 'rebuildMonthlyLocationPerformance')
    .addSeparator()
    .addItem('🎨 Apply Location Colors', 'applyLocationColorCoding')
    .addItem('💰 Format PAID VALIDATION', 'formatPaidValidationSheet')
    .addSeparator()
    .addItem('🔧 Run Sheet Migration (one-time)', 'runFullMigration')
    .addItem('⚙️ System Setup (First Time)', 'initializeSystem')
    .addToUi();
}


// ===========================================================================
// FORM SUBMISSION HANDLER
// ===========================================================================

function submitTrainingData(data) {
  console.log('=== FORM SUBMISSION RECEIVED ===');

  try {
    // --- Validate required fields ---
    const required = [
      'location', 'trainer', 'trainee', 'position', 'trainingDay',
      'performanceScore', 'knowledgeScore', 'attitudeScore',
      'whatCovered', 'whereStruggling', 'planForward', 'recap'
    ];
    for (const field of required) {
      if (!data[field] && data[field] !== 0 && data[field] !== false) {
        throw new Error('Missing required field: ' + field);
      }
    }

    // --- Convert named labels → numbers and calculate scores ---
    const labelToNum = function(label) {
      var n = CONFIG.SCORING.LABEL_MAP[label];
      if (n) return n;
      var parsed = parseInt(label);
      return (!isNaN(parsed) && parsed >= 1 && parsed <= 5) ? parsed : 0;
    };
    const scores = {
      performance: labelToNum(data.performanceScore),
      knowledge:   labelToNum(data.knowledgeScore),
      attitude:    labelToNum(data.attitudeScore)
    };
    const totalScore = scores.performance + scores.knowledge + scores.attitude;
    const percentage = Math.round((totalScore / CONFIG.SCORING.MAX_TOTAL) * 100);
    const performanceLevel = percentage >= CONFIG.SCORING.THRESHOLDS.EXCELLENT ? 'Excellent'
      : percentage >= CONFIG.SCORING.THRESHOLDS.GOOD ? 'Good'
      : 'Needs Improvement';

    // --- Duplicate check ---
    const hash = generateSubmissionHash(data);
    if (isDuplicateSubmission(hash)) {
      return { success: false, error: 'Duplicate submission detected. Please wait before resubmitting.' };
    }

    // --- Handle photo upload ---
    var photoUrl = '';
    if (data.photoData && data.photoFileName) {
      photoUrl = saveChecklistPhoto(data);
    }

    // --- Get spreadsheet and ensure sheets exist ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAllSheets(ss);

    // --- Insert record ---
    const recordId = insertRecord(ss, {
      location: data.location,
      trainer: data.trainer,
      trainee: data.trainee,
      position: data.position,
      trainingDay: data.trainingDay,
      shift: data.shift,
      performanceScore: data.performanceScore,
      knowledgeScore: data.knowledgeScore,
      attitudeScore: data.attitudeScore,
      totalScore: totalScore,
      percentage: percentage,
      performanceLevel: performanceLevel,
      whatCovered: data.whatCovered,
      whereStruggling: data.whereStruggling,
      planForward: data.planForward,
      recap: data.recap,
      overallNotes: data.overallNotes,
      photoUrl: photoUrl,
      hash: hash
    });

    // --- Send notifications ---
    if (CONFIG.EMAIL.ENABLED) {
      sendNotificationSafe(data, recordId, totalScore, percentage, performanceLevel);
    }

    // --- Check flag conditions ---
    checkAlertConditions(data, percentage, performanceLevel, recordId);

    console.log('✓ Submission complete: ' + recordId);
    return {
      success: true,
      recordId: recordId,
      message: 'Assessment submitted successfully.',
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.log('❌ Submission error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}


// ===========================================================================
// RECORD INSERTION
// ===========================================================================

function insertRecord(ss, data) {
  const sheet = ss.getSheetByName('Training Records');
  if (!sheet) throw new Error('Training Records sheet not found');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Double-check duplicate after acquiring lock
    if (isDuplicateSubmission(data.hash)) {
      throw new Error('Duplicate detected after lock acquisition');
    }

    const timestamp = new Date();

    const row = [
      timestamp,                         // A: Timestamp
      data.location,                     // B: Location
      data.trainer,                      // C: Trainer
      data.trainee,                      // D: Trainee
      data.position,                     // E: Position
      data.trainingDay,                  // F: Training Day
      data.shift || '',                  // G: Shift
      data.performanceLevel,             // H: Performance Level
      data.overallNotes || '',           // I: Overall Notes
      data.performanceScore,             // J: Performance Score
      data.knowledgeScore,               // K: Knowledge Score
      data.attitudeScore,                // L: Attitude Score
      data.totalScore,                   // M: Total Score
      data.percentage,                   // N: Percentage
      data.whatCovered || '',            // O: What Was Covered
      data.whereStruggling || '',        // P: Where Struggling
      data.planForward || '',            // Q: Plan for Next Shift
      data.recap || '',                  // R: Recap
      data.photoUrl || ''               // S: Checklist Photo URL
    ];

    sheet.appendRow(row);

    // Color-code the Performance Level cell
    const lastRow = sheet.getLastRow();
    const colors = { 'Excellent': '#d5f4e6', 'Good': '#fff3cd', 'Needs Improvement': '#f8d7da' };
    sheet.getRange(lastRow, COLUMNS.PERFORMANCE_LEVEL + 1)
      .setBackground(colors[data.performanceLevel] || '#ffffff');

    SpreadsheetApp.flush();

    // Return a generated reference for the success message
    const ref = 'TR-' +
      Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd') +
      '-' + Math.random().toString(36).substr(2, 6).toUpperCase();
    return ref;

  } finally {
    lock.releaseLock();
  }
}


// ===========================================================================
// PHOTO UPLOAD
// ===========================================================================

function saveChecklistPhoto(data) {
  try {
    // Decode the base64 photo data
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.photoData),
      data.photoMimeType || 'image/jpeg',
      data.photoFileName || 'checklist.jpg'
    );

    // Find or create the root folder
    var rootFolder;
    const folders = DriveApp.getFoldersByName(CONFIG.PHOTO_UPLOAD.FOLDER_NAME);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(CONFIG.PHOTO_UPLOAD.FOLDER_NAME);
    }

    // Organize: Location / Trainee / Day
    const locationFolder = getOrCreateSubfolder(rootFolder, data.location);
    const traineeFolder = getOrCreateSubfolder(locationFolder, data.trainee);

    // Name the file descriptively
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const fileName = 'Day' + data.trainingDay + '_' + data.trainee + '_' + dateStr + '.' + getExtension(data.photoFileName);

    const file = traineeFolder.createFile(blob.setName(fileName));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    console.log('✓ Photo saved: ' + file.getUrl());
    return file.getUrl();

  } catch (error) {
    console.log('⚠️ Photo upload failed: ' + error.toString());
    return ''; // Don't block submission — photo can be resubmitted
  }
}

function getOrCreateSubfolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parent.createFolder(name);
}

function getExtension(filename) {
  const parts = (filename || 'file.jpg').split('.');
  return parts[parts.length - 1] || 'jpg';
}


// ===========================================================================
// DUPLICATE DETECTION
// ===========================================================================

function generateSubmissionHash(data) {
  const raw = data.location + '-' + data.trainer + '-' + data.trainee + '-' + data.position + '-' + data.trainingDay;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw)
    .map(function(b) { return (b + 256).toString(16).slice(-2); })
    .join('');
}

function isDuplicateSubmission(hash) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Training Records');
    if (!sheet || sheet.getLastRow() <= 1) return false;

    const values = sheet.getDataRange().getValues();
    const cutoff = new Date(Date.now() - CONFIG.DUPLICATE_WINDOW_MS);

    // Check last 20 rows only (performance)
    const start = Math.max(1, values.length - 20);
    for (var i = start; i < values.length; i++) {
      const rowTime = new Date(values[i][COLUMNS.TIMESTAMP]);
      if (rowTime > cutoff) {
        const rowHash = generateSubmissionHash({
          location: values[i][COLUMNS.LOCATION],
          trainer: values[i][COLUMNS.TRAINER],
          trainee: values[i][COLUMNS.TRAINEE],
          position: values[i][COLUMNS.POSITION],
          trainingDay: values[i][COLUMNS.TRAINING_DAY]
        });
        if (rowHash === hash) return true;
      }
    }
    return false;

  } catch (e) {
    console.log('Duplicate check error: ' + e.toString());
    return false; // Fail open — allow submission
  }
}


// ===========================================================================
// ALERT / FLAG LOGIC
// ===========================================================================

function checkAlertConditions(data, percentage, performanceLevel, recordId) {
  const day = parseInt(data.trainingDay);
  const alerts = [];

  // Flag 1: Day 4 or 5 with score < 75%
  if ((day === 4 || day === 5) && percentage < CONFIG.SCORING.THRESHOLDS.GOOD) {
    alerts.push('⚠️ LOW SCORE on Day ' + day + ': ' + data.trainee + ' scored ' + percentage + '% (' + performanceLevel + '). ' +
      'Trainer: ' + data.trainer + ' at ' + data.location + '. May not be ready for mock service.');
  }

  // Flag 2: Recap not completed
  if (data.recap && data.recap.toString().startsWith('No')) {
    alerts.push('⚠️ MISSED RECAP: ' + data.trainer + ' did not complete End-of-Shift Recap with ' + data.trainee +
      ' (Day ' + day + ' at ' + data.location + '). ' + data.recap);
  }

  // Flag 3: Check for consecutive Needs Improvement (same trainee)
  if (performanceLevel === 'Needs Improvement') {
    const consecutive = checkConsecutiveNI(data.trainee, data.location);
    if (consecutive >= 2) {
      alerts.push('🔴 CONSECUTIVE NI: ' + data.trainee + ' has ' + consecutive + ' consecutive "Needs Improvement" ' +
        'scores at ' + data.location + '. Check if this is a trainee issue or training execution issue.');
    }
  }

  // Send alert email if any flags triggered
  if (alerts.length > 0) {
    try {
      const subject = CONFIG.EMAIL.SUBJECT_PREFIX + ' ⚠️ Training Alert — ' + data.trainee + ' (' + recordId + ')';
      const body = 'Training Alert Summary\n' +
        'Record: ' + recordId + '\n' +
        'Trainee: ' + data.trainee + ' | Trainer: ' + data.trainer + '\n' +
        'Location: ' + data.location + ' | Day ' + data.trainingDay + '\n' +
        'Score: ' + percentage + '% (' + performanceLevel + ')\n\n' +
        alerts.join('\n\n');

      CONFIG.EMAIL.RECIPIENTS.forEach(function(email) {
        MailApp.sendEmail(email, subject, body);
      });
      console.log('✓ Alert email sent: ' + alerts.length + ' flag(s)');
    } catch (e) {
      console.log('⚠️ Alert email failed: ' + e.toString());
    }
  }
}

function checkConsecutiveNI(trainee, location) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Training Records');
    if (!sheet || sheet.getLastRow() <= 1) return 0;

    const values = sheet.getDataRange().getValues();
    var consecutive = 0;

    // Walk backwards through records for this trainee at this location
    for (var i = values.length - 1; i >= 1; i--) {
      if (values[i][COLUMNS.TRAINEE] === trainee && values[i][COLUMNS.LOCATION] === location) {
        if (values[i][COLUMNS.PERFORMANCE_LEVEL] === 'Needs Improvement') {
          consecutive++;
        } else {
          break; // Streak broken
        }
      }
    }
    return consecutive;
  } catch (e) {
    return 0;
  }
}


// ===========================================================================
// EMAIL NOTIFICATIONS
// ===========================================================================

function sendNotificationSafe(data, recordId, totalScore, percentage, performanceLevel) {
  try {
    // Rate limiting via PropertiesService
    const hourKey = 'emailCount_' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHH');
    const count = parseInt(PropertiesService.getScriptProperties().getProperty(hourKey) || '0');

    if (count >= CONFIG.EMAIL.HOURLY_LIMIT) {
      console.log('Email rate limit reached, skipping');
      return;
    }

    const subject = CONFIG.EMAIL.SUBJECT_PREFIX + ' ' + performanceLevel + ' — ' + data.trainee + ' Day ' + data.trainingDay;
    const bodyParts = [
      'Training Assessment Submitted',
      '',
      'Record ID: ' + recordId,
      'Location: ' + data.location,
      'Trainer: ' + data.trainer,
      'Trainee: ' + data.trainee,
      'Position: ' + data.position,
      'Training Day: ' + data.trainingDay,
      'Shift: ' + (data.shift || 'Not specified'),
      '',
      '--- Scores ---',
      'Performance: ' + data.performanceScore,
      'Knowledge:   ' + data.knowledgeScore,
      'Attitude:    ' + data.attitudeScore,
      'Total: ' + totalScore + '/15 (' + percentage + '%)',
      'Performance Level: ' + performanceLevel,
      '',
      '--- Accountability ---',
      'What was covered: ' + (data.whatCovered || 'N/A'),
      'Where struggling: ' + (data.whereStruggling || 'N/A'),
      'Plan forward: ' + (data.planForward || 'N/A'),
      'Recap: ' + (data.recap || 'N/A')
    ];

    bodyParts.push('Checklist photo: ' + (data.photoUrl ? 'Uploaded' : 'Not uploaded'));
    bodyParts.push('');
    bodyParts.push('Overall notes: ' + (data.overallNotes || 'None'));

    const body = bodyParts.join('\n');

    CONFIG.EMAIL.RECIPIENTS.forEach(function(email) {
      MailApp.sendEmail(email, subject, body);
    });

    PropertiesService.getScriptProperties().setProperty(hourKey, (count + 1).toString());
    console.log('✓ Notification email sent');

  } catch (error) {
    console.log('⚠️ Email failed: ' + error.toString());
  }
}


// ===========================================================================
// SHEET INITIALIZATION
// ===========================================================================

function ensureAllSheets(ss) {
  if (!ss.getSheetByName('Training Records')) createTrainingRecordsSheet(ss);
  if (!ss.getSheetByName('Analytics Dashboard')) createAnalyticsSheet(ss);
  if (!ss.getSheetByName('Location Summary')) createLocationSummarySheet(ss);
  if (!ss.getSheetByName('Trainer Performance')) createTrainerPerformanceSheet(ss);
}

function createTrainingRecordsSheet(ss) {
  console.log('Creating Training Records sheet...');
  const sheet = ss.insertSheet('Training Records');

  // Headers
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setBackground('#2C5AA0').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);

  // Column widths — matches 19-col layout A–S
  var widths = [150, 150, 120, 120, 100, 80, 80, 130, 300, 80, 80, 80, 80, 80, 250, 250, 250, 200, 200];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  sheet.setFrozenRows(1);
  console.log('✓ Training Records sheet created');
  return sheet;
}

function createAnalyticsSheet(ss) {
  console.log('Creating Analytics Dashboard...');
  const sheet = ss.insertSheet('Analytics Dashboard');

  // Title
  sheet.getRange('A1').setValue('Three Points Hospitality — Training Analytics')
    .setFontSize(18).setFontWeight('bold').setBackground('#2C5AA0').setFontColor('#FFFFFF');
  sheet.getRange('A1:H1').merge();

  sheet.getRange('A3').setValue('Last Updated:').setFontWeight('bold');
  sheet.getRange('B3').setValue(new Date()).setNumberFormat('MM/dd/yyyy hh:mm:ss AM/PM');

  // Overall metrics
  sheet.getRange('A5').setValue('📊 OVERALL PERFORMANCE')
    .setFontSize(14).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
  sheet.getRange('A5:H5').merge();

  sheet.getRange('A6:E6').setValues([['Metric', 'This Week', 'Last Week', 'This Month', 'All Time']])
    .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');

  // Metrics rows (placeholders — populated by rebuildAnalyticsDashboard)
  var metrics = [
    ['Total Assessments', '', '', '', ''],
    ['Average Score (%)', '', '', '', ''],
    ['Excellent Rate (%)', '', '', '', ''],
    ['Good Rate (%)', '', '', '', ''],
    ['Needs Improvement (%)', '', '', '', ''],
    ['Unique Trainees', '', '', '', ''],
    ['Active Trainers', '', '', '', '']
  ];
  sheet.getRange(7, 1, metrics.length, 5).setValues(metrics);

  // Location section
  sheet.getRange('A15').setValue('🏢 PERFORMANCE BY LOCATION')
    .setFontSize(14).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
  sheet.getRange('A15:H15').merge();

  sheet.getRange('A16:E16').setValues([['Location', 'Total', 'Avg Score (%)', 'Excellence Rate (%)', 'NI Rate (%)']])
    .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');

  CONFIG.LOCATIONS.forEach(function(loc, i) {
    sheet.getRange(17 + i, 1).setValue(loc);
  });

  // Column widths
  [180, 130, 120, 130, 130, 150].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  console.log('✓ Analytics Dashboard created');
  return sheet;
}

function createLocationSummarySheet(ss) {
  console.log('Creating Location Summary...');
  const sheet = ss.insertSheet('Location Summary');

  sheet.getRange('A1').setValue('Location Performance Summary')
    .setFontSize(16).setFontWeight('bold').setBackground('#2C5AA0').setFontColor('#FFFFFF');
  sheet.getRange('A1:H1').merge();

  sheet.getRange('A3').setValue('Last Updated:').setFontWeight('bold');
  sheet.getRange('B3').setValue(new Date()).setNumberFormat('MM/dd/yyyy hh:mm:ss AM/PM');

  var row = 5;
  var thisMonthStart = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  var lastMonthStart = 'DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1)';
  var lastMonthEnd   = 'EOMONTH(TODAY(),-1)';

  CONFIG.LOCATIONS.forEach(function(location) {
    sheet.getRange(row, 1).setValue('📍 ' + location.toUpperCase())
      .setFontSize(12).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
    sheet.getRange(row, 1, 1, 6).merge();
    row++;

    sheet.getRange(row, 1, 1, 6).setValues([['Metric', 'All Time', 'This Month', 'Last Month', 'Trend', 'Notes']])
      .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
    row++;

    // All formulas use named ranges — never break when columns shift
    var locQ  = '"' + location + '"';
    var metricsFormulas = [
      ['Total Assessments',
        '=COUNTIF(tr_location,' + locQ + ')',
        '=COUNTIFS(tr_location,' + locQ + ',tr_timestamp,">="&' + thisMonthStart + ',tr_timestamp,"<="&TODAY())',
        '=COUNTIFS(tr_location,' + locQ + ',tr_timestamp,">="&' + lastMonthStart + ',tr_timestamp,"<="&' + lastMonthEnd + ')',
        '=IFERROR(B' + (row+0) + '-C' + (row+0) + ',"")', 'Count'],
      ['Average Score (%)',
        '=IFERROR(ROUND(AVERAGEIF(tr_location,' + locQ + ',tr_percentage),1),0)',
        '=IFERROR(ROUND(AVERAGEIFS(tr_percentage,tr_location,' + locQ + ',tr_timestamp,">="&' + thisMonthStart + ',tr_timestamp,"<="&TODAY()),1),0)',
        '=IFERROR(ROUND(AVERAGEIFS(tr_percentage,tr_location,' + locQ + ',tr_timestamp,">="&' + lastMonthStart + ',tr_timestamp,"<="&' + lastMonthEnd + '),1),0)',
        '=IFERROR(B' + (row+1) + '-C' + (row+1) + ',"")', 'Percentage column'],
      ['Excellence Rate (%)',
        '=IFERROR(ROUND(COUNTIFS(tr_location,' + locQ + ',tr_performance_level,"Excellent")/COUNTIF(tr_location,' + locQ + ')*100,1),0)',
        '=IFERROR(ROUND(COUNTIFS(tr_location,' + locQ + ',tr_performance_level,"Excellent",tr_timestamp,">="&' + thisMonthStart + ',tr_timestamp,"<="&TODAY())/COUNTIFS(tr_location,' + locQ + ',tr_timestamp,">="&' + thisMonthStart + ',tr_timestamp,"<="&TODAY())*100,1),0)',
        '=IFERROR(ROUND(COUNTIFS(tr_location,' + locQ + ',tr_performance_level,"Excellent",tr_timestamp,">="&' + lastMonthStart + ',tr_timestamp,"<="&' + lastMonthEnd + ')/COUNTIFS(tr_location,' + locQ + ',tr_timestamp,">="&' + lastMonthStart + ',tr_timestamp,"<="&' + lastMonthEnd + ')*100,1),0)',
        '=IFERROR(B' + (row+2) + '-C' + (row+2) + ',"")', '≥90% score'],
      ['Active Trainers',
        '=' + TRAINER_COUNTS[location],
        '=B' + (row+3),
        '=B' + (row+3),
        '', 'Roster count']
    ];

    sheet.getRange(row, 1, metricsFormulas.length, 6).setValues(metricsFormulas);
    row += metricsFormulas.length + 2;
  });

  [150, 120, 120, 120, 100, 150].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  console.log('✓ Location Summary created');
  return sheet;
}

function createTrainerPerformanceSheet(ss) {
  console.log('Creating Trainer Performance...');
  const sheet = ss.insertSheet('Trainer Performance');

  var colors = ['#FCE5CD', '#CFE2F3', '#D9D2E9'];
  var col = 1;

  CONFIG.LOCATIONS.forEach(function(location, i) {
    sheet.getRange(1, col).setValue(location)
      .setFontSize(12).setFontWeight('bold').setBackground(colors[i]);
    sheet.getRange(1, col, 1, 5).merge();

    sheet.getRange(2, col, 1, 5)
      .setValues([['Month', 'Trainer', 'Assessments', 'Avg Score (%)', 'Success Rate (%)']])
      .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');

    // Force number format to prevent date display bugs
    sheet.getRange(3, col + 2, 100, 3).setNumberFormat('0');

    [80, 150, 100, 110, 110].forEach(function(w, j) { sheet.setColumnWidth(col + j, w); });
    col += 6;
  });

  console.log('✓ Trainer Performance created');
  return sheet;
}


// ===========================================================================
// ANALYTICS POPULATION (single-pass from Training Records)
// ===========================================================================

/**
 * Reads all Training Records once, then populates all analytics sheets.
 * Called via the menu or after rebuilding individual sheets.
 */
function refreshAnalytics() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const records = getTrainingRecords(ss);

    populateAnalyticsDashboard(ss, records);
    populateTrainerPerformance(ss, records);
    populateMonthlyLocationPerformance(ss, records);

    // Update Location Summary timestamp
    const ls = ss.getSheetByName('Location Summary');
    if (ls) ls.getRange('B3').setValue(new Date());

    SpreadsheetApp.flush();
    ui.alert('Success', '✓ All analytics refreshed from ' + records.length + ' records.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Analytics refresh failed: ' + e.toString(), ui.ButtonSet.OK);
  }
}

/** Read all Training Records (skip header). Returns array of row arrays. */
function getTrainingRecords(ss) {
  const sheet = ss.getSheetByName('Training Records');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getDataRange().getValues().slice(1);
}


// ---------------------------------------------------------------------------
// Analytics Dashboard
// ---------------------------------------------------------------------------

function rebuildAnalyticsDashboard() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = ss.getSheetByName('Analytics Dashboard');
    if (existing) ss.deleteSheet(existing);
    createAnalyticsSheet(ss);
    populateAnalyticsDashboard(ss, getTrainingRecords(ss));
    ui.alert('Success', '✓ Analytics Dashboard rebuilt.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.toString(), ui.ButtonSet.OK);
  }
}

function populateAnalyticsDashboard(ss, records) {
  const sheet = ss.getSheetByName('Analytics Dashboard');
  if (!sheet) return;

  const now = new Date();
  const weekStart = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
  const lastWeekStart = new Date(weekStart.getTime() - 7 * 86400000);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  var buckets = {
    thisWeek: newBucket(), lastWeek: newBucket(), thisMonth: newBucket(), allTime: newBucket(),
    locations: {}
  };

  records.forEach(function(r) {
    if (!r[COLUMNS.TIMESTAMP] || !r[COLUMNS.LOCATION]) return;
    const ts = new Date(r[COLUMNS.TIMESTAMP]);
    const pct = parseFloat(r[COLUMNS.PERCENTAGE]) || 0;
    // Derive level from pct if the stored value is blank (handles un-backfilled rows)
    const storedLevel = (r[COLUMNS.PERFORMANCE_LEVEL] || '').toString().trim();
    const level = storedLevel || (pct >= CONFIG.SCORING.THRESHOLDS.EXCELLENT ? 'Excellent'
                                : pct >= CONFIG.SCORING.THRESHOLDS.GOOD      ? 'Good'
                                : pct > 0                                     ? 'Needs Improvement' : '');
    const trainer = r[COLUMNS.TRAINER];
    const trainee = r[COLUMNS.TRAINEE];
    const loc = r[COLUMNS.LOCATION];

    if (isNaN(ts.getTime())) return;

    // Location bucket
    if (!buckets.locations[loc]) buckets.locations[loc] = newBucket();
    addToBucket(buckets.locations[loc], pct, level, trainer, trainee);

    // Time buckets
    if (ts >= weekStart) addToBucket(buckets.thisWeek, pct, level, trainer, trainee);
    if (ts >= lastWeekStart && ts < weekStart) addToBucket(buckets.lastWeek, pct, level, trainer, trainee);
    if (ts >= monthStart) addToBucket(buckets.thisMonth, pct, level, trainer, trainee);
    addToBucket(buckets.allTime, pct, level, trainer, trainee);
  });

  sheet.getRange('B3').setValue(now);

  // Populate overview (rows 7-13)
  var periods = [buckets.thisWeek, buckets.lastWeek, buckets.thisMonth, buckets.allTime];
  var overviewData = [
    ['Total Assessments'].concat(periods.map(function(p) { return p.total; })),
    ['Average Score (%)'].concat(periods.map(function(p) { return p.total > 0 ? Math.round(p.scoreSum / p.total) : 0; })),
    ['Excellent Rate (%)'].concat(periods.map(function(p) { return p.total > 0 ? Math.round(p.excellent / p.total * 100) : 0; })),
    ['Good Rate (%)'].concat(periods.map(function(p) { return p.total > 0 ? Math.round(p.good / p.total * 100) : 0; })),
    ['Needs Improvement (%)'].concat(periods.map(function(p) { return p.total > 0 ? Math.round(p.ni / p.total * 100) : 0; })),
    ['Unique Trainees'].concat(periods.map(function(p) { return p.trainees.size; })),
    ['Active Trainers'].concat(periods.map(function(p) { return p.trainers.size; }))
  ];
  sheet.getRange(7, 1, overviewData.length, 5).setValues(overviewData);

  // Populate location rows (rows 17-19)
  CONFIG.LOCATIONS.forEach(function(loc, i) {
    var b = buckets.locations[loc] || newBucket();
    var row = 17 + i;
    sheet.getRange(row, 1).setValue(loc);
    sheet.getRange(row, 2).setValue(b.total);
    sheet.getRange(row, 3).setValue(b.total > 0 ? Math.round(b.scoreSum / b.total) : 0);
    sheet.getRange(row, 4).setValue(b.total > 0 ? Math.round(b.excellent / b.total * 100) : 0);
    sheet.getRange(row, 5).setValue(b.total > 0 ? Math.round(b.ni / b.total * 100) : 0);
  });
}

function newBucket() {
  return { total: 0, scoreSum: 0, excellent: 0, good: 0, ni: 0, trainers: new Set(), trainees: new Set() };
}

function addToBucket(bucket, pct, level, trainer, trainee) {
  bucket.total++;
  bucket.scoreSum += pct;
  if (level === 'Excellent') bucket.excellent++;
  if (level === 'Good') bucket.good++;
  if (level === 'Needs Improvement') bucket.ni++;
  if (trainer) bucket.trainers.add(trainer);
  if (trainee) bucket.trainees.add(trainee);
}


// ---------------------------------------------------------------------------
// Trainer Performance
// ---------------------------------------------------------------------------

function rebuildTrainerPerformance() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = ss.getSheetByName('Trainer Performance');
    if (existing) ss.deleteSheet(existing);
    createTrainerPerformanceSheet(ss);
    populateTrainerPerformance(ss, getTrainingRecords(ss));
    ui.alert('Success', '✓ Trainer Performance rebuilt.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.toString(), ui.ButtonSet.OK);
  }
}

function populateTrainerPerformance(ss, records) {
  const sheet = ss.getSheetByName('Trainer Performance');
  if (!sheet) return;

  // Clear data rows
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clear();

  // Group by location → month → trainer
  var grouped = {};
  records.forEach(function(r) {
    if (!r[COLUMNS.TRAINER] || !r[COLUMNS.LOCATION] || !r[COLUMNS.TIMESTAMP]) return;
    const loc = r[COLUMNS.LOCATION];
    const trainer = r[COLUMNS.TRAINER];
    const ts = new Date(r[COLUMNS.TIMESTAMP]);
    const pct = parseFloat(r[COLUMNS.PERCENTAGE]) || 0;
    if (isNaN(ts.getTime())) return;

    const month = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MMM yyyy');
    const key = loc + '|' + trainer + '|' + month;

    if (!grouped[key]) grouped[key] = { loc: loc, trainer: trainer, month: month, count: 0, scoreSum: 0, successes: 0 };
    grouped[key].count++;
    grouped[key].scoreSum += pct;
    if (pct >= CONFIG.SCORING.THRESHOLDS.GOOD) grouped[key].successes++;
  });

  // Sort and write by location section
  var colStarts = { 'Cantina Añejo': 1, 'Original American Kitchen': 7, 'White Buffalo': 13 };
  var sorted = Object.values(grouped).sort(function(a, b) {
    if (a.loc !== b.loc) return a.loc.localeCompare(b.loc);
    return new Date(a.month).getTime() - new Date(b.month).getTime();
  });

  Object.keys(colStarts).forEach(function(loc) {
    var col = colStarts[loc];
    var row = 3;
    var currentMonth = '';

    sorted.filter(function(d) { return d.loc === loc; }).forEach(function(d) {
      if (d.month !== currentMonth) {
        sheet.getRange(row, col).setValue(d.month).setFontWeight('bold').setBackground('#E8F4FD');
        sheet.getRange(row, col + 1, 1, 4).clearContent();
        currentMonth = d.month;
        row++;
      }
      sheet.getRange(row, col).setValue('');
      sheet.getRange(row, col + 1).setValue(d.trainer);
      sheet.getRange(row, col + 2).setValue(d.count).setNumberFormat('0');
      sheet.getRange(row, col + 3).setValue(Math.round(d.scoreSum / d.count)).setNumberFormat('0');
      sheet.getRange(row, col + 4).setValue(Math.round(d.successes / d.count * 100)).setNumberFormat('0');
      row++;
    });
  });

  SpreadsheetApp.flush();
}


// ---------------------------------------------------------------------------
// Monthly Location Performance
// ---------------------------------------------------------------------------

function rebuildMonthlyLocationPerformance() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = ss.getSheetByName('Monthly Location Performance');
    if (existing) ss.deleteSheet(existing);
    createMonthlyLocationPerformanceSheet(ss);
    populateMonthlyLocationPerformance(ss, getTrainingRecords(ss));
    ui.alert('Success', '✓ Monthly Location Performance rebuilt.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.toString(), ui.ButtonSet.OK);
  }
}

function createMonthlyLocationPerformanceSheet(ss) {
  const sheet = ss.insertSheet('Monthly Location Performance');
  sheet.getRange('A1').setValue('Monthly Location Performance Analysis')
    .setFontSize(16).setFontWeight('bold').setBackground('#2C5AA0').setFontColor('#FFFFFF');
  sheet.getRange('A1:G1').merge();
  [120, 140, 130, 120, 100, 160, 120].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  return sheet;
}

function populateMonthlyLocationPerformance(ss, records) {
  const sheet = ss.getSheetByName('Monthly Location Performance');
  if (!sheet) return;

  // Clear below title
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();

  // Group by location → month
  var grouped = {};
  records.forEach(function(r) {
    if (!r[COLUMNS.LOCATION] || !r[COLUMNS.TIMESTAMP]) return;
    const loc = r[COLUMNS.LOCATION];
    const ts = new Date(r[COLUMNS.TIMESTAMP]);
    const pct = parseFloat(r[COLUMNS.PERCENTAGE]) || 0;
    const storedLevel = (r[COLUMNS.PERFORMANCE_LEVEL] || '').toString().trim();
    const level = storedLevel || (pct >= CONFIG.SCORING.THRESHOLDS.EXCELLENT ? 'Excellent'
                                : pct >= CONFIG.SCORING.THRESHOLDS.GOOD      ? 'Good'
                                : pct > 0                                     ? 'Needs Improvement' : '');
    if (isNaN(ts.getTime())) return;

    const month = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MMM yyyy');
    const key = loc + '|' + month;

    if (!grouped[key]) grouped[key] = { loc: loc, month: month, total: 0, scoreSum: 0, excellent: 0, good: 0, ni: 0, success: 0 };
    grouped[key].total++;
    grouped[key].scoreSum += pct;
    if (level === 'Excellent') grouped[key].excellent++;
    if (level === 'Good') grouped[key].good++;
    if (level === 'Needs Improvement') grouped[key].ni++;
    if (pct >= 80) grouped[key].success++;
  });

  var sorted = Object.values(grouped).sort(function(a, b) {
    if (a.loc !== b.loc) return a.loc.localeCompare(b.loc);
    return new Date(a.month).getTime() - new Date(b.month).getTime();
  });

  var colors = { 'Cantina Añejo': '#FCE5CD', 'Original American Kitchen': '#CFE2F3', 'White Buffalo': '#D9D2E9' };
  var row = 2;

  CONFIG.LOCATIONS.forEach(function(loc) {
    // Location header
    sheet.getRange(row, 1).setValue(loc)
      .setFontSize(12).setFontWeight('bold');
    sheet.getRange(row, 1, 1, 7).setBackground(colors[loc]);
    row++;

    // Column headers
    sheet.getRange(row, 1, 1, 7)
      .setValues([['Month', 'Total', 'Avg Score', 'Excellent', 'Good', 'Needs Improvement', 'Success Rate (%)']])
      .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
    row++;

    sorted.filter(function(d) { return d.loc === loc; }).forEach(function(d) {
      sheet.getRange(row, 1, 1, 7).setValues([[
        d.month,
        d.total,
        d.total > 0 ? Math.round(d.scoreSum / d.total) : 0,
        d.excellent,
        d.good,
        d.ni,
        d.total > 0 ? Math.round(d.success / d.total * 100) : 0
      ]]);
      row++;
    });
    row += 2;
  });

  SpreadsheetApp.flush();
}


// ---------------------------------------------------------------------------
// Location Summary rebuild
// ---------------------------------------------------------------------------

function rebuildLocationSummary() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = ss.getSheetByName('Location Summary');
    if (existing) ss.deleteSheet(existing);
    createLocationSummarySheet(ss);
    ui.alert('Success', '✓ Location Summary rebuilt with fresh formulas.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.toString(), ui.ButtonSet.OK);
  }
}


// ===========================================================================
// PAID VALIDATION & COLORS
// ===========================================================================

function formatPaidValidationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PAID VALIDATION');
  if (!sheet) sheet = ss.insertSheet('PAID VALIDATION');

  var maxRow = 100;
  sheet.getRange(1, 1, maxRow, 12).clearFormat();

  var sections = [1, 5, 9];
  var colors = ['#FCE5CD', '#CFE2F3', '#D9D2E9'];

  sections.forEach(function(col, i) {
    sheet.getRange(1, col, 1, 3).merge()
      .setBackground(colors[i]).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(2, col, 1, 3)
      .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(3, col, maxRow, 1).setNumberFormat('MMM yyyy');
    sheet.getRange(3, col + 1, maxRow, 1).setNumberFormat('@');
    sheet.getRange(3, col + 2, maxRow, 1).insertCheckboxes().setHorizontalAlignment('center');
  });

  [100, 150, 60, 30, 100, 150, 60, 30, 100, 150, 60].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setFrozenRows(2);

  SpreadsheetApp.getUi().alert('✓ PAID VALIDATION formatted.');
}

function applyLocationColorCoding() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    var colors = { 'Cantina Añejo': '#FCE5CD', 'Original American Kitchen': '#CFE2F3', 'White Buffalo': '#D9D2E9' };

    // Trainer Performance
    const tp = ss.getSheetByName('Trainer Performance');
    if (tp) {
      var col = 1;
      CONFIG.LOCATIONS.forEach(function(loc) {
        tp.getRange(1, col).setBackground(colors[loc]).setFontColor('#000000');
        col += 6;
      });
    }

    // Monthly Location Performance
    const mlp = ss.getSheetByName('Monthly Location Performance');
    if (mlp) {
      var data = mlp.getDataRange().getValues();
      data.forEach(function(row, i) {
        if (colors[row[0]]) mlp.getRange(i + 1, 1, 1, 7).setBackground(colors[row[0]]);
      });
    }

    ui.alert('Success', '✓ Location colors applied.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.toString(), ui.ButtonSet.OK);
  }
}


// ===========================================================================
// SYSTEM SETUP
// ===========================================================================

function initializeSystem() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAllSheets(ss);

    // Remove default Sheet1 if other sheets exist
    const sheet1 = ss.getSheetByName('Sheet1');
    if (sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);

    // Populate analytics from any existing data
    const records = getTrainingRecords(ss);
    if (records.length > 0) {
      populateAnalyticsDashboard(ss, records);
      populateTrainerPerformance(ss, records);
      populateMonthlyLocationPerformance(ss, records);
    }

    console.log('✓ System initialized');
    ui.alert('Success', '✓ All sheets created and analytics populated. System ready.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Setup failed: ' + e.toString(), ui.ButtonSet.OK);
  }
}


// ===========================================================================
// UTILITIES
// ===========================================================================

function resolveLocationAlias(input, availableLocations) {
  var normalized = (input || '').trim().toUpperCase();
  var resolved = CONFIG.LOCATION_ALIASES[normalized] || input;
  var locs = availableLocations || CONFIG.LOCATIONS;
  for (var i = 0; i < locs.length; i++) {
    if (locs[i].toLowerCase() === resolved.toLowerCase()) return locs[i];
  }
  return null;
}
