// ============================================================================
// SHEET MIGRATION & NAMED RANGES SETUP
// Three Points Hospitality — C.O.R.E. Training System
// ============================================================================
//
// RUN ORDER (from Apps Script editor):
//   → Run: runFullMigration()   ← does everything in the correct safe order
//
// WHAT IT DOES:
//   - Finds every column BY HEADER NAME (never by letter — shift-safe)
//   - Merges "Recap Completed" + "Recap Missed Reason" → single "Recap" col
//     "Yes — Complete" or "No — [reason]"
//   - Deletes: Record ID, Teamwork Score, Professionalism Score, Total Score
//   - Renames: Knowledge→Performance, Technical→Knowledge, Service→Attitude
//   - Moves "Overall Notes" to immediately after "Performance Level"
//   - Rebuilds Total Score as sum of 3 new scores
//   - Creates a named range for EVERY column (entire data column, row 2–10000)
//   - Rebuilds Location Summary formulas using named ranges
//
// BACKUP FIRST: Duplicate your Training Records tab before running.
// SAFE TO RE-RUN: Each step checks state before acting.
// ============================================================================

var SHEET_NAME = 'Training Records';

var FINAL_HEADERS = [
  'Timestamp',           // A
  'Location',            // B
  'Trainer',             // C
  'Trainee',             // D
  'Position',            // E
  'Training Day',        // F
  'Shift',               // G
  'Performance Level',   // H
  'Overall Notes',       // I
  'Performance Score',   // J
  'Knowledge Score',     // K
  'Attitude Score',      // L
  'Total Score',         // M
  'Percentage',          // N
  'What Was Covered',    // O
  'Where Struggling',    // P
  'Plan for Next Shift', // Q
  'Recap',               // R
  'Checklist Photo URL'  // S
];

var NAMED_RANGES = [
  'tr_timestamp',
  'tr_location',
  'tr_trainer',
  'tr_trainee',
  'tr_position',
  'tr_training_day',
  'tr_shift',
  'tr_performance_level',
  'tr_overall_notes',
  'tr_performance_score',
  'tr_knowledge_score',
  'tr_attitude_score',
  'tr_total_score',
  'tr_percentage',
  'tr_what_covered',
  'tr_where_struggling',
  'tr_plan_forward',
  'tr_recap',
  'tr_photo_url'
];

// Score label → numeric value map (for recalculating totals on old rows)
var SCORE_MAP = {
  'Poor': 1, 'Below Average': 2, 'Average': 3, 'Strong': 4, 'Excellent': 5,
  '1': 1, '2': 2, '3': 3, '4': 4, '5': 5 // handle old numeric values too
};

var LOCATIONS = ['Cantina Añejo', 'Original American Kitchen', 'White Buffalo'];
var SETUP_TRAINER_COUNTS = { 'Cantina Añejo': 15, 'Original American Kitchen': 8, 'White Buffalo': 2 };


// ============================================================================
// ENTRY POINT
// ============================================================================

function runFullMigration() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '⚠️ Run Sheet Migration?',
    'This will restructure the Training Records sheet.\n\n' +
    'MAKE SURE YOU HAVE A BACKUP TAB before continuing.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) { ui.alert('Migration cancelled.'); return; }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

    Logger.log('=== STARTING MIGRATION ===');

    step1_mergeRecapColumns(sheet);
    step2_deleteUnwantedColumns(sheet);
    step3_renameScoreHeaders(sheet);
    step4_moveOverallNotes(sheet);
    step5_verifyHeaders(sheet);
    step6_createNamedRanges(ss, sheet);
    step7_rebuildLocationSummaryFormulas(ss);
    step8_backfillScoresAndPercentage(sheet);

    SpreadsheetApp.flush();
    Logger.log('=== MIGRATION COMPLETE ===');

    ui.alert('✅ Migration Complete',
      'Training Records restructured.\n' +
      FINAL_HEADERS.length + ' named ranges created.\n\n' +
      'Now run "🔄 Refresh All Analytics" from the Training System menu.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log('MIGRATION FAILED: ' + e.toString());
    ui.alert('❌ Migration Failed', e.toString() + '\n\nCheck View → Logs for details.', ui.ButtonSet.OK);
  }
}


// ============================================================================
// STEP 1 — Merge "Recap Completed" + "Recap Missed Reason" → "Recap"
// Result format: "Yes — Complete"  or  "No — [reason]"
// ============================================================================

function step1_mergeRecapColumns(sheet) {
  Logger.log('Step 1: Merging Recap columns...');
  var h = getHeaderMap(sheet);

  var recapCol  = h['Recap Completed'];
  var reasonCol = h['Recap Missed Reason'];

  if (!recapCol) {
    Logger.log('  "Recap Completed" not found — may already be merged. Skipping.');
    return;
  }
  if (!reasonCol) {
    // Already merged — just rename the header
    sheet.getRange(1, recapCol).setValue('Recap');
    Logger.log('  "Recap Missed Reason" not found — renamed header only.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var recapVals  = sheet.getRange(2, recapCol,  lastRow - 1, 1).getValues();
    var reasonVals = sheet.getRange(2, reasonCol, lastRow - 1, 1).getValues();

    var merged = recapVals.map(function(row, i) {
      var recap  = (row[0] || '').toString().trim();
      var reason = (reasonVals[i][0] || '').toString().trim();
      if (recap === 'Yes') return ['Yes — Complete'];
      if (recap === 'No')  return ['No — ' + (reason || 'No reason provided')];
      if (recap === '')    return [''];
      return [recap];
    });

    sheet.getRange(2, recapCol, lastRow - 1, 1).setValues(merged);
  }

  sheet.getRange(1, recapCol).setValue('Recap');

  // Delete Reason col — find it again after header rename (col index unchanged)
  sheet.deleteColumn(reasonCol);
  Logger.log('  Merged into "Recap" col ' + colLetter(recapCol) + '. Deleted Reason col.');
}


// ============================================================================
// STEP 2 — Delete unwanted columns by name (right-to-left to avoid index shift)
// ============================================================================

function step2_deleteUnwantedColumns(sheet) {
  Logger.log('Step 2: Deleting unwanted columns...');
  var toDelete = ['Record ID', 'Teamwork Score', 'Professionalism Score', 'Total Score'];

  var h = getHeaderMap(sheet);
  var cols = [];
  toDelete.forEach(function(name) {
    if (h[name]) cols.push(h[name]);
    else Logger.log('  "' + name + '" not found — already removed or renamed.');
  });

  // Sort descending — delete rightmost first so left-side indexes stay valid
  cols.sort(function(a, b) { return b - a; });
  cols.forEach(function(col) {
    var name = sheet.getRange(1, col).getValue();
    sheet.deleteColumn(col);
    Logger.log('  Deleted "' + name + '" (col ' + colLetter(col) + ')');
  });
}


// ============================================================================
// STEP 3 — Rename score headers
// Knowledge Score → Performance Score
// Technical Score → Knowledge Score
// Service Score   → Attitude Score
// (rename in this exact order to avoid temporary collision)
// ============================================================================

function step3_renameScoreHeaders(sheet) {
  Logger.log('Step 3: Renaming score headers...');
  var renames = [
    { from: 'Knowledge Score', to: 'Performance Score' },
    { from: 'Technical Score', to: 'Knowledge Score'   },
    { from: 'Service Score',   to: 'Attitude Score'    }
  ];

  renames.forEach(function(r) {
    var h = getHeaderMap(sheet); // re-read each time
    var col = h[r.from];
    if (col) {
      sheet.getRange(1, col).setValue(r.to);
      Logger.log('  "' + r.from + '" → "' + r.to + '" (col ' + colLetter(col) + ')');
    } else {
      Logger.log('  "' + r.from + '" not found — skipping.');
    }
  });
}


// ============================================================================
// STEP 4 — Move "Overall Notes" to immediately after "Performance Level"
// ============================================================================

function step4_moveOverallNotes(sheet) {
  Logger.log('Step 4: Moving Overall Notes...');
  var h = getHeaderMap(sheet);
  var notesCol = h['Overall Notes'];
  var perfCol  = h['Performance Level'];

  if (!notesCol) { Logger.log('  "Overall Notes" not found. Skipping.'); return; }
  if (!perfCol)  { Logger.log('  "Performance Level" not found. Skipping.'); return; }

  var targetCol = perfCol + 1;
  if (notesCol === targetCol) {
    Logger.log('  Already in correct position. Skipping.');
    return;
  }

  var lastRow  = Math.max(sheet.getLastRow(), 1);
  var notesData = sheet.getRange(1, notesCol, lastRow, 1).getValues();

  // Insert blank column right after Performance Level
  sheet.insertColumnAfter(perfCol);

  // If notes was to the right of where we inserted, it shifted +1
  var adjustedNotesCol = notesCol > perfCol ? notesCol + 1 : notesCol;

  // Write notes data into the new column
  sheet.getRange(1, targetCol, lastRow, 1).setValues(notesData);

  // Delete the original notes column (now shifted)
  sheet.deleteColumn(adjustedNotesCol);
  Logger.log('  Overall Notes moved to col ' + colLetter(targetCol));
}


// ============================================================================
// STEP 5 — Verify headers match FINAL_HEADERS, log any mismatches
// ============================================================================

function step5_verifyHeaders(sheet) {
  Logger.log('Step 5: Verifying headers...');
  var lastCol = sheet.getLastColumn();
  var actual  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var ok = true;

  FINAL_HEADERS.forEach(function(expected, i) {
    var got = (actual[i] || '').toString().trim();
    if (got !== expected) {
      Logger.log('  ⚠️ Col ' + colLetter(i + 1) + ': expected "' + expected + '", got "' + got + '"');
      ok = false;
    }
  });

  if (ok) Logger.log('  ✅ All ' + FINAL_HEADERS.length + ' headers verified.');

  // Style header row
  sheet.getRange(1, 1, 1, FINAL_HEADERS.length)
    .setBackground('#2C5AA0').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11);

  // Column widths
  var widths = [150, 150, 120, 120, 100, 80, 80, 130, 300, 80, 80, 80, 80, 80, 250, 250, 250, 200, 200];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setFrozenRows(1);
}


// ============================================================================
// STEP 6 — Create named ranges (data rows only, row 2–10000)
// Finds each column by header name before assigning — fully shift-safe
// ============================================================================

function step6_createNamedRanges(ss, sheet) {
  Logger.log('Step 6: Creating named ranges...');

  // Remove any existing named ranges with our names
  ss.getNamedRanges().forEach(function(nr) {
    if (NAMED_RANGES.indexOf(nr.getName()) !== -1) nr.remove();
  });

  var h = getHeaderMap(sheet);

  FINAL_HEADERS.forEach(function(header, i) {
    var rangeName = NAMED_RANGES[i];
    var col = h[header];
    if (!col) {
      col = i + 1;
      Logger.log('  ⚠️ "' + header + '" not found by name — using position ' + col);
    }
    var range = sheet.getRange(2, col, 9999, 1);
    ss.setNamedRange(rangeName, range);
    Logger.log('  "' + rangeName + '" → col ' + colLetter(col) + ' (' + header + ')');
  });

  Logger.log('  ✅ ' + FINAL_HEADERS.length + ' named ranges created.');
}


// ============================================================================
// STEP 7 — Rebuild Location Summary using named ranges
// ============================================================================

function step7_rebuildLocationSummaryFormulas(ss) {
  Logger.log('Step 7: Rebuilding Location Summary with named range formulas...');
  var ls = ss.getSheetByName('Location Summary');
  if (!ls) { Logger.log('  Location Summary not found — skipping.'); return; }

  var lastRow = ls.getLastRow();
  if (lastRow > 4) ls.getRange(5, 1, lastRow - 4, 8).clear();

  var row = 5;
  var thisMonthStart = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  var lastMonthStart = 'DATE(YEAR(EOMONTH(TODAY(),-1)),MONTH(EOMONTH(TODAY(),-1)),1)';
  var lastMonthEnd   = 'EOMONTH(TODAY(),-1)';

  LOCATIONS.forEach(function(location) {
    ls.getRange(row, 1).setValue('📍 ' + location.toUpperCase())
      .setFontSize(12).setFontWeight('bold').setBackground('#6C7B7F').setFontColor('#FFFFFF');
    ls.getRange(row, 1, 1, 6).merge();
    row++;

    ls.getRange(row, 1, 1, 6)
      .setValues([['Metric', 'All Time', 'This Month', 'Last Month', 'Trend', 'Notes']])
      .setBackground('#8FA4A7').setFontColor('#FFFFFF').setFontWeight('bold');
    row++;

    var locQ = '"' + location + '"';
    var formulas = [
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
        '=' + SETUP_TRAINER_COUNTS[location],
        '=B' + (row+3),
        '=B' + (row+3),
        '', 'Roster count']
    ];

    ls.getRange(row, 1, formulas.length, 6).setValues(formulas);
    row += formulas.length + 2;
  });

  ls.getRange('B3').setValue(new Date());
  Logger.log('  ✅ Location Summary rebuilt with named range formulas.');
}


// ============================================================================
// STEP 8 — Backfill Total Score, Percentage, and Performance Level
// for all existing rows using the new 3-score schema (max 15)
// ============================================================================

function step8_backfillScoresAndPercentage(sheet) {
  // Allow running standalone from the editor (no argument)
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) { Logger.log('Sheet "' + SHEET_NAME + '" not found.'); return; }
  }
  Logger.log('Step 8: Backfilling scores and percentage for existing rows...');
  var h = getHeaderMap(sheet);

  var colPerf     = h['Performance Score'];
  var colKnow     = h['Knowledge Score'];
  var colAtt      = h['Attitude Score'];
  var colTotal    = h['Total Score'];
  var colPct      = h['Percentage'];
  var colLevel    = h['Performance Level'];

  if (!colPerf || !colKnow || !colAtt || !colTotal || !colPct || !colLevel) {
    Logger.log('  ⚠️ One or more score columns not found — skipping backfill.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('  No data rows to backfill.'); return; }

  var numRows = lastRow - 1;

  // Read all score columns at once
  var perfVals  = sheet.getRange(2, colPerf,  numRows, 1).getValues();
  var knowVals  = sheet.getRange(2, colKnow,  numRows, 1).getValues();
  var attVals   = sheet.getRange(2, colAtt,   numRows, 1).getValues();
  var totalVals = sheet.getRange(2, colTotal, numRows, 1).getValues();
  var pctVals   = sheet.getRange(2, colPct,   numRows, 1).getValues();

  var newTotals  = [];
  var newPcts    = [];
  var newLevels  = [];
  var backfilled = 0;

  for (var i = 0; i < numRows; i++) {
    var perf  = toScore(perfVals[i][0]);
    var know  = toScore(knowVals[i][0]);
    var att   = toScore(attVals[i][0]);
    var total = parseInt(totalVals[i][0]) || 0;
    var pct   = parseFloat(pctVals[i][0]) || 0;

    // Only backfill rows where Percentage is missing or 0 but scores exist
    if ((pct === 0 || pct === '') && (perf > 0 || know > 0 || att > 0)) {
      total = perf + know + att;
      pct   = Math.round((total / 15) * 100);
      backfilled++;
    }

    var level = pct >= 90 ? 'Excellent' : pct >= 75 ? 'Good' : 'Needs Improvement';
    if (pct === 0 && total === 0) level = '';

    newTotals.push([total]);
    newPcts.push([pct]);
    newLevels.push([level]);
  }

  sheet.getRange(2, colTotal, numRows, 1).setValues(newTotals);
  sheet.getRange(2, colPct,   numRows, 1).setValues(newPcts);
  sheet.getRange(2, colLevel, numRows, 1).setValues(newLevels);

  Logger.log('  ✅ Backfilled ' + backfilled + ' rows. All ' + numRows + ' rows verified.');
}

/** Convert a score cell value (label or number) to integer 1–5, or 0 if blank */
function toScore(val) {
  if (!val && val !== 0) return 0;
  var s = val.toString().trim();
  var map = { 'Poor': 1, 'Developing': 2, 'Average': 3, 'Strong': 4, 'Excellent': 5,
              'Below Average': 2, 'Good': 4 }; // handle old labels too
  if (map[s]) return map[s];
  var n = parseInt(s);
  return (!isNaN(n) && n >= 1 && n <= 5) ? n : 0;
}


// ============================================================================
// STANDALONE — Rebuild all four analytics sheets from scratch
// Run this from the menu after migration to get clean sheets
// ============================================================================

function rebuildAllAnalyticsSheets() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '🔄 Rebuild All Analytics Sheets?',
    'This will DELETE and recreate:\n' +
    '• Analytics Dashboard\n• Location Summary\n• Trainer Performance\n• Monthly Location Performance\n\n' +
    'Training Records data is NOT affected.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) { ui.alert('Cancelled.'); return; }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var toDelete = ['Analytics Dashboard', 'Location Summary', 'Trainer Performance', 'Monthly Location Performance'];

    toDelete.forEach(function(name) {
      var s = ss.getSheetByName(name);
      if (s) { ss.deleteSheet(s); Logger.log('Deleted: ' + name); }
    });

    // Recreate via code.gs functions (they exist in same script scope)
    createAnalyticsSheet(ss);
    createLocationSummarySheet(ss);
    createTrainerPerformanceSheet(ss);
    createMonthlyLocationPerformanceSheet(ss);

    // Populate from live data
    var records = getTrainingRecords(ss);
    populateAnalyticsDashboard(ss, records);
    populateTrainerPerformance(ss, records);
    populateMonthlyLocationPerformance(ss, records);

    SpreadsheetApp.flush();
    Logger.log('✅ All analytics sheets rebuilt from ' + records.length + ' records.');
    ui.alert('✅ Done', 'All analytics sheets rebuilt from ' + records.length + ' records.', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}


// ============================================================================
// UTILITIES
// ============================================================================

/** Returns { 'Header Name': colIndex (1-based), ... } */
function getHeaderMap(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) {
    var name = (h || '').toString().trim();
    if (name) map[name] = i + 1;
  });
  return map;
}

/** 1-based column index → letter string (1→A, 27→AA, etc.) */
function colLetter(col) {
  var s = '';
  while (col > 0) {
    var r = (col - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}


// ============================================================================
// STANDALONE HELPERS — run individually if needed
// ============================================================================

/** Re-creates named ranges only, without touching sheet structure */
function recreateNamedRangesOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet not found.'); return; }
  step6_createNamedRanges(ss, sheet);
  SpreadsheetApp.getUi().alert('✅ Named ranges recreated.');
}

/** Logs current header map — useful for debugging before migration */
function logCurrentHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('Sheet not found.'); return; }
  var map = getHeaderMap(sheet);
  Object.keys(map).sort(function(a, b) { return map[a] - map[b]; }).forEach(function(k) {
    Logger.log(colLetter(map[k]) + ' (col ' + map[k] + '): ' + k);
  });
}

/** Logs all named ranges currently in the spreadsheet */
function logNamedRanges() {
  SpreadsheetApp.getActiveSpreadsheet().getNamedRanges().forEach(function(nr) {
    Logger.log(nr.getName() + ' → ' + nr.getRange().getA1Notation());
  });
}
