// ============================================================
// 30-DAY EVAL BUILDER — THREE POINTS HOSPITALITY (BULLETPROOF v10.0)
// OAK — Original American Kitchen
// 8-COLUMN GRID with proper Section 7 rendering
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Eval Tools')
    .addItem('Build ALL Evals (this location)', 'buildAllEvals')
    .addSeparator()
    .addItem('Build Bartender Eval', 'buildBartenderEval')
    .addItem('Build Server Eval', 'buildServerEval')
    .addItem('Build Host Eval', 'buildHostEval')
    .addToUi();
}

function buildAllEvals() {
  const cfg = getLocationConfig();
  buildEval_(cfg, 'Bartender');
  buildEval_(cfg, 'Server');
  buildEval_(cfg, 'Host');
  SpreadsheetApp.getActive().toast('All 3 role evals built.');
}

function buildBartenderEval() {
  buildEval_(getLocationConfig(), 'Bartender');
  SpreadsheetApp.getActive().toast('Bartender eval built.');
}

function buildServerEval() {
  buildEval_(getLocationConfig(), 'Server');
  SpreadsheetApp.getActive().toast('Server eval built.');
}

function buildHostEval() {
  buildEval_(getLocationConfig(), 'Host');
  SpreadsheetApp.getActive().toast('Host eval built.');
}

// ============================================================
// CONSTANTS
// ============================================================

const COLORS = {
  navy:    '#0f2c53',
  gold:    '#c9a84c',
  text:    '#434343',
  white:   '#ffffff',
  border:  '#1f1f1f',
  lightBg: '#f5f5f5',
  keyBg:   '#e8f0f8',
  rowAlt:  '#fafafa',
  green:   '#1e7e34',
  amber:   '#856404',
  red:     '#721c24',
  scoreInput: '#fffde7'
};

const FONTS = {
  header: 'Lexend',
  body:   'Poppins',
  title:  'Lexend',
  fill:   'Arial'
};

// 8-column grid (added extra column B2 for Section 7)
const COL = {
  MARGIN_L: 1,  // A: 8px
  STUB_L: 2,    // B: 30px
  STUB_L2: 3,   // C: 50px (NEW - for Section 7 "Evaluator" etc.)
  TEXT: 4,      // D: 310px (reduced since we added column)
  STUB_R: 5,    // E: 30px
  SCORE: 6,     // F: 60px
  NOTES: 7,     // G: 230px
  MARGIN_R: 8   // H: 8px
};

const SPAN = { FULL: 6, LEFT: 4, RIGHT: 2 };
const C = COL;

// ============================================================
// CORE BUILDER
// ============================================================

function buildEval_(cfg, role) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = cfg.locationShort + ' ' + role + ' 30-Day Eval';
  let sh = ss.getSheetByName(sheetName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);

  resetSheet_(sh);

  let row = buildHeader_(sh, cfg, role);
  row = buildKeyBlock_(sh, row);
  row = buildSection0_(sh, row, role);

  // Track scored rows for data validation and formulas
  const scoredRows = [];
  const setAScores = [];
  const setBScores = [];
  const setCScores = [];

  const s1Start = row;
  row = buildSection1_(sh, row, scoredRows);

  SpreadsheetApp.flush();

  const s2Start = row;
  row = buildSection2_(sh, row, cfg, role, scoredRows, setAScores, setBScores, setCScores);

  SpreadsheetApp.flush();

  const s3Start = row;
  row = buildSection3_(sh, row, cfg, role, scoredRows);

  SpreadsheetApp.flush();

  row = buildSection4_(sh, row);
  row = buildSection5_(sh, row);

  SpreadsheetApp.flush();

  row = buildSection6_(sh, row, scoredRows, sh);
  row = buildSection7_(sh, row);

  SpreadsheetApp.flush();

  // Apply data validation ONLY to scored rows (not headers)
  applyDataValidation_(sh, scoredRows);
  applyConditionalFormatting_(sh, scoredRows);

  sh.setTabColor(COLORS.navy);
}

// ============================================================
// SHEET RESET
// ============================================================

function resetSheet_(sh) {
  sh.clear();
  sh.clearFormats();
  if (sh.getMaxRows() > 1) sh.deleteRows(2, sh.getMaxRows() - 1);
  if (sh.getMaxColumns() > 1) sh.deleteColumns(2, sh.getMaxColumns() - 1);
  sh.insertRowsAfter(1, 249);
  sh.insertColumnsAfter(1, 7);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setHiddenGridlines(true);

  sh.setColumnWidth(C.MARGIN_L,  8);
  sh.setColumnWidth(C.STUB_L,   30);
  sh.setColumnWidth(C.STUB_L2,  50);
  sh.setColumnWidth(C.TEXT,    450);  // Increased from 310 to 450
  sh.setColumnWidth(C.STUB_R,   30);
  sh.setColumnWidth(C.SCORE,    60);
  sh.setColumnWidth(C.NOTES,   230);
  sh.setColumnWidth(C.MARGIN_R,  8);
}

// ============================================================
// HEADER (COMPACT VERSION)
// ============================================================

function buildHeader_(sh, cfg, role) {
  let row = 1;

  // Row 1: Top padding/margin
  sh.getRange(row, 1, 1, 8).setBackground(COLORS.white);
  sh.setRowHeight(row, 12); row++;

  // Row 2: Split header - Company (left, light) | Eval title (right, navy)
  // Left margin (column A)
  sh.getRange(row, C.MARGIN_L).setBackground(COLORS.white);

  // Left side: Company name (Columns B-D) - 14pt font
  sh.getRange(row, C.STUB_L, 1, 3).merge()
    .setValue('THREE POINTS HOSPITALITY GROUP')
    .setFontFamily(FONTS.title).setFontSize(14).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBackground('#f5f5f5');

  // Right side: Eval title (Columns E-G) - 12pt font, no borders (will be removed)
  sh.getRange(row, C.STUB_R, 1, 3).merge()
    .setValue('30-Day Performance Evaluation')
    .setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBackground(COLORS.navy)
    .setBorder(false, false, false, false, false, false);

  // Right margin (column H)
  sh.getRange(row, C.MARGIN_R).setBackground(COLORS.white);
  sh.setRowHeight(row, 34); row++;

  // Row 3: Location (left) | Role (right) - vertical alignment MIDDLE for both
  sh.getRange(row, C.MARGIN_L).setBackground(COLORS.white);

  sh.getRange(row, C.STUB_L, 1, 3).merge()
    .setValue(cfg.locationDisplay)
    .setFontFamily(FONTS.body).setFontSize(9).setFontColor('#666666')
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBackground(COLORS.white);

  sh.getRange(row, C.STUB_R, 1, 3).merge()
    .setValue(role.toUpperCase())
    .setFontFamily(FONTS.header).setFontSize(13).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBackground(COLORS.white);

  sh.getRange(row, C.MARGIN_R).setBackground(COLORS.white);
  sh.setRowHeight(row, 26); row++;

  // Row 4: Navy spacer bar (full width B-G)
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('')
    .setBackground(COLORS.navy);
  sh.setRowHeight(row, 4); row++;

  // Row 5: Date and Evaluator fields with light borders and bottom alignment
  sh.getRange(row, C.MARGIN_L).setBackground(COLORS.white);

  sh.getRange(row, C.STUB_L, 1, 3).merge()
    .setValue('Date: ______________________')
    .setFontFamily(FONTS.fill).setFontSize(9).setFontColor('#444444')
    .setVerticalAlignment('bottom').setHorizontalAlignment('left')
    .setBackground(COLORS.white)
    .setBorder(null, true, null, null, null, null, '#d6dce4', SpreadsheetApp.BorderStyle.SOLID)
    .setBorder(null, null, true, null, null, null, '#d6dce4', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, C.STUB_R, 1, 3).merge()
    .setValue('Evaluator: ______________________________________')
    .setFontFamily(FONTS.fill).setFontSize(9).setFontColor('#444444')
    .setVerticalAlignment('bottom').setHorizontalAlignment('left')
    .setBackground(COLORS.white)
    .setBorder(null, null, true, true, null, null, '#d6dce4', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, C.MARGIN_R).setBackground(COLORS.white);
  sh.setRowHeight(row, 26); row++;

  // Row 6: Bottom spacer before scoring key
  sh.getRange(row, 1, 1, 8).setBackground(COLORS.white);
  sh.setRowHeight(row, 10); row++;

  return row;
}

// ============================================================
// SCORING KEY BLOCK
// ============================================================

function buildKeyBlock_(sh, row) {
  // Scoring key with rich text: bold "SCORING KEY:" and smaller scale numbers
  const keyText = 'SCORING KEY: 0-4 SCALE\n4 = Exceeds Expectations  |  3 = Meets Expectations |  2 = Needs Improvement  |  1 = Needs Significant Improvement  |  0 = Did Not Do';
  const rich = SpreadsheetApp.newRichTextValue().setText(keyText);

  // "SCORING KEY:" bold 11pt Lexend
  const boldKeyStyle = SpreadsheetApp.newTextStyle().setFontFamily('Lexend').setFontSize(11).setBold(true).setForegroundColor(COLORS.navy).build();
  rich.setTextStyle(0, 13, boldKeyStyle);

  // " 0-4 SCALE" normal weight 11pt Lexend
  const normalScaleStyle = SpreadsheetApp.newTextStyle().setFontFamily('Lexend').setFontSize(11).setBold(false).setForegroundColor(COLORS.navy).build();
  rich.setTextStyle(13, 22, normalScaleStyle);

  // Line break
  const breakStyle = SpreadsheetApp.newTextStyle().setFontFamily('Lexend').setFontSize(11).setBold(true).setForegroundColor(COLORS.navy).build();
  rich.setTextStyle(22, 23, breakStyle);

  // "4 =" bold 9pt
  const boldNumStyle = SpreadsheetApp.newTextStyle().setFontSize(9).setBold(true).setForegroundColor(COLORS.navy).build();
  rich.setTextStyle(23, 26, boldNumStyle);
  rich.setTextStyle(52, 55, boldNumStyle);  // "3 ="
  rich.setTextStyle(55, 56, boldNumStyle);  // "|"
  rich.setTextStyle(78, 81, boldNumStyle);  // "2 ="
  rich.setTextStyle(104, 107, boldNumStyle); // "1 ="
  rich.setTextStyle(107, 108, boldNumStyle); // "|"
  rich.setTextStyle(142, 145, boldNumStyle); // "0 ="
  rich.setTextStyle(145, 146, boldNumStyle); // last section

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setRichTextValue(rich.build())
    .setBackground('#e8f0fb')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true)
    .setBorder(true, true, true, true, null, null, '#b0c4de', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;
  sh.setRowHeight(row, 6); row++;
  return row;
}

// ============================================================
// SECTION 0: EMPLOYEE INFORMATION
// ============================================================

function buildSection0_(sh, row, role) {
  row = writeSectionHeader_(sh, row, ' EMPLOYEE INFORMATION');

  const fields = [
    ['Employee Name', ''],
    ['Position', role],
    ['Hire Date', '']
  ];

  fields.forEach((f, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
      .setValue(f[0])
      .setBackground(bg).setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge()
      .setValue(f[1])
      .setBackground(bg).setFontFamily(FONTS.fill).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 28); row++;  // Changed from 24 to 28
  });

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 1: UNIVERSAL CRITERIA
// ============================================================

function buildSection1_(sh, row, scoredRows) {
  // Section header with rich text (max in italic)
  const headerText = ' 1. UNIVERSAL CRITERIA     Max: 24 pts  (6 criteria × 4)';
  const rich = SpreadsheetApp.newRichTextValue().setText(headerText);
  const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(11).setBold(true).setForegroundColor(COLORS.white).build();
  const italicStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(9).setItalic(true).setForegroundColor(COLORS.white).build();

  rich.setTextStyle(0, 27, boldStyle);  // " 1. UNIVERSAL CRITERIA  "
  rich.setTextStyle(27, headerText.length, italicStyle);  // "   Max: 24 pts  (6 criteria × 4)"

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setRichTextValue(rich.build())
    .setBackground(COLORS.navy)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;
  row = writeColumnHeaders_(sh, row, 'Criterion:', 'Score  /4', 'Observed During Shift?');

  const criteria = [
    ['Guest Interaction', 'Warm, proactive, reads the table. Acknowledges guests within standard windows.'],
    ['Hi! Method', 'Executes the Hi! Method consistently. No guest goes unacknowledged.'],
    ['Teamwork', 'Communicates proactively with the floor. Helps without being asked.'],
    ['Professionalism', 'Appearance, punctuality, and conduct all meet or exceed standard.'],
    ['Composure Under Pressure', 'Stays controlled during high volume. No visible breakdown or avoidance.'],
    ['Attendance', 'Points Earned: Never Late = 4pts | Late 1 time = 3pts | Late 2+ times = 0pts']
  ];

  criteria.forEach((c, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeScoredRow_(sh, row, c[0], c[1], bg);
    scoredRows.push(row);
    sh.setRowHeight(row, 24); row++;
  });

  row = writeSectionTotal_(sh, row, 'Section 1 Score', '24');
  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 2: ACTIVE KNOWLEDGE TEST (WITH SET FORMULAS)
// ============================================================

function buildSection2_(sh, row, cfg, role, scoredRows, setAScores, setBScores, setCScores) {
  // Section header with rich text (max in italic)
  const headerText = ' 2. ACTIVE KNOWLEDGE TEST     Max: 24 pts  (6 questions × 4)';
  const rich = SpreadsheetApp.newRichTextValue().setText(headerText);
  const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(11).setBold(true).setForegroundColor(COLORS.white).build();
  const italicStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(9).setItalic(true).setForegroundColor(COLORS.white).build();

  rich.setTextStyle(0, 30, boldStyle);  // " 2. ACTIVE KNOWLEDGE TEST  "
  rich.setTextStyle(30, headerText.length, italicStyle);  // "   Max: 24 pts  (6 questions × 4)"

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setRichTextValue(rich.build())
    .setBackground(COLORS.navy)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;

  // SET A
  row = writeQuestionSetHeader_(sh, row, 'SET A: Happy Hour & Service Standards', 'Select 2 questions. Circle the numbers used below.');
  row = writeColumnHeaders_(sh, row, 'Question', 'Score  /4', 'Notes / Answer');

  cfg.setA.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    scoredRows.push(row);
    setAScores.push(row);
    sh.setRowHeight(row, 24); row++;
  });

  row = writeQuestionsUsed_(sh, row);
  row = writeSetSubtotalWithFormula_(sh, row, 'Set A Score', '8', setAScores);
  sh.setRowHeight(row, 8); row++;

  // SET B
  const setBKey = 'setB' + role;
  const setBQuestions = cfg[setBKey];

  row = writeQuestionSetHeader_(sh, row, 'SET B: ' + role + ' Menu & Product Knowledge', 'Select 2 questions. Circle the numbers used below.');
  row = writeColumnHeaders_(sh, row, 'Question', 'Score  /4', 'Notes / Answer');

  setBQuestions.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    scoredRows.push(row);
    setBScores.push(row);
    sh.setRowHeight(row, 24); row++;
  });

  row = writeQuestionsUsed_(sh, row);
  row = writeSetSubtotalWithFormula_(sh, row, 'Set B Score', '8', setBScores);
  sh.setRowHeight(row, 8); row++;

  // SET C
  row = writeQuestionSetHeader_(sh, row, 'SET C: Situational Logic', 'Select 2 scenarios. Circle the numbers used below.');
  row = writeColumnHeaders_(sh, row, 'Scenario', 'Score  /4', 'Response Notes');

  cfg.setC.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    scoredRows.push(row);
    setCScores.push(row);
    sh.setRowHeight(row, 24); row++;
  });

  row = writeQuestionsUsed_(sh, row);
  row = writeSetSubtotalWithFormula_(sh, row, 'Set C Score', '8', setCScores);

  row = writeSectionTotal_(sh, row, 'Section 2 Total', '24');
  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 3: ROLE-SPECIFIC ACTIVE TEST
// ============================================================

function buildSection3_(sh, row, cfg, role, scoredRows) {
  // Section header with rich text (max in italic)
  const headerText = ` 3. ${role.toUpperCase()} ACTIVE TEST     Max: 12 pts  (3 tests × 4)`;
  const rich = SpreadsheetApp.newRichTextValue().setText(headerText);
  const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(11).setBold(true).setForegroundColor(COLORS.white).build();
  const italicStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(9).setItalic(true).setForegroundColor(COLORS.white).build();

  const maxIndex = headerText.indexOf('Max:');
  rich.setTextStyle(0, maxIndex, boldStyle);
  rich.setTextStyle(maxIndex, headerText.length, italicStyle);

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setRichTextValue(rich.build())
    .setBackground(COLORS.navy)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;

  const insertKey = role.toLowerCase() + 'Insert';
  const insert = cfg[insertKey];

  row = writeRoleInsertHeader_(sh, row, role.toUpperCase() + ' ACTIVE TEST');
  insert.forEach((test, i) => {
    row = writeActiveTest_(sh, row, (i + 1) + '. ' + test.title, test.instructions, test.rubric, test.notes, scoredRows);
  });

  // Section 3 Total with right alignment
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue('Section 3 Total Points:')
    .setBackground('#1f4e8c').setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle').setHorizontalAlignment('right')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge().setValue('________ / 12')
    .setBackground('#1f4e8c').setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 4: EVALUATOR NOTES
// ============================================================

function buildSection4_(sh, row) {
  row = writeSectionHeader_(sh, row, ' 4. EVALUATOR NOTES');

  const noteFields = [
    'What They Did Well:',
    'Areas for Improvement:',
    'Immediate Coaching Provided:',
    'High-Priority Flags:'
  ];

  noteFields.forEach((label, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;

    sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
      .setValue(label)
      .setBackground(bg).setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, false, true, false, false, '#d6dce4', SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 16); row++;

    sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
      .setValue('')
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9)
      .setBorder(false, true, true, true, false, false, '#d6dce4', SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 32); row++;
  });

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 5: ASSESSMENT SUMMARY
// ============================================================

function buildSection5_(sh, row) {
  row = writeSectionHeader_(sh, row, ' 5. ASSESSMENT SUMMARY');

  const issues = [
    '☐  Product Knowledge Gap',
    '☐  POS Accuracy Issue',
    '☐  Service Recovery Gap',
    '☐  Speed / Efficiency Below Standard',
    '☐  Guest Complaint Risk',
    '☐  Team Communication Issue',
    '☐  Upsell Execution Missing',
    '☐  Side Work / Closing Incomplete'
  ];

  const actions = [
    '☐  Immediate Correction',
    '☐  Additional Training',
    '☐  Follow-Up Assessment',
    '☐  Management Discussion',
    '☐  No Action Required',
    '☐  Trainer Eligible'
  ];

  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
    .setValue('Issues Observed')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge()
    .setValue('Action Required')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 18); row++;

  const maxRows = Math.max(issues.length, actions.length);
  for (let i = 0; i < maxRows; i++) {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
      .setValue(issues[i] || '')
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(8.5).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge()
      .setValue(actions[i] || '')
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(8.5).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 18); row++;
  }

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 6: SCORING + OUTCOME (UPDATED THRESHOLDS)
// ============================================================

function buildSection6_(sh, row, scoredRows, sheet) {
  row = writeSectionHeader_(sh, row, ' 6. SCORING + OUTCOME');

  const scoreRows = [
    ['Section 1: Universal Criteria', '24'],
    ['Section 2: Active Knowledge Test', '24'],
    ['Section 3: Role-Specific Active Test', '12']
  ];

  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
    .setValue('Section').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE)
    .setValue('Max').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.NOTES)
    .setValue('Score').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  scoreRows.forEach((r, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(r[0])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE).setValue(r[1])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.NOTES).setValue('')
      .setBackground(COLORS.scoreInput).setFontFamily(FONTS.fill).setFontSize(11).setFontWeight('bold')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 22); row++;
  });

  // Build SUM formula from scored rows
  const scoreRefs = scoredRows.map(r => `F${r}`).join(',');
  const formula = `=SUM(${scoreRefs})`;

  // TOTAL POINTS row - right-aligned label in B-E
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue('TOTAL POINTS:')
    .setBackground('#1f4e8c').setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.white)
    .setVerticalAlignment('middle').setHorizontalAlignment('right')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Score in F with formula
  sh.getRange(row, C.SCORE).setFormula(formula)
    .setBackground('#e8f0fb').setFontFamily(FONTS.fill).setFontSize(12).setFontWeight('bold')
    .setFontColor('#0a2540').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID)
    .setNumberFormat('0');

  // " / 60" text in G
  sh.getRange(row, C.NOTES).setValue('________ / 60')
    .setBackground('#e8f0fb').setFontFamily(FONTS.fill).setFontSize(12).setFontWeight('bold')
    .setFontColor('#0a2540').setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 26); row++;

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Percentage:       %')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;
  sh.setRowHeight(row, 6); row++;

  // UPDATED THRESHOLDS: 88+ Trainer, 75-87 Development
  const outcomes = [
    { bg: '#d4edda', color: COLORS.green, range: '88–100%  (53–60)', label: '✅  Pass — Trainer Eligible' },
    { bg: '#d4edda', color: COLORS.green, range: '75–87%  (45–52)',  label: '✅  Pass — Standard Development' },
    { bg: '#fff3cd', color: COLORS.amber, range: '60–74%  (36–44)',  label: '⚠️  Conditional Pass — PIP Required' },
    { bg: '#f8d7da', color: COLORS.red,   range: 'Below 60%  (<36)', label: '❌  Does Not Meet Standard' }
  ];

  outcomes.forEach(o => {
    sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
      .setValue(' ' + o.range)  // Space prefix
      .setBackground(o.bg).setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
      .setFontColor(o.color).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge()
      .setValue(o.label)
      .setBackground(o.bg).setFontFamily(FONTS.body).setFontSize(8.5).setFontColor(o.color)
      .setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 20); row++;
  });

  sh.setRowHeight(row, 6); row++;

  // Outcome row with rich text (bold "Outcome:")
  const outcomeText = ' Outcome:   ☐ Trainer Eligible     ☐ Standard Development     ☐ Conditional Pass     ☐ Does Not Meet Standard';
  const rich = SpreadsheetApp.newRichTextValue().setText(outcomeText);
  const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8.5).setBold(true).setForegroundColor(COLORS.text).build();
  const normalStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8.5).setBold(false).setForegroundColor(COLORS.text).build();

  rich.setTextStyle(0, 9, boldStyle);  // " Outcome:"
  rich.setTextStyle(9, outcomeText.length, normalStyle);

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setRichTextValue(rich.build())
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.body).setFontSize(8).setFontColor(COLORS.text)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(' Re-Evaluation Date (if applicable):')  // Space prefix
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.fill).setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 7: SIGN-OFF (8-COLUMN LAYOUT)
// ============================================================

function buildSection7_(sh, row) {
  row = writeSectionHeader_(sh, row, ' 7. SIGN-OFF');

  // Header row - using STUB_L and STUB_L2 for Role column
  sh.getRange(row, C.STUB_L, 1, 2).merge().setValue('Role')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.TEXT).setValue('Printed Name')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.STUB_R).setValue('')
    .setBackground('#d6dce4')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE).setValue('Signature')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.NOTES).setValue('Date')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  // Data rows - merged B-C for role labels with space prefix and left-aligned
  ['Evaluator', 'GM', 'Employee'].forEach((roleLabel, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;

    sh.getRange(row, C.STUB_L, 1, 2).merge().setValue(' ' + roleLabel)  // Space prefix
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle').setHorizontalAlignment('left')  // Changed from center to left
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, C.TEXT).setValue('')
      .setBackground(bg).setFontFamily(FONTS.fill).setFontSize(9)
      .setVerticalAlignment('middle')
      .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, C.STUB_R).setValue('')
      .setBackground(bg)
      .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, C.SCORE).setValue('')
      .setBackground(bg).setFontFamily(FONTS.fill).setFontSize(9)
      .setVerticalAlignment('middle')
      .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, C.NOTES).setValue('')
      .setBackground(bg).setFontFamily(FONTS.fill).setFontSize(9)
      .setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.setRowHeight(row, 28); row++;
  });

  sh.setRowHeight(row, 6); row++;
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Employee signature confirms review of this evaluation and acknowledgment of any required next steps.')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.body).setFontSize(8).setFontStyle('italic')
    .setFontColor('#888888').setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
  sh.setRowHeight(row, 14); row++;

  return row;
}

// ============================================================
// SHARED ROW WRITERS
// ============================================================

function writeSectionHeader_(sh, row, title) {
  // Title already includes number and space prefix, just apply formatting
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(title)
    .setBackground(COLORS.navy)
    .setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold').setFontColor(COLORS.white)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;
  return row;
}

function writeColumnHeaders_(sh, row, col1, col2, col3) {
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(col1)
    .setBackground('#e2e8f0').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor('#0a2540').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE).setValue(col2)
    .setBackground('#e2e8f0').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor('#0a2540').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.NOTES).setValue(col3)
    .setBackground('#e2e8f0').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor('#0a2540').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 18); row++;
  return row;
}

function writeScoredRow_(sh, row, label, description, bg) {
  const fullText = label + '\n' + description;
  const rich = SpreadsheetApp.newRichTextValue().setText(fullText);
  const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(10).setBold(true).setForegroundColor(COLORS.text).build();
  const italicStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8).setItalic(true).setForegroundColor('#64748b').build();
  const boldItalicStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8).setBold(true).setItalic(true).setForegroundColor('#64748b').build();

  // Apply styles
  rich.setTextStyle(0, label.length, boldStyle);
  rich.setTextStyle(label.length, fullText.length, italicStyle);

  // If "Points Earned:" exists in description, make it bold+italic
  if (description.includes('Points Earned:')) {
    const pointsEarnedStart = fullText.indexOf('Points Earned:');
    const pointsEarnedEnd = pointsEarnedStart + 'Points Earned:'.length;
    rich.setTextStyle(pointsEarnedStart, pointsEarnedEnd, boldItalicStyle);
  }

  // Content cell (B-E merged) with left border only
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
    .setRichTextValue(rich.build())
    .setBackground(bg).setVerticalAlignment('middle').setWrap(true)
    .setBorder(null, true, null, null, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  // Score cell (F) with FULL borders
  sh.getRange(row, C.SCORE)
    .setValue('')
    .setBackground(COLORS.scoreInput).setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Notes cell (G) with right border only
  sh.getRange(row, C.NOTES)
    .setValue('')
    .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9)
    .setVerticalAlignment('middle')
    .setBorder(null, null, null, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
}

function writeQuestionRow_(sh, row, num, question, answer, bg) {
  const titleText = num + '  ' + question + '\n';
  const descText = 'Ans: ' + answer;
  const fullText = titleText + descText;

  const rich = SpreadsheetApp.newRichTextValue().setText(fullText);
  const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(10).setBold(true).setForegroundColor(COLORS.text).build();
  const italicStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8).setItalic(true).setForegroundColor('#64748b').build();

  rich.setTextStyle(0, titleText.length, boldStyle);
  rich.setTextStyle(titleText.length, fullText.length, italicStyle);

  // Content cell (B-E merged) with left border only
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
    .setRichTextValue(rich.build())
    .setBackground(bg).setVerticalAlignment('middle').setWrap(true)
    .setBorder(null, true, null, null, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  // Score cell (F) with FULL borders
  sh.getRange(row, C.SCORE)
    .setValue('')
    .setBackground(COLORS.scoreInput).setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Notes cell (G) with right border only
  sh.getRange(row, C.NOTES)
    .setValue('')
    .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9)
    .setVerticalAlignment('middle')
    .setBorder(null, null, null, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
}

function writeQuestionSetHeader_(sh, row, title, instruction) {
  const fullText = title + '\n' + instruction;
  const rich = SpreadsheetApp.newRichTextValue().setText(fullText);
  const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(9.5).setBold(true).setForegroundColor(COLORS.white).build();
  const normalStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(9.5).setBold(false).setForegroundColor(COLORS.white).build();

  // Title is bold, instruction is normal
  rich.setTextStyle(0, title.length, boldStyle);
  rich.setTextStyle(title.length, fullText.length, normalStyle);

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setRichTextValue(rich.build())
    .setBackground('#4a6fa5')
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 26); row++;
  return row;
}

function writeRoleInsertHeader_(sh, row, label) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(label)
    .setBackground('#3a5276').setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeActiveTest_(sh, row, title, instructions, rubric, notes, scoredRows) {
  // Title row (full width B-G)
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(title)
    .setBackground('#e2e8f0').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  // Instructions row (full width B-G, italic)
  // For "Timed Drink Build", remove the drinks/time info (will be separate row)
  let cleanInstructions = instructions;
  const isDrinkBuild = title.includes('Timed Drink Build');
  if (isDrinkBuild) {
    // Extract just the first two lines before "Drinks Called:"
    const parts = instructions.split('\n\n');
    cleanInstructions = parts[0];
  }

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(cleanInstructions)
    .setBackground(COLORS.white).setFontFamily(FONTS.body).setFontSize(8).setFontColor('#64748b')
    .setVerticalAlignment('middle').setWrap(true).setFontStyle('italic')
    .setBorder(true, true, true, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 44); row++;

  // Add "Drinks Called" row for Timed Drink Build only
  if (isDrinkBuild) {
    const drinksText = 'Drinks Called: ________,________ Time:  _______seconds';
    const rich = SpreadsheetApp.newRichTextValue().setText(drinksText);
    const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8).setBold(true).setForegroundColor('#0a2540').build();
    const normalStyle = SpreadsheetApp.newTextStyle().setFontFamily('Arial').setFontSize(8).setBold(false).setForegroundColor('#0a2540').build();

    // "Drinks Called:" bold
    rich.setTextStyle(0, 15, boldStyle);
    // Underscores and commas - Arial normal
    rich.setTextStyle(15, 23, normalStyle);  // ________,
    rich.setTextStyle(23, 24, normalStyle);  // comma
    // "Time:" bold
    rich.setTextStyle(32, 40, boldStyle);  // "Time:  _"
    // Rest normal
    rich.setTextStyle(40, 47, normalStyle);  // "_____se"
    rich.setTextStyle(47, 54, boldStyle);  // "conds"

    sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
      .setRichTextValue(rich.build())
      .setBackground(COLORS.white).setVerticalAlignment('bottom')
      .setBorder(true, true, true, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 24); row++;
  }

  const firstRubricRow = row;

  // Rubric rows (4 rows) - ALL rubrics are B-F merged
  rubric.forEach((r, i) => {
    const bg = COLORS.white;
    const scoreNum = String(4 - i);
    const fullText = '[' + scoreNum + ']  ' + r;

    const rich = SpreadsheetApp.newRichTextValue().setText(fullText);
    const boldStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.header).setFontSize(9).setBold(true).setForegroundColor(COLORS.navy).build();
    const italicStyle = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8).setItalic(true).setForegroundColor('#64748b').build();

    rich.setTextStyle(0, scoreNum.length + 4, boldStyle);
    rich.setTextStyle(scoreNum.length + 4, fullText.length, italicStyle);

    // All rubric rows: Rubric text (B-F merged)
    sh.getRange(row, C.STUB_L, 1, 5).merge()
      .setRichTextValue(rich.build())
      .setBackground(bg).setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, true, true, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

    sh.setRowHeight(row, 22); row++;
  });

  // Notes cell (G) merged vertically across all 4 rubric rows
  sh.getRange(firstRubricRow, C.NOTES, 4, 1).merge()
    .setValue(notes ? notes : '')
    .setBackground('#e2e8f0').setFontFamily(FONTS.header).setFontSize(8).setFontWeight('bold')
    .setFontColor('#64748b').setVerticalAlignment('top').setFontStyle('italic')
    .setBorder(true, true, true, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  // Activity Score row: Label (B-F) + Score dropdown (G)
  const activityNum = title.match(/\d+/)[0];
  sh.getRange(row, C.STUB_L, 1, 5).merge()
    .setValue('Activity ' + activityNum + ' Score:')
    .setBackground('#e2e8f0').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, C.NOTES)
    .setValue('')
    .setBackground(COLORS.scoreInput).setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  scoredRows.push(row);
  sh.setRowHeight(row, 20); row++;

  sh.setRowHeight(row, 6); row++;
  return row;
}

function writeQuestionsUsed_(sh, row) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Questions Used: ________,________ Evaluator Notes: ')
    .setBackground(COLORS.rowAlt).setFontFamily(FONTS.fill).setFontSize(8.5).setFontColor(COLORS.text)
    .setVerticalAlignment('bottom')  // Changed from middle to bottom
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;
  return row;
}

function writeSetSubtotalWithFormula_(sh, row, label, max, scoreRows) {
  // Build formula for this set
  const scoreRefs = scoreRows.map(r => `F${r}`).join(',');
  const formula = `=SUM(${scoreRefs})`;

  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(label + ':')
    .setBackground('#e8f0fb').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, C.SCORE).setFormula(formula)
    .setBackground('#e8f0fb').setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID)
    .setNumberFormat('0" / ' + max + '"');

  sh.getRange(row, C.NOTES).setValue('')
    .setBackground('#e8f0fb')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeSectionTotal_(sh, row, label, max) {
  // Label in B-E (right-aligned)
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(label + ':')
    .setBackground('#1f4e8c').setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle').setHorizontalAlignment('right')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Score in F-G
  sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge().setValue('________ / ' + max)
    .setBackground('#1f4e8c').setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeInsertTotal_(sh, row, label, max) {
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(label + ':  ___ / ' + max)
    .setBackground('#3a5276').setFontFamily(FONTS.fill).setFontSize(9.5).setFontWeight('bold')
    .setFontColor(COLORS.gold).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge().setValue('')
    .setBackground('#3a5276')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 18); row++;
  return row;
}

// ============================================================
// DATA VALIDATION & CONDITIONAL FORMATTING
// ============================================================

function applyDataValidation_(sh, scoredRows) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['0', '1', '2', '3', '4'], true)
    .setAllowInvalid(false)
    .setHelpText('Select a score from 0-4')
    .build();

  scoredRows.forEach(rowNum => {
    sh.getRange(rowNum, C.SCORE).setDataValidation(rule);
  });
}

function applyConditionalFormatting_(sh, scoredRows) {
  const ranges = scoredRows.map(rowNum => sh.getRange(rowNum, C.SCORE));

  const rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(4)
    .setBackground('#d4edda')
    .setFontColor('#1e7e34')
    .setRanges(ranges)
    .build();

  const rule0 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0, 1)
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges(ranges)
    .build();

  sh.setConditionalFormatRules([rule4, rule0]);
}

// ============================================================
// LOCATION CONFIG — Original American Kitchen (OAK)
// ============================================================

function getLocationConfig() {
  return {
    locationShort: 'OAK',
    locationDisplay: 'Original American Kitchen (OAK)',

    setA: [
      { num: 'A-1', question: 'What are the Happy Hour times?', answer: 'Mon–Fri, 4:00 PM – 6:00 PM' },
      { num: 'A-2', question: 'Name the $5 craft beers on Happy Hour.', answer: 'Founders All Day IPA · Kona Big Wave · Dry Wrought Seasonal Cider · Cypress & Grove Prairie Ride' },
      { num: 'A-3', question: 'What are the three $6 classic cocktails on HH?', answer: 'Margarita · Old Fashioned · Gimlet' },
      { num: 'A-4', question: 'Which four house wines are on Happy Hour?', answer: 'Pinot Grigio · Chardonnay · Cabernet Sauvignon · Merlot' },
      { num: 'A-5', question: 'Does HH pricing apply on Gamedays or Holidays?', answer: 'No' },
      { num: 'A-6', question: 'What is the Hi! Method under pressure?', answer: 'Acknowledge immediately · finish current priority correctly · return with control and hospitality' },
      { num: 'A-7', question: 'What is the 10-Foot Rule?', answer: 'Within 10 feet of a guest, make eye contact and acknowledge them' },
      { num: 'A-8', question: 'What is the 60-Second Rule?', answer: 'Greet a newly seated table within 60 seconds' }
    ],

    setBBartender: [
      { num: 'BB-1', question: 'What is the OAK House Margarita spec?', answer: '2oz Cazadores · 1oz Triple Sec · 0.75oz Lime Juice · 0.5oz Agave' },
      { num: 'BB-2', question: 'Describe the Smoked Pecan Old Fashioned.', answer: 'Pecan-infused Buffalo Trace bourbon · turbinado 4-spice pecan simple syrup · chocolate bitters · smoked glass' },
      { num: 'BB-3', question: 'What is the OAK signature draft lineup?', answer: 'Ask manager for current tap list — rotates seasonally' },
      { num: 'BB-4', question: 'Single / Rocks / Double pour sizes?', answer: '1.25oz / 1.75oz / 2.25oz' },
      { num: 'BB-5', question: 'What makes OAK bourbon program distinct?', answer: 'Curated American whiskey selection with seasonal barrel picks' },
      { num: 'BB-6', question: 'Fist grip vs. scissor grip?', answer: 'Fist = precision · Scissor = speed' },
      { num: 'BB-7', question: 'What is a proper whiskey neat presentation?', answer: 'Proper glassware · correct pour · side water offered · no ice unless requested' },
      { num: 'BB-8', question: 'What is the house mocktail or zero-proof offering?', answer: 'Reference current menu — know at least one and describe it confidently' }
    ],

    setBServer: [
      { num: 'BS-1', question: 'Describe the OAK signature burger.', answer: 'Reference current menu — must know protein, toppings, bun, and any notable ingredient' },
      { num: 'BS-2', question: 'What are the two most common allergen concerns on the menu?', answer: 'Gluten (burger buns, breaded items) · Dairy (sauces, cheese)' },
      { num: 'BS-3', question: 'What is the upsell path on a basic burger order?', answer: 'Upgrade protein · add premium topping · suggest a craft beer or cocktail pairing' },
      { num: 'BS-4', question: 'Describe the Steps of Service for OAK.', answer: '60-sec greet · drink order + upsell · app rec · food order with mods confirmed · 2-bite check · refills/table manicure · check on signal · farewell' },
      { num: 'BS-5', question: 'What is the standard for a 2-Bite Check-Back?', answer: 'Return within 2 bites or 2 minutes of entrée delivery. Confirm satisfaction before the guest has to ask.' },
      { num: 'BS-6', question: 'How do you handle a guest with a gluten allergy?', answer: 'Confirm allergy severity. Note in POS. Communicate to kitchen verbally. Confirm with expo before delivery.' },
      { num: 'BS-7', question: 'When do you offer dessert?', answer: 'After entrée plates are cleared — proactively, before the guest asks for the check.' },
      { num: 'BS-8', question: 'Describe a strong farewell at OAK.', answer: 'Thank them by name if known. Mention a reason to return (upcoming special, seasonal item). Hold the door or acknowledge the exit.' }
    ],

    setBHost: [
      { num: 'BH-1', question: 'What is the Farewell Standard?', answer: 'Guest leaves with the same strong impression they arrived with' },
      { num: 'BH-2', question: 'How do you handle a party arriving 15 minutes early for their reservation?', answer: 'Welcome them warmly. Offer the bar if table is not ready. Give an honest ETA. Do not seat them at a dirty or unset table.' },
      { num: 'BH-3', question: 'What information do you gather when adding someone to the waitlist?', answer: 'Name · party size · phone number · any special needs (high chair, accessibility)' },
      { num: 'BH-4', question: 'How do you describe OAK to a first-time guest?', answer: 'American kitchen and bar · seasonal and locally-inspired menu · strong cocktail program · comfortable and approachable' },
      { num: 'BH-5', question: 'A guest complains about the wait time. What do you do?', answer: 'Acknowledge their frustration. Apologize without excuses. Offer a specific remedy (bar, appetizer, realistic updated time). Involve manager if needed.' }
    ],

    setC: [
      { num: 'S-1', question: 'Walk me through the LEAST Method.', answer: 'Listen · Empathize · Apologize · Solve · Thank' },
      { num: 'S-2', question: 'A glass breaks behind the bar. What do you do?', answer: 'Stop. Sweep and contain. Notify manager. Verify no contamination before resuming.' },
      { num: 'S-3', question: 'Table 1 needs refill. Table 2 has hot food in window. Table 3 needs check. Sequence it.', answer: 'Hot food first · Refill second · Check third' },
      { num: 'S-4', question: 'Guest looks intoxicated and orders another round. What do you do?', answer: 'Politely decline. Offer water/food. Involve manager if needed.' },
      { num: 'S-5', question: 'Mid-task and a new guest walks in. What do you do?', answer: 'Acknowledge immediately (Hi! Method). Finish the task. Return with full attention.' },
      { num: 'S-6', question: 'Guest sends a dish back. Walk through your response.', answer: 'Reassure guest. Communicate with kitchen immediately. Keep guest updated. Manager if needed.' },
      { num: 'S-7', question: 'You are running low on a menu item mid-service. What do you do?', answer: 'Alert manager immediately. Note table counts. Communicate the 86 to all servers. Update guests proactively.' },
      { num: 'S-8', question: 'The POS goes down. What do you do?', answer: 'Try basic reset steps. Notify management immediately. Follow manual backup process.' }
    ],

    bartenderInsert: [
      {
        title: 'Timed Drink Build',
        instructions: 'Call two drinks from the OAK pool. Time from first touch to presentation. OAK pool: House Margarita · Smoked Pecan Old Fashioned · Classic Gimlet · Seasonal Cocktail (current menu) · Whiskey Neat with water back · Craft Beer Draft Pull.\n\nDrinks Called:                     ,                          Time:         sec',
        rubric: [
          'Under 75 seconds · clean workspace · accurate build · proper presentation',
          '75–95 seconds OR one minor execution error',
          '96–115 seconds OR two errors OR presentation issue',
          'Over 115 seconds OR errors that require rebuild'
        ],
        notes: 'Errors observed:'
      },
      {
        title: 'Upsell Demonstration',
        instructions: 'Tell the employee: "A guest just ordered a well vodka soda." Ask them to demonstrate their upsell. OAK targets: Premium spirit upgrade · craft beer swap · cocktail suggestion with description.',
        rubric: [
          'Names a specific premium item with appetizing language and price point',
          'Offers an upgrade — missing item name OR price OR enthusiasm',
          'Generic "would you like to upgrade?" with no specifics',
          'Does not attempt upsell'
        ],
        notes: 'Response observed:'
      },
      {
        title: 'Priority Sequencing',
        instructions: '"You are mid-rush. Two tickets on the board. A guest at the bar is signaling. Your barback is asking about ice levels. Sequence your response — first, second, third, fourth."',
        rubric: [
          'Correct sequence · explained clearly · no hesitation: ticket → Hi! Method → barback → ticket 2',
          'Mostly correct · one item out of order',
          'Two items out of order OR unclear logic',
          'Incorrect sequence OR no attempt to explain reasoning'
        ],
        notes: null
      }
    ],

    serverInsert: [
      {
        title: 'Steps of Service Demo',
        instructions: 'Ask employee to walk through a full table from greeting to payment. Score on sequence and completeness — not speed.\n\nTouchpoints: 60-sec Hi! Method · drink order + upsell · app rec · food order with mods · 2-bite/2-min check · proactive refills + table manicure · check on signal · farewell standard\n\nTouchpoints hit:       / 8',
        rubric: [
          '7–8 touchpoints hit · confident delivery · natural flow',
          '5–6 touchpoints hit · minor gaps',
          '3–4 touchpoints hit OR major step skipped',
          '2 or fewer touchpoints OR steps out of sequence'
        ],
        notes: null
      },
      {
        title: 'Upsell Demonstration',
        instructions: '"A guest just ordered the basic burger and a water." Demonstrate your upsell approach. OAK targets: Premium protein upgrade · add a craft beer or cocktail · suggest an appetizer or premium topping.',
        rubric: [
          'Attempts two or more upsell targets with specific language and enthusiasm',
          'Attempts one upsell target with some language',
          'Generic "can I get you anything else?" with no specific item',
          'Takes the order as-is with no upsell attempt'
        ],
        notes: 'Response observed:'
      },
      {
        title: 'Priority Sequencing',
        instructions: '"Table 1 needs a refill. Table 2 has hot food in the window. Table 3 is signaling for their check. You are currently rolling silverware. What do you do?"',
        rubric: [
          'Drop silverware · hot food first · refill second · check third — stated without hesitation',
          'Hot food first · minor order issue on remaining steps',
          'Two steps out of order OR silverware not dropped first',
          'Incorrect priority or no clear answer'
        ],
        notes: null
      }
    ],

    hostInsert: [
      {
        title: 'Greeting Scenario',
        instructions: 'Stand 10 feet away. Walk toward the employee. They must execute the full greeting standard without prompting.',
        rubric: [
          'Eye contact at 10 feet · verbal greeting within 5 feet · full welcome within 60 sec · warm and natural',
          'Greeting initiated but late OR incomplete one element',
          'Greeting only after direct approach OR missing two elements',
          'No acknowledgment until directly addressed'
        ],
        notes: null
      },
      {
        title: 'Reservation Scenario',
        instructions: '"A guest walks up without a reservation on a Friday night. You have a 45-minute wait. Party of 4. Walk me through exactly what you say and do."',
        rubric: [
          'Warm greeting · clear wait time · offer waitlist · explain notification · invite to bar — all 5 beats, confident',
          '3–4 beats · minor gap in communication or tone',
          '2 beats OR abrupt communication OR wait time not stated',
          'Wait time withheld OR walk-in handled before reservations'
        ],
        notes: 'Response observed:'
      },
      {
        title: 'Priority Sequencing',
        instructions: '"Three parties walk in at the same time. Two have reservations. One does not. A server flags you about a table not turning. What do you do?"',
        rubric: [
          'Reservation parties seated first · walk-in communicated to · server acknowledged with 60-sec commitment · door never abandoned',
          'Mostly correct · one beat missed',
          'Walk-in prioritized over reservation OR server flag ignored',
          'Door abandoned OR all three handled incorrectly'
        ],
        notes: null
      }
    ]
  };
}