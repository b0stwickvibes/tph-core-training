// ============================================================
// 30-DAY EVAL BUILDER — THREE POINTS HOSPITALITY
// Builds one sheet per ROLE per LOCATION.
// Cantina: Bartender / Server / Host
// OAK:     Bartender / Server / Host  (duplicate file, swap config)
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Eval Tools')
    .addItem('Build ALL Evals (this location)', 'buildAllEvals')
    .addSeparator()
    .addItem('Build Bartender Eval', 'buildBartenderEval')
    .addItem('Build Server Eval', 'buildServerEval')
    .addItem('Build Host Eval', 'buildHostEval')
    .addSeparator()
    .addItem('Format Active Sheet', 'formatActiveSheet')
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
// CONSTANTS — same palette/fonts as checklist builder
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
  red:     '#721c24'
};

const FONTS = {
  header: 'Lexend',
  body:   'Poppins',
  title:  'Lexend',
  fill:   'Arial'    // All fill-in / underline fields — change here to restyle globally
};

// Apply the fill-in font to any range (underline fields, signature lines, etc.)
function setFillFont_(range) {
  return range.setFontFamily(FONTS.fill);
}

// Column layout — 7-column grid
// A(1)=margin(8)  B(2)=stub(30)  C(3)=content(330)  D(4)=stub(30)  E(5)=score(52)  F(6)=notes(222)  G(7)=margin(8)
// FULL  = B–F (5 cols from col 2)
// LEFT  = B–D (3 cols from col 2)  — content block
// RIGHT = E–F (2 cols from col 5)  — score+notes block
// CONTENT = C only (1)  |  SCORE = E only (1)  |  NOTES = F only (1)
const COL = { MARGIN_L: 1, STUB_L: 2, TEXT: 3, STUB_R: 4, SCORE: 5, NOTES: 6, MARGIN_R: 7 };
const SPAN = { FULL: 5, LEFT: 3, RIGHT: 2, CONTENT: 3 };
// Convenience aliases so row writers read clearly
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
  row = buildSection0_(sh, row, role);       // Employee Info
  row = buildSection1_(sh, row);             // Universal Criteria
  SpreadsheetApp.flush();
  row = buildSection2_(sh, row, cfg, role);  // Knowledge Test (role-filtered)
  SpreadsheetApp.flush();
  row = buildSection3_(sh, row, cfg, role);  // Role-Specific Active Test (single role)
  SpreadsheetApp.flush();
  row = buildSection4_(sh, row);             // Evaluator Notes
  row = buildSection5_(sh, row);             // Assessment Summary
  SpreadsheetApp.flush();
  row = buildSection6_(sh, row);             // Scoring + Outcome
  row = buildSection7_(sh, row);             // Sign-Off
  SpreadsheetApp.flush();
}

// ============================================================
// SHEET RESET
// ============================================================

function resetSheet_(sh) {
  sh.clear();
  sh.clearFormats();
  if (sh.getMaxRows() > 1)    sh.deleteRows(2, sh.getMaxRows() - 1);
  if (sh.getMaxColumns() > 1) sh.deleteColumns(2, sh.getMaxColumns() - 1);
  sh.insertRowsAfter(1, 249);
  sh.insertColumnsAfter(1, 6);  // 7 columns total: A–G
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setHiddenGridlines(true);
  sh.setTabColor(COLORS.navy);

  // Column widths
  sh.setColumnWidth(C.MARGIN_L,  8);   // A: left margin
  sh.setColumnWidth(C.STUB_L,   30);   // B: left stub (merge control)
  sh.setColumnWidth(C.TEXT,    330);   // C: main content
  sh.setColumnWidth(C.STUB_R,   30);   // D: right stub (merge control)
  sh.setColumnWidth(C.SCORE,    52);   // E: score / points
  sh.setColumnWidth(C.NOTES,   222);   // F: notes / answer
  sh.setColumnWidth(C.MARGIN_R,  8);   // G: right margin
}

// ============================================================
// HEADER
// ============================================================

function buildHeader_(sh, cfg, role) {
  let row = 1;

  // Row 1: breathing room
  sh.setRowHeight(row, 16); row++;

  // Row 2: Company name left | form label right
  sh.getRange(row, C.STUB_L, 1, 3).merge()
    .setValue('THREE POINTS HOSPITALITY GROUP')
    .setFontFamily(FONTS.title).setFontSize(13).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('bottom').setHorizontalAlignment('left')
    .setBackground(COLORS.white);
  sh.getRange(row, C.SCORE, 1, 2).merge()
    .setValue('30-Day Performance Evaluation')
    .setFontFamily(FONTS.body).setFontSize(10)
    .setFontColor('#999999').setVerticalAlignment('bottom').setHorizontalAlignment('right')
    .setBackground(COLORS.white);
  sh.setRowHeight(row, 28); row++;

  // Row 3: Location left | Role right
  sh.getRange(row, C.STUB_L, 1, 3).merge()
    .setValue(cfg.locationDisplay)
    .setFontFamily(FONTS.body).setFontSize(9)
    .setFontColor('#aaaaaa').setVerticalAlignment('top').setHorizontalAlignment('left')
    .setBackground(COLORS.white);
  sh.getRange(row, C.SCORE, 1, 2).merge()
    .setValue(role.toUpperCase())
    .setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('top').setHorizontalAlignment('right')
    .setBackground(COLORS.white);
  sh.setRowHeight(row, 20); row++;

  // Row 4: spacer
  sh.setRowHeight(row, 8); row++;

  // Row 5: thin navy rule — full content width B–F
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge().setBackground(COLORS.navy);
  sh.setRowHeight(row, 2); row++;

  // Row 6: spacer
  sh.setRowHeight(row, 8); row++;

  // Rows 7–8: 2×2 meta grid — left half = B–D, right half = E–F
  // Row 7: Location label | GM fill-in
  sh.getRange(row, C.STUB_L, 1, 3).merge()
    .setValue('Location:    ' + cfg.locationDisplay)
    .setFontFamily(FONTS.body).setFontSize(9.5).setFontColor('#444444')
    .setVerticalAlignment('middle').setHorizontalAlignment('left').setBackground(COLORS.white)
    .setBorder(true, true, true, false, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE, 1, 2).merge()
    .setValue('GM:')
    .setFontFamily(FONTS.body).setFontSize(9.5).setFontColor('#444444')
    .setVerticalAlignment('middle').setHorizontalAlignment('left').setBackground(COLORS.white)
    .setBorder(true, false, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  setFillFont_(sh.getRange(row, C.SCORE, 1, 2));
  sh.setRowHeight(row, 28); row++;

  // Row 8: Date fill-in | Evaluator fill-in
  sh.getRange(row, C.STUB_L, 1, 3).merge()
    .setValue('Date:')
    .setFontFamily(FONTS.fill).setFontSize(9.5).setFontColor('#444444')
    .setVerticalAlignment('middle').setHorizontalAlignment('left').setBackground(COLORS.white)
    .setBorder(true, true, true, false, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE, 1, 2).merge()
    .setValue('Evaluator:')
    .setFontFamily(FONTS.fill).setFontSize(9.5).setFontColor('#444444')
    .setVerticalAlignment('middle').setHorizontalAlignment('left').setBackground(COLORS.white)
    .setBorder(true, false, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 28); row++;

  // Row 9: spacer
  sh.setRowHeight(row, 14); row++;

  return row;
}

function styleMetaCell_(range) {
  range.setFontFamily(FONTS.body).setFontSize(9.5)
    .setFontColor('#444444').setHorizontalAlignment('left').setVerticalAlignment('middle');
}

// ============================================================
// SCORING KEY BLOCK
// ============================================================

function buildKeyBlock_(sh, row) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('SCORING KEY — 0–4 POINT SCALE (applies to all scored rows)')
    .setBackground(COLORS.keyBg)
    .setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBorder(true, true, false, true, false, false, '#b0c4de', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;

  const keyRows = [
    ['4', 'Exceeds Expectations — Performance exceeded the standard.'],
    ['3', 'Meets Expectations — Performance met the standard.'],
    ['2', 'Needs Improvement — Performance was below the standard.'],
    ['1', 'Needs Significant Improvement — Performance was far below the standard.'],
    ['0', 'Did Not Do — The task or expectation was not performed.']
  ];

  keyRows.forEach((kr, i) => {
    const bg = i % 2 === 0 ? COLORS.keyBg : '#dce8f5';
    sh.getRange(row, C.STUB_L)
      .setValue(kr[0])
      .setBackground(bg).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, '#b0c4de', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.TEXT, 1, 4).merge()
      .setValue(kr[1])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBorder(true, false, true, true, false, false, '#b0c4de', SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 18); row++;
  });

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('⚠  If a criterion was genuinely not observable this shift, skip the row and note it in Section 4.')
    .setBackground('#fff8e1')
    .setFontFamily(FONTS.body).setFontSize(8.5).setFontColor('#5d4037').setFontStyle('italic')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#ffe082', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 18); row++;
  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 0: EMPLOYEE INFORMATION
// ============================================================

function buildSection0_(sh, row, role) {
  row = writeSectionHeader_(sh, row, '0', 'EMPLOYEE INFORMATION');

  const fields = [
    ['Employee Name', ''],
    ['Position', role],
    ['First Solo Shift', '                              Days Since Solo:'],
    ['Training Status', '☐ New Hire     ☐ Role Transfer     ☐ Re-Training']
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
    sh.setRowHeight(row, 24); row++;
  });

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 1: UNIVERSAL CRITERIA (6 items, 0-4 scale)
// ============================================================

function buildSection1_(sh, row) {
  row = writeSectionHeader_(sh, row, '1', 'UNIVERSAL CRITERIA     Max: 24 pts  (6 criteria × 4)');
  row = writeColumnHeaders_(sh, row, 'Criterion', 'Score  /4', 'Observed During Shift?');

  const criteria = [
    ['Guest Interaction', 'Warm, proactive, reads the table. Acknowledges guests within standard windows.'],
    ['Hi! Method', 'Executes the Hi! Method consistently. No guest goes unacknowledged.'],
    ['Teamwork', 'Communicates proactively with the floor. Helps without being asked.'],
    ['Professionalism', 'Appearance, punctuality, and conduct all meet or exceed standard.'],
    ['Composure Under Pressure', 'Stays controlled during high volume. No visible breakdown or avoidance.'],
    ['Service Recovery', 'Handles complaints using LEAST Method independently. Correct escalation.']
  ];

  criteria.forEach((c, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeScoredRow_(sh, row, c[0], c[1], bg);
    sh.setRowHeight(row, 24); row++;
  });

  // Section subtotal row
  row = writeSectionTotal_(sh, row, 'Section 1 Score', '24');
  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 2: ACTIVE KNOWLEDGE TEST
// ============================================================

function buildSection2_(sh, row, cfg, role) {
  row = writeSectionHeader_(sh, row, '2', 'ACTIVE KNOWLEDGE TEST     Max: 24 pts  (6 questions × 4)');

  // SET A
  row = writeQuestionSetHeader_(sh, row, 'SET A: Happy Hour & Service Standards', 'Select 2 questions. Circle the numbers used below.');
  row = writeColumnHeaders_(sh, row, 'Question', 'Score  /4', 'Notes / Answer');

  cfg.setA.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    sh.setRowHeight(row, 24); row++;
  });

  row = writeQuestionsUsed_(sh, row);
  row = writeSetSubtotal_(sh, row, 'Set A Score', '8');
  sh.setRowHeight(row, 8); row++;

  // SET B — role-specific bank only
  const setBKey = 'setB' + role;  // setBBartender, setBServer, or setBHost
  const setBQuestions = cfg[setBKey];

  row = writeQuestionSetHeader_(sh, row, 'SET B: ' + role + ' Menu & Product Knowledge', 'Select 2 questions. Circle the numbers used below.');
  row = writeColumnHeaders_(sh, row, 'Question', 'Score  /4', 'Notes / Answer');

  setBQuestions.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    sh.setRowHeight(row, 24); row++;
  });

  row = writeQuestionsUsed_(sh, row);
  row = writeSetSubtotal_(sh, row, 'Set B Score', '8');
  sh.setRowHeight(row, 8); row++;

  // SET C
  row = writeQuestionSetHeader_(sh, row, 'SET C: Situational Logic', 'Select 2 scenarios. Circle the numbers used below.');
  row = writeColumnHeaders_(sh, row, 'Scenario', 'Score  /4', 'Response Notes');

  cfg.setC.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    sh.setRowHeight(row, 24); row++;
  });

  row = writeQuestionsUsed_(sh, row);
  row = writeSetSubtotal_(sh, row, 'Set C Score', '8');

  row = writeSectionTotal_(sh, row, 'Section 2 Total', '24');
  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 3: ROLE-SPECIFIC ACTIVE TEST
// ============================================================

function buildSection3_(sh, row, cfg, role) {
  row = writeSectionHeader_(sh, row, '3', role.toUpperCase() + ' ACTIVE TEST     Max: 12 pts  (3 tests × 4)');

  // Single role insert — map role to config key
  const insertKey = role.toLowerCase() + 'Insert';  // bartenderInsert, serverInsert, hostInsert
  const insert = cfg[insertKey];

  row = writeRoleInsertHeader_(sh, row, role.toUpperCase() + ' ACTIVE TEST');
  insert.forEach((test, i) => {
    row = writeActiveTest_(sh, row, (i + 1) + '. ' + test.title, test.instructions, test.rubric, test.notes);
  });
  row = writeInsertTotal_(sh, row, role + ' Active Test Total', '12');
  sh.setRowHeight(row, 8); row++;

  return row;
}

// ============================================================
// SECTION 4: EVALUATOR NOTES
// ============================================================

function buildSection4_(sh, row) {
  row = writeSectionHeader_(sh, row, '4', 'EVALUATOR NOTES');

  const noteFields = [
    'What They Did Well:',
    'Areas for Improvement:',
    'Immediate Coaching Provided During Evaluation:',
    'High-Priority Flags:'
  ];

  noteFields.forEach((label, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;

    sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
      .setValue(label)
      .setBackground(bg).setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, false, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 18); row++;

    sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
      .setValue('')
      .setBackground(bg)
      .setBorder(false, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 36); row++;
  });

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 5: ASSESSMENT SUMMARY
// ============================================================

function buildSection5_(sh, row) {
  row = writeSectionHeader_(sh, row, '5', 'ASSESSMENT SUMMARY');

  // Two-column layout: issues left, action right
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
    '☐  Immediate Correction — before next shift',
    '☐  Additional Targeted Training',
    '☐  Follow-Up Assessment',
    '☐  Management Discussion Required',
    '☐  No Action Required',
    '☐  Trainer / Incentive Eligible'
  ];

  // Issues header (left)
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
  sh.setRowHeight(row, 20); row++;

  const maxRows = Math.max(issues.length, actions.length);
  for (let i = 0; i < maxRows; i++) {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
      .setValue(issues[i] || '')
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge()
      .setValue(actions[i] || '')
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 20); row++;
  }

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 6: SCORING + OUTCOME
// ============================================================

function buildSection6_(sh, row) {
  row = writeSectionHeader_(sh, row, '6', 'SCORING + OUTCOME');

  // Score table
  const scoreRows = [
    ['Section 1: Universal Criteria', '24', ''],
    ['Section 2: Active Knowledge Test', '24', ''],
    ['Section 3: Role-Specific Active Test', '12', '']
  ];

  // Header
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
    .setValue('Section').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE)
    .setValue('Max').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.NOTES)
    .setValue('Score').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;

  scoreRows.forEach((r, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(r[0])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE).setValue(r[1])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.NOTES).setValue(r[2])
      .setBackground('#fffde7').setFontFamily(FONTS.fill).setFontSize(11).setFontWeight('bold')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 24); row++;
  });

  // Total row
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue('TOTAL')
    .setBackground(COLORS.navy).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE).setValue('60')
    .setBackground(COLORS.navy).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.white).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.NOTES).setValue('/ 60')
    .setBackground(COLORS.navy).setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.gold).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 28); row++;

  // Percentage row
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Percentage:       %')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 26); row++;
  sh.setRowHeight(row, 6); row++;

  // Outcome thresholds
  const outcomes = [
    { bg: '#d4edda', color: COLORS.green, range: '90–100%  (54–60 pts)', label: '✅  Pass — Trainer / Incentive Eligible. Full standard met.' },
    { bg: '#d4edda', color: COLORS.green, range: '75–89%  (45–53 pts)',  label: '✅  Pass — Standard Development. Baseline met; coach to gaps.' },
    { bg: '#fff3cd', color: COLORS.amber, range: '60–74%  (36–44 pts)',  label: '⚠️  Conditional Pass — 30-Day Improvement Plan required.' },
    { bg: '#f8d7da', color: COLORS.red,   range: 'Below 60%  (< 36)',    label: '❌  Does Not Meet Standard — PIP or Re-Training. GM decides next step.' }
  ];

  outcomes.forEach(o => {
    sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
      .setValue(o.range)
      .setBackground(o.bg).setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(o.color).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge()
      .setValue(o.label)
      .setBackground(o.bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(o.color)
      .setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 24); row++;
  });

  sh.setRowHeight(row, 6); row++;

  // Outcome selection
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Outcome:   ☐ Pass — Trainer Eligible     ☐ Pass — Standard Development     ☐ Conditional Pass     ☐ Does Not Meet Standard')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;

  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Re-Evaluation Date (if applicable):')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.fill).setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 7: SIGN-OFF
// ============================================================

function buildSection7_(sh, row) {
  row = writeSectionHeader_(sh, row, '7', 'SIGN-OFF');

  const sigLabels = ['Role', 'Printed Name', 'Signature', 'Date'];
  const sigCols   = [C.STUB_L, C.TEXT, C.SCORE, C.NOTES];

  // Sign-off header
  sigLabels.forEach((h, i) => {
    sh.getRange(row, sigCols[i]).setValue(h)
      .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.setRowHeight(row, 22); row++;

  ['Evaluator', 'GM', 'Employee'].forEach((role, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sigCols.forEach((col, j) => {
      sh.getRange(row, col).setValue(j === 0 ? role : '')
        .setBackground(bg)
        .setFontFamily(j === 0 ? FONTS.body : FONTS.fill)
        .setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    });
    sh.setRowHeight(row, 30); row++;
  });

  // Acknowledgment
  sh.setRowHeight(row, 6); row++;
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Employee signature confirms review of this evaluation and acknowledgment of any required next steps.')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.body).setFontSize(8.5).setFontStyle('italic')
    .setFontColor('#888888').setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
  sh.setRowHeight(row, 16); row++;

  return row;
}

// ============================================================
// SHARED ROW WRITERS
// ============================================================

function writeSectionHeader_(sh, row, num, title) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(`${num}. ${title}`)
    .setBackground(COLORS.navy)
    .setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold').setFontColor(COLORS.white)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;
  return row;
}

function writeColumnHeaders_(sh, row, col1, col2, col3) {
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(col1)
    .setBackground('#8fa4a7').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE).setValue(col2)
    .setBackground('#8fa4a7').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor(COLORS.white).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.NOTES).setValue(col3)
    .setBackground('#8fa4a7').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 18); row++;
  return row;
}

function writeScoredRow_(sh, row, label, description, bg) {
  // Y/N + label merged (B+C)
  const rich = SpreadsheetApp.newRichTextValue().setText(label + '\n' + description);
  const bold = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(9).setBold(true).setForegroundColor(COLORS.text).build();
  const normal = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8.5).setBold(false).setForegroundColor('#666666').build();
  rich.setTextStyle(0, label.length, bold);
  rich.setTextStyle(label.length, label.length + description.length + 1, normal);

  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
    .setRichTextValue(rich.build())
    .setBackground(bg).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Score cell
  sh.getRange(row, C.SCORE)
    .setValue('')
    .setBackground('#fffde7').setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Notes cell
  sh.getRange(row, C.NOTES)
    .setValue('')
    .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9)
    .setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
}

function writeQuestionRow_(sh, row, num, question, answer, bg) {
  const fullText = num + '  ' + question + '\n' + answer;
  const rich = SpreadsheetApp.newRichTextValue().setText(fullText);
  const bold = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(9).setBold(true).setForegroundColor(COLORS.text).build();
  const normal = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8.5).setBold(false).setForegroundColor('#888888').build();
  const numLen = (num + '  ' + question).length;
  rich.setTextStyle(0, numLen, bold);
  rich.setTextStyle(numLen, fullText.length, normal);

  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge()
    .setRichTextValue(rich.build())
    .setBackground(bg).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, C.SCORE)
    .setValue('')
    .setBackground('#fffde7').setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, C.NOTES)
    .setValue('')
    .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9)
    .setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
}

function writeQuestionSetHeader_(sh, row, title, instruction) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(title + '\n' + instruction)
    .setBackground('#4a6fa5').setFontFamily(FONTS.header).setFontSize(9.5).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 28); row++;
  return row;
}

function writeRoleBankHeader_(sh, row, label) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(label)
    .setBackground('#c9daf8').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor('#1a237e').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, '#7986cb', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeRoleInsertHeader_(sh, row, label) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(label)
    .setBackground('#3a5276').setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;
  return row;
}

function writeActiveTest_(sh, row, title, instructions, rubric, notes) {
  // Test title
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(title)
    .setBackground('#dce8f8').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, '#aac4e0', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  // Instructions
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue(instructions)
    .setBackground(COLORS.white).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, notes ? 42 : 36); row++;

  // Rubric rows (4 → 1) with score column
  rubric.forEach((r, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, C.STUB_L).setValue(String(4 - i))
      .setBackground('#fffde7').setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.TEXT, 1, 2).merge().setValue(r)
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge().setValue(i === 0 ? 'Score: ___' : '')
      .setBackground(i === 0 ? '#fffde7' : bg).setFontFamily(FONTS.header).setFontSize(10).setFontWeight(i === 0 ? 'bold' : 'normal')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 20); row++;
  });

  // Notes field if specified
  if (notes) {
    sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
      .setValue(notes + '  _______________________________________________')
      .setBackground(COLORS.rowAlt).setFontFamily(FONTS.fill).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 20); row++;
  }

  sh.setRowHeight(row, 5); row++;
  return row;
}

function writeQuestionsUsed_(sh, row) {
  sh.getRange(row, C.STUB_L, 1, SPAN.FULL).merge()
    .setValue('Questions Used: _______  ,  _______          Evaluator Notes: _______________________________________________')
    .setBackground(COLORS.rowAlt).setFontFamily(FONTS.fill).setFontSize(9).setFontColor(COLORS.text)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeSetSubtotal_(sh, row, label, max) {
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(label + ':  ___ / ' + max)
    .setBackground('#e8f0fb').setFontFamily(FONTS.fill).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge().setValue('')
    .setBackground('#e8f0fb')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeSectionTotal_(sh, row, label, max) {
  sh.getRange(row, C.STUB_L, 1, SPAN.LEFT).merge().setValue(label + ':  ___ / ' + max)
    .setBackground('#1f4e8c').setFontFamily(FONTS.fill).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, C.SCORE, 1, SPAN.RIGHT).merge().setValue('')
    .setBackground('#1f4e8c')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;
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
  sh.setRowHeight(row, 20); row++;
  return row;
}

// ============================================================
// FORMAT ACTIVE SHEET (menu item)
// ============================================================

function formatActiveSheet() {
  SpreadsheetApp.getActive().toast('Formatting complete.');
}

// ============================================================
// LOCATION CONFIG — Cantina Añejo
// Duplicate this file and swap data below for OAK / WB
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
        instructions: 'Call two drinks from the OAK pool. Time from first touch to presentation. OAK pool: House Margarita · Smoked Pecan Old Fashioned · Classic Gimlet · Seasonal Cocktail (current menu) · Whiskey Neat with water back · Craft Beer Draft Pull.\n\nDrinks Called: _____________________ , _____________________     Time: _____ sec',
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
        instructions: 'Ask employee to walk through a full table from greeting to payment. Score on sequence and completeness — not speed.\n\nTouchpoints: 60-sec Hi! Method · drink order + upsell · app rec · food order with mods · 2-bite/2-min check · proactive refills + table manicure · check on signal · farewell standard\n\nTouchpoints hit: ___ / 8',
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
        instructions: '"Three parties walk in simultaneously. Two have reservations. One does not. A server flags you about a table not turning. What do you do?"',
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
