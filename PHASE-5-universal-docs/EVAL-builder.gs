// ============================================================
// 30-DAY EVAL BUILDER — THREE POINTS HOSPITALITY
// Mirrors Bartender-builder.gs architecture exactly.
// Builds: Cantina 30-Day Eval + OAK 30-Day Eval sheets.
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Eval Tools')
    .addItem('Build ALL Evals', 'buildAllEvals')
    .addSeparator()
    .addItem('Build Cantina 30-Day Eval', 'buildCantinaEval')
    .addItem('Build OAK 30-Day Eval', 'buildOakEval')
    .addSeparator()
    .addItem('Format Active Sheet', 'formatActiveSheet')
    .addToUi();
}

function buildAllEvals() {
  buildCantinaEval();
  buildOakEval();
  SpreadsheetApp.getActive().toast('Both eval sheets built.');
}

function buildCantinaEval() {
  buildEval_(getCantinaConfig());
  SpreadsheetApp.getActive().toast('Cantina 30-Day Eval built.');
}

function buildOakEval() {
  buildEval_(getOakConfig());
  SpreadsheetApp.getActive().toast('OAK 30-Day Eval built.');
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
  title:  'Lexend'
};

// Column layout — mirrors checklist builder exactly
// A(1)=margin B(2)=Y/N C(3)=criterion/text D(4)=gutter E(5)=score F(6)=notes/answer G(7)=margin
const COL = { MARGIN_L: 1, YN: 2, TEXT: 3, GUTTER: 4, SCORE: 5, NOTES: 6, MARGIN_R: 7 };

// ============================================================
// CORE BUILDER
// ============================================================

function buildEval_(cfg) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = cfg.sheetName;
  let sh = ss.getSheetByName(sheetName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);

  resetSheet_(sh);
  let row = buildHeader_(sh, cfg);
  row = buildKeyBlock_(sh, row);
  row = buildSection0_(sh, row);       // Employee Info
  row = buildSection1_(sh, row);       // Universal Criteria
  row = buildSection2_(sh, row, cfg);  // Knowledge Test
  row = buildSection3_(sh, row, cfg);  // Role-Specific Active Test
  row = buildSection4_(sh, row);       // Evaluator Notes
  row = buildSection5_(sh, row);       // Assessment Summary
  row = buildSection6_(sh, row);       // Scoring + Outcome
  row = buildSection7_(sh, row);       // Sign-Off
}

// ============================================================
// SHEET RESET
// ============================================================

function resetSheet_(sh) {
  sh.clear();
  sh.clearFormats();
  if (sh.getMaxRows() > 1)    sh.deleteRows(2, sh.getMaxRows() - 1);
  if (sh.getMaxColumns() > 1) sh.deleteColumns(2, sh.getMaxColumns() - 1);
  sh.insertRowsAfter(1, 199);
  sh.insertColumnsAfter(1, 7);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setHiddenGridlines(true);
  sh.setTabColor(COLORS.navy);

  // Column widths — optimized for print
  sh.setColumnWidth(COL.MARGIN_L, 12);   // A: left margin
  sh.setColumnWidth(COL.YN,       38);   // B: Y/N
  sh.setColumnWidth(COL.TEXT,    420);   // C: criterion text
  sh.setColumnWidth(COL.GUTTER,   12);   // D: gutter
  sh.setColumnWidth(COL.SCORE,    54);   // E: score cell
  sh.setColumnWidth(COL.NOTES,   210);   // F: notes/answer
  sh.setColumnWidth(COL.MARGIN_R, 12);   // G: right margin
}

// ============================================================
// HEADER
// ============================================================

function buildHeader_(sh, cfg) {
  let row = 1;

  // Row 1: Thin navy top bar
  sh.getRange(row, 1, 1, 7).merge().setBackground(COLORS.navy);
  sh.setRowHeight(row, 4);
  row++;

  // Row 2: Company name left + eval title right
  sh.getRange(row, COL.YN, 1, 2).merge()
    .setValue('THREE POINTS HOSPITALITY GROUP')
    .setFontFamily(FONTS.title).setFontSize(13).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle');

  sh.getRange(row, COL.SCORE, 1, 2).merge()
    .setValue('30-DAY PERFORMANCE EVALUATION')
    .setBackground(COLORS.navy)
    .setFontFamily(FONTS.title).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.white)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 30);
  row++;

  // Row 3: Spacer
  sh.setRowHeight(row, 4); row++;

  // Row 4-5: Meta fields
  const metaBorder = '#cccccc';

  sh.getRange(row, COL.YN, 1, 2).merge()
    .setValue('LOCATION: ' + cfg.locationDisplay)
    .setBackground(COLORS.lightBg);
  styleMetaCell_(sh.getRange(row, COL.YN));

  sh.getRange(row, COL.SCORE, 1, 2).merge()
    .setValue('DATE: ____________________')
    .setBackground(COLORS.lightBg);
  styleMetaCell_(sh.getRange(row, COL.SCORE));
  sh.setRowHeight(row, 26); row++;

  sh.getRange(row, COL.YN, 1, 2).merge()
    .setValue('EVALUATOR: ____________________')
    .setBackground(COLORS.lightBg);
  styleMetaCell_(sh.getRange(row, COL.YN));
  sh.getRange(row, COL.YN).setFontFamily('Arial');

  sh.getRange(row, COL.SCORE, 1, 2).merge()
    .setValue('GM PRESENT: ____________________')
    .setBackground(COLORS.lightBg);
  styleMetaCell_(sh.getRange(row, COL.SCORE));
  sh.getRange(row, COL.SCORE).setFontFamily('Arial');

  sh.getRange(row - 1, COL.YN, 2, 4)
    .setBorder(true, true, true, true, true, true, metaBorder, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 26); row++;

  // Row 6: Thin navy bottom bar
  sh.getRange(row, 1, 1, 7).merge().setBackground(COLORS.navy);
  sh.setRowHeight(row, 3); row++;

  // Spacer
  sh.setRowHeight(row, 6); row++;

  return row;
}

function styleMetaCell_(range) {
  range.setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('left').setVerticalAlignment('middle');
}

// ============================================================
// SCORING KEY BLOCK
// ============================================================

function buildKeyBlock_(sh, row) {
  // Title
  sh.getRange(row, COL.YN, 1, 5).merge()
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
    sh.getRange(row, COL.YN).merge()
      .setValue(kr[0])
      .setBackground(bg).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, '#b0c4de', SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, COL.TEXT, 1, 4).merge()
      .setValue(kr[1])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBorder(true, false, true, true, false, false, '#b0c4de', SpreadsheetApp.BorderStyle.SOLID);

    sh.setRowHeight(row, 18); row++;
  });

  // Note about skipping
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue('⚠  If a criterion was genuinely not observable this shift, skip the row and note it in Section 4.')
    .setBackground('#fff8e1')
    .setFontFamily(FONTS.body).setFontSize(8.5).setFontColor('#5d4037').setFontStyle('italic')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#ffe082', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 18); row++;

  sh.setRowHeight(row, 8); row++; // spacer
  return row;
}

// ============================================================
// SECTION 0: EMPLOYEE INFORMATION
// ============================================================

function buildSection0_(sh, row) {
  row = writeSectionHeader_(sh, row, '0', 'EMPLOYEE INFORMATION');

  const fields = [
    ['Employee Name', '____________________________________'],
    ['Position', '☐ Bartender          ☐ Server          ☐ Host'],
    ['First Solo Shift', '________________    Days Since Solo: ________'],
    ['Training Status', '☐ New Hire     ☐ Role Transfer     ☐ Re-Training']
  ];

  fields.forEach((f, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, COL.YN, 1, 2).merge()
      .setValue(f[0])
      .setBackground(bg).setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, COL.SCORE, 1, 2).merge()
      .setValue(f[1])
      .setBackground(bg).setFontFamily('Arial').setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.setRowHeight(row, 22); row++;
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

function buildSection2_(sh, row, cfg) {
  row = writeSectionHeader_(sh, row, '2', 'ACTIVE KNOWLEDGE TEST     Max: 24 pts  (6 questions × 4)');

  // Role callout box
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue('► ROLE BEING EVALUATED:    ☐ Bartender          ☐ Server          ☐ Host\n   Circle the role above. In Set B, complete ONLY the question bank that matches the checked role.')
    .setBackground('#fff3cd')
    .setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
    .setFontColor('#7d4e00').setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, false, false, '#ffc107', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(row, 38); row++;
  sh.setRowHeight(row, 6); row++; // spacer

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

  // SET B
  row = writeQuestionSetHeader_(sh, row, 'SET B: Menu & Product Knowledge', 'Use ONLY the bank matching the role circled above. Select 2 questions.');

  // Bartender sub-header
  row = writeRoleBankHeader_(sh, row, '▸ BARTENDER');
  row = writeColumnHeaders_(sh, row, 'Question', 'Score  /4', 'Notes / Answer');
  cfg.setBBartender.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    sh.setRowHeight(row, 24); row++;
  });

  // Server sub-header
  row = writeRoleBankHeader_(sh, row, '▸ SERVER');
  cfg.setBServer.forEach((q, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    writeQuestionRow_(sh, row, q.num, q.question, q.answer, bg);
    sh.setRowHeight(row, 24); row++;
  });

  // Host sub-header
  row = writeRoleBankHeader_(sh, row, '▸ HOST');
  cfg.setBHost.forEach((q, i) => {
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

function buildSection3_(sh, row, cfg) {
  row = writeSectionHeader_(sh, row, '3', 'ROLE-SPECIFIC ACTIVE TEST     Max: 12 pts  (3 tests × 4)');

  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue('Complete ONLY the insert that matches the role circled in Section 2. Skip the other two.')
    .setBackground('#fff3cd')
    .setFontFamily(FONTS.body).setFontSize(9).setFontWeight('bold')
    .setFontColor('#7d4e00').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#ffc107', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  sh.setRowHeight(row, 6); row++;

  // --- BARTENDER INSERT ---
  row = writeRoleInsertHeader_(sh, row, 'INSERT A — BARTENDER');
  cfg.bartenderInsert.forEach((test, i) => {
    row = writeActiveTest_(sh, row, 'A' + (i + 1) + '. ' + test.title, test.instructions, test.rubric, test.notes);
  });
  row = writeInsertTotal_(sh, row, 'Bartender Insert Total', '12');
  sh.setRowHeight(row, 10); row++;

  // --- SERVER INSERT ---
  row = writeRoleInsertHeader_(sh, row, 'INSERT B — SERVER');
  cfg.serverInsert.forEach((test, i) => {
    row = writeActiveTest_(sh, row, 'B' + (i + 1) + '. ' + test.title, test.instructions, test.rubric, test.notes);
  });
  row = writeInsertTotal_(sh, row, 'Server Insert Total', '12');
  sh.setRowHeight(row, 10); row++;

  // --- HOST INSERT ---
  row = writeRoleInsertHeader_(sh, row, 'INSERT C — HOST');
  cfg.hostInsert.forEach((test, i) => {
    row = writeActiveTest_(sh, row, 'C' + (i + 1) + '. ' + test.title, test.instructions, test.rubric, test.notes);
  });
  row = writeInsertTotal_(sh, row, 'Host Insert Total', '12');
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

    sh.getRange(row, COL.YN, 1, 5).merge()
      .setValue(label)
      .setBackground(bg).setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, false, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 18); row++;

    sh.getRange(row, COL.YN, 1, 5).merge()
      .setValue('')
      .setBackground(bg)
      .setBorder(false, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 34); row++;
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
  sh.getRange(row, COL.YN, 1, 2).merge()
    .setValue('Issues Observed')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Action header (right)
  sh.getRange(row, COL.SCORE, 1, 2).merge()
    .setValue('Action Required')
    .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  const maxRows = Math.max(issues.length, actions.length);
  for (let i = 0; i < maxRows; i++) {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;

    sh.getRange(row, COL.YN, 1, 2).merge()
      .setValue(issues[i] || '')
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, COL.SCORE, 1, 2).merge()
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
    ['Section 1: Universal Criteria', '24', '___'],
    ['Section 2: Active Knowledge Test', '24', '___'],
    ['Section 3: Role-Specific Active Test', '12', '___']
  ];

  // Header
  sh.getRange(row, COL.YN, 1, 2).merge()
    .setValue('Section').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.SCORE).merge()
    .setValue('Available').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.NOTES).merge()
    .setValue('Score').setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9)
    .setFontWeight('bold').setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  scoreRows.forEach((r, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, COL.YN, 1, 2).merge().setValue(r[0])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, COL.SCORE).setValue(r[1])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, COL.NOTES).setValue(r[2])
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 22); row++;
  });

  // Total row
  sh.getRange(row, COL.YN, 1, 2).merge().setValue('TOTAL')
    .setBackground(COLORS.navy).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.SCORE).setValue('60')
    .setBackground(COLORS.navy).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.white).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.NOTES).setValue('___  / 60')
    .setBackground(COLORS.navy).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.gold).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 26); row++;

  // Percentage row
  sh.getRange(row, COL.YN, 1, 4).merge()
    .setValue('Percentage:  ___ %')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold').setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;
  sh.setRowHeight(row, 6); row++;

  // Outcome thresholds
  const outcomes = [
    { bg: '#d4edda', color: COLORS.green, range: '90–100%  (54–60 pts)', label: '✅  Pass — Trainer / Incentive Eligible. Full standard met.' },
    { bg: '#d4edda', color: COLORS.green, range: '75–89%  (45–53 pts)',  label: '✅  Pass — Standard Development. Baseline met; coach to gaps.' },
    { bg: '#fff3cd', color: COLORS.amber, range: '60–74%  (36–44 pts)',  label: '⚠️  Conditional Pass — 30-Day Improvement Plan required.' },
    { bg: '#f8d7da', color: COLORS.red,   range: 'Below 60%  (< 36 pts)', label: '❌  Does Not Meet Standard — PIP or Re-Training. GM decides next step.' }
  ];

  outcomes.forEach(o => {
    sh.getRange(row, COL.YN, 1, 2).merge()
      .setValue(o.range)
      .setBackground(o.bg).setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(o.color).setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, COL.SCORE, 1, 2).merge()
      .setValue(o.label)
      .setBackground(o.bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(o.color)
      .setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 22); row++;
  });

  sh.setRowHeight(row, 6); row++;

  // Outcome selection
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue('Outcome:   ☐ Pass — Trainer Eligible     ☐ Pass — Standard Development     ☐ Conditional Pass     ☐ Does Not Meet Standard')
    .setBackground(COLORS.lightBg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;

  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue('Re-Evaluation Date (if applicable): __________________________')
    .setBackground(COLORS.lightBg).setFontFamily('Arial').setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;

  sh.setRowHeight(row, 8); row++;
  return row;
}

// ============================================================
// SECTION 7: SIGN-OFF
// ============================================================

function buildSection7_(sh, row) {
  row = writeSectionHeader_(sh, row, '7', 'SIGN-OFF');

  const sigRows = [
    ['Evaluator', '', '', ''],
    ['GM', '', '', ''],
    ['Employee', '', '', '']
  ];

  // Sign-off header
  ['Role', 'Printed Name', 'Signature', 'Date'].forEach((h, i) => {
    const cols = [COL.YN, COL.TEXT, COL.SCORE, COL.NOTES];
    const spans = [1, 1, 1, 1];
    sh.getRange(row, cols[i]).setValue(h)
      .setBackground('#d6dce4').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.navy).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.setRowHeight(row, 20); row++;

  sigRows.forEach((r, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    [COL.YN, COL.TEXT, COL.SCORE, COL.NOTES].forEach(col => {
      const val = col === COL.YN ? r[0] : '';
      sh.getRange(row, col).setValue(val)
        .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text).setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    });
    sh.setRowHeight(row, 28); row++;
  });

  // Acknowledgment
  sh.setRowHeight(row, 6); row++;
  sh.getRange(row, COL.YN, 1, 5).merge()
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
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue(`${num}. ${title}`)
    .setBackground(COLORS.navy)
    .setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold').setFontColor(COLORS.white)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24); row++;
  return row;
}

function writeColumnHeaders_(sh, row, col1, col2, col3) {
  sh.getRange(row, COL.YN, 1, 2).merge().setValue(col1)
    .setBackground('#8fa4a7').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.SCORE).setValue(col2)
    .setBackground('#8fa4a7').setFontFamily(FONTS.header).setFontSize(8.5).setFontWeight('bold')
    .setFontColor(COLORS.white).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.NOTES).setValue(col3)
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

  sh.getRange(row, COL.YN, 1, 2).merge()
    .setRichTextValue(rich.build())
    .setBackground(bg).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Score cell
  sh.getRange(row, COL.SCORE)
    .setValue('')
    .setBackground('#fffde7').setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Notes cell
  sh.getRange(row, COL.NOTES)
    .setValue('')
    .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9)
    .setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
}

function writeQuestionRow_(sh, row, num, question, answer, bg) {
  // Num + question in B+C
  const fullText = num + '  ' + question + '\n' + answer;
  const rich = SpreadsheetApp.newRichTextValue().setText(fullText);
  const bold = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(9).setBold(true).setForegroundColor(COLORS.text).build();
  const normal = SpreadsheetApp.newTextStyle().setFontFamily(FONTS.body).setFontSize(8.5).setBold(false).setForegroundColor('#888888').build();
  const numLen = (num + '  ' + question).length;
  rich.setTextStyle(0, numLen, bold);
  rich.setTextStyle(numLen, fullText.length, normal);

  sh.getRange(row, COL.YN, 1, 2).merge()
    .setRichTextValue(rich.build())
    .setBackground(bg).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, COL.SCORE)
    .setValue('')
    .setBackground('#fffde7').setFontFamily(FONTS.header).setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sh.getRange(row, COL.NOTES)
    .setValue('')
    .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9)
    .setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
}

function writeQuestionSetHeader_(sh, row, title, instruction) {
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue(title + '\n' + instruction)
    .setBackground('#4a6fa5').setFontFamily(FONTS.header).setFontSize(9.5).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 28); row++;
  return row;
}

function writeRoleBankHeader_(sh, row, label) {
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue(label)
    .setBackground('#c9daf8').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor('#1a237e').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, '#7986cb', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeRoleInsertHeader_(sh, row, label) {
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue(label)
    .setBackground('#3a5276').setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;
  return row;
}

function writeActiveTest_(sh, row, title, instructions, rubric, notes) {
  // Test title
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue(title)
    .setBackground('#dce8f8').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, '#aac4e0', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;

  // Instructions
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue(instructions)
    .setBackground(COLORS.white).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, notes ? 42 : 36); row++;

  // Rubric rows (4 → 1) with score column
  rubric.forEach((r, i) => {
    const bg = i % 2 === 0 ? COLORS.white : COLORS.rowAlt;
    sh.getRange(row, COL.YN).setValue(String(4 - i))
      .setBackground('#fffde7').setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, COL.TEXT, 1, 2).merge().setValue(r)
      .setBackground(bg).setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, false, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, COL.SCORE, 1, 2).merge().setValue(i === 0 ? 'Score: ___' : '')
      .setBackground(i === 0 ? '#fffde7' : bg).setFontFamily(FONTS.header).setFontSize(10).setFontWeight(i === 0 ? 'bold' : 'normal')
      .setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 20); row++;
  });

  // Notes field if specified
  if (notes) {
    sh.getRange(row, COL.YN, 1, 5).merge()
      .setValue(notes + '  _______________________________________________')
      .setBackground(COLORS.rowAlt).setFontFamily('Arial').setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 20); row++;
  }

  sh.setRowHeight(row, 5); row++;
  return row;
}

function writeQuestionsUsed_(sh, row) {
  sh.getRange(row, COL.YN, 1, 5).merge()
    .setValue('Questions Used: _______  ,  _______          Evaluator Notes: _______________________________________________')
    .setBackground(COLORS.rowAlt).setFontFamily('Arial').setFontSize(9).setFontColor(COLORS.text)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeSetSubtotal_(sh, row, label, max) {
  sh.getRange(row, COL.YN, 1, 2).merge().setValue(label + ':  ___ / ' + max)
    .setBackground('#e8f0fb').setFontFamily(FONTS.header).setFontSize(9).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.SCORE, 1, 2).merge().setValue('')
    .setBackground('#e8f0fb')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 20); row++;
  return row;
}

function writeSectionTotal_(sh, row, label, max) {
  sh.getRange(row, COL.YN, 1, 2).merge().setValue(label + ':  ___ / ' + max)
    .setBackground('#1f4e8c').setFontFamily(FONTS.header).setFontSize(10).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.SCORE, 1, 2).merge().setValue('')
    .setBackground('#1f4e8c')
    .setBorder(true, false, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22); row++;
  return row;
}

function writeInsertTotal_(sh, row, label, max) {
  sh.getRange(row, COL.YN, 1, 2).merge().setValue(label + ':  ___ / ' + max)
    .setBackground('#3a5276').setFontFamily(FONTS.header).setFontSize(9.5).setFontWeight('bold')
    .setFontColor(COLORS.gold).setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, COL.SCORE, 1, 2).merge().setValue('')
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
// CANTINA CONFIG
// ============================================================

function getCantinaConfig() {
  return {
    sheetName: 'Cantina 30-Day Eval',
    locationDisplay: 'Cantina Añejo GNV',

    setA: [
      { num: 'A-1', question: 'What are the Happy Hour times?', answer: 'Mon–Thu, 5:00 PM – 7:00 PM' },
      { num: 'A-2', question: 'Name the five $4 well spirits.', answer: 'E11even Vodka, Bombay Gin, Bacardi Rum, Evan Williams Whiskey, Cazadores Blanco' },
      { num: 'A-3', question: 'True or False: HH applies on Gamedays and Holidays.', answer: 'False' },
      { num: 'A-4', question: 'What are the $7 / $9 / $10 Patron specials?', answer: '$7 Silver · $9 Reposado · $10 Añejo Cantina Barrel Select' },
      { num: 'A-5', question: 'Which Woodbridge wines are on Happy Hour?', answer: 'Cabernet, Rosé, Chardonnay, Pinot Grigio' },
      { num: 'A-6', question: 'What is the Hi! Method under pressure?', answer: 'Acknowledge immediately · finish current priority correctly · return with control and hospitality' },
      { num: 'A-7', question: 'What is the 10-Foot Rule?', answer: 'Acknowledge nearby guests with eye contact and readiness to assist' },
      { num: 'A-8', question: 'What is the 60-Second Rule?', answer: 'Greet a newly seated table within 60 seconds' }
    ],

    setBBartender: [
      { num: 'BB-1', question: 'What is the House Margarita spec?', answer: '2oz Cazadores · 1oz Triple Sec · 0.75oz Lime · 0.5oz Agave' },
      { num: 'BB-2', question: 'What makes the Smoked Añejo Old Fashioned unique?', answer: 'Tequila-based · made with Patron Barrel Select Cantina Añejo' },
      { num: 'BB-3', question: 'What does additive-free tequila mean?', answer: 'Made only with agave, water, and yeast — no legal additives' },
      { num: 'BB-4', question: 'Single / Rocks / Double pour sizes?', answer: '1.25oz / 1.75oz / 2.25oz' },
      { num: 'BB-5', question: 'Frozen Float upcharge and spirit?', answer: '$5 · Bacardi Black float' },
      { num: 'BB-6', question: 'List all draft beers and styles.', answer: 'Modelo (Lager) · Dos XX Lager · Dos XX Ambar · FM Cantina Lager · Blue Moon · Coors Light · Juicy Haze (IPA) · Swamp Head' },
      { num: 'BB-7', question: 'Fist grip vs. scissor grip?', answer: 'Fist = precision · Scissor = speed' },
      { num: 'BB-8', question: 'Añejo Espresso Martini ingredients?', answer: 'Patron Barrel Cantina Añejo · Cazadores Café · espresso · Ancho Reyes · Aztec choc bitters · saline · cinnamon' }
    ],

    setBServer: [
      { num: 'BS-1', question: 'What three items come in the Mexican Trifecta?', answer: 'Salsa Madre · Guacamole Ranchero or Salsa Verde · Queso Blanco' },
      { num: 'BS-2', question: 'Describe the Birria Jalisco tacos.', answer: 'Birria chuck roast & short rib · pickled red onion · shredded Mexican cheese · cilantro · corn tortilla · consomé' },
      { num: 'BS-3', question: 'Is the Mexican Brownie gluten free?', answer: 'No' },
      { num: 'BS-4', question: 'Protein add-on for Grande Nachos?', answer: 'Chicken or Pork +$4 · Birria Beef +$7' },
      { num: 'BS-5', question: 'Two salad dressing options?', answer: 'Cilantro crema · Mango habanero vinaigrette' },
      { num: 'BS-6', question: 'What can be added to Fajitas for upsell?', answer: 'Rice and Beans for $3' },
      { num: 'BS-7', question: 'Describe the Tijuana Caesar dressing.', answer: 'Anchovy · garlic · dijon · cracked black pepper · worcestershire · lime juice · egg yolk' },
      { num: 'BS-8', question: 'Protein options for Enchiladas?', answer: 'Guajillo Chicken · Citrus Carnitas · Tequila Lime Shrimp · Carne Asada' }
    ],

    setBHost: [
      { num: 'BH-1', question: 'Three items in the Mexican Trifecta?', answer: 'Salsa Madre · Guacamole Ranchero or Salsa Verde · Queso Blanco' },
      { num: 'BH-2', question: 'Name the five HH well spirits.', answer: 'E11even Vodka · Bombay Gin · Bacardi Rum · Evan Williams Whiskey · Cazadores Blanco' },
      { num: 'BH-3', question: 'What is the Farewell Standard?', answer: 'Guest leaves with the same strong impression they arrived with' },
      { num: 'BH-4', question: 'How do you explain the Street Kitchen concept?', answer: 'Scratch kitchen with international street-inspired dishes' },
      { num: 'BH-5', question: 'How should you handle first-time guests?', answer: 'Guide them clearly · explain the concept · make strong recommendations' }
    ],

    setC: [
      { num: 'S-1', question: 'Walk me through the LEAST Method.', answer: 'Listen · Empathize · Apologize · Solve · Thank' },
      { num: 'S-2', question: 'Glass breaks in the ice well. What do you do?', answer: 'Stop immediately. Burn the well. No exceptions.' },
      { num: 'S-3', question: 'A guest declines an appetizer. Recovery?', answer: 'Offer a quick-share option later. No pressure. Keep momentum.' },
      { num: 'S-4', question: 'Table 1 needs refill. Table 2 has hot food in window. Table 3 needs check. Sequence it.', answer: 'Hot food first · Refill second · Check third' },
      { num: 'S-5', question: 'Guest looks intoxicated and orders a double. What do you do?', answer: 'Politely decline. Offer water/food. Involve manager if needed.' },
      { num: 'S-6', question: 'Mid-task. New guest walks in. What do you do?', answer: 'Acknowledge immediately (Hi! Method). Finish the task. Return with full attention.' },
      { num: 'S-7', question: 'Guest sends a dish back. Walk through your response.', answer: 'Reassure guest. Communicate with kitchen immediately. Keep guest updated. Manager if needed.' },
      { num: 'S-8', question: 'Low on limes mid-rush. What do you do?', answer: 'Communicate before you run out. Get support. Never abandon bar without coverage.' }
    ],

    bartenderInsert: [
      {
        title: 'Timed Drink Build',
        instructions: 'Call two drinks from the pool below. Time from first touch to presentation. Cantina pool: House Margarita · Smoked Añejo Old Fashioned · Cantina Ranch Water · Category E11even · Añejo Espresso Martini · Frozen Margarita with Bacardi Black Float.\n\nDrinks Called: _____________________ , _____________________     Time: _____ sec',
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
        instructions: 'Tell the employee: "A guest just ordered a standard well margarita." Ask them to demonstrate their upsell. Cantina targets: Patron Silver upgrade ($8) · Bacardi Black float ($5) · $150 Tableside Cantarito for groups.',
        rubric: [
          'Names a specific premium item with appetizing language and correct price',
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
        instructions: 'Ask employee to walk through a full table from greeting to payment. Score on sequence and completeness — not speed.\n\nTouchpoints: 60-sec Hi! Method · drink order + upsell · appetizer rec · food order with mods · 2-bite/2-min check · proactive refills + table manicure · check on signal · farewell standard\n\nTouchpoints hit: ___ / 8',
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
        instructions: '"A guest just ordered three Chicken Street Tacos and a water." Demonstrate your upsell approach.',
        rubric: [
          'Upgrades protein to Birria (+$7) · recommends Rice & Beans ($3) · offers a drink — all with specific language',
          'Attempts two of three upsell targets with some language',
          'Attempts one upsell only OR gives generic "anything else?"',
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

// ============================================================
// OAK CONFIG
// ============================================================

function getOakConfig() {
  return {
    sheetName: 'OAK 30-Day Eval',
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
