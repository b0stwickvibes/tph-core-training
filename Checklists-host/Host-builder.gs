// ============================================================
// CHECKLIST BUILDER ENGINE
// Reads day configs from separate data files, builds sheets
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Checklist Tools')
    .addItem('Build All Host Days', 'buildAllHostDays')
    .addSeparator()
    .addItem('Build Host Day 1', 'buildHostDay1')
    .addItem('Build Host Day 2', 'buildHostDay2')
    .addItem('Build Host Day 3', 'buildHostDay3')
    .addItem('Build Host Day 4', 'buildHostDay4')
    .addItem('Build Host Day 5', 'buildHostDay5')
    .addSeparator()
    .addItem('Format Active Sheet', 'formatActiveSheet')
    .addItem('Format All Checklist Sheets', 'formatAllChecklistSheets')
    .addToUi();
}

// --- Menu entry points ---

function buildAllHostDays() {
  const ss = SpreadsheetApp.getActive();
  const days = [
    getHostDay1Data(),
    getHostDay2Data(),
    getHostDay3Data(),
    getHostDay4Data(),
    getHostDay5Data()
  ];
  days.forEach(cfg => buildDay_(ss, cfg));
  const first = ss.getSheetByName('Host Day 1');
  if (first) ss.setActiveSheet(first);
  ss.toast('All 5 Host checklists built.');
}

function buildHostDay1() { buildSingleDay_(getHostDay1Data()); }
function buildHostDay2() { buildSingleDay_(getHostDay2Data()); }
function buildHostDay3() { buildSingleDay_(getHostDay3Data()); }
function buildHostDay4() { buildSingleDay_(getHostDay4Data()); }
function buildHostDay5() { buildSingleDay_(getHostDay5Data()); }

function buildSingleDay_(cfg) {
  const ss = SpreadsheetApp.getActive();
  buildDay_(ss, cfg);
  ss.toast(`${cfg.sheetName} built.`);
}

// --- Constants ---

const COLORS = {
  navy:    '#0f2c53',
  text:    '#434343',
  white:   '#ffffff',
  border:  '#1f1f1f',
  sigText: '#333333',
  sigLine: '#555555'
};

const FONTS = {
  header: 'Lexend',
  body:   'Poppins',
  title:  'Lexend'
};

// --- Core builder ---

function buildDay_(ss, cfg) {
  const sheetName = cfg.sheetName || `Host Day ${cfg.day}`;
  let sh = ss.getSheetByName(sheetName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);

  resetSheet_(sh);
  buildHeader_(sh, cfg);

  // Separate special sections for full-width treatment
  let mainSections = cfg.sections;
  let endOfShiftSection = null;
  let finalCheckpointSection = null;

  // Pull "Final Checkpoint" if present (check from the end)
  const last = mainSections[mainSections.length - 1];
  if (last && /final checkpoint/i.test(last.title)) {
    finalCheckpointSection = last;
    mainSections = mainSections.slice(0, -1);
  }

  // Pull "End of Shift" if present (now the new last)
  const newLast = mainSections[mainSections.length - 1];
  if (newLast && /end of shift/i.test(newLast.title)) {
    endOfShiftSection = newLast;
    mainSections = mainSections.slice(0, -1);
  }

  // Layout rule: <= 8 sections = single column, > 8 = two column
  const useTwoColumns = mainSections.length > 8;
  let row = 9; // Content starts after header

  if (useTwoColumns) {
    row = writeSectionsTwoColumn_(sh, row, mainSections);
  } else {
    row = writeSectionsSingleColumn_(sh, row, mainSections);
  }

  // End of Shift: full-width below columns
  if (endOfShiftSection) {
    row += 1;
    row = writeEndOfShiftBlock_(sh, row, endOfShiftSection);
  }

  // Final Checkpoint: full-width below End of Shift
  if (finalCheckpointSection) {
    row += 1;
    row = writeEndOfShiftBlock_(sh, row, finalCheckpointSection);
  }

  row += 1;
  buildSignatures_(sh, row, cfg.day);
}

function resetSheet_(sh) {
  sh.clear();
  sh.clearFormats();
  if (sh.getMaxRows() > 1) sh.deleteRows(2, sh.getMaxRows() - 1);
  if (sh.getMaxColumns() > 1) sh.deleteColumns(2, sh.getMaxColumns() - 1);
  sh.insertRowsAfter(1, 299);
  sh.insertColumnsAfter(1, 8);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setHiddenGridlines(true);
  sh.setTabColor(COLORS.navy);

  // Column widths
  sh.setColumnWidth(1, 18);   // A: left margin
  sh.setColumnWidth(2, 34);   // B: checkbox left
  sh.setColumnWidth(3, 530);  // C: text left
  sh.setColumnWidth(4, 18);   // D: gutter
  sh.setColumnWidth(5, 34);   // E: checkbox right
  sh.setColumnWidth(6, 530);  // F: text right
  sh.setColumnWidth(7, 18);   // G: right margin
  sh.setColumnWidth(8, 18);   // H
  sh.setColumnWidth(9, 18);   // I
}

// --- Header ---

function buildHeader_(sh, cfg) {
  const pageTitle = `${cfg.role.toUpperCase()} DAY ${cfg.day}: Floor Checklist`;
  let row = 1;

  // Row 1: Thin navy accent bar (top rule)
  sh.getRange(row, 1, 1, 7).merge().setBackground(COLORS.navy);
  sh.setRowHeight(row, 4);
  row += 1;

  // Row 2: Company name (left) + page-title badge (right) – single row
  sh.getRange(row, 2, 1, 2).merge()
    .setValue('THREE POINTS HOSPITALITY GROUP');
  sh.getRange(row, 2)
    .setFontFamily(FONTS.title).setFontSize(14).setFontWeight('bold')
    .setFontColor(COLORS.navy).setVerticalAlignment('middle');

  sh.getRange(row, 5, 1, 2).merge()
    .setValue(pageTitle)
    .setBackground(COLORS.navy)
    .setFontFamily(FONTS.title).setFontSize(13).setFontWeight('bold')
    .setFontColor(COLORS.white)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setRowHeight(row, 32);
  row += 1;

  // Row 3: Spacer
  sh.setRowHeight(row, 4);
  row += 1;

  // Rows 4-5: Meta fields on light gray band
  const metaGray = '#f5f5f5';
  const metaBorder = '#cccccc';

  sh.getRange(row, 2, 1, 2).merge()
    .setValue(`DAY ${cfg.day}: ${cfg.title}`)
    .setBackground(metaGray);
  styleMetaCell_(sh.getRange(row, 2));

  sh.getRange(row, 5, 1, 2).merge()
    .setValue('TRAINER: ____________________')
    .setBackground(metaGray);
  styleMetaCell_(sh.getRange(row, 5));
  sh.getRange(row, 5).setFontFamily('Arial');
  sh.setRowHeight(row, 28);
  row += 1;

  sh.getRange(row, 2, 1, 2).merge()
    .setValue('DATE: ____________________')
    .setBackground(metaGray);
  styleMetaCell_(sh.getRange(row, 2));
  sh.getRange(row, 2).setFontFamily('Arial');

  sh.getRange(row, 5, 1, 2).merge()
    .setValue('TRAINEE: ____________________')
    .setBackground(metaGray);
  styleMetaCell_(sh.getRange(row, 5));
  sh.getRange(row, 5).setFontFamily('Arial');

  // Border around the 2-row meta block
  sh.getRange(row - 1, 2, 2, 5)
    .setBorder(true, true, true, true, true, true, metaBorder, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 28);
  row += 1;

  // Row 6: Location bar
  sh.getRange(row, 2, 1, 5).merge()
    .setValue('LOCATION:   ' + cfg.locations.map(l => `☐ ${l}`).join('     '))
    .setBackground('#f9f9f9')
    .setFontFamily(FONTS.body).setFontSize(10).setFontWeight('bold')
    .setFontColor('#000000').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(row, 26);
  row += 1;

  // Row 7: Thin navy accent bar (bottom rule)
  sh.getRange(row, 1, 1, 7).merge().setBackground(COLORS.navy);
  sh.setRowHeight(row, 3);
  row += 1;

  // Row 8: Spacer before content
  sh.setRowHeight(row, 6);
}

function styleMetaCell_(range) {
  range.setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold')
    .setFontColor(COLORS.navy).setHorizontalAlignment('left').setVerticalAlignment('middle');
}

// --- Section writers ---

function writeSectionsSingleColumn_(sh, startRow, sections) {
  let row = startRow;
  sections.forEach(section => {
    row = writeSectionBlock_(sh, row, section, 2, 3, 6); // B-F full width
    row += 1;
  });
  return row;
}

function writeSectionsTwoColumn_(sh, startRow, sections) {
  // Calculate total rows each section needs (1 header + n items + 1 spacer)
  const sectionRows = sections.map(s => 1 + s.items.length + 1);

  // Balance by total row count, not section count
  const totalRows = sectionRows.reduce((a, b) => a + b, 0);
  const halfRows = Math.ceil(totalRows / 2);
  let leftCount = 0;
  let leftTotal = 0;
  for (let i = 0; i < sectionRows.length; i++) {
    if (leftTotal + sectionRows[i] > halfRows && leftCount > 0) break;
    leftTotal += sectionRows[i];
    leftCount++;
  }

  const left = sections.slice(0, leftCount);
  const right = sections.slice(leftCount);

  let leftRow = startRow;
  let rightRow = startRow;

  left.forEach(section => {
    leftRow = writeSectionBlock_(sh, leftRow, section, 2, 3, 3); // B-C
    leftRow += 1;
  });

  right.forEach(section => {
    rightRow = writeSectionBlock_(sh, rightRow, section, 5, 6, 6); // E-F
    rightRow += 1;
  });

  return Math.max(leftRow, rightRow);
}

function writeSectionBlock_(sh, row, section, checkCol, textCol, endCol) {
  const spanCols = endCol - checkCol + 1;

  // Section header
  sh.getRange(row, checkCol, 1, spanCols).merge()
    .setValue(`${section.number}. ${section.title}`)
    .setBackground(COLORS.navy)
    .setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24);
  row += 1;

  // Content rows
  section.items.forEach(item => {
    // Checkbox
    const cb = sh.getRange(row, checkCol);
    cb.insertCheckboxes();
    cb.setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground(COLORS.white)
      .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    // Text (merge from textCol to endCol)
    const textSpan = endCol - textCol + 1;
    if (textSpan > 1) {
      sh.getRange(row, textCol, 1, textSpan).merge();
    }
    const cell = sh.getRange(row, textCol);
    cell.setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle').setWrap(true).setHorizontalAlignment('left')
      .setBackground(COLORS.white)
      .setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    applyLeadInBold_(cell, item);

    sh.setRowHeight(row, 24);
    row += 1;
  });

  return row;
}

// --- End of Shift (full-width, distinct style) ---

function writeEndOfShiftBlock_(sh, row, section) {
  const eosHeader = '#3a3a3a';  // dark charcoal – distinct from navy training headers
  const eosBg     = '#f9f9f9';  // very light gray rows
  const eosBorder = '#999999';

  // Section header – full width B-F
  sh.getRange(row, 2, 1, 5).merge()
    .setValue(`${section.number}. ${section.title}`)
    .setBackground(eosHeader)
    .setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, eosBorder, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24);
  row += 1;

  const blockStart = row;

  // Content rows – checkbox in B, text spans C-F
  section.items.forEach(item => {
    const cb = sh.getRange(row, 2);
    cb.insertCheckboxes();
    cb.setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground(eosBg)
      .setBorder(true, true, true, true, false, false, eosBorder, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(row, 3, 1, 4).merge();
    const cell = sh.getRange(row, 3);
    cell.setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
      .setVerticalAlignment('middle').setWrap(true).setHorizontalAlignment('left')
      .setBackground(eosBg)
      .setBorder(true, true, true, true, false, false, eosBorder, SpreadsheetApp.BorderStyle.SOLID);
    applyLeadInBold_(cell, item);

    sh.setRowHeight(row, 24);
    row += 1;
  });

  return row;
}

// --- Signatures ---

function buildSignatures_(sh, row, dayNumber) {
  // Header bar
  sh.getRange(row, 2, 1, 5).merge()
    .setValue('SIGNATURES')
    .setBackground(COLORS.navy)
    .setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold')
    .setFontColor(COLORS.white).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24);
  row += 1;

  const sigStartRow = row;

  const blocks = [
    {
      label: 'Trainee Confirmation: I completed all items checked above and understand the material covered today.',
      sig: 'Signature: ____________________    Date: ____________'
    },
    {
      label: 'Trainer Confirmation: I covered all checked items thoroughly and the trainee demonstrated understanding.',
      sig: 'Signature: ____________________    Date: ____________'
    }
  ];

  if (dayNumber === 1) {
    blocks.push({
      label: 'MOD Verification (Day 1 Required):',
      sig: 'Signature: ____________________    Date: ____________'
    });
  }

  blocks.forEach((block, idx) => {
    // Label
    sh.getRange(row, 2, 1, 5).merge().setValue(block.label)
      .setFontFamily(FONTS.body).setFontSize(9).setFontWeight('bold')
      .setFontColor(COLORS.sigText).setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(row, 24);
    row += 1;

    // Signature line
    sh.getRange(row, 2, 1, 5).merge().setValue(block.sig)
      .setFontFamily('Arial').setFontSize(9).setFontWeight('normal')
      .setFontColor(COLORS.sigLine).setVerticalAlignment('middle');
    sh.setRowHeight(row, 22);
    row += 1;

    // Spacer between blocks
    if (idx < blocks.length - 1) {
      sh.setRowHeight(row, 8);
      row += 1;
    }
  });

  // Border around entire signature area
  sh.getRange(sigStartRow, 2, row - sigStartRow, 5)
    .setBorder(true, true, true, true, null, null, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  return row;
}

// --- Rich text helper ---

function applyLeadInBold_(cell, text) {
  const rich = SpreadsheetApp.newRichTextValue().setText(text);

  const normal = SpreadsheetApp.newTextStyle()
    .setFontFamily(FONTS.body).setFontSize(9)
    .setForegroundColor(COLORS.text).setBold(false).build();

  const bold = SpreadsheetApp.newTextStyle()
    .setFontFamily(FONTS.body).setFontSize(9)
    .setForegroundColor(COLORS.text).setBold(true).build();

  rich.setTextStyle(0, text.length, normal);

  let cut = -1;
  const semi = text.indexOf(';');
  const colon = text.indexOf(':');
  if (semi >= 0) {
    cut = semi + 1;
  } else if (colon >= 0) {
    cut = colon + 1;
  }

  if (cut > 0) {
    rich.setTextStyle(0, cut, bold);
  }

  cell.setRichTextValue(rich.build());
}

// --- Format Cleaner ---

function formatActiveSheet() {
  formatChecklistSheet_(SpreadsheetApp.getActiveSheet());
  SpreadsheetApp.getActive().toast('Active sheet formatted.');
}

function formatAllChecklistSheets() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => {
    if (/Day\s\d+/i.test(sh.getName())) {
      formatChecklistSheet_(sh);
    }
  });
  ss.toast('All checklist sheets formatted.');
}

function formatChecklistSheet_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (!lastRow || !lastCol) return;

  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const backgrounds = sh.getRange(1, 1, lastRow, lastCol).getBackgrounds();

  for (let r = 0; r < lastRow; r++) {
    for (let c = 0; c < lastCol; c++) {
      const text = String(values[r][c] || '').trim();
      if (!text) continue;
      const cell = sh.getRange(r + 1, c + 1);
      if (cell.isPartOfMerge()) continue;

      const bg = String(backgrounds[r][c]).toLowerCase();
      const darkBgs = ['#0f2c53', '#103763', '#274e8a', '#1f4e78', '#2f5597', '#234b8c'];

      if (darkBgs.indexOf(bg) !== -1 || /^\d+\.\s/.test(text) || text === 'SIGNATURES') {
        cell.setFontFamily(FONTS.header).setFontSize(11).setFontWeight('bold')
          .setFontColor(COLORS.white).setBackground(COLORS.navy);
        continue;
      }

      if (typeof cell.isChecked === 'function') {
        const checked = cell.isChecked();
        if (checked === true || checked === false) continue;
      }

      cell.setFontFamily(FONTS.body).setFontSize(9).setFontColor(COLORS.text)
        .setFontWeight('normal').setVerticalAlignment('middle');
      applyLeadInBold_(cell, text);
    }
  }

  // Compact row heights (skip header area rows 1-8)
  for (let r = 9; r <= lastRow; r++) {
    const h = sh.getRowHeight(r);
    if (h > 26) continue;
    sh.setRowHeight(r, Math.min(h, 24));
  }
}
