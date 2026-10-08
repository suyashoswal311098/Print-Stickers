function applyConditionalFormattingToH2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange("H2"); // Targeting cell H2

  // Fetch existing rules
  const existingRules = sheet.getConditionalFormatRules();

  // Create conditional formatting rules for H2
  const rules = [
    { condition: '1 - Orange', color: '#ff3300' },
    { condition: '2 - Blue', color: '#004d84' },
    { condition: '3 - Green', color: '#268700' },
    { condition: '4 - Pink', color: '#741b47' },
    { condition: '5 - White', color: '#ffffff' }
  ];

  const newRulesForH2 = rules.map(rule =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$G2="${rule.condition}"`)
      .setBackground(rule.color)
      .setRanges([range])
      .build()
  );

  // Combine existing rules with new rules for H2
  const updatedRules = existingRules.filter(rule => {
    // Exclude rules targeting H2
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getA1Notation() === "H2");
  }).concat(newRulesForH2);

  // Apply updated rules
  sheet.setConditionalFormatRules(updatedRules);
}

function applyConditionalFormattingToC4() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange("C4"); // Targeting cell C4

  // Fetch existing rules
  const existingRules = sheet.getConditionalFormatRules();

  // Create conditional formatting rules for C4
  const newRulesForC4 = [
    // Rule 1: Red background if value > 0
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=C4>0")
      .setBackground("#ff0000") // Red background color
      .setRanges([range])
      .build(),
    // Rule 2: Background color #215967 if value = 0
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=C4=0")
      .setBackground("#215967") // Custom background color
      .setFontColor("#ffff00") // Yellow text color
      .setRanges([range])
      .build()
  ];

  // Combine existing rules with new rules for C4
  const updatedRules = existingRules.filter(rule => {
    // Exclude rules targeting C4
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getA1Notation() === "C4");
  }).concat(newRulesForC4);

  // Apply updated rules
  sheet.setConditionalFormatRules(updatedRules);
}


function applyConditionalFormattingI_N_O() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Import List');

  // Ranges
  const rangeI = sheet.getRange('I4:I');
  const rangeN = sheet.getRange('N4:N');
  const rangeO = sheet.getRange('O4:O');

  // Color constants
  const DARK_GREEN = '#006100';   // Dark Green 2
  const LIGHT_RED  = '#FFC7CE';   // Light Red 1 (berry1)

  const rules = [
    // ─── Column I: all background rules ─────────────────────────────────
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D4>$N4')
      .setBackground(DARK_GREEN)
      .setRanges([rangeI])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D4<$N4')
      .setBackground(LIGHT_RED)
      .setRanges([rangeI])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E4>$O4')
      .setBackground(DARK_GREEN)
      .setRanges([rangeI])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E4<$O4')
      .setBackground(LIGHT_RED)
      .setRanges([rangeI])
      .build(),

    // ─── Column N: background only based on D vs N ───────────────────────
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D4>$N4')
      .setBackground(DARK_GREEN)
      .setRanges([rangeN])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D4<$N4')
      .setBackground(LIGHT_RED)
      .setRanges([rangeN])
      .build(),

    // ─── Column O: background only based on E vs O ───────────────────────
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E4>$O4')
      .setBackground(DARK_GREEN)
      .setRanges([rangeO])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E4<$O4')
      .setBackground(LIGHT_RED)
      .setRanges([rangeO])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}






 function filterNonBlank() {
  const ss        = SpreadsheetApp.getActive();
  const sheet     = ss.getSheetByName('Import List');  
  const headerRow = 2;
  const lastRow   = sheet.getLastRow();
  const lastCol   = sheet.getLastColumn();

  // Remove any existing filter
  if (sheet.getFilter()) sheet.getFilter().remove();

  // Build a filter range from the header through your last row
  const filterRange = sheet.getRange(
    headerRow,       // row 2 = header
    1,               // col A
    lastRow - headerRow + 1,
    lastCol
  );
  const filter = filterRange.createFilter();

  // Show only rows where column P (16) is not empty
  const criteria = SpreadsheetApp
    .newFilterCriteria()
    .whenCellNotEmpty()
    .build();
  filter.setColumnFilterCriteria(9, criteria);
}
