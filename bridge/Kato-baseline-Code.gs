const BRIDGE_VERSION = '1.3-live-study';
const STUDY_EXTRA_TABS = [
  'A-Z Word Bank',
  'Everyday Phrasal Verbs',
  'Business Phrasal Verbs',
  'Proverbs'
];
// Replace this value for each student.
// Ishikawa — fresh project using Kato baseline
const STUDENT_SPREADSHEET_ID =
  '1q1k6IEiBQ1AMWYCzW7sxelAFBoAyGuT_h83ZQi8jKRs';
function setupBridge() {
  if (STUDENT_SPREADSHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
    throw new Error('Enter the student spreadsheet ID first.');
  }
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(
    'SPREADSHEET_ID',
    STUDENT_SPREADSHEET_ID
  );
  properties.deleteProperty('ACCESS_TOKEN');
  properties.deleteProperty('MAX_ACCESSIBLE_DAY');
  console.log('Bridge setup complete.');
}
function doGet(e) {
  try {
    const settings = getSettings_();
    return jsonResponse_({
      ok: true,
      service: 'A-Z Word Bank Bridge',
      version: BRIDGE_VERSION,
      status: 'ready',
      maxAccessibleDay: settings.maxAccessibleDay
    });
  } catch (error) {
    return errorResponse_(error);
  }
}
function doPost(e) {
  try {
    const settings = getSettings_();
    const request = parseRequest_(e);
    const action = cleanText_(request.action).toLowerCase();
    const spreadsheet = settings.spreadsheet;
    if (action === 'days') {
      const days = getAccessibleDayNames_(
        spreadsheet,
        settings.maxAccessibleDay
      );
      return jsonResponse_({
        ok: true,
        action: 'days',
        days: days,
        studySheets: getAccessibleStudySheetNames_(
          spreadsheet,
          settings.maxAccessibleDay
        )
      });
    }
    if (action === 'study') {
      const studySheetName = cleanText_(request.sheet);
      const studySheet = getAccessibleStudySheet_(
        spreadsheet,
        studySheetName,
        settings.maxAccessibleDay
      );
      return jsonResponse_({
        ok: true,
        action: 'study',
        sheet: studySheetName,
        cards: getStudyCards_(studySheet)
      });
    }
    const sheetName = cleanText_(request.sheet);
    const sheet = getAccessibleDaySheet_(
      spreadsheet,
      sheetName,
      settings.maxAccessibleDay
    );
    if (action === 'list') {
      return jsonResponse_({
        ok: true,
        action: 'list',
        sheet: sheetName,
        entries: getEntries_(sheet)
      });
    }
    if (
      action !== 'add' &&
      action !== 'update' &&
      action !== 'delete'
    ) {
      throw new Error('Unsupported action.');
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      let result;
      if (action === 'add') {
        const word = cleanText_(request.word);
        const definition = cleanText_(request.definition);
        validateEntry_(word, definition);
        const row = findNextEntryRow_(sheet);
        sheet
          .getRange(row, 1, 1, 2)
          .setValues([[word, definition]]);
        addWordBankEntry_(
          spreadsheet,
          sheetName,
          word,
          definition
        );
        result = {
          ok: true,
          action: 'add',
          sheet: sheetName,
          row: row,
          word: word
        };
      }
      if (action === 'update') {
        const row = validateRow_(request.row, sheet);
        const word = cleanText_(request.word);
        const definition = cleanText_(request.definition);
        validateEntry_(word, definition);
        ensureEntryExists_(sheet, row);
        const oldValues = sheet
          .getRange(row, 1, 1, 2)
          .getValues()[0];
        sheet
          .getRange(row, 1, 1, 2)
          .setValues([[word, definition]]);
        updateWordBankEntry_(
          spreadsheet,
          settings.maxAccessibleDay,
          sheetName,
          oldValues[0],
          oldValues[1],
          word,
          definition
        );
        result = {
          ok: true,
          action: 'update',
          sheet: sheetName,
          row: row,
          word: word
        };
      }
      if (action === 'delete') {
        const row = validateRow_(request.row, sheet);
        ensureEntryExists_(sheet, row);
        const oldValues = sheet
          .getRange(row, 1, 1, 2)
          .getValues()[0];
        sheet
          .getRange(row, 1, 1, 2)
          .clearContent();
        deleteWordBankEntry_(
          spreadsheet,
          settings.maxAccessibleDay,
          sheetName,
          oldValues[0],
          oldValues[1]
        );
        result = {
          ok: true,
          action: 'delete',
          sheet: sheetName,
          row: row
        };
      }
      SpreadsheetApp.flush();
      return jsonResponse_(result);
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return errorResponse_(error);
  }
}
function getSettings_() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId =
    properties.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error(
      'Bridge setup is incomplete. Run setupBridge first.'
    );
  }
  const spreadsheet =
    SpreadsheetApp.openById(spreadsheetId);
  const adminSheet =
    spreadsheet.getSheetByName('Admin Settings');
  if (!adminSheet) {
    throw new Error(
      'The Admin Settings sheet was not found.'
    );
  }
  const maxAccessibleDay = Number(
    adminSheet.getRange('B1').getValue()
  );
  if (
    !Number.isInteger(maxAccessibleDay) ||
    maxAccessibleDay < 1 ||
    maxAccessibleDay > 100
  ) {
    throw new Error(
      'Admin Settings!B1 must contain a whole number from 1 to 100.'
    );
  }
  return {
    spreadsheetId: spreadsheetId,
    spreadsheet: spreadsheet,
    maxAccessibleDay: maxAccessibleDay
  };
}
function parseRequest_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('No request data was received.');
  }
  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error(
      'The request data was not valid JSON.'
    );
  }
}
function getAccessibleDayNames_(spreadsheet, maxDay) {
  return spreadsheet
    .getSheets()
    .map(function(sheet) {
      return sheet.getName();
    })
    .filter(function(name) {
      const match =
        name.match(/^Day\s+([1-9]\d*)$/i);
      return (
        match &&
        Number(match[1]) <= maxDay
      );
    })
    .sort(function(a, b) {
      return getDayNumber_(a) - getDayNumber_(b);
    });
}
function getAccessibleStudySheetNames_(
  spreadsheet,
  maxDay
) {
  const names =
    getAccessibleDayNames_(spreadsheet, maxDay);
  STUDY_EXTRA_TABS.forEach(function(name) {
    if (spreadsheet.getSheetByName(name)) {
      names.push(name);
    }
  });
  return names;
}
function getAccessibleStudySheet_(
  spreadsheet,
  sheetName,
  maxDay
) {
  const dayMatch =
    sheetName.match(/^Day\s+([1-9]\d*)$/i);
  if (dayMatch) {
    if (Number(dayMatch[1]) > maxDay) {
      throw new Error(
        'This Day has not been released by the administrator.'
      );
    }
  } else if (!STUDY_EXTRA_TABS.includes(sheetName)) {
    throw new Error(
      'The selected Study sheet is not available.'
    );
  }
  const sheet =
    spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(
      'The selected Study sheet was not found.'
    );
  }
  return sheet;
}
function getStudyCards_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues()
    .map(function(row) {
      return [
        cleanText_(row[0]),
        cleanText_(row[1])
      ];
    })
    .filter(function(row) {
      return row[0] && row[1];
    });
}
function getAccessibleDaySheet_(
  spreadsheet,
  sheetName,
  maxDay
) {
  const match =
    sheetName.match(/^Day\s+([1-9]\d*)$/i);
  if (!match) {
    throw new Error(
      'Entries may be managed only in Day sheets.'
    );
  }
  if (Number(match[1]) > maxDay) {
    throw new Error(
      'This Day has not been released by the administrator.'
    );
  }
  const sheet =
    spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(
      'The selected Day sheet was not found.'
    );
  }
  return sheet;
}
function getDayNumber_(sheetName) {
  const match =
    sheetName.match(/^Day\s+([1-9]\d*)$/i);
  return match ? Number(match[1]) : 0;
}
function getEntries_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues();
  const entries = [];
  values.forEach(function(row, index) {
    const word = cleanText_(row[0]);
    const definition = cleanText_(row[1]);
    if (word || definition) {
      entries.push({
        row: index + 2,
        word: word,
        definition: definition
      });
    }
  });
  return entries;
}
function findNextEntryRow_(sheet) {
  const lastRow =
    Math.max(sheet.getLastRow(), 1);
  if (lastRow >= 2) {
    const values = sheet
      .getRange(2, 1, lastRow - 1, 2)
      .getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      const word =
        cleanText_(values[i][0]);
      const definition =
        cleanText_(values[i][1]);
      if (!word && !definition) {
        return i + 2;
      }
    }
  }
  return Math.max(lastRow + 1, 2);
}
function validateRow_(value, sheet) {
  const row = Number(value);
  if (
    !Number.isInteger(row) ||
    row < 2 ||
    row > sheet.getLastRow()
  ) {
    throw new Error(
      'The selected vocabulary entry is invalid.'
    );
  }
  return row;
}
function ensureEntryExists_(sheet, row) {
  const values = sheet
    .getRange(row, 1, 1, 2)
    .getDisplayValues()[0];
  if (
    !cleanText_(values[0]) &&
    !cleanText_(values[1])
  ) {
    throw new Error(
      'That vocabulary entry no longer exists. Refresh the list.'
    );
  }
}
function validateEntry_(word, definition) {
  if (!word) {
    throw new Error(
      'Please enter a word or phrase.'
    );
  }
  if (!definition) {
    throw new Error(
      'Please enter a definition.'
    );
  }
  if (word.length > 250) {
    throw new Error(
      'The word or phrase is too long.'
    );
  }
  if (definition.length > 2000) {
    throw new Error(
      'The definition is too long.'
    );
  }
}
function getWordBankSheet_(spreadsheet) {
  const target =
    spreadsheet.getSheetByName('A-Z Word Bank');
  if (!target) {
    throw new Error(
      'The A-Z Word Bank sheet was not found.'
    );
  }
  return target;
}
function getWordBankData_(target) {
  const lastRow = target.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return target
    .getRange(2, 1, lastRow - 1, 3)
    .getValues()
    .filter(function(row) {
      return cleanText_(row[0]);
    });
}
function sortWordBankData_(data) {
  data.sort(function(a, b) {
    return cleanText_(a[0]).localeCompare(
      cleanText_(b[0])
    );
  });
}
function writeWordBankData_(
  target,
  data,
  previousCount
) {
  const requiredRows = data.length + 1;
  if (requiredRows > target.getMaxRows()) {
    target.insertRowsAfter(
      target.getMaxRows(),
      requiredRows - target.getMaxRows()
    );
  }
  const rowsToClear =
    Math.max(previousCount, data.length);
  if (rowsToClear > 0) {
    target
      .getRange(2, 1, rowsToClear, 3)
      .clearContent();
  }
  if (data.length > 0) {
    target
      .getRange(2, 1, data.length, 3)
      .setValues(data);
  }
}
function findWordBankEntryIndex_(
  data,
  sheetName,
  word,
  definition
) {
  const cleanSheetName =
    cleanText_(sheetName);
  const cleanWord =
    cleanText_(word);
  const cleanDefinition =
    cleanText_(definition);
  return data.findIndex(function(row) {
    return (
      cleanText_(row[0]) === cleanWord &&
      cleanText_(row[1]) === cleanDefinition &&
      cleanText_(row[2]) === cleanSheetName
    );
  });
}
function addWordBankEntry_(
  spreadsheet,
  sheetName,
  word,
  definition
) {
  const target =
    getWordBankSheet_(spreadsheet);
  const data =
    getWordBankData_(target);
  const previousCount = data.length;
  data.push([
    word,
    definition,
    sheetName
  ]);
  sortWordBankData_(data);
  writeWordBankData_(
    target,
    data,
    previousCount
  );
}
function updateWordBankEntry_(
  spreadsheet,
  maxDay,
  sheetName,
  oldWord,
  oldDefinition,
  newWord,
  newDefinition
) {
  const target =
    getWordBankSheet_(spreadsheet);
  const data =
    getWordBankData_(target);
  const previousCount = data.length;
  const index = findWordBankEntryIndex_(
    data,
    sheetName,
    oldWord,
    oldDefinition
  );
  if (index === -1) {
    rebuildWordBank_(spreadsheet, maxDay);
    return;
  }
  data[index] = [
    newWord,
    newDefinition,
    sheetName
  ];
  sortWordBankData_(data);
  writeWordBankData_(
    target,
    data,
    previousCount
  );
}
function deleteWordBankEntry_(
  spreadsheet,
  maxDay,
  sheetName,
  word,
  definition
) {
  const target =
    getWordBankSheet_(spreadsheet);
  const data =
    getWordBankData_(target);
  const previousCount = data.length;
  const index = findWordBankEntryIndex_(
    data,
    sheetName,
    word,
    definition
  );
  if (index === -1) {
    rebuildWordBank_(spreadsheet, maxDay);
    return;
  }
  data.splice(index, 1);
  writeWordBankData_(
    target,
    data,
    previousCount
  );
}
function rebuildWordBank_(spreadsheet, maxDay) {
  const target =
    getWordBankSheet_(spreadsheet);
  const data = [];
  getAccessibleDayNames_(spreadsheet, maxDay)
    .forEach(function(sheetName) {
      const sheet =
        spreadsheet.getSheetByName(sheetName);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return;
      }
      const values = sheet
        .getRange(2, 1, lastRow - 1, 2)
        .getValues();
      values.forEach(function(row) {
        const word = cleanText_(row[0]);
        if (word) {
          data.push([
            word,
            row[1] || '',
            sheetName
          ]);
        }
      });
    });
  sortWordBankData_(data);
  const previousCount =
    getWordBankData_(target).length;
  writeWordBankData_(
    target,
    data,
    previousCount
  );
  SpreadsheetApp.flush();
}
/**
 * Manual recovery tool.
 * Not used during normal Add/Edit/Delete requests.
 */
function repairWordBank() {
  const settings = getSettings_();
  rebuildWordBank_(
    settings.spreadsheet,
    settings.maxAccessibleDay
  );
  console.log(
    'A-Z Word Bank repair complete.'
  );
}
function cleanText_(value) {
  return value == null
    ? ''
    : String(value).trim();
}
function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(
      JSON.stringify(payload)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
function errorResponse_(error) {
  return jsonResponse_({
    ok: false,
    error:
      error && error.message
        ? error.message
        : String(error)
  });
}
/**
 * Read-only test: lists the released Days.
 */
function testListAccessibleDays() {
  const testEvent = {
    postData: {
      contents: JSON.stringify({
        action: 'days'
      })
    }
  };
  console.log(
    doPost(testEvent).getContent()
  );
}
