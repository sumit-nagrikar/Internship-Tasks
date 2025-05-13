function onEdit(e) {
  if (!e || !e.source) {
    Logger.log('No valid event object');
    return;
  }

  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();

  if (row === 1) {
    Logger.log('Header row edited, skipping');
    return;
  }

  validateAllRows(sheet);
  updateTotalErrors(sheet);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Validator")
    .addItem("Validate All Rows", "validateSheetManually")
    .addToUi();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  validateAllRows(sheet);
  updateTotalErrors(sheet);
}

function validateSheetManually() {
  const sheet = SpreadsheetApp.getActiveSheet();
  validateAllRows(sheet);
}

function getColumnIndices(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Headers: ' + headers.join(', '));

  // Case-insensitive header matching
  const nameCol = headers.findIndex(h => h.toString().toUpperCase() === 'NAME') + 1;
  const emailCol = headers.findIndex(h => h.toString().toUpperCase() === 'EMAIL') + 1;
  const phoneCol = headers.findIndex(h => h.toString().toUpperCase() === 'PHONE') + 1;
  let statusCol = headers.findIndex(h => h.toString().toUpperCase() === 'STATUS') + 1;
  let errorCountCol = headers.findIndex(h => h.toString().toUpperCase() === 'ERRORSCOUNT') + 1;

  // Append status and errorsCount columns if missing
  if (!statusCol) {
    statusCol = headers.length + 1;
    sheet.getRange(1, statusCol).setValue('status');
    Logger.log('Added missing status column at position ' + statusCol);
  }
  if (!errorCountCol) {
    errorCountCol = headers.length + (statusCol === headers.length + 1 ? 2 : 1);
    sheet.getRange(1, errorCountCol).setValue('errorsCount');
    Logger.log('Added missing errorsCount column at position ' + errorCountCol);
  }

  if (!nameCol || !emailCol || !phoneCol || !statusCol || !errorCountCol) {
    const missing = [];
    if (!nameCol) missing.push('NAME');
    if (!emailCol) missing.push('EMAIL');
    if (!phoneCol) missing.push('PHONE');
    if (!statusCol) missing.push('STATUS');
    if (!errorCountCol) missing.push('ERRORSCOUNT');
    Logger.log(`Missing required columns: ${missing.join(', ')}`);
    SpreadsheetApp.getUi().alert('Validation Error', `Missing required columns: ${missing.join(', ')}`, SpreadsheetApp.ButtonSet.OK);
    return null;
  }

  return { nameCol, emailCol, phoneCol, statusCol, errorCountCol };
}

function cleanEditedValues(sheet, columns) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const ranges = [
    sheet.getRange(2, columns.emailCol, lastRow - 1, 1),
    sheet.getRange(2, columns.phoneCol, lastRow - 1, 1)
  ];

  ranges.forEach(range => {
    const values = range.getValues();
    let editedFound = false;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === "Edited") {
        values[i][0] = "";
        editedFound = true;
      }
    }
    if (editedFound) {
      range.setValues(values);
      Logger.log(`Cleared "Edited" from ${range.getA1Notation()}`);
    }
  });
}

function validateInputs(sheet, row, columns) {
  const name = sheet.getRange(row, columns.nameCol).getValue().toString().trim();
  const email = sheet.getRange(row, columns.emailCol).getValue().toString().trim();
  const phone = sheet.getRange(row, columns.phoneCol).getValue().toString().trim();

  let errors = [];
  if (!name) errors.push("Missing Name");
  if (!email) errors.push("Missing Email");
  if (!phone) errors.push("Missing Phone");
  if (email && !isValidEmail(email)) errors.push("Invalid Email");
  if (phone && !isValidPhone(phone)) errors.push("Invalid Phone");

  Logger.log(`Row ${row}: Name=${name}, Email=${email}, Phone=${phone}, Errors=${errors.join(", ") || "none"}`);
  return { name, email, phone, errors };
}

function validateAllRows(sheet) {
  const columns = getColumnIndices(sheet);
  if (!columns) {
    Logger.log('Failed to get column indices');
    return;
  }

  cleanEditedValues(sheet, columns);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('No data rows to validate');
    return;
  }

  try {
    const updates = [];
    const backgrounds = [];
    let dataLastRow = lastRow;

    const lastRowStatus = sheet.getRange(lastRow, columns.statusCol).getValue();
    if (lastRowStatus === "Total Errors") {
      dataLastRow = lastRow - 1;
    }

    for (let row = 2; row <= dataLastRow; row++) {
      const { errors } = validateInputs(sheet, row, columns);
      updates.push([errors.length === 0 ? "Valid" : errors.join(", "), errors.length]);
      backgrounds.push(errors.length > 0 ? "#ffe0e0" : null);
      Logger.log(`Processed row ${row}`);
    }

    if (updates.length > 0) {
      sheet.getRange(2, columns.statusCol, dataLastRow - 1, 2).setValues(updates);
      sheet.getRange(2, 1, dataLastRow - 1, sheet.getLastColumn()).setBackgrounds(backgrounds.map(bg => [bg]));
      Logger.log(`Batch updated rows 2 to ${dataLastRow}`);
    }

    updateTotalErrors(sheet);
    Logger.log(`Validated all rows from 2 to ${dataLastRow}`);
  } catch (error) {
    Logger.log(`Error in validateAllRows: ${error.message}`);
    SpreadsheetApp.getUi().alert('Validation Error', `Failed to validate rows: ${error.message}`, SpreadsheetApp.ButtonSet.OK);
  }
}

function updateStatusAndErrors(sheet, row, columns, errors) {
  const updates = [[errors.length === 0 ? "Valid" : errors.join(", "), errors.length]];
  sheet.getRange(row, columns.statusCol, 1, 2).setValues(updates);
  Logger.log(`Row ${row}: Status updated to "${updates[0][0]}", errorsCount to ${updates[0][1]}`);
}

function setRowBackground(sheet, row, errors) {
  const background = errors.length > 0 ? "#ffe0e0" : null;
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(background);
  Logger.log(`Row ${row}: Background set to ${background || "null"}`);
}

function isValidEmail(email) {
  const isValid = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(email);
  Logger.log(`Validating email "${email}": ${isValid}`);
  return isValid;
}

function isValidPhone(phone) {
  const isValid = /^((\+91[-\s]?)?[6-9]\d{9})$/.test(phone);
  Logger.log(`Validating phone "${phone}": ${isValid}`);
  return isValid;
}

function updateTotalErrors(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const errorCountCol = headers.findIndex(h => h.toString().toUpperCase() === 'ERRORSCOUNT') + 1;
  const statusCol = headers.findIndex(h => h.toString().toUpperCase() === 'STATUS') + 1;
  if (!errorCountCol || !statusCol) {
    Logger.log('errorsCount or status column not found');
    SpreadsheetApp.getUi().alert('Validation Error', 'Missing status or errorsCount column', SpreadsheetApp.ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  let dataLastRow = lastRow;

  // Robustly find dataLastRow by checking for Total Errors
  if (lastRow > 1) {
    for (let row = lastRow; row >= 2; row--) {
      const status = sheet.getRange(row, statusCol).getValue();
      if (status === "Total Errors") {
        sheet.getRange(row, statusCol, 1, 2).clearContent();
        dataLastRow = row - 1;
        break;
      }
    }
  }

  // Calculate total errors
  let totalErrors = 0;
  if (dataLastRow > 1) {
    const errorCounts = sheet.getRange(2, errorCountCol, dataLastRow - 1, 1).getValues();
    totalErrors = errorCounts.reduce((sum, [count]) => {
      const num = Number(count);
      Logger.log(`Row ${2 + errorCounts.indexOf([count])}: errorsCount=${count}, converted=${num}`);
      return sum + (isNaN(num) ? 0 : num);
    }, 0);
  }

  // Write total errors
  const totalRow = dataLastRow + 1;
  sheet.getRange(totalRow, statusCol, 1, 2).setValues([["Total Errors", totalErrors]]);
  Logger.log(`Total errors updated: ${totalErrors} in row ${totalRow}, column ${statusCol}:${errorCountCol}, dataLastRow=${dataLastRow}`);
}