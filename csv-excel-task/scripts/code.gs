function onEdit(e) {
  if (!e || !e.source) {
    Logger.log('No valid event object');
    return;
  }

  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row === 1) {
    Logger.log('Header row edited, skipping');
    return;
  }

  const columns = getColumnIndices(sheet);
  if (!columns) {
    Logger.log('Failed to get column indices');
    return;
  }

  // Clean "Edited" values in Email/Phone columns on any edit
  cleanEditedValues(sheet, columns, row);

  // Optionally: Only validate when relevant columns are edited (e.g., Email/Phone columns)
  const editedColumnIsRelevant = col === columns.emailCol || col === columns.phoneCol;
  if (editedColumnIsRelevant) {
    Logger.log(`Validating row ${row}`);
    validateSingleRow(sheet, row);
    updateTotalErrors(sheet);
  }
}


function validateSheetManually() {
  const sheet = SpreadsheetApp.getActiveSheet();
  validateAllRows(sheet);
}

function getColumnIndices(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Headers: ' + headers.join(', '));

  const nameCol = headers.indexOf("Name") + 1;
  const emailCol = headers.indexOf("Email") + 1;
  const phoneCol = headers.indexOf("Phone") + 1;
  const statusCol = headers.indexOf("status") + 1;
  const errorCountCol = headers.indexOf("errorsCount") + 1;

  if (!nameCol || !emailCol || !phoneCol || !statusCol || !errorCountCol) {
    Logger.log(`Missing required columns: Name=${nameCol}, Email=${emailCol}, Phone=${phoneCol}, Status=${statusCol}, ErrorCount=${errorCountCol}`);
    return null;
  }

  return { nameCol, emailCol, phoneCol, statusCol, errorCountCol };
}

function cleanEditedValues(sheet, columns, row = null) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const rowsToCheck = row ? [row] : Array.from({ length: lastRow - 1 }, (_, i) => i + 2); // Rows 2 to lastRow

  const emailRange = sheet.getRange(2, columns.emailCol, lastRow - 1, 1).getValues();
  const phoneRange = sheet.getRange(2, columns.phoneCol, lastRow - 1, 1).getValues();

  let emailEdited = false;
  let phoneEdited = false;

  rowsToCheck.forEach((r) => {
    const rowIndex = r - 2; // Because emailRange starts from row 2
    if (emailRange[rowIndex][0] === "Edited") {
      emailRange[rowIndex][0] = "";
      emailEdited = true;
    }
    if (phoneRange[rowIndex][0] === "Edited") {
      phoneRange[rowIndex][0] = "";
      phoneEdited = true;
    }
  });

  if (emailEdited) {
    sheet.getRange(2, columns.emailCol, lastRow - 1, 1).setValues(emailRange);
    Logger.log(`Cleared "Edited" flags from Email column`);
  }
  if (phoneEdited) {
    sheet.getRange(2, columns.phoneCol, lastRow - 1, 1).setValues(phoneRange);
    Logger.log(`Cleared "Edited" flags from Phone column`);
  }
}

function validateInputsFromRow(rowData, columns) {
  const name = rowData[columns.nameCol - 1]?.toString().trim() || "";
  const email = rowData[columns.emailCol - 1]?.toString().trim() || "";
  const phone = rowData[columns.phoneCol - 1]?.toString().trim() || "";

  let errors = [];
  if (!name) errors.push("Missing Name");
  if (!email) errors.push("Missing Email");
  if (!phone) errors.push("Missing Phone");
  if (email && !isValidEmail(email)) errors.push("Invalid Email");
  if (phone && !isValidPhone(phone)) errors.push("Invalid Phone");

  return { name, email, phone, errors };
}

function validateSingleRow(sheet, row) {
  const columns = getColumnIndices(sheet);
  if (!columns) {
    Logger.log('Failed to get column indices');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (row === 1) {
    Logger.log('Skipping header row');
    return;
  }

  const lastRowStatus = sheet.getRange(lastRow, columns.statusCol).getValue();
  if (row === lastRow && lastRowStatus === "Total Errors") {
    Logger.log('Skipping Total Errors row');
    return;
  }

  try {
    cleanEditedValues(sheet, columns, row);

    // Batch read row data
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const { errors } = validateInputsFromRow(rowData, columns);

    const status = errors.length === 0 ? "Valid" : errors.join(", ");
    const errorsCount = errors.length;
    const background = errors.length > 0 ? "#ffe0e0" : null;

    // Update status and errors (write back)
    sheet.getRange(row, columns.statusCol, 1, 2).setValues([[status, errorsCount]]);

    // Only update background for the edited row (single row optimization)
    const currentBackground = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getBackgrounds()[0];
    
    // Only set background if it differs
    if (currentBackground.some(bg => bg !== background)) {
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(background);
      Logger.log(`Updated background for row ${row}`);
    } else {
      Logger.log(`No background change needed for row ${row}`);
    }

    Logger.log(`Validated row ${row}`);
  } catch (error) {
    Logger.log(`Error in validateSingleRow for row ${row}: ${error.message}`);
  }
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

    const allValues = sheet.getRange(2, 1, dataLastRow - 1, sheet.getLastColumn()).getValues();

    for (let i = 0; i < allValues.length; i++) {
      const rowData = allValues[i];
      const { errors } = validateInputsFromRow(rowData, columns);
      updates.push([errors.length === 0 ? "Valid" : errors.join(", "), errors.length]);
      backgrounds.push([errors.length > 0 ? "#ffe0e0" : null]);
      Logger.log(`Processed row ${i + 2}`);
    }

    if (updates.length > 0) {
      sheet.getRange(2, columns.statusCol, updates.length, 2).setValues(updates);
      sheet.getRange(2, 1, backgrounds.length, 1).setBackgrounds(backgrounds);
      Logger.log(`Batch updated rows 2 to ${dataLastRow}`);
    }

    Logger.log(`Validated all rows from 2 to ${dataLastRow}`);
  } catch (error) {
    Logger.log(`Error in validateAllRows: ${error.message}`);
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
  const errorCountCol = headers.indexOf("errorsCount") + 1;
  const statusCol = headers.indexOf("status") + 1;
  if (!errorCountCol || !statusCol) {
    Logger.log('errorsCount or status column not found');
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