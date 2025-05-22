const { authClient, sheets, drive } = require('../utils/googlesheet.utils');


const TEMPLATE_SHEET_ID = "1rKT9Q-zZ-vQA6CZNrMd10qFPSwbHlT51EJ6UUc4g_uI";

async function createSheetFromTemplate() {
  if (!TEMPLATE_SHEET_ID) throw new Error("TEMPLATE_SHEET_ID not available");

  const { data } = await drive.files.copy({
    fileId: TEMPLATE_SHEET_ID,
    requestBody: {
      name: `Upload Sheet ${new Date().toISOString()}`,
    },
  });

  const spreadsheetId = data.id;
  const metadata = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties",
  });

  const existingSheets = metadata.data.sheets;
  if (existingSheets.length > 1) {
    const requests = existingSheets.slice(1).map((sheet) => ({
      deleteSheet: {
        sheetId: sheet.properties.sheetId,
      },
    }));
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });
  }

  if (existingSheets.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId: existingSheets[0].properties.sheetId,
                title: "TempSheet",
              },
              fields: "title",
            },
          },
        ],
      },
    });
  }

  await drive.permissions.create({
    fileId: spreadsheetId,
    requestBody: {
      role: "writer",
      type: "anyone",
    },
  });

  return {
    sheetId: spreadsheetId,
    sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
  };
}

async function initializeSchoolSheets(schools) {
  try {
    if (!TEMPLATE_SHEET_ID) throw new Error("TEMPLATE_SHEET_ID not available");
    const { data } = await drive.files.copy({
      fileId: TEMPLATE_SHEET_ID,
      requestBody: {
        name: `Upload Sheet ${new Date().toISOString()}`,
      },
    });

    const spreadsheetId = data.id;
    const metadata = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties",
    });

    const existingSheets = metadata.data.sheets;
    if (existingSheets.length > 1) {
      const requests = existingSheets.slice(1).map((sheet) => ({
        deleteSheet: {
          sheetId: sheet.properties.sheetId,
        },
      }));
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });
    }

    await drive.permissions.create({
      fileId: spreadsheetId,
      requestBody: {
        role: "writer",
        type: "anyone",
      },
    });

    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties.title",
    });

    const existingSheetTitles = response.data.sheets.map((sheet) =>
      sheet.properties.title.toUpperCase()
    );

    const sheetsToCreate = schools.filter(
      (school) => !existingSheetTitles.includes(school.toUpperCase())
    );

    if (sheetsToCreate.length === 0) {
      console.log("All sheets already exist. Proceeding with data upload...");
      return;
    }

    const requests = sheetsToCreate.map((school, index) => ({
      addSheet: {
        properties: {
          title: school,
          index,
          gridProperties: {
            frozenRowCount: 2,
          },
        },
      },
    }));

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });

    console.log("Created sheets:", sheetsToCreate);
    return {
      sheetId: spreadsheetId,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
    };
  } catch (error) {
    console.error("Error initializing sheets:", error.message);
    throw error;
  }
}


async function writeToGoogleSheet(data) {
  const { firstRow, groupedData } = data;
  const firstRowWithBlankCols = [" ", " ", ...firstRow];
  if (!groupedData || Object.keys(groupedData).length === 0) {
    console.warn("No grouped data provided");
    return;
  }

  // Validate school names
  const schoolNames = Object.keys(groupedData).map(
    (school) => groupedData[school].schoolDisplayName || school
  );
  for (const schoolName of schoolNames) {
    if (!schoolName || schoolName.trim() === "") {
      console.error("School name is empty or invalid");
      throw new Error("School name cannot be empty");
    }
  }

  const { sheetId, sheetUrl } = await initializeSchoolSheets(schoolNames);

  for (const schoolName of Object.keys(groupedData)) {
    const schoolData = groupedData[schoolName].rows;
    if (!schoolData || !Array.isArray(schoolData) || schoolData.length === 0) {
      console.warn(`No valid data for school: ${schoolName}`);
      continue;
    }
    const sheetTitle = groupedData[schoolName].schoolDisplayName || schoolName;

    if (!schoolData[0]) {
      console.warn(`No data rows for school: ${schoolName}`);
      continue;
    }
    const headers = Object.keys(schoolData[0]);

    const finalHeaders = [...headers];

    const values = schoolData.map((row) =>
      finalHeaders.map((header) => row[header] ?? "")
    );

    const totalErrors = schoolData.reduce(
      (sum, row) => sum + (row.ErrorsCount || 0),
      0
    );
    const totalRow = Array(finalHeaders.length).fill("");
    totalRow[finalHeaders.indexOf("Status")] = "Total Errors";
    totalRow[finalHeaders.indexOf("ErrorsCount")] = totalErrors;

    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `'${sheetTitle}'!A1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [firstRowWithBlankCols, finalHeaders, ...values, totalRow],
      },
    });

    const sheetIdNum = await getSheetIdByName(sheets, sheetId, sheetTitle);
    if (sheetIdNum !== null) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: {
          requests: [
            {
              autoResizeDimensions: {
                dimensions: {
                  sheetId: sheetIdNum,
                  dimension: "COLUMNS",
                  startIndex: 0,
                  endIndex: Math.max(
                    firstRowWithBlankCols.length,
                    finalHeaders.length
                  ),
                },
              },
            },
          ],
        },
      });
    }

    // Apply formatting and validation
    await applyConditionalFormatting(
      authClient,
      sheets,
      sheetId,
      finalHeaders,
      sheetTitle
    );
    await applyDataValidation(
      authClient,
      sheets,
      sheetId,
      finalHeaders,
      schoolData.length,
      sheetTitle
    );
  }
  return { sheetId, sheetUrl };
}
function isValidEmail(email) {
  return /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(email);
}

async function getSheetIdByName(sheets, spreadsheetId, sheetTitle) {
  const response = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties",
  });
  const sheet = response.data.sheets.find(
    (s) => s.properties.title.toUpperCase() === sheetTitle.toUpperCase()
  );
  return sheet ? sheet.properties.sheetId : null;
}

async function applyConditionalFormatting(
  auth,
  sheets,
  sheetId,
  headers,
  sheetTitle
) {
  const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const sheetIdMap = sheetMeta.data.sheets.reduce((map, sheet) => {
    map[sheet.properties.title.toUpperCase()] = sheet.properties.sheetId;
    return map;
  }, {});

  const sheetIdValue = sheetIdMap[sheetTitle.toUpperCase()];
  if (sheetIdValue === undefined) {
    console.error(`Sheet '${sheetTitle}' not found for conditional formatting`);
    return;
  }

  const requests = [];
  const errorCountColIndex = headers.indexOf("ErrorsCount");
  if (errorCountColIndex !== -1) {
    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `'${sheetTitle}'`,
    });
    const rowCount = sheetData.data.values.length;

    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [
            {
              sheetId: sheetIdValue,
              startRowIndex: 2,
              endRowIndex: rowCount - 1,
              startColumnIndex: 0,
              endColumnIndex: headers.length,
            },
          ],
          booleanRule: {
            condition: {
              type: "CUSTOM_FORMULA",
              values: [
                {
                  userEnteredValue: `=INDIRECT(ADDRESS(ROW(), ${
                    errorCountColIndex + 1
                  }))>0`,
                },
              ],
            },
            format: {
              backgroundColor: { red: 0.984, green: 0.486, blue: 0.486 },
            },
          },
        },
        index: 0,
      },
    });
  }

  if (requests.length) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { requests },
    });
    console.log(`Conditional formatting applied to sheet '${sheetTitle}'`);
  }
}

async function applyDataValidation(
  auth,
  sheets,
  sheetId,
  headers,
  rowCount,
  sheetTitle
) {
  if (!headers || headers.length === 0 || rowCount <= 0) {
    console.warn(
      `Invalid input for data validation: headers=${headers?.length}, rowCount=${rowCount}, sheetTitle=${sheetTitle}`
    );
    return;
  }

  const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const sheetIdMap = sheetMeta.data.sheets.reduce((map, sheet) => {
    map[sheet.properties.title.toUpperCase()] = sheet.properties.sheetId;
    return map;
  }, {});

  const sheetIdValue = sheetIdMap[sheetTitle.toUpperCase()];
  if (sheetIdValue === undefined) {
    console.error(
      `Sheet '${sheetTitle}' not found for data validation in spreadsheet ${sheetId}`
    );
    return;
  }

  const requests = [];
  const endRowIndex = rowCount + 2;

  // Email validation
  const emailColIndex = headers.findIndex((h) =>
    ["CONTACT_EMAIL", "EMAIL", "E-MAIL", "EMAIL_ADDRESS"].includes(
      h.toUpperCase()
    )
  );

  const nameColIndex = headers.findIndex((h) => h.toUpperCase() === "NAME");
  if (nameColIndex !== -1) {
    requests.push({
      setDataValidation: {
        range: {
          sheetId: sheetIdValue,
          startRowIndex: 2,
          endRowIndex,
          startColumnIndex: nameColIndex,
          endColumnIndex: nameColIndex + 1,
        },
        rule: {
          condition: {
            type: "CUSTOM_FORMULA",
            values: [
              {
                userEnteredValue: '=NOT(ISBLANK(INDIRECT("R[0]C[0]", FALSE)))',
              },
            ],
          },
          inputMessage: "Name field cannot be empty.",
          strict: true,
          showCustomUi: true,
        },
      },
    });
  }

  const phoneColIndex = headers.findIndex((h) => h.toUpperCase() === "PHONE");
  if (phoneColIndex !== -1) {
    requests.push({
      setDataValidation: {
        range: {
          sheetId: sheetIdValue,
          startRowIndex: 2,
          endRowIndex,
          startColumnIndex: phoneColIndex,
          endColumnIndex: phoneColIndex + 1,
        },
        rule: {
          condition: {
            type: "CUSTOM_FORMULA",
            values: [
              {
                userEnteredValue: `=REGEXMATCH(TO_TEXT(INDIRECT(ADDRESS(ROW(), COLUMN()))), "^((\\+91[-\\s]?)?[6-9][0-9]{9})$")`,
              },
            ],
          },
          inputMessage: "Enter a valid 10-digit phone number.",
          strict: false,
          showCustomUi: true,
        },
      },
    });
  }

  if (emailColIndex !== -1) {
    requests.push({
      setDataValidation: {
        range: {
          sheetId: sheetIdValue,
          startRowIndex: 2,
          endRowIndex,
          startColumnIndex: emailColIndex,
          endColumnIndex: emailColIndex + 1,
        },
        rule: {
          condition: {
            type: "CUSTOM_FORMULA",
            values: [
              {
                userEnteredValue:
                  '=REGEXMATCH(INDIRECT("R[0]C[0]", FALSE), "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$")',
              },
            ],
          },
          inputMessage: "Must be a valid email (e.g., user@domain.com)",
          strict: false,
          showCustomUi: true,
        },
      },
    });
  }

  if (requests.length) {
    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: { requests },
      });
      console.log(`Data validation applied to sheet '${sheetTitle}'`);
    } catch (error) {
      console.error(
        `Failed to apply data validation to sheet '${sheetTitle}': ${error.message}`
      );
    }
  }
}

async function getSheetData(sheetId, sheetTitle) {

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${sheetTitle}'`,
  });

  const [headers, ...rows] = res.data.values || [[]];
  const data = rows.map((row) => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] || "";
    });
    return obj;
  });

  return { data, headers };
}

async function revalidateSheetData(sheetId) {
  const response = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    fields: "sheets.properties.title",
  });

  const sheetTitles = response.data.sheets.map(
    (sheet) => sheet.properties.title
  );
  const groupedData = {};

  for (const sheetTitle of sheetTitles) {
    if (sheetTitle === "TempSheet") continue;
    const { data } = await getSheetData(sheetId, sheetTitle);
    groupedData[sheetTitle] = {
      schoolDisplayName: sheetTitle,
      rows: data,
    };
  }

  await writeToGoogleSheet(sheetId, { firstRow: ["Metadata"], groupedData });
}


module.exports = {
  createSheetFromTemplate,
  writeToGoogleSheet,
  getSheetData,
  revalidateSheetData,
};
