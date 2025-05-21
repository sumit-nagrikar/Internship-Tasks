const { google } = require("googleapis");
const fs = require("fs").promises;
const {
  getAuthenticatedClient,
  getSheetsClient,
} = require("../utils/googlesheet.utils");

const TEMPLATE_SHEET_ID = "1rKT9Q-zZ-vQA6CZNrMd10qFPSwbHlT51EJ6UUc4g_uI";

const dropdownValues = [
  "ORG_CODE",
  "ORG_NAME",
  "ORG_NAME_L",
  "ORG_ADDRESS",
  "ORG_CITY",
  "ORG_STATE",
  "ORG_PIN",
  "ACADEMIC_COURSE_ID",
  "COURSE_NAME",
  "COURSE_NAME_L",
  "COURSE_SUBTITLE",
  "ADMISSION_YEAR",
  "INTR_COURSE_NAME_FIRST",
  "INTR_COURSE_NAME_SECOND",
  "DEPARTMENT",
  "STREAM",
  "STREAM_L",
  "STREAM_SECOND",
  "STREAM_SECOND_L",
  "SESSION",
  "SPECIALIZATION_MAJOR",
  "SPECIALIZATION_MINOR",
  "REGN_NO",
  "RROLL",
  "ABC_ACCOUNT_ID",
  "CNAME",
  "AADHAAR_NAME",
  "DOB",
  "GENDER",
  "CASTE",
  "RELIGION",
  "NATIONALITY",
  "BLOOD_GROUP",
  "PH",
  "MOBILE",
  "EMAIL",
  "FNAME",
  "MNAME",
  "GNAME",
  "STUDENT_ADDRESS",
  "PHOTO",
  "MRKS_REC_STATUS",
  "RESULT",
  "RESULT_TH",
  "RESULT_PR",
  "YEAR",
  "MONTH",
  "DIVISION",
  "GRADE",
  "PERCENT",
  "DOR",
  "DOI",
  "DOV",
  "DOE",
  "DOP",
  "DOQ",
  "DOS",
  "THESIS",
  "REMARKS",
  "CERT_NO",
  "MEDIUM",
  "SEM",
  "CENTRE_NAME",
  "EXAM_TYPE",
  "TERM_TYPE",
  "TOT",
  "TOT_MIN",
  "TOT_MRKS",
  "TOT_MRKS_WRDS",
  "TOT_MRKS_MIN",
  "TOT_TH_MAX",
  "TOT_TH_MIN",
  "TOT_TH_MRKS",
  "TOT_PR_MAX",
  "TOT_PR_MIN",
  "TOT_PR_MRKS",
  "TOT_CE_MAX",
  "TOT_CE_MIN",
  "TOT_CE_MRKS",
  "TOT_VV_MAX",
  "TOT_VV_MIN",
  "TOT_VV_MRKS",
  "TOT_PR_CE_MAX",
  "TOT_PR_CE_MRKS",
  "TOT_TH_CE_MAX",
  "TOT_TH_CE_MRKS",
  "PREV_TOT_MRKS",
  "PREV_TOT_MRKS_MAX",
  "PREV_TOT_MRKS_MIN",
  "GRAND_TOT_MAX",
  "GRAND_TOT_MIN",
  "GRAND_TOT_MRKS",
  "GRAND_TOT_GRADE_POINTS",
  "GRAND_TOT_CREDIT",
  "GRAND_TOT_CREDIT_POINTS",
  "TOT_GRADE",
  "TOT_GRADE_POINTS",
  "TOT_CREDIT",
  "TOT_CREDIT_POINTS",
  "CGPA",
  "TOT_CGPA_MINOR",
  "TOT_CREDIT_MINOR",
  "FINAL_GMAX_TOTAL",
  "CGPA_SCALE",
  "GPA",
  "SGPA",
  "INTR_CGPA_FIRST",
  "INTR_CGPA_SECOND",
  "SUB1NM",
  "SUB1",
  "SUB1MAX",
  "SUB1MIN",
  "SUB1_SESSION",
  "SUB1_TH_MAX",
  "SUB1_TH_MIN",
  "SUB1_PR_MAX",
  "SUB1_PR_MIN",
  "SUB1_CE_MAX",
  "SUB1_CE_MIN",
  "SUB1_VV_MAX",
  "SUB1_VV_MIN",
  "SUB1_VV_GRADE",
  "SUB1_TH_MRKS",
  "SUB1_TH_CE_MAX",
  "SUB1_TH_CE_MRKS",
  "SUB1_TH_GRADE",
  "SUB1_TH_AGGREGATE",
  "SUB1_PR_AGGREGATE",
  "SUB1_PR_MRKS",
  "SUB1_PR_GRADE",
  "SUB1_PR_CE_MAX",
  "SUB1_PR_CE_MRKS",
  "SUB1_PR_HOURS",
  "SUB1_TH_HOURS",
  "SUB1_TT_HOURS",
  "SUB1_CE_WEIGHT_MRKS",
  "SUB1_CE_MRKS",
  "SUB1_CE_GRADE",
  "SUB1_CE1_MRKS",
  "SUB1_CE1_GRADE",
  "SUB1_CE2_MRKS",
  "SUB1_CE2_GRADE",
  "SUB1_CE3_MRKS",
  "SUB1_CE3_GRADE",
  "SUB1_CE4_MRKS",
  "SUB1_CE4_GRADE",
  "SUB1_VV_MRKS",
  "SUB1_PAPER1_MRKS",
  "SUB1_PAPER2_MRKS",
  "SUB1_PAPER3_MRKS",
  "SUB1_PAPER4_MRKS",
  "SUB1_PAPER1_PR_MRKS",
  "SUB1_PAPER2_PR_MRKS",
  "SUB1_PAPER3_PR_MRKS",
  "SUB1_PAPER1_CE_MRKS",
  "SUB1_PAPER2_CE_MRKS",
  "SUB1_PAPER3_CE_MRKS",
  "SUB1_PAPER1_MRKS_SH",
  "SUB1_PAPER1_CE_MRKS_SH",
  "SUB1_PAPER1_PR_MRKS_SH",
  "SUB1_PAPER2_MRKS_SH",
  "SUB1_PAPER2_CE_MRKS_SH",
  "SUB1_PAPER2_PR_MRKS_SH",
  "SUB1_PAPER3_MRKS_SH",
  "SUB1_PAPER3_CE_MRKS_SH",
  "SUB1_PAPER3_PR_MRKS_SH",
  "SUB1_MAX_MRKS_SH",
  "SUB1_MAX_CE_MRKS_SH",
  "SUB1_MAX_PR_MRKS_SH",
  "SUB1_MIN_MRKS_SH",
  "SUB1_MIN_CE_MRKS_SH",
  "SUB1_MIN_PR_MRKS_SH",
  "SUB1_LAB1_MRKS",
  "SUB1_LAB2_MRKS",
  "SUB1_LAB3_MRKS",
  "SUB1_LAB4_MRKS",
  "SUB1_LAB1_GRADE",
  "SUB1_LAB2_GRADE",
  "SUB1_LAB3_GRADE",
  "SUB1_LAB4_GRADE",
  "SUB1_REPORT_MRKS",
  "SUB1_REPORT_GRADE",
  "SUB1_PRO_MRKS",
  "SUB1_PRO_CE_MRKS",
  "SUB1_TEE_PR_MRKS",
  "SUB1_TEE_TH_MRKS",
  "SUB1_TEE_PR_GRADE",
  "SUB1_TEE_TH_GRADE",
  "SUB1_TEE_WEIGHT_MRKS",
  "SUB1_TYPE",
  "SUB1_TOT",
  "SUB1_CE_TOT",
  "SUB1_PR_TOT",
  "SUB1_REMARKS",
  "SUB1_STATUS",
  "SUB1_GRADE",
  "SUB1_GRADE_POINTS",
  "SUB1_CREDIT",
  "SUB1_CREDIT_POINTS",
  "SUB1_CREDIT_ELIGIBILITY",
  "SUB1_CREDIT_HOURS",
  "SUB1_PAPER1_STATUS",
  "SUB1_PAPER2_STATUS",
  "SUB1_PAPER3_STATUS",
  "SUB1_PAPER4_STATUS",
  "SUB1_GRACE",
  "SUB1_GROUP",
  "SUB1_GROUP_CODE",
  "SUB1_GROUP_MAX",
  "SUB1_GROUP_MIN",
  "SUB1_GROUP_TOT",
  "NON_CREDIT_HOURS",
];

async function createSheetFromTemplate() {
  if (!TEMPLATE_SHEET_ID) throw new Error("TEMPLATE_SHEET_ID not available");

  const auth = await getAuthenticatedClient();
  const drive = google.drive({ version: "v3", auth });
  const sheets = getSheetsClient(auth);

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
    sheetUrl: `htps://docs.googtle.com/spreadsheets/d/${spreadsheetId}`,
  };
}

async function initializeSchoolSheets(schools) {
  try {
    if (!TEMPLATE_SHEET_ID) throw new Error("TEMPLATE_SHEET_ID not available");

    const auth = await getAuthenticatedClient();
    const drive = google.drive({ version: "v3", auth });
    const sheets = getSheetsClient(auth);

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
  const auth = await getAuthenticatedClient();
  const sheets = getSheetsClient(auth);
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
      auth,
      sheets,
      sheetId,
      finalHeaders,
      sheetTitle
    );
    await applyDataValidation(
      auth,
      sheets,
      sheetId,
      finalHeaders,
      schoolData.length,
      sheetTitle
    );
  }
  return { sheetId, sheetUrl };
}

async function writeToGoogleSheetNew(sheetId, data, sheetTitle = "Sheet1") {
  const auth = await getAuthenticatedClient();
  const sheets = getSheetsClient(auth);

  let displayHeaders = data.headers[0] || [];
  let fieldKeys = data.headers[1] || [];

  const systemFields = [
    { key: "Status", label: "Status" },
    { key: "ErrorsCount", label: "ErrorsCount" },
  ];

  systemFields.forEach(({ key }) => {
    const index = fieldKeys.indexOf(key);
    if (index !== -1) {
      fieldKeys.splice(index, 1);
      displayHeaders.splice(index, 1);
    }
  });

  fieldKeys = systemFields.map((f) => f.key).concat(fieldKeys);
  displayHeaders = systemFields.map((f) => f.label).concat(displayHeaders);

  const allRows = Object.values(data.groupedData).flatMap(
    (group) => group.rows
  );
  const dataRows = allRows.map((row) => fieldKeys.map((key) => row[key] ?? ""));

  const totalErrors = allRows.reduce(
    (sum, row) => sum + (row.ErrorsCount || 0),
    0
  );
  const totalRow = Array(fieldKeys.length).fill("");
  totalRow[fieldKeys.indexOf("Status")] = "Total Errors";
  totalRow[fieldKeys.indexOf("ErrorsCount")] = totalErrors;

  const values = [displayHeaders, fieldKeys, ...dataRows, totalRow];

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `'${sheetTitle}'!A1`,
    valueInputOption: "RAW",
    requestBody: { values },
  });

  const sheetIdNum = await getSheetIdByTitle(sheets, sheetId, sheetTitle);

  const errorColor = { red: 1, green: 0.8, blue: 0.8 };

  const errorHighlightRequests = dataRows
    .map((row, i) => {
      const errorCount = parseInt(row[fieldKeys.indexOf("ErrorsCount")]) || 0;
      const rowIndex = i + 2;
      if (errorCount > 0) {
        return {
          repeatCell: {
            range: {
              sheetId: sheetIdNum,
              startRowIndex: rowIndex,
              endRowIndex: rowIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: errorColor,
              },
            },
            fields: "userEnteredFormat.backgroundColor",
          },
        };
      }
      return null;
    })
    .filter(Boolean);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: {
      requests: [
        {
          updateDimensionProperties: {
            range: {
              sheetId: sheetIdNum,
              dimension: "ROWS",
              startIndex: 0,
              endIndex: values.length,
            },
            properties: {
              pixelSize: 30,
            },
            fields: "pixelSize",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId: sheetIdNum,
              startRowIndex: 1,
              endRowIndex: 2,
            },
            cell: {
              dataValidation: null,
            },
            fields: "dataValidation",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId: sheetIdNum,
              startRowIndex: 0,
              endRowIndex: 2,
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  bold: true,
                },
              },
            },
            fields: "userEnteredFormat.textFormat.bold",
          },
        },
        ...errorHighlightRequests,
      ],
    },
  });

  const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;
  console.log(`Validations Added to Spreadsheet : ${sheetUrl}`);
  return { sheetId, sheetUrl };
}

async function getSheetIdByTitle(sheets, spreadsheetId, title) {
  const metadata = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = metadata.data.sheets.find((s) => s.properties.title === title);
  return sheet?.properties.sheetId;
}

async function buildAndCreateSheetFromParsedData(
  parsedData,
  sheetTitle = "Sheet1"
) {
  const auth = await getAuthenticatedClient();
  const sheets = getSheetsClient(auth);
  const drive = google.drive({ version: "v3", auth });

  const firstRow = parsedData[0];
  const valueRows = parsedData.slice(1);

  const firstRowClean = [...firstRow];

  const createResponse = await sheets.spreadsheets.create({
    requestBody: {
      properties: {
        title: "Student Data Sheet",
      },
    },
  });

  const sheetId = createResponse.data.spreadsheetId;

  if (sheetTitle !== "Sheet1") {
    const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    const defaultSheet = sheetMeta.data.sheets.find(
      (s) => s.properties.sheetId === 0 || s.properties.title === "Sheet1"
    );
    const defaultSheetId = defaultSheet.properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId: defaultSheetId,
                title: sheetTitle,
              },
              fields: "title",
            },
          },
        ],
      },
    });
  }

  const sheetMetaFinal = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
  });
  const sheetObj = sheetMetaFinal.data.sheets.find(
    (s) => s.properties.title === sheetTitle
  );
  const sheetIdNum = sheetObj?.properties?.sheetId;

  const dropdownRow = Array(firstRowClean.length).fill("");
  for (let i = 0; i < firstRowClean.length; i++) {
    dropdownRow[i] = dropdownValues[i % dropdownValues.length];
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `'${sheetTitle}'!A1`,
    valueInputOption: "RAW",
    requestBody: {
      values: [firstRowClean, dropdownRow, ...valueRows],
    },
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: {
              sheetId: sheetIdNum,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: 0,
              endColumnIndex: firstRowClean.length,
            },
            cell: {
              dataValidation: {
                condition: {
                  type: "ONE_OF_LIST",
                  values: dropdownValues.map((val) => ({
                    userEnteredValue: val,
                  })),
                },
                showCustomUi: true,
                strict: true,
              },
            },
            fields: "dataValidation",
          },
        },
        {
          autoResizeDimensions: {
            dimensions: {
              sheetId: sheetIdNum,
              dimension: "COLUMNS",
              startIndex: 0,
              endIndex: firstRowClean.length,
            },
          },
        },
        {
          repeatCell: {
            range: {
              sheetId: sheetIdNum,
              startRowIndex: 0,
              endRowIndex: 2,
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  bold: true,
                },
              },
            },
            fields: "userEnteredFormat.textFormat.bold",
          },
        },
      ],
    },
  });

  // 7️⃣ Make sheet public (optional)
  await drive.permissions.create({
    fileId: sheetId,
    requestBody: {
      role: "writer",
      type: "anyone",
    },
  });

  const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;
  console.log(`Sheet created: ${sheetUrl}`);
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
  const sheetData = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${sheetTitle}'`,
  });
  const rowCount = sheetData.data.values.length;
  if (errorCountColIndex !== -1) {
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

  requests.push({
    updateDimensionProperties: {
      range: {
        sheetId: sheetIdValue,
        dimension: "ROWS",
        startIndex: 2,
        endIndex: rowCount - 1,
      },
      properties: {
        pixelSize: 30,
      },
      fields: "pixelSize",
    },
  });

  const dropdownColumnStart = 2;
  const dropdownColumnEnd = headers.length;

  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetIdValue,
        startRowIndex: 1,
        endRowIndex: 2,
        startColumnIndex: dropdownColumnStart,
        endColumnIndex: dropdownColumnEnd,
      },
      cell: {
        dataValidation: {
          condition: {
            type: "ONE_OF_LIST",
            values: dropdownValues.map((value) => ({
              userEnteredValue: value,
            })),
          },
          strict: true,
          showCustomUi: true,
        },
      },
      fields: "dataValidation",
    },
  });

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
  const auth = await getAuthenticatedClient();
  const sheets = getSheetsClient(auth);

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
  const auth = await getAuthenticatedClient();
  const sheets = getSheetsClient(auth);

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

async function attachScriptToSheet(sheetId) {
  try {
    const auth = await getAuthenticatedClient();
    const script = google.script({ version: "v1", auth });

    console.log("Reading code.gs file...");
    const code = await fs.readFile("scripts/code.gs", "utf8");

    console.log("Creating script project for sheet:", sheetId);
    const { data: scriptProject } = await script.projects.create({
      requestBody: {
        title: `Script for Sheet ${sheetId}`,
        parentId: sheetId,
      },
    });
    console.log("Script project created:", scriptProject.scriptId);

    console.log("Uploading script content...");
    await script.projects.updateContent({
      scriptId: scriptProject.scriptId,
      requestBody: {
        files: [
          {
            name: "Code",
            type: "SERVER_JS",
            source: code,
          },
          {
            name: "appsscript",
            type: "JSON",
            source: JSON.stringify({
              timeZone: "Asia/Kolkata",
              exceptionLogging: "STACKDRIVER",
              runtimeVersion: "V8",
            }),
          },
        ],
      },
    });

    await new Promise((resolve) => setTimeout(resolve, 2000));

    console.log(
      `Script injected successfully into spreadsheet: https://docs.google.com/spreadsheets/d/${sheetId}`
    );
    return scriptProject.scriptId;
  } catch (error) {
    console.error("Error injecting script:", error.message);
    throw error;
  }
}

module.exports = {
  createSheetFromTemplate,
  writeToGoogleSheet,
  getSheetData,
  revalidateSheetData,
  attachScriptToSheet,
  buildAndCreateSheetFromParsedData,
  writeToGoogleSheetNew,
};
