const { google } = require('googleapis');
const fs = require('fs').promises;
const { validateRows } = require('./fileParser.service');
const { getAuthenticatedClient, getSheetsClient } = require('../utils/googlesheet.utils');

const TEMPLATE_SHEET_ID = '1rKT9Q-zZ-vQA6CZNrMd10qFPSwbHlT51EJ6UUc4g_uI';

async function createSheetFromTemplate() {
    if (!TEMPLATE_SHEET_ID) throw new Error("TEMPLATE_SHEET_ID not available");

    const auth = await getAuthenticatedClient();
    const drive = google.drive({ version: 'v3', auth });

    const { data } = await drive.files.copy({
        fileId: TEMPLATE_SHEET_ID,
        requestBody: {
            name: `Upload Sheet ${new Date().toISOString()}`,
        },
    });

    await drive.permissions.create({
        fileId: data.id,
        requestBody: {
            role: 'writer',
            type: 'anyone',
        },
    });

    return {
        sheetId: data.id,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${data.id}`,
    };
}


async function initializeSchoolSheets(auth, sheets, spreadsheetId, schools) {
  try {
    // Fetch existing sheets
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties.title",
    });

    const existingSheets = response.data.sheets.map((sheet) =>
      sheet.properties.title.toUpperCase()
    );
    console.log("Existing sheets:", existingSheets);

    // Filter out schools that already have sheets (case-insensitive)
    const sheetsToCreate = schools.filter(
      (school) => !existingSheets.includes(school.toUpperCase())
    );

    if (sheetsToCreate.length === 0) {
      console.log("All sheets already exist. Proceeding with data upload...");
      return;
    }

    // Create batch update request for new sheets
    const requests = sheetsToCreate.map((school) => ({
      addSheet: {
        properties: {
          title: school,
          gridProperties: {
            frozenRowCount: 2, // Freeze first row (metadata) and second row (headers)
          },
        },
      },
    }));

    // Execute batch update to create new sheets
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: { requests },
    });

    console.log("Created sheets:", sheetsToCreate);
  } catch (error) {
    console.error("Error initializing sheets:", error.message);
    throw error;
  }
}

async function writeToGoogleSheet(sheetId, data) {
    const auth = await getAuthenticatedClient();
    const sheets = getSheetsClient(auth);

    if (!data || data.length === 0) return;

    // Step 1: Normalize fields
    const fieldAliases = {
        Name: ['name', 'FIRST_NAME'],
        Email: ['email', 'CONTACT_EMAIL'],
        Phone: ['phone', 'mobile', 'CONTACT_NUMBER'],
    };

    function normalizeKey(row, fieldName) {
        const aliases = fieldAliases[fieldName];
        for (let alias of aliases) {
            const match = Object.keys(row).find(k => k.toLowerCase().includes(alias));
            if (match) return row[match];
        }
        return '';
    }

    // Step 2: Process and normalize rows
    const validatedData = data.map(row => {
        const name = normalizeKey(row, 'Name');
        const email = normalizeKey(row, 'Email');
        const phone = normalizeKey(row, 'Phone');

        const normalized = {
            Name: String(name).trim(),
            Email: String(email).trim(),
            Phone: String(phone).trim(),
        };

        // Add any extra fields (preserve as-is)
        const extraFields = {};
        for (let key of Object.keys(row)) {
            const lowerKey = key.toLowerCase();
            if (!['name', 'full name', 'email', 'email address', 'phone', 'mobile', 'contact'].includes(lowerKey)) {
                extraFields[key] = row[key];
            }
        }

        // Validation
        let errors = [];
        if (!normalized.Name) errors.push("Missing Name");
        if (!normalized.Email) errors.push("Missing Email");
        if (!normalized.Phone) errors.push("Missing Phone");
        if (normalized.Email && !isValidEmail(normalized.Email)) errors.push("Invalid Email");
        if (normalized.Phone && !isValidPhone(normalized.Phone)) errors.push("Invalid Phone");

        return {
            ...normalized,
            ...extraFields,
            status: errors.length === 0 ? "Valid" : errors.join(", "),
            errorsCount: errors.length,
        };
    });

    // Step 3: Collect all headers dynamically
    const allHeaders = new Set();
    validatedData.forEach(obj => Object.keys(obj).forEach(key => allHeaders.add(key)));

    // Fix header order: Known → Extra → Final
    const knownHeaders = ["Name", "Email", "Phone"];
    const extraHeaders = [...allHeaders].filter(h => !knownHeaders.includes(h) && h !== "status" && h !== "errorsCount");
    const finalHeaders = [...knownHeaders, ...extraHeaders, "status", "errorsCount"];

    // Step 4: Prepare values
    const values = validatedData.map(obj => finalHeaders.map(h => obj[h] ?? ""));

    // Step 5: Total Errors Row
    const totalErrors = validatedData.reduce((sum, row) => sum + row.errorsCount, 0);
    const totalRow = Array(finalHeaders.length).fill("");
    totalRow[finalHeaders.indexOf("status")] = "Total Errors";
    totalRow[finalHeaders.indexOf("errorsCount")] = totalErrors;

    // Step 6: Write to Sheet
    await sheets.spreadsheets.values.append({
        spreadsheetId: sheetId,
        range: 'Sheet1!A1',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
            values: [finalHeaders, ...values, totalRow],
        },
    });

    await applyConditionalFormatting(sheetId, finalHeaders);
    await applyDataValidation(sheetId, finalHeaders, data.length);

}


function isValidEmail(email) {
    return /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(email);
}

function isValidPhone(phone) {
    return /^((\+91[-\s]?)?[6-9][0-9]{9})$/.test(phone);
}

async function applyConditionalFormatting(sheetId, headers) {
    const auth = await getAuthenticatedClient();
    const sheets = getSheetsClient(auth);

    const statusColIndex = headers.findIndex(h => h === "status");
    if (statusColIndex === -1) return;

    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: {
            requests: [
                {
                    addConditionalFormatRule: {
                        rule: {
                            ranges: [
                                {
                                    sheetId: 0,
                                    startRowIndex: 2,
                                    startColumnIndex: 0,
                                    endColumnIndex: headers.length,
                                },
                            ],
                            booleanRule: {
                                condition: {
                                    type: 'CUSTOM_FORMULA',
                                    values: [
                                        {
                                            userEnteredValue: `=OR(ISNUMBER(SEARCH("Invalid", INDIRECT(ADDRESS(ROW(), ${statusColIndex + 1})))), ISNUMBER(SEARCH("Missing", INDIRECT(ADDRESS(ROW(), ${statusColIndex + 1})))))`,
                                        },
                                    ],
                                },
                                format: {
                                    backgroundColor: { red: 1, green: 0.88, blue: 0.88 },
                                },
                            },
                        },
                        index: 0,
                    },
                },
            ],
        },
    });
}

async function applyDataValidation(sheetId, headers, rowCount) {
    const auth = await getAuthenticatedClient();
    const sheets = getSheetsClient(auth);

    const requests = [];
    const endRowIndex = rowCount + 1;

    headers.forEach((header, colIndex) => {
        if (header === "Email") {
            requests.push({
                setDataValidation: {
                    range: {
                        sheetId: 0,
                        startRowIndex: 2,
                        endRowIndex,
                        startColumnIndex: colIndex,
                        endColumnIndex: colIndex + 1,
                    },
                    rule: {
                        condition: {
                            type: 'CUSTOM_FORMULA',
                            values: [
                                {
                                    userEnteredValue: `=REGEXMATCH(INDIRECT(ADDRESS(ROW(), COLUMN())), "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$")`,
                                },
                            ],
                        },
                        inputMessage: 'Enter a valid email address.',
                        strict: false,
                        showCustomUi: true,
                    },
                },
            });
        }

        if (header === "Phone") {
            requests.push({
                setDataValidation: {
                    range: {
                        sheetId: 0,
                        startRowIndex: 2,
                        endRowIndex,
                        startColumnIndex: colIndex,
                        endColumnIndex: colIndex + 1,
                    },
                    rule: {
                        condition: {
                            type: 'CUSTOM_FORMULA',
                            values: [
                                {
                                    userEnteredValue: `=REGEXMATCH(TO_TEXT(INDIRECT(ADDRESS(ROW(), COLUMN()))), "^((\\+91[-\\s]?)?[6-9][0-9]{9})$")`,
                                },
                            ],
                        },
                        inputMessage: 'Enter a valid 10-digit phone number.',
                        strict: false,
                        showCustomUi: true,
                    },
                },
            });
        }
    });

    if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: sheetId,
            requestBody: { requests },
        });
    }
}

async function getSheetData(sheetId) {
    const auth = await getAuthenticatedClient();
    const sheets = getSheetsClient(auth);

    const res = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: 'Sheet1',
    });

    const [headers, ...rows] = res.data.values;
    const data = rows.map(row => {
        const obj = {};
        headers.forEach((header, i) => {
            obj[header] = row[i] || '';
        });
        return obj;
    });

    return { data, headers };
}

async function revalidateSheetData(sheetId) {
    const { data } = await getSheetData(sheetId);
    const validated = validateRows(data);
    await writeToGoogleSheet(sheetId, validated);
    // return validated;
}

async function attachScriptToSheet(sheetId) {
    try {
        const auth = await getAuthenticatedClient();
        const script = google.script({ version: 'v1', auth });

        console.log('Reading code.gs file...');
        const code = await fs.readFile('scripts/code.gs', 'utf8');
        // console.log('Code.gs content:', code);

        // console.log('Creating script project for sheet:', sheetId);
        const { data: scriptProject } = await script.projects.create({
            requestBody: {
                title: `Script for Sheet ${sheetId}`,
                parentId: sheetId,
            },
        });
        console.log('Script project created:', scriptProject.scriptId);

        console.log('Uploading script content...');
        await script.projects.updateContent({
            scriptId: scriptProject.scriptId,
            requestBody: {
                files: [
                    {
                        name: 'Code',
                        type: 'SERVER_JS',
                        source: code,
                    },
                    {
                        name: 'appsscript',
                        type: 'JSON',
                        source: JSON.stringify({
                            timeZone: 'Asia/Kolkata',
                            exceptionLogging: 'STACKDRIVER',
                            runtimeVersion: 'V8',
                        }),
                    },
                ],
            },
        });

        await new Promise(resolve => setTimeout(resolve, 2000));

        console.log(`Script injected successfully into spreadsheet: https://docs.google.com/spreadsheets/d/${sheetId}`);
        return scriptProject.scriptId;
    } catch (error) {
        console.error('Error injecting script:', error.message);
        throw error;
    }
}

module.exports = {
    createSheetFromTemplate,
    writeToGoogleSheet,
    getSheetData,
    revalidateSheetData,
    attachScriptToSheet,
};