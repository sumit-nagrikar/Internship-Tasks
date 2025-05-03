const { google } = require('googleapis');
const credentials = require('../config/credentials.json');
const { validateRows } = require('./fileParser.service');

// Auth client setup
function getAuthClient() {
    return new google.auth.GoogleAuth({
        credentials,
        scopes: [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
        ],
    });
}

// Sheets client
function getSheetsClient(auth) {
    return google.sheets({ version: 'v4', auth });
}

// Create a new Google Sheet and make it public
async function createNewGoogleSheet() {
    const auth = await getAuthClient().getClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const drive = google.drive({ version: 'v3', auth });

    const { data } = await sheets.spreadsheets.create({
        requestBody: {
            properties: {
                title: `Upload Sheet ${new Date().toISOString()}`,
            },
        },
    });

    const sheetId = data.spreadsheetId;

    await drive.permissions.create({
        fileId: sheetId,
        requestBody: {
            role: 'writer',
            type: 'anyone',
        },
    });

    return {
        sheetId,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${sheetId}`,
    };
}

// Write headers + data to a Google Sheet
async function writeToGoogleSheet(sheetId, data) {
    const auth = await getAuthClient().getClient();
    const sheets = getSheetsClient(auth);

    if (!data || data.length === 0) return;

    const headers = Object.keys(data[0]);
    const values = data.map(obj => headers.map(key => obj[key]));

    await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: 'Sheet1!A1',
        valueInputOption: 'RAW',
        requestBody: {
            values: [headers, ...values],
        },
    });

    await applyConditionalFormatting(sheetId, headers);
}

// Highlight "Invalid" status rows using conditional formatting
async function applyConditionalFormatting(sheetId, headers) {
    const auth = await getAuthClient().getClient();
    const sheets = getSheetsClient(auth);

    const statusColIndex = headers.findIndex(h => h.toLowerCase() === 'status');
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
                                    startRowIndex: 1,
                                    startColumnIndex: 0,
                                    endColumnIndex: headers.length,
                                },
                            ],
                            booleanRule: {
                                condition: {
                                    type: 'CUSTOM_FORMULA',
                                    values: [
                                        {
                                            userEnteredValue: `=ISNUMBER(SEARCH("Invalid", INDIRECT(ADDRESS(ROW(), ${statusColIndex + 1}))))`,
                                        },
                                    ],
                                },
                                format: {
                                    backgroundColor: { red: 1, green: 0.8, blue: 0.8 },
                                    textFormat: { bold: true },
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

// Get data from an existing sheet
async function getSheetData(sheetId) {
    const auth = await getAuthClient().getClient();
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

// Re-validate and update Google Sheet with fresh validation results
async function revalidateSheetData(sheetId) {
    const { data } = await getSheetData(sheetId);
    const validated = validateRows(data);
    await writeToGoogleSheet(sheetId, validated);
    return validated;
}

module.exports = {
    createNewGoogleSheet,
    writeToGoogleSheet,
    getSheetData,
    revalidateSheetData,
};
