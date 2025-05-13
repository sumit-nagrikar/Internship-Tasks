const { google } = require('googleapis');
const fs = require('fs').promises;
const { getAuthenticatedClient, getSheetsClient } = require('../utils/googlesheet.utils');

const TEMPLATE_SHEET_ID = '1rKT9Q-zZ-vQA6CZNrMd10qFPSwbHlT51EJ6UUc4g_uI';

async function createSheetFromTemplate() {
    if (!TEMPLATE_SHEET_ID) throw new Error("TEMPLATE_SHEET_ID not available");

    const auth = await getAuthenticatedClient();
    const drive = google.drive({ version: 'v3', auth });
    const sheets = getSheetsClient(auth);

    // Copy the template spreadsheet
    const { data } = await drive.files.copy({
        fileId: TEMPLATE_SHEET_ID,
        requestBody: {
            name: `Upload Sheet ${new Date().toISOString()}`,
        },
    });

    // Clear existing sheets except one (to ensure a clean slate)
    const spreadsheetId = data.id;
    const metadata = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'sheets.properties',
    });

    const existingSheets = metadata.data.sheets;
    if (existingSheets.length > 1) {
        const requests = existingSheets.slice(1).map(sheet => ({
            deleteSheet: {
                sheetId: sheet.properties.sheetId,
            },
        }));
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests },
        });
    }

    // Rename the remaining sheet to a temporary name
    if (existingSheets.length > 0) {
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: [{
                    updateSheetProperties: {
                        properties: {
                            sheetId: existingSheets[0].properties.sheetId,
                            title: 'TempSheet',
                        },
                        fields: 'title',
                    },
                }],
            },
        });
    }

    await drive.permissions.create({
        fileId: spreadsheetId,
        requestBody: {
            role: 'writer',
            type: 'anyone',
        },
    });

    return {
        sheetId: spreadsheetId,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
    };
}

async function initializeSchoolSheets(auth, spreadsheetId, schools) {
    try {
        const sheets = getSheetsClient(auth);

        // Fetch existing sheets
        const response = await sheets.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties.title',
        });

        const existingSheets = response.data.sheets.map(
            (sheet) => sheet.properties.title.toUpperCase()
        );
        console.log('Existing sheets:', existingSheets);

        // Filter out schools that already have sheets (case-insensitive)
        const sheetsToCreate = schools.filter(
            (school) => !existingSheets.includes(school.toUpperCase())
        );

        if (sheetsToCreate.length === 0) {
            console.log('All sheets already exist. Proceeding with data upload...');
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
            resource: {
                requests,
            },
        });

        console.log('Created sheets:', sheetsToCreate);
    } catch (error) {
        console.error('Error initializing sheets:', error.message);
        throw error;
    }
}

async function writeToGoogleSheet(sheetId, data, scriptId) {
    const { firstRow, groupedData } = data;
    if (!groupedData || Object.keys(groupedData).length === 0) {
        console.warn('No grouped data provided');
        return;
    }

    const auth = await getAuthenticatedClient();
    const sheets = getSheetsClient(auth);

    // Initialize sheets for all schools
    const schoolNames = Object.keys(groupedData).map(
        school => groupedData[school].schoolDisplayName || school
    );
    await initializeSchoolSheets(auth, sheetId, schoolNames);

    // Write data to each sheet
    for (const schoolName of Object.keys(groupedData)) {
        const data = groupedData[schoolName].rows;
        if (!data || !Array.isArray(data) || data.length === 0) {
            console.warn(`No valid data for school: ${schoolName}`);
            continue;
        }

        const sheetTitle = groupedData[schoolName].schoolDisplayName || schoolName;

        // Dynamically generate headers from the first row's keys
        if (!data[0]) {
            console.warn(`No data rows for school: ${schoolName}`);
            continue;
        }
        const rawHeaders = Object.keys(data[0]);
     
        const headers = rawHeaders.map(header => {
            const upperHeader = header.toUpperCase();
            if (upperHeader === 'SCHOOL_NAME') return 'SCHOOL_NAME';
            if (upperHeader === 'FIRST_NAME') return 'Name';
            if (upperHeader === 'CONTACT_EMAIL') return 'Email';
            if (upperHeader.includes('PHONE') || upperHeader.includes('MOBILE') || upperHeader.includes('CONTACT')) return 'Phone';
            return header;
        });
        if (headers.length === 0) {
            console.warn(`No valid headers for school: ${schoolName}`);
            continue;
        }

        const rows = data.map(row => headers.map(header => row[header] || row[header.toLowerCase()] || row[header.toUpperCase()] || ''));

        await sheets.spreadsheets.values.update({
            spreadsheetId: sheetId,
            range: `'${sheetTitle}'!A1`,
            valueInputOption: 'RAW',
            requestBody: {
                values: [firstRow, headers, ...rows],
            },
        });

        // Resize columns
        const sheetIdNum = await getSheetIdByName(sheets, sheetId, sheetTitle);
        if (sheetIdNum !== null) {
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: sheetId,
                requestBody: {
                    requests: [{
                        autoResizeDimensions: {
                            dimensions: {
                                sheetId: sheetIdNum,
                                dimension: 'COLUMNS',
                                startIndex: 0,
                                endIndex: Math.max(firstRow.length, headers.length),
                            },
                        },
                    }],
                },
            });
        }

        // Apply formatting and validation
        await applyConditionalFormatting(sheetId, headers, sheetTitle);
        await applyDataValidation(sheetId, headers, data.length, sheetTitle);
    }

    // Run Apps Script validation
    if (scriptId) {
        const script = google.script({ version: 'v1', auth });
        try {
            console.log(`Attempting to run script with ID: ${scriptId}`);
            // Verify script content
            const content = await script.projects.getContent({ scriptId });
            const hasValidateAllRows = content.data.files.some(file => 
                file.source.includes('function validateAllRows')
            );
            if (!hasValidateAllRows) {
                throw new Error('validateAllRows function not found in script');
            }
            // Retry script execution
            const response = await runScriptWithRetry(script, scriptId);
            console.log('Apps Script validateAllRows executed:', response.data);
        } catch (err) {
            console.error(`Script execution failed for scriptId ${scriptId}: ${err.message}`);
            console.error('Error details:', JSON.stringify(err, null, 2));
            if (err.message.includes('Requested entity was not found')) {
                console.error('Possible causes: Invalid scriptId, script project not propagated, missing permissions, Google Apps Script API not enabled, or deployment not active.');
                console.error(`Check script project: https://script.google.com/home/projects/${scriptId}`);
            }
        }
    } else {
        console.warn('No scriptId provided; skipping Apps Script validation');
    }
}

async function runScriptWithRetry(script, scriptId, maxRetries = 3, delay = 2000) {
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const response = await script.scripts.run({
                scriptId,
                resource: {
                    function: 'validateAllRows',
                    parameters: [],
                },
            });
            return response;
        } catch (err) {
            console.warn(`Attempt ${attempt} failed for scriptId ${scriptId}: ${err.message}`);
            if (attempt === maxRetries) {
                throw new Error(`Failed to run script after ${maxRetries} attempts: ${err.message}`);
            }
            await new Promise(resolve => setTimeout(resolve, delay));
        }
    }
}

async function attachScriptToSheet(sheetId) {
    try {
        const auth = await getAuthenticatedClient();
        const script = google.script({ version: 'v1', auth });

        console.log('Reading code.gs file...');
        const code = await fs.readFile('scripts/code.gs', 'utf8');

        console.log('Creating script project...');
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

        // Create a deployment to ensure script is executable
        console.log('Creating script deployment...');
        try {
            await script.projects.deployments.create({
                scriptId: scriptProject.scriptId,
                requestBody: {
                    versionNumber: 1,
                    manifestFileName: 'appsscript',
                    description: 'Initial deployment for validation script',
                },
            });
            console.log('Script deployment created');
        } catch (deployErr) {
        throw err;
        }

        // Increased delay to ensure scriptellisenproject and deployment propagation
        console.log('Waiting for script project to propagate...');
        await new Promise(resolve => setTimeout(resolve, 10000)); // 10 seconds

        // Verify script project exists
        try {
            await script.projects.get({ scriptId: scriptProject.scriptId });
            console.log(`Script project verified: https://script.google.com/home/projects/${scriptProject.scriptId}`);
        } catch (verifyErr) {
            console.error(`Failed to verify script project ${scriptProject.scriptId}: ${verifyErr.message}`);
            throw new Error(`Script project creation failed: ${verifyErr.message}`);
        }

        console.log(`Script injected successfully into spreadsheet: https://docs.google.com/spreadsheets/d/${sheetId}`);
        return scriptProject.scriptId;
    } catch (error) {
        throw error;
    }
}

async function getSheetIdByName(sheets, spreadsheetId, sheetTitle) {
    const response = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'sheets.properties',
    });
    const sheet = response.data.sheets.find(s => s.properties.title.toUpperCase() === sheetTitle.toUpperCase());
    return sheet ? sheet.properties.sheetId : null;
}

async function applyConditionalFormatting(sheetId, headers, sheetTitle) {
    const auth = await getAuthenticatedClient();
    const sheets = getSheetsClient(auth);

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

    // Find CGPA column (case-insensitive)
    const cgpaColIndex = headers.findIndex(header => header.toUpperCase() === 'CGPA');
    if (cgpaColIndex !== -1) {
        requests.push({
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetIdValue,
                        startRowIndex: 2, // Start from data rows (after metadata and headers)
                        startColumnIndex: cgpaColIndex,
                        endColumnIndex: cgpaColIndex + 1,
                    }],
                    booleanRule: {
                        condition: {
                            type: 'NUMBER_LESS',
                            values: [{ userEnteredValue: '6' }],
                        },
                        format: {
                            backgroundColor: { red: 1, green: 0.8, blue: 0.8 },
                            textFormat: { bold: true },
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
    }
}

async function applyDataValidation(sheetId, headers, rowCount, sheetTitle) {
    if (!headers || headers.length === 0 || rowCount <= 0) {
        console.warn(`Invalid input for data validation: headers=${headers?.length}, rowCount=${rowCount}, sheetTitle=${sheetTitle}`);
        return;
    }

    const auth = await getAuthenticatedClient();
    const sheets = getSheetsClient(auth);

    const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    const sheetIdMap = sheetMeta.data.sheets.reduce((map, sheet) => {
        map[sheet.properties.title.toUpperCase()] = sheet.properties.sheetId;
        return map;
    }, {});

    const sheetIdValue = sheetIdMap[sheetTitle.toUpperCase()];
    if (sheetIdValue === undefined) {
        console.error(`Sheet '${sheetTitle}' not found for data validation in spreadsheet ${sheetId}`);
        return;
    }

    const requests = [];

    // Find Email column (case-insensitive, multiple variations)
    const emailColIndex = headers.findIndex(header => 
        ['EMAIL', 'CONTACT_EMAIL', 'E-MAIL', 'EMAIL ADDRESS'].includes(header.toUpperCase())
    );
    if (emailColIndex !== -1) {
        requests.push({
            setDataValidation: {
                range: {
                    sheetId: sheetIdValue,
                    startRowIndex: 2, // Start from data rows
                    endRowIndex: rowCount + 2,
                    startColumnIndex: emailColIndex,
                    endColumnIndex: emailColIndex + 1,
                },
                rule: {
                    condition: {
                        type: 'CUSTOM_FORMULA',
                        values: [{ userEnteredValue: '=REGEXMATCH(INDIRECT("R[0]C[0]", FALSE), "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$")' }],
                    },
                    inputMessage: 'Must be a valid email (e.g., user@domain.com)',
                    strict: true,
                    showCustomUi: true,
                },
            },
        });
    }

    // Find Phone column (case-insensitive, multiple variations)
    const phoneColIndex = headers.findIndex(header => 
        ['PHONE', 'PHONE NUMBER', 'MOBILE', 'CONTACT'].includes(header.toUpperCase())
    );
    if (phoneColIndex !== -1) {
        requests.push({
            setDataValidation: {
                range: {
                    sheetId: sheetIdValue,
                    startRowIndex: 2, // Start from data rows
                    endRowIndex: rowCount + 2,
                    startColumnIndex: phoneColIndex,
                    endColumnIndex: phoneColIndex + 1,
                },
                rule: {
                    condition: {
                        type: 'CUSTOM_FORMULA',
                        values: [{ userEnteredValue: '=AND(LEN(INDIRECT("R[0]C[0]", FALSE))=10, ISNUMBER(VALUE(INDIRECT("R[0]C[0]", FALSE))))' }],
                    },
                    inputMessage: 'Must be a 10-digit phone number',
                    strict: true,
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
            console.error(`Failed to apply data validation to sheet '${sheetTitle}': ${error.message}`);
        }
    }
}

async function revalidateSheetData(sheetId) {
    console.warn('revalidateSheetData is not implemented');
    return { message: 'Revalidation not implemented' };
}

module.exports = {
    createSheetFromTemplate,
    writeToGoogleSheet,
    initializeSchoolSheets,
    attachScriptToSheet,
    revalidateSheetData,
};