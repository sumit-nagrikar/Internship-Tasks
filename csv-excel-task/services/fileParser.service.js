const XLSX = require('xlsx');

const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const phoneRegex = /^(\+91[-\s]?)?[6-9]\d{9}$/;

function findEmailKeyFromHeaders(row) {
    const keys = Object.keys(row);
    return keys.find(k => k.trim().toLowerCase() === 'contact_email');
}

function findPhoneKeysFromHeaders(row) {
    const keys = Object.keys(row);
    const phoneKeywords = ['phone', 'mobile', 'contact', 'contact_number'];
    return keys.filter(k =>
        phoneKeywords.some(keyword =>
            k.toLowerCase().includes(keyword) && k.toLowerCase() !== 'contact_email'
        )
    );
}

function validateEmail(email) {
    return typeof email === 'string' && emailRegex.test(email.trim());
}

function validatePhone(phone) {
    return typeof phone === 'string' && phoneRegex.test(phone.trim());
}

function validateRows(rows) {
    if (!rows || rows.length === 0) return [];

    const emailKey = findEmailKeyFromHeaders(rows[0]);
    const phoneKeys = findPhoneKeysFromHeaders(rows[0]);

    const groupedBySchool = {};

    rows.forEach((row, index) => {
        const newRow = { ...row };
        const statusMessages = [];
        let errorCount = 0;

        // Normalize school name for grouping
        let schoolName = row['SCHOOL_NAME']?.toString().trim();
        const normalizedSchoolName = schoolName?.toLowerCase();

        if (!schoolName) {
            statusMessages.push('SCHOOL_NAME is required');
            errorCount++;
        }

        // Name Validation
        const firstName = row['FIRST_NAME']?.toString().trim();
        if (!firstName) {
            statusMessages.push('FIRST_NAME is required');
            errorCount++;
        }

        // Email Validation
        if (emailKey) {
            const email = row[emailKey]?.toString().trim();
            if (!validateEmail(email)) {
                statusMessages.push('Invalid Email');
                errorCount++;
            }
        } else {
            statusMessages.push('CONTACT_EMAIL Field Missing');
            errorCount++;
        }

        // Phone Validation (commented out due to missing phone column)
        /*
        if (phoneKeys.length > 0) {
            const hasValidPhone = phoneKeys.some(key => {
                const phone = row[key]?.toString().trim();
                return validatePhone(phone);
            });

            if (!hasValidPhone) {
                statusMessages.push('Invalid Phone');
                errorCount++;
            }
        } else {
            console.warn(`No phone-related columns found in headers for row ${index + 3}`);
            statusMessages.push('Phone Field Missing');
            errorCount++;
        }
        */

        newRow.status = statusMessages.length ? statusMessages.join(', ') : 'Valid';
        newRow.errorsCount = errorCount;

        if (!normalizedSchoolName) {
            if (!groupedBySchool['__UNKNOWN__']) groupedBySchool['__UNKNOWN__'] = { schoolDisplayName: 'Unknown', rows: [] };
            groupedBySchool['__UNKNOWN__'].rows.push(newRow);
        } else {
            if (!groupedBySchool[normalizedSchoolName]) {
                groupedBySchool[normalizedSchoolName] = {
                    schoolDisplayName: schoolName,
                    rows: [],
                };
            }
            groupedBySchool[normalizedSchoolName].rows.push(newRow);
        }
    });

    return groupedBySchool;
}

function parseExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Read first row (metadata to be written to spreadsheet)
    const firstRow = XLSX.utils.sheet_to_json(sheet, {
        defval: '',
        range: 0,
        header: 1
    })[0];

    // Read data starting from second row as headers, third row as data
    const jsonWithHeaderRow = XLSX.utils.sheet_to_json(sheet, {
        defval: '',
        range: 1
    });

    const groupedData = validateRows(jsonWithHeaderRow);
    return { firstRow, groupedData };
}

module.exports = { parseExcel, validateRows };