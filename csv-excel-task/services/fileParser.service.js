const fs = require('fs');
const csv = require('csv-parser');
const XLSX = require('xlsx');

const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const phoneRegex = /^(\+91[-\s]?)?[6-9]\d{9}$/;

function findEmailKeyFromHeaders(row) {
    const keys = Object.keys(row);
    return keys.find(k => k.trim().toLowerCase() === 'email');
}

function findPhoneKeysFromHeaders(row) {
    const keys = Object.keys(row);
    const phoneKeywords = ['phone', 'mobile', 'contact'];
    return keys.filter(k =>
        phoneKeywords.some(keyword =>
            k.toLowerCase().includes(keyword)
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

    return rows.map((row, index) => {
        const newRow = { ...row };
        const statusMessages = [];
        let errorCount = 0;

        // Email Validation
        if (emailKey) {
            const email = row[emailKey]?.toString().trim();
            if (!validateEmail(email)) {
                statusMessages.push('Invalid Email');
                errorCount++;
            }
        } else {
            statusMessages.push('Email Field Missing');
            errorCount++;
        }

        // Phone Validation
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
            statusMessages.push('Phone Field Missing');
            errorCount++;
        }

        newRow.status = statusMessages.length ? statusMessages.join(', ') : 'Valid';//changed header names
        newRow.errorsCount = errorCount;

        return newRow;
    });
}


function parseCSV(filePath) {
    return new Promise((resolve, reject) => {
        const rows = [];
        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (row) => rows.push(row))
            .on('end', () => resolve(validateRows(rows)))
            .on('error', reject);
    });
}

function parseExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    return validateRows(rawRows);
}

module.exports = { parseCSV, parseExcel, validateRows };
