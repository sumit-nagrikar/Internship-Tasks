const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function findEmailKeyFromHeaders(row) {
    const keys = Object.keys(row);
    return keys.find((k) => k.trim().toLowerCase() === "contact_email" || k.trim().toLowerCase() === "email");
}

function validateEmail(email) {
    return typeof email === "string" && emailRegex.test(email.trim());
}

function validateRows(rows) {
    if (!rows || rows.length === 0) return { validatedRows: [], totalErrors: 0 };

    const emailKey = findEmailKeyFromHeaders(rows[0]);
    const validatedRows = [];
    let totalErrors = 0;

    rows.forEach((row) => {
        if (!Object.values(row).some(cell => cell?.toString().trim())) return;

        let newRow = { ...row };
        const statusMessages = [];
        let errorCount = 0;

        const orgCode = row["ORG_CODE"]?.toString().trim();
        if (orgCode) {
            const normalizedOrgCode = orgCode.toUpperCase().replace(/-/g, "");
            if (!/^U\d{5}$/.test(normalizedOrgCode)) {
                statusMessages.push(`ORG_CODE should be in format U-XXXXX, got ${orgCode}`);
                errorCount++;
            }
        } else {
            statusMessages.push("ORG_CODE is required");
            errorCount++;
        }

        const orgPin = row["ORG_PIN"]?.toString().trim();
        if (orgPin && !/^\d{6}$/.test(orgPin)) {
            statusMessages.push("ORG_PIN should be a 6-digit numeric value");
            errorCount++;
        }

        if (emailKey) {
            const email = row[emailKey]?.toString().trim();
            if (email && !validateEmail(email)) {
                statusMessages.push("Invalid Email");
                errorCount++;
            }
        }

        if (errorCount > 0) totalErrors += 1;

        newRow = {
            Status: statusMessages.length ? statusMessages.join("\n") : "Valid",
            ErrorsCount: errorCount,
            ...newRow,
        };

        validatedRows.push(newRow);
    });

    return validatedRows;
}


function chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
        chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
}

module.exports = { validateRows, chunkArray };