const express = require('express');
const router = express.Router();
const upload = require('../utils/multerConfig');
const fs = require('fs').promises;
const path = require('path');
const { parseExcel } = require('../services/fileParser.service');
const { createSheetFromTemplate, writeToGoogleSheet, attachScriptToSheet, revalidateSheetData } = require('../services/googleSheet.service');

let open;
(async () => {
    open = (await import('open')).default;
})();

router.post('/data', upload.single('file'), async (req, res) => {
    const file = req.file;
    const extension = path.extname(file.originalname);

    try {
        if (!file) throw new Error('No file uploaded');
        if (extension !== '.xlsx') throw new Error('Only .xlsx Excel files are supported.');

        const data = await parseExcel(file.path);


        if (!data.groupedData || Object.keys(data.groupedData).length === 0) {
            throw new Error('No valid data found in the uploaded file.');
        }

        // Create a new spreadsheet
        const { sheetId, sheetUrl } = await createSheetFromTemplate();

        // Attach Google Apps Script
        let scriptId;
        try {
            scriptId = await attachScriptToSheet(sheetId);
        } catch (err) {
            console.warn('Script attachment failed, proceeding without script:', err.message);
            scriptId = null;
        }

        // Write data to sheets
        await writeToGoogleSheet(sheetId, data, scriptId);

        await open(sheetUrl);

        // Delete file
        await fs.unlink(file.path).catch(err => console.error('File deletion error:', err));

        return res.json({
            message: 'Data processed and written to Google Sheet successfully',
            sheetUrl,
        });

    } catch (err) {
        // Check if file exists before deleting
        try {
            await fs.access(file.path);
            await fs.unlink(file.path).catch(err => console.error('File deletion error:', err));
        } catch (accessErr) {
            console.warn('File does not exist or already deleted:', accessErr.message);
        }
        console.error('Upload error:', err);
        return res.status(500).json({ error: err.message });
    }
});

module.exports = router;