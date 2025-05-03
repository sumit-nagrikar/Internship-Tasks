const express = require('express');
const router = express.Router();
const upload = require('../utils/multerConfig');
const fs = require('fs');
const path = require('path');
const { parseCSV, parseExcel } = require('../services/fileParser.service');
const { writeToGoogleSheet, createNewGoogleSheet } = require('../services/googleSheet.service');
const { revalidateSheetData } = require('../services/googleSheet.service');

let open;
(async () => {
  open = (await import('open')).default;
})();


router.post('/upload', upload.single('file'), async (req, res) => {
    const file = req.file;
    const extension = path.extname(file.originalname);

    try {
        let parsedData = [];

        const { sheetId, sheetUrl } = await createNewGoogleSheet();

        if (extension === '.csv') {
            parsedData = await parseCSV(file.path);
        } else if (extension === '.xlsx') {
            parsedData = parseExcel(file.path);
        } else {
            throw new Error('Only .csv or .xlsx files are supported.');
        }

        await writeToGoogleSheet(sheetId, parsedData);

        fs.unlinkSync(file.path);

        await open(sheetUrl);

        return res.json({
            message: 'Data written to new Google Sheet successfully.',
            sheetUrl
        });

    } catch (err) {
        if (file && fs.existsSync(file.path)) fs.unlinkSync(file.path);
        return res.status(500).json({ error: err.message });
    }
});


router.get('/recheck/:sheetId', async (req, res) => {
    try {console.log("req.params",req.params);
    const {sheetId} =req.params;
    
        const result = await revalidateSheetData(sheetId);

        res.json({ success: true, data: result });
    } catch (err) {
        console.error('Recheck error:', err);
        res.status(500).json({ success: false, error: 'Something went wrong' });
    }
});

module.exports = router;
