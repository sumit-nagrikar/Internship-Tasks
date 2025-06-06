// routes/upload.route.js
const express = require('express');
const router = express.Router();
const upload = require("../utils/multerConfig");
const fs = require("fs").promises;
const path = require("path");
const { parseAndProcessFile } = require("../services/fileParser.service");
const { getSheetData } = require("../services/googleSheet.service");
const { mongoQueue } = require("../utils/queue.bullmq");
let open;
(async () => {
  open = (await import('open')).default;
})();

// Existing routes (adjusted to match your style)
router.post('/data', upload.single('file'), async (req, res) => {
  const file = req.file;

  try {
    if (!file) throw new Error('No file uploaded');
    if (path.extname(file.originalname) !== '.xlsx') {
      throw new Error('Only .xlsx Excel files are supported.');
    }

    const sheetId = await parseAndProcessFile(file.path);

    const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;

    await open(sheetUrl);

    await fs.unlink(file.path).catch(err => console.error('File deletion error:', err));

    return res.json({
      message: 'Data validated and queued for Google Sheet successfully',
      sheetUrl,
      sheetId,
    });
  } catch (err) {
    try {
      await fs.access(file?.path);
      await fs.unlink(file.path).catch(err => console.error('File deletion error:', err));
    } catch (accessErr) {
      console.warn('File does not exist or already deleted:', accessErr.message);
    }
    console.error('Upload error:', err);
    return res.status(500).json({ error: err.message });
  }
});

router.post("/excel", upload.single("file"), async (req, res) => {
  const file = req.file;

  try {
    if (!file) throw new Error("No file uploaded");

    const extension = path.extname(file.originalname);
    if (extension !== ".xlsx") {
      throw new Error("Only .xlsx Excel files are supported.");
    }

    const sheetId = await parseAndProcessFile(file.path);

    console.log("Sheet created and data queued:", sheetId);

    const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;
    await open(sheetUrl);

    await fs.unlink(file.path).catch((err) =>
      console.error("File deletion error:", err)
    );

    return res.json({
      message: "Data processed and queued to Google Sheet successfully",
      sheetUrl,
      sheetId,
    });
  } catch (err) {
    try {
      if (file?.path) {
        await fs.access(file.path);
        await fs.unlink(file.path).catch((err) => console.error("File deletion error:", err));
      }
    } catch (accessErr) {
      console.warn("File does not exist or already deleted:", accessErr.message);
    }
    console.error("Upload error:", err);
    return res.status(500).json({ error: err.message });
  }
});

router.post("/validate/:sheetId", async (req, res) => {
  const { sheetId } = req.params;

  try {
    const data = await parseAndProcessFile(`googleSheet:${sheetId}`);
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;

    await open(sheetUrl);

    return res.json({
      message: "Data revalidated and written to Google Sheet successfully",
      sheetUrl,
      sheetId,
    });
  } catch (err) {
    console.error("Validation error:", err);
    return res.status(500).json({ error: err.message });
  }
});

router.post('/submit/:sheetId', async (req, res) => {
  const { sheetId } = req.params;
  try {
    const { data } = await getSheetData(sheetId, 'Sheet1');
    if (!data || !data.length) {
      return res.status(400).json({ error: 'No data found in sheet' });
    }
    const cleanedData = data.map(row => {
      const cleanedRow = { ...row };
      delete cleanedRow.Status;
      delete cleanedRow.ErrorsCount;
      return cleanedRow;
    }).filter(row => row.ORGNAME && row.ORGNAME !== 'ORGNAME' && row.ORGNAME.trim() !== '');

    const groupedByOrg = cleanedData.reduce((result, item) => {
      (result[item.ORGNAME] = result[item.ORGNAME] || []).push(item);
      return result;
    }, {});

    if (Object.keys(groupedByOrg).length === 0) {
      return res.status(400).json({ error: 'No valid organizations found in data' });
    }

    const sessionId = Date.now();
    const jobIds = [];
    for (const [orgName, orgData] of Object.entries(groupedByOrg)) {
      const jobId = `mongo_submit_${sheetId}_${orgName}_${sessionId}`;
      await mongoQueue.add(jobId, {
        orgData,
        sheetId,
        orgName,
      }, {
        attempts: 3,
        backoff: {
          type: 'exponential',
          delay: 1000,
        },
      });
      jobIds.push(jobId);
    }
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;
    // await open(sheetUrl);
    return res.json({
      message: 'Data queued for MongoDB submission',
      sheetUrl,
      sheetId,
      jobIds,
      totalOrganizations: Object.keys(groupedByOrg).length,
    });
  } catch (err) {
    console.error('Submission error:', err);
    return res.status(500).json({ error: err.message });
  }
});

module.exports = router;