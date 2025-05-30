const express = require('express');
const router = express.Router();
const upload = require("../utils/multerConfig");
const fs = require("fs").promises;
const path = require("path");
const { parseAndProcessFile } = require("../services/fileParser.service");
let open;
(async () => {
  open = (await import('open')).default;
})();

router.post('/data', upload.single('file'), async (req, res) => {
  const file = req.file;

  try {
    if (!file) throw new Error('No file uploaded');
    if (path.extname(file.originalname) !== '.xlsx') {
      throw new Error('Only .xlsx Excel files are supported.');
    }

    // Process the file using parseAndProcessFile
    const { status, message, sheetId } = await parseAndProcessFile(file.path);

    if (status !== 'success') {
      throw new Error(message);
    }

    const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;

    // Open the sheet in the browser
    await open(sheetUrl);

    // Delete the uploaded file
    await fs.unlink(file.path).catch(err => console.error('File deletion error:', err));

    return res.json({
      message: 'Data validated and queued for Google Sheet successfully',
      sheetUrl,
      sheetId,
    });

  } catch (err) {
    // Clean up file if it exists
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
  const extension = path.extname(file.originalname);

  try {
    if (!file) throw new Error("No file uploaded");
    if (extension !== ".xlsx") {
      throw new Error("Only .xlsx Excel files are supported.");
    }

    const data = await parseAndProcessFile(file.path);
    if (data.status !== 'success') {
      throw new Error(data.message);
    }

    const { sheetId } = data;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;

    await open(sheetUrl);

    await fs.unlink(file.path).catch((err) => console.error("File deletion error:", err));

    return res.json({
      message: "Data processed and queued to Google Sheet successfully",
      sheetUrl,
      sheetId,
    });
  } catch (err) {
    try {
      await fs.access(file?.path);
      await fs.unlink(file.path).catch((err) => console.error("File deletion error:", err));
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
    // Placeholder: Need parseGoogleSheet or equivalent
    const data = await parseAndProcessFile(`googleSheet:${sheetId}`); // Hypothetical, adjust based on implementation
    if (data.status !== 'success') {
      throw new Error(data.message);
    }

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

module.exports = router;