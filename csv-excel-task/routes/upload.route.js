const express = require("express");
const router = express.Router();
const upload = require("../utils/multerConfig");
const fs = require("fs").promises;
const path = require("path");
const {
  parseExcel,
  parseExcelInOneSource,
  parseGoogleSheet,
} = require("../services/fileParser.service");
const {
  createSheetFromTemplate,
  writeToGoogleSheet,
  attachScriptToSheet,
  revalidateSheetData,
  buildAndCreateSheetFromParsedData,
  writeToGoogleSheetNew,
} = require("../services/googleSheet.service");

let open;
(async () => {
  open = (await import("open")).default;
})();

router.post("/data", upload.single("file"), async (req, res) => {
  const file = req.file;
  const extension = path.extname(file.originalname);

  try {
    if (!file) throw new Error("No file uploaded");
    if (extension !== ".xlsx")
      throw new Error("Only .xlsx Excel files are supported.");

    const data = await parseExcel(file.path);

    if (!data.groupedData || Object.keys(data.groupedData).length === 0) {
      throw new Error("No valid data found in the uploaded file.");
    }

    const { sheetId, sheetUrl } = await writeToGoogleSheet(data);

    await open(sheetUrl);

    await fs
      .unlink(file.path)
      .catch((err) => console.error("File deletion error:", err));

    return res.json({
      message: "Data processed and written to Google Sheet successfully",
      sheetUrl,
    });
  } catch (err) {
    try {
      await fs.access(file.path);
      await fs
        .unlink(file.path)
        .catch((err) => console.error("File deletion error:", err));
    } catch (accessErr) {
      console.warn(
        "File does not exist or already deleted:",
        accessErr.message
      );
    }
    console.error("Upload error:", err);
    return res.status(500).json({ error: err.message });
  }
});

router.post("/excel", upload.single("file"), async (req, res) => {
  const file = req.file;
  const extension = path.extname(file.originalname);

  try {
    if (!file) throw new Error("No file uploaded");
    if (extension !== ".xlsx")
      throw new Error("Only .xlsx Excel files are supported.");

    const data = await parseExcelInOneSource(file.path);

    if (!data || Object.keys(data).length === 0) {
      throw new Error("No valid data found in the uploaded file.");
    }

    const { sheetUrl } = await buildAndCreateSheetFromParsedData(data);

    await open(sheetUrl);

    await fs
      .unlink(file.path)
      .catch((err) => console.error("File deletion error:", err));

    return res.json({
      message: "Data processed and written to Google Sheet successfully",
      sheetUrl,
    });
  } catch (err) {
    try {
      await fs.access(file.path);
      await fs
        .unlink(file.path)
        .catch((err) => console.error("File deletion error:", err));
    } catch (accessErr) {
      console.warn(
        "File does not exist or already deleted:",
        accessErr.message
      );
    }
    console.error("Upload error:", err);
    return res.status(500).json({ error: err.message });
  }
});

router.post("/validate/:sheetId", async (req, res) => {
  const { sheetId } = req.params;
  try {
    const data = await parseGoogleSheet(sheetId);

    if (!data || Object.keys(data).length === 0) {
      throw new Error("No valid data found sheet.");
    }

    const { sheetUrl } = await writeToGoogleSheetNew(sheetId, data);

    await open(sheetUrl);

    return res.json({
      message: "Data processed and written to Google Sheet successfully",
      sheetUrl,
    });
  } catch (err) {
    console.error("Upload error:", err);
    return res.status(500).json({ error: err.message });
  }
});

module.exports = router;
