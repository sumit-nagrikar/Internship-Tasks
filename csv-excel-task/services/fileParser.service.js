const XLSX = require("xlsx");
const { queue } = require("../utils/queue.bullmq");
const { createSheetFromTemplate, getSheetData } = require("../services/googleSheet.service");
const { validateRows, chunkArray } = require("../utils/validation");

const DEFAULT_SHEET_NAME = "Sheet1";

async function addToBullMQQueue(validatedData, firstRow, sheetId) {
  const chunks = chunkArray(validatedData, 20);
  const sessionId = Date.now();

  console.log(`Total chunks to add: ${chunks.length}`);

  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    const jobNumber = i + 1;
    const jobId = `${sheetId}_${sessionId}_${jobNumber}`;
    const jobData = {
      chunkIndex: i,
      headers: [firstRow, Object.keys(chunk[0])],
      validatedData: chunk,
      sheetId,
      isFirstChunk: i === 0,
      isLastChunk: i === chunks.length - 1,
      jobNumber,
    };

    await queue.add(jobId, jobData, {
      attempts: 3,
      backoff: {
        type: "exponential",
        delay: 1000,
      },
    });

  }
}


async function parseAndProcessFile(source) {
  try {
    let rows, firstRow, sheetId;

    if (source.startsWith("googleSheet:")) {
      sheetId = source.replace("googleSheet:", "");
      const { data, headers } = await getSheetData(sheetId, DEFAULT_SHEET_NAME);
      if (!data || data.length === 0) throw new Error("No data found in the Google Sheet");
      rows = data;
      firstRow = headers;
    } else {
      const workbook = XLSX.readFile(source);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
      if (!rows || rows.length === 0) throw new Error("No data found in the Excel file");
      firstRow = Object.keys(rows[0]);
      sheetId = (await createSheetFromTemplate()).sheetId;
    }

    const validatedRows = validateRows(rows);
    await addToBullMQQueue(validatedRows, firstRow, sheetId);


    return { status: "success", message: `Data validated and queued for Google Sheet ${sheetId}`, sheetId };
  } catch (error) {
    console.error("Error processing source:", error.stack);
    return { status: "error", message: error.message };
  }
}

module.exports = { parseAndProcessFile };