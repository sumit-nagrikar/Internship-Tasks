// utils/worker.js
const { Worker } = require("bullmq");
const { connection, queueName } = require("./queue.bullmq");
const { writeToGoogleSheetNew } = require("../services/googleSheet.service");
const { processMongoJob } = require("./mongoJobHandler");

function startWorker() {
  const worker = new Worker(
    queueName,
    async (job) => {
      const { type, chunkIndex, headers, validatedData, sheetId, isFirstChunk, isLastChunk, jobNumber } = job.data;

      try {
        await writeToGoogleSheetNew(
          sheetId,
          { headers, validatedData, isLastChunk },
          "Sheet1",
          !isFirstChunk
        );
        console.log(`Processed chunk ${chunkIndex + 1} (Job ${jobNumber}) to Google Sheet ${sheetId}`);
      } catch (error) {
        console.error(`Error processing job ${jobNumber}:`, error.stack);
        throw error;
      }
    },
    {
      connection,
      concurrency: 1,
    }
  );

  worker.on("failed", (job, err) => {
    console.error(`Job ${job.data.jobNumber} failed: ${err.stack}`);
  });

  worker.on("completed", (job) => {
    console.log(`Job ${job.data.jobNumber} completed`);
  });

  worker.on("error", (err) => {
    console.error("Worker error:", err.stack);
  });

  console.log(`Worker started for queue ${queueName}`);
}

if (require.main === module) {
  startWorker();
}

module.exports = { startWorker };
