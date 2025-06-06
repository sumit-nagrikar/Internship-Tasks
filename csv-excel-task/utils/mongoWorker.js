// utils/mongoWorker.js
const { Worker } = require('bullmq');
const { connection, mongoQueueName } = require('./queue.bullmq');
const { processMongoJob } = require('./mongoJobHandler.js');

function startMongoWorker() {
    const worker = new Worker(
        mongoQueueName,
        async (job) => {
            const { orgData, sheetId, orgName } = job.data;
            // console.log(`Processing MongoDB job for orgName ${orgName} from sheet ${sheetId} and ${orgData.length} rows`);
            try {
                await processMongoJob({ orgData, orgName, sheetId });
                console.log(`Processed MongoDB submission for orgName ${orgName} from sheet ${sheetId}`);
            } catch (error) {
                console.error(`Error processing MongoDB job ${orgName}:`, error.stack);
                throw error;
            }
        },
        {
            connection,
            concurrency: 3,
        }
    );

    worker.on('failed', (job, err) => {
        console.error(`MongoDB job ${job.data.orgName} (sheet ${job.data.sheetId}) failed: ${err.stack}`);
    });

    worker.on('completed', (job) => {
        console.log(`MongoDB job ${job.data.orgName} (sheet ${job.data.sheetId}) completed`);
    });

    worker.on('error', (err) => {
        console.error('Mongo worker error:', err.stack);
    });

    console.log(`Mongo worker started for queue ${mongoQueueName}`);
}

if (require.main === module) {
    startMongoWorker();
}

module.exports = { startMongoWorker };