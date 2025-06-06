const { Queue } = require('bullmq');

const connection = {
    host: '127.0.0.1',
    port: 6379,
};

const queueName = 'data_queue';
const queue = new Queue(queueName, { connection });

const mongoQueueName = 'mongo_queue';
const mongoQueue = new Queue(mongoQueueName, { connection });

module.exports = { queue, mongoQueue, connection, queueName, mongoQueueName };