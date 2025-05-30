const { Queue } = require("bullmq");

const connection = {
    host: "127.0.0.1",
    port: 6379,
};

const queueName = "data_queue";
const queue = new Queue(queueName, { connection });

module.exports = { queue, connection, queueName };