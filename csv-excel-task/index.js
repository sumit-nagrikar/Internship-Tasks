const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const uploadRoute = require('./routes/upload.route');
const { startWorker } = require('./utils/sheet.worker');
const { startMongoWorker } = require('./utils/mongoWorker');
require('dotenv').config();


const app = express();
const PORT = 3051;

mongoose.connect(process.env.MONGO_URI, { useNewUrlParser: true, useUnifiedTopology: true })
    .then(() => console.log('MongoDB connected'))
    .catch(err => console.error('MongoDB connection error:', err));

app.use(cors());
app.use(express.json());
app.use('/upload', uploadRoute);

startWorker();
startMongoWorker();

app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});