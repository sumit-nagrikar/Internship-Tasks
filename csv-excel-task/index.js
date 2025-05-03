const express = require('express');
const cors = require('cors');
const uploadRoute = require('./routes/upload.route');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use('/', uploadRoute);

app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});