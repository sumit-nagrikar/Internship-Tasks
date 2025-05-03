// app.js
const express = require('express');
const connectDB = require('./db');
const app = express();
const PORT = 3000;

app.use(express.json());

connectDB();

const productRoutes = require('./routes/product.routes');
app.use('/product', productRoutes);

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});