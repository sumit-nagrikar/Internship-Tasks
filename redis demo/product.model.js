const mongoose = require('mongoose');

const productSchema = new mongoose.Schema({
  name: String,
  stock: Number,
});

module.exports = mongoose.model('Product', productSchema);
