const express = require('express');
const Product = require('../product.model');
const redisClient = require('../redisClient');
const { cacheMiddleware, cacheMiddlewareDynamic } = require('../middlewares/cache.middleware');

const router = express.Router();

// GET all products
router.get('/', cacheMiddleware('all_products', 60), async (req, res) => {
    try {
        const products = await Product.find({});
        if (!products || products.length === 0) {
            return res.status(404).json({ message: 'No products found' });
        }

        res.json({
            source: 'mongoDB',
            data: products,
        });
    } catch (err) {
        console.error('Error fetching products:', err);
        res.status(500).json({ message: 'Server error' });
    }
});

// GET product
router.get('/:id', cacheMiddlewareDynamic((req) => `product:${req.params.id}`, 60), async (req, res) => {
    const { id } = req.params;

    try {
        const product = await Product.findById(id);
        if (!product) return res.status(404).json({ message: 'Product not found' });

        res.json({
            source: 'mongoDB',
            data: product,
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ message: 'Server error' });
    }
});

// POST product creation
router.post('/', async (req, res) => {
    const { name, stock } = req.body;

    try {
        const product = new Product({ name, stock });
        await product.save();
        await redisClient.del('all_products');

        res.json(product);
    } catch (err) {
        console.error(err);
        res.status(500).json({ message: 'Create failed' });
    }
});

module.exports = router;