const redisClient = require('../redisClient');

const cacheMiddleware = (key, duration = 60) => async (req, res, next) => {
  const cached = await redisClient.get(key);
  if (cached) {
    return res.json({ source: 'redis cache', data: JSON.parse(cached) });
  }

  const originalJson = res.json.bind(res);
  res.json = async (body) => {
    if (body?.data) {
        console.log(`Caching data with key: ${key} for ${duration} seconds`);
      await redisClient.set(key, JSON.stringify(body.data), { EX: duration });
    }
    return originalJson(body);
  };

  next();
};

const cacheMiddlewareDynamic = (keyFn, duration = 60) => async (req, res, next) => {
  const key = keyFn(req);
  const cached = await redisClient.get(key);
  if (cached) {
    return res.json({ source: 'redis cache', data: JSON.parse(cached) });
  }

  const originalJson = res.json.bind(res);
  res.json = async (body) => {
    if (body?.data) {
        console.log(`Caching data with key: ${key} for ${duration} seconds`);
      await redisClient.set(key, JSON.stringify(body.data), { EX: duration });
    }
    return originalJson(body);
  };

  next();
};

module.exports = {
    cacheMiddleware,
    cacheMiddlewareDynamic,
  };