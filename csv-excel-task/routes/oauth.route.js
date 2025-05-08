const express = require('express');
const router = express.Router();
const { oauth2Client, setTokens } = require('../utils/googlesheet.utils');

router.get('/oauth2callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send("No code provided");

  try {
    const { tokens } = await oauth2Client.getToken(code);
    setTokens(tokens);
    res.send('Authentication successful! You can close this window.');
  } catch (err) {
    console.error(err);
    res.status(500).send('Authentication failed');
  }
});

module.exports = router;
