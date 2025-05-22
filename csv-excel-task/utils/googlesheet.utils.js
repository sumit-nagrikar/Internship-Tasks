const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

const CREDENTIALS_PATH = path.join(__dirname, '../config/credentials.json');

const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));

const authClient = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive']
});

const sheets = google.sheets({ version: 'v4', auth: authClient });
const drive = google.drive({ version: 'v3', auth: authClient });

module.exports = { authClient, sheets, drive };