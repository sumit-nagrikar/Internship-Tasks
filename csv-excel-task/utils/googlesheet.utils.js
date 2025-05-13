const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
const readline = require('readline');

const CREDENTIALS_PATH = path.join(__dirname, '../config/oauth2.json');
const TOKEN_PATH = path.join(__dirname, '../config/token.json');

const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
const { client_secret, client_id, redirect_uris } = credentials.installed;

const oauth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
);


// SCOPES needed
const SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.external_request',
];

// Get authenticated client (if token.json exists)
async function getAuthenticatedClient() {
    if (fs.existsSync(TOKEN_PATH)) {
        const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
        oauth2Client.setCredentials(token);
    } else {
        await getNewToken(); // generates token.json interactively
    }
    return oauth2Client;
}

// Step to generate new token
function getNewToken() {
    const authUrl = oauth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
        prompt: 'consent',
    });

    console.log('\n Authorize this app by visiting this URL:\n', authUrl);

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    return new Promise((resolve, reject) => {
        rl.question('\nEnter the code from that page here: ', (code) => {
            rl.close();
            oauth2Client.getToken(code, (err, token) => {
                if (err) {
                    console.error('Error retrieving access token', err);
                    return reject(err);
                }
                oauth2Client.setCredentials(token);
                fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
                console.log('Token stored to', TOKEN_PATH);
                resolve(oauth2Client);
            });
        });
    });
}

// Sheets client (after auth)
function getSheetsClient(auth) {
    return google.sheets({ version: 'v4', auth });
}

function setTokens(tokens) {
    oauth2Client.setCredentials(tokens);
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
    console.log('Token saved to', TOKEN_PATH);
}

module.exports = {
    oauth2Client,
    getAuthenticatedClient,
    getSheetsClient,
    setTokens,
};
