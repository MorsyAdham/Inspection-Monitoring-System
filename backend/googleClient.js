// googleClient.js
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

function createSheetsClient() {
  const credPath = path.join(__dirname, 'credentials.json');
  if (!fs.existsSync(credPath)) {
    throw new Error('credentials.json not found in backend folder. For local dev, place the service account JSON as backend/credentials.json. For deployment, use env var.');
  }
  const auth = new google.auth.GoogleAuth({
    keyFile: credPath,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  return google.sheets({ version: 'v4', auth });
}

module.exports = { createSheetsClient };
