// index.js
require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const { createSheetsClient } = require('./googleClient');

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: '8mb' }));

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

if (!SPREADSHEET_ID) {
  console.error('SPREADSHEET_ID not set in env');
  process.exit(1);
}

// For Render: if env has GOOGLE_SERVICE_ACCOUNT (base64 JSON) write to credentials.json
if (process.env.GOOGLE_SERVICE_ACCOUNT_BASE64 && !fs.existsSync(path.join(__dirname, 'credentials.json'))) {
  const buf = Buffer.from(process.env.GOOGLE_SERVICE_ACCOUNT_BASE64, 'base64');
  fs.writeFileSync(path.join(__dirname, 'credentials.json'), buf);
}

const sheets = createSheetsClient();

// helpers
function normalizeKey(s){ return String(s||'').replace(/\u00A0/g,'').replace(/[^\w]/g,'').toLowerCase(); }
function colNumberToLetter(num) {
  let s = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    num = Math.floor((num - mod) / 26);
  }
  return s;
}

// read whole tab
async function readSheetRange(tab) {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${tab}`
  });
  const values = resp.data.values || [];
  const header = values[0] || [];
  const rows = values.slice(1).map(r => header.map((_,i) => r[i] === undefined ? '' : r[i]));
  return { header, rows, startRow: 2 };
}

// GET /files  => list sheet tabs
app.get('/files', async (req, res) => {
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const tabs = (meta.data.sheets || []).map(s => s.properties.title);
    res.json({ tabs });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// GET /sheet/:tab => header + rows
app.get('/sheet/:tab', async (req, res) => {
  const tab = req.params.tab;
  try {
    const data = await readSheetRange(tab);
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// POST /sheet/:tab/batchUpdate  { updates: [{ rowIndex:0, values:[...]}] }
app.post('/sheet/:tab/batchUpdate', async (req, res) => {
  const tab = req.params.tab;
  const updates = req.body.updates;
  if (!Array.isArray(updates)) return res.status(400).json({ error: 'updates must be array' });

  try {
    const { header, startRow } = await readSheetRange(tab);
    const nCols = header.length || (updates[0] && updates[0].values.length) || 1;

    const data = updates.map(u => {
      const sheetRow = startRow + u.rowIndex;
      const endLetter = colNumberToLetter(nCols);
      const range = `${tab}!A${sheetRow}:${endLetter}${sheetRow}`;
      return { range, values: [u.values.map(v => v === null ? '' : v)] };
    });

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'RAW', data }
    });

    res.json({ ok: true, updated: updates.length });
  } catch (err) {
    console.error('batch update error', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /sheet/:tab/export downloads xlsx Data+Analytics
app.get('/sheet/:tab/export', async (req, res) => {
  const tab = req.params.tab;
  try {
    const { header, rows } = await readSheetRange(tab);
    const objs = rows.map(r => {
      const o = {};
      header.forEach((h,i) => o[h || `Col${i+1}`] = r[i] ?? '');
      return o;
    });

    // analytics per container
    const map = {};
    // try to find indices
    const normHeader = header.map(h => normalizeKey(h));
    const idxContainer = normHeader.indexOf('containernum') !== -1 ? normHeader.indexOf('containernum') : normHeader.indexOf('container');
    const idxRemarks = normHeader.indexOf('remarks') !== -1 ? normHeader.indexOf('remarks') : (normHeader.indexOf('remark') !== -1 ? normHeader.indexOf('remark') : -1);

    rows.forEach(r => {
      const cont = idxContainer !== -1 ? String(r[idxContainer] ?? 'NA') : 'NA';
      const rem = idxRemarks !== -1 ? String(r[idxRemarks] ?? '') : '';
      if (!map[cont]) map[cont] = { total:0, finished:0, remaining:0 };
      map[cont].total++;
      if (/done/i.test(rem)) map[cont].finished++;
      else map[cont].remaining++;
    });

    const analytics = Object.keys(map).sort().map(cont => {
      const v = map[cont];
      return { Container: cont, TotalBoxes: v.total, Finished: v.finished, Remaining: v.remaining, CompletionPercent: v.total===0?0:Math.round((v.finished/v.total)*100) };
    });
    const totals = analytics.reduce((acc,a)=> { acc.TotalBoxes+=a.TotalBoxes; acc.Finished+=a.Finished; acc.Remaining+=a.Remaining; return acc; }, { TotalBoxes:0, Finished:0, Remaining:0 });
    analytics.push({ Container:'ALL', ...totals, CompletionPercent: totals.TotalBoxes===0?0:Math.round((totals.Finished/totals.TotalBoxes)*100) });

    const wb = XLSX.utils.book_new();
    const wsData = XLSX.utils.json_to_sheet(objs);
    const wsAnalytics = XLSX.utils.json_to_sheet(analytics);
    XLSX.utils.book_append_sheet(wb, wsData, 'Data');
    XLSX.utils.book_append_sheet(wb, wsAnalytics, 'Analytics');

    const buf = XLSX.write(wb, { bookType:'xlsx', type:'buffer' });
    const fname = `${tab}_${(new Date()).toISOString().slice(0,16).replace('T','_').replace(':','-')}.xlsx`;
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (err) {
    console.error('export error', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Server listening on ${PORT}`));
