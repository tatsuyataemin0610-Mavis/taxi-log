const { google } = require('googleapis');
const SPREADSHEET_ID = '1mdHyJ7VUbfzC4TFSoMZUNfDuXYJmp9FdnDriCJJpopY';

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  return new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
}

async function getOrCreateSheet(sheets, yearMonth) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existing = spreadsheet.data.sheets.find(s => s.properties.title === yearMonth);
  if (existing) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: yearMonth, gridProperties: { rowCount: 70, columnCount: 9 } } } }] }
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${yearMonth}!A1:I1`,
    valueInputOption: 'RAW',
    requestBody: { values: [['日期','固定開銷','加油/ETC/其他','公里數','上線小時','實收','當日支出','淨利','備註']] }
  });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });

    if (req.method === 'POST') {
      const { date, fixed, fuel, km, hours, income } = req.body;
      const d = new Date(date);
      const yearMonth = `${d.getFullYear()}年${String(d.getMonth()+1).padStart(2,'0')}月`;
      const dayLabel  = `${d.getMonth()+1}月${d.getDate()}日`;
      const profit    = income - Number(fixed) - Number(fuel);

      await getOrCreateSheet(sheets, yearMonth);

      // 找下一個空行
      const existing = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${yearMonth}!A2:A65`,
      });
      const rows = existing.data.values || [];
      const nextRow = rows.length + 2;

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${yearMonth}!A${nextRow}:H${nextRow}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[dayLabel, fixed, fuel, km, hours, income, fuel, profit]] }
      });

      return res.status(200).json({ success: true, profit });
    }

    if (req.method === 'GET') {
      const { yearMonth } = req.query;
      if (!yearMonth) return res.status(400).json({ error: 'yearMonth required' });
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `${yearMonth}!A1:I65`,
        });
        return res.status(200).json({ rows: result.data.values || [] });
      } catch { return res.status(200).json({ rows: [] }); }
    }

    res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
