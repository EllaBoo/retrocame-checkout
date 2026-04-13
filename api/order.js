const { google } = require('googleapis');

const SHEET_ID = process.env.SHEET_ID;
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const body = req.body;
    if (!body.email && !body.name) return res.status(400).json({ error: 'Missing required fields' });

    const auth = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET);
    auth.setCredentials({ refresh_token: REFRESH_TOKEN });

    const sheets = google.sheets({ version: 'v4', auth });

    // Ensure header exists
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheetNames = meta.data.sheets.map(s => s.properties.title);

    if (!sheetNames.includes('Заказы')) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        resource: { requests: [{ addSheet: { properties: { title: 'Заказы' } } }] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: 'Заказы!A1',
        valueInputOption: 'RAW',
        resource: { values: [['Дата','Имя','Email','Телефон','Страна','Регион','Город','Индекс','Адрес','Товары','Сумма','Доставка (₽)','Итого (₽)','Вес (г)','Комментарий','Статус']] }
      });
    }

    const now = new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' });
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: 'Заказы!A:P',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      resource: {
        values: [[
          now,
          body.name || '',
          body.email || '',
          body.phone || '',
          body.country || 'Россия',
          body.region || '',
          body.city || '',
          body.zip || '',
          body.address || '',
          body.products || '',
          parseFloat(body.subtotal) || 0,
          parseFloat(body.shipping_cost) || 0,
          parseFloat(body.total) || 0,
          parseInt(body.weight) || 0,
          body.comment || '',
          'Новый'
        ]]
      }
    });

    res.status(200).json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
