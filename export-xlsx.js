// api/export-xlsx.js
// Vercel serverless funkce: POST -> vrátí XLSX jako attachment
// Kompatibilní s application/json i application/x-www-form-urlencoded (form POST)

const XLSX = require('xlsx');

async function getRawBody(req){
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => data += chunk);
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });
}

function parseBody(req, raw){
  const ct = (req.headers['content-type'] || '').toLowerCase();
  try {
    if (ct.startsWith('application/json')) {
      return JSON.parse(raw || '{}');
    }
    if (ct.startsWith('application/x-www-form-urlencoded')) {
      const params = new URLSearchParams(raw || '');
      const payload = params.get('payload');
      if (payload) {
        try { return JSON.parse(payload); } catch(_) { return { payload }; }
      }
      // nebo mapovat všechny klíče
      const obj = {};
      for (const [k,v] of params.entries()) obj[k] = v;
      return obj;
    }
    // fallback: zkusit JSON
    return raw ? JSON.parse(raw) : {};
  } catch(e) {
    return {};
  }
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).send('Method Not Allowed');
  }

  try {
    const raw = await getRawBody(req);
    const body = parseBody(req, raw);
    const { rows, sheetName = 'Výsledky', filename = 'vysledky.xlsx' } = body || {};

    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: 'Invalid rows' });
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);

    // Fixace hlavičky a přiměřené šířky sloupců
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    ws['!cols'] = rows[0].map(h => ({ wch: Math.max(10, String(h || '').length + 2) }));

    XLSX.utils.book_append_sheet(wb, ws, String(sheetName).substring(0,31));

    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('Cache-Control', 'no-store');

    return res.status(200).send(buf);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Export failed' });
  }
};
