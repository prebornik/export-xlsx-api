// api/export-xlsx.js
// POST -> vrátí XLSX jako attachment
// Vstup: { rows: Array<Array<any>>, sheetName?: string, filename?: string }
// Podporuje application/json i application/x-www-form-urlencoded

import { utils, write } from 'xlsx';

/** Pomocná funkce: parse application/x-www-form-urlencoded do objektu */
async function parseFormUrlEncoded(request) {
  const text = await request.text();
  const params = new URLSearchParams(text);
  const payload = params.get('payload');

  // 1) Pokud přijde payload jako JSON string (časté u formulářů), zkus ho rozparsovat
  if (payload) {
    try { return JSON.parse(payload); } catch { return { payload }; }
  }

  // 2) Jinak převedeme všechny páry na objekt
  const obj = Object.fromEntries(params.entries());

  // Pokud je rows stringem, zkusíme ho rozparsovat jako JSON
  if (typeof obj.rows === 'string') {
    try { obj.rows = JSON.parse(obj.rows); } catch { /* noop */ }
  }

  return obj;
}

export async function POST(request) {
  try {
    // --- 1) Parsování vstupu -----------------------------------------------
    const ctype = (request.headers.get('content-type') || '').toLowerCase();
    let body = {};
    if (ctype.startsWith('application/json')) {
      body = await request.json();
    } else if (ctype.startsWith('application/x-www-form-urlencoded')) {
      body = await parseFormUrlEncoded(request);
    } else {
      // fallback: pokus o JSON
      try { body = await request.json(); } catch { body = {}; }
    }

    const {
      rows,
      sheetName = 'Výsledky',
      filename = 'vysledky.xlsx'
    } = body || {};

    // --- 2) Validace -------------------------------------------------------
    if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(rows[0])) {
      return Response.json(
        { error: 'Body musí obsahovat { rows: Array<Array<any>> } s minimálně jedním řádkem hlavičky.' },
        { status: 400 }
      );
    }

    // --- 3) Tvorba XLSX ----------------------------------------------------
    const wb = utils.book_new();
    const ws = utils.aoa_to_sheet(rows);

    // Šířky sloupců dle délky hlaviček (min 10 znaků)
    ws['!cols'] = rows[0].map(h => ({ wch: Math.max(10, String(h ?? '').length + 2) }));

    // Zapnout filtr v hlavičce (nad celým rozsahem)
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };

    // POZN.: Freeze panes (zamrznutí horního řádku) není v CE SheetJS podporováno.

    utils.book_append_sheet(wb, ws, String(sheetName).substring(0, 31));

    // Zápis do binárního bufferu
    const buf = write(wb, { bookType: 'xlsx', type: 'buffer' });

    // --- 4) Odpověď jako download -----------------------------------------
    // Bezpečný Content-Disposition pro diakritiku: ASCII fallback + RFC 5987 filename*
    const ascii = String(filename).replace(/[^A-Za-z0-9_.-]/g, '_');
    const headers = {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(filename)}"`,
      'Cache-Control': 'no-store'
    };

    return new Response(buf, { status: 200, headers });
  } catch (err) {
    return Response.json({ error: 'Export failed', detail: String(err?.message || err) }, { status: 500 });
  }
}
