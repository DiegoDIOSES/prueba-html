const { timingSafeEqual, randomUUID } = require('node:crypto');
const { google } = require('googleapis');

const FIELDS = ['id','tarea','desc','modulo','tipo','responsable','prioridad','estado','sprint','fsolicitud','flimite','hest','hreal','comentarios','version','actualizado'];
const SHEET_NAME = 'Tareas';
const RANGE_ALL = `${SHEET_NAME}!A2:P`;
const RANGE_HEADER = `${SHEET_NAME}!A1:P1`;

function equal(a, b) {
  const x = Buffer.from(String(a || '')), y = Buffer.from(String(b || ''));
  return x.length === y.length && timingSafeEqual(x, y);
}

function fail(message, code) {
  const e = new Error(message);
  e.code = code;
  throw e;
}

let sheetsClientPromise = null;
function getSheetsClient() {
  if (!sheetsClientPromise) {
    const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
    const rawKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY || '';
    const key = rawKey.replace(/\\n/g, '\n');
    if (!email || !key) fail('Falta configurar GOOGLE_SERVICE_ACCOUNT_EMAIL o GOOGLE_SERVICE_ACCOUNT_KEY en Vercel.', 503);
    const auth = new google.auth.GoogleAuth({
      credentials: { client_email: email, private_key: key },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    sheetsClientPromise = auth.getClient().then((authClient) => google.sheets({ version: 'v4', auth: authClient }));
  }
  return sheetsClientPromise;
}

async function ensureSheet(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  let sheet = meta.data.sheets.find((s) => s.properties.title === SHEET_NAME);

  if (!sheet) {
    const created = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_NAME } } }] },
    });
    sheet = created.data.replies[0].addSheet;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: RANGE_HEADER,
      valueInputOption: 'RAW',
      requestBody: { values: [FIELDS] },
    });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: sheet.properties.sheetId, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.12, green: 0.23, blue: 0.37 },
                  textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true },
                },
              },
              fields: 'userEnteredFormat(backgroundColor,textFormat)',
            },
          },
          {
            updateSheetProperties: {
              properties: { sheetId: sheet.properties.sheetId, gridProperties: { frozenRowCount: 1 } },
              fields: 'gridProperties.frozenRowCount',
            },
          },
        ],
      },
    });
  } else {
    const headerRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: RANGE_HEADER });
    const header = (headerRes.data.values && headerRes.data.values[0]) || [];
    if (header.join('|') !== FIELDS.join('|')) {
      fail('Los encabezados de Tareas no coinciden. Conserva su orden y nombres.', 409);
    }
  }
  return sheet.properties.sheetId;
}

async function readRows(sheets, spreadsheetId) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: RANGE_ALL });
  const values = res.data.values || [];
  return values
    .map((row, i) => {
      const task = {};
      FIELDS.forEach((f, j) => (task[f] = row[j] || ''));
      task.hest = Number(task.hest) || 0;
      task.hreal = Number(task.hreal) || 0;
      return { task, row: i + 2 };
    })
    .filter((x) => x.task.id);
}

// Nombres asignables por defecto. El equipo puede crecer: cualquier nombre con
// formato razonable es aceptado (ver validación de "responsable" más abajo),
// esta lista solo documenta el equipo base.
const KNOWN_TEAM = ['Diego', 'Edith', 'Paula'];
const NAME_PATTERN = /^[\p{L}\p{M}0-9 .'-]{1,60}$/u;

function validateResponsable(raw) {
  const value = String(raw == null ? '' : raw).trim();
  if (!value) return '';
  const names = value.split(',').map((s) => s.trim()).filter(Boolean);
  if (names.length > 10) fail('Se pueden asignar como máximo 10 personas por tarea.', 400);
  names.forEach((n) => {
    if (!NAME_PATTERN.test(n)) fail('Nombre de responsable inválido: ' + n, 400);
  });
  // Normaliza duplicados y el formato de separación.
  return Array.from(new Set(names)).join(', ');
}

function validate(input) {
  if (!input || typeof input !== 'object') fail('Tarea inválida.', 400);
  const t = {};
  FIELDS.slice(0, 14).forEach((f) => {
    if (f === 'responsable') return; // se valida aparte, admite lista separada por comas
    if (['hest', 'hreal'].includes(f)) {
      t[f] = Number(input[f] || 0);
      if (!Number.isFinite(t[f]) || t[f] < 0 || t[f] > 100000) fail('Horas inválidas.', 400);
    } else {
      t[f] = String(input[f] == null ? '' : input[f]).trim();
      if (t[f].length > (['desc', 'comentarios'].includes(f) ? 5000 : 300)) fail('Texto demasiado largo.', 400);
    }
  });
  t.responsable = validateResponsable(input.responsable);
  if (t.responsable.length > 300) fail('Texto demasiado largo.', 400);
  if (!/^[a-zA-Z0-9_-]{1,80}$/.test(t.id) || !t.tarea) fail('Falta el nombre o identificador de la tarea.', 400);
  const enums = {
    prioridad: ['Alta', 'Media', 'Baja'],
    estado: ['Por hacer', 'En progreso', 'En revisión', 'Bloqueado', 'Hecho'],
    tipo: ['Bug', 'Mejora', 'Nueva funcionalidad', 'Consulta / Soporte', 'Reunión'],
  };
  Object.keys(enums).forEach((f) => {
    if (!enums[f].includes(t[f])) fail('Valor inválido: ' + f, 400);
  });
  ['fsolicitud', 'flimite'].forEach((f) => {
    if (
      t[f] &&
      (!/^\d{4}-\d{2}-\d{2}$/.test(t[f]) ||
        !Number.isFinite(Date.parse(t[f])) ||
        new Date(t[f]).toISOString().slice(0, 10) !== t[f])
    ) {
      fail('Fecha inválida.', 400);
    }
  });
  return t;
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (!['GET', 'POST'].includes(req.method)) return res.status(405).json({ error: 'Método no permitido.' });

  const { SPREADSHEET_ID: spreadsheetId, EDITOR_KEY: editor, PUBLIC_EDIT } = process.env;
  if (!spreadsheetId) return res.status(503).json({ error: 'Falta configurar SPREADSHEET_ID en Vercel.' });

  let payload = { action: 'list' };
  if (req.method === 'POST') {
    if (PUBLIC_EDIT !== 'true' && !editor) return res.status(503).json({ error: 'Falta configurar EDITOR_KEY en Vercel.' });
    if (PUBLIC_EDIT !== 'true' && !equal(req.headers['x-editor-key'], editor)) {
      return res.status(401).json({ error: 'La clave de edición no es válida.' });
    }
    try {
      payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    } catch {
      return res.status(400).json({ error: 'Solicitud inválida.' });
    }
    if (!payload || !['save', 'delete'].includes(payload.action) || JSON.stringify(payload).length > 25000) {
      return res.status(400).json({ error: 'Acción inválida.' });
    }
  }

  try {
    const sheets = await getSheetsClient();
    const sheetId = await ensureSheet(sheets, spreadsheetId);
    let rows = await readRows(sheets, spreadsheetId);

    if (payload.action === 'list') {
      return res.status(200).json({ ok: true, tasks: rows.map((x) => x.task), requiresEditorKey: PUBLIC_EDIT !== 'true' });
    }

    const id = payload.action === 'save' ? payload.task && payload.task.id : payload.id;
    const existing = rows.find((x) => x.task.id === id);

    if (existing && existing.task.version !== String(payload.version || '')) {
      return res.status(409).json({ error: 'Otra persona modificó esta tarea. Cierra el formulario, actualiza y vuelve a abrirla.' });
    }
    if (!existing && payload.version) {
      return res.status(409).json({ error: 'La tarea ya fue eliminada. Actualiza la página.' });
    }

    if (payload.action === 'delete') {
      if (existing) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [
              {
                deleteDimension: {
                  range: { sheetId, dimension: 'ROWS', startIndex: existing.row - 1, endIndex: existing.row },
                },
              },
            ],
          },
        });
      }
    } else {
      const t = validate(payload.task);
      t.version = randomUUID();
      t.actualizado = new Date().toISOString();
      const values = FIELDS.map((f) => String(t[f] ?? ''));
      const targetRow = existing ? existing.row : rows.length + 2;
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${SHEET_NAME}!A${targetRow}:P${targetRow}`,
        valueInputOption: 'RAW',
        requestBody: { values: [values] },
      });
    }

    rows = await readRows(sheets, spreadsheetId);
    return res.status(200).json({ ok: true, tasks: rows.map((x) => x.task), requiresEditorKey: PUBLIC_EDIT !== 'true' });
  } catch (err) {
    if (err.code && Number.isInteger(err.code)) return res.status(err.code).json({ error: err.message });
    console.error('Sheets API error:', (err.errors && JSON.stringify(err.errors)) || err.message || err);
    return res.status(502).json({ error: 'No se pudo completar la operación en Google Sheets. Actualiza antes de reintentar.' });
  }
};
