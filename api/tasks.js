const { timingSafeEqual } = require('node:crypto');
function equal(a, b) {
  const x = Buffer.from(String(a || '')), y = Buffer.from(String(b || ''));
  return x.length === y.length && timingSafeEqual(x, y);
}
module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (!['GET', 'POST'].includes(req.method)) return res.status(405).json({error:'Método no permitido.'});
  const { SHEETS_SCRIPT_URL: url, SHEETS_API_KEY: key, EDITOR_KEY: editor, PUBLIC_EDIT } = process.env;
  if (!url || !key) return res.status(503).json({error:'Falta configurar la conexión con Google Sheets en Vercel.'});
  let payload = {action:'list'};
  if (req.method === 'POST') {
    if (PUBLIC_EDIT !== 'true' && !editor) return res.status(503).json({error:'Falta configurar EDITOR_KEY en Vercel.'});
    if (PUBLIC_EDIT !== 'true' && !equal(req.headers['x-editor-key'], editor)) return res.status(401).json({error:'La clave de edición no es correcta.'});
    try { payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
    catch { return res.status(400).json({error:'Solicitud inválida.'}); }
    if (!payload || !['save','delete'].includes(payload.action) || JSON.stringify(payload).length > 25000) return res.status(400).json({error:'Solicitud inválida.'});
  }
  try {
    const upstream = await fetch(url, {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({...payload, key}), signal:AbortSignal.timeout(25000), redirect:'follow'
    });
    if (!upstream.ok) throw new Error('upstream');
    const data = await upstream.json();
    if (!data.ok) return res.status(data.code || 502).json({error:data.error || 'No se pudo guardar.'});
    return res.status(200).json({...data, requiresEditorKey:PUBLIC_EDIT !== 'true'});
  } catch {
    return res.status(502).json({error:'No se pudo confirmar la operación con Google Sheets. Actualiza antes de reintentar.'});
  }
};
