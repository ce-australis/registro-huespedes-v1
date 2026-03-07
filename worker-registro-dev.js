/**
 * CLOUDFLARE WORKER — ENTORNO DEV
 * Worker-dev: API gateway entre el frontend DEV y Google Apps Script DEV.
 *
 * VARIABLES DE ENTORNO (Cloudflare Dashboard → Worker → Settings → Variables):
 *   APPS_SCRIPT_URL  = https://script.google.com/macros/s/XXXXXX/exec
 *   ADMIN_TOKEN      = contraseña de acceso al panel de administración
 *
 * RUTAS:
 *   POST /           → formulario de registro (comportamiento original, sin cambios)
 *   GET  /api/registros      → admin: lista todos los registros
 *   GET  /api/registros/:id  → admin: detalle de un registro por nº de fila
 *   OPTIONS *        → CORS preflight
 */

export default {
  async fetch(request, env) {

    const CORS = {
      'Access-Control-Allow-Origin':  '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    };

    /* ── Preflight CORS ── */
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }

    /* ── Validación de configuración ── */
    if (!env.APPS_SCRIPT_URL) {
      return jsonErr(CORS, 500, 'Variable APPS_SCRIPT_URL no configurada en el Worker');
    }

    const url = new URL(request.url);
    const path = url.pathname;

    /* ════════════════════════════════════════════════════
       RUTAS ADMIN — GET /api/registros (y /:id)
       Autenticación por header: Authorization: Bearer {ADMIN_TOKEN}
    ════════════════════════════════════════════════════ */
    if (request.method === 'GET' && path.startsWith('/api/registros')) {

      /* Validar ADMIN_TOKEN */
      if (!env.ADMIN_TOKEN) {
        return jsonErr(CORS, 500, 'Variable ADMIN_TOKEN no configurada en el Worker');
      }
      const authHeader = request.headers.get('Authorization') || '';
      const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
      if (token !== env.ADMIN_TOKEN) {
        return jsonErr(CORS, 401, 'No autorizado');
      }

      /* Determinar acción: /api/registros vs /api/registros/42 */
      const segments = path.split('/').filter(Boolean); // ['api','registros','42']
      const id = segments[2] ? Number(segments[2]) : null;
      const action = id ? 'getRegistro' : 'getRegistros';

      /* Llamar a Apps Script con doGet */
      const gasUrl = new URL(env.APPS_SCRIPT_URL);
      gasUrl.searchParams.set('action', action);
      if (id) gasUrl.searchParams.set('id', String(id));

      try {
        const gsResponse = await fetch(gasUrl.toString(), { redirect: 'follow' });
        const text = await gsResponse.text();
        console.log('[DEV Worker] GET', action, '→ status:', gsResponse.status);

        let data;
        try { data = JSON.parse(text); }
        catch { data = { status: 'raw', rawResponse: text.substring(0, 500) }; }

        return new Response(JSON.stringify(data), {
          status: 200,
          headers: { ...CORS, 'Content-Type': 'application/json' },
        });
      } catch (err) {
        console.error('[DEV Worker] Error GET admin:', err.message);
        return jsonErr(CORS, 500, err.message);
      }
    }

    /* ════════════════════════════════════════════════════
       RUTA FORMULARIO — POST /
       Comportamiento original: proxy al Apps Script.
       SIN modificaciones respecto a la versión anterior.
    ════════════════════════════════════════════════════ */
    if (request.method === 'POST') {
      try {
        const body = await request.text();

        const gsResponse = await fetch(env.APPS_SCRIPT_URL, {
          method:   'POST',
          headers:  { 'Content-Type': 'text/plain' },
          body:     body,
          redirect: 'follow',
        });

        const text = await gsResponse.text();
        console.log('[DEV Worker] POST status:', gsResponse.status);
        console.log('[DEV Worker] POST response (500 chars):', text.substring(0, 500));

        let data;
        try { data = JSON.parse(text); }
        catch { data = { status: 'raw', rawResponse: text.substring(0, 500) }; }

        return new Response(JSON.stringify(data), {
          status:  200,
          headers: { ...CORS, 'Content-Type': 'application/json' },
        });
      } catch (err) {
        console.error('[DEV Worker] Error POST:', err.message);
        return jsonErr(CORS, 500, err.message);
      }
    }

    /* ── Ruta no encontrada ── */
    return jsonErr(CORS, 404, 'Ruta no encontrada');
  }
};

/* Helper: respuesta de error JSON */
function jsonErr(cors, status, message) {
  return new Response(
    JSON.stringify({ status: 'error', message }),
    { status, headers: { ...cors, 'Content-Type': 'application/json' } }
  );
}
