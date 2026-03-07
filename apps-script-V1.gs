/**
 * ============================================================
 *  REGISTRO DE HUÉSPEDES — Ático Marbella Centro
 *  Google Apps Script (Web App) — ENTORNO DEV
 *  v3 — Soporte de múltiples huéspedes por reserva
 * ============================================================
 *
 *  CAMBIOS v3 respecto a v2:
 *  - doPost ahora acepta payload { token, reserva, huespedes[] }
 *  - Se genera ReservaID (R + 4 dígitos) y CodigoCaja (4 dígitos)
 *  - Una carpeta Drive por reserva (en vez de por huésped)
 *  - Una fila en Sheets por huésped, todas con el mismo ReservaID
 *  - HEADERS actualizado: añadidos ReservaID, Código Caja, Fecha Expedición
 *    y eliminados campos de dirección (Tipo Vía, Nombre Vía, etc.)
 *  - Email de notificación actualizado con ReservaID, CodigoCaja y nº huéspedes
 *
 *  SETUP (solo si es un nuevo despliegue):
 *  1. Ejecutar autorizar() → aceptar todos los permisos
 *  2. Implementar → Nueva implementación (NO actualizar)
 *     Tipo: Aplicación web · Ejecutar como: Yo · Acceso: Cualquier persona
 *  3. Copiar la URL /exec → actualizar APPS_SCRIPT_URL en Worker-dev (Cloudflare)
 *
 * ============================================================
 */

const SPREADSHEET_ID    = '1cvVZ0WUF-lDxa5KjpTWCTMbOAMlGBF5RhCwNj2PqRGg'; // Sheet
const SHEET_NAME        = 'Registros';
const API_TOKEN         = 'AtMb2025!xK9#qR7vL';
const DRIVE_FOLDER_ID   = '1SluqT8ZB-DJuLNlSK2_2JKj5PovbM7tZ';              // Carpeta raíz en Drive
const DRIVE_FOLDER_NAME = 'REGISTRO HUÉSPEDES_2026';              // (fallback por nombre)
const EMAIL_NOTIFICACION = "ce.australis@gmail.com,aticomarbellacentro@gmail.com";
/* ── Cabeceras de la hoja (21 columnas) ── */
const HEADERS = [
  'Timestamp',          // col  1
  'ReservaID',          // col  2
  'Código Caja',        // col  3
  'Nombre',             // col  4
  'Primer Apellido',    // col  5
  'Segundo Apellido',   // col  6
  'Fecha Nacimiento',   // col  7
  'Sexo',               // col  8
  'Nacionalidad',       // col  9
  'Tipo Documento',     // col 10
  'Nº Documento',       // col 11
  'Fecha Expedición',   // col 12
  'Nº Soporte',         // col 13
  'Fecha Entrada',      // col 14
  'Fecha Salida',       // col 15
  'Teléfono',           // col 16
  'Email',              // col 17
  'Idioma',             // col 18
  'Carpeta Documentos', // col 19
  'Foto Anverso',       // col 20
  'Foto Reverso',       // col 21
];

/* ══════════════════════════════════════════════════
   doGet — endpoints de lectura para el panel admin
   (El Worker ya validó el ADMIN_TOKEN; aquí solo lógica.)
══════════════════════════════════════════════════ */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;

    if (action === 'getRegistros') return getRegistros();

    if (action === 'getRegistro') {
      const id = Number(e.parameter.id);
      if (!id || id < 2) return jsonResponse({ status: 'error', message: 'ID inválido' });
      return getRegistro(id);
    }

    if (action === 'debug') return debugSheet();

    return jsonResponse({ status: 'ok', message: 'API DEV activa — v3 multi-huésped.' });

  } catch (err) {
    Logger.log('ERROR en doGet: ' + err.toString());
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/* Devuelve todos los registros de la Sheet como array JSON */
function getRegistros() {
  const ss               = SpreadsheetApp.openById(SPREADSHEET_ID);
  const todasLasPestanas = ss.getSheets().map(s => s.getName());
  Logger.log('Pestañas en el spreadsheet: ' + JSON.stringify(todasLasPestanas));

  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('Pestaña "' + SHEET_NAME + '" NO encontrada.');
    return jsonResponse({
      status: 'error',
      message: 'Pestaña "' + SHEET_NAME + '" no encontrada',
      debug: { spreadsheetId: SPREADSHEET_ID, sheetNameBuscado: SHEET_NAME, pestanasDisponibles: todasLasPestanas },
    });
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  Logger.log('Sheet encontrada. lastRow=' + lastRow + ', lastCol=' + lastCol);

  if (lastRow < 2) {
    return jsonResponse({ status: 'ok', total: 0, registros: [], debug: { lastRow, lastCol, sheetName: sheet.getName() } });
  }

  const values   = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const registros = values.map((row, i) => filaAObjeto(row, i + 2));
  Logger.log('Registros devueltos: ' + registros.length);
  return jsonResponse({ status: 'ok', total: registros.length, registros });
}

/* Devuelve un único registro por número de fila (empieza en 2) */
function getRegistro(id) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || id > sheet.getLastRow()) {
    return jsonResponse({ status: 'error', message: 'Registro no encontrado' });
  }
  const lastCol = sheet.getLastColumn();
  const row     = sheet.getRange(id, 1, 1, lastCol).getValues()[0];
  return jsonResponse({ status: 'ok', registro: filaAObjeto(row, id) });
}

/* Mapea una fila de Sheets a un objeto con claves legibles — v3 (21 columnas) */
function filaAObjeto(row, filaNum) {
  return {
    id:              filaNum,
    timestamp:       row[0]  ? Utilities.formatDate(new Date(row[0]), 'Europe/Madrid', 'yyyy-MM-dd HH:mm:ss') : '',
    reservaId:       String(row[1]  || ''),
    codigoCaja:      String(row[2]  || ''),
    nombre:          String(row[3]  || ''),
    apellido1:       String(row[4]  || ''),
    apellido2:       String(row[5]  || ''),
    fechaNacimiento: String(row[6]  || ''),
    sexo:            String(row[7]  || ''),
    nacionalidad:    String(row[8]  || ''),
    tipoDocumento:   String(row[9]  || ''),
    numeroDocumento: String(row[10] || ''),
    fechaExpedicion: String(row[11] || ''),
    numeroSoporte:   String(row[12] || ''),
    fechaEntrada:    String(row[13] || ''),
    fechaSalida:     String(row[14] || ''),
    telefono:        String(row[15] || ''),
    email:           String(row[16] || ''),
    idioma:          String(row[17] || ''),
    carpetaDrive:    String(row[18] || ''),
    fotoAnverso:     String(row[19] || ''),
    fotoReverso:     String(row[20] || ''),
  };
}

/* ══════════════════════════════════════════════════
   doPost — recibe la reserva completa con N huéspedes
   Payload esperado: { token, reserva: {...}, huespedes: [{...}, ...] }
══════════════════════════════════════════════════ */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents)
      ? e.postData.contents
      : (e && e.parameter && e.parameter.data)
        ? e.parameter.data
        : null;

    if (!raw) return jsonResponse({ status: 'error', message: 'No data' });

    const data = JSON.parse(raw);

    if (data.token !== API_TOKEN) {
      return jsonResponse({ status: 'error', message: 'Token inválido' });
    }

    const reserva   = data.reserva;
    const huespedes = data.huespedes;

    if (!reserva || !Array.isArray(huespedes) || huespedes.length === 0) {
      return jsonResponse({ status: 'error', message: 'Payload inválido: falta reserva o huespedes[]' });
    }

    /* ── Generar IDs de reserva ── */
    const reservaId  = 'R' + Math.floor(1000 + Math.random() * 9000);
    const codigoCaja = String(Math.floor(1000 + Math.random() * 9000));

    /* ── Sheets ── */
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    /* ── Drive: una carpeta por reserva ── */
    const rootFolder    = getRootFolder();
    const primerH       = huespedes[0];
    const folderName    = reservaId + '_' + sanitize(primerH.nombre) + '_' + sanitize(primerH.apellido1);
    const reservaFolder = obtenerOCrearCarpeta(folderName, rootFolder);
    const carpetaUrl    = reservaFolder.getUrl();

    Logger.log('[doPost] ReservaID=' + reservaId + ' CodigoCaja=' + codigoCaja + ' Huéspedes=' + huespedes.length);

    /* ── Procesar cada huésped: guardar fotos + añadir fila ── */
    const filasGuardadas = [];

    for (let i = 0; i < huespedes.length; i++) {
      const h = huespedes[i];
      let linkFrontal = '';
      let linkTrasero = '';

      if (h.fotoFrontalB64 && h.fotoFrontalMime) {
        const ext  = extensionDeMime(h.fotoFrontalMime);
        const file = guardarImagen(reservaFolder, 'huesped' + (i + 1) + '_anverso' + ext, h.fotoFrontalB64, h.fotoFrontalMime);
        if (file) linkFrontal = file.getUrl();
      }
      if (h.fotoTraseroB64 && h.fotoTraseroMime) {
        const ext  = extensionDeMime(h.fotoTraseroMime);
        const file = guardarImagen(reservaFolder, 'huesped' + (i + 1) + '_reverso' + ext, h.fotoTraseroB64, h.fotoTraseroMime);
        if (file) linkTrasero = file.getUrl();
      }

      const row = [
        new Date(),
        reservaId,
        codigoCaja,
        clean(h.nombre),
        clean(h.apellido1),
        clean(h.apellido2),
        clean(h.fechaNacimiento),
        clean(h.sexo),
        clean(h.nacionalidad),
        clean(h.tipoDocumento),
        clean(h.numeroDocumento),
        clean(h.fechaExpedicion),
        clean(h.numeroSoporte),
        clean(reserva.fechaEntrada),
        clean(reserva.fechaSalida),
        clean(reserva.telefono),
        clean(reserva.email),
        clean(reserva.idioma || ''),
        carpetaUrl,
        linkFrontal,
        linkTrasero,
      ];

      sheet.appendRow(row);
      filasGuardadas.push({ huesped: i + 1, nombre: clean(h.nombre) + ' ' + clean(h.apellido1) });
      Logger.log('[doPost] Fila guardada: huésped ' + (i + 1) + ' — ' + clean(h.nombre));
    }

    /* ── Email de notificación ──
     * Aislado en su propio try/catch: si falla NO aborta el doPost.
     * El resultado (enviado: true/false) se devuelve en la respuesta JSON
     * y aparece en el panel DEBUG del formulario.
     */
    let emailResult;
    try {
      enviarNotificacion(reservaId, codigoCaja, reserva, huespedes);
      emailResult = { enviado: true };
    } catch (emailErr) {
      Logger.log('ERROR en enviarNotificacion: ' + emailErr.message);
      emailResult = { enviado: false, error: emailErr.message };
    }

    let emailHuespedResult;
    try {
      enviarEmailHuesped(reservaId, codigoCaja, reserva, huespedes);
      emailHuespedResult = { enviado: true };
    } catch (emailErr) {
      Logger.log('ERROR en enviarEmailHuesped: ' + emailErr.message);
      emailHuespedResult = { enviado: false, error: emailErr.message };
    }

    return jsonResponse({
      status:       'ok',
      reservaId,
      codigoCaja,
      huespedes:    filasGuardadas.length,
      filas:        filasGuardadas,
      email:        emailResult,
      emailHuesped: emailHuespedResult,
    });

  } catch (err) {
    Logger.log('ERROR en doPost: ' + err.toString());
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/* ── Notificación por email ── */
function enviarNotificacion(reservaId, codigoCaja, reserva, huespedes) {
  const primerH  = huespedes[0];
  const numH     = huespedes.length;
  const sufijo   = numH > 1 ? ' +' + (numH - 1) + ' más' : '';
  const asunto   = 'Reserva ' + reservaId + ': ' + clean(primerH.nombre) + ' ' + clean(primerH.apellido1) + sufijo;

  /* Datos completos requeridos por el sistema SES del Ministerio del Interior */
  let detalleHuespedes = '';
  huespedes.forEach(function(h, i) {
    detalleHuespedes +=
      '\nHuésped ' + (i + 1) + ':\n' +
      '  Nombre completo : ' + clean(h.nombre) + ' ' + clean(h.apellido1) + (h.apellido2 ? ' ' + clean(h.apellido2) : '') + '\n' +
      '  Fecha nacimiento: ' + clean(h.fechaNacimiento) + '\n' +
      '  Sexo            : ' + clean(h.sexo)            + '\n' +
      '  Nacionalidad    : ' + clean(h.nacionalidad)    + '\n' +
      '  Tipo documento  : ' + clean(h.tipoDocumento)   + '\n' +
      '  Nº documento    : ' + clean(h.numeroDocumento) + '\n' +
      '  Fecha expedición: ' + clean(h.fechaExpedicion) + '\n' +
      '  Nº soporte      : ' + clean(h.numeroSoporte)   + '\n';
  });

  const cuerpo =
    '════════════════════════════════\n' +
    '  NUEVA RESERVA \n' +
    '  ReservaID  : ' + reservaId  + '\n' +
    '  Código caja: ' + codigoCaja + '\n' +
    '  Huéspedes  : ' + numH       + '\n' +
    '════════════════════════════════\n\n' +
    'ESTANCIA:\n' +
    '  Entrada : ' + clean(reserva.fechaEntrada) + '\n' +
    '  Salida  : ' + clean(reserva.fechaSalida)  + '\n\n' +
    'CONTACTO:\n' +
    '  Teléfono: ' + clean(reserva.telefono) + '\n' +
    '  Email   : ' + clean(reserva.email)    + '\n\n' +
    'HUÉSPEDES (datos SES):' + detalleHuespedes + '\n';

  MailApp.sendEmail(EMAIL_NOTIFICACION, asunto, cuerpo);
}

/* ── Email de bienvenida al huésped ── */
function enviarEmailHuesped(reservaId, codigoCaja, reserva, huespedes) {
  const primerH      = huespedes[0];
  const nombre       = clean(primerH.nombre);
  const apellido1    = clean(primerH.apellido1);
  const fechaEntrada = clean(reserva.fechaEntrada);
  const fechaSalida  = clean(reserva.fechaSalida);
  const idioma       = clean(reserva.idioma || 'es');

  const htmlES = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0; padding:0; font-family:Arial, Helvetica, sans-serif; background:#e0e0e0;">

  <div style="background-color:#f2f2f2; padding:30px 16px; font-family:Arial, Helvetica, sans-serif;">

    <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
           style="max-width:600px; margin:0 auto; background:#ffffff;
                  border-radius:10px; overflow:hidden;
                  box-shadow:0 2px 12px rgba(0,0,0,0.10);">

      <!-- ── HEADER ─────────────────────────────── -->
      <tr>
        <td style="background-color:#012030; padding:36px 32px 30px; text-align:center;">
       <h1 style="margin:0; color:#ffffff;
           font-family:'Rozha One', Georgia, serif;
           font-size:26px;
           font-weight:normal;
           letter-spacing:0.06em;
           line-height:1.25;">
  Ático Marbella Centro
</h1>
          <p style="margin:8px 0 0; color:rgba(255,255,255,0.60);
                    font-size:11px; letter-spacing:0.14em; text-transform:uppercase;">
            Instrucciones de llegada
          </p>
        </td>
      </tr>

      <!-- ── WELCOME ─────────────────────────────── -->
      <tr>
        <td style="padding:32px 32px 0;">
          <p style="margin:0; font-size:17px; color:#012030; font-weight:bold; line-height:1.4;">
            Hola, <span style="color:#C4724A;">${nombre}</span> 👋🏻
          </p>
          <p style="margin:12px 0 0; font-size:15px; color:#444444; line-height:1.7;">
            Gracias por rellenar el formulario de registro. Te damos la bienvenida y te presentamos a continuación toda la información para el acceso al apartamento.
          </p>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

      <!-- ── RESUMEN DE ESTANCIA ─────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            📋 &nbsp;Tu reserva
          </p>

          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="background:#f7f9fb; border-radius:8px; border:1px solid #e2e8ed;
                        overflow:hidden; font-size:14px;">
            <tr>
              <td style="padding:12px 16px; color:#777777; width:40%; border-bottom:1px solid #e2e8ed;">
                Nombre
              </td>
              <td style="padding:12px 16px; color:#012030; font-weight:bold; border-bottom:1px solid #e2e8ed;">
                ${nombre} ${apellido1}
              </td>
            </tr>
            <tr>
              <td style="padding:12px 16px; color:#777777; border-bottom:1px solid #e2e8ed;">
                Check-in
              </td>
              <td style="padding:12px 16px; color:#012030; font-weight:bold; border-bottom:1px solid #e2e8ed;">
                ${fechaEntrada}
              </td>
            </tr>
            <tr>
              <td style="padding:12px 16px; color:#777777;">
                Check-out
              </td>
              <td style="padding:12px 16px; color:#012030; font-weight:bold;">
                ${fechaSalida}<br> antes de las 11:00
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

      <!-- ── DIRECCIÓN ───────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            📍 &nbsp;Dirección
          </p>
          <p style="margin:0 0 4px; font-size:16px; color:#1a1a1a; font-weight:bold; line-height:1.4;">
            Calle Jacinto Benavente, 8
          </p>
          <p style="margin:0 0 18px; font-size:14px; color:#555555; line-height:1.5;">
            Edificio Marbelsun III<br>
            8º planta · puerta 3
          </p>
          <a href="https://maps.app.goo.gl/pRVaVEH8yxymNwq99"
           target="_blank"
   style="display:inline-flex; align-items:center; justify-content:center; gap:8px;
          background-color:#012030; color:#ffffff;
          font-size:14px; font-weight:bold; text-decoration:none;
          padding:12px 24px; border-radius:6px; letter-spacing:0.03em;">
  ↗️
  Abrir en Google Maps
</a>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

      <!-- ── CAJA DE LLAVES ──────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            🔑 &nbsp;Acceso · Caja de llaves
          </p>
          <p style="margin:0 0 14px; font-size:14px; color:#444444; line-height:1.7;">
            El acceso al apartamento se realiza de forma autónoma mediante una caja de seguridad con código. Sigue estos pasos:
          </p>

          <!-- Paso 1 -->
          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="margin-bottom:10px;">
            <tr>
              <td style="vertical-align:top; width:32px;">
                <div style="background:#012030; color:#ffffff; width:24px; height:24px;
                            border-radius:50%; text-align:center; line-height:24px;
                            font-size:12px; font-weight:bold;">1</div>
              </td>
              <td style="padding-left:10px; font-size:14px; color:#444444; line-height:1.6;">
                Al llegar al portal, justo enfrente encontrarás una escalera con barandilla metálica. La <strong>caja de llaves</strong> está colgada de esa barandilla,(mirando desde el portal).<br>
                <a href="https://photos.app.goo.gl/dW7efj3aK2EZeHm1A">📷 Ver foto de la ubicación</a>
              </td>
            </tr>
          </table>

          <!-- Paso 2 -->
          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="margin-bottom:10px;">
            <tr>
              <td style="vertical-align:top; width:32px;">
                <div style="background:#012030; color:#ffffff; width:24px; height:24px;
                            border-radius:50%; text-align:center; line-height:24px;
                            font-size:12px; font-weight:bold;">2</div>
              </td>
              <td style="padding-left:10px; font-size:14px; color:#444444; line-height:1.6;">
                Introduce el <strong>código de acceso</strong>: <span style="background:#f0f0f0;
                padding:2px 8px; border-radius:4px; font-family:monospace; font-size:14px;
                color:#012030; font-weight:bold;">${codigoCaja}</span>
              </td>
            </tr>
          </table>

          <!-- Paso 3 -->
          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="margin-bottom:10px;">
            <tr>
              <td style="vertical-align:top; width:32px;">
                <div style="background:#012030; color:#ffffff; width:24px; height:24px;
                            border-radius:50%; text-align:center; line-height:24px;
                            font-size:12px; font-weight:bold;">3</div>
              </td>
              <td style="padding-left:10px; font-size:14px; color:#444444; line-height:1.6;">
                Dentro encontrarás un llavero con:<br>
                🔵 <strong>Llave / chip azul</strong> → acceso al portal del edificio<br>
                🔑 <strong>Llave grande</strong> → acceso a la puerta del apartamento
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>


      <!-- ── ASCENSOR ────────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            🛗 &nbsp;Ascensor
          </p>
          <p style="margin:0 0 10px; font-size:14px; color:#444444; line-height:1.7;">
            Al entrar al portal, los ascensores para subir al apartamento se encuentran inmediatamente a mano izquierda.<br>
            Ten en cuenta que en el edificio hay otros ascensores/montacargas; utiliza únicamente estos para acceder a la vivienda.
          </p>
          <p style="margin:0; font-size:14px; color:#444444; line-height:1.7;">
            Una vez en el ascensor, sube hasta la <strong>planta 8</strong>. El apartamento es la <strong>puerta nº 3</strong>.
          </p>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>
 <!-- ── WIFI ────────────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            📶 &nbsp;WiFi
          </p>

          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="background:#f7f9fb; border-radius:8px; border:1px solid #e2e8ed;
                        overflow:hidden; font-size:14px;">
            <tr>
              <td style="padding:12px 16px; color:#777777; width:40%; border-bottom:1px solid #e2e8ed;">
                Red
              </td>
              <td style="padding:12px 16px; font-family:monospace; color:#012030;
                         font-weight:bold; font-size:15px; border-bottom:1px solid #e2e8ed;">
                MarbelsunWifi
              </td>
            </tr>
            <tr>
              <td style="padding:12px 16px; color:#777777;">
                Contraseña
              </td>
              <td style="padding:12px 16px; font-family:monospace; color:#012030;
                         font-weight:bold; font-size:15px;">
                marbella2026
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>
      <!-- ── CONTACTO ────────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            💬 &nbsp;¿Necesitas ayuda?
          </p>
          <p style="margin:0 0 10px; font-size:14px; color:#444444; line-height:1.7;">
            Si tienes cualquier duda o surge algún problema durante la llegada, no dudes en llamarnos.
          </p>
        <p style="margin:0; font-size:15px; color:#444444; line-height:1.6;">
  <a href="tel:+34611164242" style="color:#012030; text-decoration:none; font-weight:bold;">
    +34 611 164 242
  </a>
</p>
        </td>
      </tr>

      <!-- ── FIRMA ───────────────────────────────── -->
      <tr>
        <td style="padding:28px 32px 32px;">
          <p style="margin:0; font-size:15px; color:#444444; line-height:1.7;">
            ¡Os deseamos una feliz estancia!
          </p>
          <p style="margin:8px 0 0; font-size:16px; color:#012030; font-weight:bold;">
            Andrés
          </p>
          <p style="margin:4px 0 0; font-size:13px; color:#888888;">
            Ático Marbella Centro
          </p>
        </td>
      </tr>

      <!-- ── FOOTER ──────────────────────────────── -->
      <tr>
        <td style="background-color:#f7f9fb; border-top:1px solid #e8e8e8;
                   padding:18px 32px; text-align:center;">
          <p style="margin:0; font-size:11px; color:#aaaaaa; line-height:1.6;">
            Ático Marbella Centro · Calle Jacinto Benavente, 8 · Marbella, Málaga<br>
            Este correo ha sido enviado automáticamente. Por favor, no respondas directamente a este mensaje.
          </p>
        </td>
      </tr>

    </table>

  </div>

</body>
</html>`;

  const htmlEN = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0; padding:0; font-family:Arial, Helvetica, sans-serif; background:#e0e0e0;">

  <div style="background-color:#f2f2f2; padding:30px 16px; font-family:Arial, Helvetica, sans-serif;">

    <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
           style="max-width:600px; margin:0 auto; background:#ffffff;
                  border-radius:10px; overflow:hidden;
                  box-shadow:0 2px 12px rgba(0,0,0,0.10);">

      <!-- ── HEADER ─────────────────────────────── -->
      <tr>
        <td style="background-color:#012030; padding:36px 32px 30px; text-align:center;">
          <h1 style="margin:0; color:#ffffff; font-family:Georgia, 'Times New Roman', serif;
                     font-size:24px; font-weight:normal; letter-spacing:0.04em; line-height:1.3;">
            Ático Marbella Centro
          </h1>
          <p style="margin:8px 0 0; color:rgba(255,255,255,0.60);
                    font-size:11px; letter-spacing:0.14em; text-transform:uppercase;">
            Arrival Instructions
          </p>
        </td>
      </tr>

      <!-- ── WELCOME ─────────────────────────────── -->
      <tr>
        <td style="padding:32px 32px 0;">
          <p style="margin:0; font-size:17px; color:#012030; font-weight:bold; line-height:1.4;">
            Hello, <span style="color:#C4724A;">${nombre}</span> 👋🏻
          </p>
          <p style="margin:12px 0 0; font-size:15px; color:#444444; line-height:1.7;">
            Thank you for completing the registration form.<br>
            Below you will find all the information you need to access the apartment.
          </p>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

      <!-- ── RESUMEN DE ESTANCIA ─────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            📋 &nbsp;Your booking
          </p>

          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="background:#f7f9fb; border-radius:8px; border:1px solid #e2e8ed;
                        overflow:hidden; font-size:14px;">
            <tr>
              <td style="padding:12px 16px; color:#777777; width:40%; border-bottom:1px solid #e2e8ed;">
                Name
              </td>
              <td style="padding:12px 16px; color:#012030; font-weight:bold; border-bottom:1px solid #e2e8ed;">
                ${nombre} ${apellido1}
              </td>
            </tr>
            <tr>
              <td style="padding:12px 16px; color:#777777; border-bottom:1px solid #e2e8ed;">
                Check-in
              </td>
              <td style="padding:12px 16px; color:#012030; font-weight:bold; border-bottom:1px solid #e2e8ed;">
                ${fechaEntrada}
              </td>
            </tr>
            <tr>
              <td style="padding:12px 16px; color:#777777;">
                Check-out
              </td>
              <td style="padding:12px 16px; color:#012030; font-weight:bold;">
                ${fechaSalida}<br>before 11:00
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

      <!-- ── DIRECCIÓN ───────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            📍 &nbsp;Address
          </p>
          <p style="margin:0 0 4px; font-size:16px; color:#1a1a1a; font-weight:bold; line-height:1.4;">
            Calle Jacinto Benavente, 8
          </p>
          <p style="margin:0 0 18px; font-size:14px; color:#555555; line-height:1.5;">
            Edificio Marbelsun III<br>
            8th floor · apartment 3
          </p>
             <a href="https://maps.app.goo.gl/pRVaVEH8yxymNwq99"
              target="_blank"
   style="display:inline-flex; align-items:center; justify-content:center; gap:8px;
          background-color:#012030; color:#ffffff;
          font-size:14px; font-weight:bold; text-decoration:none;
          padding:12px 24px; border-radius:6px; letter-spacing:0.03em;">
  ↗️
  Open in Google Maps
</a>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

      <!-- ── CAJA DE LLAVES ──────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            🔑 &nbsp;Access · Key Box
          </p>
          <p style="margin:0 0 14px; font-size:14px; color:#444444; line-height:1.7;">
            Access to the apartment is self-managed using a coded lock box.<br>
            Follow these steps:
          </p>

          <!-- Paso 1 -->
          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="margin-bottom:10px;">
            <tr>
              <td style="vertical-align:top; width:32px;">
                <div style="background:#012030; color:#ffffff; width:24px; height:24px;
                            border-radius:50%; text-align:center; line-height:24px;
                            font-size:12px; font-weight:bold;">1</div>
              </td>
              <td style="padding-left:10px; font-size:14px; color:#444444; line-height:1.6;">
                When you arrive at the building entrance, right in front of you there is a <strong>staircase with a metal railing</strong>.<br>
                The <strong>key box</strong> hangs from the railing (as seen from the entrance).<br>
                <a href="https://photos.app.goo.gl/dW7efj3aK2EZeHm1A">📷 View photo of the location</a>
              </td>
            </tr>
          </table>

          <!-- Paso 2 -->
          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="margin-bottom:10px;">
            <tr>
              <td style="vertical-align:top; width:32px;">
                <div style="background:#012030; color:#ffffff; width:24px; height:24px;
                            border-radius:50%; text-align:center; line-height:24px;
                            font-size:12px; font-weight:bold;">2</div>
              </td>
              <td style="padding-left:10px; font-size:14px; color:#444444; line-height:1.6;">
                Enter the <strong>access code</strong>: <span style="background:#f0f0f0;
                padding:2px 8px; border-radius:4px; font-family:monospace; font-size:14px;
                color:#012030; font-weight:bold;">${codigoCaja}</span>
              </td>
            </tr>
          </table>

          <!-- Paso 3 -->
          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="margin-bottom:10px;">
            <tr>
              <td style="vertical-align:top; width:32px;">
                <div style="background:#012030; color:#ffffff; width:24px; height:24px;
                            border-radius:50%; text-align:center; line-height:24px;
                            font-size:12px; font-weight:bold;">3</div>
              </td>
              <td style="padding-left:10px; font-size:14px; color:#444444; line-height:1.6;">
                Inside you will find a keyring with:<br>
                🔵 <strong>Blue chip</strong> → opens the building entrance<br>
                🔑 <strong>Large key</strong> → opens the apartment door
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

     
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>

      <!-- ── ASCENSOR ────────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            🛗 &nbsp;Elevator
          </p>
          <p style="margin:0 0 10px; font-size:14px; color:#444444; line-height:1.7;">
            As you enter the building, the elevators to the apartment are immediately on your left.<br>
            Please note there are other lifts/service elevators in the building — use only these to access the apartment.
          </p>
          <p style="margin:0; font-size:14px; color:#444444; line-height:1.7;">
            Once in the elevator, go up to <strong>floor 8</strong>. The apartment is <strong>door number 3</strong>.
          </p>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
        </td>
      </tr>
 <!-- ── WIFI ────────────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            📶 &nbsp;WiFi
          </p>

          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                 style="background:#f7f9fb; border-radius:8px; border:1px solid #e2e8ed;
                        overflow:hidden; font-size:14px;">
            <tr>
              <td style="padding:12px 16px; color:#777777; width:40%; border-bottom:1px solid #e2e8ed;">
                Network
              </td>
              <td style="padding:12px 16px; font-family:monospace; color:#012030;
                         font-weight:bold; font-size:15px; border-bottom:1px solid #e2e8ed;">
                MarbelsunWifi
              </td>
            </tr>
            <tr>
              <td style="padding:12px 16px; color:#777777;">
                Password
              </td>
              <td style="padding:12px 16px; font-family:monospace; color:#012030;
                         font-weight:bold; font-size:15px;">
                marbella2026
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- ── SEPARADOR ──────────────────────────── -->
      <tr>
<td style="padding:24px 32px 0;">
<hr style="border:none; border-top:1px solid #e8e8e8; margin:0;">
</td>
</tr>
      <!-- ── CONTACTO ────────────────────────────── -->
      <tr>
        <td style="padding:24px 32px 0;">
          <p style="margin:0 0 14px; font-size:11px; font-weight:bold; letter-spacing:0.12em;
                    text-transform:uppercase; color:#012030;">
            💬 &nbsp;Need help?
          </p>
          <p style="margin:0 0 10px; font-size:14px; color:#444444; line-height:1.7;">
            If you have any questions or run into any issues upon arrival, please don't hesitate to call us.
          </p>
          <p style="margin:0; font-size:15px; color:#444444; line-height:1.6;">
  <a href="tel:+34611164242" style="color:#012030; text-decoration:none; font-weight:bold;">
    +34 611 164 242
  </a>
</p>
        </td>
      </tr>

      <!-- ── FIRMA ───────────────────────────────── -->
      <tr>
        <td style="padding:28px 32px 32px;">
          <p style="margin:0; font-size:15px; color:#444444; line-height:1.7;">
            We hope you have a wonderful stay!
          </p>
          <p style="margin:8px 0 0; font-size:16px; color:#012030; font-weight:bold;">
            Andrés
          </p>
          <p style="margin:4px 0 0; font-size:13px; color:#888888;">
            Ático Marbella Centro
          </p>
        </td>
      </tr>

      <!-- ── FOOTER ──────────────────────────────── -->
      <tr>
        <td style="background-color:#f7f9fb; border-top:1px solid #e8e8e8;
                   padding:18px 32px; text-align:center;">
          <p style="margin:0; font-size:11px; color:#aaaaaa; line-height:1.6;">
            Ático Marbella Centro · Calle Jacinto Benavente, 8 · Marbella, Málaga<br>
            This email was sent automatically. Please do not reply directly to this message.
          </p>
        </td>
      </tr>

    </table>

  </div>

</body>
</html>`;

  const html = (idioma === 'en') ? htmlEN : htmlES;

const asunto = (idioma === 'en')
  ? 'Check-in information – Marbella apartment'
  : 'Información de acceso – Apartamento Marbella';

MailApp.sendEmail({
  to: reserva.email.trim(),
  subject: asunto,
  htmlBody: html,
  name: "Ático Marbella Centro"
});
}

/* ── Helpers Drive ── */

function getRootFolder() {
  try {
    return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (err) {
    Logger.log('No se pudo acceder por ID, buscando por nombre: ' + err.message);
    return obtenerOCrearCarpeta(DRIVE_FOLDER_NAME);
  }
}

function obtenerOCrearCarpeta(nombre, parent) {
  const iter = parent ? parent.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return parent ? parent.createFolder(nombre) : DriveApp.createFolder(nombre);
}

function guardarImagen(carpeta, nombre, base64, mimeType) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nombre);
    const file = carpeta.createFile(blob);
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    return file;
  } catch (err) {
    console.error('Error guardando imagen:', err.toString());
    return null;
  }
}

function extensionDeMime(mime) {
  const map = { 'image/jpeg': '.jpg', 'image/jpg': '.jpg', 'image/png': '.png' };
  return map[mime] || '.img';
}

/* ── Helpers datos ── */
function sanitize(val) {
  if (!val) return 'sin-dato';
  return String(val).trim()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_\-\.]/g, '_')
    .substring(0, 40);
}

function clean(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ══════════════════════════════════════════════════
   FUNCIONES DE TEST / DIAGNÓSTICO — ejecutar desde el editor
══════════════════════════════════════════════════ */

/** DIAGNÓSTICO — Ejecutar desde el editor → Ver → Registros de ejecución */
function testGetRegistros() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('=== DIAGNÓSTICO getRegistros ===');
  Logger.log('SPREADSHEET_ID : ' + SPREADSHEET_ID);
  Logger.log('SHEET_NAME     : ' + SHEET_NAME);
  const todasLasPestanas = ss.getSheets().map(s => s.getName());
  Logger.log('Pestañas disponibles: ' + JSON.stringify(todasLasPestanas));
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('❌ Pestaña "' + SHEET_NAME + '" NO encontrada.'); return; }
  Logger.log('✅ Pestaña: "' + sheet.getName() + '" — lastRow=' + sheet.getLastRow() + ' lastCol=' + sheet.getLastColumn());
  if (sheet.getLastRow() >= 2) {
    Logger.log('Primera fila de datos: ' + JSON.stringify(sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]));
  }
  Logger.log('=== FIN DIAGNÓSTICO ===');
}

/** Devuelve info de la sheet como JSON (llamable vía doGet?action=debug, sin ADMIN_TOKEN) */
function debugSheet() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tabs  = ss.getSheets().map(s => s.getName());
    const sheet = ss.getSheetByName(SHEET_NAME);
    return jsonResponse({
      status:              'ok',
      spreadsheetId:       SPREADSHEET_ID,
      sheetNameBuscado:    SHEET_NAME,
      pestanasDisponibles: tabs,
      sheetEncontrada:     !!sheet,
      lastRow:             sheet ? sheet.getLastRow()    : null,
      lastCol:             sheet ? sheet.getLastColumn() : null,
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/** Verifica que el email funciona de forma aislada */
function testEmail() {
  MailApp.sendEmail(EMAIL_NOTIFICACION, '[DEV] Test registro', 'Prueba manual desde Apps Script DEV — v3.\n\nSi recibes esto, MailApp funciona correctamente.');
  Logger.log('Email de test enviado a ' + EMAIL_NOTIFICACION);
}

/** Fuerza la re-autorización de todos los scopes (ejecutar antes de cada nuevo despliegue) */
function autorizar() {
  const folder = DriveApp.createFolder('TEST_AUTORIZAR_DEV');
  folder.setTrashed(true);
  SpreadsheetApp.openById(SPREADSHEET_ID);
  MailApp.getRemainingDailyQuota();
  Logger.log('Autorización completada. Quota MailApp restante: ' + MailApp.getRemainingDailyQuota());
}
