/**
 * ============================================================
 *  REGISTRO DE HUÉSPEDES — Ático Marbella Centro
 *  Google Apps Script (Web App) — V02
 *  v6 — Dirección a nivel de reserva (formulario rediseñado DEV01)
 * ============================================================
 *
 *  NOVEDADES v6 respecto a v5:
 *  - La dirección es ÚNICA por reserva (7 campos: Tipo Vía, Nombre Vía,
 *    Número, CP, Población, Ciudad, País), repetida en cada fila de huésped.
 *  - Sheet pasa a 28 columnas (cols 14–20 = dirección de la reserva).
 *  - El XML SES usa la dirección de la reserva para cada viajero.
 *  - Compatible con el formulario multi-huésped rediseñado (DEV01/index.html).
 *
 *  NOVEDADES v5 respecto a v4:
 *  - Templates HTML del email al huésped completamente rediseñados
 *    (tipografía Adobe Fonts: Rosalind, Inge Variable, Mozaic Geo)
 *  - Nueva sección PARKING en el email (ES y EN)
 *  - Diseño oscuro en header/footer (#1c1c1a), paleta crema/malva
 *
 *  NOVEDADES v4 respecto a v3:
 *  - 3 campos nuevos por huésped: domicilio, codigoPostal, paisResidencia
 *  - Sheet ahora tiene 24 columnas (en vez de 21)
 *  - Función generarXMLSES() → XML formato "Partes de Viajeros" (spec v1.2.0)
 *  - Email de notificación admin incluye el XML como adjunto (.xml)
 *    con nombre: SES_{ReservaID}_{Fecha}.xml
 *
 *  SETUP:
 *  1. Ejecutar autorizar() → aceptar todos los permisos
 *  2. Implementar → Nueva implementación (NO actualizar)
 *     Tipo: Aplicación web · Ejecutar como: Yo · Acceso: Cualquier persona
 *  3. Copiar la URL /exec → actualizar APPS_SCRIPT_URL en el Worker (Cloudflare)
 *
 * ============================================================
 */

/* ── Configuración general ── */
const SPREADSHEET_ID    = '1cvVZ0WUF-lDxa5KjpTWCTMbOAMlGBF5RhCwNj2PqRGg';
const SHEET_NAME        = 'Registros';
const API_TOKEN         = 'AtMb2025!xK9#qR7vL';
const DRIVE_FOLDER_ID   = '1SluqT8ZB-DJuLNlSK2_2JKj5PovbM7tZ';
const DRIVE_FOLDER_NAME = 'REGISTRO HUÉSPEDES_2026';
const EMAIL_NOTIFICACION = 'ce.australis@gmail.com,aticomarbellacentro@gmail.com';

/* ── Configuración SES.HOSPEDAJES ── */
const SES_CODIGO_ESTABLECIMIENTO = 'ESFCTU0000290290000970200000000000000000VFT/MA/475163';
const SES_TIPO_PAGO_DEFAULT      = 'PLATF'; // Plataforma de pago (Airbnb/Booking)

/* ── Cabeceras de la hoja (28 columnas) ──
   v6: la dirección es a nivel de RESERVA (7 campos), repetida en cada
   fila de huésped. Coincide con el formulario rediseñado (DEV01). */
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
  'Tipo Vía',           // col 14 ← dirección reserva
  'Nombre Vía',         // col 15 ← dirección reserva
  'Número',             // col 16 ← dirección reserva
  'Código Postal',      // col 17 ← dirección reserva
  'Población',          // col 18 ← dirección reserva
  'Ciudad',             // col 19 ← dirección reserva
  'País Residencia',    // col 20 ← dirección reserva
  'Fecha Entrada',      // col 21
  'Fecha Salida',       // col 22
  'Teléfono',           // col 23
  'Email',              // col 24
  'Idioma',             // col 25
  'Carpeta Documentos', // col 26
  'Foto Anverso',       // col 27
  'Foto Reverso',       // col 28
];

/* ══════════════════════════════════════════════════
   doGet — endpoints de lectura para el panel admin
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

    return jsonResponse({ status: 'ok', message: 'API activa — v4 SES XML.' });

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
    return jsonResponse({
      status: 'error',
      message: 'Pestaña "' + SHEET_NAME + '" no encontrada',
      debug: { spreadsheetId: SPREADSHEET_ID, pestanasDisponibles: todasLasPestanas },
    });
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    return jsonResponse({ status: 'ok', total: 0, registros: [] });
  }

  const values   = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const registros = values.map((row, i) => filaAObjeto(row, i + 2));
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

/* Mapea una fila de Sheets a un objeto con claves legibles — v6 (28 columnas) */
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
    viaTipo:         String(row[13] || ''),   // ← col 14 dirección reserva
    viaNombre:       String(row[14] || ''),   // ← col 15 dirección reserva
    viaNumero:       String(row[15] || ''),   // ← col 16 dirección reserva
    codigoPostal:    String(row[16] || ''),   // ← col 17 dirección reserva
    poblacion:       String(row[17] || ''),   // ← col 18 dirección reserva
    ciudad:          String(row[18] || ''),   // ← col 19 dirección reserva
    paisResidencia:  String(row[19] || ''),   // ← col 20 dirección reserva
    fechaEntrada:    String(row[20] || ''),
    fechaSalida:     String(row[21] || ''),
    telefono:        String(row[22] || ''),
    email:           String(row[23] || ''),
    idioma:          String(row[24] || ''),
    carpetaDrive:    String(row[25] || ''),
    fotoAnverso:     String(row[26] || ''),
    fotoReverso:     String(row[27] || ''),
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
        new Date(),                           // col  1 Timestamp
        reservaId,                            // col  2 ReservaID
        codigoCaja,                           // col  3 Código Caja
        clean(h.nombre),                      // col  4 Nombre
        clean(h.apellido1),                   // col  5 Primer Apellido
        clean(h.apellido2),                   // col  6 Segundo Apellido
        clean(h.fechaNacimiento),             // col  7 Fecha Nacimiento
        clean(h.sexo),                        // col  8 Sexo
        clean(h.nacionalidad),                // col  9 Nacionalidad
        clean(h.tipoDocumento),               // col 10 Tipo Documento
        clean(h.numeroDocumento),             // col 11 Nº Documento
        clean(h.fechaExpedicion),             // col 12 Fecha Expedición
        clean(h.numeroSoporte),               // col 13 Nº Soporte
        clean(reserva.viaTipo),               // col 14 Tipo Vía ← dirección reserva
        clean(reserva.viaNombre),             // col 15 Nombre Vía ← dirección reserva
        clean(reserva.viaNumero),             // col 16 Número ← dirección reserva
        clean(reserva.codigoPostal),          // col 17 Código Postal ← dirección reserva
        clean(reserva.poblacion),             // col 18 Población ← dirección reserva
        clean(reserva.ciudad),                // col 19 Ciudad ← dirección reserva
        clean(reserva.pais),                  // col 20 País Residencia ← dirección reserva
        clean(reserva.fechaEntrada),          // col 21 Fecha Entrada
        clean(reserva.fechaSalida),           // col 22 Fecha Salida
        clean(reserva.telefono),              // col 23 Teléfono
        clean(reserva.email),                 // col 24 Email
        clean(reserva.idioma || ''),          // col 25 Idioma
        carpetaUrl,                           // col 26 Carpeta Documentos
        linkFrontal,                          // col 27 Foto Anverso
        linkTrasero,                          // col 28 Foto Reverso
      ];

      sheet.appendRow(row);
      filasGuardadas.push({ huesped: i + 1, nombre: clean(h.nombre) + ' ' + clean(h.apellido1) });
      Logger.log('[doPost] Fila guardada: huésped ' + (i + 1) + ' — ' + clean(h.nombre));
    }

    /* ── Email de notificación admin (con XML SES adjunto) ── */
    let emailResult;
    try {
      enviarNotificacion(reservaId, codigoCaja, reserva, huespedes);
      emailResult = { enviado: true };
    } catch (emailErr) {
      Logger.log('ERROR en enviarNotificacion: ' + emailErr.message);
      emailResult = { enviado: false, error: emailErr.message };
    }

    /* ── Email de bienvenida al huésped ── */
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

/* ══════════════════════════════════════════════════
   GENERACIÓN XML SES.HOSPEDAJES — Partes de Viajeros
   Spec v1.2.0 — Ministerio del Interior
══════════════════════════════════════════════════ */

/**
 * Genera el fichero XML "Parte de Viajeros" según el formato
 * exigido por SES.HOSPEDAJES (Real Decreto 933/2021).
 * @returns {string} XML completo como string UTF-8
 */
function generarXMLSES(reservaId, reserva, huespedes) {
  const hoy            = formatFecha(new Date());
  const numPersonas    = huespedes.length;
  const fechaEntradaXml = fechaConHora(clean(reserva.fechaEntrada), '14:00:00');
  const fechaSalidaXml  = fechaConHora(clean(reserva.fechaSalida),  '11:00:00');

  /* ── Dirección a nivel de RESERVA (compartida por todos los viajeros) ── */
  const dirCalle     = [clean(reserva.viaTipo), clean(reserva.viaNombre), clean(reserva.viaNumero)]
                         .filter(function(x) { return x; }).join(' ');
  const dirCP        = clean(reserva.codigoPostal);
  const dirPaisISO   = paisAIso3(clean(reserva.pais));
  const dirMunicipio = clean(reserva.poblacion) || clean(reserva.ciudad);
  const municipioXml = dirMunicipio
    ? '          <nombreMunicipio>' + escXml(dirMunicipio) + '</nombreMunicipio>\n'
    : '';

  let personasXml = '';
  huespedes.forEach(function(h) {
    const tipoDocSES = mapearTipoDoc(clean(h.tipoDocumento));
    const sexoSES    = mapearSexo(clean(h.sexo));
    const nacISO     = paisAIso3(clean(h.nacionalidad));

    // apellido2 solo si tipoDocumento = NIF
    const apellido2Xml = (tipoDocSES === 'NIF' && clean(h.apellido2))
      ? '      <apellido2>' + escXml(clean(h.apellido2)) + '</apellido2>\n'
      : '';

    // tipoDocumento y numeroDocumento: incluir siempre (el formulario los exige)
    const docXml = tipoDocSES
      ? '      <tipoDocumento>' + tipoDocSES + '</tipoDocumento>\n' +
        '      <numeroDocumento>' + escXml(clean(h.numeroDocumento)) + '</numeroDocumento>\n'
      : '';

    // soporteDocumento: solo si NIF o NIE
    const soporteXml = ((tipoDocSES === 'NIF' || tipoDocSES === 'NIE') && clean(h.numeroSoporte))
      ? '      <soporteDocumento>' + escXml(clean(h.numeroSoporte)) + '</soporteDocumento>\n'
      : '';

    personasXml +=
      '    <persona>\n' +
      '      <rol>VI</rol>\n' +
      '      <nombre>' + escXml(clean(h.nombre)) + '</nombre>\n' +
      '      <apellido1>' + escXml(clean(h.apellido1)) + '</apellido1>\n' +
      apellido2Xml +
      docXml +
      soporteXml +
      '      <fechaNacimiento>' + escXml(clean(h.fechaNacimiento)) + '</fechaNacimiento>\n' +
      (nacISO ? '      <nacionalidad>' + nacISO + '</nacionalidad>\n' : '') +
      '      <sexo>' + sexoSES + '</sexo>\n' +
      '      <direccion>\n' +
      '          <direccion>' + escXml(dirCalle) + '</direccion>\n' +
      '          <codigoPostal>' + escXml(dirCP) + '</codigoPostal>\n' +
      '          <pais>' + dirPaisISO + '</pais>\n' +
      municipioXml +
      '      </direccion>\n' +
      '      <correo>' + escXml(clean(reserva.email)) + '</correo>\n' +
      '      <telefono>' + escXml(clean(reserva.telefono)) + '</telefono>\n' +
      '    </persona>\n';
  });

  return '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<peticion>\n' +
    '  <solicitud>\n' +
    '    <codigoEstablecimiento>' + SES_CODIGO_ESTABLECIMIENTO + '</codigoEstablecimiento>\n' +
    '    <comunicacion>\n' +
    '      <contrato>\n' +
    '        <referencia>' + escXml(reservaId) + '</referencia>\n' +
    '        <fechaContrato>' + hoy + '</fechaContrato>\n' +
    '        <fechaEntrada>' + fechaEntradaXml + '</fechaEntrada>\n' +
    '        <fechaSalida>' + fechaSalidaXml + '</fechaSalida>\n' +
    '        <numPersonas>' + numPersonas + '</numPersonas>\n' +
    '        <pago>\n' +
    '          <tipoPago>' + SES_TIPO_PAGO_DEFAULT + '</tipoPago>\n' +
    '        </pago>\n' +
    '      </contrato>\n' +
    personasXml +
    '    </comunicacion>\n' +
    '  </solicitud>\n' +
    '</peticion>';
}

/* ── Funciones auxiliares para el XML ── */

/** Escapa caracteres especiales XML */
function escXml(s) {
  if (!s) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Mapea el sexo del formulario (M/F/X) al código SES (H/M/O).
 * El formulario usa M=Masculino, F=Femenino — SES usa H=Hombre, M=Mujer.
 */
function mapearSexo(sexoForm) {
  if (sexoForm === 'M') return 'H'; // Masculino → Hombre
  if (sexoForm === 'F') return 'M'; // Femenino  → Mujer
  return 'O';                       // No especificado → Otro
}

/** Mapea el tipo de documento del formulario al código SES */
function mapearTipoDoc(tipoForm) {
  if (tipoForm === 'DNI')               return 'NIF';
  if (tipoForm === 'Pasaporte')         return 'PAS';
  if (tipoForm === 'NIE')               return 'NIE';
  return 'OTRO';
}

/**
 * Convierte el nombre de un país (en español o inglés) al código ISO 3166-1 Alfa-3.
 * Cubre los ~40 países más frecuentes en turismo en España.
 * Fallback: 'XXX' (aceptado por el Ministerio como "desconocido").
 */
function paisAIso3(texto) {
  var mapa = {
    // Español
    'españa': 'ESP', 'alemania': 'DEU', 'francia': 'FRA', 'reino unido': 'GBR',
    'italia': 'ITA', 'portugal': 'PRT', 'países bajos': 'NLD', 'holanda': 'NLD',
    'bélgica': 'BEL', 'suiza': 'CHE', 'austria': 'AUT', 'estados unidos': 'USA',
    'ee.uu.': 'USA', 'usa': 'USA', 'canadá': 'CAN', 'canada': 'CAN',
    'argentina': 'ARG', 'méxico': 'MEX', 'mexico': 'MEX', 'brasil': 'BRA',
    'marruecos': 'MAR', 'polonia': 'POL', 'suecia': 'SWE', 'dinamarca': 'DNK',
    'noruega': 'NOR', 'finlandia': 'FIN', 'irlanda': 'IRL', 'australia': 'AUS',
    'japón': 'JPN', 'japon': 'JPN', 'china': 'CHN', 'rusia': 'RUS',
    'ucrania': 'UKR', 'rumanía': 'ROU', 'rumania': 'ROU', 'rep. checa': 'CZE',
    'república checa': 'CZE', 'hungría': 'HUN', 'hungria': 'HUN',
    'grecia': 'GRC', 'turquía': 'TUR', 'turquia': 'TUR', 'india': 'IND',
    'colombia': 'COL', 'chile': 'CHL', 'perú': 'PER', 'peru': 'PER',
    'ecuador': 'ECU', 'venezuela': 'VEN', 'cuba': 'CUB',
    // Inglés
    'spain': 'ESP', 'germany': 'DEU', 'france': 'FRA', 'united kingdom': 'GBR',
    'uk': 'GBR', 'gb': 'GBR', 'great britain': 'GBR', 'england': 'GBR',
    'italy': 'ITA', 'netherlands': 'NLD', 'switzerland': 'CHE',
    'sweden': 'SWE', 'norway': 'NOR', 'denmark': 'DNK', 'finland': 'FIN',
    'poland': 'POL', 'ireland': 'IRL', 'greece': 'GRC', 'turkey': 'TUR',
    'romania': 'ROU', 'hungary': 'HUN', 'czech republic': 'CZE',
    'united states': 'USA', 'brazil': 'BRA', 'morocco': 'MAR',
    'russia': 'RUS', 'ukraine': 'UKR', 'japan': 'JPN', 'belgium': 'BEL',
    'austria': 'AUT',
    // Códigos ISO directos (por si el usuario escribe el código)
    'esp': 'ESP', 'deu': 'DEU', 'fra': 'FRA', 'gbr': 'GBR', 'ita': 'ITA',
    'prt': 'PRT', 'nld': 'NLD', 'bel': 'BEL', 'che': 'CHE', 'aut': 'AUT',
    'usa': 'USA', 'can': 'CAN', 'arg': 'ARG', 'mex': 'MEX', 'bra': 'BRA',
    'mar': 'MAR', 'pol': 'POL', 'swe': 'SWE', 'dnk': 'DNK', 'nor': 'NOR',
    'fin': 'FIN', 'irl': 'IRL', 'aus': 'AUS', 'jpn': 'JPN', 'chn': 'CHN',
  };

  var clave = String(texto).trim().toLowerCase();
  return mapa[clave] || 'XXX'; // XXX = desconocido (aceptado por el Ministerio)
}

/** Añade la hora a una fecha "YYYY-MM-DD" → "YYYY-MM-DDThh:mm:ss" */
function fechaConHora(fecha, hora) {
  return String(fecha).trim() + 'T' + hora;
}

/** Formatea un objeto Date a "YYYY-MM-DD" */
function formatFecha(d) {
  return Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd');
}

/* ══════════════════════════════════════════════════
   EMAIL DE NOTIFICACIÓN ADMIN (con XML adjunto)
══════════════════════════════════════════════════ */
function enviarNotificacion(reservaId, codigoCaja, reserva, huespedes) {
  const primerH  = huespedes[0];
  const numH     = huespedes.length;
  const sufijo   = numH > 1 ? ' +' + (numH - 1) + ' más' : '';
  const asunto   = '[SES XML adjunto] Reserva ' + reservaId + ': ' +
                   clean(primerH.nombre) + ' ' + clean(primerH.apellido1) + sufijo;

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

  /* Dirección a nivel de reserva (compartida por todos los huéspedes) */
  const dirReserva =
    [clean(reserva.viaTipo), clean(reserva.viaNombre), clean(reserva.viaNumero)]
      .filter(function(x) { return x; }).join(' ') +
    (clean(reserva.codigoPostal) ? ', ' + clean(reserva.codigoPostal) : '') +
    (clean(reserva.poblacion)    ? ' ' + clean(reserva.poblacion)     : '') +
    (clean(reserva.ciudad)       ? ', ' + clean(reserva.ciudad)       : '') +
    (clean(reserva.pais)         ? ' (' + clean(reserva.pais) + ')'   : '');

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
    '  Teléfono : ' + clean(reserva.telefono) + '\n' +
    '  Email    : ' + clean(reserva.email)    + '\n' +
    '  Dirección: ' + dirReserva              + '\n\n' +
    'HUÉSPEDES (datos SES):' + detalleHuespedes + '\n' +
    '────────────────────────────────\n' +
    'Se adjunta el fichero XML listo para subir a SES.HOSPEDAJES.\n' +
    'Portal: https://hospedajes.ses.mir.es\n';

  /* Generar XML SES y adjuntarlo */
  const xmlString     = generarXMLSES(reservaId, reserva, huespedes);
  const nombreFichero = 'SES_' + reservaId + '_' +
                        Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyyMMdd') + '.xml';
  const xmlBlob       = Utilities.newBlob(xmlString, 'application/xml', nombreFichero);

  MailApp.sendEmail({
    to:          EMAIL_NOTIFICACION,
    subject:     asunto,
    body:        cuerpo,
    attachments: [xmlBlob],
  });

  Logger.log('[enviarNotificacion] Email enviado con adjunto: ' + nombreFichero);
}

/* ══════════════════════════════════════════════════
   EMAIL DE BIENVENIDA AL HUÉSPED
══════════════════════════════════════════════════ */
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
  <title>Instrucciones de llegada · Ático Marbella Centro</title>
  <link rel="stylesheet" href="https://use.typekit.net/hzo3vlq.css">
</head>
<body style="margin:0;padding:0;background-color:#f5f3ee;">

  <!-- ─── HEADER ────────────────────────────────────────── -->
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#1c1c1a;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;">
          <tr>
            <td align="center" style="padding:44px 48px 40px;">
              <p style="margin:0 0 8px 0;font-family:'rosalind',Georgia,serif;font-weight:400;font-size:28px;color:#e2dae1;letter-spacing:0.02em;line-height:1.2;">
                Ático Marbella Centro
              </p>
              <p style="margin:0;font-family:'inge-variable',Georgia,serif;font-style:normal;font-weight:400;font-size:52px;color:#f5f3ee;line-height:1.05;letter-spacing:-0.01em;">
                instrucciones de llegada
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>

  <!-- CUERPO ──────────────────────────────────────────── -->
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f5f3ee;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;">


          <!-- ─── SALUDO ─────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:28px;color:#867281;line-height:1.3;">
                Hola, ${nombre}
              </p>
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Gracias por rellenar el formulario de registro.<br>
                A continuación toda la información para el acceso al apartamento.
              </p>
            </td>
          </tr>


          <!-- ─── TU RESERVA ────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 24px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                TU RESERVA
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="vertical-align:top;padding-right:32px;">
                    <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">
                      Check in
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          ${fechaEntrada}
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td style="vertical-align:top;">
                    <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">
                      Check out
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          ${fechaSalida}
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── DIRECCIÓN ──────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="vertical-align:top;">
                    <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">
                      Dirección
                    </p>
                    <p style="margin:0 0 4px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:20px;color:#867281;line-height:1.3;">
                      Calle Jacinto Benavente, 8
                    </p>
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:17px;color:#3b3a3d;line-height:1.5;">
                      Edificio Marbelsun III<br>
                      8ª planta · puerta 3
                    </p>
                  </td>
                  <td style="vertical-align:bottom;width:170px;text-align:right;padding-left:16px;">
                    <a href="https://maps.app.goo.gl/pRVaVEH8yxymNwq99"
                       style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;white-space:nowrap;">
                      Abrir en Google Maps
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── ACCESO ─────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                ACCESO
              </p>
              <p style="margin:0 0 32px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                El acceso al apartamento se realiza de forma autónoma mediante una caja de seguridad con código.
                <span style="color:#867281;font-weight:300;">Sigue estos pasos</span>:
              </p>

              <!-- Paso 1 -->
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:36px;">
                <tr>
                  <td style="vertical-align:top;width:56px;padding-right:16px;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:56px;color:#867281;opacity:0.6;line-height:1;">1</p>
                  </td>
                  <td style="vertical-align:top;padding-top:4px;">
                    <p style="margin:0 0 18px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      Al llegar al portal del edificio, justo enfrente verás una escalera con
                      <span style="color:#867281;font-weight:300;">barandilla metálica</span>.
                      La caja de llaves está colgada en la barandilla, en el lado derecho de la escalera (mirando desde el portal).
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                      <tr>
                        <td style="text-align:right;padding-top:16px;">
                          <a href="https://photos.app.goo.gl/dW7efj3aK2EZeHm1A"
                             style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;">
                            Foto ubicación
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Paso 2 -->
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:36px;">
                <tr>
                  <td style="vertical-align:top;width:56px;padding-right:16px;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:56px;color:#867281;opacity:0.6;line-height:1;">2</p>
                  </td>
                  <td style="vertical-align:top;padding-top:4px;">
                    <p style="margin:0 0 18px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      Introduce el <span style="color:#867281;font-weight:300;">código</span> de acceso en la caja:
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                      <tr>
                        <td style="text-align:right;padding-top:16px;">
                          <span style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                            ${codigoCaja}
                          </span>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Paso 3 -->
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="vertical-align:top;width:56px;padding-right:16px;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:56px;color:#867281;opacity:0.6;line-height:1;">3</p>
                  </td>
                  <td style="vertical-align:top;padding-top:4px;">
                    <p style="margin:0 0 8px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      Dentro encontrarás un llavero con:
                    </p>
                    <p style="margin:0 0 4px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      <span style="color:#867281;font-weight:300;">Chip azul</span> → abre el portal del edificio
                    </p>
                    <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      <span style="color:#867281;font-weight:300;">Llave grande</span> → abre la puerta del apartamento
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                      <tr>
                        <td style="text-align:right;padding-top:16px;">
                          <a href="https://photos.app.goo.gl/Vcpi36tqdTF5AjAY8"
                             style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;">
                            Instrucciones extra
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── ASCENSOR ───────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                ASCENSOR
              </p>
              <p style="margin:0 0 12px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Al entrar al portal, los ascensores para subir al apartamento se encuentran inmediatamente a mano izquierda.<br>
                Ten en cuenta que en el edificio hay otros ascensores/montacargas; utiliza únicamente estos para acceder a la vivienda.
              </p>
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Una vez en el ascensor, sube hasta la <span style="color:#867281;font-weight:300;">planta 8</span>. El apartamento es la <span style="color:#867281;font-weight:300;">puerta nº 3</span>.
              </p>
            </td>
          </tr>


          <!-- ─── PARKING ────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                PARKING
              </p>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Acceso en coche:<br>
                La <span style="color:#867281;font-weight:300;">entrada al aparcamiento</span> se encuentra en:<br>
                Jacinto Benavente 8, junto a Ferretería Villalón.
              </p>

              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:14px;">
                <tr>
                  <td style="text-align:right;">
                    <a href="https://www.google.com/maps/search/?api=1&query=Jacinto%20Benavente%208%20Marbella"
                       style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;white-space:nowrap;">
                      Abrir en Google Maps
                    </a>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                La plaza de aparcamiento está en el <span style="color:#867281;font-weight:300;">sótano -3 (B3)</span>, <span style="color:#867281;font-weight:300;">plaza número 394</span>.<br>
                Una vez en el nivel -3, es la <span style="color:#867281;font-weight:300;">tercera plaza a su izquierda</span>.
              </p>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                <span style="color:#867281;font-weight:300;">Acceso a pie</span>:<br>
                La entrada se encuentra en:<br>
                <span style="color:#867281;font-weight:300;">Calle Jacinto Benavente 12</span>
              </p>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Por favor, utilice la puerta que aparece en las fotos adjuntas. Puede abrirla con la <span style="color:#867281;font-weight:300;">llave pequeña</span> que va unida al mando a distancia.<br>
                La llave se encuentra junto al resto de llaves del apartamento.
              </p>

              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                <span style="color:#867281;font-weight:300;">IMPORTANTE</span>:<br>
                El ascensor solo va del nivel -3 al nivel -1.<br>
                No hay acceso en ascensor entre la calle y el nivel -1, por lo que <span style="color:#867281;font-weight:300;">deberá usar las escaleras para ese tramo</span>.
              </p>

              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="text-align:right;">
                    <a href="https://photos.app.goo.gl/NxWiwZZJmfcNphbQA"
                       style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;">
                      Instrucciones extra
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── WIFI ───────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                WIFI
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="padding-right:14px;vertical-align:middle;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">Red</p>
                  </td>
                  <td style="padding-right:32px;vertical-align:middle;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          MarbelsunWifi
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td style="padding-right:14px;vertical-align:middle;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">Contraseña</p>
                  </td>
                  <td style="vertical-align:middle;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          marbella2026
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── MÁS ────────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 56px;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                MÁS
              </p>
              <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:11px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;">
                ¿Necesitas ayuda?
              </p>
              <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:14px;color:#3b3a3d;line-height:1.8;">
                Si tienes cualquier duda o surge algún problema durante la llegada, no dudes en llamarnos.
              </p>
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:14px;color:#3b3a3d;line-height:1.8;">
                +34 611 164 242
              </p>
            </td>
          </tr>


        </table>
      </td>
    </tr>
  </table>

  <!-- ─── FOOTER ─────────────────────────────────────────── -->
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#1c1c1a;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;">
          <tr>
            <td align="center" style="padding:24px 48px;">
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:11px;color:#f5f3ee;line-height:1.8;letter-spacing:0.04em;">
                Ático Marbella Centro · Calle Jacinto Benavente, 8 · Marbella, Málaga
              </p>
              <p style="margin:4px 0 0 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:11px;color:#f5f3ee;line-height:1.8;letter-spacing:0.04em;">
                Este correo ha sido enviado automáticamente.<br>
                Por favor, no respondas directamente a este mensaje.
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>

</body>
</html>`;

  const htmlEN = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Arrival instructions · Ático Marbella Centro</title>
  <link rel="stylesheet" href="https://use.typekit.net/hzo3vlq.css">
</head>
<body style="margin:0;padding:0;background-color:#f5f3ee;">

  <!-- ─── HEADER ────────────────────────────────────────── -->
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#1c1c1a;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;">
          <tr>
            <td align="center" style="padding:44px 48px 40px;">
              <p style="margin:0 0 8px 0;font-family:'rosalind',Georgia,serif;font-weight:400;font-size:28px;color:#e2dae1;letter-spacing:0.02em;line-height:1.2;">
                Ático Marbella Centro
              </p>
              <p style="margin:0;font-family:'inge-variable',Georgia,serif;font-style:normal;font-weight:400;font-size:52px;color:#f5f3ee;line-height:1.05;letter-spacing:-0.01em;">
                arrival instructions
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>

  <!-- BODY ──────────────────────────────────────────────── -->
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f5f3ee;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;">


          <!-- ─── GREETING ──────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:28px;color:#867281;line-height:1.3;">
                Hello, ${nombre}
              </p>
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Thank you for completing the registration form.<br>
                Below you will find all the information you need to access the apartment.
              </p>
            </td>
          </tr>


          <!-- ─── YOUR BOOKING ──────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 24px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                YOUR BOOKING
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="vertical-align:top;padding-right:32px;">
                    <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">
                      Check in
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          ${fechaEntrada}
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td style="vertical-align:top;">
                    <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">
                      Check out
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          ${fechaSalida}
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── ADDRESS ───────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="vertical-align:top;">
                    <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">
                      Address
                    </p>
                    <p style="margin:0 0 4px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:20px;color:#867281;line-height:1.3;">
                      Calle Jacinto Benavente, 8
                    </p>
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:17px;color:#3b3a3d;line-height:1.5;">
                      Edificio Marbelsun III<br>
                      8th floor · apartment 3
                    </p>
                  </td>
                  <td style="vertical-align:bottom;width:170px;text-align:right;padding-left:16px;">
                    <a href="https://maps.app.goo.gl/pRVaVEH8yxymNwq99"
                       style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;white-space:nowrap;">
                      Open in Google Maps
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── ACCESS ────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                ACCESS
              </p>
              <p style="margin:0 0 32px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Access to the apartment is self-managed using a coded lock box.
                <span style="color:#867281;font-weight:300;">Follow these steps</span>:
              </p>

              <!-- Step 1 -->
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:36px;">
                <tr>
                  <td style="vertical-align:top;width:56px;padding-right:16px;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:56px;color:#867281;opacity:0.6;line-height:1;">1</p>
                  </td>
                  <td style="vertical-align:top;padding-top:4px;">
                    <p style="margin:0 0 18px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      When you arrive at the building entrance, right in front of you there is a staircase with a
                      <span style="color:#867281;font-weight:300;">metal railing</span>.
                      The key box hangs from the railing, on the right side of the staircase (as seen from the entrance).
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                      <tr>
                        <td style="text-align:right;padding-top:16px;">
                          <a href="https://photos.app.goo.gl/dW7efj3aK2EZeHm1A"
                             style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;">
                            Location photo
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Step 2 -->
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:36px;">
                <tr>
                  <td style="vertical-align:top;width:56px;padding-right:16px;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:56px;color:#867281;opacity:0.6;line-height:1;">2</p>
                  </td>
                  <td style="vertical-align:top;padding-top:4px;">
                    <p style="margin:0 0 18px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      Enter the <span style="color:#867281;font-weight:300;">code</span> into the lock box:
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                      <tr>
                        <td style="text-align:right;padding-top:16px;">
                          <span style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                            ${codigoCaja}
                          </span>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Step 3 -->
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="vertical-align:top;width:56px;padding-right:16px;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:56px;color:#867281;opacity:0.6;line-height:1;">3</p>
                  </td>
                  <td style="vertical-align:top;padding-top:4px;">
                    <p style="margin:0 0 8px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      Inside you will find a keyring with:
                    </p>
                    <p style="margin:0 0 4px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      <span style="color:#867281;font-weight:300;">Blue chip</span> → opens the building entrance
                    </p>
                    <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                      <span style="color:#867281;font-weight:300;">Large key</span> → opens the apartment door
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                      <tr>
                        <td style="text-align:right;padding-top:16px;">
                          <a href="https://photos.app.goo.gl/Vcpi36tqdTF5AjAY8"
                             style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;">
                            Extra instructions
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── ELEVATOR ───────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                ELEVATOR
              </p>
              <p style="margin:0 0 12px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                As you enter the building, the elevators to the apartment are immediately on your left.<br>
                Please note there are other lifts/service elevators in the building — use only these to access the apartment.
              </p>
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Once in the elevator, go up to <span style="color:#867281;font-weight:300;">floor 8</span>. The apartment is <span style="color:#867281;font-weight:300;">door number 3</span>.
              </p>
            </td>
          </tr>


          <!-- ─── PARKING ────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                PARKING
              </p>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                By car:<br>
                The <span style="color:#867281;font-weight:300;">parking entrance</span> is located at:<br>
                Jacinto Benavente 8, next to Ferretería Villalón.
              </p>

              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:14px;">
                <tr>
                  <td style="text-align:right;">
                    <a href="https://www.google.com/maps/search/?api=1&query=Jacinto%20Benavente%208%20Marbella"
                       style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;white-space:nowrap;">
                      Open in Google Maps
                    </a>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                The parking space is on <span style="color:#867281;font-weight:300;">basement level -3 (B3)</span>, <span style="color:#867281;font-weight:300;">space number 394</span>.<br>
                Once on level -3, it is the <span style="color:#867281;font-weight:300;">third space on your left</span>.
              </p>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                <span style="color:#867281;font-weight:300;">On foot</span>:<br>
                The entrance is located at:<br>
                <span style="color:#867281;font-weight:300;">Calle Jacinto Benavente 12</span>
              </p>

              <p style="margin:0 0 14px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                Please use the door shown in the attached photos. You can open it with the <span style="color:#867281;font-weight:300;">small key</span> attached to the remote control.<br>
                The key is kept with the rest of the apartment keys.
              </p>

              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:15px;color:#3b3a3d;line-height:1.8;">
                <span style="color:#867281;font-weight:300;">IMPORTANT</span>:<br>
                The elevator only goes from level -3 to level -1.<br>
                There is no elevator access between street level and level -1, so <span style="color:#867281;font-weight:300;">you will need to use the stairs for that section</span>.
              </p>

              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="text-align:right;">
                    <a href="https://photos.app.goo.gl/NxWiwZZJmfcNphbQA"
                       style="display:inline-block;background-color:#e2dae1;border-radius:4px;padding:11px 18px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:10px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;text-decoration:none;">
                      Extra instructions
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── WIFI ───────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 0;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                WIFI
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="padding-right:14px;vertical-align:middle;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">Network</p>
                  </td>
                  <td style="padding-right:32px;vertical-align:middle;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          MarbelsunWifi
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td style="padding-right:14px;vertical-align:middle;">
                    <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:100;font-size:11px;color:#867281;letter-spacing:0.14em;text-transform:uppercase;">Password</p>
                  </td>
                  <td style="vertical-align:middle;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="background-color:#e2dae1;border-radius:4px;padding:13px 22px;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:300;font-size:15px;color:#867281;">
                          marbella2026
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>


          <!-- ─── MORE ───────────────────────────────────────────── -->
          <tr>
            <td style="padding:44px 48px 56px;">
              <p style="margin:0 0 20px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:400;font-size:32px;color:#444441;line-height:1;letter-spacing:0.02em;opacity:0.6;">
                MORE
              </p>
              <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:11px;color:#867281;letter-spacing:0.12em;text-transform:uppercase;">
                Need help?
              </p>
              <p style="margin:0 0 10px 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:14px;color:#3b3a3d;line-height:1.8;">
                If you have any questions or run into any issues upon arrival, please don't hesitate to call us.
              </p>
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:14px;color:#3b3a3d;line-height:1.8;">
                +34 611 164 242
              </p>
            </td>
          </tr>


        </table>
      </td>
    </tr>
  </table>

  <!-- ─── FOOTER ─────────────────────────────────────────── -->
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#1c1c1a;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;">
          <tr>
            <td align="center" style="padding:24px 48px;">
              <p style="margin:0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:11px;color:#f5f3ee;line-height:1.8;letter-spacing:0.04em;">
                Ático Marbella Centro · Calle Jacinto Benavente, 8 · Marbella, Málaga
              </p>
              <p style="margin:4px 0 0 0;font-family:'mozaic-geo-variable',Arial,sans-serif;font-weight:200;font-size:11px;color:#f5f3ee;line-height:1.8;letter-spacing:0.04em;">
                This email was sent automatically.<br>
                Please do not reply directly to this message.
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>

</body>
</html>`;

  const html  = (idioma === 'en') ? htmlEN : htmlES;
  const asunto = (idioma === 'en')
    ? 'Check-in information – Marbella apartment'
    : 'Información de acceso – Apartamento Marbella';

  MailApp.sendEmail({
    to:       reserva.email.trim(),
    subject:  asunto,
    htmlBody: html,
    name:     'Ático Marbella Centro',
  });
}

/* ══════════════════════════════════════════════════
   HELPERS DRIVE
══════════════════════════════════════════════════ */
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
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
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
   FUNCIONES DE TEST / DIAGNÓSTICO
══════════════════════════════════════════════════ */

/** Devuelve info de la sheet como JSON (llamable vía doGet?action=debug) */
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
  MailApp.sendEmail(EMAIL_NOTIFICACION, '[TEST V02] Test registro', 'Prueba manual desde Apps Script V02.\nSi recibes esto, MailApp funciona correctamente.');
  Logger.log('Email de test enviado a ' + EMAIL_NOTIFICACION);
}

/** Prueba la generación del XML con datos ficticios y lo adjunta en un email de test */
function testXML() {
  const reservaFicti = {
    fechaEntrada: '2026-06-01',
    fechaSalida:  '2026-06-05',
    telefono:     '+34611164242',
    email:        EMAIL_NOTIFICACION.split(',')[0],
    idioma:       'es',
    viaTipo:      'Calle',
    viaNombre:    'Mayor',
    viaNumero:    '10',
    codigoPostal: '28001',
    poblacion:    'Madrid',
    ciudad:       'Madrid',
    pais:         'España',
  };
  const huespedesFicti = [{
    nombre:          'Juan',
    apellido1:       'García',
    apellido2:       'López',
    fechaNacimiento: '1985-03-20',
    sexo:            'M',
    nacionalidad:    'España',
    tipoDocumento:   'DNI',
    numeroDocumento: '12345678A',
    fechaExpedicion: '2020-01-15',
    numeroSoporte:   'AAA123456',
  }];

  const xml = generarXMLSES('R9999', reservaFicti, huespedesFicti);
  const blob = Utilities.newBlob(xml, 'application/xml', 'SES_TEST_R9999.xml');

  MailApp.sendEmail({
    to:          EMAIL_NOTIFICACION,
    subject:     '[TEST V02] XML SES generado — revisar adjunto',
    body:        'XML de prueba generado por generarXMLSES(). Revisa que el formato es correcto.\n\n' + xml,
    attachments: [blob],
  });
  Logger.log('Email de test XML enviado. Contenido:\n' + xml);
}

/** Fuerza la re-autorización de todos los scopes (ejecutar antes de cada nuevo despliegue) */
function autorizar() {
  const folder = DriveApp.createFolder('TEST_AUTORIZAR_V02');
  folder.setTrashed(true);
  SpreadsheetApp.openById(SPREADSHEET_ID);
  MailApp.getRemainingDailyQuota();
  Logger.log('Autorización completada. Quota MailApp restante: ' + MailApp.getRemainingDailyQuota());
}
