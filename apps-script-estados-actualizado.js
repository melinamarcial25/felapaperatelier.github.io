/*
 * PORTAL FELA — RESPUESTAS EN COLUMNAS
 *
 * 1. Conserva el resto de tu Code.gs.
 * 2. Agrega este bloque completo al proyecto.
 * 3. Elimina o reemplaza las funciones doPost, load_, save_ y savePhoto_
 *    anteriores para evitar nombres duplicados.
 * 4. Guarda y ejecuta migrarRespuestasAColumnas() UNA SOLA VEZ.
 * 5. Revisa la hoja de respaldo creada automáticamente.
 * 6. Crea una nueva versión de la implementación web.
 */

const RESPONSE_HEADERS_V2 = [
  // Identificación y estado.
  'cliente_id',
  'nombre_cliente',
  'estado',

  // 01 · Los festejados.
  'nombres_pareja',
  'orden_pareja',
  'incluir_padres',
  'nombre_padre_1',
  'nombre_madre_1',
  'incluir_segundos_padres',
  'nombre_padre_2',
  'nombre_madre_2',
  'incluir_padrinos_corte',
  'padrinos',
  'corte_honor',

  // 02 · Ceremonia y recepción.
  'fecha_ceremonia',
  'hora_ceremonia',
  'nombre_lugar',
  'direccion_completa',
  'enlace_mapa',

  // 03 · Confirmación de asistencia.
  'tipo_confirmacion',
  'telefono_confirmacion',
  'fecha_limite_rsvp',
  'nota_invitados',
  'comentarios_rsvp',
  'politica_ninos_no',
  'politica_ninos_guarderia',

  // 04 · Vestimenta.
  'dress_code',
  'nota_estilo',

  // 05 · Horario del día.
  'hora_ceremonia_evento',
  'hora_coctel',
  'hora_cena',
  'hora_fiesta',
  'momentos_extra',

  // 06 · Alojamiento y transporte.
  'sugerencias_alojamiento',
  'detalles_transporte',
  'indicaciones_especiales',

  // 07 · Regalos.
  'sugerencia_regalo',
  'mesa_regalo',
  'tienda_codigo',
  'cuenta_banco',
  'instrucciones_regalo',

  // 08 · Inspiración y estilo.
  'inspiracion',
  'paleta_colores',
  'cancion',
  'link_youtube',

  // 09 · Fotografías.
  'enlaces_fotografias',

  // 10 · Cuenta regresiva.
  'incluir_cuenta_regresiva',

  // 11 · Detalles finales.
  'imprescindibles',
  'notas_finales',

  // Control y respaldo técnico.
  'ultima_actualizacion',
  'datos_json',
  'fotos_json'
];

const RESPONSE_FORM_FIELDS_V2 = RESPONSE_HEADERS_V2.filter(header =>
  [
    'cliente_id',
    'nombre_cliente',
    'estado',
    'enlaces_fotografias',
    'ultima_actualizacion',
    'datos_json',
    'fotos_json'
  ].indexOf(header) < 0
);

function doPost(e) {
  try {
    const request = JSON.parse(
      (e.postData && e.postData.contents) || '{}'
    );
    let result;

    switch (request.action) {
      case 'login':
        result = login_(request);
        break;
      case 'load':
        result = load_(request);
        break;
      case 'save':
        result = save_(request);
        break;
      case 'deletePhoto':
        result = deletePhoto_(request);
        break;
      case 'getPhotoThumbnail':
        result = getPhotoThumbnail_(request);
        break;
      default:
        throw new Error('Acción no permitida.');
    }

    return json_({ ok: true, ...result });
  } catch (error) {
    console.error(error);
    return json_({
      ok: false,
      message:
        error.message ||
        'No fue posible completar la solicitud.'
    });
  }
}

/*
 * Ejecuta esta función una sola vez desde el editor de Apps Script.
 * Crea una copia completa de la hoja anterior antes de modificarla.
 */
function migrarRespuestasAColumnas() {
  const book = book_();
  const source = sheet_(RESPONSES_SHEET);
  const values = source.getDataRange().getValues();
  const oldHeaders = values.length ? values[0].map(String) : [];

  const timeZone = Session.getScriptTimeZone() || 'America/Hermosillo';
  const stamp = Utilities.formatDate(
    new Date(),
    timeZone,
    'yyyyMMdd-HHmmss'
  );
  const backupName = uniqueSheetName_(
    book,
    'Respuestas respaldo ' + stamp
  );
  source.copyTo(book).setName(backupName);

  const migratedRows = values.slice(1).map(row => {
    const oldRecord = {};
    oldHeaders.forEach((header, index) => {
      oldRecord[header] = row[index];
    });

    const data = parseJson_(oldRecord.datos_json, {});
    const photos = parseJson_(oldRecord.fotos_json, []);
    const status =
      String(data.estado_proyecto || '').trim() ||
      String(oldRecord.estado || '').trim() ||
      'Borrador';

    return buildResponseRowV2_({
      clientId: oldRecord.cliente_id,
      clientName: oldRecord.nombre_cliente,
      status: status,
      data: data,
      photos: photos,
      updatedAt: oldRecord.ultima_actualizacion
    });
  });

  source.clearContents();
  source.getRange(1, 1, 1, RESPONSE_HEADERS_V2.length)
    .setValues([RESPONSE_HEADERS_V2]);

  if (migratedRows.length) {
    source.getRange(
      2,
      1,
      migratedRows.length,
      RESPONSE_HEADERS_V2.length
    ).setValues(migratedRows);
  }

  source.setFrozenRows(1);
  source.getRange(1, 1, 1, RESPONSE_HEADERS_V2.length)
    .setFontWeight('bold')
    .setBackground('#5a4032')
    .setFontColor('#ffffff');
  source.autoResizeColumns(1, RESPONSE_HEADERS_V2.length);

  console.log(
    'Migración terminada. Respaldo creado: ' + backupName
  );
}

function load_(request) {
  const session = session_(request.token);
  const response = responseFor_(session.clientId);

  return {
    client: {
      id: session.clientId,
      name: session.name
    },
    draft: response ? parseJson_(response.datos_json, {}) : {},
    photos: response ? parseJson_(response.fotos_json, []) : [],
    status: response ? String(response.estado || 'Borrador') : 'Borrador',
    updatedAt: response
      ? String(response.ultima_actualizacion || '')
      : ''
  };
}

function save_(request) {
  const session = session_(request.token);
  const data = request.data;

  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('Datos del formulario inválidos.');
  }

  const uploads = Array.isArray(request.uploads)
    ? request.uploads
    : [];

  if (uploads.length > CONFIG.MAX_UPLOADS_PER_REQUEST) {
    throw new Error(
      'Puedes subir un máximo de ' +
      CONFIG.MAX_UPLOADS_PER_REQUEST +
      ' fotos por vez.'
    );
  }

  const allowedStatuses = [
    'Borrador',
    'Detalles sin completar',
    'Listo para diseñar'
  ];
  const requestedStatus =
    allowedStatuses.indexOf(data.estado_proyecto) >= 0
      ? data.estado_proyecto
      : 'Borrador';

  assertResponseHeadersV2_();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const existing = responseFor_(session.clientId);

    if (
      existing &&
      String(existing.estado || '') === 'Listo para diseñar'
    ) {
      throw new Error(
        'Este formulario ya fue marcado como listo para diseñar y no admite cambios.'
      );
    }

    const oldPhotos = existing
      ? parseJson_(existing.fotos_json, [])
      : [];

    clientFolder_(session);

    const newPhotos = uploads.map(upload =>
      savePhoto_(upload, session)
    );
    const allPhotos = oldPhotos.concat(newPhotos);
    const now = new Date().toISOString();

    const row = buildResponseRowV2_({
      clientId: session.clientId,
      clientName: session.name,
      status: requestedStatus,
      data: data,
      photos: allPhotos,
      updatedAt: now
    });

    const target = sheet_(RESPONSES_SHEET);
    if (existing) {
      target
        .getRange(existing._row, 1, 1, RESPONSE_HEADERS_V2.length)
        .setValues([row]);
    } else {
      target.appendRow(row);
    }

    return {
      updatedAt: now,
      status: requestedStatus,
      photos: allPhotos
    };
  } finally {
    lock.releaseLock();
  }
}

function deletePhoto_(request) {
  const session = session_(request.token);
  const photoId = String(request.photoId || '').trim();

  if (!photoId) {
    throw new Error('No se recibió la fotografía que deseas eliminar.');
  }

  assertResponseHeadersV2_();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const existing = responseFor_(session.clientId);
    if (!existing) {
      throw new Error('No encontramos el expediente del cliente.');
    }
    if (String(existing.estado || '') === 'Listo para diseñar') {
      throw new Error(
        'Este formulario está listo para diseñar y ya no admite cambios.'
      );
    }

    const photos = parseJson_(existing.fotos_json, []);
    const photo = photos.find(item =>
      String(item.id || '') === photoId
    );

    if (!photo) {
      throw new Error(
        'La fotografía no pertenece a este expediente o ya fue eliminada.'
      );
    }

    // La imagen se mueve a la papelera para que pueda recuperarse desde Drive.
    DriveApp.getFileById(photoId).setTrashed(true);

    const remainingPhotos = photos.filter(item =>
      String(item.id || '') !== photoId
    );
    const data = parseJson_(existing.datos_json, {});
    const now = new Date().toISOString();
    const row = buildResponseRowV2_({
      clientId: session.clientId,
      clientName: session.name,
      status: String(existing.estado || 'Borrador'),
      data: data,
      photos: remainingPhotos,
      updatedAt: now
    });

    sheet_(RESPONSES_SHEET)
      .getRange(existing._row, 1, 1, RESPONSE_HEADERS_V2.length)
      .setValues([row]);

    removeFileRecord_(photo);

    return {
      updatedAt: now,
      photos: remainingPhotos
    };
  } finally {
    lock.releaseLock();
  }
}

function getPhotoThumbnail_(request) {
  const session = session_(request.token);
  const photoId = String(request.photoId || '').trim();
  const existing = responseFor_(session.clientId);

  if (!existing || !photoId) {
    throw new Error('No encontramos la fotografía solicitada.');
  }

  const photos = parseJson_(existing.fotos_json, []);
  const belongsToClient = photos.some(photo =>
    String(photo.id || '') === photoId
  );

  if (!belongsToClient) {
    throw new Error('La fotografía no pertenece a este expediente.');
  }

  const thumbnail = DriveApp.getFileById(photoId).getThumbnail();
  if (!thumbnail) return { dataUrl: '' };

  return {
    dataUrl:
      'data:' +
      thumbnail.getContentType() +
      ';base64,' +
      Utilities.base64Encode(thumbnail.getBytes())
  };
}

/*
 * Conserva la categoría enviada por el portal para que la hoja Archivos
 * distinga inspiración y fotografías de los festejados.
 */
function savePhoto_(upload, session) {
  const name = cleanFileName_(upload.name);
  const mime = String(upload.mime || '');
  const encoded = String(upload.base64 || '');
  const type = String(upload.type || 'sin_clasificar');

  if (
    !name ||
    !/^image\/(jpeg|png|webp|heic|heif)$/i.test(mime) ||
    !encoded
  ) {
    throw new Error(
      'Solo se permiten imágenes JPG, PNG, WEBP o HEIC.'
    );
  }

  const bytes = Utilities.base64Decode(encoded);
  if (bytes.length > CONFIG.MAX_FILE_BYTES) {
    throw new Error(name + ' supera el límite de 7 MB por foto.');
  }

  const clientFolder = clientFolder_(session);
  const file = clientFolder.createFile(
    Utilities.newBlob(bytes, mime, name)
  );

  const metadata = {
    id: file.getId(),
    name: file.getName(),
    mime: mime,
    type: type,
    size: bytes.length,
    uploadedAt: new Date().toISOString(),
    driveUrl: file.getUrl()
  };

  sheet_(FILES_SHEET).appendRow([
    session.clientId,
    session.name,
    metadata.name,
    metadata.type,
    metadata.size,
    metadata.driveUrl,
    metadata.uploadedAt
  ]);

  return metadata;
}

function buildResponseRowV2_(options) {
  const data = options.data || {};
  const photos = Array.isArray(options.photos)
    ? options.photos
    : [];

  const values = {
    cliente_id: String(options.clientId || ''),
    nombre_cliente: String(options.clientName || ''),
    estado: String(options.status || 'Borrador'),
    enlaces_fotografias: photos
      .map(photo => String(photo.driveUrl || ''))
      .filter(Boolean)
      .join('\n'),
    ultima_actualizacion: String(options.updatedAt || ''),
    datos_json: JSON.stringify(data),
    fotos_json: JSON.stringify(photos)
  };

  RESPONSE_FORM_FIELDS_V2.forEach(field => {
    values[field] = readableSheetValue_(data[field]);
  });

  return RESPONSE_HEADERS_V2.map(header =>
    values[header] === undefined ? '' : values[header]
  );
}

function readableSheetValue_(value) {
  if (value === true) return 'Sí';
  if (value === false) return 'No';
  if (value === null || value === undefined) return '';
  if (Array.isArray(value) || typeof value === 'object') {
    return JSON.stringify(value);
  }
  return String(value);
}

function assertResponseHeadersV2_() {
  const target = sheet_(RESPONSES_SHEET);
  const lastColumn = target.getLastColumn();
  const headers = lastColumn
    ? target.getRange(1, 1, 1, lastColumn).getValues()[0].map(String)
    : [];

  if (
    headers.length !== RESPONSE_HEADERS_V2.length ||
    headers.some((header, index) =>
      header !== RESPONSE_HEADERS_V2[index]
    )
  ) {
    throw new Error(
      'La hoja Respuestas todavía usa el formato anterior. Ejecuta migrarRespuestasAColumnas() desde Apps Script.'
    );
  }
}

function uniqueSheetName_(book, baseName) {
  let name = baseName.slice(0, 100);
  let counter = 2;

  while (book.getSheetByName(name)) {
    const suffix = ' ' + counter;
    name = baseName.slice(0, 100 - suffix.length) + suffix;
    counter++;
  }

  return name;
}

function removeFileRecord_(photo) {
  const target = sheet_(FILES_SHEET);
  const values = target.getDataRange().getDisplayValues();
  if (values.length < 2) return;

  const headers = values[0];
  const urlIndex = headers.indexOf('url_drive');
  if (urlIndex < 0) return;

  for (let rowIndex = values.length - 1; rowIndex >= 1; rowIndex--) {
    if (
      String(values[rowIndex][urlIndex] || '') ===
      String(photo.driveUrl || '')
    ) {
      target.deleteRow(rowIndex + 1);
    }
  }
}
