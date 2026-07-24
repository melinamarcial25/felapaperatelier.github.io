/*
 * Sustituye en Google Apps Script las funciones load_ y save_ por estas
 * versiones. Conservan el estado visible en Sheets y bloquean definitivamente
 * un expediente después de marcarlo como "Listo para diseñar".
 */

function load_(request) {
  const session = session_(request.token);
  const response = responseFor_(session.clientId);
  return {
    client: { id: session.clientId, name: session.name },
    draft: response ? parseJson_(response.datos_json, {}) : {},
    photos: response ? parseJson_(response.fotos_json, []) : [],
    status: response ? String(response.estado || 'Borrador') : 'Borrador',
    updatedAt: response ? String(response.ultima_actualizacion || '') : ''
  };
}

function save_(request) {
  const session = session_(request.token);
  const data = request.data;
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('Datos del formulario inválidos.');
  }

  const uploads = Array.isArray(request.uploads) ? request.uploads : [];
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
  const requestedStatus = allowedStatuses.indexOf(data.estado_proyecto) >= 0
    ? data.estado_proyecto
    : 'Borrador';

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
    const newPhotos = uploads.map(upload => savePhoto_(upload, session));
    const now = new Date().toISOString();

    const row = [
      session.clientId,
      session.name,
      requestedStatus,
      field_(data, 'nombres_pareja'),
      field_(data, 'fecha_ceremonia'),
      field_(data, 'nombre_lugar'),
      field_(data, 'tipo_confirmacion'),
      field_(data, 'dress_code'),
      now,
      JSON.stringify(data),
      JSON.stringify(oldPhotos.concat(newPhotos))
    ];

    const target = sheet_(RESPONSES_SHEET);
    if (existing) {
      target.getRange(existing._row, 1, 1, row.length).setValues([row]);
    } else {
      target.appendRow(row);
    }

    return {
      updatedAt: now,
      status: requestedStatus,
      photos: oldPhotos.concat(newPhotos)
    };
  } finally {
    lock.releaseLock();
  }
}
