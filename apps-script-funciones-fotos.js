/*
 * PEGA ESTE ARCHIVO COMPLETO AL FINAL DE CODE.GS.
 * No elimines doPost ni ninguna otra función.
 */

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

    // Se mueve a la papelera para que la propietaria pueda recuperarla.
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

function removeFileRecord_(photo) {
  const target = sheet_(FILES_SHEET);
  const values = target.getDataRange().getDisplayValues();
  if (values.length < 2) return;

  const headers = values[0];
  const urlIndex = headers.indexOf('url_drive');
  if (urlIndex < 0) return;

  for (
    let rowIndex = values.length - 1;
    rowIndex >= 1;
    rowIndex--
  ) {
    if (
      String(values[rowIndex][urlIndex] || '') ===
      String(photo.driveUrl || '')
    ) {
      target.deleteRow(rowIndex + 1);
    }
  }
}
