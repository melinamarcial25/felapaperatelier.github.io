/*
 * Sustituye únicamente tu función login_ actual por esta versión y vuelve
 * a desplegar la aplicación web. Permite acceder con cliente_id o nombre.
 */
function login_(request) {
  const identifier = String(request.clientId || '').trim();
  const password = String(request.password || '');
  if (!identifier || !password) {
    throw new Error('Escribe tu ID o nombre de cliente y contraseña.');
  }

  const normalizedIdentifier = normalize_(identifier);
  const rows = objects_(sheet_(CLIENTS_SHEET));
  const matches = rows.filter(row =>
    truthy_(row.activo) &&
    (
      normalize_(row.cliente_id) === normalizedIdentifier ||
      normalize_(row.nombre) === normalizedIdentifier
    )
  );

  // Si dos clientes tienen exactamente el mismo nombre, deberán usar su ID.
  if (matches.length > 1) {
    throw new Error('Hay más de un cliente con ese nombre. Ingresa tu ID de cliente.');
  }

  const client = matches[0];
  if (!client || !safeEqual_(String(client.password_hash), sha256_(password))) {
    throw new Error('ID, nombre o contraseña incorrectos.');
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('session:' + token, JSON.stringify({
    clientId: String(client.cliente_id),
    name: String(client.nombre)
  }), CONFIG.SESSION_SECONDS);

  return {
    token: token,
    client: {
      id: String(client.cliente_id),
      name: String(client.nombre)
    }
  };
}
