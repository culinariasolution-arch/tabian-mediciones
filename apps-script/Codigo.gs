// =====================================================
//  TABIAN · API de la app de mediciones
//  Va en: hoja "Mediciones Tabian" > Extensiones > Apps Script
//  Implementar como Aplicación web:
//    Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
//  (la seguridad la pone la comprobación del token de Google de abajo)
// =====================================================

const CLIENT_ID = '237381041487-4k9g42ec2f9n9bh6p4ndgk4mnoo0v3hp.apps.googleusercontent.com';
const DOMINIO = 'tabian.es';
const HOJAS_PERMITIDAS = ['Registro', 'Tipos', 'Ofertas', 'Ajustes'];
const CARPETAS_PERMITIDAS = [
  '1k9M5WWDNolywj923m2QU6OkhhMgbh0TQ', // 02 OBRAS / 2026
  '124TQ25CQDb12VxurtRrNOK2qG_P1k_Sb', // 02 OBRAS / 2027
  '1E8ogOkPsPH-1rMHvW18ymjjWBauFSBWe', // 02 OBRAS
  '1v6cy2sFHstUbnc5kTY9bTHMuZIolK7rL'  // 01 COMERCIAL
];

function doGet() {
  return json_({ ok: true, app: 'Mediciones Tabian' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const usuario = comprobarToken_(req.token);
    const datos = ejecutar_(req.accion, req.datos || {}, usuario);
    return json_({ ok: true, datos: datos });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err), auth: !!(err && err.auth) });
  }
}

// ---------- seguridad ----------
function comprobarToken_(token) {
  if (!token) throw errAuth_('Falta iniciar sesión');
  const cache = CacheService.getScriptCache();
  const clave = 'tk_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 40);
  const guardado = cache.get(clave);
  if (guardado) return guardado;
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw errAuth_('Sesión caducada');
  const info = JSON.parse(r.getContentText());
  const email = String(info.email || '').toLowerCase();
  if (info.aud !== CLIENT_ID) throw errAuth_('Aplicación no autorizada');
  if (String(info.email_verified) !== 'true') throw errAuth_('Correo no verificado');
  if (!email.endsWith('@' + DOMINIO)) throw errAuth_('Solo cuentas @' + DOMINIO);
  const resto = Number(info.exp) - Math.floor(Date.now() / 1000);
  if (resto <= 0) throw errAuth_('Sesión caducada');
  cache.put(clave, email, Math.min(resto, 600));
  return email;
}
function errAuth_(m) { const e = new Error(m); e.auth = true; return e; }

function hojaDeRango_(rango) {
  const m = String(rango).match(/^(\w+)!/);
  if (!m || HOJAS_PERMITIDAS.indexOf(m[1]) < 0) throw new Error('Rango no permitido: ' + rango);
  return m[1];
}

// ---------- acciones ----------
function ejecutar_(accion, d, usuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  switch (accion) {

    case 'yo':
      return { email: usuario };

    case 'get_values': {
      hojaDeRango_(d.range);
      const vals = ss.getRange(d.range).getDisplayValues();
      // quitar filas vacías del final, como hace la API de Sheets
      while (vals.length && vals[vals.length - 1].every(v => v === '')) vals.pop();
      return { values: vals.map(fila => { const f = fila.slice(); while (f.length && f[f.length - 1] === '') f.pop(); return f; }) };
    }

    case 'update_values': {
      hojaDeRango_(d.range);
      const lock = LockService.getScriptLock(); lock.waitLock(15000);
      try {
        const rango = ss.getRange(d.range);
        const vals = d.values.map(f => f.map(v => (v === null || v === undefined) ? '' : v));
        rango.offset(0, 0, vals.length, vals[0].length).setValues(vals);
        SpreadsheetApp.flush();
        return { status: 'success' };
      } finally { lock.releaseLock(); }
    }

    case 'borrar_fila': {
      if (HOJAS_PERMITIDAS.indexOf(d.hoja) < 0) throw new Error('Hoja no permitida');
      const lock = LockService.getScriptLock(); lock.waitLock(15000);
      try {
        const sh = ss.getSheetByName(d.hoja);
        const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(d.id)) { sh.deleteRow(i + 2); return { borrada: true }; }
        }
        return { borrada: false };
      } finally { lock.releaseLock(); }
    }

    case 'guardar_pdf': {
      if (CARPETAS_PERMITIDAS.indexOf(d.carpeta) < 0) throw new Error('Carpeta no permitida');
      const blob = Utilities.newBlob(Utilities.base64Decode(d.base64), 'application/pdf', d.nombre);
      const f = DriveApp.getFolderById(d.carpeta).createFile(blob);
      return { id: f.getId(), url: f.getUrl() };
    }
  }
  throw new Error('Acción desconocida: ' + accion);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Ejecuta esta función una vez desde el editor para dar permisos (Drive y conexiones externas)
function autorizar() {
  DriveApp.getFolderById(CARPETAS_PERMITIDAS[0]).getName();
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true });
  Logger.log('Permisos OK');
}
