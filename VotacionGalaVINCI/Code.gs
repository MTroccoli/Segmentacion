/**
 * Gala VINCI - Sistema de Votacion
 * Septiembre 2026
 *
 * INSTRUCCIONES DE CONFIGURACION:
 *
 * OPCION A - Manual:
 * 1. Crear un nuevo proyecto en Google Apps Script (script.google.com)
 * 2. Copiar este archivo como Code.gs
 * 3. Copiar Index.html como archivo HTML en el proyecto
 * 4. Ejecutar la funcion inicializarSistema() UNA VEZ desde el editor
 * 5. Desplegar > Nueva implementacion > Aplicacion web
 *
 * OPCION B - Con clasp (recomendada):
 * 1. npm install -g @google/clasp
 * 2. clasp login
 * 3. clasp create --type webapp --title "Gala VINCI"
 *    (o clasp clone <scriptId> si ya existe el proyecto)
 * 4. clasp push
 * 5. clasp deploy
 * 6. Ejecutar inicializarSistema() una vez desde el editor
 *
 * DESPLIEGUE:
 * - Ejecutar como: Tu cuenta
 * - Acceso: "Cualquier persona de tu organizacion" (auto-detecta email)
 *           O "Cualquier persona" (requiere login manual)
 */

const CONFIG = {
  ADMIN_PASSWORD: 'gvarguru26',
  PESO_PONDERADO: 0.80,
  PESO_REGULAR: 0.20,
  CACHE_EMAILS_VOTARON: 300,
  CACHE_RESULTADOS: 15,
  SLIDES_ID: '1Mkj7SD3SOSkDjyAjGjIIJob-TqvHu9rrrAAjsBvGNXI'
};

const PAISES = [
  'Argentina',
  'Chile',
  'CIB',
  'Colombia',
  'Espana',
  'Holding',
  'Mexico',
  'Peru',
  'Turquia',
  'Uruguay',
  'Venezuela'
];

const HOJAS = {
  VOTOS: 'Votos',
  LISTA_PONDERADA: 'Lista Ponderada',
  IMAGENES: 'Imagenes'
};

// =============================================
// PUNTO DE ENTRADA
// =============================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gala VINCI - Votacion')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =============================================
// INICIALIZACION (ejecutar una sola vez)
// =============================================

function inicializarSistema() {
  var ss = obtenerHoja_();
  var nombres = ss.getSheets().map(function(s) { return s.getName(); });

  if (nombres.indexOf(HOJAS.VOTOS) === -1) {
    var sv = ss.insertSheet(HOJAS.VOTOS);
    sv.appendRow(['Timestamp', 'Email', 'PaisVotante', 'PaisVotado', 'EsPonderado']);
    sv.getRange('A1:E1').setFontWeight('bold');
  }

  if (nombres.indexOf(HOJAS.LISTA_PONDERADA) === -1) {
    var sp = ss.insertSheet(HOJAS.LISTA_PONDERADA);
    sp.appendRow(['Email', 'Nombre', 'Pais', 'Cargo']);
    sp.getRange('A1:D1').setFontWeight('bold');
    sp.appendRow(['director.innovacion@bbva.com', 'Ejemplo Director', 'Espana', 'Director de Innovacion']);
  }

  try {
    var h1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
    if (h1 && ss.getSheets().length > 1) ss.deleteSheet(h1);
  } catch(e) {}

  Logger.log('Sistema listo. Hoja: ' + ss.getUrl());
  return { url: ss.getUrl(), id: ss.getId() };
}

// =============================================
// UTILIDADES PRIVADAS
// =============================================

function obtenerHoja_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  var ss = SpreadsheetApp.create('Gala VINCI - Votacion');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function leerHoja_(nombre) {
  var sheet = obtenerHoja_().getSheetByName(nombre);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function obtenerEmailsVotaron_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('emails_votaron');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
  if (!sheet || sheet.getLastRow() <= 1) {
    cache.put('emails_votaron', '[]', CONFIG.CACHE_EMAILS_VOTARON);
    return [];
  }

  var emails = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return r[0] ? r[0].toString().toLowerCase() : ''; })
    .filter(function(e) { return e; });

  cache.put('emails_votaron', JSON.stringify(emails), CONFIG.CACHE_EMAILS_VOTARON);
  return emails;
}

function obtenerSetPonderados_() {
  var sheet = obtenerHoja_().getSheetByName(HOJAS.LISTA_PONDERADA);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return r[0] ? r[0].toString().toLowerCase() : ''; })
    .filter(function(e) { return e; });
}

function invalidarCacheVotos_() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['emails_votaron', 'resultados']);
}

function invalidarCacheResultados_() {
  CacheService.getScriptCache().remove('resultados');
}

// =============================================
// API PUBLICA - DETECCION DE EMAIL
// =============================================

function obtenerEmailUsuario() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email && email.indexOf('@') !== -1) {
      return { ok: true, email: email };
    }
  } catch(e) {}
  return { ok: false, email: '' };
}

// =============================================
// API PUBLICA - REGISTRO
// =============================================

function obtenerPaises() {
  return PAISES;
}

function validarUsuario(email, pais) {
  email = email.toLowerCase().trim();
  if (!email || email.indexOf('@') === -1) {
    return { ok: false, msg: 'Ingresa un email valido.' };
  }
  if (PAISES.indexOf(pais) === -1) {
    return { ok: false, msg: 'Selecciona un pais valido.' };
  }

  var votaron = obtenerEmailsVotaron_();
  if (votaron.indexOf(email) !== -1) {
    return { ok: false, msg: 'Ya registraste tu voto. Solo se permite un voto por persona.', yaVoto: true };
  }

  var ponderados = obtenerSetPonderados_();
  var esPonderado = ponderados.indexOf(email) !== -1;

  return { ok: true, email: email, pais: pais, esPonderado: esPonderado };
}

// =============================================
// API PUBLICA - VOTACION
// =============================================

function obtenerPaisesParaVotar(paisUsuario) {
  return PAISES.filter(function(p) { return p !== paisUsuario; });
}

function registrarVoto(email, paisVotado, paisVotante) {
  email = email.toLowerCase().trim();

  if (PAISES.indexOf(paisVotado) === -1) {
    return { ok: false, msg: 'Pais no valido.' };
  }
  if (paisVotado === paisVotante) {
    return { ok: false, msg: 'No puedes votar por tu propio pais.' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
    if (!sheet) {
      return { ok: false, msg: 'Error de configuracion: no se encontro la hoja de votos. Ejecuta inicializarSistema().' };
    }

    if (sheet.getLastRow() > 1) {
      var emails = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < emails.length; i++) {
        if (emails[i][0] && emails[i][0].toString().toLowerCase() === email) {
          return { ok: false, msg: 'Ya registraste tu voto anteriormente.' };
        }
      }
    }

    var ponderados = obtenerSetPonderados_();
    var esPond = ponderados.indexOf(email) !== -1;

    sheet.appendRow([
      new Date(), email, paisVotante, paisVotado, esPond ? 'SI' : 'NO'
    ]);

    invalidarCacheVotos_();

    return {
      ok: true,
      msg: 'Voto registrado exitosamente!',
      paisVotado: paisVotado
    };
  } catch(e) {
    return { ok: false, msg: 'Error: ' + (e.message || 'Hubo un error, por favor reintenta en unos segundos.') };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// =============================================
// API PUBLICA - ADMIN
// =============================================

function loginAdmin(password) {
  return password === CONFIG.ADMIN_PASSWORD;
}

function obtenerResultadosAdmin(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };

  var cache = CacheService.getScriptCache();
  var cached = cache.get('resultados');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var votos = leerHoja_(HOJAS.VOTOS);

  var totalPond = 0;
  var totalReg = 0;
  var conteosPond = {};
  var conteosReg = {};
  PAISES.forEach(function(p) { conteosPond[p] = 0; conteosReg[p] = 0; });

  votos.forEach(function(v) {
    var pais = v.PaisVotado;
    if (v.EsPonderado === 'SI') {
      totalPond++;
      if (conteosPond[pais] !== undefined) conteosPond[pais]++;
    } else {
      totalReg++;
      if (conteosReg[pais] !== undefined) conteosReg[pais]++;
    }
  });

  var ranking = PAISES.map(function(pais) {
    var vp = conteosPond[pais];
    var vr = conteosReg[pais];
    var sp = totalPond > 0 ? (vp / totalPond) * CONFIG.PESO_PONDERADO * 100 : 0;
    var sr = totalReg  > 0 ? (vr / totalReg)  * CONFIG.PESO_REGULAR   * 100 : 0;

    return {
      pais: pais,
      votosPond: vp, votosReg: vr, totalVotos: vp + vr,
      scorePond: Math.round(sp * 100) / 100,
      scoreReg:  Math.round(sr * 100) / 100,
      scoreTotal: Math.round((sp + sr) * 100) / 100
    };
  });

  ranking.sort(function(a, b) { return b.scoreTotal - a.scoreTotal; });

  var detallePond = votos
    .filter(function(v) { return v.EsPonderado === 'SI'; })
    .map(function(v) {
      return {
        email: v.Email,
        paisVotante: v.PaisVotante,
        paisVotado: v.PaisVotado,
        fecha: v.Timestamp
      };
    });

  var resultado = {
    ok: true,
    ranking: ranking,
    stats: {
      total: votos.length,
      ponderados: totalPond,
      regulares: totalReg,
      paises: PAISES.length
    },
    detallePond: detallePond,
    actualizado: new Date().toISOString()
  };

  try { cache.put('resultados', JSON.stringify(resultado), CONFIG.CACHE_RESULTADOS); } catch(e) {}
  return resultado;
}

function obtenerListaPonderada(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  return { ok: true, lista: leerHoja_(HOJAS.LISTA_PONDERADA) };
}

function agregarUsuarioPonderado(password, datos) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };

  var ponderados = obtenerSetPonderados_();
  if (ponderados.indexOf(datos.email.toLowerCase()) !== -1) {
    return { ok: false, msg: 'El usuario ya esta en la lista.' };
  }

  obtenerHoja_().getSheetByName(HOJAS.LISTA_PONDERADA).appendRow([
    datos.email, datos.nombre || '', datos.pais || '', datos.cargo || ''
  ]);
  invalidarCacheResultados_();
  return { ok: true };
}

function eliminarUsuarioPonderado(password, email) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var sheet = obtenerHoja_().getSheetByName(HOJAS.LISTA_PONDERADA);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      invalidarCacheResultados_();
      return { ok: true };
    }
  }
  return { ok: false, msg: 'Usuario no encontrado.' };
}

function resetearVotos(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  invalidarCacheVotos_();
  return { ok: true };
}

// =============================================
// IMAGENES DESDE GOOGLE SLIDES
// =============================================

/**
 * Ejecutar UNA VEZ (o cuando cambien las slides).
 * Exporta cada slide como PNG a una carpeta de Drive
 * y guarda el mapeo Pais -> URL en la hoja "Imagenes".
 *
 * PREREQUISITO: Habilitar el servicio avanzado de Slides:
 *   En el editor de Apps Script > Servicios (+) > Google Slides API > Agregar
 *
 * Despues de ejecutar, revisar la hoja "Imagenes" y
 * corregir el mapeo si el orden de slides no coincide con PAISES.
 */
function exportarImagenesSlides() {
  var presId = CONFIG.SLIDES_ID;
  var token = ScriptApp.getOAuthToken();

  var presUrl = 'https://slides.googleapis.com/v1/presentations/' + presId;
  var presResp = UrlFetchApp.fetch(presUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (presResp.getResponseCode() !== 200) {
    Logger.log('Error accediendo a la presentacion: ' + presResp.getContentText());
    return { ok: false, msg: 'No se pudo acceder a la presentacion. Verifica el ID y que tengas permisos de lectura.' };
  }

  var presData = JSON.parse(presResp.getContentText());
  var slides = presData.slides;

  var folders = DriveApp.getFoldersByName('Gala VINCI - Imagenes');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Gala VINCI - Imagenes');
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var ss = obtenerHoja_();
  var sheet = ss.getSheetByName(HOJAS.IMAGENES);
  if (!sheet) {
    sheet = ss.insertSheet(HOJAS.IMAGENES);
  } else {
    sheet.clear();
  }
  sheet.appendRow(['Pais', 'ImagenURL']);
  sheet.getRange('A1:B1').setFontWeight('bold');

  var total = 0;
  for (var i = 0; i < slides.length; i++) {
    var pageId = slides[i].objectId;

    var thumbUrl = 'https://slides.googleapis.com/v1/presentations/' + presId +
                   '/pages/' + pageId + '/thumbnail?thumbnailProperties.thumbnailSize=MEDIUM';
    var thumbResp = UrlFetchApp.fetch(thumbUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (thumbResp.getResponseCode() !== 200) continue;

    var thumbData = JSON.parse(thumbResp.getContentText());
    var imageBlob = UrlFetchApp.fetch(thumbData.contentUrl).getBlob();

    var fileName = 'slide_' + (i + 1) + '.png';
    imageBlob.setName(fileName);

    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    var file = folder.createFile(imageBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var pais = i < PAISES.length ? PAISES[i] : 'Slide ' + (i + 1);
    var url = 'https://drive.google.com/uc?export=view&id=' + file.getId();

    sheet.appendRow([pais, url]);
    total++;
  }

  Logger.log('Exportadas ' + total + ' imagenes.');
  return { ok: true, total: total, msg: 'Revisa la hoja "Imagenes" y corrige el mapeo Pais si es necesario.' };
}

function obtenerImagenesPaises() {
  var sheet = obtenerHoja_().getSheetByName(HOJAS.IMAGENES);
  if (!sheet || sheet.getLastRow() <= 1) return {};

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var mapping = {};
  data.forEach(function(row) {
    if (row[0] && row[1]) mapping[row[0]] = row[1];
  });
  return mapping;
}
