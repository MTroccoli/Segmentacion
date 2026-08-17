/**
 * Gala VINCI - Sistema de Votacion
 *
 * INSTRUCCIONES DE CONFIGURACION:
 * 1. Crear un nuevo proyecto en Google Apps Script (script.google.com)
 * 2. Copiar este archivo como Code.gs
 * 3. Copiar Index.html como archivo HTML en el proyecto
 * 4. Ejecutar la funcion inicializarSistema() UNA VEZ desde el editor
 * 5. Desplegar > Nueva implementacion > Aplicacion web
 *    - Ejecutar como: Tu cuenta
 *    - Acceso: Cualquier persona
 * 6. Abrir la hoja de calculo generada y configurar:
 *    - Pestana "Iniciativas": agregar las iniciativas con imagenes
 *    - Pestana "Lista Ponderada": agregar los emails con voto ponderado
 *
 * NOTA SOBRE ESCALABILIDAD:
 * Apps Script soporta ~30 ejecuciones simultaneas. Para 500 usuarios
 * concurrentes se usa cache agresivo y funciones rapidas.
 * Si necesitas mas capacidad, considera migrar a Firebase/Supabase.
 */

const CONFIG = {
  ADMIN_PASSWORD: 'gvarguru26',
  CACHE_DURACION: 60,
  PESO_PONDERADO: 0.80,
  PESO_REGULAR: 0.20
};

const HOJAS = {
  INICIATIVAS: 'Iniciativas',
  VOTOS: 'Votos',
  LISTA_PONDERADA: 'Lista Ponderada'
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

  if (nombres.indexOf(HOJAS.INICIATIVAS) === -1) {
    var si = ss.insertSheet(HOJAS.INICIATIVAS);
    si.appendRow(['ID', 'Pais', 'Nombre', 'Descripcion', 'ImagenURL', 'Activa']);
    si.getRange('A1:F1').setFontWeight('bold');
    var ejemplos = [
      [1, 'Espana',    'IA Banking 360',     'Asistente bancario con inteligencia artificial generativa para asesoramiento personalizado',            '', 'SI'],
      [2, 'Mexico',    'Green Finance Hub',   'Plataforma de financiamiento verde para PyMEs con medicion de impacto ambiental',                     '', 'SI'],
      [3, 'Colombia',  'Crypto Bridge',       'Puente entre banca tradicional y activos digitales con cumplimiento regulatorio',                     '', 'SI'],
      [4, 'Peru',      'Rural Connect',       'Inclusion financiera digital para comunidades rurales sin acceso bancario',                           '', 'SI'],
      [5, 'Argentina', 'DataShield',          'Sistema de prevencion de fraude con machine learning en tiempo real',                                 '', 'SI'],
      [6, 'Chile',     'Open Banking+',       'Ecosistema de APIs abiertas para integracion con fintechs locales',                                  '', 'SI'],
      [7, 'Uruguay',   'Smart Mortgage',      'Proceso hipotecario 100% digital con valuacion automatica por IA',                                   '', 'SI'],
      [8, 'Turquia',   'NeoPayments',         'Sistema de pagos instantaneos con tecnologia blockchain',                                            '', 'SI'],
      [9, 'Venezuela', 'MicroLend',           'Plataforma de microcreditos digitales con scoring alternativo basado en datos transaccionales',       '', 'SI']
    ];
    ejemplos.forEach(function(r) { si.appendRow(r); });
    si.autoResizeColumns(1, 6);
  }

  if (nombres.indexOf(HOJAS.VOTOS) === -1) {
    var sv = ss.insertSheet(HOJAS.VOTOS);
    sv.appendRow(['Timestamp', 'Email', 'PaisVotante', 'IDIniciativa', 'PaisIniciativa', 'NombreIniciativa', 'EsPonderado']);
    sv.getRange('A1:G1').setFontWeight('bold');
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

// =============================================
// API PUBLICA - REGISTRO
// =============================================

function obtenerPaises() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('paises');
  if (cached) return JSON.parse(cached);

  var ini = leerHoja_(HOJAS.INICIATIVAS);
  var set = {};
  ini.forEach(function(i) { if (i.Activa === 'SI') set[i.Pais] = true; });
  var paises = Object.keys(set).sort();
  cache.put('paises', JSON.stringify(paises), 300);
  return paises;
}

function validarUsuario(email, pais) {
  email = email.toLowerCase().trim();
  if (!email || email.indexOf('@') === -1) {
    return { ok: false, msg: 'Ingresa un email valido.' };
  }

  var votos = leerHoja_(HOJAS.VOTOS);
  var yaVoto = votos.some(function(v) {
    return v.Email && v.Email.toLowerCase() === email;
  });
  if (yaVoto) {
    return { ok: false, msg: 'Ya registraste tu voto. Solo se permite un voto por persona.', yaVoto: true };
  }

  var lista = leerHoja_(HOJAS.LISTA_PONDERADA);
  var esPonderado = lista.some(function(u) {
    return u.Email && u.Email.toLowerCase() === email;
  });

  return { ok: true, email: email, pais: pais, esPonderado: esPonderado };
}

// =============================================
// API PUBLICA - VOTACION
// =============================================

function obtenerIniciativas(paisUsuario) {
  var cache = CacheService.getScriptCache();
  var key = 'ini_' + paisUsuario;
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  var ini = leerHoja_(HOJAS.INICIATIVAS);
  var filtradas = ini
    .filter(function(i) { return i.Activa === 'SI' && i.Pais !== paisUsuario; })
    .map(function(i) {
      return {
        id: i.ID,
        pais: i.Pais,
        nombre: i.Nombre,
        descripcion: i.Descripcion,
        imagenUrl: i.ImagenURL
      };
    });

  cache.put(key, JSON.stringify(filtradas), CONFIG.CACHE_DURACION);
  return filtradas;
}

function registrarVoto(email, idIniciativa, paisVotante) {
  email = email.toLowerCase().trim();
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var ss = obtenerHoja_();
    var sheet = ss.getSheetByName(HOJAS.VOTOS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase() === email) {
        return { ok: false, msg: 'Ya registraste tu voto anteriormente.' };
      }
    }

    var inis = leerHoja_(HOJAS.INICIATIVAS);
    var ini = null;
    for (var j = 0; j < inis.length; j++) {
      if (inis[j].ID == idIniciativa) { ini = inis[j]; break; }
    }
    if (!ini) return { ok: false, msg: 'Iniciativa no encontrada.' };
    if (ini.Pais === paisVotante) return { ok: false, msg: 'No puedes votar por tu propio pais.' };

    var lista = leerHoja_(HOJAS.LISTA_PONDERADA);
    var esPond = lista.some(function(u) {
      return u.Email && u.Email.toLowerCase() === email;
    });

    sheet.appendRow([
      new Date(), email, paisVotante, idIniciativa,
      ini.Pais, ini.Nombre, esPond ? 'SI' : 'NO'
    ]);

    try { CacheService.getScriptCache().remove('resultados'); } catch(e) {}

    return {
      ok: true,
      msg: 'Voto registrado exitosamente!',
      iniciativa: { nombre: ini.Nombre, pais: ini.Pais }
    };
  } catch(e) {
    return { ok: false, msg: 'Hubo un error, por favor reintenta en unos segundos.' };
  } finally {
    lock.releaseLock();
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
  var inis = leerHoja_(HOJAS.INICIATIVAS).filter(function(i) { return i.Activa === 'SI'; });

  var totalPond = votos.filter(function(v) { return v.EsPonderado === 'SI'; }).length;
  var totalReg  = votos.filter(function(v) { return v.EsPonderado === 'NO'; }).length;

  var ranking = inis.map(function(ini) {
    var votosIni = votos.filter(function(v) { return v.IDIniciativa == ini.ID; });
    var vp = votosIni.filter(function(v) { return v.EsPonderado === 'SI'; }).length;
    var vr = votosIni.filter(function(v) { return v.EsPonderado === 'NO'; }).length;

    var sp = totalPond > 0 ? (vp / totalPond) * CONFIG.PESO_PONDERADO * 100 : 0;
    var sr = totalReg  > 0 ? (vr / totalReg)  * CONFIG.PESO_REGULAR   * 100 : 0;

    return {
      id: ini.ID, pais: ini.Pais, nombre: ini.Nombre,
      votosPond: vp, votosReg: vr, totalVotos: votosIni.length,
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
        iniciativa: v.NombreIniciativa,
        paisIni: v.PaisIniciativa,
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
      iniciativas: inis.length
    },
    detallePond: detallePond,
    actualizado: new Date().toISOString()
  };

  try { cache.put('resultados', JSON.stringify(resultado), 15); } catch(e) {}
  return resultado;
}

function obtenerTodasIniciativas(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var inis = leerHoja_(HOJAS.INICIATIVAS);
  return {
    ok: true,
    iniciativas: inis.map(function(i) {
      return {
        id: i.ID, pais: i.Pais, nombre: i.Nombre,
        descripcion: i.Descripcion, imagenUrl: i.ImagenURL,
        activa: i.Activa === 'SI'
      };
    })
  };
}

function agregarIniciativa(password, datos) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var ss = obtenerHoja_();
  var sheet = ss.getSheetByName(HOJAS.INICIATIVAS);
  var last = sheet.getLastRow();
  var lastId = last > 1 ? sheet.getRange(last, 1).getValue() : 0;
  sheet.appendRow([lastId + 1, datos.pais, datos.nombre, datos.descripcion, datos.imagenUrl || '', 'SI']);
  CacheService.getScriptCache().removeAll(['paises']);
  return { ok: true, id: lastId + 1 };
}

function actualizarIniciativa(password, datos) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var ss = obtenerHoja_();
  var sheet = ss.getSheetByName(HOJAS.INICIATIVAS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == datos.id) {
      if (datos.pais !== undefined)        sheet.getRange(i+1, 2).setValue(datos.pais);
      if (datos.nombre !== undefined)      sheet.getRange(i+1, 3).setValue(datos.nombre);
      if (datos.descripcion !== undefined) sheet.getRange(i+1, 4).setValue(datos.descripcion);
      if (datos.imagenUrl !== undefined)   sheet.getRange(i+1, 5).setValue(datos.imagenUrl);
      if (datos.activa !== undefined)      sheet.getRange(i+1, 6).setValue(datos.activa ? 'SI' : 'NO');
      break;
    }
  }
  CacheService.getScriptCache().removeAll(['paises']);
  return { ok: true };
}

function obtenerListaPonderada(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  return { ok: true, lista: leerHoja_(HOJAS.LISTA_PONDERADA) };
}

function agregarUsuarioPonderado(password, datos) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var lista = leerHoja_(HOJAS.LISTA_PONDERADA);
  var existe = lista.some(function(u) {
    return u.Email && u.Email.toLowerCase() === datos.email.toLowerCase();
  });
  if (existe) return { ok: false, msg: 'El usuario ya esta en la lista.' };

  var ss = obtenerHoja_();
  ss.getSheetByName(HOJAS.LISTA_PONDERADA).appendRow([
    datos.email, datos.nombre || '', datos.pais || '', datos.cargo || ''
  ]);
  return { ok: true };
}

function eliminarUsuarioPonderado(password, email) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var ss = obtenerHoja_();
  var sheet = ss.getSheetByName(HOJAS.LISTA_PONDERADA);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, msg: 'Usuario no encontrado.' };
}

function subirImagenDrive(base64Data, fileName, mimeType) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
    return { ok: true, url: url };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function resetearVotos(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var ss = obtenerHoja_();
  var sheet = ss.getSheetByName(HOJAS.VOTOS);
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  CacheService.getScriptCache().remove('resultados');
  return { ok: true };
}
