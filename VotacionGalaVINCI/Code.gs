/**
 * Gala VINCI - Sistema de Votacion
 * Septiembre 2026
 *
 * INSTRUCCIONES DE CONFIGURACION:
 * 1. Crear un nuevo proyecto en Google Apps Script (script.google.com)
 * 2. Copiar este archivo como Code.gs
 * 3. Copiar Index.html como archivo HTML en el proyecto
 * 4. Ejecutar la funcion inicializarSistema() UNA VEZ desde el editor
 * 5. Desplegar > Nueva implementacion > Aplicacion web
 *    - Ejecutar como: Tu cuenta
 *    - Acceso: Cualquier persona de tu organizacion (para auto-detectar email)
 *              O "Cualquier persona" (requiere login manual)
 * 6. Abrir la hoja de calculo generada y configurar:
 *    - Pestana "Lista Ponderada": agregar los emails con voto ponderado
 *
 * DETECCION AUTOMATICA DE EMAIL:
 * Si despliegas con acceso "Cualquier persona de [tu organizacion]",
 * el sistema detecta automaticamente el email del usuario logueado.
 * Si despliegas con acceso "Cualquier persona", el usuario debera
 * ingresar su email manualmente.
 */

const CONFIG = {
  ADMIN_PASSWORD: 'gvarguru26',
  CACHE_DURACION: 60,
  PESO_PONDERADO: 0.80,
  PESO_REGULAR: 0.20
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
  'Venezuela'
];

const HOJAS = {
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

function obtenerPaisesParaVotar(paisUsuario) {
  return PAISES.filter(function(p) { return p !== paisUsuario; });
}

function registrarVoto(email, paisVotado, paisVotante) {
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

    if (PAISES.indexOf(paisVotado) === -1) {
      return { ok: false, msg: 'Pais no valido.' };
    }
    if (paisVotado === paisVotante) {
      return { ok: false, msg: 'No puedes votar por tu propio pais.' };
    }

    var lista = leerHoja_(HOJAS.LISTA_PONDERADA);
    var esPond = lista.some(function(u) {
      return u.Email && u.Email.toLowerCase() === email;
    });

    sheet.appendRow([
      new Date(), email, paisVotante, paisVotado, esPond ? 'SI' : 'NO'
    ]);

    try { CacheService.getScriptCache().remove('resultados'); } catch(e) {}

    return {
      ok: true,
      msg: 'Voto registrado exitosamente!',
      paisVotado: paisVotado
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

  var totalPond = votos.filter(function(v) { return v.EsPonderado === 'SI'; }).length;
  var totalReg  = votos.filter(function(v) { return v.EsPonderado === 'NO'; }).length;

  var ranking = PAISES.map(function(pais) {
    var votosP = votos.filter(function(v) { return v.PaisVotado === pais; });
    var vp = votosP.filter(function(v) { return v.EsPonderado === 'SI'; }).length;
    var vr = votosP.filter(function(v) { return v.EsPonderado === 'NO'; }).length;

    var sp = totalPond > 0 ? (vp / totalPond) * CONFIG.PESO_PONDERADO * 100 : 0;
    var sr = totalReg  > 0 ? (vr / totalReg)  * CONFIG.PESO_REGULAR   * 100 : 0;

    return {
      pais: pais,
      votosPond: vp, votosReg: vr, totalVotos: votosP.length,
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

  try { cache.put('resultados', JSON.stringify(resultado), 15); } catch(e) {}
  return resultado;
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
