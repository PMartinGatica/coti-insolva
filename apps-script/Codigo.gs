/**
 * ==================================================================
 * INSOLVA - Sincronizacion Google Sheet -> GitHub
 * ==================================================================
 *
 * Lee las hojas Encabezados / Productos / Instalacion / Tiempos,
 * las une por "N° COTIZACION" y publica un unico JSON en:
 *
 *   PMartinGatica/coti-insolva  ->  cotizaciones/cotizaciones_data.json
 *
 * Ese archivo es el que consume cotizacion.html.
 *
 * ------------------------------------------------------------------
 * PUESTA EN MARCHA (una sola vez)
 * ------------------------------------------------------------------
 * 1. Extensiones > Apps Script > pega este archivo.
 * 2. Ejecuta  guardarToken()  una vez, con el PAT nuevo escrito
 *    en la constante de abajo. Queda guardado en las Propiedades del
 *    script (cifradas) y despues BORRAS el valor de la constante.
 * 3. Ejecuta  verificar()  y revisa el registro: te dice cuantas
 *    cotizaciones salen y que filas quedaron afuera, sin publicar nada.
 * 4. Ejecuta  syncAll()  una vez y acepta los permisos.
 * 5. Activadores (reloj) > Agregar activador:
 *       funcion: syncAll
 *       origen:  Basado en tiempo
 *       cada:    1 hora  (o lo que prefieras)
 *
 * IMPORTANTE: el activador tiene que ser "instalable" (el del reloj).
 * Un onEdit(e) simple NO sirve: los activadores simples corren sin
 * autorizacion y tienen prohibido usar UrlFetchApp, asi que la
 * sincronizacion automatica al editar fallaba siempre en silencio.
 * ==================================================================
 */

var GITHUB_REPO   = 'PMartinGatica/coti-insolva';
var GITHUB_PATH   = 'cotizaciones/cotizaciones_data.json';
var GITHUB_BRANCH = 'main';

var SHEETS = {
  ENCA: 'Encabezados',
  PROD: 'Productos',
  INST: 'Instalación',
  TIEM: 'Tiempos'
};


/**
 * Ejecutar UNA vez con el token pegado aca, y despues vaciar la cadena.
 * El token queda en las Propiedades del script, no en el codigo.
 */
function guardarToken() {
  var TOKEN_NUEVO = '';   // <-- pega el PAT aca, ejecuta, y volve a vaciarlo

  if (!TOKEN_NUEVO) {
    throw new Error('Pega el token en TOKEN_NUEVO antes de ejecutar guardarToken().');
  }
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', TOKEN_NUEVO);
  Logger.log('Token guardado. Borra ahora el valor de TOKEN_NUEVO del codigo.');
}

function getToken_() {
  var t = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!t) {
    throw new Error('Falta el token. Ejecuta guardarToken() primero.');
  }
  return t;
}


// ==================================================================
// PUNTOS DE ENTRADA
// ==================================================================

/** Esta es la funcion que engancha el activador de tiempo. */
function syncAll() {
  sincronizar_(false);
}

/** Simulacro: arma el JSON y lo reporta, pero NO toca GitHub. */
function verificar() {
  sincronizar_(true);
}


// ==================================================================
// NUCLEO
// ==================================================================

function sincronizar_(soloVerificar) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var enca = leerHoja_(ss, SHEETS.ENCA);
  if (!enca) {
    throw new Error('No existe la hoja "' + SHEETS.ENCA + '".');
  }
  if (!enca.filas.length) {
    throw new Error('La hoja "' + SHEETS.ENCA + '" no tiene datos.');
  }

  var colKey = buscarCol_(enca.head, ['N COTIZACION', 'NRO COTIZACION', 'N PRESUPUESTO', 'COTIZACION']);
  if (colKey < 0) {
    throw new Error('No encuentro la columna "N° COTIZACION" en ' + SHEETS.ENCA + '.');
  }

  var colCliente = buscarCol_(enca.head, ['CLIENTE']);
  var colFecha   = buscarCol_(enca.head, ['FECHA']);
  var colTipo    = buscarCol_(enca.head, ['TIPO']);
  var colEstado  = buscarCol_(enca.head, ['ESTADO']);

  // Las hojas de detalle se indexan por clave: evita recorrerlas
  // enteras por cada cotizacion.
  // CODIGO es opcional: si la hoja Productos no tiene esa columna,
  // el campo sale vacio y el cotizador lo deja editable a mano.
  var prodPorKey = indexar_(ss, SHEETS.PROD, ['PRODUCTO', 'CANTIDAD', 'PRECIO UNITARIO', 'SUBTOTAL', 'CODIGO']);
  var instPorKey = indexar_(ss, SHEETS.INST, ['ITEM', 'VALOR']);
  var tiemPorKey = indexar_(ss, SHEETS.TIEM, ['ETAPA', 'HORAS']);

  var cotizaciones = [];
  var omitidas = [];
  var num = 0;

  for (var i = 0; i < enca.filas.length; i++) {
    var row = enca.filas[i];
    var key = texto_(row[colKey]);

    // Sin N° COTIZACION no hay forma de unir las otras hojas.
    if (!key) {
      var pista = texto_(row[colCliente]) || texto_(row[0]) || '(fila vacia)';
      omitidas.push('fila ' + (i + 2) + ': ' + pista);
      continue;
    }

    num++;

    var join = key.toUpperCase();   // la union no distingue mayusculas

    // r = [PRODUCTO, CANTIDAD, PRECIO UNITARIO, SUBTOTAL, CODIGO]
    var productos = (prodPorKey[join] || []).map(function (r) {
      var cant = parseInt(r[1], 10) || 1;
      var unit = monto_(r[2]);
      var sub  = monto_(r[3]);
      return {
        codigo:   texto_(r[4]),
        nombre:   texto_(r[0]),
        cant:     cant,
        unit:     unit,
        subtotal: sub || (cant * unit)
      };
    }).filter(function (p) { return p.nombre; });

    // r = [ITEM, VALOR]
    var instalacion = (instPorKey[join] || []).map(function (r) {
      return {
        item:  texto_(r[0]) || 'Instalación y mano de obra',
        valor: monto_(r[1])
      };
    });

    // r = [ETAPA, HORAS]
    var tiempos = (tiemPorKey[join] || []).map(function (r) {
      return {
        etapa: texto_(r[0]) || 'Instalación',
        horas: parseInt(r[1], 10) || 0
      };
    });

    var anticipo = productos.reduce(function (a, p) { return a + p.subtotal; }, 0);
    var saldo    = instalacion.reduce(function (a, x) { return a + x.valor; }, 0);

    cotizaciones.push({
      num:            num,
      nroPresupuesto: key,
      cliente:        texto_(row[colCliente]),
      fecha:          fecha_(row[colFecha]),
      tipo:           texto_(row[colTipo]),
      estado:         texto_(row[colEstado]) || 'Pendiente',
      productos:      productos,
      instalacion:    instalacion,
      tiempos:        tiempos,
      anticipo:       anticipo,   // 100% de los equipos
      saldo:          saldo,      // instalacion / mano de obra
      total:          anticipo + saldo
    });
  }

  // ------------------------------------------------------------------
  // Informe
  // ------------------------------------------------------------------
  Logger.log('Cotizaciones generadas: ' + cotizaciones.length);

  if (omitidas.length) {
    Logger.log('OMITIDAS por no tener N° COTIZACION (' + omitidas.length + '):');
    omitidas.forEach(function (o) { Logger.log('   - ' + o); });
    Logger.log('   Cargales un N° en "' + SHEETS.ENCA + '" para que aparezcan en el cotizador.');
  }

  var sinProd = cotizaciones.filter(function (c) { return !c.productos.length; });
  if (sinProd.length) {
    Logger.log('SIN PRODUCTOS asociados en la hoja "' + SHEETS.PROD + '" (' + sinProd.length + '):');
    sinProd.forEach(function (c) { Logger.log('   - ' + c.nroPresupuesto + ' / ' + c.cliente); });
  }

  var payload = JSON.stringify(cotizaciones, null, 2);

  if (soloVerificar) {
    Logger.log('--- MODO VERIFICACION: no se publico nada en GitHub ---');
    return;
  }

  publicar_(payload);
}


// ==================================================================
// GITHUB
// ==================================================================

function publicar_(payload) {

  var url = 'https://api.github.com/repos/' + GITHUB_REPO + '/contents/' + GITHUB_PATH;

  var headers = {
    'Authorization': 'Bearer ' + getToken_(),
    'Accept': 'application/vnd.github+json'
  };

  // 1) Leer el archivo actual para obtener su SHA (y comparar).
  var sha = '';
  var actual = '';

  var get = UrlFetchApp.fetch(url + '?ref=' + GITHUB_BRANCH, {
    method: 'GET',
    headers: headers,
    muteHttpExceptions: true
  });

  if (get.getResponseCode() === 200) {
    var meta = JSON.parse(get.getContentText());
    sha = meta.sha || '';
    try {
      actual = Utilities.newBlob(
        Utilities.base64Decode(meta.content || '')
      ).getDataAsString('UTF-8');
    } catch (e) {
      actual = '';
    }
  } else if (get.getResponseCode() !== 404) {
    throw new Error('GitHub GET ' + get.getResponseCode() + ': ' + get.getContentText());
  }

  // 2) Si no cambio nada, no commitear. Evita un commit por cada
  //    pasada del activador.
  if (actual && actual.replace(/\s+$/, '') === payload.replace(/\s+$/, '')) {
    Logger.log('Sin cambios: no se genero commit.');
    return;
  }

  // 3) Publicar.
  var body = {
    message: 'Sync cotizaciones ' + Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm'),
    content: Utilities.base64Encode(payload, Utilities.Charset.UTF_8),
    branch:  GITHUB_BRANCH
  };
  if (sha) body.sha = sha;

  var put = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: headers,
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = put.getResponseCode();

  if (code >= 200 && code < 300) {
    Logger.log('✓ Publicado en ' + GITHUB_REPO + '/' + GITHUB_PATH);
  } else {
    throw new Error('GitHub PUT ' + code + ': ' + put.getContentText());
  }
}


// ==================================================================
// UTILIDADES
// ==================================================================

/** Normaliza un encabezado: mayusculas, sin acentos, sin simbolos. */
function normHead_(v) {
  return (v == null ? '' : String(v))
    .normalize('NFD').replace(/[̀-ͯ]/g, '')  // saca acentos
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, '')                        // saca °, ., etc.
    .replace(/\s+/g, ' ')
    .trim();
}

/** Devuelve el indice de la primera columna cuyo encabezado coincida. */
function buscarCol_(head, alias) {
  for (var a = 0; a < alias.length; a++) {
    var objetivo = normHead_(alias[a]);
    for (var i = 0; i < head.length; i++) {
      if (normHead_(head[i]) === objetivo) return i;
    }
  }
  // Segunda pasada, mas laxa: por contenido. Exige 3+ caracteres para
  // no aparear encabezados sueltos tipo "N" con "N COTIZACION".
  for (var a2 = 0; a2 < alias.length; a2++) {
    var obj2 = normHead_(alias[a2]);
    for (var j = 0; j < head.length; j++) {
      var h = normHead_(head[j]);
      if (h.length < 3) continue;
      if (h.indexOf(obj2) > -1 || obj2.indexOf(h) > -1) return j;
    }
  }
  return -1;
}

/** Lee una hoja y separa encabezado de filas. Tolera nombres con/sin acento. */
function leerHoja_(ss, nombre) {
  var sheet = ss.getSheetByName(nombre);

  if (!sheet) {
    // Reintento ignorando acentos y mayusculas
    var objetivo = normHead_(nombre);
    var todas = ss.getSheets();
    for (var i = 0; i < todas.length; i++) {
      if (normHead_(todas[i].getName()) === objetivo) { sheet = todas[i]; break; }
    }
  }
  if (!sheet) return null;

  var d = sheet.getDataRange().getValues();
  return {
    head:  d.length ? d[0] : [],
    filas: d.length > 1 ? d.slice(1) : []
  };
}

/**
 * Indexa una hoja de detalle por N° COTIZACION.
 * Devuelve { clave: [ [col1, col2, ...], ... ] } respetando el orden
 * de `campos`, resuelto por encabezado y no por posicion fija.
 */
function indexar_(ss, nombre, campos) {
  var hoja = leerHoja_(ss, nombre);
  if (!hoja) {
    Logger.log('AVISO: no existe la hoja "' + nombre + '".');
    return {};
  }

  var colKey = buscarCol_(hoja.head, ['N COTIZACION', 'NRO COTIZACION', 'N PRESUPUESTO', 'COTIZACION']);
  if (colKey < 0) {
    Logger.log('AVISO: la hoja "' + nombre + '" no tiene columna N° COTIZACION.');
    return {};
  }

  var cols = campos.map(function (c) { return buscarCol_(hoja.head, [c]); });

  var out = {};
  var huerfanas = 0;

  hoja.filas.forEach(function (row) {
    var key = texto_(row[colKey]).toUpperCase();
    if (!key) { huerfanas++; return; }
    if (!out[key]) out[key] = [];
    out[key].push(cols.map(function (ci) { return ci < 0 ? '' : row[ci]; }));
  });

  if (huerfanas) {
    Logger.log('AVISO: "' + nombre + '" tiene ' + huerfanas +
               ' fila(s) sin N° COTIZACION; no se pueden asociar.');
  }
  return out;
}

function texto_(v) {
  if (v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return fecha_(v);
  return String(v).trim();
}

/** Formatea fechas como dd/MM/yyyy; deja pasar el texto tal cual. */
function fecha_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
  }
  return String(v).trim();
}

/**
 * Convierte a numero. Acepta el numero crudo de la celda, "$1.143.000",
 * "1,234.56" y devuelve 0 para "Pendiente", "?" o vacio.
 */
function monto_(v) {
  if (typeof v === 'number') return isFinite(v) ? Math.round(v) : 0;
  if (v == null) return 0;

  var s = String(v).replace(/[^0-9.,]/g, '').trim();
  if (!s) return 0;

  var u = s.lastIndexOf('.');
  var c = s.lastIndexOf(',');

  if (u > -1 && c > -1) {
    // Hay punto y coma: el ultimo de los dos es el separador decimal
    s = (u > c)
      ? s.replace(/,/g, '')                        // 1,234.56
      : s.replace(/\./g, '').replace(',', '.');    // 1.234,56
  } else if (c > -1) {
    var pc = s.split(',');
    s = (pc.length === 2 && pc[1].length <= 2)
      ? s.replace(',', '.')                        // 1234,50
      : s.replace(/,/g, '');                       // 1,234,567
  } else if (u > -1) {
    var pu = s.split('.');
    if (pu.length !== 2 || pu[1].length > 2) {
      s = s.replace(/\./g, '');                    // 1.143.000
    }
  }

  return Math.round(parseFloat(s)) || 0;
}
