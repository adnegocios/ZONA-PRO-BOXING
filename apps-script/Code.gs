/**
 * Zona Pro Boxing - API para Google Sheets
 *
 * Se pega en Extensiones > Apps Script de la hoja "ZONA PRO BOXING - Base de datos"
 * y se implementa como Aplicación web (ejecutar como: Yo, acceso: Cualquier persona).
 * La página de GitHub se conecta a esta URL y manda siempre la clave de acceso.
 *
 * IMPORTANTE: cambia CLAVE antes de implementar. No subas tu clave real a GitHub.
 */
var CLAVE = 'CAMBIA-ESTA-CLAVE';

var TABLAS = {
  students:   { hoja: 'Estudiantes', cols: ['id', 'apellidos', 'nombres', 'cedula', 'fechaNacimiento', 'peso', 'estatura', 'nivel', 'fechaInicio', 'tipoPlan', 'fechaPago', 'estado', 'inactivoDesde', 'notas'] },
  payments:   { hoja: 'Pagos',       cols: ['id', 'studentId', 'nombre', 'valor', 'metodo', 'fecha', 'tipoPlan', 'entrenamiento'] },
  attendance: { hoja: 'Asistencia',  cols: ['id', 'studentId', 'nombre', 'fecha', 'hora'] }
};
var COLUMNAS_NUMERICAS = { valor: true };

function doGet() {
  return salida({ ok: true, mensaje: 'API de Zona Pro Boxing activa' });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return salida({ ok: false, error: 'peticion' });
  }
  if (!CLAVE || CLAVE === 'CAMBIA-ESTA-CLAVE' || req.key !== CLAVE) {
    return salida({ ok: false, error: 'clave' });
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.action === 'list') return salida(listar());
    var t = TABLAS[req.table];
    if (!t) return salida({ ok: false, error: 'tabla' });
    if (req.action === 'upsert') return salida(guardar(t, req.row));
    if (req.action === 'delete') return salida(borrar(t, req.id));
    return salida({ ok: false, error: 'accion' });
  } catch (err) {
    return salida({ ok: false, error: 'servidor', detalle: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecuta esta funcion una vez desde el editor para crear las hojas y los encabezados. */
function configurar() {
  Object.keys(TABLAS).forEach(function (k) {
    var t = TABLAS[k];
    var sh = hoja(t);
    // Todo como texto (fechas ISO, cedulas con ceros) excepto los valores numericos.
    var filas = Math.max(sh.getMaxRows() - 1, 1);
    var f = formatos(t)[0];
    sh.getRange(2, 1, filas, t.cols.length).setNumberFormats(Array.apply(null, Array(filas)).map(function () { return f; }));
  });
}

function hoja(t) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(t.hoja);
  if (!sh) sh = ss.insertSheet(t.hoja);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, t.cols.length).setValues([t.cols]).setFontWeight('bold');
    sh.setFrozenRow(1);
  }
  return sh;
}

function formatos(t) {
  return [t.cols.map(function (c) { return COLUMNAS_NUMERICAS[c] ? '0.00' : '@'; })];
}

function normalizar(v, col) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (COLUMNAS_NUMERICAS[col]) return Number(v) || 0;
  return v === null || v === undefined ? '' : String(v);
}

function leer(t) {
  var sh = hoja(t);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var datos = sh.getRange(2, 1, n - 1, t.cols.length).getValues();
  var out = [];
  datos.forEach(function (fila) {
    if (fila[0] === '' || fila[0] === null) return;
    var o = {};
    t.cols.forEach(function (c, i) { o[c] = normalizar(fila[i], c); });
    out.push(o);
  });
  return out;
}

function listar() {
  return {
    ok: true,
    students: leer(TABLAS.students),
    payments: leer(TABLAS.payments),
    attendance: leer(TABLAS.attendance)
  };
}

function buscarFila(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return -1;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function guardar(t, row) {
  if (!row || !row.id) return { ok: false, error: 'id' };
  var sh = hoja(t);
  var f = buscarFila(sh, row.id);
  var actual = f > 0 ? sh.getRange(f, 1, 1, t.cols.length).getValues()[0] : [];
  var nueva = t.cols.map(function (c, i) {
    if (Object.prototype.hasOwnProperty.call(row, c)) {
      var v = row[c];
      if (COLUMNAS_NUMERICAS[c]) return Number(v) || 0;
      return v === null || v === undefined ? '' : String(v);
    }
    return f > 0 ? actual[i] : '';
  });
  if (f < 0) f = sh.getLastRow() + 1;
  var rango = sh.getRange(f, 1, 1, t.cols.length);
  rango.setNumberFormats(formatos(t));
  rango.setValues([nueva]);
  return { ok: true };
}

function borrar(t, id) {
  var sh = hoja(t);
  var f = buscarFila(sh, id);
  if (f > 0) sh.deleteRow(f);
  return { ok: true };
}
