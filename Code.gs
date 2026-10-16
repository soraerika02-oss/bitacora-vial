/**
 * BITÁCORA VIAL — Conexión con Google Sheets
 * Pegue este código en la hoja de cálculo: Extensiones > Apps Script.
 * 1. Cambie la CLAVE por una propia y compártala solo con el equipo.
 * 2. Implementar > Nueva implementación > Aplicación web
 *    Ejecutar como: Yo   |   Quién tiene acceso: Cualquier usuario
 * 3. Copie la URL que termina en /exec y péguela en Ajustes de la app.
 */

const CLAVE = 'CAMBIE-ESTA-CLAVE';
const NOMBRE_HOJA = 'Registros';
const NOMBRE_CARPETA = 'Bitácora Vial - Fotos';

const COLUMNAS = [
  'id', 'proyecto', 'fecha', 'medido_por', 'vereda_sector', 'via', 'abscisa_referencia',
  'latitud', 'longitud', 'precision_gps_m', 'tipo_pavimento', 'tipo_dano', 'severidad',
  'mediciones_LxAxP', 'sobrecorte', 'longitud_fisuras_m', 'area_dano_m2', 'area_corte_m2', 'volumen_m3',
  'mezcla_asfaltica_t', 'riego_liga_L', 'base_granular_m3', 'concreto_m3', 'material_granular_m3',
  'estado', 'observaciones', 'fotos', 'mapa', 'actualizado', 'eliminado', 'json'
];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.clave !== CLAVE) return salida({ ok: false, error: 'Clave incorrecta' });
  if (p.accion === 'listar') {
    const datos = obtenerHoja().getDataRange().getValues();
    const iJson = COLUMNAS.indexOf('json');
    const registros = [];
    for (let i = 1; i < datos.length; i++) {
      const txt = datos[i][iJson];
      if (!txt) continue;
      try { registros.push(JSON.parse(txt)); } catch (err) { /* fila editada a mano: se ignora */ }
    }
    return salida({ ok: true, registros: registros });
  }
  return salida({ ok: true, mensaje: 'Conexión correcta con la Bitácora Vial' });
}

function doPost(e) {
  const candado = LockService.getScriptLock();
  candado.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.clave !== CLAVE) return salida({ ok: false, error: 'Clave incorrecta' });
    if (d.accion !== 'guardar' || !d.registro || !d.registro.id) return salida({ ok: false, error: 'Solicitud no válida' });

    const r = d.registro;
    const urls = {};
    if (d.fotos && d.fotos.length) {
      const carpeta = obtenerCarpeta();
      d.fotos.forEach(function (f) {
        const blob = Utilities.newBlob(Utilities.base64Decode(f.base64), f.mime || 'image/jpeg', f.nombre || (f.id + '.jpg'));
        urls[f.id] = carpeta.createFile(blob).getUrl();
      });
    }
    (r.fotos || []).forEach(function (f) { if (urls[f.id]) f.url = urls[f.id]; });
    delete r.pendiente;
    guardarRegistro(r);
    return salida({ ok: true, fotos: urls });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  } finally {
    candado.releaseLock();
  }
}

function guardarRegistro(r) {
  const hoja = obtenerHoja();
  const fila = construirFila(r);
  const celda = hoja.getRange('A:A').createTextFinder(r.id).matchEntireCell(true).findNext();
  if (celda) hoja.getRange(celda.getRow(), 1, 1, fila.length).setValues([fila]);
  else hoja.appendRow(fila);
  guardarMediciones(r);
}

/* Pestaña "Mediciones": una fila por cada medida, con largo, ancho y profundidad separados */
const NOMBRE_HOJA_MED = 'Mediciones';
const COLUMNAS_MED = [
  'id', 'proyecto', 'fecha', 'medido_por', 'vereda_sector', 'via', 'abscisa_referencia', 'tipo_pavimento', 'tipo_dano', 'severidad',
  'medida_n', 'largo_m', 'ancho_m', 'profundidad_cm', 'area_m2', 'area_corte_m2', 'volumen_m3', 'estado'
];

function obtenerHojaMed() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA_MED);
  if (!hoja) hoja = libro.insertSheet(NOMBRE_HOJA_MED);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLUMNAS_MED);
    hoja.getRange(1, 1, 1, COLUMNAS_MED.length).setFontWeight('bold').setBackground('#1f2933').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
    hoja.hideColumns(1);
  }
  return hoja;
}

function filasMediciones(r) {
  if (r.eliminado) return [];
  const n = function (v) { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : ''; };
  const red = function (x, d) { const f = Math.pow(10, d); return Math.round(x * f) / f; };
  const txt = function (v) { const s = v == null ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; };
  const s = r.aplicaCorte && r.factores ? Number(r.factores.sobrecorte) || 0 : 0;
  return (r.mediciones || []).map(function (m, i) {
    const l = n(m.l), a = n(m.a), p = n(m.p);
    const area = l !== '' && a !== '' ? red(l * a, 3) : '';
    const corte = area !== '' ? red((l + 2 * s) * (a + 2 * s), 3) : '';
    const vol = corte !== '' && p !== '' ? red(corte * p / 100, 4) : '';
    return [r.id, txt(r.proyecto), r.fecha, txt(r.usuario), txt(r.sector), txt(r.via), txt(r.abscisa), r.tipoPav, r.tipoDano, r.severidad,
      i + 1, l, a, p, area, corte, vol, r.estado];
  });
}

function guardarMediciones(r) {
  const hoja = obtenerHojaMed();
  const celdas = hoja.getRange('A:A').createTextFinder(r.id).matchEntireCell(true).findAll();
  celdas.map(function (c) { return c.getRow(); }).sort(function (a, b) { return b - a; })
    .forEach(function (fila) { hoja.deleteRow(fila); });
  const filas = filasMediciones(r);
  if (filas.length) hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, COLUMNAS_MED.length).setValues(filas);
}

/** Ejecute esta función una vez para llenar la pestaña Mediciones con los registros que ya existían. */
function reconstruirMediciones() {
  const datos = obtenerHoja().getDataRange().getValues();
  const iJson = COLUMNAS.indexOf('json');
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const vieja = libro.getSheetByName(NOMBRE_HOJA_MED);
  if (vieja) libro.deleteSheet(vieja);
  const hoja = obtenerHojaMed();
  let filas = [];
  for (let i = 1; i < datos.length; i++) {
    try { filas = filas.concat(filasMediciones(JSON.parse(datos[i][iJson]))); } catch (err) { /* fila sin json válido */ }
  }
  if (filas.length) hoja.getRange(2, 1, filas.length, COLUMNAS_MED.length).setValues(filas);
}

function construirFila(r) {
  const c = r.calc || {};
  const red = function (n, d) { const f = Math.pow(10, d || 3); return Math.round((Number(n) || 0) * f) / f; };
  const txt = function (v) { const s = v == null ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; };
  const mediciones = (r.mediciones || []).map(function (m) { return m.l + ' x ' + (m.a || '-') + ' x ' + (m.p || '-') + ' cm'; }).join('; ');
  const fotos = (r.fotos || []).map(function (f) { return f.url; }).filter(String).join(' | ');
  const mapa = r.lat && r.lon ? 'https://www.google.com/maps?q=' + r.lat + ',' + r.lon : '';
  return [
    r.id, txt(r.proyecto), r.fecha, txt(r.usuario), txt(r.sector), txt(r.via), txt(r.abscisa),
    r.lat ? Number(r.lat) : '', r.lon ? Number(r.lon) : '', r.precision || '', r.tipoPav, r.tipoDano, r.severidad,
    txt(mediciones), r.aplicaCorte ? 'SÍ' : 'NO', red(c.lineal, 2), red(c.aDano, 2), red(c.aCorte, 2), red(c.vol, 3),
    red(c.mezcla, 2), red(c.liga, 1), red(c.base, 3), red(c.concreto, 3), red(c.granular, 3),
    r.estado, txt(r.obs), fotos, mapa, r.actualizado, r.eliminado ? 'SÍ' : 'NO', JSON.stringify(r)
  ];
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) hoja = libro.insertSheet(NOMBRE_HOJA);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLUMNAS);
    hoja.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold').setBackground('#1f2933').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
    hoja.hideColumns(1);                 // id interno
    hoja.hideColumns(COLUMNAS.length);   // json de respaldo para la app
  }
  return hoja;
}

function obtenerCarpeta() {
  const it = DriveApp.getFoldersByName(NOMBRE_CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(NOMBRE_CARPETA);
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecute esta función una vez desde el editor para autorizar los permisos de Hoja y Drive. */
function autorizar() {
  obtenerHoja();
  obtenerCarpeta();
}
