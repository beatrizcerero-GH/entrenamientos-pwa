/*************************************************************
 *  REGISTRO DE ENTRENAMIENTOS - Google Apps Script
 *  Backend que usa Google Sheets como base de datos.
 *
 *  INSTRUCCIONES RÁPIDAS:
 *   1. Pega este archivo en el editor de Apps Script.
 *   2. Pega el contenido de Index.html en un archivo HTML
 *      llamado exactamente "Index".
 *   3. Ejecuta una vez la función  setup()  (menú Ejecutar).
 *      Autoriza los permisos cuando lo pida.
 *   4. Implementar > Nueva implementación > Aplicación web.
 *      "Ejecutar como: yo"  /  "Acceso: cualquiera con la cuenta".
 *   5. Abre la URL que te da. ¡Listo!
 *************************************************************/

/* ===== CONFIGURACIÓN ===== */
// Déjalo vacío si este script está dentro de la propia hoja de cálculo.
// Si lo creas aparte, pon aquí el ID de la hoja (lo que va entre /d/ y /edit en la URL).
const SS_ID = '';

const SHEETS = {
  USUARIOS:       'Usuarios',
  GRUPOS:         'Grupos',
  EJERCICIOS:     'Ejercicios',
  ENTRENAMIENTOS: 'Entrenamientos',
  SERIES:         'Series'
};

function getSS() {
  var ss = SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No encuentro la hoja de cálculo. Si creaste el script SUELTO (no desde Extensiones › Apps Script de una hoja), copia el ID de tu hoja y pégalo en la variable SS_ID, arriba del Code.gs.');
  }
  return ss;
}

// Crea las pestañas y datos de ejemplo si todavía no existen (autorreparación).
function ensureSetup() {
  var ss = getSS();
  if (!ss.getSheetByName(SHEETS.USUARIOS) || !ss.getSheetByName(SHEETS.GRUPOS)) {
    setup();
  }
}

/* ===== WEB APP ===== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Registro de Entrenamientos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
 *  SETUP: crea las hojas y datos de ejemplo. Ejecutar 1 vez.
 * ========================================================= */
function setup() {
  const ss = getSS();

  crearHoja(ss, SHEETS.USUARIOS,       ['id', 'nombre', 'color', 'avatar_json']);
  crearHoja(ss, SHEETS.GRUPOS,         ['id', 'nombre', 'imagen_url', 'orden', 'activo']);
  crearHoja(ss, SHEETS.EJERCICIOS,     ['id', 'grupo_id', 'nombre', 'imagen_url', 'orden', 'activo']);
  crearHoja(ss, SHEETS.ENTRENAMIENTOS, ['id', 'usuario_id', 'fecha', 'grupo_id', 'hora_inicio', 'hora_fin']);
  crearHoja(ss, SHEETS.SERIES,         ['id', 'entrenamiento_id', 'ejercicio_id', 'serie_num', 'peso', 'repeticiones', 'timestamp']);

  // ---- Usuarios (3) ----
  const shU = ss.getSheetByName(SHEETS.USUARIOS);
  if (shU.getLastRow() < 2) {
    shU.getRange(2, 1, 3, 4).setValues([
      ['u1', 'Usuario 1', '#e85d2a', ''],
      ['u2', 'Usuario 2', '#2a8ae8', ''],
      ['u3', 'Usuario 3', '#2ab07a', '']
    ]);
  }

  // ---- Grupos musculares (configurables) ----
  const shG = ss.getSheetByName(SHEETS.GRUPOS);
  if (shG.getLastRow() < 2) {
    shG.getRange(2, 1, 7, 5).setValues([
      ['pecho',     'Pecho',     '', 1, 'SI'],
      ['hombro',    'Hombro',    '', 2, 'SI'],
      ['espalda',   'Espalda',   '', 3, 'SI'],
      ['pierna',    'Pierna',    '', 4, 'SI'],
      ['biceps',    'Bíceps',    '', 5, 'SI'],
      ['triceps',   'Tríceps',   '', 6, 'SI'],
      ['funcional', 'Funcional', '', 7, 'SI']
    ]);
  }

  // ---- Ejercicios de ejemplo (configurables) ----
  const shE = ss.getSheetByName(SHEETS.EJERCICIOS);
  if (shE.getLastRow() < 2) {
    shE.getRange(2, 1, 12, 6).setValues([
      ['ej01', 'pecho',   'Press banca',             '', 1, 'SI'],
      ['ej02', 'pecho',   'Press inclinado',         '', 2, 'SI'],
      ['ej03', 'pecho',   'Aperturas',               '', 3, 'SI'],
      ['ej04', 'hombro',  'Press militar',           '', 1, 'SI'],
      ['ej05', 'hombro',  'Elevaciones laterales',   '', 2, 'SI'],
      ['ej06', 'espalda', 'Dominadas',               '', 1, 'SI'],
      ['ej07', 'espalda', 'Remo con barra',          '', 2, 'SI'],
      ['ej08', 'pierna',  'Sentadilla',              '', 1, 'SI'],
      ['ej09', 'pierna',  'Prensa',                  '', 2, 'SI'],
      ['ej10', 'biceps',  'Curl con barra',          '', 1, 'SI'],
      ['ej11', 'triceps', 'Extensión en polea',      '', 1, 'SI'],
      ['ej12', 'funcional','Burpees',                '', 1, 'SI']
    ]);
  }

  SpreadsheetApp.getUi && SpreadsheetApp.flush();
  return 'Setup completado. Ya puedes implementar la app web.';
}

function crearHoja(ss, nombre, cabeceras) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras])
      .setFontWeight('bold').setBackground('#1f1f1f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ===== HELPERS ===== */
function leerTabla(nombre) {
  const sh = getSS().getSheetByName(nombre);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function esActivo(v) {
  v = String(v).trim().toLowerCase();
  return v === 'si' || v === 'sí' || v === 'true' || v === '1' || v === 'x';
}

// Devuelve la fecha siempre como texto 'yyyy-MM-dd', venga como texto o como objeto Date.
function fechaStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// Convierte un enlace de Drive en una URL que se puede mostrar como imagen.
function normalizarImagen(url) {
  if (!url) return '';
  url = String(url).trim();
  if (url.indexOf('drive.google') !== -1 || url.indexOf('docs.google') !== -1) {
    const m = url.match(/[-\w]{25,}/);
    if (m) return 'https://drive.google.com/thumbnail?id=' + m[0] + '&sz=w600';
  }
  return url;
}

/* =========================================================
 *  API que llama el front-end (google.script.run)
 * ========================================================= */

function getDatosIniciales() {
  ensureSetup();
  const usuarios = leerTabla(SHEETS.USUARIOS).map(function (u) {
    return { id: u.id, nombre: u.nombre, color: u.color || '#e85d2a' };
  });
  const grupos = leerTabla(SHEETS.GRUPOS)
    .filter(function (g) { return esActivo(g.activo); })
    .sort(function (a, b) { return (a.orden || 0) - (b.orden || 0); })
    .map(function (g) {
      return { id: g.id, nombre: g.nombre, imagen: normalizarImagen(g.imagen_url) };
    });
  return { usuarios: usuarios, grupos: grupos };
}

function getEjercicios(grupoId) {
  return leerTabla(SHEETS.EJERCICIOS)
    .filter(function (e) { return String(e.grupo_id) === String(grupoId) && esActivo(e.activo); })
    .sort(function (a, b) { return (a.orden || 0) - (b.orden || 0); })
    .map(function (e) {
      return { id: e.id, nombre: e.nombre, imagen: normalizarImagen(e.imagen_url) };
    });
}

function iniciarEntrenamiento(usuarioId, grupoId) {
  const sh = getSS().getSheetByName(SHEETS.ENTRENAMIENTOS);
  const id = Utilities.getUuid();
  const now = new Date();
  const fecha = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([id, usuarioId, fecha, grupoId, now, '']);
  return id;
}

function finalizarEntrenamiento(entrenamientoId) {
  const sh = getSS().getSheetByName(SHEETS.ENTRENAMIENTOS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(entrenamientoId)) {
      sh.getRange(i + 1, 6).setValue(new Date()); // columna hora_fin
      return true;
    }
  }
  return false;
}

function guardarSerie(entrenamientoId, ejercicioId, serieNum, peso, reps) {
  const sh = getSS().getSheetByName(SHEETS.SERIES);
  sh.appendRow([
    Utilities.getUuid(), entrenamientoId, ejercicioId,
    serieNum, Number(peso), Number(reps), new Date()
  ]);
  return true;
}

// Series ya registradas en ESTE entrenamiento para ESTE ejercicio (por si se vuelve a entrar).
function getSeriesActuales(entrenamientoId, ejercicioId) {
  return leerTabla(SHEETS.SERIES)
    .filter(function (s) {
      return String(s.entrenamiento_id) === String(entrenamientoId) &&
             String(s.ejercicio_id) === String(ejercicioId);
    })
    .sort(function (a, b) { return (a.serie_num || 0) - (b.serie_num || 0); })
    .map(function (s) {
      return { serie: s.serie_num, peso: s.peso, reps: s.repeticiones };
    });
}

// Recordatorio: último día, pesos de ese día y peso medio de ese usuario en ese ejercicio.
function getHistorialEjercicio(usuarioId, ejercicioId) {
  const entrenos = {};
  leerTabla(SHEETS.ENTRENAMIENTOS).forEach(function (e) {
    if (String(e.usuario_id) === String(usuarioId)) entrenos[e.id] = e;
  });

  const series = leerTabla(SHEETS.SERIES).filter(function (s) {
    return String(s.ejercicio_id) === String(ejercicioId) && entrenos[s.entrenamiento_id];
  });

  if (!series.length) return { hayHistorial: false };

  const pesos = series.map(function (s) { return Number(s.peso); })
                      .filter(function (n) { return !isNaN(n); });
  const pesoMedio = pesos.length ? pesos.reduce(function (a, b) { return a + b; }, 0) / pesos.length : 0;

  let ultimo = null;
  series.forEach(function (s) {
    const e = entrenos[s.entrenamiento_id];
    const f = fechaStr(e.fecha);
    if (!ultimo || f > ultimo.fecha) ultimo = { fecha: f, entrenoId: e.id };
  });

  const seriesUltimas = series
    .filter(function (s) { return String(s.entrenamiento_id) === String(ultimo.entrenoId); })
    .sort(function (a, b) { return (a.serie_num || 0) - (b.serie_num || 0); })
    .map(function (s) { return { serie: s.serie_num, peso: s.peso, reps: s.repeticiones }; });

  return {
    hayHistorial: true,
    ultimaFecha: ultimo.fecha,
    seriesUltimas: seriesUltimas,
    pesoMedio: Math.round(pesoMedio * 10) / 10
  };
}

/* =========================================================
 *  FASE 2: INFORMES Y PROGRESO
 * ========================================================= */

function nombreGrupo(id){
  var g = leerTabla(SHEETS.GRUPOS).filter(function(x){ return String(x.id)===String(id); })[0];
  return g ? g.nombre : id;
}
function nombreEjercicio(id){
  var e = leerTabla(SHEETS.EJERCICIOS).filter(function(x){ return String(x.id)===String(id); })[0];
  return e ? e.nombre : id;
}
function nombreUsuario(id){
  var u = leerTabla(SHEETS.USUARIOS).filter(function(x){ return String(x.id)===String(id); })[0];
  return u ? u.nombre : id;
}
function fmtHora(v){
  if(!v) return '-';
  try{ return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'HH:mm'); }catch(e){ return '-'; }
}
function duracionMin(ini, fin){
  if(!ini || !fin) return null;
  try{ return Math.round((new Date(fin).getTime() - new Date(ini).getTime())/60000); }catch(e){ return null; }
}

// Lista de entrenamientos de un usuario (para elegir informe por entrenamiento)
function getEntrenamientosUsuario(usuarioId){
  return leerTabla(SHEETS.ENTRENAMIENTOS)
    .filter(function(e){ return String(e.usuario_id)===String(usuarioId); })
    .map(function(e){
      return { id:e.id, fecha:fechaStr(e.fecha), grupo:nombreGrupo(e.grupo_id),
               duracion:duracionMin(e.hora_inicio, e.hora_fin) };
    })
    .sort(function(a,b){ return a.fecha < b.fecha ? 1 : -1; });
}

// Fechas distintas con entrenamiento (para informe por día)
function getFechasUsuario(usuarioId){
  var set = {};
  leerTabla(SHEETS.ENTRENAMIENTOS)
    .filter(function(e){ return String(e.usuario_id)===String(usuarioId); })
    .forEach(function(e){ set[fechaStr(e.fecha)] = true; });
  return Object.keys(set).sort().reverse();
}

function detalleEntreno(e){
  var series = leerTabla(SHEETS.SERIES).filter(function(s){ return String(s.entrenamiento_id)===String(e.id); });
  var porEj = {};
  series.forEach(function(s){
    var k = s.ejercicio_id;
    if(!porEj[k]) porEj[k] = { ejercicio:nombreEjercicio(k), series:[] };
    porEj[k].series.push({ serie:s.serie_num, peso:Number(s.peso), reps:Number(s.repeticiones) });
  });
  var ejercicios = Object.keys(porEj).map(function(k){
    var o = porEj[k];
    o.series.sort(function(a,b){ return a.serie-b.serie; });
    o.volumen = o.series.reduce(function(a,s){ return a + s.peso*s.reps; }, 0);
    return o;
  });
  return {
    grupo:nombreGrupo(e.grupo_id), fecha:fechaStr(e.fecha),
    inicio:fmtHora(e.hora_inicio), fin:fmtHora(e.hora_fin),
    duracion:duracionMin(e.hora_inicio, e.hora_fin),
    ejercicios:ejercicios
  };
}

function getDetalleEntrenamiento(entrenamientoId){
  var e = leerTabla(SHEETS.ENTRENAMIENTOS).filter(function(x){ return String(x.id)===String(entrenamientoId); })[0];
  if(!e) return null;
  var d = detalleEntreno(e);
  d.usuario = nombreUsuario(e.usuario_id);
  return d;
}

function getDetalleDia(usuarioId, fecha){
  var entrenos = leerTabla(SHEETS.ENTRENAMIENTOS)
    .filter(function(e){ return String(e.usuario_id)===String(usuarioId) && fechaStr(e.fecha)===String(fecha); });
  return { usuario:nombreUsuario(usuarioId), fecha:fecha, entrenos: entrenos.map(detalleEntreno) };
}

/* ---- Generación de PDF (HTML -> PDF, sin librerías externas) ---- */
function estiloInforme(){
  return '<style>'+
    'body{font-family:Arial,Helvetica,sans-serif;color:#222;padding:24px;}'+
    'h1{color:#e85d2a;border-bottom:3px solid #e85d2a;padding-bottom:6px;font-size:22px;}'+
    'h2{color:#333;margin-top:26px;font-size:17px;border-bottom:1px solid #ccc;padding-bottom:4px;}'+
    'h3{margin:16px 0 4px;font-size:14px;}'+
    'small{color:#888;font-weight:normal;}'+
    'p{font-size:13px;line-height:1.6;}'+
    'table{border-collapse:collapse;width:100%;margin-top:6px;font-size:12px;}'+
    'th,td{border:1px solid #ccc;padding:6px 8px;text-align:center;}'+
    'th{background:#f3f3f3;}'+
    '.resumen{background:#faf2ee;border-left:4px solid #e85d2a;padding:10px 14px;font-size:13px;}'+
    '</style>';
}

function tablaEjercicio(ej){
  var filas = ej.series.map(function(s){
    return '<tr><td>'+s.serie+'</td><td>'+s.peso+'</td><td>'+s.reps+'</td></tr>';
  }).join('');
  return '<h3>'+ej.ejercicio+' <small>(Volumen total: '+ej.volumen+' kg)</small></h3>'+
    '<table><tr><th>Serie</th><th>Peso (kg)</th><th>Reps</th></tr>'+filas+'</table>';
}

function htmlInformeEntreno(d){
  var cuerpo = d.ejercicios.map(tablaEjercicio).join('');
  return estiloInforme()+
    '<h1>Entrenamiento de '+d.grupo+'</h1>'+
    '<div class="resumen"><b>Usuario:</b> '+d.usuario+' &nbsp;|&nbsp; <b>Fecha:</b> '+d.fecha+'<br>'+
    '<b>Inicio:</b> '+d.inicio+' &nbsp;|&nbsp; <b>Fin:</b> '+d.fin+' &nbsp;|&nbsp; '+
    '<b>Duración:</b> '+(d.duracion!=null?d.duracion+' min':'-')+'</div>'+ cuerpo;
}

function htmlInformeDia(dd){
  var bloques = dd.entrenos.map(function(d){
    return '<h2>'+d.grupo+' &nbsp;<small>('+d.inicio+'–'+d.fin+', '+
      (d.duracion!=null?d.duracion+' min':'-')+')</small></h2>'+
      d.ejercicios.map(tablaEjercicio).join('');
  }).join('');
  return estiloInforme()+
    '<h1>Resumen del día '+dd.fecha+'</h1>'+
    '<div class="resumen"><b>Usuario:</b> '+dd.usuario+' &nbsp;|&nbsp; '+
    '<b>Entrenamientos:</b> '+dd.entrenos.length+'</div>'+ bloques;
}

// tipo: 'entreno' | 'dia'
function generarInformePDF(tipo, idOrFecha, usuarioId){
  var html, nombre;
  if(tipo==='entreno'){
    var d = getDetalleEntrenamiento(idOrFecha);
    html = htmlInformeEntreno(d);
    nombre = 'Entrenamiento_'+d.grupo+'_'+d.fecha;
  } else {
    var dd = getDetalleDia(usuarioId, idOrFecha);
    html = htmlInformeDia(dd);
    nombre = 'Dia_'+dd.fecha;
  }
  var blob = Utilities.newBlob(html, 'text/html', nombre+'.html').getAs('application/pdf');
  return { b64:Utilities.base64Encode(blob.getBytes()), mime:'application/pdf', nombre:nombre+'.pdf' };
}

/* ---- Progreso: evolución por ejercicio ---- */
function getProgresoEjercicio(usuarioId, ejercicioId){
  var entrenos = {};
  leerTabla(SHEETS.ENTRENAMIENTOS).forEach(function(e){
    if(String(e.usuario_id)===String(usuarioId)) entrenos[e.id] = fechaStr(e.fecha);
  });
  var series = leerTabla(SHEETS.SERIES).filter(function(s){
    return String(s.ejercicio_id)===String(ejercicioId) && entrenos[s.entrenamiento_id];
  });
  var porFecha = {};
  series.forEach(function(s){
    var f = entrenos[s.entrenamiento_id];
    if(!porFecha[f]) porFecha[f] = { pesos:[], reps:0, vol:0 };
    porFecha[f].pesos.push(Number(s.peso));
    porFecha[f].reps += Number(s.repeticiones);
    porFecha[f].vol  += Number(s.peso)*Number(s.repeticiones);
  });
  return Object.keys(porFecha).sort().map(function(f){
    var o = porFecha[f];
    var media = o.pesos.reduce(function(a,b){ return a+b; }, 0)/o.pesos.length;
    return {
      fecha:f,
      pesoMax: Math.max.apply(null, o.pesos),
      pesoMedio: Math.round(media*10)/10,
      reps:o.reps, volumen:o.vol
    };
  });
}
