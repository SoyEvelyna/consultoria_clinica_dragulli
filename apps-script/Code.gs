/**
 * Clínica Dra. Gulli — backend de Google Apps Script.
 *
 * Se pega en un proyecto INDEPENDIENTE de Apps Script (script.google.com),
 * creado desde una cuenta Gmail personal con permiso de edición sobre el
 * Sheet, y se publica como Web App. La web del tablero le habla por HTTP.
 * Va en una cuenta personal porque el Workspace de soyevelyna.com bloquea
 * la publicación de Web Apps por política de la organización.
 *
 * DISEÑO:
 * - Un solo Sheet con TRES tableros de trabajo: Administración ("02"),
 *   Médicos ("03") y Comercial ("04"). La web los muestra con un selector.
 * - "01 I Plan de trabajo" son las reuniones, compartidas por los tres.
 * - Cada hoja de proceso tiene sus propias columnas: se ubican por el texto
 *   del encabezado (TAREA, RESPONSABLE, INICIO, DEADLINE/FIN, ESTADO…), así
 *   que si se agrega o mueve una columna no hay que tocar código.
 * - Las tareas se leen y ESCRIBEN en la hoja que corresponde. Lo que no tiene
 *   columna (notas de la web) va a las pestañas "WebApp - *".
 * - Los links del entregable van en OBSERVACIONES, como texto clickeable.
 * - Autenticación: un token compartido (ver ACCESS_TOKEN), el mismo que se
 *   configura en la web. No hay login individual.
 */

/* Para ejecutar a mano desde el editor (el menú no muestra funciones con "_").
   Van primero porque el editor corre la primera función del archivo. */
function verTableros() { Logger.log(JSON.stringify(diagnostico_(), null, 2)); }

/* Suma una prioridad al trimestre de un tablero (por ejemplo una campaña
   puntual). Se edita NUEVA_PRIORIDAD y se ejecuta agregarPrioridad(): escribe
   la fila en el bloque PRIORIDADES de la hoja y suma la etiqueta al
   desplegable de la columna PRIORIDAD. */
var NUEVA_PRIORIDAD = { tablero: "admin", texto: "DÍA DE LA MADRE" };

/* Unifica el desplegable de RESPONSABLE de un tablero: toma los nombres que
   ya ofrece la hoja más todos los que están efectivamente usados en la
   columna, y aplica esa única lista a toda la columna de tareas. Hace falta
   porque quedaron filas con reglas más angostas que la realidad: nombres como
   Ailén o Dani estaban escritos en celdas pero el desplegable los rechazaba,
   así que al asignarlos desde la web terminaban en OBSERVACIONES. */
function unificarResponsables() {
  var resumen = TABLEROS.map(function (t) {
    var L = procesoLayout_(t.id);
    if (L.C.responsable === undefined) return { tablero: t.id, error: "sin columna RESPONSABLE" };
    var col = L.C.responsable + 1;
    var nombres = [], vistos = {};
    function sumar(v) {
      var x = String(v === null || v === undefined ? "" : v).trim();
      if (!x || vistos[norm_(x)]) return;
      vistos[norm_(x)] = true;
      nombres.push(x);
    }
    L.tasks.forEach(function (task) {
      dropdownValues_(L.sheet.getRange(task.row, col)).forEach(sumar);
      sumar(L.sheet.getRange(task.row, col).getValue());
    });
    if (!nombres.length) return { tablero: t.id, error: "no encontré nombres" };
    ["Anto/Caro", "Caro/Anto"].forEach(sumar);
    nombres.sort(function (a, b) { return a.localeCompare(b); });
    /* setAllowInvalid(true): la lista sugiere, no bloquea. Si se elige una
       combinación que no está, igual se escribe en vez de perderse. */
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(nombres, true).setAllowInvalid(true).build();
    var desde = L.C.header + 2;
    L.sheet.getRange(desde, col, L.sheet.getMaxRows() - desde + 1, 1).setDataValidation(regla);
    return { tablero: t.id, hoja: L.sheet.getName(), opciones: nombres };
  });
  Logger.log(JSON.stringify(resumen, null, 2));
}

function agregarPrioridad() {
  var t = tablero_(NUEVA_PRIORIDAD.tablero);
  var sheet = hojaPorPrefijo_(t.prefijo);
  var values = sheet.getDataRange().getValues();
  var C = colsProceso_(values);
  var prioridadesRow = findLabelRow_(values, "PRIORIDADES");
  var iniciativasRow = findLabelRow_(values, "INICIATIVAS");
  var fin = iniciativasRow !== -1 ? iniciativasRow : C.header;

  /* La última fila numerada del bloque, y en qué columnas van número y texto. */
  var ultima = -1, ultimoN = 0, colN = -1, colTexto = -1;
  for (var r = (prioridadesRow !== -1 ? prioridadesRow + 1 : 0); r < fin; r++) {
    var n = null, cN = -1, cT = -1;
    for (var c = 0; c < Math.min(values[r].length, 6); c++) {
      var val = cell_(values[r], c);
      if (val === null) continue;
      var num = typeof val === "number" ? val : Number(String(val).replace(/[.)]/g, "").trim());
      if (n === null && !isNaN(num) && num >= 1 && num <= 9 && String(val).length <= 3) { n = num; cN = c; continue; }
      if (n !== null && cT === -1) { cT = c; break; }
    }
    if (n === null || cT === -1) continue;
    if (n > ultimoN) { ultimoN = n; ultima = r; colN = cN; colTexto = cT; }
  }
  if (ultima === -1) throw new Error("No encuentro el bloque de prioridades en " + sheet.getName());
  if (ultimoN >= MAX_PRIORIDADES) throw new Error("Ya hay " + ultimoN + " prioridades; subí MAX_PRIORIDADES");

  var filaOrigen = ultima + 1;
  var filaNueva = filaOrigen + 1;
  var ancho = Math.max(sheet.getLastColumn(), colTexto + 1);
  sheet.insertRowAfter(filaOrigen);
  sheet.getRange(filaOrigen, 1, 1, ancho).copyTo(sheet.getRange(filaNueva, 1, 1, ancho), { formatOnly: true });

  /* El texto de la prioridad vive en celdas combinadas: se repite el combinado. */
  sheet.getRange(filaOrigen, 1, 1, ancho).getMergedRanges().forEach(function (m) {
    if (m.getNumRows() !== 1) return;
    sheet.getRange(filaNueva, m.getColumn(), 1, m.getNumColumns()).merge();
  });

  sheet.getRange(filaNueva, colN + 1).setValue(ultimoN + 1);
  sheet.getRange(filaNueva, colTexto + 1).setValue(NUEVA_PRIORIDAD.texto);

  /* La columna PRIORIDAD ofrece ahora también la nueva etiqueta. */
  var etiquetas = [];
  for (var i = 1; i <= ultimoN + 1; i++) etiquetas.push("P" + i);
  if (C.prioridad !== undefined) {
    var desde = C.header + 3; // el encabezado bajó una fila al insertar
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(etiquetas, true).setAllowInvalid(true).build();
    sheet.getRange(desde, C.prioridad + 1, Math.max(sheet.getMaxRows() - desde + 1, 1), 1).setDataValidation(regla);
  }
  Logger.log(JSON.stringify({ hoja: sheet.getName(), fila: filaNueva, n: ultimoN + 1,
    texto: NUEVA_PRIORIDAD.texto, opciones: etiquetas }));
}
function quienEnvia() { Logger.log(JSON.stringify({ efectivo: Session.getEffectiveUser().getEmail() })); }

var ACCESS_TOKEN = "FnFqqvwQQLcBSjOjhiCxEJZdmcpJXuGu";

/* Cuántas prioridades por trimestre admite un tablero. */
var MAX_PRIORIDADES = 6;

/* La web tiene sus propios nombres de estado y la hoja los suyos. Sin esta
   traducción el desplegable de ESTADO rechaza el valor y la tarea se guarda
   sin estado, con el texto cayendo en OBSERVACIONES. */
var ESTADO_A_SHEET = {
  "Por hacer": "Pendiente",
  "En proceso": "En proceso",
  "Seguimiento": "Seguimiento",
  "En testeo": "En testeo",
  "Completado": "Finalizada",
  "Bloqueado": "Bloqueado"
};

/* Id del Google Sheet "Evelyna I CONSULTORÍA: Dra. Daniela Gulli". */
var SHEET_ID = "1fz9FefdaKjy3MRH371d7hjIdxiFryatsAIqwgAaVj4o";
var SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(SHEET_ID)); }

var PREFIJO_ETAPA1 = "01";
var SHEET_OVERRIDES = "WebApp - Overrides";
var SHEET_NOTES = "WebApp - Notas";

/* Los tres tableros del cliente. Las hojas se buscan por su número al
   principio del nombre, así un cambio de título no rompe nada. */
var TABLEROS = [
  { id: "admin", label: "Administración", prefijo: "02" },
  { id: "medicos", label: "Médicos", prefijo: "03" },
  { id: "comercial", label: "Comercial", prefijo: "04" }
];

function tablero_(id) {
  for (var i = 0; i < TABLEROS.length; i++) if (TABLEROS[i].id === id) return TABLEROS[i];
  throw new Error("Tablero desconocido: " + id);
}

/* Hoja cuyo nombre empieza con ese número (ej. "02"). */
function hojaPorPrefijo_(prefijo) {
  var hojas = ss_().getSheets();
  for (var i = 0; i < hojas.length; i++) {
    var nombre = String(hojas[i].getName() || "").trim();
    if (nombre.indexOf(prefijo) === 0) return hojas[i];
  }
  throw new Error("No encuentro la hoja que empieza con '" + prefijo + "'. Hay: " +
    hojas.map(function (h) { return h.getName(); }).join(" | "));
}

var SHEET_SCHEMAS = {};
SHEET_SCHEMAS[SHEET_OVERRIDES] = ["task_id", "estado", "link", "inicio", "cierre", "hidden", "updated_at"];
SHEET_SCHEMAS[SHEET_NOTES] = ["id", "text", "author", "createdAt"];

/* =====================================================================
   ENTRY POINTS
   ===================================================================== */

function doGet(e) {
  try {
    checkToken_(e.parameter.token);
    if (e.parameter.action !== "read") {
      return jsonOut_({ ok: false, error: "acción GET no soportada: " + e.parameter.action });
    }
    ensureSheets_();
    return jsonOut_({
      ok: true,
      seed: readSeed_(),
      overrides: readOverrides_(),
      notes: readSimpleRows_(SHEET_NOTES, SHEET_SCHEMAS[SHEET_NOTES])
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var body = JSON.parse(e.postData.contents || "{}");
    checkToken_(body.token);
    ensureSheets_();
    return jsonOut_({ ok: true, result: handleAction_(body.action, body.payload || {}) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function handleAction_(action, p) {
  switch (action) {
    case "updateTask": return updateTask_(p.tablero, p.id, p.fields || {});
    case "addTask": return addTask_(p.tablero, p.fields || {});
    case "deleteTask": return deleteTask_(p.tablero, p.id);
    case "setOverride": return setOverride_(p.taskId, p.patch || {});
    case "addNote": return addRow_(SHEET_NOTES, SHEET_SCHEMAS[SHEET_NOTES],
      Object.assign({ id: "n" + Date.now(), createdAt: nowIso_() }, p));
    case "deleteNote": return deleteRow_(SHEET_NOTES, p.id);
    case "addMeetingSheet": return addMeetingSheet_(p);
    case "deleteMeetingSheet": return deleteMeetingSheet_(p);
    case "migrarPrioridad": return migrarPrioridad_(p.tablero);
    default: throw new Error("acción desconocida: " + action);
  }
}

function checkToken_(token) {
  if (!token || token !== ACCESS_TOKEN) throw new Error("token inválido");
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function nowIso_() { return new Date().toISOString(); }

/* =====================================================================
   UTILIDADES DE CELDAS
   ===================================================================== */

function cell_(row, idx) {
  if (idx === undefined || idx === null) return null;
  var v = row[idx];
  if (v === null || v === undefined) return null;
  if (typeof v === "string") { v = v.trim(); return v === "" ? null : v; }
  return v;
}

function norm_(s) {
  return String(s === null || s === undefined ? "" : s)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ").trim().toUpperCase();
}

/* Fecha de la hoja -> "yyyy-mm-dd". Texto no reconocido se deja igual. */
function toIsoDate_(value) {
  if (value === null || value === undefined || value === "") return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "America/Argentina/Buenos_Aires", "yyyy-MM-dd");
  }
  var s = String(value).trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var d = m[1], mo = m[2], y = m[3];
    if (y.length === 2) y = "20" + y;
    return y + "-" + ("0" + mo).slice(-2) + "-" + ("0" + d).slice(-2);
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

/* "yyyy-mm-dd" -> fecha a mediodía, para que no se corra de día por zona. */
function toSheetDate_(iso) {
  if (!iso) return "";
  var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return iso;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

/* Opciones de un desplegable que NO acepta otros valores (null si acepta). */
function listOptions_(cell) {
  var dv = cell.getDataValidation();
  if (!dv || dv.getAllowInvalid()) return null;
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return crit[0];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return crit[0].getValues().map(function (r) { return r[0]; }).filter(function (v) { return v !== ""; });
  }
  return null;
}

/* Valores de un desplegable (acepte o no otros), para ofrecerlos en la web. */
function dropdownValues_(cell) {
  var dv = cell.getDataValidation();
  if (!dv) return [];
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  var vals = [];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) vals = crit[0];
  else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    vals = crit[0].getValues().map(function (r) { return r[0]; });
  }
  return vals.map(function (v) { return String(v).trim(); }).filter(Boolean);
}

/* Escribe el primer candidato que la celda acepte. false si lo rechaza. */
function setSafe_(cell, candidates) {
  var options = listOptions_(cell);
  var value = candidates[0];
  if (options) {
    var hit = null;
    candidates.forEach(function (c) {
      if (hit !== null) return;
      options.forEach(function (o) { if (hit === null && norm_(o) === norm_(c)) hit = o; });
    });
    if (hit === null) return false;
    value = hit;
  }
  try { cell.setValue(value); return true; } catch (err) { return false; }
}

/* =====================================================================
   LECTURA DE LAS HOJAS DE PROCESO (02 / 03 / 04)
   ===================================================================== */

/* Ubica las columnas de la tabla de tareas por el texto del encabezado.
   Cada hoja tiene las suyas: 02 y 04 usan DEADLINE, 03 usa FIN; TEMA y
   FINALIZADO pueden no estar. Índices 0-based. */
var NOMBRES_COL = {
  "PRIORIDAD": "prioridad", "PRIO": "prioridad", "P": "prioridad",
  "TEMA": "tema", "TAREA": "tarea", "RESPONSABLE": "responsable",
  "INICIO": "inicio", "DEADLINE": "cierre", "FIN": "cierre", "CIERRE": "cierre",
  "FINALIZADO": "finalizado", "ESTADO": "estado", "OBSERVACIONES": "obs"
};

function colsProceso_(values) {
  for (var r = 0; r < values.length; r++) {
    var fila = values[r], map = null;
    for (var c = 0; c < fila.length; c++) {
      var texto = norm_(fila[c]);
      // "TAREA" puede venir con un cartel al lado ("TAREA 📄 Clientes I Funnel").
      var k = NOMBRES_COL[texto] || (texto.indexOf("TAREA") === 0 ? "tarea" : null);
      if (!k) continue;
      map = map || {};
      if (map[k] === undefined) map[k] = c;
    }
    if (map && map.tarea !== undefined && map.estado !== undefined) {
      map.header = r;
      return map;
    }
  }
  throw new Error("No encuentro el encabezado (TAREA / ESTADO) de la tabla de tareas");
}

function colsEscritura_(C) {
  var out = {};
  ["prioridad", "tema", "tarea", "responsable", "inicio", "cierre", "finalizado", "estado", "obs"].forEach(function (k) {
    if (C[k] !== undefined) out[k] = C[k] + 1;
  });
  return out;
}

function findLabelRow_(values, texto, desde) {
  var target = norm_(texto);
  for (var i = desde || 0; i < values.length; i++) {
    for (var c = 0; c < Math.min(values[i].length, 4); c++) {
      if (norm_(values[i][c]) === target) return i;
    }
  }
  return -1;
}

function tieneTarea_(row, C) {
  return !!(cell_(row, C.tarea) || (C.tema !== undefined && cell_(row, C.tema)));
}

function leerTarea_(row, C, richObs) {
  function v(k) { return C[k] === undefined ? null : cell_(row, C[k]); }
  var obsText = v("obs");
  var links = obsLinks_(richObs, obsText);
  return {
    link: links.length ? links.join("\n") : null,
    prioridad: v("prioridad"), tema: v("tema"), tarea: v("tarea"), responsable: v("responsable"),
    inicio: toIsoDate_(v("inicio")), cierre: toIsoDate_(v("cierre")), finalizado: toIsoDate_(v("finalizado")),
    estado: v("estado"), obs: obsSinLinks_(obsText, links)
  };
}

/* Lee un tablero completo: objetivo, prioridades, tareas y finalizadas. */
function readTablero_(t) {
  var sheet = hojaPorPrefijo_(t.prefijo);
  var values = sheet.getDataRange().getValues();
  var C = colsProceso_(values);
  var richObs = C.obs === undefined ? null : sheet.getRange(1, C.obs + 1, values.length, 1).getRichTextValues();

  var objetivoRow = findLabelRow_(values, "OBJETIVO 1");
  var prioridadesRow = findLabelRow_(values, "PRIORIDADES");
  var iniciativasRow = findLabelRow_(values, "INICIATIVAS");
  var finalizadasRow = findLabelRow_(values, "FINALIZADAS");
  if (finalizadasRow === -1) finalizadasRow = findLabelRow_(values, "FINALIZADOS");

  var objetivo = "";
  var desdeObj = objetivoRow !== -1 ? objetivoRow + 1 : 0;
  var hastaObj = prioridadesRow !== -1 ? prioridadesRow : Math.min(desdeObj + 6, values.length);
  for (var o = desdeObj; o < hastaObj; o++) {
    var texto = primerTexto_(values[o]);
    if (texto && texto.indexOf("¿") !== 0 && !/^(OBJETIVO|PRIORIDADES|INICIATIVAS)/i.test(texto)) {
      objetivo = texto; break;
    }
  }

  // Prioridades: filas numeradas 1..3. Si el rótulo PRIORIDADES no está
  // (la hoja se reordena seguido), se buscan igual arriba de la tabla.
  var prioridades = [];
  var desdePrio = prioridadesRow !== -1 ? prioridadesRow + 1 : 0;
  {
    var finPrio = iniciativasRow !== -1 ? iniciativasRow : C.header;
    for (var pr = desdePrio; pr < finPrio; pr++) {
      var n = null, texto2 = null;
      for (var c2 = 0; c2 < Math.min(values[pr].length, 6); c2++) {
        var val = cell_(values[pr], c2);
        if (val === null) continue;
        var num = typeof val === "number" ? val : Number(String(val).replace(/[.)]/g, "").trim());
        if (n === null && !isNaN(num) && num >= 1 && num <= 9 && String(val).length <= 3) { n = num; continue; }
        if (n !== null && texto2 === null) { texto2 = String(val).trim(); break; }
      }
      if (n === null || !texto2 || texto2.length < 8) continue;
      if (/^KRs|^%|M[ée]tricas/i.test(texto2)) continue; // bloque de métricas: no es prioridad
      if (prioridades.some(function (x) { return Number(x.n) === Number(n); })) continue;
      prioridades.push({ n: n, tt: texto2, desc: "" });
      if (prioridades.length === MAX_PRIORIDADES) break;
    }
    prioridades.sort(function (a, b) { return Number(a.n) - Number(b.n); });
  }

  var iniciativas = [], finalizados = [];
  var desde = (iniciativasRow === -1 ? C.header : Math.max(iniciativasRow, C.header)) + 1;
  var hasta = finalizadasRow === -1 ? values.length : finalizadasRow;
  for (var ir = desde; ir < hasta; ir++) {
    if (!tieneTarea_(values[ir], C)) continue;
    iniciativas.push(leerTarea_(values[ir], C, richObs ? richObs[ir][0] : null));
  }
  if (finalizadasRow !== -1) {
    for (var fr = finalizadasRow + 1; fr < values.length; fr++) {
      if (!tieneTarea_(values[fr], C)) continue;
      finalizados.push(leerTarea_(values[fr], C, richObs ? richObs[fr][0] : null));
    }
  }

  return {
    id: t.id, label: t.label, hoja: sheet.getName(), objetivo: objetivo, prioridades: prioridades,
    iniciativas: iniciativas, finalizados: finalizados, opciones: opcionesTablero_(sheet, C, iniciativas.length ? desde : null)
  };
}

function primerTexto_(fila) {
  for (var c = 0; c < Math.min(fila.length, 6); c++) {
    var v = cell_(fila, c);
    if (typeof v === "string" && v.length > 3) return v;
  }
  return null;
}

/* Opciones de los desplegables de la hoja, para que la web ofrezca lo mismo. */
function opcionesTablero_(sheet, C, filaEjemplo) {
  var out = { responsable: [], tema: [], estado: [] };
  if (filaEjemplo === null || filaEjemplo === undefined) return out;
  try {
    ["responsable", "tema", "estado"].forEach(function (k) {
      if (C[k] !== undefined) out[k] = dropdownValues_(sheet.getRange(filaEjemplo + 1, C[k] + 1));
    });
  } catch (err) {}
  return out;
}

function readSeed_() {
  return {
    tableros: TABLEROS.map(readTablero_),
    etapa1: readEtapa1_(),
    opcionesReuniones: opcionesReuniones_()
  };
}

/* =====================================================================
   LECTURA DE "01 I Plan de trabajo" (reuniones)
   ===================================================================== */

function etapa1Table_() {
  var sheet = hojaPorPrefijo_(PREFIJO_ETAPA1);
  var values = sheet.getDataRange().getValues();
  var header = -1;
  for (var i = 0; i < values.length; i++) {
    if (norm_(values[i][0]) === "FECHA") { header = i; break; }
  }
  if (header === -1) throw new Error("No encuentro la tabla de reuniones (encabezado 'Fecha') en '" + sheet.getName() + "'");
  var last = header;
  while (last + 1 < values.length && cell_(values[last + 1], 0)) last++;
  return { sheet: sheet, values: values, header: header, last: last };
}

function readEtapa1_() {
  var T = etapa1Table_();
  var out = [];
  for (var r = T.header + 1; r <= T.last; r++) {
    var row = T.values[r];
    out.push({
      fecha: toIsoDate_(cell_(row, 0)), hs: cell_(row, 1), tarea: cell_(row, 2),
      responsable: cell_(row, 3), estado: cell_(row, 4), resultado: cell_(row, 5), obs: cell_(row, 6)
    });
  }
  return out;
}

function opcionesReuniones_() {
  try {
    var T = etapa1Table_();
    return {
      responsable: dropdownValues_(T.sheet.getRange(T.header + 2, 4)),
      estado: dropdownValues_(T.sheet.getRange(T.header + 2, 5))
    };
  } catch (err) { return { responsable: [], estado: [] }; }
}

/* =====================================================================
   LINKS EN OBSERVACIONES
   ===================================================================== */

var MAX_LINKS = 3;

function obsLinks_(richValue, text) {
  var out = [];
  function add(u) { if (u && out.indexOf(u) === -1 && out.length < MAX_LINKS) out.push(u); }
  if (richValue) {
    add(richValue.getLinkUrl());
    richValue.getRuns().forEach(function (run) { add(run.getLinkUrl()); });
  }
  (String(text || "").match(/https?:\/\/\S+/g) || []).forEach(add);
  return out;
}

function obsSinLinks_(text, links) {
  if (!text) return null;
  var lines = String(text).split(/\n/).filter(function (line) {
    var l = line.trim();
    return l && links.indexOf(l) === -1 && !/^https?:\/\/\S+$/.test(l);
  });
  var t = lines.join("\n").trim();
  return t || null;
}

/* Deja en OBSERVACIONES el texto y, debajo, un link por línea (clickeable). */
function escribirObsYLinks_(cell, fields, perdidos) {
  var currentText = String(cell.getValue() || "");
  var currentLinks = obsLinks_(cell.getRichTextValue(), currentText);
  var partes = fields.link !== undefined
    ? String(fields.link || "").split(/\s*\n\s*|\s+\|\s+/).map(function (x) { return x.trim(); }).filter(Boolean)
    : currentLinks;
  var links = [], conTexto = [];
  partes.forEach(function (x) {
    if (/^https?:\/\/\S+$/.test(x)) { if (links.length < MAX_LINKS) links.push(x); }
    else conTexto.push(x);
  });
  var base = fields.obs !== undefined ? (fields.obs || "") : (obsSinLinks_(currentText, currentLinks) || "");
  base = [base].concat(conTexto).concat(perdidos || []).filter(Boolean).join("\n");
  var text = [base].concat(links).filter(Boolean).join("\n");
  if (!text) { cell.setValue(""); return; }
  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  var re = /https?:\/\/[^\s]+/g, m;
  while ((m = re.exec(text)) !== null) {
    try { builder = builder.setLinkUrl(m.index, m.index + m[0].length, m[0]); } catch (err) {}
  }
  cell.setRichTextValue(builder.build());
}

/* =====================================================================
   ESCRITURA EN LAS HOJAS DE PROCESO
   ===================================================================== */

function hashId_(str) {
  var h = 5381;
  for (var i = 0; i < str.length; i++) h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
  return "t" + h.toString(36);
}

/* Ubica cada tarea de un tablero (id, fila, sección), igual que la web. */
function procesoLayout_(tableroId) {
  var t = tablero_(tableroId);
  var sheet = hojaPorPrefijo_(t.prefijo);
  var values = sheet.getDataRange().getValues();
  var C = colsProceso_(values);
  var iniciativasRow = findLabelRow_(values, "INICIATIVAS");
  var finalizadasRow = findLabelRow_(values, "FINALIZADAS");
  if (finalizadasRow === -1) finalizadasRow = findLabelRow_(values, "FINALIZADOS");

  var tasks = [], seen = {};
  function scan(desde, hasta, source) {
    var last = desde - 1;
    for (var r = desde; r < hasta; r++) {
      if (!tieneTarea_(values[r], C)) continue;
      var key = [t.id, source, cell_(values[r], C.tema) || "", cell_(values[r], C.tarea) || ""].join("|");
      seen[key] = (seen[key] || 0) + 1;
      tasks.push({ id: hashId_(key + "#" + seen[key]), row: r + 1, source: source });
      last = r;
    }
    return last + 1; // 1-based
  }
  var desde = (iniciativasRow === -1 ? C.header : Math.max(iniciativasRow, C.header)) + 1;
  var hasta = finalizadasRow === -1 ? values.length : finalizadasRow;
  var lastIniciativaRow = scan(desde, hasta, "iniciativa");
  var lastFinalizadoRow = finalizadasRow === -1 ? values.length : scan(finalizadasRow + 1, values.length, "finalizado");

  return {
    tablero: t, sheet: sheet, values: values, C: C, cols: colsEscritura_(C), tasks: tasks,
    lastIniciativaRow: lastIniciativaRow, lastFinalizadoRow: lastFinalizadoRow
  };
}

function findTask_(L, id) {
  for (var i = 0; i < L.tasks.length; i++) if (L.tasks[i].id === id) return L.tasks[i];
  return null;
}

function idAtRow_(tableroId, row) {
  var L = procesoLayout_(tableroId);
  for (var i = 0; i < L.tasks.length; i++) if (L.tasks[i].row === row) return L.tasks[i].id;
  return null;
}

/* Escribe los campos en su columna. Lo que un desplegable rechaza no tira
   abajo el guardado: queda anotado en OBSERVACIONES y se avisa a la web. */
function writeTaskCells_(L, row, fields) {
  var perdidos = [];
  Object.keys(L.cols).forEach(function (k) {
    if (k === "obs" || fields[k] === undefined) return;
    var cell = L.sheet.getRange(row, L.cols[k]);
    var v = fields[k];
    if (v === null || v === "") { try { cell.setValue(""); } catch (err) {} return; }
    var candidatos = (k === "inicio" || k === "cierre" || k === "finalizado") ? [toSheetDate_(v)]
      : k === "estado" ? [ESTADO_A_SHEET[v] || v, v]
      : [v];
    if (!setSafe_(cell, candidatos)) {
      perdidos.push(k.charAt(0).toUpperCase() + k.slice(1) + ": " + v);
    }
  });
  if (L.cols.obs !== undefined && (fields.obs !== undefined || fields.link !== undefined || perdidos.length)) {
    escribirObsYLinks_(L.sheet.getRange(row, L.cols.obs), fields, perdidos);
  }
  return perdidos;
}

/* Mueve una fila (valores y formato) debajo de afterRow. Devuelve su fila. */
function moveRow_(sheet, fromRow, afterRow) {
  sheet.insertRowAfter(afterRow);
  var to = afterRow + 1;
  if (fromRow > afterRow) fromRow++;
  var width = sheet.getMaxColumns();
  sheet.getRange(fromRow, 1, 1, width).copyTo(sheet.getRange(to, 1, 1, width));
  sheet.deleteRow(fromRow);
  return fromRow < to ? to - 1 : to;
}

function updateTask_(tableroId, id, fields) {
  var L = procesoLayout_(tableroId);
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en la hoja. Puede haber cambiado: recargá la página.");
  var perdidos = writeTaskCells_(L, t.row, fields);
  var row = t.row;
  if (fields.estado !== undefined) {
    var done = esFinalizada_(fields.estado);
    if (done && t.source === "iniciativa") row = moveRow_(L.sheet, row, L.lastFinalizadoRow);
    else if (!done && t.source === "finalizado") row = moveRow_(L.sheet, row, L.lastIniciativaRow);
  }
  SpreadsheetApp.flush();
  var newId = idAtRow_(tableroId, row);
  migrateOverride_(id, newId);
  return { id: newId, enObservaciones: perdidos };
}

function addTask_(tableroId, fields) {
  var L = procesoLayout_(tableroId);
  var after = esFinalizada_(fields.estado) ? L.lastFinalizadoRow : L.lastIniciativaRow;
  L.sheet.insertRowAfter(after);
  var row = after + 1;
  L.sheet.getRange(row, 1, 1, L.sheet.getMaxColumns()).clearContent();
  var perdidos = writeTaskCells_(L, row, fields);
  SpreadsheetApp.flush();
  return { id: idAtRow_(tableroId, row), enObservaciones: perdidos };
}

function deleteTask_(tableroId, id) {
  var L = procesoLayout_(tableroId);
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en la hoja. Recargá la página.");
  L.sheet.deleteRow(t.row);
  deleteRow_(SHEET_OVERRIDES, id);
  return { deleted: id };
}

/* El estado que la hoja usa para lo terminado. */
function esFinalizada_(estado) {
  return /^final/i.test(String(estado || ""));
}

/* =====================================================================
   REUNIONES (hoja 01)
   ===================================================================== */

function horas_(duracion) {
  var s = String(duracion || "").trim().toLowerCase().replace("hs", "").replace("h", "").trim();
  var n = Number(s.replace(",", "."));
  return isNaN(n) ? (duracion || "") : n;
}

function addMeetingSheet_(p) {
  var T = etapa1Table_();
  var sheet = T.sheet;
  var src = T.last + 1;
  sheet.insertRowAfter(src);
  var row = src + 1;
  sheet.getRange(src, 1, 1, 7).copyTo(sheet.getRange(row, 1, 1, 7));

  var campos = [
    { col: 1, label: "Fecha", value: toSheetDate_(p.fecha) },
    { col: 2, label: "Hs", value: horas_(p.duracion) },
    { col: 3, label: "Tarea", value: p.tarea || "Reunión de trabajo" },
    { col: 4, label: "Responsable", value: p.responsable || "" },
    { col: 5, label: "Estado", value: p.estado || "Finalizado" },
    { col: 6, label: "Resultado", value: p.resumen || "" }
  ];
  var perdidos = [];
  campos.forEach(function (w) {
    if (w.value === "") { sheet.getRange(row, w.col).setValue(""); return; }
    if (!setSafe_(sheet.getRange(row, w.col), [w.value])) perdidos.push(w.label + ": " + w.value);
  });
  sheet.getRange(row, 7).setValue(perdidos.join(" | "));
  SpreadsheetApp.flush();
  return { row: row, guardado: sheet.getRange(row, 1, 1, 7).getDisplayValues()[0], enObservaciones: perdidos };
}

function deleteMeetingSheet_(p) {
  var T = etapa1Table_();
  for (var r = T.header + 1; r <= T.last; r++) {
    var row = T.values[r];
    if (toIsoDate_(row[0]) === p.fecha &&
        String(cell_(row, 2) || "") === String(p.tarea || "") &&
        String(cell_(row, 5) || "") === String(p.resultado || "")) {
      T.sheet.deleteRow(r + 1);
      return { deleted: true };
    }
  }
  throw new Error("No encuentro esa reunión en la hoja de reuniones. Recargá la página.");
}

/* =====================================================================
   PESTAÑAS DE LA WEB (WebApp - *)
   ===================================================================== */

function ensureSheets_() {
  var ss = ss_();
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      var sheet = ss.insertSheet(name);
      sheet.appendRow(SHEET_SCHEMAS[name]);
      sheet.setFrozenRows(1);
    }
  });
}

function readSimpleRows_(sheetName, cols) {
  var sheet = ss_().getSheetByName(sheetName);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!cell_(values[r], 0)) continue;
    var obj = {};
    cols.forEach(function (c, i) { obj[c] = cell_(values[r], i); });
    out.push(obj);
  }
  return out;
}

function readOverrides_() {
  var rows = readSimpleRows_(SHEET_OVERRIDES, SHEET_SCHEMAS[SHEET_OVERRIDES]);
  var out = {};
  rows.forEach(function (r) {
    out[r.task_id] = {
      estado: r.estado || null, link: r.link || null,
      inicio: toIsoDate_(r.inicio), cierre: toIsoDate_(r.cierre),
      hidden: r.hidden === true || String(r.hidden).toUpperCase() === "TRUE",
      updatedAt: r.updated_at || null
    };
  });
  return out;
}

function findRowIndexById_(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) if (String(values[r][0]) === String(id)) return r + 1;
  return -1;
}

function addRow_(sheetName, cols, obj) {
  var sheet = ss_().getSheetByName(sheetName);
  sheet.appendRow(cols.map(function (c) { return obj[c] === undefined || obj[c] === null ? "" : obj[c]; }));
  return obj;
}

function deleteRow_(sheetName, id) {
  var sheet = ss_().getSheetByName(sheetName);
  if (!sheet) return { deleted: false };
  var rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx === -1) return { deleted: false };
  sheet.deleteRow(rowIdx);
  return { deleted: true };
}

function setOverride_(taskId, patch) {
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var rowIdx = findRowIndexById_(sheet, taskId);
  var updatedAt = nowIso_();
  if (rowIdx === -1) {
    var row = { task_id: taskId, estado: "", link: "", inicio: "", cierre: "", hidden: false, updated_at: updatedAt };
    Object.keys(patch).forEach(function (k) { row[k] = patch[k] === null ? "" : patch[k]; });
    addRow_(SHEET_OVERRIDES, SHEET_SCHEMAS[SHEET_OVERRIDES], row);
  } else {
    var cols = SHEET_SCHEMAS[SHEET_OVERRIDES];
    Object.keys(patch).forEach(function (key) {
      var colIdx = cols.indexOf(key);
      if (colIdx === -1) return;
      sheet.getRange(rowIdx, colIdx + 1).setValue(patch[key] === null ? "" : patch[key]);
    });
    sheet.getRange(rowIdx, cols.indexOf("updated_at") + 1).setValue(updatedAt);
  }
  return { taskId: taskId, patch: patch };
}

/* Al editar una tarea cambia su id (es un hash del contenido): se muda la
   fila de overrides al id nuevo para no perder lo oculto. */
function migrateOverride_(oldId, newId) {
  if (!newId || oldId === newId) return;
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var rowIdx = findRowIndexById_(sheet, oldId);
  if (rowIdx === -1) return;
  sheet.getRange(rowIdx, 1).setValue(newId);
  sheet.getRange(rowIdx, SHEET_SCHEMAS[SHEET_OVERRIDES].indexOf("updated_at") + 1).setValue(nowIso_());
}

/* =====================================================================
   MIGRACIÓN OPCIONAL: columna PRIORIDAD en un tablero
   ===================================================================== */

var PRIORIDADES_SHEET = ["P1", "P2", "P3"];

/* Agrega la columna PRIORIDAD (antes de TAREA) con desplegable P1/P2/P3.
   No completa nada: las tareas quedan sin prioridad hasta que se cargue. */
function migrarPrioridad_(tableroId) {
  var L = procesoLayout_(tableroId);
  if (L.C.prioridad !== undefined) return { columnaCreada: false, tablero: tableroId };
  var col = (L.C.tema !== undefined ? L.C.tema : L.C.tarea) + 1; // 1-based
  L.sheet.insertColumnBefore(col);
  var filaHeader = L.C.header + 1;
  L.sheet.getRange(filaHeader, col + 1).copyTo(L.sheet.getRange(filaHeader, col), { formatOnly: true });
  L.sheet.getRange(filaHeader, col).setValue("PRIORIDAD");
  L.sheet.setColumnWidth(col, 92);
  SpreadsheetApp.flush();

  var L2 = procesoLayout_(tableroId);
  var regla = SpreadsheetApp.newDataValidation().requireValueInList(PRIORIDADES_SHEET, true).setAllowInvalid(true).build();
  L2.tasks.forEach(function (t) { L2.sheet.getRange(t.row, L2.cols.prioridad).setDataValidation(regla); });
  return { columnaCreada: true, tablero: tableroId, tareas: L2.tasks.length };
}

/* Diagnóstico rápido desde el editor. */
function diagnostico_() {
  return TABLEROS.map(function (t) {
    try {
      var d = readTablero_(t);
      return { tablero: t.id, hoja: d.hoja, objetivo: d.objetivo, prioridades: d.prioridades.length,
        tareas: d.iniciativas.length, finalizadas: d.finalizados.length, opciones: d.opciones };
    } catch (err) {
      return { tablero: t.id, prefijo: t.prefijo, error: String(err && err.message || err) };
    }
  });
}
