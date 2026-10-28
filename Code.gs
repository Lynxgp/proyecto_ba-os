/*
 * Registro de salidas
 * Copyright (C) 2026 Daniel Montiel <dmonrui264@g.educaand.es> · IES Campanillas (Málaga)
 *
 * Este programa es software libre: puedes redistribuirlo y/o modificarlo bajo los términos de la
 * Licencia Pública General GNU publicada por la Free Software Foundation, ya sea la versión 3 de la
 * Licencia o (a tu elección) cualquier versión posterior.
 *
 * Se distribuye con la esperanza de que sea útil, pero SIN NINGUNA GARANTÍA; incluso sin la garantía
 * implícita de COMERCIABILIDAD o IDONEIDAD PARA UN PROPÓSITO PARTICULAR. Consulta la Licencia Pública
 * General GNU para más detalles: <https://www.gnu.org/licenses/gpl-3.0.html>
 * (texto completo en el archivo LICENSE).
 *
 * SPDX-License-Identifier: GPL-3.0-or-later
 */
// Registro de salidas · servidor (Google Apps Script)
// Pestañas: Alumnos | Profesores | Configuración | Registro | Fuera | Incidencias | Resumen general | Resumen por curso

const REG = 'Registro', CFG = 'Configuración', INC = 'Incidencias', FUE = 'Fuera', OBS = 'Observaciones';
const GENERICOS_ = ['guardia', 'conserjeria']; // usuarios compartidos: no pueden cambiar su contraseña
const FUERA_MIN = 15; // minutos máximos fuera: pasado ese tiempo se marca la vuelta sola

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Registro de salidas')
    .addItem('Actualizar resúmenes', 'actualizarResumenes')
    .addItem('Cargar datos de ejemplo (borra el registro)', 'generarDatosEjemplo').addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Registro de salidas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Ejecutar UNA vez desde el editor: crea las pestañas que falten
function preparar() {
  const ss = SpreadsheetApp.getActive();
  const crear = (nombre, filas) => {
    if (ss.getSheetByName(nombre)) return;
    const sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
    sh.getRange(1, 1, 1, filas[0].length).setFontWeight('bold');
    sh.setFrozenRows(1);
  };
  crear('Alumnos', [['Curso', 'Alumno (Apellidos, Nombre)', 'Código de baño'], ['1º ESO A', 'Apellidos, Nombre', '1234']]);
  const sa = ss.getSheetByName('Alumnos'); // hojas anteriores: añade la columna del código si falta
  if (!String(sa.getRange(1, 3).getValue()).trim()) sa.getRange(1, 3).setValue('Código de baño').setFontWeight('bold');
  sa.getRange('C:C').setNumberFormat('@'); // los códigos se guardan como texto (no pierden ceros a la izquierda)
  crear('Profesores', [['Profesor', 'Usuario', 'Contraseña'], ['Apellidos, Nombre', 'usuario1', 'cambiar-clave']]);
  crear(CFG, [['Parámetro', 'Valor'], ['Máximo semanal', 4],
    ['Inicio 1.º trimestre', new Date(2026, 8, 15)], ['Fin 1.º trimestre', new Date(2026, 11, 22)],
    ['Inicio 2.º trimestre', new Date(2027, 0, 8)], ['Fin 2.º trimestre', new Date(2027, 2, 26)],
    ['Inicio 3.º trimestre', new Date(2027, 3, 5)], ['Fin 3.º trimestre', new Date(2027, 5, 22)]]);
  crear(REG, [['ID', 'Salida', 'Alumno', 'Curso', 'Profesor', 'Vuelta', 'Minutos', 'Salida (ms)']]);
  ss.getSheetByName(CFG).getRange('B3:B8').setNumberFormat('dd/mm/yyyy');
  ss.getSheetByName(REG).getRangeList(['B:B', 'F:F']).setNumberFormat('dd/mm/yyyy hh:mm');
  hojaFuera_();
  hojaInc_();
  hojaObs_();
  crearAdmin();
  actualizarResumenes();
}

// Comprueba qué separador de argumentos admite setFormulas en esta hoja (según su idioma): "," o ";"
function separador_(sh) {
  const z = sh.getRange('J1');
  z.setFormula('=MAX(1,2)');
  SpreadsheetApp.flush();
  const ok = z.getValue() === 2;
  z.clear();
  return ok ? ',' : ';';
}

// Crea o rehace las hojas de resumen (con fórmulas vivas sobre "Registro")
function actualizarResumenes() {
  const ss = SpreadsheetApp.getActive();
  const hoja = n => { const s = ss.getSheetByName(n) || ss.insertSheet(n); s.clear(); return s; };

  // Resumen general
  const g = hoja('Resumen general');
  const SEP = separador_(g), F = f => SEP === ';' ? f.replace(/,/g, ';') : f;
  g.getRange('A1:F1').setValues([['Periodo', 'Desde', 'Hasta', 'Salidas', 'Alumnos distintos', 'Minutos de media']]).setFontWeight('bold');
  const P = [
    ['Semana actual', '=TODAY()-WEEKDAY(TODAY(),2)+1', '=B2+6'],
    ['Mes actual', '=EOMONTH(TODAY(),-1)+1', '=EOMONTH(TODAY(),0)'],
    ['1.º trimestre', `='${CFG}'!B3`, `='${CFG}'!B4`],
    ['2.º trimestre', `='${CFG}'!B5`, `='${CFG}'!B6`],
    ['3.º trimestre', `='${CFG}'!B7`, `='${CFG}'!B8`],
    ['Total', '=DATE(2000,1,1)', '=DATE(2100,1,1)']];
  P.forEach((p, i) => {
    const r = i + 2, c = `Registro!$B:$B,">="&B${r},Registro!$B:$B,"<"&C${r}+1`;
    g.getRange(r, 1).setValue(p[0]); // las etiquetas son texto, no fórmulas
    g.getRange(r, 2, 1, 5).setFormulas([[F(p[1]), F(p[2]),
      F(`=COUNTIFS(${c})`),
      F(`=IFERROR(COUNTUNIQUE(FILTER(Registro!$C:$C,Registro!$B:$B>=B${r},Registro!$B:$B<C${r}+1)),0)`),
      F(`=IFERROR(ROUND(AVERAGEIFS(Registro!$G:$G,${c}),1),0)`)]]);
  });
  g.getRange('B2:C7').setNumberFormat('dd/mm/yyyy');
  g.getRange('A9').setValue('Para ver otra semana o mes, cambia las fechas Desde/Hasta. Todo el historial está en la hoja Registro.');

  // Resumen por curso
  const cursos = [...new Set(ss.getSheetByName('Alumnos').getDataRange().getValues().slice(1)
    .map(r => String(r[0]).trim()).filter(String))];
  const c = hoja('Resumen por curso');
  c.getRange('A1:G1').setValues([['Curso', 'Semana actual', 'Mes actual', '1.º trimestre', '2.º trimestre', '3.º trimestre', 'Total']]).setFontWeight('bold');
  cursos.forEach((cu, i) => {
    const r = i + 2, f = [];
    for (let j = 0; j < 6; j++) {
      const pr = j + 2;
      f.push(F(`=COUNTIFS(Registro!$D:$D,$A${r},Registro!$B:$B,">="&'Resumen general'!$B$${pr},Registro!$B:$B,"<"&'Resumen general'!$C$${pr}+1)`));
    }
    c.getRange(r, 1).setValue(cu);
    c.getRange(r, 2, 1, 6).setFormulas([f]);
  });
  const t = cursos.length + 2;
  c.getRange(t, 1).setValue('Todos los cursos').setFontWeight('bold');
  c.getRange(t, 2, 1, 6).setFormulas([[...'BCDEFG'.split('').map(l => F(`=SUM(${l}2:${l}${t - 1})`))]]).setFontWeight('bold');
  c.setFrozenRows(1);
}

// --- Acceso -------------------------------------------------------------

function entrar(usuario, clave) {
  const cache = CacheService.getScriptCache(), u = String(usuario).trim().toLowerCase();
  const fk = 'f_' + u, fallos = +cache.get(fk) || 0;
  if (!u || fallos >= 5) return null; // 5 fallos = bloqueo de 10 minutos
  const fila = SpreadsheetApp.getActive().getSheetByName('Profesores').getDataRange().getValues().slice(1)
    .find(r => String(r[1]).trim().toLowerCase() === u && String(r[2]) === String(clave));
  if (!fila) { cache.put(fk, fallos + 1, 600); return null; }
  const token = Utilities.getUuid();
  cache.put('t_' + token, String(fila[0]).trim(), 21600); // sesión de 6 horas
  cache.put('u_' + token, u, 21600);
  return token;
}

// Lista para el desplegable "Profesorado" del acceso: solo nombre y usuario, nunca la contraseña.
// Guardia y conserjería tienen su propio botón, así que no salen aquí.
function listaLogin() {
  return enviar_(() => {
    const norm = s => String(s).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    return SpreadsheetApp.getActive().getSheetByName('Profesores').getDataRange().getValues().slice(1)
      .filter(r => String(r[1]).trim() && !['guardia', 'conserjeria', 'admin'].includes(norm(r[1])))
      .map(r => ({ n: String(r[0]).trim() || String(r[1]).trim(), u: String(r[1]).trim() }))
      .sort((a, b) => a.n.localeCompare(b.n, 'es'));
  });
}

// --- Alumnado: el alumno apunta su salida con su código de baño (sin sesión y sin acceso a la app) ---
// El código está en la columna C de la hoja "Alumnos". Si el alumno ya figura fuera (lo registró un profesor),
// no se duplica: solo se le muestra la información. Si no figuraba, se anota con profesor "(alumnado)".
const ALU_ = '(alumnado)';

// Código comparable: sin espacios, en minúsculas y, si son solo cifras, sin ceros a la izquierda
function codNorm_(s) {
  const t = String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, '');
  return /^\d+$/.test(t) ? t.replace(/^0+(?=\d)/, '') : t;
}

function inicioDia_(tz, t) {
  return Utilities.parseDate(Utilities.formatDate(new Date(t), tz, 'yyyy-MM-dd') + ' 00:00:00', tz, 'yyyy-MM-dd HH:mm:ss').getTime();
}

function registrarAlumno(codigo, cid) { return enviar_(() => registrarAlumno_(codigo, cid)); }

function registrarAlumno_(codigo, cid) {
  const cod = codNorm_(codigo);
  if (!cod || cod.length > 40) return { ok: false, motivo: 'vacio' };
  // Freno a quien prueba códigos al azar: 8 fallos desde el mismo dispositivo = 5 minutos de espera
  const cache = CacheService.getScriptCache(), fk = 'fa_' + String(cid || '').replace(/[^\w]/g, '').slice(0, 24);
  const fallos = +cache.get(fk) || 0;
  if (fallos >= (cid ? 8 : 30)) return { ok: false, motivo: 'bloqueo' };
  const ss = SpreadsheetApp.getActive();
  const hit = ss.getSheetByName('Alumnos').getDataRange().getValues().slice(1)
    .filter(r => String(r[1]).trim() && codNorm_(r[2]) === cod);
  if (!hit.length) { cache.put(fk, fallos + 1, 300); return { ok: false, motivo: 'noexiste' }; }
  if (hit.length > 1) return { ok: false, motivo: 'repetido' };
  const c = String(hit[0][0]).trim(), n = String(hit[0][1]).trim();
  return conCerrojo_(() => {
    const reg = ss.getSheetByName(REG), fs = hojaFuera_();
    const ya = estadoFuera_().find(x => x.n === n); // ¿ya figura fuera ahora mismo?
    let nuevo = false, prof = '';
    if (ya) prof = ya.p === ALU_ ? '' : ya.p;
    else {
      const ahora = new Date(), id = 'a' + ahora.getTime() + Math.random().toString(36).slice(2, 5);
      if (!reg.getRange(1, 8).getValue()) reg.getRange(1, 8).setValue('Salida (ms)').setFontWeight('bold');
      reg.appendRow([id, ahora, n, c, ALU_, '', '', ahora.getTime()]);
      fs.appendRow([id, ahora, n, c, ALU_, ahora.getTime()]);
      nuevo = true;
    }
    const tz = ss.getSpreadsheetTimeZone(), t0 = inicioDia_(tz, Date.now());
    const u = +Utilities.formatDate(new Date(), tz, 'u'); // 1 = lunes … 7 = domingo
    const s0 = inicioDia_(tz, t0 + 12 * 36e5 - (u - 1) * 864e5);
    let hoy = 0, sem = 0;
    reg.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0] || String(r[2]).trim() !== n) return;
      const o = ms_(r, 7, 1);
      if (o >= s0) sem++;
      if (o >= t0) hoy++;
    });
    return { ok: true, n, c, hoy, sem, prof, nuevo };
  });
}

const sinTildes_ = s => String(s).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
function usuarioDe_(token) { return token ? (CacheService.getScriptCache().get('u_' + token) || '') : ''; }
function esGenerico_(token) { return GENERICOS_.includes(sinTildes_(usuarioDe_(token))); }
function esAdmin_(token) { return sinTildes_(usuarioDe_(token)) === 'admin'; }

// Cambio de contraseña del propio usuario (no disponible para guardia ni conserjería)
// Devuelve 'ok' | 'actual' | 'corta' | 'igual' | 'bloqueo' | 'generico'
function cambiarClave(token, actual, nueva) {
  const me = quien_(token);
  if (!me) throw new Error('Sesión caducada');
  if (esGenerico_(token)) return 'generico';
  nueva = String(nueva == null ? '' : nueva);
  if (nueva.length < 6 || nueva.length > 64) return 'corta';
  const cache = CacheService.getScriptCache(), u = usuarioDe_(token);
  const fk = 'fc_' + (u || me).toLowerCase(), fallos = +cache.get(fk) || 0;
  if (fallos >= 5) return 'bloqueo'; // 5 fallos = espera de 10 minutos
  return conCerrojo_(() => {
    const sh = SpreadsheetApp.getActive().getSheetByName('Profesores'), v = sh.getDataRange().getValues();
    let i = -1;
    for (let k = 1; k < v.length && i < 0; k++)
      if (u ? String(v[k][1]).trim().toLowerCase() === u : String(v[k][0]).trim() === me) i = k;
    if (i < 0) throw new Error('Usuario no encontrado');
    if (GENERICOS_.includes(sinTildes_(v[i][1]))) return 'generico';
    if (String(v[i][2]) !== String(actual)) { cache.put(fk, fallos + 1, 600); return 'actual'; }
    if (String(actual) === nueva) return 'igual';
    sh.getRange(i + 1, 3).setNumberFormat('@').setValue(nueva);
    cache.remove(fk);
    return 'ok';
  });
}

function quien_(token) { return token ? CacheService.getScriptCache().get('t_' + token) : null; }

// --- Datos --------------------------------------------------------------

// Las respuestas viajan como texto JSON: así nunca se pierden por valores no serializables (fechas, etc.)
// y, si algo falla en el servidor, el navegador recibe el motivo en lugar de un null mudo.
function enviar_(fn) {
  try { return JSON.stringify(fn()); }
  catch (e) { return JSON.stringify({ error: String(e && e.message || e) }); }
}
function getInicial(token) { return enviar_(() => getInicial_(token)); }

function getInicial_(token) {
  const me = quien_(token);
  if (!me) return null;
  const ss = SpreadsheetApp.getActive(), courses = {}, codes = {}; // codes: nombre → código de baño (solo viaja a personal con sesión)
  ss.getSheetByName('Alumnos').getDataRange().getValues().slice(1).forEach(r => {
    const c = String(r[0]).trim();
    if (c && r[1]) {
      (courses[c] = courses[c] || []).push(String(r[1]).trim());
      const k = String(r[2] == null ? '' : r[2]).trim();
      if (k) codes[String(r[1]).trim()] = k;
    }
  });
  const max = Number(ss.getSheetByName(CFG).getRange('B2').getValue()) || 4;
  const desde = Date.now() - 35 * 864e5;
  const fuera = conCerrojo_(estadoFuera_); // cierra las caducadas y devuelve quién está fuera ahora
  const vals = ss.getSheetByName(REG).getDataRange().getValues();
  const log = vals.slice(1)
    .filter(r => r[0] && ms_(r, 7, 1) >= desde)
    .map(r => ({ id: String(r[0]), out: ms_(r, 7, 1), n: String(r[2]), c: String(r[3]), p: String(r[4]),
                 back: r[5] instanceof Date ? r[5].getTime() : null }));
  const si = ss.getSheetByName(INC);
  const inc = si ? si.getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1] instanceof Date && r[1].getTime() >= desde)
    .map(r => ({ id: String(r[0]), t: r[1].getTime(), p: String(r[2]),
                 s: String(r[3]).split(' | ').filter(String), note: String(r[4] || '') })) : [];
  const so = ss.getSheetByName(OBS);
  const obs = so ? so.getDataRange().getValues().slice(1)
    .filter(r => r[0] && String(r[2]).trim())
    .map(r => ({ id: String(r[0]), t: r[1] instanceof Date ? r[1].getTime() : 0, n: String(r[2]).trim(),
                 c: String(r[3]), p: String(r[4]), x: String(r[5] || '') })) : [];
  const cv = ss.getSheetByName(CFG).getRange('B3:B8').getValues().map(r => r[0]);
  const tri = [0, 1, 2].map(i => cv[2 * i] instanceof Date && cv[2 * i + 1] instanceof Date ? [cv[2 * i].getTime(), cv[2 * i + 1].getTime()] : null);
  return { courses, codes, max, me, log, inc, obs, fuera, tri, ahora: Date.now(), cp: !esGenerico_(token), adm: esAdmin_(token) };
}

// Quién está fuera ahora mismo (hoja "Fuera"): lectura ligera, la consultan todos los navegadores
function getFuera(token) { return enviar_(() => getFuera_(token)); }

function getFuera_(token) {
  if (!quien_(token)) return null;
  const sh = SpreadsheetApp.getActive().getSheetByName(FUE);
  if (sh) { // lectura rápida sin cerrojo si no hay nadie caducado
    const lim = FUERA_MIN * 60000, ahora = Date.now();
    const v = sh.getDataRange().getValues().slice(1).filter(r => r[0] && ms_(r, 5, 1) > 0);
    if (!v.some(r => ahora - ms_(r, 5, 1) >= lim)) return { ahora: Date.now(), lista: v.map(aFuera_) };
  }
  return { ahora: Date.now(), lista: conCerrojo_(estadoFuera_) };
}

// Devuelve 'ok' o 'ya_fuera' (otro profesor ya lo había registrado como fuera)
function registrarSalida(token, rec) {
  const me = quien_(token);
  if (!me) throw new Error('Sesión caducada');
  return conCerrojo_(() => {
    const reg = SpreadsheetApp.getActive().getSheetByName(REG), fs = hojaFuera_(), nom = String(rec.n).trim();
    if (buscar_(reg, String(rec.id))) return 'ok'; // ya registrada (reintento)
    const hit = fs.getRange('C:C').createTextFinder(nom).matchEntireCell(true).findNext();
    if (hit && hit.getRow() > 1) return fs.getRange(hit.getRow(), 5).getValue() === ALU_ ? 'ya_alumnado' : 'ya_fuera';
    const ahora = new Date();
    if (!reg.getRange(1, 8).getValue()) reg.getRange(1, 8).setValue('Salida (ms)').setFontWeight('bold');
    reg.appendRow([rec.id, ahora, nom, rec.c, me, '', '', ahora.getTime()]);
    fs.appendRow([rec.id, ahora, nom, rec.c, me, ahora.getTime()]);
    return 'ok';
  });
}

function registrarVuelta(token, id) {
  if (!quien_(token)) throw new Error('Sesión caducada');
  conCerrojo_(() => {
    const ss = SpreadsheetApp.getActive(), reg = ss.getSheetByName(REG), fs = hojaFuera_();
    const f = buscar_(reg, id);
    if (f && !reg.getRange(f, 6).getValue()) {
      const sal = ms_(reg.getRange(f, 1, 1, 8).getValues()[0], 7, 1);
      const vuelta = new Date(Math.min(Date.now(), sal + FUERA_MIN * 60000));
      reg.getRange(f, 6, 1, 2).setValues([[vuelta, Math.round((vuelta.getTime() - sal) / 60000)]]);
    }
    const g = buscar_(fs, id);
    if (g > 1) fs.deleteRow(g);
  });
}

function deshacer(token, id) {
  if (!quien_(token)) throw new Error('Sesión caducada');
  conCerrojo_(() => {
    const reg = SpreadsheetApp.getActive().getSheetByName(REG), fs = hojaFuera_(), f = buscar_(reg, id);
    if (f && !reg.getRange(f, 6).getValue()) reg.deleteRow(f);
    const g = buscar_(fs, id);
    if (g > 1) fs.deleteRow(g);
  });
}

// Incidencias (baño sucio): se guardan en la pestaña "Incidencias" (se crea sola si falta)
function registrarIncidencia(token, rec) {
  const me = quien_(token);
  if (!me) throw new Error('Sesión caducada');
  const ahora = Date.now(), t = Number(rec.t);
  const cuando = new Date(t > 0 ? Math.min(t, ahora) : ahora);
  const alumnos = (Array.isArray(rec.s) ? rec.s : []).slice(0, 30).map(s => String(s).trim()).filter(String);
  conCerrojo_(() => hojaInc_().appendRow([String(rec.id), cuando, me,
    texto_(alumnos.join(' | ')), texto_(String(rec.note || '').slice(0, 300))]));
}

function borrarIncidencia(token, id) {
  if (!quien_(token)) throw new Error('Sesión caducada');
  conCerrojo_(() => {
    const sh = SpreadsheetApp.getActive().getSheetByName(INC), f = sh ? buscar_(sh, String(id)) : 0;
    if (f > 1) sh.deleteRow(f);
  });
}

// Observaciones sobre un alumno (visibles solo para el personal). Se guardan en la pestaña "Observaciones".
function guardarObservacion(token, rec) {
  const me = quien_(token);
  if (!me) throw new Error('Sesión caducada');
  const n = String(rec.n || '').trim().slice(0, 200), x = String(rec.x || '').trim().slice(0, 500);
  if (!n || !x) throw new Error('Faltan datos');
  conCerrojo_(() => hojaObs_().appendRow([String(rec.id), new Date(Math.min(Number(rec.t) || Date.now(), Date.now())),
    texto_(n), texto_(String(rec.c || '').trim().slice(0, 100)), me, texto_(x)]));
}

// Solo quien escribió la observación puede eliminarla
function borrarObservacion(token, id) {
  const me = quien_(token);
  if (!me) throw new Error('Sesión caducada');
  conCerrojo_(() => {
    const sh = SpreadsheetApp.getActive().getSheetByName(OBS), f = sh ? buscar_(sh, String(id)) : 0;
    if (f < 2) return;
    if (!esAdmin_(token) && String(sh.getRange(f, 5).getValue()).trim() !== me) throw new Error('Solo puedes eliminar tus observaciones');
    sh.deleteRow(f);
  });
}

function hojaObs_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(OBS);
  if (!sh) {
    sh = ss.insertSheet(OBS);
    sh.getRange(1, 1, 1, 6).setValues([['ID', 'Fecha', 'Alumno', 'Curso', 'Autor', 'Observación']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
  }
  return sh;
}

// Elimina un registro del historial, solo si lo introdujo quien lo pide. Devuelve 'ok' | 'ajeno' | 'no_existe'
function eliminarRegistro(token, id) {
  const me = quien_(token);
  if (!me) throw new Error('Sesión caducada');
  return conCerrojo_(() => {
    const ss = SpreadsheetApp.getActive(), reg = ss.getSheetByName(REG), fs = hojaFuera_(), f = buscar_(reg, String(id));
    if (f < 2) return 'no_existe';
    if (!esAdmin_(token) && String(reg.getRange(f, 5).getValue()).trim() !== me) return 'ajeno'; // el admin puede borrar cualquiera
    reg.deleteRow(f);
    const g = buscar_(fs, String(id));
    if (g > 1) fs.deleteRow(g);
    return 'ok';
  });
}

function hojaInc_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(INC);
  if (!sh) {
    sh = ss.insertSheet(INC);
    sh.getRange(1, 1, 1, 5).setValues([['ID', 'Fecha', 'Profesor', 'Alumnos señalados', 'Nota']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
  }
  return sh;
}

// Evita que un texto que empiece por = + - @ se interprete como fórmula
function texto_(s) { return /^[=+\-@]/.test(s) ? ' ' + s : s; }

// Hoja "Fuera": solo quienes están fuera en este momento. Si no existe se crea con lo que ya estuviera abierto en Registro.
// Llamar siempre dentro del cerrojo.
function hojaFuera_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FUE);
  if (!sh) {
    sh = ss.insertSheet(FUE);
    sh.getRange(1, 1, 1, 6).setValues([['ID', 'Salida', 'Alumno', 'Curso', 'Profesor', 'Salida (ms)']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm:ss');
    const abiertas = ss.getSheetByName(REG).getDataRange().getValues().slice(1)
      .filter(r => r[0] && ms_(r, 7, 1) > 0 && !(r[5] instanceof Date))
      .map(r => [r[0], r[1], r[2], r[3], r[4], ms_(r, 7, 1)]);
    if (abiertas.length) sh.getRange(2, 1, abiertas.length, 6).setValues(abiertas);
  } else if (!sh.getRange(1, 6).getValue()) {
    sh.getRange(1, 6).setValue('Salida (ms)').setFontWeight('bold'); // hojas creadas con la versión anterior
  }
  return sh;
}

// Hora de salida en milisegundos: usa la columna numérica (no depende de zonas horarias) y, si falta, la fecha de la hoja
const ms_ = (r, iMs, iFecha) => {
  const n = Number(r[iMs]);
  return n > 0 ? n : (r[iFecha] instanceof Date ? r[iFecha].getTime() : 0);
};
const aFuera_ = r => ({ id: String(r[0]), out: ms_(r, 5, 1), n: String(r[2]), c: String(r[3]), p: String(r[4]) });

// Cierra (vuelta automática) a quien lleva FUERA_MIN minutos o más y devuelve quién sigue fuera. Dentro del cerrojo.
function estadoFuera_() {
  const reg = SpreadsheetApp.getActive().getSheetByName(REG), fs = hojaFuera_();
  const fv = fs.getDataRange().getValues(), lim = FUERA_MIN * 60000, ahora = Date.now(), vivos = [], venc = [];
  for (let i = 1; i < fv.length; i++) {
    const r = fv[i];
    if (!r[0] || !(ms_(r, 5, 1) > 0)) continue;
    (ahora - ms_(r, 5, 1) >= lim ? venc : vivos).push(i);
  }
  venc.forEach(i => {
    const f = buscar_(reg, String(fv[i][0]));
    if (f && !reg.getRange(f, 6).getValue())
      reg.getRange(f, 6, 1, 2).setValues([[new Date(ms_(fv[i], 5, 1) + lim), FUERA_MIN]]);
  });
  venc.slice().reverse().forEach(i => fs.deleteRow(i + 1)); // de abajo arriba
  return vivos.map(i => aFuera_(fv[i]));
}

function buscar_(sh, id) {
  const f = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

function conCerrojo_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

// Historial completo de "Registro", filtrado por fechas. desde/hasta en ms (hasta excluido); null = sin límite.
function getHistorial(token, desde, hasta) { return enviar_(() => getHistorial_(token, desde, hasta)); }

function getHistorial_(token, desde, hasta) {
  if (!quien_(token)) return null;
  const d = Number(desde) || 0, h = Number(hasta) || Infinity, CAP = 3000, ahora = Date.now();
  const todas = SpreadsheetApp.getActive().getSheetByName(REG).getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => {
      const out = ms_(r, 7, 1), mm = r[6] !== '' && r[6] !== null && !isNaN(Number(r[6])) ? Number(r[6]) : null;
      return { id: String(r[0]), out, n: String(r[2]), c: String(r[3]), p: String(r[4]),
               m: mm !== null ? mm : (ahora - out >= FUERA_MIN * 60000 ? FUERA_MIN : null) };
    })
    .filter(x => x.out >= d && x.out < h)
    .sort((a, b) => b.out - a.out);
  return { total: todas.length, rows: todas.slice(0, CAP) };
}

// Resumen agregado de un periodo cualquiera (día, semana o trimestre). desde/hasta en ms (hasta excluido).
function getResumen(token, desde, hasta) { return enviar_(() => getResumen_(token, desde, hasta)); }

function getResumen_(token, desde, hasta) {
  if (!quien_(token)) return null;
  const ss = SpreadsheetApp.getActive(), tz = ss.getSpreadsheetTimeZone();
  const d = Number(desde) || 0, h = Number(hasta) || Infinity, ahora = Date.now(), lim = FUERA_MIN * 60000;
  const stu = {}, cur = {}, hor = new Array(24).fill(0), dia = {}, memo = {};
  let tot = 0, sumM = 0, nM = 0;
  ss.getSheetByName(REG).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    const out = ms_(r, 7, 1);
    if (!(out >= d && out < h)) return;
    tot++;
    const n = String(r[2]), c = String(r[3]);
    (stu[n] || (stu[n] = [n, c, 0]))[2]++;
    const hk = Math.floor(out / 3600000); // una sola conversión de zona horaria por hora
    const f = memo[hk] || (memo[hk] = Utilities.formatDate(new Date(out), tz, 'yyyy-MM-dd-H').split('-'));
    hor[+f[3]]++;
    const dk = f[0] + '-' + f[1] + '-' + f[2];
    dia[dk] = (dia[dk] || 0) + 1;
    (cur[c] || (cur[c] = [0, 0]))[0]++;
    const g = r[6], m = g !== '' && g !== null && !isNaN(Number(g)) ? Number(g) : (ahora - out >= lim ? FUERA_MIN : null);
    if (m !== null) { sumM += m; nM++; }
  });
  const lista = Object.values(stu);
  lista.forEach(k => { cur[k[1]][1]++; });
  return { tot, dist: lista.length, min: nM ? Math.round(sumM / nM * 10) / 10 : 0, cur, stu: lista, hor, dia };
}

// --- Datos de ejemplo ----------------------------------------------------

const FESTIVOS_ = ['2026-10-12', '2026-12-07', '2026-12-08'];

// Borra Registro, Fuera e Incidencias y genera un historial simulado desde el 15/09/2026 hasta hoy
// con el alumnado y el profesorado reales de la hoja.
function generarDatosEjemplo() {
  const ss = SpreadsheetApp.getActive();
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('Cargar datos de ejemplo',
      'Se BORRARÁN todas las salidas e incidencias actuales y se generará un historial simulado desde el 15/09/2026 hasta hoy. ¿Continuar?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  } catch (e) { /* ejecutado desde el editor sin hoja abierta: se continúa */ }

  const tz = ss.getSpreadsheetTimeZone(), p2 = n => String(n).padStart(2, '0');
  const alumnos = ss.getSheetByName('Alumnos').getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1]).map(r => [String(r[0]).trim(), String(r[1]).trim()]);
  const profes = ss.getSheetByName('Profesores').getDataRange().getValues().slice(1)
    .filter(r => sinTildes_(r[1]) !== 'admin').map(r => String(r[0]).trim()).filter(String);
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-').map(Number);
  const dias = [];
  for (let t = Date.UTC(2026, 8, 15); t <= Date.UTC(hoy[0], hoy[1] - 1, hoy[2]); t += 864e5) {
    const dt = new Date(t), dow = dt.getUTCDay(), k = dt.toISOString().slice(0, 10);
    if (dow > 0 && dow < 6 && !FESTIVOS_.includes(k) && !(k >= '2026-12-23' && k <= '2027-01-07'))
      dias.push([dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()]);
  }
  const medianoche = (y, m, d) =>
    Utilities.parseDate(`${y}-${p2(m + 1)}-${p2(d)} 00:00:00`, tz, 'yyyy-MM-dd HH:mm:ss').getTime();
  const sim = simularSalidas_(alumnos, profes, dias, medianoche, Date.now());

  conCerrojo_(() => {
    const reg = ss.getSheetByName(REG), fs = hojaFuera_(), si = hojaInc_();
    [reg, fs, si].forEach(sh => { if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent(); });
    reg.getRange(1, 8).setValue('Salida (ms)').setFontWeight('bold');
    if (sim.salidas.length) {
      reg.getRange(2, 1, sim.salidas.length, 8).setValues(sim.salidas.map(s =>
        [s.id, new Date(s.out), s.n, s.c, s.p, new Date(s.back), s.min, s.out]));
    }
    if (sim.inc.length) {
      si.getRange(2, 1, sim.inc.length, 5).setValues(sim.inc.map(i =>
        [i.id, new Date(i.t), i.p, texto_(i.s.join(' | ')), i.note]));
    }
    reg.getRangeList(['B:B', 'F:F']).setNumberFormat('dd/mm/yyyy hh:mm');
  });
  actualizarResumenes();
  SpreadsheetApp.flush();
  try { SpreadsheetApp.getUi().alert('Listo', `${sim.salidas.length} salidas y ${sim.inc.length} incidencias generadas en ${dias.length} días lectivos.`, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
}

function mulberry_(a) {
  return function () {
    a |= 0; a = a + 0x6D2B79F5 | 0;
    let t = Math.imul(a ^ a >>> 15, 1 | a);
    t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  };
}

// Simulación (función pura): alumnos = [[curso, nombre]], dias = [[año, mes0, día]], medianoche(y,m,d) → ms
function simularSalidas_(alumnos, profes, dias, medianoche, ahora) {
  const r = mulberry_(20260915), gauss = () => Math.sqrt(-2 * Math.log(1 - r())) * Math.cos(2 * Math.PI * r());
  const b36 = n => { let s = ''; for (let i = 0; i < n; i++) s += 'abcdefghijklmnopqrstuvwxyz0123456789'[Math.floor(r() * 36)]; return s; };
  const cursos = [...new Set(alumnos.map(a => a[0]))], fc = {};
  cursos.forEach(c => fc[c] = 0.7 + 0.6 * r());
  const w = alumnos.map(a => Math.exp(1.1 * gauss()) * fc[a[0]]);
  for (let i = 0; i < 8; i++) w[Math.floor(r() * w.length)] *= 6; // alumnado "frecuente"
  const acum = []; w.reduce((s, x, i) => (acum[i] = s + x), 0);
  const total = acum[acum.length - 1];
  const pick = () => { const x = r() * total; let lo = 0, hi = acum.length - 1; while (lo < hi) { const m = (lo + hi) >> 1; if (acum[m] < x) lo = m + 1; else hi = m; } return lo; };
  const espec = profes.filter(p => /^(guardia|conserjería)$/i.test(p)), norm = profes.filter(p => !espec.includes(p));
  const pool = {};
  cursos.forEach(c => pool[c] = Array.from({ length: 5 }, () => norm[Math.floor(r() * norm.length)] || 'Profesor'));
  const prof = c => espec.length && r() < 0.08 ? espec[Math.floor(r() * espec.length)] : pool[c][Math.floor(r() * 5)];
  const clases = [[495, .12], [555, .18], [615, .20], [705, .18], [765, .18], [825, .14]]; // inicio de cada clase (min) y peso
  const salidas = [];
  dias.forEach(([y, m, d]) => {
    const mid = medianoche(y, m, d), dow = new Date(Date.UTC(y, m, d)).getUTCDay(), hoy = {};
    const n = Math.round(85 * (dow === 1 || dow === 5 ? 0.9 : 1.05) * (0.75 + 0.5 * r()));
    for (let i = 0; i < n; i++) {
      let x = r(), cl = clases[clases.length - 1];
      for (const c of clases) { if (x < c[1]) { cl = c; break; } x -= c[1]; }
      const out = mid + ((cl[0] + 8 + Math.floor(r() * 45)) * 60 + Math.floor(r() * 60)) * 1000;
      if (out > ahora - 20 * 60000) continue;
      let k = -1;
      for (let t = 0; t < 12 && k < 0; t++) {
        const c = pick(), prev = hoy[c] || [];
        if (prev.length < 3 && !prev.some(o => Math.abs(o - out) < 20 * 60000)) k = c;
      }
      if (k < 0) continue;
      (hoy[k] = hoy[k] || []).push(out);
      const olvido = r() < 0.07;
      const dur = olvido ? 900 : Math.max(60, Math.min(840, Math.round((2.5 + Math.exp(1 + 0.5 * gauss())) * 60 + (r() - 0.5) * 40)));
      const [cu, nom] = alumnos[k];
      salidas.push({ id: 's' + out + b36(3), out, n: nom, c: cu, p: prof(cu), back: out + dur * 1000, min: Math.round(dur / 60) });
    }
  });
  salidas.sort((a, b) => a.out - b.out);
  const inc = [], notas = ['Baño sucio', 'Baño sucio', 'Sin papel', 'Sin jabón', 'Grifo abierto', ''];
  for (let i = 0; i < 14 && salidas.length; i++) {
    const s = salidas[Math.floor(r() * salidas.length)], t = s.out + (5 + Math.floor(r() * 15)) * 60000;
    if (t > ahora) continue;
    const cerca = salidas.filter(x => Math.abs(x.out - s.out) < 15 * 60000).slice(0, 3).map(x => x.n);
    inc.push({ id: 'i' + t + b36(3), t, p: s.p, s: [...new Set([s.n, ...cerca])].slice(0, 3), note: notas[Math.floor(r() * notas.length)] });
  }
  inc.sort((a, b) => a.t - b.t);
  return { salidas, inc };
}

// --- Administración (solo usuario "admin") ----------------------------------------------------------

const PROT_ = ['guardia', 'conserjeria', 'admin']; // filas de Profesores que la importación nunca toca
const lim_ = (s, n) => String(s == null ? '' : s).trim().slice(0, n);

// Crea el usuario admin si no existe (se ejecuta sola desde preparar(); también se puede lanzar a mano).
// Muestra la contraseña inicial una sola vez: cámbiala al entrar con el botón "Contraseña".
function crearAdmin() {
  const pw = crearAdmin_();
  if (!pw) { try { SpreadsheetApp.getUi().alert('Usuario admin', 'El usuario admin ya existe.', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {} return; }
  Logger.log('Usuario admin creado. Contraseña inicial: ' + pw);
  try { SpreadsheetApp.getUi().alert('Usuario admin creado', 'Usuario: admin\nContraseña inicial: ' + pw + '\n\nCámbiala al entrar (botón "Contraseña").', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
}

function crearAdmin_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Profesores');
  if (sh.getDataRange().getValues().slice(1).some(r => sinTildes_(r[1]) === 'admin')) return null;
  const pw = Utilities.getUuid().replace(/-/g, '').slice(0, 10), r0 = sh.getLastRow() + 1;
  sh.getRange(r0, 3).setNumberFormat('@');
  sh.getRange(r0, 1, 1, 3).setValues([['Administración', 'admin', pw]]);
  return pw;
}

// Código de baño: 1.ª letra del nombre + 3 del 1.er apellido + 3 del 2.º (minúsculas, sin tildes y sin partículas "de, la…")
function codigoBano_(nombre) {
  const t = sinTildes_(nombre), i = t.indexOf(',');
  const ap = (i < 0 ? t : t.slice(0, i)).replace(/\b(de|del|la|las|los|y|van|von|da|das|do|dos|di)\s+/g, '').trim().split(/\s+/);
  const no = i < 0 ? '' : t.slice(i + 1).trim();
  return (no.charAt(0) || '') + (ap[0] || '').slice(0, 3) + (ap[1] || '').slice(0, 3);
}

function claveCurso_(c) {
  const m = /^(\d+)\s*º?\s*(.*)$/.exec(c), resto = m ? m[2] : c;
  return [/^ESO/i.test(resto) ? 0 : /^BACH/i.test(resto) ? 1 : 2, m ? +m[1] : 99, resto];
}
function cmpAlu_(a, b) {
  const x = claveCurso_(a[0]), y = claveCurso_(b[0]);
  for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] < y[i] ? -1 : 1;
  return a[1].localeCompare(b[1], 'es');
}

// Importa las listas de Séneca. tipo 'prof': filas [Empleado/a, Usuario IdEA, DNI/Pasaporte];
// tipo 'alu': filas [Unidad, Alumno/a, Estado Matrícula]. aplicar=false → solo vista previa. borrar → quita lo que ya no figura.
function adminImportar(token, tipo, filas, aplicar, borrar) {
  return enviar_(() => {
    if (!quien_(token)) return null;
    if (!esAdmin_(token)) throw new Error('Solo el administrador puede importar listas');
    if (!Array.isArray(filas) || filas.length > 3000) throw new Error('Fichero no válido');
    if (tipo !== 'prof' && tipo !== 'alu') throw new Error('Tipo de lista desconocido');
    return conCerrojo_(() => tipo === 'prof' ? impProf_(filas, !!aplicar, !!borrar) : impAlu_(filas, !!aplicar, !!borrar));
  });
}

function impProf_(filas, aplicar, borrar) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Profesores'), v = sh.getDataRange().getValues();
  const u2i = new Map();
  for (let i = 1; i < v.length; i++) { const u = String(v[i][1]).trim().toLowerCase(); if (u && !u2i.has(u)) u2i.set(u, i); }
  const vistos = new Set(), nuevos = [], act = [], omit = [], avisos = [];
  let igual = 0, dup = 0;
  filas.forEach(f => {
    const n = lim_(f[0], 150), u = lim_(f[1], 60).toLowerCase(), d = lim_(f[2], 40);
    if (!n || !u) { omit.push(n || u || '(fila vacía)'); return; }
    if (PROT_.includes(sinTildes_(u))) { omit.push(n + ' (usuario reservado)'); return; }
    if (vistos.has(u)) { dup++; return; }
    const i = u2i.get(u);
    if (i === undefined) {
      if (!d) { omit.push(n + ' (sin DNI/Pasaporte)'); return; }
      nuevos.push([n, u, d]);
    } else if (String(v[i][0]).trim() !== n) act.push({ i, n, antes: String(v[i][0]).trim() });
    else igual++;
    vistos.add(u);
  });
  if (!vistos.size) throw new Error('El fichero no contiene profesorado importable');
  const bajas = [];
  if (borrar) for (let i = 1; i < v.length; i++) {
    const u = String(v[i][1]).trim().toLowerCase();
    if (u && !PROT_.includes(sinTildes_(u)) && !vistos.has(u)) bajas.push(i);
  }
  if (dup) avisos.push(dup + ' fila(s) repetida(s) en el fichero (se usa la primera).');
  if (nuevos.length) avisos.push('A las personas nuevas se les asigna su DNI/Pasaporte como contraseña inicial; a quien ya existía no se le cambia la contraseña.');
  if (aplicar) {
    act.forEach(a => sh.getRange(a.i + 1, 1).setValue(a.n));
    bajas.slice().reverse().forEach(i => sh.deleteRow(i + 1));
    if (nuevos.length) {
      const r0 = sh.getLastRow() + 1;
      sh.getRange(r0, 3, nuevos.length, 1).setNumberFormat('@');
      sh.getRange(r0, 1, nuevos.length, 3).setValues(nuevos);
    }
  }
  return { aplicado: aplicar, nuevos: nuevos.length, actualizados: act.length, bajas: bajas.length, sinCambios: igual, omitidos: omit.length,
    ej: { nuevos: nuevos.slice(0, 6).map(x => x[0]), actualizados: act.slice(0, 6).map(a => a.antes + ' → ' + a.n),
          bajas: bajas.slice(0, 8).map(i => String(v[i][0]).trim() || String(v[i][1])), omitidos: omit.slice(0, 6) }, avisos };
}

function impAlu_(filas, aplicar, borrar) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Alumnos'), v = sh.getDataRange().getValues();
  const ex = new Map();
  for (let i = 1; i < v.length; i++) { const n = String(v[i][1]).trim(); if (n && !ex.has(n)) ex.set(n, { c: String(v[i][0]).trim(), k: String(v[i][2] == null ? '' : v[i][2]).trim() }); }
  const avisos = [], lista0 = [], vistos = new Set(), cursosDe = new Map();
  let omEstado = 0, omInv = 0, repet = 0;
  filas.forEach(f => {
    const c = lim_(f[0], 60), n = lim_(f[1], 150), e = lim_(f[2], 40);
    if (e) { omEstado++; return; }                  // matrícula anulada, trasladada… → no se importa
    if (!c || !n) { omInv++; return; }
    if (vistos.has(c + '|' + n)) { repet++; return; } // misma persona dos veces en el mismo curso
    vistos.add(c + '|' + n);
    lista0.push([c, n]);
    if (!cursosDe.has(n)) cursosDe.set(n, new Set());
    cursosDe.get(n).add(c);
  });
  // Mismo nombre en cursos distintos = personas distintas: se añade el curso al nombre para distinguirlas
  const fin = new Map(), homonimos = [];
  lista0.forEach(([c, n]) => {
    const nn = cursosDe.get(n).size > 1 ? n + ' (' + c + ')' : n;
    if (nn !== n) homonimos.push(nn);
    fin.set(nn, c);
  });
  if (!fin.size) throw new Error('El fichero no contiene alumnado importable');
  const nuevos = [], cambios = [], bajas = [];
  let igual = 0;
  fin.forEach((c, n) => { const e = ex.get(n); if (!e) nuevos.push(n); else if (e.c !== c) cambios.push(n + ': ' + e.c + ' → ' + c); else igual++; });
  if (borrar) ex.forEach((e, n) => { if (!fin.has(n)) bajas.push(n); });
  const conFormula = String(sh.getRange(2, 3).getFormula() || '').charAt(0) === '=';
  if (omEstado) avisos.push(omEstado + ' matrícula(s) con Estado relleno (anulada, trasladada…) no se importan.');
  if (omInv) avisos.push(omInv + ' fila(s) sin curso o sin nombre no se importan.');
  if (homonimos.length) avisos.push(homonimos.length + ' alumno(s) comparten nombre con otro de distinto curso: se importan los dos y se les añade el curso al nombre (por ejemplo, «' + homonimos[0] + '»).');
  if (repet) avisos.push(repet + ' fila(s) con el mismo nombre y el mismo curso no se importan (repetidas).');
  avisos.push(conFormula ? 'La columna C (código de baño) tiene una fórmula: no se modifica. Si es una fórmula por fila, arrástrala hasta el final.'
                         : 'A cada alumno nuevo se le genera su código de baño; los que ya existían conservan el suyo.');
  if (aplicar) {
    const lista = [];
    fin.forEach((c, n) => lista.push([c, n]));
    if (!borrar) ex.forEach((e, n) => { if (!fin.has(n)) lista.push([e.c, n]); });
    lista.sort(cmpAlu_);
    const antes = v.length - 1;
    if (antes > 0) sh.getRange(2, 1, antes, 2).clearContent();      // la columna C no se toca al limpiar
    sh.getRange(2, 1, lista.length, 2).setValues(lista);
    if (!conFormula) {
      const usados = new Set(), codes = lista.map(([c, n]) => { const e = ex.get(n), k = e ? e.k : ''; if (k) usados.add(codNorm_(k)); return k; });
      lista.forEach(([c, n], j) => {
        if (codes[j]) return;
        const base = codigoBano_(n) || 'alumno';
        let k = base, s = 0;
        while (usados.has(codNorm_(k))) k = base + (++s);
        usados.add(codNorm_(k)); codes[j] = k;
      });
      sh.getRange(2, 3, lista.length, 1).setNumberFormat('@').setValues(codes.map(k => [k]));
      if (antes > lista.length) sh.getRange(lista.length + 2, 3, antes - lista.length, 1).clearContent();
      if (!String(sh.getRange(1, 3).getValue()).trim()) sh.getRange(1, 3).setValue('Código de baño').setFontWeight('bold');
    }
  }
  return { aplicado: aplicar, nuevos: nuevos.length, actualizados: cambios.length, bajas: bajas.length, sinCambios: igual, omitidos: omEstado + omInv,
    ej: { nuevos: nuevos.slice(0, 6), actualizados: cambios.slice(0, 6), bajas: bajas.slice(0, 8), omitidos: [] }, avisos };
}

// --- Logo del centro (lo cambia el admin; se guarda en la pestaña "Configuración", fila "Logo (URL)") ---

const LOGO_KEY_ = 'Logo (URL)';

function logoCfg_() {
  const v = SpreadsheetApp.getActive().getSheetByName(CFG).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === LOGO_KEY_) return String(v[i][1] || '').trim();
  return '';
}

// Público (se necesita en la pantalla de acceso): devuelve la dirección del logo o '' para usar el original
function getLogo() {
  try {
    const c = CacheService.getScriptCache();
    let u = c.get('logo');
    if (u === null) { u = logoCfg_(); c.put('logo', u || '-', 21600); }
    return u === '-' ? '' : u;
  } catch (e) { return ''; }
}

// Acepta solo https; convierte los enlaces de Google Drive en una dirección de imagen directa
function normalizarLogo_(url) {
  let u = String(url == null ? '' : url).trim();
  if (!u) return '';
  if (u.length > 500 || /[\s"'<>]/.test(u)) throw new Error('La dirección no es válida');
  const d = /^https:\/\/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=\w+&)?id=)([\w-]{10,})/i.exec(u);
  if (d) u = 'https://drive.google.com/thumbnail?id=' + d[1] + '&sz=w600';
  if (!/^https:\/\/[^\/\s]+\.[^\/\s]+/i.test(u)) throw new Error('La dirección debe empezar por https://');
  return u;
}

// Solo admin. aplicar=false → solo normaliza y devuelve la dirección (para probarla antes de guardar)
function adminLogo(token, url, aplicar) {
  return enviar_(() => {
    if (!quien_(token)) return null;
    if (!esAdmin_(token)) throw new Error('Solo el administrador puede cambiar el logo');
    const u = normalizarLogo_(url);
    if (aplicar) conCerrojo_(() => {
      const sh = SpreadsheetApp.getActive().getSheetByName(CFG), v = sh.getDataRange().getValues();
      let f = 0;
      for (let i = 1; i < v.length && !f; i++) if (String(v[i][0]).trim() === LOGO_KEY_) f = i + 1;
      if (!f) f = sh.getLastRow() + 1;
      sh.getRange(f, 1, 1, 2).setValues([[LOGO_KEY_, u]]);
      CacheService.getScriptCache().remove('logo');
    });
    return { url: u };
  });
}
