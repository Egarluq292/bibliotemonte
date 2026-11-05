/**
 * BIBLIOTECA ESCOLAR — Servidor (Google Apps Script)
 * ---------------------------------------------------
 * La hoja de cálculo es la base de datos compartida. Todas las reservas pasan
 * por aquí y se comprueban con un candado (LockService), así que dos maestros
 * nunca pueden reservar la misma hora aunque pulsen a la vez.
 *
 * Instalación:
 *  1. Crea una hoja de Google nueva → Extensiones → Apps Script.
 *  2. Pega este archivo, guarda y ejecuta la función `instalar` una vez
 *     (acepta los permisos). En el registro aparecerá el CÓDIGO DE CENTRO.
 *  3. Implementar → Nueva implementación → Aplicación web:
 *       Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario.
 *     Copia la URL /exec: la usará la web de GitHub Pages.
 *  4. El primer maestro que se registre queda como ADMINISTRADOR (regístrate tú primero).
 *
 * Cómo llama la web (sin cabeceras, para evitar problemas de CORS):
 *   fetch(URL + '?action=estado&mes=2026-11')
 *   fetch(URL, { method: 'POST', body: JSON.stringify({ action: 'reservar', token, fecha, tramo, grupo }) })
 */

const TZ = 'Europe/Madrid';

// Festivos y efemérides que se mantienen en GitHub (además de la pestaña Festivos de la hoja).
// Clases del centro (nombre y número aproximado de alumnos). Cada maestro elige en cuáles da clase.
const CLASES_CENTRO = [
  ['Infantil', 15], ['1.er ciclo', 6], ['2.º ciclo', 12], ['3.er ciclo', 6], ['ESO', 9], ['Aula específica', 5]
];

const URL_CALENDARIO = 'https://egarluq292.github.io/bibliotemonte/calendario.json';

const TRAMOS = [
  { n: 1, inicio: '09:30', fin: '10:30' },
  { n: 2, inicio: '10:30', fin: '11:30' },
  { n: 3, inicio: '12:00', fin: '13:00' },
  { n: 4, inicio: '13:00', fin: '14:00' }
];

const HOJAS = {
  Config:      ['clave', 'valor'],
  Maestros:    ['id', 'nombre', 'nombreClave', 'sal', 'hash', 'rol', 'token', 'alta'],
  Clases:      ['codigo', 'nombre', 'maestroId', 'activa', 'alta', 'alumnos', 'maestros'],
  Reservas:    ['id', 'fecha', 'tramo', 'maestroId', 'maestro', 'grupo', 'creada', 'estado', 'anuladaPor'],
  Sugerencias: ['id', 'creada', 'tipoAutor', 'alias', 'clase', 'titulo', 'texto', 'estado', 'moderadaPor', 'codigoClase', 'maestroClase'],
  Resenas:     ['id', 'creada', 'tipoAutor', 'alias', 'clase', 'libro', 'autorLibro', 'valoracion', 'texto', 'estado', 'moderadaPor', 'codigoClase', 'maestroClase'],
  Festivos:    ['fecha', 'motivo']
};

const ACCIONES_ESCRITURA = [
  'registrarMaestro', 'entrar', 'reservar', 'anular', 'enviarSugerencia', 'enviarResena',
  'moderar', 'crearClase', 'desactivarClase',
  'cambiarClave', 'restablecerClave', 'bajaMaestro', 'cambiarRol', 'cambiarCodigoCentro', 'alumnosClase', 'eliminarClase', 'retirar', 'misClases', 'cambiarCodigoClase'
];

/* ===================== CALENDARIO ESCOLAR ===================== */

/** Festivos de la hoja + los de calendario.json (días sueltos o periodos «desde/hasta»). */
function todosFestivos() {
  const hojaF = filas('Festivos').map(f => ({ fecha: f.fecha, motivo: f.motivo }));
  return hojaF.concat(festivosCalendario());
}

function festivosCalendario() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('calendario');
  if (guardado) return JSON.parse(guardado);
  let lista = [];
  try {
    const r = UrlFetchApp.fetch(URL_CALENDARIO + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      const datos = JSON.parse(r.getContentText());
      (datos.festivos || []).forEach(f => {
        if (f.fecha) lista.push({ fecha: f.fecha, motivo: f.motivo || 'Festivo' });
        else if (f.desde && f.hasta) {
          for (let d = f.desde; d <= f.hasta; d = sumarDias(d, 1)) lista.push({ fecha: d, motivo: f.motivo || 'Vacaciones' });
        }
      });
    }
  } catch (e) { /* Si GitHub no responde, se usan solo los festivos de la hoja. */ }
  cache.put('calendario', JSON.stringify(lista), 1800);
  return lista;
}

/** Ejecútala una vez a mano tras actualizar el servidor: pide el permiso para leer GitHub. */
function probarCalendario() {
  CacheService.getScriptCache().remove('calendario');
  Logger.log('Festivos leídos de GitHub: ' + festivosCalendario().length);
}

/* ===================== INSTALACIÓN ===================== */

function instalar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(HOJAS).forEach(nombre => {
    let h = ss.getSheetByName(nombre);
    if (!h) h = ss.insertSheet(nombre);
    const cab = HOJAS[nombre];
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
    // Todo como texto: evita que Sheets convierta fechas y horas por su cuenta.
    h.getRange(1, 1, h.getMaxRows(), cab.length).setNumberFormat('@');
  });
  const h1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (h1 && ss.getSheets().length > 1) ss.deleteSheet(h1);

  const cfg = leerConfig();
  if (!cfg.CODIGO_CENTRO) guardarConfig('CODIGO_CENTRO', generarCodigo());
  if (!cfg.DIAS_ANTELACION) guardarConfig('DIAS_ANTELACION', '60');
  Logger.log('CÓDIGO DE CENTRO (para que los maestros se registren): ' + leerConfig().CODIGO_CENTRO);
}

/* ===================== ENTRADA ===================== */

function doGet(e) { return responder(() => manejar((e.parameter || {}).action, e.parameter || {})); }

function doPost(e) {
  return responder(() => {
    const p = JSON.parse((e.postData && e.postData.contents) || '{}');
    return manejar(p.action, p);
  });
}

function responder(fn) {
  let salida;
  try { salida = Object.assign({ ok: true }, fn()); }
  catch (err) { salida = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

function manejar(accion, p) {
  const acciones = {
    estado, publicas, registrarMaestro, entrar, reservar, anular,
    enviarSugerencia, enviarResena, pendientes, moderar, listarClases,
    cambiarClave, adminInfo, restablecerClave, bajaMaestro, cambiarRol, cambiarCodigoCentro, alumnosClase, eliminarClase, retirar,
    misClases, cambiarCodigoClase
  };
  if (!acciones[accion]) throw new Error('Acción desconocida.');
  prepararClases();
  if (ACCIONES_ESCRITURA.indexOf(accion) === -1) return acciones[accion](p);

  const candado = LockService.getScriptLock();
  if (!candado.tryLock(15000)) throw new Error('El servidor está ocupado. Inténtalo de nuevo en unos segundos.');
  try { return acciones[accion](p); }
  finally { SpreadsheetApp.flush(); candado.releaseLock(); }
}

/* ===================== CONSULTAS PÚBLICAS ===================== */

/** Reservas de un mes + situación de hoy. ?mes=AAAA-MM (por defecto, el actual). */
function estado(p) {
  const ahora = ahoraLocal();
  const mes = /^\d{4}-\d{2}$/.test(p.mes || '') ? p.mes : ahora.fecha.slice(0, 7);
  const activas = filas('Reservas').filter(r => r.estado === 'activa');
  const limpiar = r => ({ id: r.id, fecha: r.fecha, tramo: Number(r.tramo), maestro: r.maestro, maestroId: r.maestroId, grupo: r.grupo });
  return {
    ahora: { fecha: ahora.fecha, hora: ahora.hora, tramoActual: tramoEn(ahora.hora) },
    tramos: TRAMOS,
    mes,
    reservas: activas.filter(r => r.fecha.slice(0, 7) === mes).map(limpiar),
    hoy: activas.filter(r => r.fecha === ahora.fecha).map(limpiar),
    festivos: todosFestivos(),
    diasAntelacion: Number(leerConfig().DIAS_ANTELACION || 60)
  };
}

/** Sugerencias y reseñas ya aprobadas (lo que ve todo el mundo). */
function publicas() {
  const aprobadas = (hoja) => filas(hoja).filter(x => x.estado === 'aprobada').reverse().slice(0, 150);
  // Número de alumnos por nombre de clase (para el ranking). No se envían los códigos.
  const tam = {};
  filas('Clases').filter(c => c.activa === 'sí' && Number(c.alumnos) > 0).forEach(c => {
    tam[c.nombre] = Math.max(tam[c.nombre] || 0, Number(c.alumnos));
  });
  return {
    clases: Object.keys(tam).map(n => ({ nombre: n, alumnos: tam[n] })),
    sugerencias: aprobadas('Sugerencias').map(s => ({ id: s.id, creada: s.creada, tipoAutor: s.tipoAutor, alias: s.alias, clase: s.clase, titulo: s.titulo, texto: s.texto })),
    resenas: aprobadas('Resenas').map(r => ({ id: r.id, creada: r.creada, tipoAutor: r.tipoAutor, alias: r.alias, clase: r.clase, libro: r.libro, autorLibro: r.autorLibro, valoracion: Number(r.valoracion), texto: r.texto }))
  };
}

/* ===================== MAESTROS ===================== */

/** { codigoCentro, nombre, pin } → token. El primero en registrarse es administrador. */
function registrarMaestro(p) {
  if (String(p.codigoCentro || '').trim().toUpperCase() !== String(leerConfig().CODIGO_CENTRO).toUpperCase())
    throw new Error('El código de centro no es correcto.');
  const nombre = texto(p.nombre, 40, 'nombre');
  const pin = String(p.pin || '');
  if (pin.length < 4) throw new Error('La clave debe tener al menos 4 caracteres.');
  const clave = normalizar(nombre);
  const maestros = filas('Maestros');
  if (maestros.some(m => m.nombreClave === clave)) throw new Error('Ya hay un maestro registrado con ese nombre. Usa «Entrar».');

  const sal = Utilities.getUuid();
  const m = {
    id: 'M' + Date.now().toString(36), nombre, nombreClave: clave, sal, hash: resumen(sal, pin),
    rol: maestros.length === 0 ? 'admin' : 'maestro', token: Utilities.getUuid(), alta: ahoraLocal().completo
  };
  anadir('Maestros', m);
  return sesion(m);
}

/** { nombre, pin } → token. Bloquea 15 min tras 5 fallos. */
function entrar(p) {
  const clave = normalizar(texto(p.nombre, 40, 'nombre'));
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('fallos_' + clave) || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos. Espera 15 minutos.');
  const m = filas('Maestros').find(x => x.nombreClave === clave);
  if (!m || m.hash !== resumen(m.sal, String(p.pin || ''))) {
    cache.put('fallos_' + clave, String(fallos + 1), 900);
    throw new Error('Nombre o clave incorrectos.');
  }
  cache.remove('fallos_' + clave);
  return sesion(m);
}

function sesion(m) { return { token: m.token, maestro: { id: m.id, nombre: m.nombre, rol: m.rol } }; }

function maestroPorToken(token) {
  const m = token ? filas('Maestros').find(x => x.token === token) : null;
  if (!m) throw new Error('Sesión no válida. Vuelve a entrar.');
  return m;
}

/** { token, actual, nueva } — cualquier maestro cambia su propia clave. */
function cambiarClave(p) {
  const m = maestroPorToken(p.token);
  if (m.hash !== resumen(m.sal, String(p.actual || ''))) throw new Error('La clave actual no es correcta.');
  const nueva = String(p.nueva || '');
  if (nueva.length < 4) throw new Error('La clave nueva debe tener al menos 4 caracteres.');
  const sal = Utilities.getUuid();
  escribir('Maestros', m._fila, { sal, hash: resumen(sal, nueva) });
  return { cambiada: true };
}

/* ===================== ADMINISTRACIÓN ===================== */

function soloAdmin(token) {
  const m = maestroPorToken(token);
  if (m.rol !== 'admin') throw new Error('Solo el administrador puede hacer esto.');
  return m;
}

/** { token } → código de centro y lista de maestros. */
function adminInfo(p) {
  soloAdmin(p.token);
  return {
    codigoCentro: leerConfig().CODIGO_CENTRO,
    maestros: filas('Maestros').map(m => ({ id: m.id, nombre: m.nombre, rol: m.rol, alta: m.alta }))
  };
}

/** { token, id } → clave provisional. Cierra las sesiones abiertas de ese maestro. */
function restablecerClave(p) {
  soloAdmin(p.token);
  const m = filas('Maestros').find(x => x.id === p.id);
  if (!m) throw new Error('Ese maestro no existe.');
  const provisional = String(100000 + Math.floor(Math.random() * 900000));
  const sal = Utilities.getUuid();
  escribir('Maestros', m._fila, { sal, hash: resumen(sal, provisional), token: Utilities.getUuid() });
  return { nombre: m.nombre, provisional };
}

/** { token, id } — no borra sus reservas ya hechas. */
function bajaMaestro(p) {
  const admin = soloAdmin(p.token);
  if (p.id === admin.id) throw new Error('No puedes darte de baja a ti mismo.');
  const m = filas('Maestros').find(x => x.id === p.id);
  if (!m) throw new Error('Ese maestro no existe.');
  hoja('Maestros').deleteRow(m._fila);
  return { baja: m.nombre };
}

/** { token, id, rol:'admin'|'maestro' } */
function cambiarRol(p) {
  const admin = soloAdmin(p.token);
  if (['admin', 'maestro'].indexOf(p.rol) === -1) throw new Error('Rol no válido.');
  if (p.id === admin.id && p.rol !== 'admin') throw new Error('No puedes quitarte a ti mismo el papel de administrador.');
  const m = filas('Maestros').find(x => x.id === p.id);
  if (!m) throw new Error('Ese maestro no existe.');
  escribir('Maestros', m._fila, { rol: p.rol });
  return { id: m.id, rol: p.rol };
}

/** { token } → código de centro nuevo. Los maestros ya registrados no se ven afectados. */
function cambiarCodigoCentro(p) {
  soloAdmin(p.token);
  const codigo = generarCodigo();
  guardarConfig('CODIGO_CENTRO', codigo);
  return { codigoCentro: codigo };
}

/* ===================== RESERVAS ===================== */

/** { token, fecha:'AAAA-MM-DD', tramo:1-4, grupo } */
function reservar(p) {
  const m = maestroPorToken(p.token);
  const fecha = validarFecha(p.fecha);
  const tramo = TRAMOS.find(t => t.n === Number(p.tramo));
  if (!tramo) throw new Error('Tramo horario no válido.');
  const grupo = texto(p.grupo, 40, 'grupo o motivo');

  const ahora = ahoraLocal();
  if (fecha + ' ' + tramo.fin <= ahora.completo) throw new Error('Esa hora ya ha pasado.');
  const dias = Number(leerConfig().DIAS_ANTELACION || 60);
  if (fecha > sumarDias(ahora.fecha, dias)) throw new Error('Solo se puede reservar con ' + dias + ' días de antelación como máximo.');

  // Comprobación dentro del candado: aquí es donde se impiden los solapamientos.
  const ocupada = filas('Reservas').find(r => r.estado === 'activa' && r.fecha === fecha && Number(r.tramo) === tramo.n);
  if (ocupada) throw new Error('Esa hora acaba de reservarla ' + ocupada.maestro + '.');

  const r = {
    id: 'R' + Date.now().toString(36), fecha, tramo: String(tramo.n), maestroId: m.id, maestro: m.nombre,
    grupo, creada: ahora.completo, estado: 'activa', anuladaPor: ''
  };
  anadir('Reservas', r);
  return { reserva: { id: r.id, fecha, tramo: tramo.n, maestro: m.nombre, maestroId: m.id, grupo } };
}

/** { token, id } — solo quien reservó o el administrador. */
function anular(p) {
  const m = maestroPorToken(p.token);
  const r = filas('Reservas').find(x => x.id === p.id && x.estado === 'activa');
  if (!r) throw new Error('La reserva no existe o ya estaba anulada.');
  if (r.maestroId !== m.id && m.rol !== 'admin') throw new Error('Solo puedes anular tus propias reservas.');
  const tramo = TRAMOS.find(t => t.n === Number(r.tramo));
  if (r.fecha + ' ' + tramo.fin <= ahoraLocal().completo) throw new Error('No se puede anular una reserva pasada.');
  escribir('Reservas', r._fila, { estado: 'anulada', anuladaPor: m.nombre });
  return { anulada: r.id };
}

/* ===================== SUGERENCIAS Y RESEÑAS ===================== */

/** Alumnado: { codigoClase, alias, ... } · Maestros: { token, ... } */
function enviarSugerencia(p) {
  const a = autor(p);
  const s = {
    id: 'S' + Date.now().toString(36), creada: ahoraLocal().completo, tipoAutor: a.tipo, alias: a.alias, clase: a.clase,
    titulo: texto(p.titulo, 80, 'título'), texto: texto(p.texto, 1000, 'texto'),
    estado: a.tipo === 'maestro' ? 'aprobada' : 'pendiente', moderadaPor: '',
    codigoClase: a.codigo, maestroClase: a.maestroId
  };
  anadir('Sugerencias', s);
  return { estado: s.estado };
}

function enviarResena(p) {
  const a = autor(p);
  const val = Number(p.valoracion);
  if (!(val >= 1 && val <= 5)) throw new Error('La valoración debe estar entre 1 y 5 estrellas.');
  const r = {
    id: 'L' + Date.now().toString(36), creada: ahoraLocal().completo, tipoAutor: a.tipo, alias: a.alias, clase: a.clase,
    libro: texto(p.libro, 100, 'título del libro'), autorLibro: texto(p.autorLibro || '-', 80, 'autor'),
    valoracion: String(Math.round(val)), texto: texto(p.texto, 1000, 'reseña'),
    estado: a.tipo === 'maestro' ? 'aprobada' : 'pendiente', moderadaPor: '',
    codigoClase: a.codigo, maestroClase: a.maestroId
  };
  anadir('Resenas', r);
  return { estado: r.estado };
}

function autor(p) {
  if (p.token) { const m = maestroPorToken(p.token); return { tipo: 'maestro', alias: m.nombre, clase: '', codigo: '', maestroId: m.id }; }
  const codigo = String(p.codigoClase || '').trim().toUpperCase();
  const clase = filas('Clases').find(c => c.codigo === codigo && c.activa === 'sí');
  if (!clase) throw new Error('El código de clase no es correcto.');
  // Freno anti-avalancha: máximo 20 envíos por clase cada 10 minutos.
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('envios_' + codigo) || 0);
  if (n >= 20) throw new Error('Tu clase ha enviado muchos mensajes seguidos. Espera un poco.');
  cache.put('envios_' + codigo, String(n + 1), 600);
  return { tipo: 'alumno', alias: texto(p.alias, 20, 'nombre'), clase: clase.nombre, codigo: clase.codigo, maestroId: '' };
}

/* ===================== CLASES DEL CENTRO ===================== */

/** Crea las clases fijas que falten y desactiva las antiguas creadas a mano. Se ejecuta sola (una vez cada 6 h). */
function prepararClases() {
  const cache = CacheService.getScriptCache();
  if (cache.get('clases_v2')) return;
  const candado = LockService.getScriptLock();
  if (!candado.tryLock(15000)) return;
  try {
    const nombres = CLASES_CENTRO.map(c => c[0]);
    let clases = filas('Clases');
    clases.filter(c => c.activa === 'sí' && nombres.indexOf(c.nombre) === -1).forEach(c => escribir('Clases', c._fila, { activa: 'no' }));
    clases = filas('Clases');
    const codigos = clases.map(c => c.codigo);
    CLASES_CENTRO.forEach(([nombre, alumnos]) => {
      if (clases.some(c => c.nombre === nombre && c.activa === 'sí')) return;
      let codigo; do { codigo = generarCodigo(); } while (codigos.indexOf(codigo) !== -1);
      codigos.push(codigo);
      anadir('Clases', { codigo, nombre, maestroId: '', activa: 'sí', alta: ahoraLocal().completo, alumnos: String(alumnos), maestros: '' });
    });
    SpreadsheetApp.flush();
    cache.put('clases_v2', '1', 21600);
  } finally { candado.releaseLock(); }
}

const listaMaestros = c => String(c.maestros || '').split(',').filter(Boolean);
const clasesDe = m => filas('Clases').filter(c => c.activa === 'sí' && listaMaestros(c).indexOf(m.id) !== -1).map(c => c.nombre);
/** ¿Puede este maestro gestionar lo escrito en esa clase? */
function esDeSusClases(m, x, mias) {
  return m.rol === 'admin' || x.maestroClase === m.id || (x.clase && mias.indexOf(x.clase) !== -1);
}

/** { token, clases:[nombres] } — las clases en las que da clase este maestro. */
function misClases(p) {
  const m = maestroPorToken(p.token);
  const elegidas = Array.isArray(p.clases) ? p.clases : [];
  filas('Clases').filter(c => c.activa === 'sí').forEach(c => {
    const lista = listaMaestros(c).filter(id => id !== m.id);
    if (elegidas.indexOf(c.nombre) !== -1) lista.push(m.id);
    escribir('Clases', c._fila, { maestros: lista.join(',') });
  });
  return { clases: elegidas };
}

/** { token, nombre } — da un código nuevo a una clase (el anterior deja de funcionar). */
function cambiarCodigoClase(p) {
  const m = maestroPorToken(p.token);
  const c = filas('Clases').find(x => x.nombre === p.nombre && x.activa === 'sí');
  if (!c || (m.rol !== 'admin' && listaMaestros(c).indexOf(m.id) === -1)) throw new Error('No puedes cambiar el código de esa clase.');
  const codigos = filas('Clases').map(x => x.codigo);
  let codigo; do { codigo = generarCodigo(); } while (codigos.indexOf(codigo) !== -1);
  escribir('Clases', c._fila, { codigo });
  return { nombre: c.nombre, codigo };
}

/** { token } → lo que espera revisión: cada maestro, lo de sus clases; el administrador, todo. */
function pendientes(p) {
  const m = maestroPorToken(p.token);
  const mias = clasesDe(m);
  const pend = (h) => filas(h)
    .filter(x => x.estado === 'pendiente' && esDeSusClases(m, x, mias))
    .map(x => { const o = Object.assign({}, x); delete o._fila; o.mia = x.maestroClase === m.id || mias.indexOf(x.clase) !== -1; return o; });
  return { sugerencias: pend('Sugerencias'), resenas: pend('Resenas'), todas: m.rol === 'admin' };
}

/** { token, tipo:'sugerencia'|'resena', id, decision:'aprobada'|'rechazada' } */
function moderar(p) {
  const m = maestroPorToken(p.token);
  const hoja = p.tipo === 'resena' ? 'Resenas' : 'Sugerencias';
  if (['aprobada', 'rechazada'].indexOf(p.decision) === -1) throw new Error('Decisión no válida.');
  const x = filas(hoja).find(f => f.id === p.id);
  if (!x) throw new Error('No encontrado.');
  if (!esDeSusClases(m, x, clasesDe(m))) throw new Error('Solo puedes revisar los comentarios de tus clases.');
  if (x.estado !== 'pendiente') throw new Error('Otro maestro ya lo ha revisado.');
  escribir(hoja, x._fila, { estado: p.decision, moderadaPor: m.nombre });
  return { id: x.id, estado: p.decision };
}

function numAlumnos(v) {
  const n = Math.round(Number(v));
  if (!(n >= 1 && n <= 60)) throw new Error('Indica cuántos alumnos tiene la clase (entre 1 y 60).');
  return String(n);
}

/** { token, codigo, alumnos } — corrige el número de alumnos de una clase. */
function alumnosClase(p) {
  const m = maestroPorToken(p.token);
  const c = filas('Clases').find(x => x.codigo === p.codigo);
  if (!c || (m.rol !== 'admin' && listaMaestros(c).indexOf(m.id) === -1)) throw new Error('No puedes modificar esa clase.');
  escribir('Clases', c._fila, { alumnos: numAlumnos(p.alumnos) });
  return { codigo: c.codigo, alumnos: Number(p.alumnos) };
}

/** Las clases del centro (y, para el administrador, los códigos antiguos) con el alumnado que ha participado. */
function listarClases(p) {
  const m = maestroPorToken(p.token);
  const envios = filas('Sugerencias').concat(filas('Resenas')).filter(x => x.tipoAutor === 'alumno');
  const participantes = (filtro) => {
    const porAlumno = {};
    envios.filter(filtro).forEach(x => {
      const k = normalizar(x.alias);
      const a = porAlumno[k] || (porAlumno[k] = { alias: x.alias, envios: 0, publicados: 0, ultimo: '', mensajes: [] });
      a.envios++;
      a.mensajes.push(x.libro !== undefined
        ? { tipo: 'resena', libro: x.libro, valoracion: x.valoracion, texto: x.texto, estado: x.estado, creada: x.creada }
        : { tipo: 'sugerencia', titulo: x.titulo, texto: x.texto, estado: x.estado, creada: x.creada });
      if (x.estado === 'aprobada') a.publicados++;
      if (x.creada > a.ultimo) a.ultimo = x.creada;
    });
    const lista = Object.keys(porAlumno).map(k => porAlumno[k]).sort((a, b) => b.envios - a.envios);
    lista.forEach(a => a.mensajes.sort((x, y) => x.creada < y.creada ? 1 : -1));
    return lista;
  };
  const todas = filas('Clases');
  const activas = CLASES_CENTRO.map(([nombre]) => todas.find(c => c.nombre === nombre && c.activa === 'sí')).filter(Boolean);
  const antiguas = m.rol === 'admin' ? todas.filter(c => c.activa !== 'sí') : [];
  return {
    clases: activas.map(c => {
      const mia = listaMaestros(c).indexOf(m.id) !== -1;
      const ve = mia || m.rol === 'admin';
      return { nombre: c.nombre, codigo: ve ? c.codigo : '', activa: 'sí', mia, alumnos: Number(c.alumnos) || 0,
        participantes: ve ? participantes(x => x.clase === c.nombre && (!x.codigoClase || x.codigoClase === c.codigo || activas.every(o => o.codigo !== x.codigoClase))) : [] };
    }),
    antiguas: antiguas.map(c => ({ nombre: c.nombre, codigo: c.codigo, activa: 'no', participantes: participantes(x => x.codigoClase === c.codigo) }))
  };
}

/** { token, codigo } — borra un código antiguo (desactivado) y todo lo que se escribió con él. */
function eliminarClase(p) {
  const m = maestroPorToken(p.token);
  const c = filas('Clases').find(x => x.codigo === p.codigo);
  if (!c || (c.maestroId !== m.id && m.rol !== 'admin')) throw new Error('No puedes eliminar esa clase.');
  if (c.activa === 'sí') throw new Error('Solo se pueden eliminar códigos desactivados.');
  let borrados = 0;
  ['Sugerencias', 'Resenas'].forEach(h => {
    filas(h).filter(x => x.codigoClase === c.codigo).map(x => x._fila).sort((a, b) => b - a)
      .forEach(f => { hoja(h).deleteRow(f); borrados++; });
  });
  hoja('Clases').deleteRow(c._fila);
  return { codigo: c.codigo, borrados };
}

/** { token, tipo, id } — retira algo ya publicado (lo devuelve a «rechazada»). */
function retirar(p) {
  const m = maestroPorToken(p.token);
  const h = p.tipo === 'resena' ? 'Resenas' : 'Sugerencias';
  const x = filas(h).find(f => f.id === p.id);
  if (!x) throw new Error('No encontrado.');
  const suyo = esDeSusClases(m, x, clasesDe(m)) || (x.tipoAutor === 'maestro' && x.alias === m.nombre);
  if (!suyo) throw new Error('Solo puedes retirar lo de tus clases.');
  escribir(h, x._fila, { estado: 'rechazada', moderadaPor: m.nombre + ' (retirado)' });
  return { id: x.id };
}

/* ===================== UTILIDADES ===================== */

function hoja(nombre) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre); }

function filas(nombre) {
  const h = hoja(nombre), cab = HOJAS[nombre];
  const n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2, 1, n, cab.length).getDisplayValues().map((v, i) => {
    const o = { _fila: i + 2 };
    cab.forEach((c, j) => o[c] = v[j]);
    return o;
  });
}

function anadir(nombre, obj) {
  const cab = HOJAS[nombre];
  const h = hoja(nombre);
  // Si la hoja se creó con una versión anterior, completa la fila de títulos.
  if (h.getLastColumn() < cab.length) h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  const fila = h.getLastRow() + 1;
  h.getRange(fila, 1, 1, cab.length).setNumberFormat('@').setValues([cab.map(c => obj[c] == null ? '' : String(obj[c]))]);
}

function escribir(nombre, fila, cambios) {
  const cab = HOJAS[nombre], h = hoja(nombre);
  Object.keys(cambios).forEach(c => h.getRange(fila, cab.indexOf(c) + 1).setValue(cambios[c]));
}

function leerConfig() {
  const o = {};
  filas('Config').forEach(f => o[f.clave] = f.valor);
  return o;
}

function guardarConfig(clave, valor) {
  const f = filas('Config').find(x => x.clave === clave);
  if (f) escribir('Config', f._fila, { valor }); else anadir('Config', { clave, valor });
}

function ahoraLocal() {
  const d = new Date();
  const fecha = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const hora = Utilities.formatDate(d, TZ, 'HH:mm');
  return { fecha, hora, completo: fecha + ' ' + hora };
}

function tramoEn(hora) {
  const t = TRAMOS.find(x => hora >= x.inicio && hora < x.fin);
  return t ? t.n : null;
}

function validarFecha(f) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f || '')) throw new Error('Fecha no válida.');
  const d = new Date(f + 'T12:00:00Z');
  if (isNaN(d) || d.toISOString().slice(0, 10) !== f) throw new Error('Fecha no válida.');
  const dia = d.getUTCDay();
  if (dia === 0 || dia === 6) throw new Error('La biblioteca solo se reserva de lunes a viernes.');
  const festivo = todosFestivos().find(x => x.fecha === f);
  if (festivo) throw new Error('Ese día no hay clase (' + (festivo.motivo || 'festivo') + ').');
  return f;
}

function sumarDias(fecha, n) {
  const d = new Date(fecha + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function texto(v, max, campo) {
  const s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  if (!s) throw new Error('Falta el campo «' + campo + '».');
  if (s.length > max) throw new Error('El campo «' + campo + '» es demasiado largo (máx. ' + max + ').');
  // Evita que Sheets interprete el texto como fórmula.
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function normalizar(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function resumen(sal, pin) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + pin, Utilities.Charset.UTF_8));
}

function generarCodigo() {
  const palabras = ['LUNA', 'SOL', 'RIO', 'OLIVO', 'LINCE', 'BUHO', 'NUBE', 'ROBLE', 'FARO', 'COMETA', 'ATLAS', 'PLUMA', 'TINTA', 'CUENTO', 'MAPA', 'ARCO'];
  return palabras[Math.floor(Math.random() * palabras.length)] + '-' + (10 + Math.floor(Math.random() * 90));
}
