/**
 * TALONARIO DEL SORTEO — web pública + panel de responsables con login
 *
 * Enlace público  (URL de la aplicación web)          → formulario para participantes.
 * Panel privado   (misma URL terminada en ?p=panel)    → acceso con correo y contraseña.
 * Base de datos   (esta hoja de Google)                → Participantes, Resultados, Ajustes, Responsables.
 *
 * Seguridad: todas las funciones del panel exigen una sesión válida. Las funciones internas
 * terminan en "_" (no se pueden llamar desde el navegador). Las contraseñas se guardan cifradas.
 */

const HOJA_PART = 'Participantes';
const HOJA_AJ = 'Ajustes';
const HOJA_RES = 'Resultados';
const HOJA_RESP = 'Responsables';

const COLS = ['Fecha', 'Nombre', 'Correo', 'Teléfono', 'Boletos', 'Números', 'Desde', 'Hasta',
  'Canal', 'Cobrado', 'Papeleta enviada', 'Anulado', 'Notas', 'ID', 'Registrado por'];
const C = {FECHA: 1, NOMBRE: 2, CORREO: 3, TEL: 4, CANT: 5, NUMS: 6, DESDE: 7, HASTA: 8,
  CANAL: 9, COBRADO: 10, ENVIADA: 11, ANULADO: 12, NOTAS: 13, ID: 14, POR: 15};
const COLS_RES = ['Fecha', 'Premio', 'Número', 'Nombre', 'Correo', 'ID', 'Sacado por'];
const COLS_RESP = ['Nombre', 'Correo (usuario)', 'Clave cifrada (no tocar)', 'Activo', 'Último acceso'];

const AJ = {
  nombre: ['Nombre del sorteo', 'Sorteo'],
  premio: ['Premio', ''],
  precio: ['Precio por boleto (€)', 2],
  fecha: ['Fecha del sorteo', ''],
  lugar: ['Lugar o forma del sorteo', ''],
  organiza: ['Organiza', ''],
  contacto: ['Correo de contacto', ''],
  pago: ['Cómo pagar', 'Bizum al 600 000 000 indicando tu nombre'],
  inicio: ['Primer número', 1],
  total: ['Total de boletos (0 = sin límite)', 500],
  maxPersona: ['Máximo de boletos por persona en el formulario', 20],
  soloCobrados: ['Solo entran en el sorteo los cobrados (SI/NO)', 'SI'],
  abierto: ['Formulario abierto (SI/NO)', 'SI'],
  mensaje: ['Mensaje del correo', '¡Gracias por participar! Guarda este correo: es tu papeleta y lleva tus números.']
};

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const SESION_SEG = 21600; // 6 horas

/* ===================== Utilidades ===================== */

function ss_() {
  const activa = SpreadsheetApp.getActiveSpreadsheet();
  if (activa) return activa;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('El sorteo aún no está preparado.');
  return SpreadsheetApp.openById(id);
}

function hoja_(nombre) {
  const sh = ss_().getSheetByName(nombre);
  if (!sh) throw new Error('El sorteo aún no está preparado.');
  return sh;
}

function siNo_(v) { return String(v).trim().toUpperCase().indexOf('S') === 0; }

function fechaISO_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function fechaTxt_(v) {
  const iso = fechaISO_(v);
  if (iso) { const p = iso.split('-'); return Number(p[2]) + ' de ' + MESES[Number(p[1]) - 1] + ' de ' + p[0]; }
  return String(v || '').trim();
}

function ajustes_() {
  const sh = hoja_(HOJA_AJ);
  const vals = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 2).getValues();
  const m = {};
  vals.forEach(function (r) { m[String(r[0]).trim()] = r[1]; });
  const g = function (k) { const lbl = AJ[k][0]; return m[lbl] !== undefined && m[lbl] !== '' ? m[lbl] : AJ[k][1]; };
  const inicio = parseInt(g('inicio'), 10), total = parseInt(g('total'), 10), maxP = parseInt(g('maxPersona'), 10);
  return {
    nombre: String(g('nombre')).trim(), premio: String(g('premio')).trim(), precio: Number(g('precio')) || 0,
    fecha: fechaTxt_(g('fecha')), fechaISO: fechaISO_(g('fecha')), lugar: String(g('lugar')).trim(),
    organiza: String(g('organiza')).trim(), contacto: String(g('contacto')).trim(), pago: String(g('pago')).trim(),
    inicio: isNaN(inicio) ? 1 : inicio, total: isNaN(total) ? 0 : Math.max(0, total),
    maxPersona: isNaN(maxP) || maxP < 1 ? 20 : maxP,
    soloCobrados: siNo_(g('soloCobrados')), abierto: siNo_(g('abierto')), mensaje: String(g('mensaje')).trim()
  };
}

function digitos_(aj) { return Math.max(3, String(aj.total > 0 ? aj.inicio + aj.total - 1 : 9999).length); }
function fmt_(n, aj) { return String(n).padStart(digitos_(aj), '0'); }
function eur_(n) { return (Math.round((Number(n) || 0) * 100) / 100).toFixed(2).replace('.', ',') + ' €'; }
function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c];
  });
}

function ultimoNumero_(sh, aj) {
  let max = aj.inicio - 1;
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, C.HASTA, last - 1, 1).getValues().forEach(function (r) {
    const v = Number(r[0]); if (!isNaN(v) && v > max) max = v;
  });
  return max;
}

function filaPorId_(sh, col, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function bloquear_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('Hay mucha actividad ahora mismo. Inténtalo de nuevo en unos segundos.');
  return lock;
}

/* ===================== Menú de la hoja (solo para quien abre la hoja) ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sorteo')
    .addItem('Añadir responsable o nueva contraseña', 'anadirResponsable')
    .addItem('Ver enlaces de la web', 'verEnlace')
    .addSeparator()
    .addItem('Preparar hojas (solo la primera vez)', 'preparar')
    .addToUi();
}

function preparar() {
  const ui = SpreadsheetApp.getUi(); // solo funciona dentro de la hoja
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  let p = ss.getSheetByName(HOJA_PART);
  if (!p) p = ss.insertSheet(HOJA_PART, 0);
  if (p.getLastRow() === 0) {
    p.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#0b6b6a').setFontColor('#ffffff');
    p.setFrozenRows(1);
    p.getRange(2, C.COBRADO, 3000, 3).insertCheckboxes();
    p.getRange(2, C.FECHA, 3000, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    p.setColumnWidth(C.NOMBRE, 200); p.setColumnWidth(C.CORREO, 220); p.setColumnWidth(C.NUMS, 200);
    p.hideColumns(C.DESDE, 2); p.hideColumns(C.ID);
    p.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$L2=TRUE')
      .setFontColor('#9aa5a7').setStrikethrough(true).setRanges([p.getRange('A2:O3001')]).build()]);
  }
  let a = ss.getSheetByName(HOJA_AJ);
  if (!a) {
    a = ss.insertSheet(HOJA_AJ);
    const filas = Object.keys(AJ).map(function (k) { return AJ[k]; });
    a.getRange(1, 1, filas.length, 2).setValues(filas);
    a.getRange(1, 1, filas.length, 1).setFontWeight('bold');
    a.setColumnWidth(1, 360); a.setColumnWidth(2, 380);
  }
  let r = ss.getSheetByName(HOJA_RES);
  if (!r) { r = ss.insertSheet(HOJA_RES); r.getRange(1, 1, 1, COLS_RES.length).setValues([COLS_RES]).setFontWeight('bold'); r.setFrozenRows(1); r.hideColumns(6); }
  let s = ss.getSheetByName(HOJA_RESP);
  if (!s) {
    s = ss.insertSheet(HOJA_RESP);
    s.getRange(1, 1, 1, COLS_RESP.length).setValues([COLS_RESP]).setFontWeight('bold'); s.setFrozenRows(1);
    s.getRange(2, 4, 200, 1).insertCheckboxes(); s.setColumnWidth(2, 240); s.setColumnWidth(3, 120);
  }
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
  ui.alert('Hojas listas.\n\n1) Usa Sorteo → Añadir responsable para crear tu propio acceso.\n2) Publica la web: Implementar → Nueva implementación → Aplicación web.');
}

function anadirResponsable() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Añadir responsable', 'Nombre de la persona:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const nombre = r1.getResponseText().trim();
  const r2 = ui.prompt('Añadir responsable', 'Su correo (será su usuario para entrar):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const email = r2.getResponseText().trim().toLowerCase();
  if (!nombre || !EMAIL_RE.test(email)) { ui.alert('Falta el nombre o el correo no es válido.'); return; }

  const pass = generarClave_();
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const clave = salt + ':' + hash_(salt, pass);
  const sh = hoja_(HOJA_RESP);
  const fila = filaPorId_(sh, 2, email);
  if (fila > 0) sh.getRange(fila, 1, 1, 4).setValues([[nombre, email, clave, true]]);
  else sh.appendRow([nombre, email, clave, true, '']);

  const url = ScriptApp.getService().getUrl();
  ui.alert('Acceso creado para ' + nombre,
    'Usuario: ' + email + '\nContraseña: ' + pass + '\n\n' +
    (url ? 'Panel: ' + url + '?p=panel\n\n' : 'Cuando publiques la web, el panel estará en la URL de la web terminada en ?p=panel\n\n') +
    'Apúntala ahora y pásasela a esa persona: no se vuelve a mostrar.\n' +
    'Para quitarle el acceso, desmarca «Activo» en la hoja Responsables. Si pierde la contraseña, vuelve a añadirla con el mismo correo.',
    ui.ButtonSet.OK);
}

function verEnlace() {
  const ui = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();
  if (!url) { ui.alert('Todavía no has publicado la web. En Apps Script: Implementar → Nueva implementación → Aplicación web (Ejecutar como: yo · Acceso: cualquier usuario).'); return; }
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<p><b>Para participantes</b> (compártelo):</p><input style="width:100%;padding:8px" value="' + esc_(url) + '" onclick="this.select()" readonly>' +
    '<p><b>Panel de responsables</b>:</p><input style="width:100%;padding:8px" value="' + esc_(url) + '?p=panel" onclick="this.select()" readonly></div>'
  ).setWidth(480).setHeight(220);
  ui.showModalDialog(html, 'Enlaces del sorteo');
}

/* ===================== Contraseñas y sesiones ===================== */

function hash_(salt, pass) {
  let d = salt + '|' + pass;
  let b;
  for (let i = 0; i < 500; i++) {
    b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d, Utilities.Charset.UTF_8);
    d = Utilities.base64Encode(b) + salt;
  }
  return Utilities.base64Encode(b);
}

function generarClave_() {
  const abc = 'abcdefghjkmnpqrstuvwxyz23456789';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now());
  let s = '';
  for (let i = 0; i < 12; i++) { s += abc[(bytes[i] + 256) % abc.length]; if (i === 3 || i === 7) s += '-'; }
  return s;
}

function responsable_(email) {
  const sh = hoja_(HOJA_RESP);
  const fila = filaPorId_(sh, 2, String(email).trim().toLowerCase());
  if (fila < 0) return null;
  const r = sh.getRange(fila, 1, 1, 5).getValues()[0];
  return {fila: fila, nombre: String(r[0]), email: String(r[1]), clave: String(r[2]), activo: r[3] === true};
}

/** Público: entrar al panel. */
function login(email, pass) {
  email = String(email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const k = 'intentos_' + email;
  const n = Number(cache.get(k) || 0);
  if (n >= 5) throw new Error('Demasiados intentos fallidos. Espera 15 minutos.');
  const r = responsable_(email);
  let ok = false;
  if (r && r.activo && r.clave.indexOf(':') > 0) {
    const partes = r.clave.split(':');
    ok = hash_(partes[0], String(pass || '')) === partes[1];
  }
  if (!ok) {
    cache.put(k, String(n + 1), 900);
    Utilities.sleep(700);
    throw new Error('Correo o contraseña incorrectos.');
  }
  cache.remove(k);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('sesion_' + token, JSON.stringify({email: r.email, nombre: r.nombre}), SESION_SEG);
  hoja_(HOJA_RESP).getRange(r.fila, 5).setValue(new Date());
  return {token: token, nombre: r.nombre, email: r.email};
}

function logout(token) { if (token) CacheService.getScriptCache().remove('sesion_' + token); return true; }

function sesion_(token) {
  if (!token || !/^[a-f0-9]{64}$/.test(String(token))) throw new Error('SESION');
  const v = CacheService.getScriptCache().get('sesion_' + token);
  if (!v) throw new Error('SESION');
  const s = JSON.parse(v);
  const r = responsable_(s.email);
  if (!r || !r.activo) { logout(token); throw new Error('SESION'); }
  return s;
}

/* ===================== Web ===================== */

function doGet(e) {
  const url = ScriptApp.getService().getUrl();
  try {
    if (e && e.parameter && e.parameter.p === 'panel') {
      const t = HtmlService.createTemplateFromFile('Panel');
      t.url = url;
      t.nombre = ajustes_().nombre;
      return t.evaluate().setTitle('Panel · ' + ajustes_().nombre).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    const t = HtmlService.createTemplateFromFile('Formulario');
    t.url = url;
    t.datos = JSON.stringify(datosPublicos()).replace(/</g, '\\u003c');
    return t.evaluate().setTitle(ajustes_().nombre).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput('<p style="font-family:Arial;padding:24px">El sorteo todavía no está preparado.</p>');
  }
}

/** Público: datos que ve cualquiera. */
function datosPublicos() {
  const aj = ajustes_();
  const ultimo = ultimoNumero_(hoja_(HOJA_PART), aj);
  const res = hoja_(HOJA_RES);
  const ganadores = res.getLastRow() > 1 ? res.getRange(2, 1, res.getLastRow() - 1, 4).getValues().map(function (r) {
    const partes = String(r[3]).trim().split(/\s+/);
    return {premio: String(r[1]), numero: fmt_(r[2], aj), nombre: partes[0] + (partes[1] ? ' ' + partes[1][0].toUpperCase() + '.' : '')};
  }) : [];
  return {
    nombre: aj.nombre, premio: aj.premio, precio: aj.precio, fecha: aj.fecha, lugar: aj.lugar,
    organiza: aj.organiza, contacto: aj.contacto, pago: aj.pago, maxPersona: aj.maxPersona, abierto: aj.abierto,
    quedan: aj.total > 0 ? Math.max(0, aj.inicio + aj.total - 1 - ultimo) : null,
    siguiente: fmt_(ultimo + 1, aj), ganadores: ganadores
  };
}

/** Público: inscripción desde el formulario. */
function registrarWeb(d) {
  return registrar_(d, 'web', 'Formulario web');
}

function registrar_(d, canal, por) {
  d = d || {};
  const nombre = String(d.nombre || '').trim().replace(/\s+/g, ' ').slice(0, 120);
  const email = String(d.email || '').trim().toLowerCase().slice(0, 160);
  const tel = String(d.telefono || '').trim().slice(0, 40);
  const notas = String(d.notas || '').trim().slice(0, 300);
  const cant = parseInt(d.cantidad, 10);
  const aj = ajustes_();
  const cache = CacheService.getScriptCache();
  const kWeb = 'web_' + email;

  if (canal === 'web') {
    if (!aj.abierto) throw new Error('El plazo para pedir boletos está cerrado.');
    if (nombre.split(' ').length < 2) throw new Error('Escribe tu nombre y al menos un apellido.');
    if (!EMAIL_RE.test(email)) throw new Error('Revisa tu correo: es donde te llegará la papeleta.');
    if (!(cant >= 1 && cant <= aj.maxPersona)) throw new Error('Puedes pedir entre 1 y ' + aj.maxPersona + ' boletos.');
    if (cache.get(kWeb)) throw new Error('Acabas de apuntarte con este correo. Si quieres más boletos, espera un minuto.');
  } else {
    if (!nombre) throw new Error('Escribe el nombre completo.');
    if (email && !EMAIL_RE.test(email)) throw new Error('El correo no parece válido.');
    if (!(cant >= 1 && cant <= 1000)) throw new Error('Pon un número de boletos válido.');
  }

  const sh = hoja_(HOJA_PART);
  const lock = bloquear_();
  let desde, hasta, id;
  try {
    const ultimo = ultimoNumero_(sh, aj);
    desde = ultimo + 1;
    hasta = desde + cant - 1;
    if (aj.total > 0 && hasta > aj.inicio + aj.total - 1) {
      const quedan = Math.max(0, aj.inicio + aj.total - 1 - ultimo);
      throw new Error(quedan ? 'Solo quedan ' + quedan + ' boletos.' : 'Se han agotado los boletos.');
    }
    const nums = [];
    for (let n = desde; n <= hasta; n++) nums.push(fmt_(n, aj));
    id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    sh.appendRow([new Date(), nombre, email, tel, cant, nums.join(', '), desde, hasta,
      canal === 'web' ? 'Online' : 'En mano', canal === 'mano' && !!d.cobrado, false, false, notas, id, por]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  // El límite de 60 s solo se activa si la inscripción se ha guardado de verdad.
  if (canal === 'web') cache.put(kWeb, '1', 60);

  let enviado = false, aviso = '';
  if (email && (canal === 'web' || d.enviar !== false)) {
    try {
      enviarPapeleta_({nombre: nombre, email: email, desde: desde, hasta: hasta, cobrado: canal === 'mano' && !!d.cobrado}, aj);
      const f = filaPorId_(sh, C.ID, id);
      if (f > 0) sh.getRange(f, C.ENVIADA).setValue(true);
      enviado = true;
    } catch (e) {
      aviso = canal === 'web' ? 'No hemos podido enviarte el correo ahora. La organización te lo mandará.' : 'No se pudo enviar el correo: ' + e.message;
    }
  }
  const numeros = [];
  for (let n = desde; n <= hasta; n++) numeros.push(fmt_(n, aj));
  return {id: id, numeros: numeros, nombre: nombre, email: email, enviado: enviado, aviso: aviso,
    total: eur_(cant * aj.precio), datos: datosPublicos()};
}

/* ===================== Correo con la papeleta ===================== */

function enviarPapeleta_(p, aj) {
  const nums = [];
  for (let n = p.desde; n <= p.hasta; n++) nums.push(fmt_(n, aj));
  const cant = nums.length;
  const importe = aj.precio ? eur_(cant * aj.precio) : '';
  const asunto = 'Tu papeleta · ' + aj.nombre + ' · ' + (cant === 1 ? 'nº ' + nums[0] : cant + ' boletos');
  const texto = [
    'Hola ' + p.nombre + ',', '', aj.mensaje, '',
    'Sorteo: ' + aj.nombre, aj.premio ? 'Premio: ' + aj.premio : '', aj.fecha ? 'Fecha del sorteo: ' + aj.fecha : '',
    aj.lugar ? 'Dónde: ' + aj.lugar : '', '',
    'Tus números (' + cant + '): ' + nums.join(', '),
    importe ? 'Importe: ' + importe + (p.cobrado ? ' (cobrado)' : ' (pendiente de pago)') : '',
    !p.cobrado && aj.pago ? 'Cómo pagar: ' + aj.pago : '', '',
    aj.organiza ? '— ' + aj.organiza : '', aj.contacto ? 'Contacto: ' + aj.contacto : ''
  ].filter(function (l, i, a) { return l !== '' || a[i - 1] !== ''; }).join('\n');

  const chips = nums.map(function (n) {
    return '<span style="display:inline-block;font-family:\'Courier New\',monospace;font-size:18px;border:1px solid #c9b77a;border-radius:6px;padding:4px 10px;margin:0 6px 6px 0;background:#fffaf0">' + n + '</span>';
  }).join('');
  const pagoHtml = !p.cobrado && aj.pago
    ? '<tr><td style="padding:14px 0 0"><div style="background:#f6e7cc;border-radius:8px;padding:12px 14px;font-size:14px;color:#5a3a00"><b>Pendiente de pago (' + esc_(importe) + ').</b><br>' + esc_(aj.pago) + '</div></td></tr>' : '';
  const html =
    '<div style="background:#eef3f2;padding:24px 12px;font-family:Arial,Helvetica,sans-serif;color:#14262a">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;margin:0 auto">' +
    '<tr><td style="padding:0 0 14px;font-size:15px;line-height:1.5">Hola ' + esc_(p.nombre) + ',<br>' + esc_(aj.mensaje) + '</td></tr>' +
    '<tr><td style="background:#fbf1cf;border-radius:10px;border:1px solid #e6d7a2">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td style="padding:20px 20px 16px;vertical-align:top">' +
    '<div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#6b5d2e">' + esc_(aj.organiza || 'Papeleta') + (aj.fecha ? ' · ' + esc_(aj.fecha) : '') + '</div>' +
    '<div style="font-family:\'Arial Narrow\',Arial,sans-serif;font-weight:bold;font-size:26px;text-transform:uppercase;color:#2b2410;margin:6px 0 4px">' + esc_(aj.nombre) + '</div>' +
    (aj.premio ? '<div style="font-size:14px;color:#2b2410;margin-bottom:12px">Premio: <b>' + esc_(aj.premio) + '</b></div>' : '') +
    '<div>' + chips + '</div>' +
    '<div style="border-top:1px solid #c9b77a;margin-top:10px;padding-top:8px;font-size:13px;color:#2b2410">' + esc_(p.nombre) + ' · ' + cant + ' boleto' + (cant === 1 ? '' : 's') + (importe ? ' · ' + importe : '') + '</div></td>' +
    '<td width="110" style="border-left:2px dashed #c9b77a;text-align:center;vertical-align:middle;padding:16px 10px">' +
    '<div style="font-size:10px;letter-spacing:2px;color:#6b5d2e">Nº</div>' +
    '<div style="font-family:\'Courier New\',monospace;font-size:28px;color:#2b2410">' + nums[0] + '</div>' +
    (cant > 1 ? '<div style="font-size:11px;color:#6b5d2e">y ' + (cant - 1) + ' más</div>' : '') +
    '</td></tr></table></td></tr>' + pagoHtml +
    '<tr><td style="padding:14px 0 0;font-size:13px;color:#56696d;line-height:1.5">' +
    (aj.lugar ? 'El sorteo se hará: ' + esc_(aj.lugar) + '.<br>' : '') + (aj.contacto ? '¿Dudas? Escribe a ' + esc_(aj.contacto) + '.' : '') +
    '</td></tr></table></div>';
  const op = {to: p.email, subject: asunto, body: texto, htmlBody: html, name: aj.organiza || aj.nombre};
  if (aj.contacto && EMAIL_RE.test(aj.contacto)) op.replyTo = aj.contacto;
  MailApp.sendEmail(op);
}

/* ===================== API del panel (todas exigen sesión) ===================== */

function estado(token) { return estado_(sesion_(token)); }

function estado_(s) {
  const aj = ajustes_();
  const sh = hoja_(HOJA_PART);
  const last = sh.getLastRow();
  const parts = last > 1 ? sh.getRange(2, 1, last - 1, COLS.length).getValues().filter(function (r) { return r[C.ID - 1]; }).map(function (r) {
    return {id: String(r[C.ID - 1]), fecha: r[0] instanceof Date ? r[0].toISOString() : '', nombre: String(r[1]), email: String(r[2]),
      tel: String(r[3]), cant: Number(r[4]) || 0, desde: Number(r[6]), hasta: Number(r[7]), canal: String(r[8]),
      cobrado: r[9] === true, enviada: r[10] === true, anulado: r[11] === true, notas: String(r[12] || ''), por: String(r[14] || '')};
  }) : [];
  const rs = hoja_(HOJA_RES);
  const res = rs.getLastRow() > 1 ? rs.getRange(2, 1, rs.getLastRow() - 1, COLS_RES.length).getValues().map(function (r) {
    return {fecha: r[0] instanceof Date ? r[0].toISOString() : '', premio: String(r[1]), numero: Number(r[2]), nombre: String(r[3]), id: String(r[5]), por: String(r[6])};
  }) : [];
  aj.digitos = digitos_(aj);
  aj.siguiente = ultimoNumero_(sh, aj) + 1;
  return {aj: aj, parts: parts, res: res, yo: s, url: ScriptApp.getService().getUrl(), cupo: MailApp.getRemainingDailyQuota(), ahora: new Date().toISOString()};
}

function ventaMano(token, d) {
  const s = sesion_(token);
  const r = registrar_(d, 'mano', s.nombre);
  r.estado = estado_(s);
  return r;
}

function actualizar(token, id, campos) {
  const s = sesion_(token);
  const sh = hoja_(HOJA_PART);
  const lock = bloquear_();
  try {
    const f = filaPorId_(sh, C.ID, id);
    if (f < 0) throw new Error('No se encuentra esa fila. Puede que alguien la haya borrado de la hoja.');
    if (campos.email && !EMAIL_RE.test(String(campos.email).trim())) throw new Error('El correo no parece válido.');
    const mapa = {nombre: C.NOMBRE, email: C.CORREO, tel: C.TEL, notas: C.NOTAS, cobrado: C.COBRADO, anulado: C.ANULADO};
    Object.keys(campos || {}).forEach(function (k) {
      if (!mapa[k]) return;
      let v = campos[k];
      if (k === 'cobrado' || k === 'anulado') v = v === true;
      else v = String(v || '').trim().slice(0, 300);
      if (k === 'email') v = v.toLowerCase();
      sh.getRange(f, mapa[k]).setValue(v);
    });
    if (campos.anulado === true) sh.getRange(f, C.NOTAS).setValue(String(sh.getRange(f, C.NOTAS).getValue() || '') + ' [anulado por ' + s.nombre + ']');
  } finally { lock.releaseLock(); }
  return estado_(s);
}

function reenviar(token, id) {
  const s = sesion_(token);
  const aj = ajustes_();
  const sh = hoja_(HOJA_PART);
  const f = filaPorId_(sh, C.ID, id);
  if (f < 0) throw new Error('No se encuentra esa fila.');
  const r = sh.getRange(f, 1, 1, COLS.length).getValues()[0];
  const email = String(r[C.CORREO - 1]).trim();
  if (!EMAIL_RE.test(email)) throw new Error('Esta persona no tiene un correo válido. Corrígelo en su ficha.');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Gmail no deja enviar más correos hoy. Prueba mañana.');
  enviarPapeleta_({nombre: r[C.NOMBRE - 1], email: email, desde: Number(r[C.DESDE - 1]), hasta: Number(r[C.HASTA - 1]), cobrado: r[C.COBRADO - 1] === true}, aj);
  sh.getRange(f, C.ENVIADA).setValue(true);
  return estado_(s);
}

function enviarPendientes(token) {
  const s = sesion_(token);
  const t0 = Date.now();
  const aj = ajustes_();
  const sh = hoja_(HOJA_PART);
  const last = sh.getLastRow();
  let ok = 0, fallos = 0;
  if (last > 1) {
    const filas = sh.getRange(2, 1, last - 1, COLS.length).getValues();
    for (let i = 0; i < filas.length; i++) {
      const r = filas[i];
      const email = String(r[C.CORREO - 1]).trim();
      if (r[C.ENVIADA - 1] === true || r[C.ANULADO - 1] === true || !EMAIL_RE.test(email)) continue;
      if (MailApp.getRemainingDailyQuota() < 1 || Date.now() - t0 > 240000) break;
      try {
        enviarPapeleta_({nombre: r[C.NOMBRE - 1], email: email, desde: Number(r[C.DESDE - 1]), hasta: Number(r[C.HASTA - 1]), cobrado: r[C.COBRADO - 1] === true}, aj);
        sh.getRange(i + 2, C.ENVIADA).setValue(true); ok++;
      } catch (e) { fallos++; }
    }
  }
  const est = estado_(s);
  const quedan = est.parts.filter(function (p) { return !p.enviada && !p.anulado && EMAIL_RE.test(p.email); }).length;
  return {ok: ok, fallos: fallos, quedan: quedan, estado: est};
}

/** Entero uniforme en [0, max) sin sesgo de módulo (muestreo por rechazo sobre 48 bits). */
function azar_(max) {
  const RANGO = 281474976710656; // 2^48
  const limite = RANGO - (RANGO % max);
  for (;;) {
    const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Math.random());
    let x = 0;
    for (let i = 0; i < 6; i++) x = x * 256 + ((b[i] + 256) % 256);
    if (x < limite) return x % max;
  }
}

function sortear(token, premio) {
  const s = sesion_(token);
  const lock = bloquear_();
  try {
    const aj = ajustes_();
    const est = estado_(s);
    const salidos = {};
    est.res.forEach(function (r) { salidos[r.numero] = true; });
    const bolsa = [];
    est.parts.forEach(function (p) {
      if (p.anulado || (aj.soloCobrados && !p.cobrado)) return;
      for (let n = p.desde; n <= p.hasta; n++) if (!salidos[n]) bolsa.push({n: n, p: p});
    });
    if (!bolsa.length) throw new Error('No hay boletos en el bombo.');
    const g = bolsa[azar_(bolsa.length)];
    const muestra = [];
    for (let i = 0; i < 36; i++) muestra.push(fmt_(bolsa[azar_(bolsa.length)].n, aj));
    hoja_(HOJA_RES).appendRow([new Date(), String(premio || 'Premio').slice(0, 80), g.n, g.p.nombre, g.p.email,
      Utilities.getUuid().replace(/-/g, '').slice(0, 10), s.nombre]);
    SpreadsheetApp.flush();
    return {numero: fmt_(g.n, aj), nombre: g.p.nombre, muestra: muestra, estado: estado_(s)};
  } finally { lock.releaseLock(); }
}

function quitarResultado(token, id) {
  const s = sesion_(token);
  const sh = hoja_(HOJA_RES);
  const f = filaPorId_(sh, 6, id);
  if (f > 0) sh.deleteRow(f);
  return estado_(s);
}

function guardarAjustes(token, o) {
  const s = sesion_(token);
  o = o || {};
  const actual = ajustes_();
  const usados = ultimoNumero_(hoja_(HOJA_PART), actual) - actual.inicio + 1;
  const inicio = parseInt(o.inicio, 10), total = parseInt(o.total, 10);
  if (usados > 0 && inicio !== actual.inicio) throw new Error('El primer número no se puede cambiar cuando ya hay boletos repartidos.');
  if (total > 0 && total < usados) throw new Error('Ya hay ' + usados + ' números repartidos: el total no puede ser menor.');
  if (o.contacto && !EMAIL_RE.test(String(o.contacto).trim())) throw new Error('El correo de contacto no parece válido.');
  const valores = {
    nombre: String(o.nombre || '').trim() || 'Sorteo', premio: String(o.premio || '').trim(), precio: Math.max(0, Number(o.precio) || 0),
    fecha: fechaISO_(o.fecha), lugar: String(o.lugar || '').trim(), organiza: String(o.organiza || '').trim(),
    contacto: String(o.contacto || '').trim(), pago: String(o.pago || '').trim(),
    inicio: isNaN(inicio) ? actual.inicio : Math.max(0, inicio), total: isNaN(total) ? actual.total : Math.max(0, total),
    maxPersona: Math.max(1, parseInt(o.maxPersona, 10) || 20), soloCobrados: o.soloCobrados ? 'SI' : 'NO',
    abierto: o.abierto ? 'SI' : 'NO', mensaje: String(o.mensaje || '').trim() || AJ.mensaje[1]
  };
  const sh = hoja_(HOJA_AJ);
  const lbls = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues().map(function (r) { return String(r[0]).trim(); });
  Object.keys(valores).forEach(function (k) {
    const i = lbls.indexOf(AJ[k][0]);
    if (i >= 0) sh.getRange(i + 1, 2).setValue(valores[k]);
    else { sh.appendRow([AJ[k][0], valores[k]]); lbls.push(AJ[k][0]); }
  });
  if (valores.fecha) { const i = lbls.indexOf(AJ.fecha[0]); if (i >= 0) sh.getRange(i + 1, 2).setNumberFormat('@').setValue(valores.fecha); }
  return estado_(s);
}
