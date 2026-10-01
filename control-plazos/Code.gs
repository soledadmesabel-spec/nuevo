/**
 * Control de Plazos de Expedientes – Google Apps Script
 * Versión 1.1 – agrega alertas por correo de expedientes vencidos.
 *
 * Lee la hoja de cálculo con las columnas:
 *   SecretarioInicial | Nro. Expediente | Sentido | Dependencia |
 *   Fecha de Asignacion | SECRETARIO | Fecha Limite | Estado
 * y publica una aplicación web (Index.html) que muestra los expedientes
 * admitidos y el plazo que tiene cada tramitador.
 */

// ===== CONFIGURACIÓN =====
const CONFIG = {
  // Nombre de la pestaña con los datos. Si está vacío se usa la primera hoja.
  NOMBRE_HOJA: '',
  // Si el script NO está vinculado a la hoja, pegue aquí el ID del Google Sheets
  // (lo que va entre /d/ y /edit en la URL). Vacío = hoja vinculada.
  ID_HOJA: '',
  // Días (calendario) que tiene el tramitador desde la fecha de asignación.
  PLAZO_DIAS: 5,
  // Días antes del vencimiento en que se marca "Por vencer".
  ALERTA_DIAS: 2,
  ZONA_HORARIA: 'America/Lima',

  // ----- Alertas por correo (v1.1) -----
  // Correo del supervisor o responsable del área: recibe el reporte general de vencidos.
  // Déjelo vacío ('') para no enviar el reporte general.
  CORREO_SUPERVISOR: '',
  // Correo de cada tramitador. La clave es el nombre en minúsculas, tal como aparece
  // en la columna SECRETARIO. Si un tramitador no tiene correo, se omite y se
  // avisa en el reporte del supervisor.
  CORREOS_TRAMITADORES: {
    'kenia': '',
    'kevin': '',
    'yeraldo': ''
  },
  NOMBRE_REMITENTE: 'Control de Plazos Judiciales'
};

const VERSION = '1.1';

// Nombres de encabezado aceptados para cada campo (se comparan sin tildes ni mayúsculas).
const ENCABEZADOS = {
  secretarioInicial: ['secretarioinicial', 'secretario inicial'],
  expediente: ['nro. expediente', 'nro expediente', 'expediente', 'n° expediente'],
  sentido: ['sentido'],
  dependencia: ['dependencia', 'juzgado'],
  fechaAsignacion: ['fecha de asignacion', 'fecha asignacion'],
  tramitador: ['secretario', 'tramitador'],
  fechaLimite: ['fecha limite', 'fecha de vencimiento', 'vencimiento'],
  estado: ['estado']
};

/** Punto de entrada de la aplicación web. */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Control de Plazos – Expedientes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Menú en la hoja de cálculo. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Control de Plazos')
    .addItem('Actualizar fecha límite y estado', 'actualizarHoja')
    .addItem('Enviar alertas de vencidos por correo', 'notificarDesdeMenu')
    .addItem('Abrir panel', 'abrirPanel')
    .addToUi();
}

/** Abre el aplicativo como diálogo dentro de la hoja. */
function abrirPanel() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate().setWidth(1200).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Control de Plazos');
}

/**
 * Devuelve los expedientes ya procesados para el HTML.
 * Las fechas se envían como texto (google.script.run no transmite objetos Date).
 */
function obtenerExpedientes() {
  const hoja = obtenerHoja_();
  const valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return { version: VERSION, hoy: formatear_(hoy_()), plazoDias: CONFIG.PLAZO_DIAS, registros: [] };

  const col = mapearColumnas_(valores[0]);
  if (col.expediente < 0 || col.fechaAsignacion < 0) {
    throw new Error('No se encontraron las columnas "Nro. Expediente" y "Fecha de Asignacion" en la fila 1.');
  }

  const hoy = hoy_();
  const vistos = {};
  const registros = [];

  for (let i = 1; i < valores.length; i++) {
    const f = valores[i];
    const expediente = texto_(f[col.expediente]);
    if (!expediente) continue;

    const asignacion = aFecha_(f[col.fechaAsignacion]);
    let limite = col.fechaLimite >= 0 ? aFecha_(f[col.fechaLimite]) : null;
    if (!limite && asignacion) limite = sumarDias_(asignacion, CONFIG.PLAZO_DIAS);

    const dias = limite ? diferenciaDias_(hoy, limite) : null;
    const sentido = normalizarSentido_(f[col.sentido]);
    const clave = expediente + '|' + sentido;
    vistos[clave] = (vistos[clave] || 0) + 1;

    registros.push({
      fila: i + 1,
      expediente: expediente,
      secretarioInicial: nombre_(f[col.secretarioInicial]),
      sentido: sentido,
      sentidoOriginal: texto_(f[col.sentido]),
      dependencia: normalizarDependencia_(f[col.dependencia]),
      tramitador: nombre_(f[col.tramitador]),
      fechaAsignacion: asignacion ? formatear_(asignacion) : '',
      fechaLimite: limite ? formatear_(limite) : '',
      diasRestantes: dias,
      estado: estadoPlazo_(dias),
      duplicado: false
    });
  }

  // Marca expedientes registrados más de una vez con el mismo sentido.
  registros.forEach(function (r) {
    r.duplicado = vistos[r.expediente + '|' + r.sentido] > 1;
  });

  return { version: VERSION, hoy: formatear_(hoy), plazoDias: CONFIG.PLAZO_DIAS, registros: registros };
}

/**
 * Escribe en la hoja la Fecha Limite (si falta) y recalcula el Estado de cada fila
 * con la fecha de hoy. Se puede programar con un activador diario.
 */
function actualizarHoja() {
  const hoja = obtenerHoja_();
  const rango = hoja.getDataRange();
  const valores = rango.getValues();
  const encabezados = valores[0];
  const col = mapearColumnas_(encabezados);

  // Si faltan, usa las columnas a la derecha de SECRETARIO y les pone encabezado.
  if (col.fechaLimite < 0) col.fechaLimite = col.tramitador + 1;
  if (col.estado < 0) col.estado = col.tramitador + 2;
  if (!texto_(encabezados[col.fechaLimite])) hoja.getRange(1, col.fechaLimite + 1).setValue('Fecha Limite');
  if (!texto_(encabezados[col.estado])) hoja.getRange(1, col.estado + 1).setValue('Estado');

  const hoy = hoy_();
  const salidaLimite = [];
  const salidaEstado = [];
  for (let i = 1; i < valores.length; i++) {
    const f = valores[i];
    const asignacion = aFecha_(f[col.fechaAsignacion]);
    let limite = aFecha_(f[col.fechaLimite]);
    if (!limite && asignacion) limite = sumarDias_(asignacion, CONFIG.PLAZO_DIAS);
    const dias = limite ? diferenciaDias_(hoy, limite) : null;
    salidaLimite.push([limite || '']);
    salidaEstado.push([texto_(f[col.expediente]) ? estadoPlazo_(dias).etiqueta : '']);
  }
  if (salidaLimite.length) {
    hoja.getRange(2, col.fechaLimite + 1, salidaLimite.length, 1)
      .setValues(salidaLimite).setNumberFormat('dd-MM-yyyy');
    hoja.getRange(2, col.estado + 1, salidaEstado.length, 1).setValues(salidaEstado);
  }
}

// ===== ALERTAS POR CORREO (v1.1) =====

/**
 * Envía un solo correo por tramitador con todos sus expedientes vencidos y un
 * reporte general al supervisor. Devuelve un mensaje con el resultado.
 */
function notificarExpedientesVencidos() {
  const datos = obtenerExpedientes();

  // Expedientes vencidos, sin repetir el mismo expediente con el mismo sentido.
  const vistos = {};
  const vencidos = datos.registros.filter(function (r) {
    const clave = r.expediente + '|' + r.sentido;
    if (r.estado.codigo !== 'vencido' || vistos[clave]) return false;
    vistos[clave] = true;
    return true;
  }).sort(function (a, b) { return a.diasRestantes - b.diasRestantes; });

  if (!vencidos.length) return 'No hay expedientes vencidos para notificar.';

  const porTramitador = {};
  vencidos.forEach(function (r) {
    const clave = (r.tramitador || 'Sin tramitador').toLowerCase().trim();
    (porTramitador[clave] = porTramitador[clave] || []).push(r);
  });

  const correosNecesarios = Object.keys(porTramitador).length + (CONFIG.CORREO_SUPERVISOR ? 1 : 0);
  const cuota = MailApp.getRemainingDailyQuota();
  if (cuota < correosNecesarios) {
    throw new Error('Cuota diaria de correo insuficiente: quedan ' + cuota + ' y se necesitan ' + correosNecesarios + '.');
  }

  let enviados = 0;
  const sinCorreo = [];

  // 1. Un correo consolidado por tramitador.
  Object.keys(porTramitador).forEach(function (clave) {
    const lista = porTramitador[clave];
    const destinatario = texto_(CONFIG.CORREOS_TRAMITADORES[clave]);
    if (!destinatario) { sinCorreo.push(lista[0].tramitador || 'Sin tramitador'); return; }
    MailApp.sendEmail({
      to: destinatario,
      subject: '⚠️ Alerta: tiene ' + lista.length + ' expediente(s) con plazo vencido',
      htmlBody: construirCuerpoCorreoHtml_(lista[0].tramitador, lista, datos.hoy),
      name: CONFIG.NOMBRE_REMITENTE
    });
    enviados++;
  });

  // 2. Reporte general al supervisor.
  if (CONFIG.CORREO_SUPERVISOR) {
    MailApp.sendEmail({
      to: CONFIG.CORREO_SUPERVISOR,
      subject: '📊 Reporte general: expedientes vencidos (' + vencidos.length + ')',
      htmlBody: construirCuerpoCorreoSupervisor_(vencidos, datos.hoy, sinCorreo),
      name: CONFIG.NOMBRE_REMITENTE
    });
    enviados++;
  }

  let msg = 'Se enviaron ' + enviados + ' correo(s) de alerta por ' + vencidos.length + ' expediente(s) vencido(s).';
  if (sinCorreo.length) msg += ' Sin correo configurado: ' + sinCorreo.join(', ') + '.';
  if (!enviados) msg += ' Configure CORREOS_TRAMITADORES y CORREO_SUPERVISOR al inicio de Code.gs.';
  return msg;
}

/** Versión para el menú de la hoja: muestra el resultado en pantalla. */
function notificarDesdeMenu() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Enviar alertas', '¿Enviar los correos de alerta de expedientes vencidos?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  ui.alert(notificarExpedientesVencidos());
}

/** Rutina diaria: actualiza la hoja y envía las alertas. */
function ejecutarRutinaDiaria() {
  actualizarHoja();
  Logger.log(notificarExpedientesVencidos());
}

/** Crea (o reemplaza) el activador que ejecuta la rutina diaria a las 7 a.m. */
function crearActivadorDiario() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) {
      const f = t.getHandlerFunction();
      return f === 'ejecutarRutinaDiaria' || f === 'actualizarHoja';
    })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('ejecutarRutinaDiaria').timeBased().everyDays(1).atHour(7).create();
}

function celda_(contenido, estilo) {
  return '<td style="padding:8px;border:1px solid #e5e7eb;' + (estilo || '') + '">' + contenido + '</td>';
}

function encabezado_(titulos, fondo, borde, color) {
  return '<tr style="background:' + fondo + ';color:' + color + ';text-align:left;">' + titulos.map(function (t) {
    return '<th style="padding:8px;border:1px solid ' + borde + ';">' + t + '</th>';
  }).join('') + '</tr>';
}

/** Correo para el tramitador. */
function construirCuerpoCorreoHtml_(nombre, expedientes, hoy) {
  let tabla = '<table style="border-collapse:collapse;width:100%;font-family:sans-serif;font-size:13px;">'
    + encabezado_(['Expediente', 'Sentido', 'Dependencia', 'F. Asignación', 'F. Límite', 'Atraso'], '#fee2e2', '#fca5a5', '#991b1b');
  expedientes.forEach(function (e) {
    tabla += '<tr>'
      + celda_(esc_(e.expediente), 'font-weight:bold;')
      + celda_(esc_(e.sentido))
      + celda_(esc_(e.dependencia))
      + celda_(esc_(e.fechaAsignacion))
      + celda_(esc_(e.fechaLimite), 'color:#b91c1c;')
      + celda_(Math.abs(e.diasRestantes) + ' días', 'color:#b91c1c;font-weight:bold;')
      + '</tr>';
  });
  tabla += '</table>';

  return '<div style="font-family:sans-serif;color:#1f2937;max-width:700px;">'
    + '<h2 style="color:#b91c1c;margin-bottom:4px;">Atención: expedientes con plazo vencido</h2>'
    + '<p>Estimado(a) <b>' + esc_(nombre) + '</b>:</p>'
    + '<p>Al corte de hoy (<b>' + esc_(hoy) + '</b>) se registran los siguientes expedientes asignados a usted que superaron el plazo de ' + CONFIG.PLAZO_DIAS + ' días:</p>'
    + tabla
    + '<p style="margin-top:16px;font-size:12px;color:#6b7280;">Por favor, proceda con la tramitación correspondiente a la brevedad.</p>'
    + '</div>';
}

/** Correo para el supervisor. */
function construirCuerpoCorreoSupervisor_(vencidos, hoy, sinCorreo) {
  const lista = vencidos.slice().sort(function (a, b) {
    return a.tramitador.localeCompare(b.tramitador) || a.diasRestantes - b.diasRestantes;
  });
  let tabla = '<table style="border-collapse:collapse;width:100%;font-family:sans-serif;font-size:13px;">'
    + encabezado_(['Tramitador', 'Expediente', 'Sentido', 'Dependencia', 'F. Límite', 'Días vencido'], '#f3f4f6', '#e5e7eb', '#1f2937');
  lista.forEach(function (e) {
    tabla += '<tr>'
      + celda_(esc_(e.tramitador), 'font-weight:bold;')
      + celda_(esc_(e.expediente))
      + celda_(esc_(e.sentido))
      + celda_(esc_(e.dependencia))
      + celda_(esc_(e.fechaLimite))
      + celda_(Math.abs(e.diasRestantes) + ' días', 'color:#b91c1c;font-weight:bold;')
      + '</tr>';
  });
  tabla += '</table>';

  const aviso = sinCorreo.length
    ? '<p style="background:#fef3c7;color:#92400e;padding:8px 12px;border-radius:6px;">Tramitadores sin correo configurado (no fueron notificados): <b>' + esc_(sinCorreo.join(', ')) + '</b></p>'
    : '';

  return '<div style="font-family:sans-serif;color:#1f2937;max-width:700px;">'
    + '<h2 style="color:#b91c1c;">Reporte de plazos vencidos</h2>'
    + '<p>Se identificaron <b>' + vencidos.length + '</b> expedientes vencidos al <b>' + esc_(hoy) + '</b>.</p>'
    + aviso + tabla
    + '</div>';
}

// ===== AUXILIARES =====

function obtenerHoja_() {
  const libro = CONFIG.ID_HOJA
    ? SpreadsheetApp.openById(CONFIG.ID_HOJA)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!libro) throw new Error('Configure CONFIG.ID_HOJA con el ID de su Google Sheets.');
  const hoja = CONFIG.NOMBRE_HOJA ? libro.getSheetByName(CONFIG.NOMBRE_HOJA) : libro.getSheets()[0];
  if (!hoja) throw new Error('No existe la hoja "' + CONFIG.NOMBRE_HOJA + '".');
  return hoja;
}

function limpiar_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

function mapearColumnas_(encabezados) {
  const norm = encabezados.map(limpiar_);
  const col = {};
  Object.keys(ENCABEZADOS).forEach(function (campo) {
    col[campo] = norm.findIndex(function (h) { return ENCABEZADOS[campo].indexOf(h) >= 0; });
  });
  // "SECRETARIO" sin "Inicial": si no se encontró, toma la 6.ª columna (F) por defecto.
  if (col.tramitador < 0 && encabezados.length >= 6) col.tramitador = 5;
  // Columnas sin encabezado tras SECRETARIO: G = fecha límite, H = estado.
  if (col.fechaLimite < 0 && col.tramitador >= 0 && !norm[col.tramitador + 1] && encabezados.length > col.tramitador + 1) {
    col.fechaLimite = col.tramitador + 1;
  }
  if (col.estado < 0 && col.tramitador >= 0 && !norm[col.tramitador + 2] && encabezados.length > col.tramitador + 2) {
    col.estado = col.tramitador + 2;
  }
  return col;
}

/** Escapa texto de la hoja antes de ponerlo en un correo HTML. */
function esc_(v) {
  return texto_(v).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function texto_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

/** "kenia" -> "Kenia" para que el mismo tramitador no aparezca dos veces. */
function nombre_(v) {
  return texto_(v).toLowerCase().replace(/(^|\s)\S/g, function (c) { return c.toUpperCase(); });
}

function normalizarSentido_(v) {
  const s = limpiar_(v);
  if (!s) return 'Sin sentido';
  if (s.indexOf('admit') === 0 || s === 'admite') return 'Admite';
  if (s.indexOf('conc') === 0) return 'Concede'; // incluye "conceede"
  if (s.indexOf('improc') === 0) return 'Improcedente';
  if (s.indexOf('inadmis') === 0) return 'Inadmisible';
  if (s.indexOf('desist') === 0) return 'Desistimiento';
  if (s.indexOf('incompet') === 0) return 'Incompetencia';
  return s.charAt(0).toUpperCase() + s.slice(1);
}

/** "1ero" / "1ro" -> "1°", "7mo" -> "7°". */
function normalizarDependencia_(v) {
  const s = texto_(v);
  const m = s.match(/^(\d+)\s*(ero|ro|do|ero|to|mo|vo|no|er|°|º)?\.?$/i);
  return m ? m[1] + '° Juzgado' : s;
}

function hoy_() {
  const t = Utilities.formatDate(new Date(), CONFIG.ZONA_HORARIA, 'yyyy-MM-dd').split('-');
  return new Date(Number(t[0]), Number(t[1]) - 1, Number(t[2]));
}

/** Acepta fechas de Sheets o texto "dd-MM-yyyy" / "dd/MM/yyyy". */
function aFecha_(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = texto_(v).match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
  if (!m) return null;
  let anio = Number(m[3]);
  if (anio < 100) anio += 2000;
  return new Date(anio, Number(m[2]) - 1, Number(m[1]));
}

function sumarDias_(fecha, dias) {
  const d = new Date(fecha.getTime());
  d.setDate(d.getDate() + dias);
  return d;
}

function diferenciaDias_(desde, hasta) {
  return Math.round((hasta.getTime() - desde.getTime()) / 86400000);
}

function formatear_(fecha) {
  return Utilities.formatDate(fecha, CONFIG.ZONA_HORARIA, 'dd-MM-yyyy');
}

function estadoPlazo_(dias) {
  if (dias === null) return { codigo: 'sin', etiqueta: 'Sin fecha' };
  if (dias < 0) return { codigo: 'vencido', etiqueta: 'Plazo Vencido' };
  if (dias === 0) return { codigo: 'hoy', etiqueta: 'Vence hoy' };
  if (dias <= CONFIG.ALERTA_DIAS) return { codigo: 'proximo', etiqueta: 'Por vencer' };
  return { codigo: 'plazo', etiqueta: 'En plazo' };
}
