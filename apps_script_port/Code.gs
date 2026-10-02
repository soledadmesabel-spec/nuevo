/**
 * MESA DE PEDIDOS JUDICIALES - WhatsApp
 * Port a Google Apps Script (Web App + Google Sheets como base de datos)
 * Original: React + Express + JSON local ("Remix Remix Remix Registro de
 * Pedidos Judiciales WhatsApp", generado en Google AI Studio).
 *
 * Este script debe pegarse en un proyecto de Apps Script VINCULADO a una
 * hoja de cálculo de Google (Extensiones > Apps Script desde el Sheet).
 * Ver README.md para instrucciones de despliegue paso a paso.
 */

// ============================================================
// CONFIGURACIÓN GENERAL
// ============================================================

const SHEET_PEDIDOS = 'Pedidos';
const SHEET_JUZGADOS = 'Juzgados';
const SHEET_CONFIG = 'Config';

const PEDIDOS_HEADERS = [
  'id', 'fechaCreacion', 'saludo', 'tipoTramite', 'expediente', 'juzgado',
  'materia', 'especialista', 'requerimiento', 'solicitante', 'telefono',
  'colegiaturaOCasilla', 'prioridad', 'estado', 'observaciones', 'fechaEscritoPendiente'
];

const JUZGADOS_HEADERS = [
  'id', 'nombre', 'especialidad', 'sede', 'juez', 'especialistas'
];

const CONFIG_KEYS = [
  'nombreGrupo', 'destinatarioDefault', 'telefonoCoordinador',
  'enlaceGrupoWhatsapp', 'tiposTramite', 'materiasFrecuentes'
];

const DEFAULT_ADMIN_PASSWORD = 'admin123';

// ============================================================
// MENÚ / INICIALIZACIÓN (ejecutar una vez desde el editor o desde el menú)
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mesa de Pedidos Judiciales')
    .addItem('Inicializar / Reparar hojas', 'inicializarHojas')
    .addItem('Restablecer directorio de juzgados', 'restablecerJuzgados')
    .addItem('Restablecer configuración', 'restablecerConfig')
    .addToUi();
}

/**
 * Crea las hojas necesarias si no existen y siembra datos iniciales.
 * Es seguro volver a ejecutarla: no borra datos ya cargados.
 */
function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Pedidos ---
  let shPedidos = ss.getSheetByName(SHEET_PEDIDOS);
  if (!shPedidos) {
    shPedidos = newTextSheet_(ss, SHEET_PEDIDOS, PEDIDOS_HEADERS);
    seedInitialPedidos_(shPedidos);
  }

  // --- Juzgados ---
  let shJuzgados = ss.getSheetByName(SHEET_JUZGADOS);
  if (!shJuzgados) {
    shJuzgados = newTextSheet_(ss, SHEET_JUZGADOS, JUZGADOS_HEADERS);
    seedInitialJuzgados_(shJuzgados);
  }

  // --- Config ---
  let shConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!shConfig) {
    shConfig = newTextSheet_(ss, SHEET_CONFIG, ['clave', 'valor']);
    seedInitialConfig_(shConfig);
  }

  // --- Password admin (Script Properties, nunca en la hoja) ---
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PASSWORD')) {
    props.setProperty('ADMIN_PASSWORD', DEFAULT_ADMIN_PASSWORD);
  }

  SpreadsheetApp.getUi().alert('Hojas listas. Contraseña admin por defecto: ' + DEFAULT_ADMIN_PASSWORD);
}

// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.view = (e && e.parameter && e.parameter.view) || '';
  return t.evaluate()
    .setTitle('Mesa de Pedidos Judiciales')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// HELPERS DE HOJA
// ============================================================

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Falta la hoja "' + name + '". Ejecuta "Inicializar / Reparar hojas" desde el menú Mesa de Pedidos Judiciales.');
  return sh;
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(r => r.join('') !== '') // ignorar filas vacías
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = serializable_(h, r[i]); });
      return obj;
    });
}

/** Sheets devuelve Date para celdas con formato fecha; google.script.run no puede serializarlas. */
function serializable_(header, v) {
  if (v instanceof Date) {
    return header === 'fechaEscritoPendiente'
      ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : v.toISOString();
  }
  return v;
}

/** Fuerza texto plano para que Sheets no convierta fechas ISO ni teléfonos "+51..." en fecha/fórmula. */
function newTextSheet_(ss, name, headers) {
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  return sh;
}

function findRowIndexById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // 1-indexed, +1 por header
  }
  return -1;
}

function isoNow_() {
  return new Date().toISOString();
}

function newId_(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
}

// ============================================================
// PEDIDOS (equivalente a /api/pedidos)
// ============================================================

/** Devuelve todos los pedidos, más recientes primero. */
function obtenerPedidos() {
  const sh = getSheet_(SHEET_PEDIDOS);
  const list = sheetToObjects_(sh);
  list.sort((a, b) => new Date(b.fechaCreacion) - new Date(a.fechaCreacion));
  return list;
}

/** Crea o actualiza un pedido (equivalente a POST /api/pedidos). */
function guardarPedido(pedido) {
  if (!pedido || !pedido.expediente) {
    throw new Error('Datos del pedido incompletos: falta el número de expediente.');
  }
  const sh = getSheet_(SHEET_PEDIDOS);

  if (!pedido.id) pedido.id = newId_('ped');
  if (!pedido.fechaCreacion) pedido.fechaCreacion = isoNow_();
  if (!pedido.estado) pedido.estado = 'Pendiente';

  const rowIndex = findRowIndexById_(sh, pedido.id);
  const rowValues = PEDIDOS_HEADERS.map(h => pedido[h] !== undefined ? pedido[h] : '');

  if (rowIndex > 0) {
    sh.getRange(rowIndex, 1, 1, PEDIDOS_HEADERS.length).setValues([rowValues]);
  } else {
    sh.appendRow(rowValues);
  }
  return pedido;
}

/** Actualiza estado / observaciones / prioridad de un pedido (equivalente a PUT). */
function actualizarPedido(id, cambios) {
  const sh = getSheet_(SHEET_PEDIDOS);
  const rowIndex = findRowIndexById_(sh, id);
  if (rowIndex < 0) throw new Error('Pedido no encontrado: ' + id);

  const colIndex = {};
  PEDIDOS_HEADERS.forEach((h, i) => { colIndex[h] = i + 1; });

  ['estado', 'observaciones', 'prioridad'].forEach(field => {
    if (cambios[field] !== undefined) {
      sh.getRange(rowIndex, colIndex[field]).setValue(cambios[field]);
    }
  });

  const updatedRow = sh.getRange(rowIndex, 1, 1, PEDIDOS_HEADERS.length).getValues()[0];
  const obj = {};
  PEDIDOS_HEADERS.forEach((h, i) => { obj[h] = updatedRow[i]; });
  return obj;
}

/** Elimina un pedido (equivalente a DELETE). */
function eliminarPedido(id) {
  const sh = getSheet_(SHEET_PEDIDOS);
  const rowIndex = findRowIndexById_(sh, id);
  if (rowIndex > 0) sh.deleteRow(rowIndex);
  return { success: true };
}

/** Devuelve un CSV (string) de todos los pedidos, para exportar. */
function obtenerCsvPedidos() {
  const pedidos = obtenerPedidos();
  const headers = ['ID', 'Fecha', 'Expediente', 'Tipo de Tramite', 'Juzgado', 'Materia', 'Especialista', 'Requerimiento', 'Solicitante', 'Telefono', 'Prioridad', 'Estado', 'Observaciones'];
  const esc = (v) => '"' + String(v || '').replace(/"/g, '""') + '"';
  const rows = pedidos.map(p => [
    p.id || '',
    p.fechaCreacion ? new Date(p.fechaCreacion).toLocaleString() : '',
    esc(p.expediente), esc(p.tipoTramite), esc(p.juzgado), esc(p.materia),
    esc(p.especialista), esc(p.requerimiento), esc(p.solicitante), esc(p.telefono),
    p.prioridad || 'Normal', p.estado || 'Pendiente', esc(p.observaciones)
  ]);
  return [headers.join(','), ...rows.map(r => r.join(','))].join('\r\n');
}

// ============================================================
// JUZGADOS / DIRECTORIO
// ============================================================

function obtenerJuzgados() {
  const sh = getSheet_(SHEET_JUZGADOS);
  return sheetToObjects_(sh).map(j => ({
    ...j,
    especialistas: j.especialistas ? String(j.especialistas).split('|').map(s => s.trim()).filter(Boolean) : []
  }));
}

function guardarJuzgado(juzgado) {
  const sh = getSheet_(SHEET_JUZGADOS);
  if (!juzgado.id) juzgado.id = newId_('juz');

  const rowValues = [
    juzgado.id, juzgado.nombre, juzgado.especialidad, juzgado.sede || '', juzgado.juez || '',
    Array.isArray(juzgado.especialistas) ? juzgado.especialistas.join('|') : (juzgado.especialistas || '')
  ];

  const rowIndex = findRowIndexById_(sh, juzgado.id);
  if (rowIndex > 0) {
    sh.getRange(rowIndex, 1, 1, JUZGADOS_HEADERS.length).setValues([rowValues]);
  } else {
    // los juzgados nuevos van primero, como en el original (unshift)
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, JUZGADOS_HEADERS.length).setValues([rowValues]);
  }
  return juzgado;
}

function eliminarJuzgado(id) {
  const sh = getSheet_(SHEET_JUZGADOS);
  const rowIndex = findRowIndexById_(sh, id);
  if (rowIndex > 0) sh.deleteRow(rowIndex);
  return { success: true };
}

function restablecerJuzgados() {
  const sh = getSheet_(SHEET_JUZGADOS);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, JUZGADOS_HEADERS.length).clearContent();
  seedInitialJuzgados_(sh);
  return obtenerJuzgados();
}

// ============================================================
// CONFIGURACIÓN
// ============================================================

function obtenerConfig() {
  const sh = getSheet_(SHEET_CONFIG);
  const values = sh.getDataRange().getValues().slice(1);
  const map = {};
  values.forEach(r => { map[r[0]] = r[1]; });

  return {
    nombreGrupo: map.nombreGrupo || '',
    destinatarioDefault: map.destinatarioDefault || '',
    telefonoCoordinador: String(map.telefonoCoordinador || ''),
    enlaceGrupoWhatsapp: map.enlaceGrupoWhatsapp || '',
    tiposTramite: map.tiposTramite ? JSON.parse(map.tiposTramite) : [],
    materiasFrecuentes: map.materiasFrecuentes ? JSON.parse(map.materiasFrecuentes) : []
  };
  // NOTA: adminPassword nunca se expone aquí (vive en Script Properties).
}

function guardarConfig(config) {
  const sh = getSheet_(SHEET_CONFIG);
  const values = sh.getDataRange().getValues();
  const rowOfKey = {};
  for (let i = 1; i < values.length; i++) rowOfKey[values[i][0]] = i + 1;

  const setKey = (key, val) => {
    const serialized = (key === 'tiposTramite' || key === 'materiasFrecuentes') ? JSON.stringify(val) : val;
    if (rowOfKey[key]) {
      sh.getRange(rowOfKey[key], 2).setValue(serialized);
    } else {
      sh.appendRow([key, serialized]);
    }
  };

  CONFIG_KEYS.forEach(key => {
    if (config[key] !== undefined) setKey(key, config[key]);
  });

  return obtenerConfig();
}

function restablecerConfig() {
  const sh = getSheet_(SHEET_CONFIG);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 2).clearContent();
  seedInitialConfig_(sh);
  return obtenerConfig();
}

// ============================================================
// ADMIN (login / password)
// ============================================================

function loginAdmin(password) {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('ADMIN_PASSWORD') || DEFAULT_ADMIN_PASSWORD;
  if (password === current) {
    return { success: true, token: 'adm-' + Date.now() };
  }
  return { success: false, message: 'Contraseña incorrecta' };
}

function cambiarPasswordAdmin(currentPassword, newPassword) {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('ADMIN_PASSWORD') || DEFAULT_ADMIN_PASSWORD;

  if (currentPassword !== current) {
    return { success: false, message: 'La contraseña actual no coincide' };
  }
  if (!newPassword || newPassword.trim().length < 4) {
    return { success: false, message: 'La nueva contraseña debe tener al menos 4 caracteres' };
  }
  props.setProperty('ADMIN_PASSWORD', newPassword.trim());
  return { success: true, message: 'Contraseña actualizada correctamente' };
}

// ============================================================
// CARGA INICIAL COMBINADA (para minimizar llamadas desde el cliente)
// ============================================================

function obtenerDatosIniciales() {
  return {
    pedidos: obtenerPedidos(),
    config: obtenerConfig(),
    juzgados: obtenerJuzgados(),
    webAppUrl: ScriptApp.getService().getUrl()
  };
}

// ============================================================
// SEMILLAS (datos iniciales, réplica de src/data/initialData.ts)
// ============================================================

function seedInitialPedidos_(sheet) {
  const now = Date.now();
  const rows = [
    ['ped-001', new Date(now - 3600 * 1000 * 2).toISOString(), 'Buen dia dra. Yuly', 'Proveido Escrito ( Indicar si es reterativo)', '02098-2015-8-0401-JR-CI-10', '10mo Juzgado Civil', 'Reivindicación', 'Francis Zegarra Cardenas', 'Desde el 07 de julio del año en curso se encuentra pendiente de resolver escritos para poder impulsar el proceso. El expediente se tramita con la nueva especialista Dra. Francis Zegarra Cardenas. Pese a que existe un proceso disciplinario, la demora sigue perjudicando.', 'Dr. Marco Aurelio Vargas', '+51 958 123 456', 'CAA 45892 / Casilla SINOE 11420', 'Muy Urgente', 'Pendiente', '', '2026-07-07'],
    ['ped-002', new Date(now - 3600 * 1000 * 18).toISOString(), 'Buen día Dra. Yuly', 'Emisión de Sentencia primera instancia  y segunda instancia', '01452-2021-0-0401-JR-CI-02', '2do Juzgado Civil', 'Obligación de Dar Suma de Dinero', 'Gisella Ramos Fernández', 'El proceso se encuentra con informe oral realizado y al despacho para sentenciar desde hace más de 45 días hábiles, habiendo vencido en exceso el plazo legal para emitir sentencia.', 'Dra. Beatriz Lucero', '+51 984 765 432', 'CAA 38710', 'Urgente', 'En Trámite', 'Coordinado con especialista para revisión el viernes.', ''],
    ['ped-003', new Date(now - 3600 * 1000 * 36).toISOString(), 'Buenas tardes Dra. Yuly', 'Notificación', '00874-2023-0-0401-JR-CI-01', '1er Juzgado Civil', 'Desalojo por Ocupación Precaria', 'Rosa Salazar Quispe', 'Se solicita impulsar la notificación de la Resolución N° 04 a la parte demandada en su domicilio real, habiendo transcurrido más de 30 días sin que se adjunte el cargo de notificación respectivo a los autos.', 'Dr. Carlos Enrique Mendoza', '+51 991 234 876', 'CAA 50122', 'Normal', 'Atendido', 'Cédula remitida a Central de Notificaciones SERNOT.', '']
  ];
  sheet.getRange(2, 1, rows.length, PEDIDOS_HEADERS.length).setValues(rows);
}

function seedInitialJuzgados_(sheet) {
  const j = (id, nombre, especialidad, sede, juez, especialistas) => [id, nombre, especialidad, sede, juez, especialistas.join('|')];
  const rows = [
    j('juz-civ-10', '10mo Juzgado Civil', 'Civil', 'Sede Central - Palacio de Justicia', 'Dra. SALAS FLORES, ZORAIDA JULIA', ['Francis Zegarra Cardenas', 'GARCIA JURADO MARISOL', 'LIZARZABURU ROMERO CHRISTIAN EDUARDO', 'MANTILLA VALDIVIA ERIK ALEXANDER', 'FERNANDEZ HUAQUIPACO NORMA HILDA', 'PEÑA CONDORI JESUS WANPIAU']),
    j('juz-civ-01', '1er Juzgado Civil', 'Civil', 'Sede Central', 'Dra. Patricia Arispe', ['Mario Benavente Ponce', 'Rosa Salazar Quispe', 'Alonso Delgado Morales']),
    j('juz-civ-02', '2do Juzgado Civil', 'Civil', 'Sede Central', 'Dr. Víctor Guzmán', ['Gisella Ramos Fernández', 'Jorge Luis Cáceres', 'Ana María Valdivia']),
    j('juz-civ-03', '3er Juzgado Civil', 'Civil', 'Sede Central', 'Dra. Elena Cornejo', ['Héctor Velásquez Neyra', 'Claudia Medina Torres']),
    j('juz-civ-04', '4to Juzgado Civil', 'Civil', 'Sede Central', 'Dr. Javier Monroy', ['Sandra Beltrán Zúñiga', 'Christian Portugal Apaza']),
    j('juz-civ-05', '5to Juzgado Civil', 'Civil', 'Sede Central', 'Dra. Carmen Montes', ['Lucía Carpio Rivero', 'Daniel Flores Mamani']),
    j('juz-civ-06', '6to Juzgado Civil', 'Civil', 'Sede Central', 'Dr. Fernando Vizcarra', ['Rodrigo Pacheco Díaz', 'Karina Gutiérrez Vera']),
    j('juz-civ-07', '7mo Juzgado Civil', 'Civil', 'Sede Central', 'Dra. Teresa Alarcón', ['Miguel Ángel Pinto', 'Vanessa Rojas Condori']),
    j('juz-civ-08', '8vo Juzgado Civil', 'Civil', 'Sede Central', 'Dr. Luis Guillén', ['Esteban Quiroz Meza', 'Monica Ticona Huanca']),
    j('juz-civ-09', '9no Juzgado Civil', 'Civil', 'Sede Central', 'Dra. Patricia Vilca', ['Diego Manrique Soto', 'Paola Yáñez Barreda']),
    j('juz-lab-01', '1er Juzgado de Trabajo (NLPT)', 'Laboral', 'Sede Laboral', 'Dr. Hugo Barrera', ['Silvia Colquehuanca', 'Gonzalo Miranda Paz']),
    j('juz-lab-02', '2do Juzgado de Trabajo (NLPT)', 'Laboral', 'Sede Laboral', 'Dra. Rocío Cuadros', ['Marcos Espinoza Cruz', 'Evelyn Choque Villegas']),
    j('juz-fam-01', '1er Juzgado de Familia', 'Familia', 'Sede Familia', 'Dra. Clara Fuentes', ['Jessica Huamán Arias', 'Walter Zevallos Ramos']),
    j('juz-const-01', '1er Juzgado Constitucional', 'Constitucional', 'Palacio de Justicia', 'Dr. Oscar Ballón', ['César Huanca Portilla', 'Diana Cárdenas Becerra']),
    j('juz-paz-01', '1er Juzgado de Paz Letrado Civil', 'Paz Letrado', 'Sede Módulo Básico', 'Dra. Nadia Portocarrero', ['Manuel Bustamante Vera', 'Liliana Cuentas Salas'])
  ];
  sheet.getRange(2, 1, rows.length, JUZGADOS_HEADERS.length).setValues(rows);
}

function seedInitialConfig_(sheet) {
  const tiposTramite = [
    { id: 1, titulo: 'Proveido Escrito ( Indicar si es reterativo)', descripcion: 'Solicitud de proveído de escrito pendiente de resolver o calificar (indicar si es reiterativo).', icono: 'FileClock', tagColor: 'blue' },
    { id: 2, titulo: 'Emisión de Sentencia primera instancia  y segunda instancia', descripcion: 'Expediente expedito para resolver o emitir sentencia en primera o segunda instancia / auto final.', icono: 'Scale', tagColor: 'amber' },
    { id: 3, titulo: 'Notificación', descripcion: 'Impulso de diligenciamiento de cédulas físicas, electrónicas y devolución de cargos.', icono: 'Send', tagColor: 'emerald' },
    { id: 4, titulo: 'Elevacion de Expedientes', descripcion: 'Elevación de actuados a Sala Superior o Corte Suprema por apelación o casación concedida.', icono: 'Layers', tagColor: 'purple' },
    { id: 5, titulo: 'Diligencias o Audiencias', descripcion: 'Programación, reprogramación o realización de audiencias, declaraciones o inspecciones.', icono: 'Calendar', tagColor: 'rose' },
    { id: 6, titulo: 'Trámite Documentario', descripcion: 'Expedición de copias certificadas, oficios, exhortos, endoses y desarchivamiento.', icono: 'Files', tagColor: 'cyan' },
    { id: 7, titulo: 'Otros y Sugerencias', descripcion: 'Otras solicitudes procesales, incidencias administrativas o sugerencias de atención judicial.', icono: 'Sparkles', tagColor: 'slate' }
  ];
  const materiasFrecuentes = [
    'Reivindicación', 'Obligación de Dar Suma de Dinero', 'Desalojo por Ocupación Precaria',
    'Nulidad de Acto Jurídico', 'Prescripción Adquisitiva de Dominio', 'Otorgamiento de Escritura Pública',
    'División y Partición de Bienes', 'Indemnización por Daños y Perjuicios', 'Alimentos y Aumento de Alimentos',
    'Pago de Beneficios Sociales (NLPT)', 'Reposición Laboral', 'Acción de Amparo', 'Sucesión Intestada',
    'Ejecución de Garantías', 'Tercería de Propiedad', 'Medida Cautelar Fuera de Proceso'
  ];

  const rows = [
    ['nombreGrupo', 'Grupo de Impulso Procesal Judicial - WhatsApp'],
    ['destinatarioDefault', 'Buen día Dra. Yuly'],
    ['telefonoCoordinador', '51987654321'],
    ['enlaceGrupoWhatsapp', 'https://chat.whatsapp.com/pedidos-judiciales-oficial'],
    ['tiposTramite', JSON.stringify(tiposTramite)],
    ['materiasFrecuentes', JSON.stringify(materiasFrecuentes)]
  ];
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}
