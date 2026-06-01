// ══════════════════════════════════════════
//  API Web (doGet / doPost)
// ══════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data._type === 'budgets') {
      const sheet = getBudgetsSheet();
      sheet.clearContents();
      sheet.appendRow(['Clave', 'Valor']);
      Object.entries(data.budgets).forEach(([k, v]) => sheet.appendRow([k, v]));
      return json({ok: true});
    }
    const sheet = getSheet();
    if (sheet.getLastRow() === 0) addHeader(sheet);
    const ms = data.mesSueldo || {};
    sheet.appendRow([
      data.id, data.fecha, data.descripcion, data.categoria, data.subcategoria,
      data.tipo === 'ingreso' ? data.monto : '',
      data.tipo === 'egreso'  ? data.monto : '',
      new Date(data.fecha + 'T12:00:00').getMonth() + 1,
      new Date(data.fecha + 'T12:00:00').getFullYear(),
      (ms.mes !== undefined ? ms.mes + 1 : ''), ms.año || ''
    ]);
    return json({ok: true});
  } catch(e) {
    return json({ok: false, error: e.message});
  }
}

function doGet(e) {
  try {
    const sheet = getSheet();
    let registros = [];
    if (sheet.getLastRow() > 1) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
      registros = rows
        .filter(r => r[0])
        .map(r => ({
          id:           String(r[0]),
          fecha:        formatFecha(r[1]),
          descripcion:  r[2] || '',
          categoria:    r[3] || '',
          subcategoria: r[4] || '',
          monto:        Number(r[5] || r[6]) || 0,
          tipo:         r[5] ? 'ingreso' : 'egreso',
          mesSueldo:    { mes: (Number(r[9]) || 1) - 1, año: Number(r[10]) || 0 },
          sincronizado: true
        }));
    }
    const budgets = {};
    const bSheet = getBudgetsSheet();
    if (bSheet.getLastRow() > 1) {
      const bRows = bSheet.getRange(2, 1, bSheet.getLastRow() - 1, 2).getValues();
      bRows.filter(r => r[0]).forEach(r => { budgets[String(r[0])] = Number(r[1]); });
    }
    return json({ registros, budgets });
  } catch(e) {
    return json({error: e.message});
  }
}

// ══════════════════════════════════════════
//  Lectura automática de correos bancarios
// ══════════════════════════════════════════

function procesarCorreosBanco() {
  const LABEL_NOMBRE = 'finanzas-procesado';
  const label = GmailApp.getUserLabelByName(LABEL_NOMBRE) || GmailApp.createLabel(LABEL_NOMBRE);
  const sheet = getSheet();
  if (sheet.getLastRow() === 0) addHeader(sheet);

  // IDs ya registrados (para evitar duplicados)
  const idsExistentes = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      .forEach(r => { if (r[0]) idsExistentes.add(String(r[0])); });
  }

  const reglas = [
    { query: 'from:transferencias@bci.cl',                    parser: parsearBCI          },
    { query: 'from:contacto@bci.cl',                          parser: parsearBCITarjeta   },
    { query: 'from:serviciodetransferencias@bancochile.cl',   parser: parsearBancoChile   },
    { query: 'from:noreply@correo.bancoestado.cl',            parser: parsearBancoEstado  },
    { query: 'from:mensajeria@santander.cl',                  parser: parsearSantander    },
  ];

  reglas.forEach(function(regla) {
    const threads = GmailApp.search(regla.query + ' -label:' + LABEL_NOMBRE);
    threads.forEach(function(thread) {
      thread.getMessages().forEach(function(msg) {
        const body = msg.getPlainBody();
        const record = regla.parser(body, msg.getDate());
        if (record && !idsExistentes.has(record.id)) {
          agregarRegistro(sheet, record);
          idsExistentes.add(record.id);
        }
      });
      thread.addLabel(label);
    });
  });
}

// ── Parsers por banco ──────────────────────

function parsearBCI(body, emailDate) {
  if (!body) return null;
  // Ignorar confirmación de recepción en cuenta Copec Pay (es traspaso propio)
  if (/Has recibido una transferencia/i.test(body) && /Copec\s*Pay/i.test(body)) return null;

  const esEgreso  = /Realizaste una transferencia/i.test(body);
  const esIngreso = /Has recibido una transferencia/i.test(body) && !esEgreso;
  if (!esEgreso && !esIngreso) return null;

  const monto = parsearMonto(body, /Monto\s*(?:transferido|recibido)\s*[\s\S]{0,5}\$?([\d.,]+)/i)
             || parsearMonto(body, /\$\s*([\d.,]+)/);
  if (!monto) return null;

  const fecha = parsearFecha(body, /Fecha\s*(?:de\s*abono|de\s*la\s*transferencia)\s*[\s\S]{0,5}(\d{2}\/\d{2}\/\d{4})/i)
             || parsearFecha(body, /(\d{2}\/\d{2}\/\d{4})/)
             || formatearFecha(emailDate);

  const comprobante = (body.match(/N[uú]mero de comprobante[^\d]*(\d+)/i) || [])[1] || '';
  const id = 'bci_' + (comprobante || Utilities.formatDate(emailDate, 'America/Santiago', 'yyyyMMddHHmmss'));

  // Mensaje puede estar en línea separada o misma línea con espacios/tabs
  const mensaje      = campo(body, 'Mensaje');
  const destinatario = campo(body, 'Nombre del destinatario');
  // Para ingresos, extraer nombre del remitente: "de [NOMBRE] hacia tu cuenta"
  const remitente    = (body.match(/Has recibido.*?de\s+([^\n\r]+?)\s+hacia/i) || [])[1] || '';
  const descripcion  = (mensaje || destinatario || remitente).trim();

  return {
    id, fecha, descripcion, monto,
    tipo: esEgreso ? 'egreso' : 'ingreso',
    categoria: esEgreso ? 'Gastos del mes' : categorizarIngreso(descripcion),
    subcategoria: ''
  };
}

function parsearBancoChile(body, emailDate) {
  // "[NOMBRE] le ha transferido $X.XXX, el día DD de MES de YYYY"
  if (!/le ha transferido/i.test(body)) return null;

  const monto = parsearMonto(body, /le ha transferido[^$\d]*\$?([\d.,]+)/i)
             || parsearMonto(body, /Monto[^$\d]*\$?([\d.,]+)/i);
  if (!monto) return null;

  const fecha = parsearFechaTexto(body)
             || parsearFecha(body, /(\d{2}\/\d{2}\/\d{4})/)
             || formatearFecha(emailDate);
  const comprobante = (body.match(/N[uú]mero de comprobante[^\d]*(\d+)/i) || [])[1] || '';
  const id = 'bch_' + (comprobante || Utilities.formatDate(emailDate, 'America/Santiago', 'yyyyMMddHHmmss'));

  const mensaje     = campo(body, 'Mensaje');
  const remitente   = (body.match(/que\s+([A-ZÁÉÍÓÚÑ][^,]+)\s+le ha transferido/i) || [])[1] || '';
  const descripcion = (mensaje || remitente).trim();

  return {
    id, fecha, descripcion, monto,
    tipo: 'ingreso',
    categoria: categorizarIngreso(descripcion),
    subcategoria: ''
  };
}

function parsearBancoEstado(body, emailDate) {
  if (!/Has recibido una Transferencia/i.test(body)) return null;

  const monto = parsearMonto(body, /Monto[^$\d]*\$?([\d.,]+)/i);
  if (!monto) return null;

  // Fecha y hora: DD/MM/YYYY HH:MM:SS
  const fecha = parsearFecha(body, /Fecha y hora\s*[\n\r]+(\d{2}\/\d{2}\/\d{4})/i)
             || parsearFecha(body, /(\d{2}\/\d{2}\/\d{4})/)
             || formatearFecha(emailDate);
  const nTrans = (body.match(/N[°º]\s*transacci[oó]n[^\d]*(\d+)/i) || [])[1] || '';
  const id = 'bce_' + (nTrans || Utilities.formatDate(emailDate, 'America/Santiago', 'yyyyMMddHHmmss'));

  const mensaje     = campo(body, 'Mensaje');
  const remitente   = (body.match(/cliente\s+([^\n\r]+)/i) || [])[1] || '';
  const descripcion = (mensaje || remitente).trim();

  return {
    id, fecha, descripcion, monto,
    tipo: 'ingreso',
    categoria: categorizarIngreso(descripcion),
    subcategoria: ''
  };
}

function parsearSantander(body, emailDate) {
  if (!body) return null;
  if (!/realiz[oó] una transferencia a tu cuenta/i.test(body)) return null;

  const monto = parsearMonto(body, /Monto transferido\s*\$?\s*([\d.,]+)/i)
             || parsearMonto(body, /\$\s*([\d.,]+)/);
  if (!monto) return null;

  const fecha = parsearFecha(body, /con fecha\s+(\d{2}\/\d{2}\/\d{4})/i)
             || parsearFecha(body, /(\d{2}\/\d{2}\/\d{4})/)
             || formatearFecha(emailDate);

  // El nombre puede ocupar 1 o 2 líneas antes de "realizó"
  const remitenteMatch = body.match(/nuestro cliente\s+([\s\S]+?)\s+realiz[oó]/i);
  const remitente = remitenteMatch ? remitenteMatch[1].replace(/\s+/g, ' ').trim() : '';

  // Comentario: solo usarlo si es corto y no contiene texto del footer
  const comentarioRaw = campo(body, 'Comentario');
  const comentario = (comentarioRaw && comentarioRaw.length < 80 && !/imprimir|nota|seguridad/i.test(comentarioRaw))
    ? comentarioRaw : '';

  const descripcion = (comentario || remitente).trim();

  // ID único: fecha + monto + timestamp del email
  const ts = Utilities.formatDate(emailDate, 'America/Santiago', 'yyyyMMddHHmmss');
  const id = 'san_' + ts + '_' + monto;

  return {
    id, fecha, descripcion, monto,
    tipo: 'ingreso',
    categoria: categorizarIngreso(descripcion),
    subcategoria: ''
  };
}

function parsearBCITarjeta(body, emailDate) {
  // "Realizaste un(a) compra en comercio nacional/internacional con tu tarjeta de débito/crédito"
  if (!/compra en comercio/i.test(body)) return null;

  const monto = parsearMonto(body, /Monto[^$\d]*\$?([\d.,]+)/i);
  if (!monto) return null;

  const fecha = parsearFecha(body, /Fecha[^\d]*(\d{2}\/\d{2}\/\d{4})/i)
             || formatearFecha(emailDate);
  const hora  = (body.match(/Hora[^\d]*(\d{2}:\d{2})/i) || [])[1] || '';
  const comercio = campo(body, 'Comercio');

  // ID único: fecha + hora + monto
  const idBase = fecha.replace(/-/g, '') + hora.replace(':', '') + monto;
  const id = 'bci_td_' + idBase;

  return {
    id,
    fecha,
    descripcion: comercio.trim(),
    monto,
    tipo: 'egreso',
    categoria: 'Gastos del mes',
    subcategoria: ''
  };
}

// ── Helpers ───────────────────────────────

// Extrae el valor que sigue a una etiqueta, sea en la misma línea o en la siguiente
function campo(body, etiqueta) {
  const esc = etiqueta.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const m = body.match(new RegExp(esc + '\\s*[:\\t ]*\\s*([^\\n\\r]+)', 'i'));
  return m ? m[1].trim() : '';
}

function parsearMonto(body, regex) {
  const m = body.match(regex);
  if (!m) return null;
  const n = parseInt(m[1].replace(/\./g, '').replace(/,/g, ''));
  return isNaN(n) || n === 0 ? null : n;
}

function parsearFecha(body, regex) {
  const m = body.match(regex);
  if (!m) return null;
  const parts = m[1].split('/');
  if (parts.length !== 3) return null;
  return parts[2] + '-' + parts[1] + '-' + parts[0];
}

// Parsea "el día 29 de mayo de 2026"
function parsearFechaTexto(body) {
  const MESES = {enero:1,febrero:2,marzo:3,abril:4,mayo:5,junio:6,
                 julio:7,agosto:8,septiembre:9,octubre:10,noviembre:11,diciembre:12};
  const m = body.match(/el d[ií]a\s+(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})/i);
  if (!m) return null;
  const mes = MESES[m[2].toLowerCase()];
  if (!mes) return null;
  return m[3] + '-' + String(mes).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
}

function formatearFecha(date) {
  return Utilities.formatDate(date, 'America/Santiago', 'yyyy-MM-dd');
}

function categorizarIngreso(desc) {
  if (/javi/i.test(desc)) return 'Depósitos Javi';
  return 'Otros depósitos';
}

function calcularMesSueldo(fechaStr) {
  const d = new Date(fechaStr + 'T12:00:00');
  const y = d.getFullYear(), m = d.getMonth();
  const ultimoDia = ultimoDiaHabil(y, m);
  const diaFecha  = d.getDate();
  if (diaFecha >= ultimoDia) return (m === 11) ? {mes: 0, año: y + 1} : {mes: m + 1, año: y};
  return {mes: m, año: y};
}

function ultimoDiaHabil(year, month) {
  let d = new Date(year, month + 1, 0);
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() - 1);
  return d.getDate();
}

function agregarRegistro(sheet, record) {
  const ms = calcularMesSueldo(record.fecha);
  sheet.appendRow([
    record.id,
    record.fecha,
    record.descripcion,
    record.categoria,
    record.subcategoria,
    record.tipo === 'ingreso' ? record.monto : '',
    record.tipo === 'egreso'  ? record.monto : '',
    new Date(record.fecha + 'T12:00:00').getMonth() + 1,
    new Date(record.fecha + 'T12:00:00').getFullYear(),
    ms.mes + 1,
    ms.año
  ]);
}

// ══════════════════════════════════════════
//  TEST — ejecutar desde el editor para revisar
//  sin escribir nada en el Sheet
// ══════════════════════════════════════════

function testParsers() {
  const LABEL_NOMBRE = 'finanzas-procesado';
  const reglas = [
    { query: 'from:transferencias@bci.cl',                    parser: parsearBCI,        banco: 'BCI transferencia'    },
    { query: 'from:contacto@bci.cl',                          parser: parsearBCITarjeta, banco: 'BCI tarjeta débito'   },
    { query: 'from:serviciodetransferencias@bancochile.cl',   parser: parsearBancoChile, banco: 'Banco de Chile'       },
    { query: 'from:noreply@correo.bancoestado.cl',            parser: parsearBancoEstado,banco: 'BancoEstado'          },
    { query: 'from:mensajeria@santander.cl',                  parser: parsearSantander,  banco: 'Santander'            },
  ];

  var resultados = [];
  var ignorados  = [];

  reglas.forEach(function(regla) {
    // Busca los últimos 20 correos de cada banco (ya procesados o no)
    const threads = GmailApp.search(regla.query, 0, 20);
    threads.forEach(function(thread) {
      thread.getMessages().forEach(function(msg) {
        const body   = msg.getPlainBody();
        const fecha  = msg.getDate();
        const asunto = msg.getSubject();
        const record = regla.parser(body, fecha);
        if (record) {
          resultados.push({
            banco:      regla.banco,
            asunto:     asunto,
            tipo:       record.tipo,
            monto:      record.monto,
            fecha:      record.fecha,
            desc:       record.descripcion,
            categoria:  record.categoria,
            id:         record.id
          });
        } else {
          ignorados.push(regla.banco + ' | ' + asunto);
        }
      });
    });
  });

  Logger.log('══════════════════════════════════');
  Logger.log('REGISTROS QUE SE CREARÍAN (' + resultados.length + ')');
  Logger.log('══════════════════════════════════');
  resultados.forEach(function(r) {
    Logger.log(
      '[' + r.tipo.toUpperCase() + '] ' + r.fecha +
      '  $' + r.monto +
      '  | ' + r.desc +
      '  | Cat: ' + r.categoria +
      '  | Banco: ' + r.banco +
      '  | ID: ' + r.id
    );
  });

  Logger.log('');
  Logger.log('CORREOS IGNORADOS (' + ignorados.length + ')');
  ignorados.forEach(function(s) { Logger.log('  ' + s); });
}

// ══════════════════════════════════════════
//  Trigger — ejecutar UNA VEZ desde el editor
// ══════════════════════════════════════════

function configurarTrigger() {
  // Eliminar triggers anteriores de esta función
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'procesarCorreosBanco') ScriptApp.deleteTrigger(t);
  });
  // Crear trigger cada 15 minutos
  ScriptApp.newTrigger('procesarCorreosBanco')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Trigger creado: procesarCorreosBanco cada 15 minutos');
}

// ══════════════════════════════════════════
//  Helpers Sheet
// ══════════════════════════════════════════

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Registros') || ss.insertSheet('Registros');
}

function getBudgetsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Presupuestos') || ss.insertSheet('Presupuestos');
}

function addHeader(sheet) {
  sheet.appendRow(['ID','Fecha','Descripción','Categoría','Subcategoría',
                   'Ingreso','Egreso','Mes','Año','Mes Sueldo','Año Sueldo']);
}

function formatFecha(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString().split('T')[0];
  return String(val).split('T')[0];
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
