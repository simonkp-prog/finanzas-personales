function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
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
    if (sheet.getLastRow() <= 1) return json([]);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    const registros = rows
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
    return json(registros);
  } catch(e) {
    return json({error: e.message});
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Registros') || ss.insertSheet('Registros');
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
