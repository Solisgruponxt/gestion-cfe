/**
 * Gestión CFE — Apps Script Web App (solo API JSON)
 * Lee en vivo la pestaña "Gestion CFE" del libro "Req. Servicios | Ops"
 * y la expone como JSON para el dashboard publicado en GitHub Pages
 * (solisgruponxt.github.io/gestion-cfe).
 *
 * INSTALACIÓN:
 * 1) Abre el Google Sheet "Req. Servicios | Ops".
 * 2) Extensions > Apps Script.
 * 3) Borra el contenido de Code.gs (si lo hay) y pega este archivo completo.
 * 4) Guarda (nombra el proyecto, ej. "Gestion CFE API").
 * 5) Deploy > New deployment > tipo "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6) Copia la liga que termina en /exec y pégala como SCRIPT_URL en index.html.
 * 7) Si editas el código después, usa "Manage deployments" > editar (lápiz)
 *    > nueva versión, para que la MISMA liga /exec sirva el código actualizado.
 */

var SHEET_NAME = 'Gestion CFE';

function doGet(e) {
  var payload;
  try {
    payload = { ok: true, data: getData_() };
  } catch (err) {
    payload = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getData_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('No se encontró la pestaña "' + SHEET_NAME + '". Revisa que el nombre coincida exactamente.');
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { updatedAt: new Date().toISOString(), headers: [], rows: [] };
  }

  var rawHeaders = values[0].map(function (h) { return String(h).trim(); });

  var rawRows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var isEmpty = row.every(function (c) { return c === '' || c === null; });
    if (isEmpty) continue;
    rawRows.push(row);
  }

  // Se omiten columnas sin ningún dato en toda la hoja: reduce el peso de la
  // respuesta (y el tiempo de transferencia) sin perder ninguna columna real.
  var colIndexes = [];
  var headers = [];
  for (var c = 0; c < rawHeaders.length; c++) {
    var hasData = rawRows.some(function (row) { return row[c] !== '' && row[c] !== null && row[c] !== undefined; });
    if (hasData) {
      colIndexes.push(c);
      headers.push(rawHeaders[c] || ('col' + c));
    }
  }

  var rows = rawRows.map(function (row) {
    var obj = {};
    for (var i = 0; i < colIndexes.length; i++) {
      var v = row[colIndexes[i]];
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
      }
      obj[headers[i]] = v;
    }
    return obj;
  });

  return { updatedAt: new Date().toISOString(), headers: headers, rows: rows };
}
