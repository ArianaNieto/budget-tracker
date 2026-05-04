// ─────────────────────────────────────────────────────────────
//  Budget Tracker — Google Apps Script
//  Pegá TODO este código en script.google.com
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = "Gastos";
const HEADERS    = ["ID", "Fecha", "Descripción", "Categoría", "Monto", "Moneda", "Timestamp"];

function doPost(e) {
  try {
    const data = e.parameter;
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet   = ss.getSheetByName(SHEET_NAME);

    // Crea la hoja si no existe
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      formatHeaders(sheet);
    }

    // Si la hoja existe pero no tiene encabezados, los agrega
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      formatHeaders(sheet);
    }

    // Agrega la fila de gasto
    sheet.appendRow([
      data.id,
      data.date,
      data.desc,
      data.cat,
      data.amount,
      data.currency,
      new Date().toLocaleString("es-AR"),
    ]);

    // Auto-resize columnas
    sheet.autoResizeColumns(1, HEADERS.length);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
}

function formatHeaders(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground("#7c6af7");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
}

// Función de prueba — correla manualmente para verificar que funciona
function testInsert() {
  doPost({
    postData: {
      contents: JSON.stringify({
        id: Date.now(),
        date: "2026-05-04",
        desc: "Prueba de conexión",
        cat: "Otro",
        amount: 1000,
        currency: "ARS",
      })
    }
  });
}
