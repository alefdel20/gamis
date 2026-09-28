// Referencia para pegar en Google Apps Script (script.google.com), vinculado
// a la hoja de cálculo donde se guardarán las confirmaciones. No se despliega
// desde este repo: copia este contenido al editor de Apps Script y publica
// el proyecto como Web App (Implementar > Nueva implementación > Aplicación
// web, acceso "Cualquier usuario"). Pega la URL /exec resultante en
// SHEETS_URL dentro de index.html.
//
// Encabezados esperados en la hoja "Confirmaciones" (fila 1):
// Fecha | Nombre | Asistirá | Lleva acompañante | Acompañante | Personas

const SHEET_NAME = 'Confirmaciones';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);

    const nombre = clean(d.fullname);
    if (!nombre) throw new Error('Nombre vacío');

    const asiste = d.attendance === true;
    const conAcomp = asiste && d.companion === true;
    const personas = asiste ? (conAcomp ? 2 : 1) : 0;

    const fecha = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy HH:mm');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    sh.appendRow([
      fecha,
      nombre,
      asiste ? 'Sí' : 'No',
      conAcomp ? 'Sí' : 'No',
      conAcomp ? clean(d.companionName) : '',
      personas
    ]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function clean(s) {
  return String(s || '').trim().slice(0, 120).replace(/^[=+\-@]/, "'$&");
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
