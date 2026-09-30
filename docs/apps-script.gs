// Referencia para pegar en Google Apps Script (script.google.com), vinculado
// a la hoja de cálculo donde se guardarán las confirmaciones. No se despliega
// desde este repo: copia este contenido al editor de Apps Script y publica
// el proyecto como Web App (Implementar > Nueva implementación > Aplicación
// web, acceso "Cualquier usuario"). Pega la URL /exec resultante en
// SHEETS_URL dentro de index.html.
//
// Encabezados esperados en la hoja "Confirmaciones" (fila 1):
// Fecha | Invitado | Asistirá | Personas

const SHEET_NAME = 'Confirmaciones';

// SINCRONIZAR con index.html
const GUEST_LIST = [
  "Jorge García",
  "Andrea García y Alejandro Flores",
  "Hugo García",
  "Rebeca Rivera",
  "Romina Rivera",
  "Jorge Armando García",
  "Victoria Contreras",
  "Julieta Contreras",
  "Victor Delgado",
  "Victor Contreras",
  "Claudia León",
  "Ximena Contreras",
  "Fernanda Contreras y Óscar Reyes",
  "Mariana Contreras",
  "Julián Mendoza",
  "Pina Contreras",
  "Gabriela Ruiz y Rubén",
  "Martha de Haro y Eoin Vaughan",
  "Julio Morales",
  "Andrea Amezcua",
  "Leslie Briseño y Miguel Ángel Bello",
  "Susana Bourde",
  "Daniel Peña",
  "Miguel Hernández y Lulú Soto",
  "Lepoldo Hernández",
  "Marco Hernández y Laura Roa",
  "Carolina Hernández",
  "Leticia Bonifacio",
  "Rosa Bonifacio",
  "Isabella Paz y Efraín Marbán",
  "Laura Facio",
  "Ana Baez y Dylan",
  "María Fernanda Maya",
  "Celina Maya"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);

    const invitado = clean(d.invitado);
    if (!GUEST_LIST.includes(invitado)) {
      throw new Error('Invitado no reconocido');
    }

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // Columna B = "Invitado" (Fecha | Invitado | Asistirá | Personas)
    const columnaInvitados = sh.getRange(2, 2, Math.max(sh.getLastRow() - 1, 0), 1).getValues().flat();
    const yaConfirmado = columnaInvitados.some(v => String(v).trim().toLowerCase() === invitado.toLowerCase());
    if (yaConfirmado) {
      return json({ ok: false, code: 'DUPLICADO', error: 'Ya habíamos registrado la confirmación de ' + invitado + '. Si necesitas corregirla, contacta directamente a Lorena y Julio.' });
    }

    const asiste = d.attendance === true;
    const esPareja = /\sy\s/.test(invitado);
    const personas = asiste ? (esPareja ? 2 : 1) : 0;

    const fecha = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy HH:mm');

    sh.appendRow([
      fecha,
      invitado,
      asiste ? 'Sí' : 'No',
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
