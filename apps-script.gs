// Tramear event form: saves each answer as a new row in this Google Sheet.
// Paste into Extensions > Apps Script of the sheet, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Copy the web app URL into the page.

const SHEET_NAME = 'Respostas';
const HEADERS = [
  'Data e hora', 'Nome', 'Cargo', 'Empresa', 'E-mail',
  'Principal conversa sobre IA', 'Momento da jornada de IA',
  'Principal desafio de liderança e pessoas', 'Temas para continuar a conversa',
  'Autoriza contato e conteúdos', 'Origem',
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = (e && e.parameter) || {};
    if (p.website) return json({ ok: true }); // bot filled the hidden field
    const sheet = getSheet();
    const temas = (e.parameters && e.parameters.temas) ? e.parameters.temas.join('; ') : '';
    sheet.appendRow([
      new Date(), clean(p.nome), clean(p.cargo), clean(p.empresa), clean(p.email),
      clean(p.conversa), clean(p.jornada), clean(p.desafio), clean(temas),
      p.autoriza === 'sim' ? 'Sim' : 'Não', clean(p.origem),
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, service: 'tramear-form' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

// Keeps text as text: a value starting with = + - @ would otherwise run as a formula.
function clean(v) {
  const s = String(v || '').trim().slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
