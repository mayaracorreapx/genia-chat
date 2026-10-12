/**
 * GenIA — recebe as respostas do front e grava na planilha "GenIA - Respostas".
 *
 * Ações (corpo JSON enviado pelo front):
 *   { action: "start",  email }            -> cria a linha, devolve { ok, id }
 *   { action: "answer", id, pergunta }     -> completa a linha, devolve { ok }
 */
const SHEET_ID = '1LtD3_J7cwWoZNltOr1Mp_3aJ3P-4S5Zj7z5yDNr8lbA';
const HEADERS = ['Data/Hora', 'E-mail', 'Pergunta para a GenIA', 'Status', 'ID'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet_();

    if (body.action === 'start') {
      const email = String(body.email || '').trim();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false, error: 'E-mail inválido' });
      const id = Utilities.getUuid();
      sheet.appendRow([new Date(), email, '', 'Aguardando pergunta', id]);
      return json_({ ok: true, id: id });
    }

    if (body.action === 'answer') {
      const pergunta = String(body.pergunta || '').trim();
      if (!pergunta) return json_({ ok: false, error: 'Pergunta vazia' });
      const row = findRow_(sheet, body.id);
      if (!row) return json_({ ok: false, error: 'Registro não encontrado' });
      sheet.getRange(row, 3, 1, 2).setValues([[pergunta, 'Recebida']]);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'Ação desconhecida' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Abrir a URL do app no navegador mostra se está no ar.
function doGet() {
  return json_({ ok: true, service: 'GenIA' });
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (current.join('|') !== HEADERS.join('|')) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRow_(sheet, id) {
  if (!id) return null;
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const ids = sheet.getRange(2, 5, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === id) return i + 2;
  }
  return null;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
