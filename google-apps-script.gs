/**
 * Check-in do voo GM 1709 · Gabriel & Maeva
 * Recebe as confirmações do site e grava-as nesta folha de cálculo.
 * Se a mesma pessoa (mesmo nome) fizer o check-in outra vez, a linha é atualizada.
 */
const SHEET_NAME = 'Check-in';
const HEADERS = ['Data/hora', 'Nome(s)', 'Presença', 'Lugares', 'Restrições alimentares', 'Pedido musical', 'Mensagem', 'Idioma'];

function clean_(v) {
  v = String(v == null ? '' : v).trim().slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // evita fórmulas injetadas
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#5A1F2C').setFontColor('#F7EDE6');
    }
    const p = (e && e.parameter) || {};
    const row = [new Date(), clean_(p.nome), clean_(p.presenca), Number(p.lugares) || 0,
                 clean_(p.restricoes), clean_(p.musica), clean_(p.mensagem), clean_(p.idioma)];
    const last = sh.getLastRow();
    let target = 0;
    if (last > 1) {
      const names = sh.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase());
      const i = names.indexOf(String(p.nome || '').trim().toLowerCase());
      if (i >= 0) target = i + 2;
    }
    if (target) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Torre de controle GM 1709: online ✈');
}
