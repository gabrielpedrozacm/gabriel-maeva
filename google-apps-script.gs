/**
 * Check-in do voo GM 1709 · Gabriel & Maeva
 * Recebe as confirmações do site e grava-as nesta folha de cálculo.
 * Se a mesma pessoa (mesmo nome) fizer o check-in outra vez, a linha é atualizada.
 */
const SHEET_NAME = 'Check-in';
// Opcional: cole aqui o ID da folha (a parte do link entre /d/ e /edit).
// Só é preciso se o script tiver sido criado fora da folha, em script.google.com.
const SHEET_ID = '';

function getBook_() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  // Script sem folha associada: cria uma no seu Google Drive e lembra-se dela.
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('BOOK_ID');
  if (saved) return SpreadsheetApp.openById(saved);
  const book = SpreadsheetApp.create('Check-in GM 1709 · Gabriel & Maeva');
  props.setProperty('BOOK_ID', book.getId());
  return book;
}
const HEADERS = ['Data/hora', 'Nome(s)', 'Presença', 'Lugares', 'Restrições alimentares', 'Pedido musical', 'Mensagem', 'Idioma'];

function clean_(v) {
  v = String(v == null ? '' : v).trim().slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // evita fórmulas injetadas
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getBook_();
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
    return ContentService.createTextOutput(JSON.stringify({ ok: true, sheet: ss.getUrl() })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Torre de controle GM 1709: online ✈');
}

// Para testar dentro do Apps Script: escolha "testeCheckIn" no menu e clique em Executar.
function testeCheckIn() {
  const out = doPost({ parameter: { nome: 'Teste torre', presenca: 'Sim', lugares: '1', idioma: 'PT' } });
  console.log(out.getContent());
}
