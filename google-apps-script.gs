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

function key_(v) {
  return String(v == null ? '' : v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

// Aba "Playlist": Música | Votos | Ocultar. Marque "Ocultar" para tirar uma música do site.
function playlistSheet_(ss) {
  let sh = ss.getSheetByName('Playlist');
  if (sh) return sh;
  sh = ss.insertSheet('Playlist');
  sh.appendRow(['Música', 'Votos', 'Ocultar']);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#5A1F2C').setFontColor('#F7EDE6');
  sh.getRange('C2:C1000').insertCheckboxes();
  const ci = ss.getSheetByName(SHEET_NAME);
  if (ci && ci.getLastRow() > 1) {
    ci.getRange(2, 6, ci.getLastRow() - 1, 1).getValues().forEach(r => addSong_(sh, r[0]));
  }
  return sh;
}

function findSong_(sh, k) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (key_(vals[i][0]) === k) return i + 2;
  return 0;
}

function addSong_(sh, song) {
  song = clean_(song);
  const k = key_(song);
  if (!k || findSong_(sh, k)) return;
  sh.appendRow([song, 0, false]);
}

function clean_(v) {
  v = String(v == null ? '' : v).trim().slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // evita fórmulas injetadas
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p0 = (e && e.parameter) || {};
    if (p0.tipo === 'aviao') {
      const props = PropertiesService.getScriptProperties();
      const n = Math.max(1, Math.min(20, Number(p0.n) || 1));
      const total = (Number(props.getProperty('AVIOES')) || 0) + n;
      props.setProperty('AVIOES', String(total));
      return ContentService.createTextOutput(JSON.stringify({ ok: true, avioes: total })).setMimeType(ContentService.MimeType.JSON);
    }
    if (p0.tipo === 'voto') {
      const pl = playlistSheet_(getBook_());
      const r = findSong_(pl, key_(p0.musica));
      if (!r) return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);
      const v = (Number(pl.getRange(r, 2).getValue()) || 0) + 1;
      pl.getRange(r, 2).setValue(v);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, votos: v })).setMimeType(ContentService.MimeType.JSON);
    }
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
    if (p.musica) addSong_(playlistSheet_(ss), p.musica);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, sheet: ss.getUrl() })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  // ?playlist=1 devolve só as músicas e os votos (sem nomes), sem as marcadas como "Ocultar"
  if (e && e.parameter && e.parameter.playlist) {
    const pl = playlistSheet_(getBook_());
    let list = [];
    if (pl.getLastRow() > 1) {
      list = pl.getRange(2, 1, pl.getLastRow() - 1, 3).getValues()
        .filter(r => String(r[0]).trim() && r[2] !== true)
        .map(r => ({ m: String(r[0]).replace(/^'/, ''), v: Number(r[1]) || 0 }))
        .sort((a, b) => b.v - a.v)
        .slice(0, 60);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, musicas: list })).setMimeType(ContentService.MimeType.JSON);
  }
  // ?stats=1 devolve só totais (nunca nomes nem mensagens) para o contador do site
  if (e && e.parameter && e.parameter.stats) {
    let respostas = 0, passageiros = 0;
    const sh = getBook_().getSheetByName(SHEET_NAME);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 3, sh.getLastRow() - 1, 2).getValues().forEach(r => {
        if (String(r[0]).trim() === 'Sim') { respostas++; passageiros += Number(r[1]) || 0; }
      });
    }
    const avioes = Number(PropertiesService.getScriptProperties().getProperty('AVIOES')) || 0;
    return ContentService.createTextOutput(JSON.stringify({ ok: true, respostas: respostas, passageiros: passageiros, avioes: avioes }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Torre de controle GM 1709: online ✈');
}

// Para testar dentro do Apps Script: escolha "testeCheckIn" no menu e clique em Executar.
function testeCheckIn() {
  const out = doPost({ parameter: { nome: 'Teste torre', presenca: 'Sim', lugares: '1', idioma: 'PT' } });
  console.log(out.getContent());
}
