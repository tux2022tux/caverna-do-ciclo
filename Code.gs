/**
 * Mather — A Caverna do Ciclo · receptor de resultados
 * Cole este arquivo em Extensões → Apps Script da planilha e
 * implante como "App da Web" (Executar como: eu · Quem pode acessar: qualquer pessoa).
 */
const ABA = 'Partidas';
const CABECALHO = ['Data', 'Nome', 'Turma', 'Modo', 'Início', 'Cavernas', 'Pontos',
  'Acertos', 'Perguntas', '% acerto', 'Maior sequência', 'Resultado', 'Erros', 'Código', 'Válido'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const nome = String(d.nome || '').trim().slice(0, 40);
    const turma = normTurma_(d.turma);
    const modo = d.modo === 'D' ? 'D' : 'C';
    const n = (v, max) => Math.max(0, Math.min(max, Math.floor(Number(v) || 0)));
    const inicio = n(d.inicio, 5), cavernas = n(d.cavernas, 5), pontos = n(d.pontos, 999999);
    const acertos = n(d.acertos, 9999), total = n(d.total, 9999), pct = n(d.pct, 100), seq = n(d.seq, 9999);
    const codigo = String(d.codigo || '').slice(0, 40);
    const esperado = codigo_(nome, modo, inicio, cavernas, pontos, pct);
    const valido = codigo === esperado && total >= acertos;

    abaPartidas_().appendRow([new Date(), seguro_(nome), turma, modo === 'D' ? 'Desafio' : 'Calmo',
      inicio, cavernas, pontos, acertos, total, pct / 100, seq,
      d.resultado === 'vitoria' ? 'Escapou' : 'Perdeu', seguro_(String(d.erros || '').slice(0, 800)),
      codigo, valido ? '✔' : '✖']);

    return json_({ ok: true, valido: valido, ranking: ranking_(turma, modo) });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  return json_({ ok: true, ranking: ranking_(normTurma_(p.turma), p.modo === 'D' ? 'D' : 'C') });
}

/* ---------- auxiliares ---------- */
function abaPartidas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(CABECALHO);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold');
    sh.getRange('J:J').setNumberFormat('0%');
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  }
  return sh;
}

// Melhor pontuação de cada aluno na turma e no modo, só partidas válidas. Top 10.
function ranking_(turma, modo) {
  if (!turma) return [];
  const sh = abaPartidas_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, CABECALHO.length).getValues();
  const modoTxt = modo === 'D' ? 'Desafio' : 'Calmo';
  const best = {};
  rows.forEach(r => {
    if (r[2] !== turma || r[3] !== modoTxt || r[14] !== '✔') return;
    const nome = String(r[1]).replace(/^'/, ''), key = nome.toLowerCase(), pts = Number(r[6]);
    if (!best[key] || pts > best[key].pontos) best[key] = { nome: nome, pontos: pts, cavernas: Number(r[5]) };
  });
  return Object.values(best).sort((a, b) => b.pontos - a.pontos).slice(0, 10);
}

// Mesmo cálculo do jogo: M-<modo><início><cavernas>-<pontos>-<%>-<verificador>
function codigo_(nome, modo, inicio, cavernas, pontos, pct) {
  const b = modo + inicio + cavernas + '-' + pontos + '-' + pct;
  return 'M-' + b + '-' + hash_(nome.toLowerCase() + '|' + b);
}
function hash_(s) {
  let h = 7;
  for (const ch of s) h = (h * 31 + ch.charCodeAt(0)) % 46656;
  return h.toString(36).toUpperCase().padStart(3, '0');
}
function normTurma_(t) {
  return String(t || '').toUpperCase().normalize('NFD').replace(/[^0-9A-Z]/g, '').slice(0, 10);
}
// Evita que um nome comece com =, +, - ou @ e vire fórmula na planilha.
function seguro_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
