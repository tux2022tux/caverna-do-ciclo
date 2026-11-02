/**
 * A Caverna do Ciclo · servidor de turmas e resultados (v2)
 *
 * Instalação: Extensões → Apps Script → cole este arquivo → salvar.
 * Atualização sem trocar a URL: Implantar → Gerenciar implantações → ✏️ Editar → Versão: Nova versão → Implantar.
 *
 * Na planilha aparece o menu "🦇 Caverna do Ciclo":
 *   - Nova turma…  → gera o código que os alunos digitam no jogo
 *   - Preparar planilha → cria as abas Turmas, Partidas e Config
 * Para bloquear uma turma, desmarque a caixa "Ativa" na aba Turmas.
 */
const ABA_P = 'Partidas', ABA_T = 'Turmas', ABA_C = 'Config';
const CAB_P = ['Data', 'Nome', 'Turma', 'Modo', 'Início', 'Cavernas', 'Pontos',
  'Acertos', 'Perguntas', '% acerto', 'Maior sequência', 'Resultado', 'Erros', 'Código', 'Válido'];
const CAB_T = ['Turma', 'Código', 'Ativa', 'Criada em', 'Observação'];
const SESSAO_SEG = 6 * 60 * 60;      // sessão do aluno dura 6 h
const MAX_PARTIDAS_SESSAO = 40;      // limite de envios por sessão
const ALFABETO = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; // sem 0/O, 1/I/L

/* ---------- menu da planilha ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🦇 Caverna do Ciclo')
    .addItem('Nova turma…', 'novaTurma')
    .addItem('Preparar planilha', 'preparar')
    .addToUi();
}
function preparar() {
  abaPartidas_(); abaTurmas_(); abaConfig_();
  SpreadsheetApp.getUi().alert('Planilha pronta.\n\nUse "Nova turma…" para gerar o código de cada turma.');
}
function novaTurma() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Nova turma', 'Nome da turma (ex.: 3B, 2A-manhã):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const turma = normTurma_(r.getResponseText());
  if (!turma) { ui.alert('Nome de turma inválido.'); return; }
  const sh = abaTurmas_();
  const rows = linhas_(sh, CAB_T.length);
  if (rows.some(x => x[0] === turma)) { ui.alert('A turma ' + turma + ' já existe na aba Turmas.'); return; }
  const usados = new Set(rows.map(x => x[1]));
  let codigo; do { codigo = gerarCodigo_(); } while (usados.has(codigo));
  sh.appendRow([turma, codigo, true, new Date(), '']);
  sh.getRange(sh.getLastRow(), 3).insertCheckboxes().setValue(true);
  ui.alert('Turma ' + turma + ' criada.\n\nCódigo para os alunos:  ' + codigo +
    '\n\nPara bloquear o jogo dessa turma, desmarque "Ativa" na aba Turmas.');
}

/* ---------- API do jogo ---------- */
function doGet() {
  return json_({ ok: true, app: 'caverna-do-ciclo', versao: 2 });
}
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (_) { return json_({ ok: false, motivo: 'formato' }); }
  try {
    switch (d.acao) {
      case 'config': return json_({ ok: true, treinoLivre: config_().treino_livre !== false });
      case 'entrar': return json_(entrar_(d));
      case 'resultado': return json_(resultado_(d));
      case 'ranking': return json_(rankingSessao_(d));
      default: return json_({ ok: false, motivo: 'acao' });
    }
  } catch (err) {
    return json_({ ok: false, motivo: 'erro', erro: String(err) });
  }
}

function entrar_(d) {
  const codigo = String(d.codigo || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 12);
  const nome = limpaNome_(d.nome);
  if (!nome) return { ok: false, motivo: 'nome' };
  const t = turmaPorCodigo_(codigo);
  if (!t) return { ok: false, motivo: 'codigo' };
  if (!t.ativa) return { ok: false, motivo: 'bloqueada', turma: t.turma };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token, JSON.stringify({ turma: t.turma, nome: nome, codigo: codigo, n: 0 }), SESSAO_SEG);
  return { ok: true, token: token, turma: t.turma, nome: nome };
}

function sessao_(token) {
  const raw = CacheService.getScriptCache().get('s:' + String(token || ''));
  return raw ? JSON.parse(raw) : null;
}

function resultado_(d) {
  const s = sessao_(d.token);
  if (!s) return { ok: false, motivo: 'sessao' };
  const t = turmaPorCodigo_(s.codigo);
  if (!t || !t.ativa) return { ok: false, motivo: 'bloqueada' };
  if (s.n >= MAX_PARTIDAS_SESSAO) return { ok: false, motivo: 'limite' };
  s.n++;
  CacheService.getScriptCache().put('s:' + d.token, JSON.stringify(s), SESSAO_SEG);

  const n = (v, max) => Math.max(0, Math.min(max, Math.floor(Number(v) || 0)));
  const modo = d.modo === 'D' ? 'D' : 'C';
  const inicio = n(d.inicio, 5), cavernas = n(d.cavernas, 5), pontos = n(d.pontos, 999999);
  const acertos = n(d.acertos, 9999), total = n(d.total, 9999), pct = n(d.pct, 100), seq = n(d.seq, 9999);
  const codigo = String(d.codigo || '').slice(0, 40);

  let status = '✔';
  if (codigo !== codigo_(s.nome, modo, inicio, cavernas, pontos, pct)) status = '✖ código';
  else if (acertos > total || seq > acertos || pontos > teto_(acertos, cavernas) ||
           (total > 0 && Math.round(100 * acertos / total) !== pct) || inicio < 1) status = '⚠ suspeito';

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    abaPartidas_().appendRow([new Date(), seguro_(s.nome), t.turma, modo === 'D' ? 'Desafio' : 'Calmo',
      inicio, cavernas, pontos, acertos, total, pct / 100, seq,
      d.resultado === 'vitoria' ? 'Escapou' : 'Perdeu', seguro_(String(d.erros || '').slice(0, 800)),
      codigo, status]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, status: status, ranking: ranking_(t.turma, modo, s.nome) };
}

function rankingSessao_(d) {
  const s = sessao_(d.token);
  if (!s) return { ok: false, motivo: 'sessao' };
  return { ok: true, turma: s.turma, calmo: ranking_(s.turma, 'C', s.nome), desafio: ranking_(s.turma, 'D', s.nome) };
}

// Melhor pontuação de cada aluno (só partidas ✔). Top 10 + posição de quem pediu.
function ranking_(turma, modo, nome) {
  const modoTxt = modo === 'D' ? 'Desafio' : 'Calmo';
  const best = {};
  linhas_(abaPartidas_(), CAB_P.length).forEach(r => {
    if (r[2] !== turma || r[3] !== modoTxt || r[14] !== '✔') return;
    const nm = String(r[1]).replace(/^'/, ''), key = nm.toLowerCase(), pts = Number(r[6]);
    if (!best[key] || pts > best[key].pontos) best[key] = { nome: nm, pontos: pts };
  });
  const lista = Object.values(best).sort((a, b) => b.pontos - a.pontos);
  const pos = nome ? lista.findIndex(x => x.nome.toLowerCase() === String(nome).toLowerCase()) + 1 : 0;
  return { top: lista.slice(0, 10), pos: pos || null, n: lista.length };
}

/* ---------- regras ---------- */
// Pontuação máxima possível: 540 por acerto (100×4 de combo + 140 de tempo)
// e 1500 por caverna (chave, porta, cristais e vidas), com folga.
function teto_(acertos, cavernas) {
  return acertos * 540 + (cavernas + 1) * 1500;
}
// Mesmo cálculo do jogo: M-<modo><início><cavernas>-<pontos>-<%>-<verificador>
function codigo_(nome, modo, inicio, cavernas, pontos, pct) {
  const b = modo + inicio + cavernas + '-' + pontos + '-' + pct;
  return 'M-' + b + '-' + hash_(String(nome).toLowerCase() + '|' + b);
}
function hash_(s) {
  let h = 7;
  for (const ch of s) h = (h * 31 + ch.charCodeAt(0)) % 46656;
  return h.toString(36).toUpperCase().padStart(3, '0');
}

/* ---------- planilha ---------- */
function abaPartidas_() {
  return aba_(ABA_P, CAB_P, sh => {
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange('J:J').setNumberFormat('0%');
  });
}
function abaTurmas_() {
  return aba_(ABA_T, CAB_T, sh => {
    sh.getRange('B:B').setFontFamily('Roboto Mono').setFontWeight('bold');
    sh.getRange('D:D').setNumberFormat('dd/MM/yyyy');
  });
}
function abaConfig_() {
  return aba_(ABA_C, ['Opção', 'Valor', 'O que faz'], sh => {
    sh.appendRow(['treino_livre', true, 'Marcado: qualquer pessoa pode jogar sem código (modo treino, nada é enviado). Desmarcado: só entra com código de turma.']);
    sh.getRange(2, 2).insertCheckboxes().setValue(true);
  });
}
function aba_(nome, cab, extra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(cab);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cab.length).setFontWeight('bold');
    if (extra) extra(sh);
  }
  return sh;
}
function linhas_(sh, cols) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, cols).getValues();
}
function turmaPorCodigo_(codigo) {
  if (!codigo) return null;
  const r = linhas_(abaTurmas_(), 3).find(x => String(x[1]).toUpperCase() === codigo);
  return r ? { turma: String(r[0]), ativa: r[2] === true || String(r[2]).toUpperCase() === 'SIM' } : null;
}
function config_() {
  const o = {};
  linhas_(abaConfig_(), 2).forEach(r => { o[String(r[0])] = r[1]; });
  return o;
}

/* ---------- utilidades ---------- */
function gerarCodigo_() {
  let c = '';
  for (let i = 0; i < 6; i++) c += ALFABETO[Math.floor(Math.random() * ALFABETO.length)];
  return c;
}
function normTurma_(t) {
  return String(t || '').toUpperCase().normalize('NFD').replace(/[^0-9A-Z]/g, '').slice(0, 10);
}
function limpaNome_(s) {
  const n = String(s || '').replace(/[\u0000-\u001f<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, 40);
  return /\p{L}/u.test(n) ? n : '';
}
// Evita que um texto comece com =, +, - ou @ e vire fórmula na planilha.
function seguro_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
