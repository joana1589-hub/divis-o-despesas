/**
 * Finanças Joana & Toni — cópia automática da app no Google Sheets.
 *
 * Lê os dados do Firebase (projeto "divisao-despesas") e reescreve os
 * separadores desta folha. Só de leitura: o que se mudar aqui NÃO volta para a app
 * e é substituído na atualização seguinte.
 *
 * Primeira vez: no menu "💰 Finanças" → "⏱️ Ativar atualização automática".
 */

const PROJECT_ID = 'divisao-despesas';
const API_KEY = 'AIzaSyCAO6gr87Mu3XRDtN2mGl7GjsutZ0Qw9bQ';   // chave pública da app (a mesma do index.html)
const MINUTOS = 5;
const EUR = '#,##0.00 "€"';
const DATA_FMT = 'dd/mm/yyyy';

const CATS = [
  ['supermercado', '🛒', 'Supermercado'], ['restaurante', '🍽️', 'Restaurante'], ['casa', '🏠', 'Casa'], ['quinta', '🏡', 'Quinta Guarda'],
  ['transporte', '🚗', 'Transporte'], ['saude', '💊', 'Saúde'], ['lazer', '🎬', 'Lazer'], ['roupa', '👗', 'Roupa'],
  ['viagem', '✈️', 'Viagem'], ['animal', '🐾', 'Animal'], ['tecnologia', '💻', 'Tecnologia'],
  ['servicos', '📱', 'Serviços'], ['outro', '📦', 'Outro']
].map(([id, icon, label]) => ({ id, icon, label }));
const DEFAULT_SUBS = {
  casa: ['Empréstimo', 'Energia e água', 'Comunicações', 'Limpeza', 'Seguros', 'Manutenção', 'Mobília / Decoração'],
  quinta: ['Empréstimo', 'Seguros', 'Impostos', 'Manutenção'],
  supermercado: ['Compras do mês', 'Compras pequenas'],
  restaurante: ['Jantar fora', 'Take-away', 'Café / Lanche'],
  transporte: ['Combustível', 'Portagens', 'Manutenção do carro', 'Estacionamento'],
  saude: ['Farmácia', 'Consultas / Exames'],
  lazer: ['Saídas', 'Desporto', 'Subscrições']
};
const WED_CATS = {
  espaco: 'Quinta / Espaço', catering: 'Catering / Bolo', foto: 'Fotografia / Vídeo', roupa: 'Vestido / Fato',
  musica: 'Música / DJ', flores: 'Flores / Decoração', aliancas: 'Alianças', convites: 'Convites / Lembranças',
  beleza: 'Cabelo / Maquilhagem', igreja: 'Igreja / Registo', lua: 'Lua de mel', outro: 'Outro'
};
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// ═════════════════════════ MENU E ATUALIZAÇÃO AUTOMÁTICA ═════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('💰 Finanças')
    .addItem('🔄 Atualizar agora', 'atualizar')
    .addItem('⏱️ Ativar atualização automática (5 min)', 'ativarAutomatico')
    .addItem('⏹️ Parar atualização automática', 'pararAutomatico')
    .addToUi();
}

function ativarAutomatico() {
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', SpreadsheetApp.getActive().getId());
  pararAutomatico(true);
  ScriptApp.newTrigger('atualizar').timeBased().everyMinutes(MINUTOS).create();
  atualizar();
  SpreadsheetApp.getActive().toast(`A folha passa a atualizar sozinha a cada ${MINUTOS} minutos.`, '💰 Finanças', 6);
}

function pararAutomatico(silencioso) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'atualizar')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (silencioso !== true) SpreadsheetApp.getActive().toast('Atualização automática parada.', '💰 Finanças', 5);
}

function folha_() {
  const ss = SpreadsheetApp.getActive();
  if (ss) return ss;
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
}

function atualizar() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const dados = lerFirebase_();
    escreverTudo_(folha_(), dados, new Date());
  } catch (e) {
    console.error(e);
    try {
      const sh = folha_().getSheetByName('Resumo') || folha_().insertSheet('Resumo', 0);
      sh.getRange('A2').setValue('⚠️ Erro na última atualização (' + Utilities.formatDate(new Date(), 'Europe/Lisbon', 'dd/MM HH:mm') + '): ' + e.message);
    } catch (e2) { /* nada */ }
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ═════════════════════════ LER DO FIREBASE ═════════════════════════

const BASE = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents`;

function lerFirebase_() {
  return {
    registos: listar_('registos'),
    obra: listar_('obra'),
    casamento: listar_('casamento', true),
    config: obterDoc_('settings/config') || {}
  };
}

// Autorização: primeiro como dona do projeto (Google Cloud); se não der, entra no
// Firebase com a conta Google (e aí aplicam-se as regras do Firestore).
function cabecalhos_() {
  const modo = CacheService.getScriptCache().get('AUTH_MODE');
  if (modo === 'firebase') return { Authorization: 'Bearer ' + tokenFirebase_() };
  return { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Goog-User-Project': PROJECT_ID };
}

function tokenFirebase_() {
  const cache = CacheService.getScriptCache();
  const t = cache.get('FB_ID_TOKEN');
  if (t) return t;
  const r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:signInWithIdp?key=' + API_KEY, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({
      postBody: 'access_token=' + ScriptApp.getOAuthToken() + '&providerId=google.com',
      requestUri: 'http://localhost', returnSecureToken: true, returnIdpCredential: false
    })
  });
  if (r.getResponseCode() !== 200) throw new Error('Não foi possível entrar no Firebase: ' + r.getContentText().slice(0, 300));
  const id = JSON.parse(r.getContentText()).idToken;
  cache.put('FB_ID_TOKEN', id, 50 * 60);
  return id;
}

function pedir_(url) {
  let r = UrlFetchApp.fetch(url, { headers: cabecalhos_(), muteHttpExceptions: true });
  const code = r.getResponseCode();
  if ((code === 401 || code === 403) && CacheService.getScriptCache().get('AUTH_MODE') !== 'firebase') {
    CacheService.getScriptCache().put('AUTH_MODE', 'firebase', 6 * 3600);
    r = UrlFetchApp.fetch(url, { headers: cabecalhos_(), muteHttpExceptions: true });
  }
  return r;
}

function listar_(colecao, opcional) {
  const out = [];
  let token = '';
  do {
    const url = `${BASE}/${colecao}?pageSize=300` + (token ? '&pageToken=' + encodeURIComponent(token) : '');
    const r = pedir_(url);
    if (r.getResponseCode() !== 200) {
      if (opcional) return out;
      throw new Error(`Erro a ler "${colecao}" (${r.getResponseCode()}): ` + r.getContentText().slice(0, 300));
    }
    const j = JSON.parse(r.getContentText());
    (j.documents || []).forEach(d => out.push(Object.assign({ id: d.name.split('/').pop() }, campos_(d.fields || {}))));
    token = j.nextPageToken || '';
  } while (token);
  return out;
}

function obterDoc_(caminho) {
  const r = pedir_(`${BASE}/${caminho}`);
  if (r.getResponseCode() !== 200) return null;
  return campos_(JSON.parse(r.getContentText()).fields || {});
}

function campos_(f) {
  const o = {};
  Object.keys(f).forEach(k => o[k] = valor_(f[k]));
  return o;
}
function valor_(v) {
  if (v == null) return null;
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return Number(v.doubleValue);
  if ('booleanValue' in v) return v.booleanValue;
  if ('nullValue' in v) return null;
  if ('timestampValue' in v) return v.timestampValue;
  if ('mapValue' in v) return campos_(v.mapValue.fields || {});
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(valor_);
  return null;
}

// ═════════════════════════ CONTAS (iguais às da app) ═════════════════════════

const num = v => Number(v) || 0;
const r2 = v => Math.round(num(v) * 100) / 100;
const outro = p => p === 'Joana' ? 'Toni' : 'Joana';
const isP = n => n === 'Joana' || n === 'Toni';
const normText = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
const chave = s => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
const slug = t => normText(t).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 30) || 'sub';
const catLabel = id => (CATS.find(c => c.id === id) || { label: id || '' }).label;
const para = r => r.para || outro(r.quem);

function subsFor_(cfg, cat) {
  const s = (cfg.subcats || {})[cat];
  if (Array.isArray(s)) return s;
  return (DEFAULT_SUBS[cat] || []).map(l => ({ id: slug(l), label: l }));
}
function subOf_(cfg, r) {
  if (!r.cat) return '';
  const map = (cfg.descSub || {})[r.cat] || {};
  const k = chave(r.desc);
  const id = Object.prototype.hasOwnProperty.call(map, k) ? map[k] : (r.sub || null);
  const s = id && subsFor_(cfg, r.cat).find(x => x.id === id);
  return s ? s.label : '';
}

function contasCasa_(regs, cfg) {
  let tj = 0, tt = 0, pj = 0, pt = 0;
  regs.forEach(r => {
    if (r.type === 'EXP') { if (r.quem === 'Joana') tj += num(r.val); else tt += num(r.val); }
    else if (r.type === 'PAY') {
      if (r.quem === 'Joana' && para(r) === 'Toni') pj += num(r.val);
      if (r.quem === 'Toni' && para(r) === 'Joana') pt += num(r.val);
    }
  });
  const aj = cfg.dev25 === 'Toni' ? num(cfg.d25) : -num(cfg.d25);
  const saldo = (tj / 2 - tt / 2) + aj - pt + pj;   // > 0: Toni deve à Joana
  return { tj, tt, pj, pt, aj, saldo };
}

function wedNorm_(r) {
  const v = num(r.val);
  const metodo = r.metodo === 'dinheiro' || r.metodo === 'entradas' ? 'dinheiro' : 'banco';
  const pessoa = isP(r.quem) ? r.quem : 'Joana';
  if (r.kind === 'receita') return { kind: 'receita', v, metodo, conta: metodo === 'banco' ? pessoa : null };
  if (r.kind === 'despesa') {
    if (metodo === 'banco') return { kind: 'despesa', v, metodo, fonte: 'conta', quem: pessoa };
    if (r.fonte === 'prendas' || r.metodo === 'entradas') return { kind: 'despesa', v, metodo, fonte: 'prendas', quem: 'Conjunto' };
    return { kind: 'despesa', v, metodo, fonte: 'proprio', quem: pessoa };
  }
  if (r.kind === 'retirada') return { kind: 'retirada', v, quem: pessoa };
  if (r.kind === 'acerto') return { kind: 'acerto', v, quem: pessoa, para: r.para || outro(r.quem) };
  return { kind: 'outro', v };
}
function wedMetodo_(r) {
  const n = wedNorm_(r);
  if (n.kind === 'receita') return n.metodo === 'dinheiro' ? 'Dinheiro (junto)' : 'MB conta ' + n.conta;
  if (n.kind === 'despesa') return n.fonte === 'conta' ? 'MB conta ' + n.quem : n.fonte === 'prendas' ? 'Dinheiro das prendas' : `Dinheiro do bolso (${n.quem})`;
  if (n.kind === 'retirada') return 'Tirou do dinheiro junto';
  return 'Transferência';
}

// Igual à app, com a opção "recuperar primeiro das prendas na própria conta".
function contasCasamento_(rows) {
  const P = {};
  ['Joana', 'Toni'].forEach(n => P[n] = { mbIn: 0, mbOut: 0, ownDin: 0, ret: 0, adj: 0 });
  const T = { recDin: 0, recBan: 0, nRec: 0, despBan: 0, despPDin: 0, despOwnDin: 0, cashOut: 0 };
  rows.forEach(r => {
    const n = wedNorm_(r);
    if (n.kind === 'receita') { T.nRec++; if (n.metodo === 'dinheiro') T.recDin += n.v; else { T.recBan += n.v; P[n.conta].mbIn += n.v; } }
    else if (n.kind === 'despesa') {
      if (n.fonte === 'conta') { T.despBan += n.v; P[n.quem].mbOut += n.v; }
      else if (n.fonte === 'prendas') T.despPDin += n.v;
      else { T.despOwnDin += n.v; P[n.quem].ownDin += n.v; }
    } else if (n.kind === 'retirada') { P[n.quem].ret += n.v; T.cashOut += n.v; }
    else if (n.kind === 'acerto') { P[n.quem].adj += n.v; P[n.para].adj -= n.v; }
  });
  const rec = T.recDin + T.recBan, desp = T.despBan + T.despPDin + T.despOwnDin;
  const cash = T.recDin - T.despPDin - T.cashOut;
  let fromCashTot = 0, conj = 0;
  Object.keys(P).forEach(n => {
    const x = P[n];
    x.B = x.mbIn - x.mbOut; x.surplus = Math.max(x.B, 0); x.deficit = Math.max(-x.B, 0);
    x.pend = x.ownDin + x.deficit - x.ret + x.adj;
    const need = Math.max(x.pend, 0);
    x.fromAcc = Math.min(need, x.surplus); x.fromCash = need - x.fromAcc; x.toJoint = x.surplus - x.fromAcc;
    fromCashTot += x.fromCash; conj += x.toJoint;
  });
  return { P, T, rec, desp, lucro: rec - desp, cash, endCash: cash - fromCashTot, conj };
}

// ═════════════════════════ ESCREVER NA FOLHA ═════════════════════════

const dataDe = s => {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || '');
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : (s || '');
};
const quando = iso => iso ? Utilities.formatDate(new Date(iso), 'Europe/Lisbon', 'dd/MM/yyyy HH:mm') : '';
const porData = (a, b) => String(b.date || '').localeCompare(String(a.date || ''));

function escreverTudo_(ss, d, agora) {
  const cfg = d.config || {};
  // Os registos eliminados na app ficam arquivados (deleted: true, ou SKIP com os dados
  // originais nas recorrentes). Saem dos separadores de cada área, mas ficam no "Tudo".
  const vivo = r => r.date && !r.deleted;
  const regs = d.registos.filter(vivo);
  const exp = regs.filter(r => r.type === 'EXP').sort(porData);
  const pay = regs.filter(r => r.type === 'PAY').sort(porData);
  const obra = d.obra.filter(vivo).sort(porData);
  const wed = d.casamento.filter(vivo).sort(porData);
  const elimRegs = d.registos.filter(r => r.date && (
    (r.deleted && (r.type === 'EXP' || r.type === 'PAY')) || (r.type === 'SKIP' && r.origType && r.desc)));
  const elimObra = d.obra.filter(r => r.date && r.deleted);
  const elimWed = d.casamento.filter(r => r.date && r.deleted);

  // ── Casa ──
  tabela_(ss, 'Casa', ['Data', 'Categoria', 'Subcategoria', 'Descrição', 'Quem pagou', 'Valor (€)', 'Metade de cada (€)', 'Recorrente', 'Nota', 'Registado por', 'Registado em'],
    exp.map(r => [dataDe(r.date), catLabel(r.cat), subOf_(cfg, r), r.desc || '', r.quem || '', r2(r.val), r2(num(r.val) / 2),
      r.recur ? 'Sim' : '', r.nota || '', r.createdBy || '', quando(r.createdAt)]),
    { money: [6, 7], dates: [1], widths: [90, 110, 160, 180, 90, 90, 110, 80, 200, 100, 120] });

  // ── Acertos ──
  tabela_(ss, 'Acertos', ['Data', 'De', 'Para', 'Valor (€)', 'Nota', 'Registado por', 'Registado em'],
    pay.map(r => [dataDe(r.date), r.quem, para(r), r2(r.val), r.nota || '', r.createdBy || '', quando(r.createdAt)]),
    { money: [4], dates: [1], widths: [90, 70, 70, 90, 220, 100, 120] });

  // ── Quinta Guarda ──
  tabela_(ss, 'Quinta Guarda', ['Data', 'Descrição', 'Quem pagou', 'Valor (€)', 'Link', 'Registado por', 'Registado em'],
    obra.map(r => [dataDe(r.date), r.desc || '', r.quem || '', r2(r.val), r.link || '', r.createdBy || '', quando(r.createdAt)]),
    { money: [4], dates: [1], widths: [90, 220, 100, 90, 220, 100, 120] });

  // ── Casamento ──
  tabela_(ss, 'Casamento', ['Data', 'Tipo', 'Convidado / Descrição', 'Categoria', 'Quem / Conta', 'Método', 'Valor (€)', 'Nota', 'Registado por', 'Registado em'],
    wed.map(r => {
      const n = wedNorm_(r);
      return [dataDe(r.date),
        n.kind === 'receita' ? 'Prenda' : n.kind === 'despesa' ? 'Despesa' : n.kind === 'retirada' ? 'Retirada' : 'Acerto',
        n.kind === 'retirada' ? `${n.quem} tirou do dinheiro junto` : n.kind === 'acerto' ? `Acerto ${n.quem} → ${n.para}` : (r.desc || ''),
        n.kind === 'despesa' ? (WED_CATS[r.cat] || 'Outro') : '',
        n.kind === 'receita' ? (n.conta || 'Dinheiro junto') : (n.kind === 'despesa' && n.fonte === 'prendas' ? 'Dinheiro junto' : n.quem),
        wedMetodo_(r), r2(r.val), r.nota || '', r.createdBy || '', quando(r.createdAt)];
    }),
    { money: [7], dates: [1], widths: [90, 80, 220, 140, 110, 190, 90, 200, 100, 120] });

  const prendas = wed.filter(r => r.kind === 'receita')
    .sort((a, b) => String(a.desc || '').localeCompare(String(b.desc || ''), 'pt'));
  tabela_(ss, 'Prendas', ['Convidado', 'Valor (€)', 'Como', 'Data', 'Nota'],
    prendas.map(r => [r.desc || '', r2(r.val), wedMetodo_(r), dataDe(r.date), r.nota || '']),
    { money: [2], dates: [4], widths: [240, 90, 150, 90, 200] });

  // ── Tudo (com histórico do que foi eliminado) ──
  const TUDO_CAB = ['Data', 'Área', 'Movimento', 'Descrição', 'Categoria', 'Quem', 'Para', 'Método', 'Valor (€)',
    'Estado', 'Registado por', 'Registado em', 'Eliminado por', 'Eliminado em', 'ID'];
  const estado = (r, elim) => elim ? ['Eliminado', r.createdBy || '', quando(r.createdAt), r.deletedBy || '', quando(r.deletedAt)]
                                   : ['Ativo', r.createdBy || '', quando(r.createdAt), '', ''];
  const tudo = [];
  const casaRow = (r, elim) => {
    const t = r.type === 'SKIP' ? r.origType : r.type;
    if (t === 'PAY') tudo.push([dataDe(r.date), 'Acerto Casa', 'Acerto', r.nota || `Acerto ${r.quem} → ${para(r)}`, '', r.quem, para(r), '', r2(r.val)]
      .concat(estado(r, elim), ['registos/' + r.id]));
    else tudo.push([dataDe(r.date), 'Casa', 'Saída', r.desc || '', [catLabel(r.cat), subOf_(cfg, r)].filter(Boolean).join(' · '), r.quem, '', '', r2(r.val)]
      .concat(estado(r, elim), ['registos/' + r.id]));
  };
  const obraRow = (r, elim) => tudo.push([dataDe(r.date), 'Quinta Guarda', 'Saída', r.desc || '', '', r.quem, '', '', r2(r.val)]
    .concat(estado(r, elim), ['obra/' + r.id]));
  const wedRow = (r, elim) => {
    const n = wedNorm_(r);
    tudo.push([dataDe(r.date), 'Casamento', n.kind === 'receita' ? 'Entrada' : n.kind === 'despesa' ? 'Saída' : 'Acerto',
      n.kind === 'retirada' ? `${n.quem} tirou do dinheiro junto` : (r.desc || ''), n.kind === 'despesa' ? (WED_CATS[r.cat] || '') : '',
      n.kind === 'receita' ? (n.conta || 'Conjunto') : n.quem, n.para || '', wedMetodo_(r), r2(r.val)]
      .concat(estado(r, elim), ['casamento/' + r.id]));
  };
  exp.forEach(r => casaRow(r, false)); pay.forEach(r => casaRow(r, false)); elimRegs.forEach(r => casaRow(r, true));
  obra.forEach(r => obraRow(r, false)); elimObra.forEach(r => obraRow(r, true));
  wed.forEach(r => wedRow(r, false)); elimWed.forEach(r => wedRow(r, true));

  // O que já estava no "Tudo" e desapareceu do Firebase (apagado de vez) fica como eliminado.
  const vistos = new Set(tudo.map(t => t[14]));
  anteriores_(ss, 'Tudo').forEach(o => {
    const id = o['ID'];
    if (!id || vistos.has(id)) return;
    vistos.add(id);
    tudo.push(TUDO_CAB.map(h => {
      if (h === 'Estado') return 'Eliminado';
      if (h === 'Eliminado em') return o[h] || ('detetado ' + Utilities.formatDate(agora, 'Europe/Lisbon', 'dd/MM/yyyy HH:mm'));
      return o[h] === undefined ? '' : o[h];
    }));
  });
  const iso = v => v instanceof Date ? Utilities.formatDate(v, 'Europe/Lisbon', 'yyyy-MM-dd') : String(v || '');
  tudo.sort((a, b) => iso(b[0]).localeCompare(iso(a[0])));
  tabela_(ss, 'Tudo', TUDO_CAB, tudo,
    { money: [9], dates: [1], widths: [90, 110, 80, 200, 200, 80, 70, 170, 90, 80, 100, 120, 100, 130, 60] });
  const shT = ss.getSheetByName('Tudo');
  if (shT.setConditionalFormatRules && tudo.length) {
    const rg = shT.getRange(2, 1, tudo.length, TUDO_CAB.length);
    shT.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2="Eliminado"').setFontColor('#a0a4b0').setStrikethrough(true).setRanges([rg]).build()]);
  }

  // ── Resumo ──
  resumo_(ss, regs, exp, pay, obra, wed, cfg, agora);

  // Ordem dos separadores (só mexe se estiver diferente)
  const ordem = ['Resumo', 'Casa', 'Acertos', 'Quinta Guarda', 'Casamento', 'Prendas', 'Tudo'];
  const atual = ss.getSheets().map(x => x.getName()).filter(n => ordem.includes(n));
  if (atual.join('|') !== ordem.join('|')) {
    ordem.forEach((n, i) => { const sh = ss.getSheetByName(n); if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); } });
    ss.setActiveSheet(ss.getSheetByName('Resumo'));
  }
  ['Folha1', 'Página1', 'Sheet1'].forEach(nm => {
    const f = ss.getSheetByName(nm);
    if (f && f.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(f);
  });
}

// Lê as linhas atuais de um separador como objetos { cabeçalho: valor }.
function anteriores_(ss, nome) {
  const sh = ss.getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const cab = v[0].map(String);
  if (cab.indexOf('ID') < 0) return [];
  return v.slice(1).map(row => { const o = {}; cab.forEach((h, i) => o[h] = row[i]); return o; });
}

function obterFolha_(ss, nome) {
  return ss.getSheetByName(nome) || ss.insertSheet(nome);
}

function tabela_(ss, nome, cab, linhas, o) {
  const sh = obterFolha_(ss, nome);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clearContents();
  const dados = [cab].concat(linhas.length ? linhas : [cab.map((_, i) => i === 0 ? 'Sem registos' : '')]);
  sh.getRange(1, 1, dados.length, cab.length).setValues(dados);
  sh.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#1a1d23').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (linhas.length) {
    (o.money || []).forEach(c => sh.getRange(2, c, linhas.length, 1).setNumberFormat(EUR));
    (o.dates || []).forEach(c => sh.getRange(2, c, linhas.length, 1).setNumberFormat(DATA_FMT));
    sh.getRange(1, 1, linhas.length + 1, cab.length).createFilter();
  }
  const props = PropertiesService.getDocumentProperties();
  if (o.widths && !props.getProperty('W_' + nome)) {   // só define larguras na 1ª vez (respeita ajustes vossos)
    o.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    props.setProperty('W_' + nome, '1');
  }
}

function resumo_(ss, regs, exp, pay, obra, wed, cfg, agora) {
  const sh = obterFolha_(ss, 'Resumo');
  sh.clearContents(); sh.clearFormats();
  const L = [];          // linhas [a, b, c, d, e]
  const fmt = [];        // [linha, tipo]
  const titulo = t => { L.push([t, '', '', '', '']); fmt.push([L.length, 'titulo']); };
  const cab = (...c) => { L.push(c.concat(Array(5 - c.length).fill(''))); fmt.push([L.length, 'cab']); };
  const lin = (...c) => { L.push(c.concat(Array(5 - c.length).fill(''))); fmt.push([L.length, 'eur']); };
  const vazio = () => L.push(['', '', '', '', '']);

  L.push(['💰 Finanças Joana & Toni', '', '', '', '']); fmt.push([1, 'grande']);
  L.push(['Atualizado em ' + Utilities.formatDate(agora, 'Europe/Lisbon', 'dd/MM/yyyy HH:mm') + ' · cópia só de leitura: os registos fazem-se na app', '', '', '', '']);
  vazio();

  // Casa
  const c = contasCasa_(regs, cfg);
  titulo('🏠 CASA — balanço');
  lin(Math.abs(c.saldo) < 0.05 ? '✅ Contas em dia' : c.saldo > 0 ? '🔴 Toni deve à Joana' : '🔵 Joana deve ao Toni', r2(Math.abs(c.saldo)));
  lin('Joana pagou (sempre)', r2(c.tj));
  lin('Toni pagou (sempre)', r2(c.tt));
  lin('Acertos Toni → Joana', r2(c.pt));
  lin('Acertos Joana → Toni', r2(c.pj));
  if (num(cfg.d25)) lin(`Ajuste manual (${cfg.dev25} deve)`, r2(num(cfg.d25)));
  vazio();

  // Por mês
  titulo('📅 CASA — por mês');
  cab('Mês', 'Total (€)', 'Joana (€)', 'Toni (€)', 'Balanço do mês');
  const meses = {};
  exp.forEach(r => {
    const k = r.date.slice(0, 7);
    const m = meses[k] = meses[k] || { t: 0, j: 0, o: 0 };
    m.t += num(r.val); if (r.quem === 'Joana') m.j += num(r.val); else m.o += num(r.val);
  });
  Object.keys(meses).sort().reverse().forEach(k => {
    const m = meses[k], b = (m.j - m.o) / 2;
    lin(`${MESES[Number(k.slice(5, 7)) - 1]} ${k.slice(0, 4)}`, r2(m.t), r2(m.j), r2(m.o),
      Math.abs(b) < 0.05 ? 'equilibrado' : (b > 0 ? 'Toni deve ' : 'Joana deve ') + Math.abs(b).toFixed(2) + ' €');
  });
  vazio();

  // Por categoria e subcategoria (por ano)
  titulo('🧾 CASA — por categoria e subcategoria');
  cab('Ano', 'Categoria', 'Subcategoria', 'Total (€)', 'Nº');
  const grupos = {};
  exp.forEach(r => {
    const k = [r.date.slice(0, 4), catLabel(r.cat || 'outro'), subOf_(cfg, r) || '(sem subcategoria)'].join('|');
    const g = grupos[k] = grupos[k] || { t: 0, n: 0 }; g.t += num(r.val); g.n++;
  });
  Object.keys(grupos).sort((a, b) => {
    const [ya] = a.split('|'), [yb] = b.split('|');
    return yb.localeCompare(ya) || grupos[b].t - grupos[a].t;
  }).forEach(k => { const [y, ca, su] = k.split('|'); lin(y, ca, su, r2(grupos[k].t), grupos[k].n); });
  vazio();

  // Quinta
  const q = { Joana: 0, Toni: 0, Conjunto: 0 };
  obra.forEach(r => { q[r.quem] = (q[r.quem] || 0) + num(r.val); });
  const qd = (q.Joana - q.Toni) / 2;
  titulo('🏡 QUINTA GUARDA');
  lin('Total investido', r2(q.Joana + q.Toni + q.Conjunto));
  lin('Pago pela Joana', r2(q.Joana));
  lin('Pago pelo Toni', r2(q.Toni));
  lin('Pago em conjunto', r2(q.Conjunto));
  lin(Math.abs(qd) < 0.01 ? 'Balanço: em equilíbrio' : qd > 0 ? 'Balanço: Toni deve à Joana' : 'Balanço: Joana deve ao Toni', r2(Math.abs(qd)));
  vazio();

  // Casamento
  if (wed.length) {
    const w = contasCasamento_(wed);
    titulo('💍 CASAMENTO');
    lin(w.lucro >= 0 ? 'Lucro' : 'Prejuízo', r2(Math.abs(w.lucro)));
    lin('Prendas recebidas', r2(w.rec));
    lin('Despesas', r2(w.desp));
    lin('💵 Dinheiro em mão agora', r2(w.cash));
    lin('🏦 Prendas que sobram na conta da Joana', r2(w.P.Joana.B));
    lin('🏦 Prendas que sobram na conta do Toni', r2(w.P.Toni.B));
    ['Joana', 'Toni'].forEach(n => {
      if (w.P[n].fromCash > 0.004) lin(`${n} ainda tira do dinheiro em mão`, r2(w.P[n].fromCash));
    });
    lin('🏁 No fim fica em mão', r2(w.endCash));
    lin('🏁 Joana transfere para a conjunta', r2(w.P.Joana.toJoint));
    lin('🏁 Toni transfere para a conjunta', r2(w.P.Toni.toJoint));
    lin('🏁 Total na conta conjunta', r2(w.conj));
  }

  sh.getRange(1, 1, L.length, 5).setValues(L);
  fmt.forEach(([row, tipo]) => {
    const rg = sh.getRange(row, 1, 1, 5);
    if (tipo === 'grande') rg.setFontSize(16).setFontWeight('bold');
    else if (tipo === 'titulo') rg.setFontWeight('bold').setBackground('#f4f5f8');
    else if (tipo === 'cab') rg.setFontWeight('bold').setFontColor('#7a7f8e');
    else if (tipo === 'eur') sh.getRange(row, 2, 1, 4).setNumberFormat(EUR);
  });
  // A coluna "Nº" e "Ano" não são euros
  L.forEach((l, i) => {
    if (/^\d{4}$/.test(String(l[0]))) { sh.getRange(i + 1, 5).setNumberFormat('0'); sh.getRange(i + 1, 2, 1, 2).setNumberFormat('@'); }
  });
  sh.setColumnWidth(1, 300); [2, 3, 4, 5].forEach(c2 => sh.setColumnWidth(c2, 130));
}
