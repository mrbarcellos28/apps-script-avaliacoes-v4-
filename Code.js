/*******************************************************
 * AVALIAÇÕES — v4 (menu único + unificação robusta de nomes)
 * Autor: Klenda (Apps Script Dev)
 *******************************************************/

const CFG = {
  SHEET_RESPOSTAS: 'Respostas ao formulário 1',
  SHEET_MEDIAS_MEMBROS: 'Médias Individuais',
  SHEET_MEDIAS_PROJETOS: 'Médias por Projeto',
  SHEET_RESUMO: 'Resumo',
  TIMEZONE: 'America/Sao_Paulo',

  // Ativa o gatilho onFormSubmit automaticamente na 1ª execução
  AUTO_ENSURE_TRIGGER: true,

  // Padrões para capturar NOME em colunas do Forms (membros)
  MEMBER_COL_PATTERNS: [
    /Avalie\s+(?:o|a)?\s*(?:Diretor(?:a)?(?: de [^:]+)?|Diretoria(?: de [^:]+)?|Coordenador(?:a)?|Gerente|Consultor(?:a)?|Membro)\s+([^:]+?)\s*:/i,
    /Avalie\s+([^:]+?)\s*:/i // fallback genérico: "Avalie Fulano:"
  ],

  // Padrões para capturar PROJETO
  PROJECT_COL_PATTERNS: [
    /Avalie\s+(?:o|a)?\s*Projeto\s+([^:]+?)\s*:/i,
    /Avalie\s+o\s*projeto\s+([^:]+?)\s*:/i,
    /Projeto\s*:\s*([^:]+?)\s*$/i
  ],

  // Palavras que marcam colunas textuais (feedback etc.)
  EXCLUDE_KEYWORDS: ['feedback', 'coment', 'sugest', 'texto', 'por que', 'justific', 'descri', 'observa'],

  // Cabeçalhos padrão do Forms para ignorar
  EXCLUDE_EXACT_HEADERS: [
    'Carimbo de data/hora', 'Timestamp', 'Endereço de e-mail', 'E-mail',
    'Email address', 'Nome', 'Nome completo', 'Sobrenome'
  ],

  // Escala de cores para a média
  COLOR_SCALE: [
    { max: 3.0, color: '#FF6B6B' },   // vermelho
    { max: 4.0, color: '#FFD166' },   // amarelo
    { max: 10,  color: '#06D6A0' }    // verde
  ]
};

/* ========== MENU ÚNICO (seguro em headless) ========== */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Avaliações')
      .addItem('Atualizar médias', 'runAtualizarMedias')
      .addToUi();
    SpreadsheetApp.getActive().toast('Use: Avaliações → Atualizar médias', 'Pronto', 3);
  } catch (_) {
    // silencioso quando rodar por gatilho/headless
  }
}

// Botão único do menu
function runAtualizarMedias() {
  try {
    if (CFG.AUTO_ENSURE_TRIGGER) _ensureOnSubmitTriggerOnce();
    recalcularMedias();
  } catch (err) {
    Logger.log(err && err.stack || err);
    SpreadsheetApp.getActive().toast('Erro ao atualizar. Veja Executions → Logs.', 'Erro', 5);
  }
}

/* ========== GATILHO (silencioso) ========== */
function aoReceberNovaResposta(e) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_RESPOSTAS);
    if (!sh || sh.getLastRow() < 2) return;
    recalcularMedias();
  } catch (err) {
    Logger.log('onFormSubmit error: ' + (err && err.stack || err));
  }
}
function _ensureOnSubmitTriggerOnce() {
  const prop = PropertiesService.getDocumentProperties();
  const flag = prop.getProperty('onSubmitTriggerInstalled');
  if (flag === '1') return;

  const has = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'aoReceberNovaResposta');
  if (!has) {
    ScriptApp.newTrigger('aoReceberNovaResposta')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onFormSubmit()
      .create();
  }
  prop.setProperty('onSubmitTriggerInstalled', '1');
}

/* ========== CORE ========== */
function recalcularMedias() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_RESPOSTAS);
  if (!sh) { toast(`Aba "${CFG.SHEET_RESPOSTAS}" não encontrada.`); return; }

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) { toast('Sem dados suficientes.'); return; }

  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map(h => String(h || ''));

  // Classificar colunas relevantes
  const columns = classifyColumns(headers);

  // Acumular notas
  const membrosRaw = new Map();   // key -> {display,sum,count}
  const projetos   = new Map();   // key -> {display,sum,count}
  let globalSum = 0, globalCount = 0;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    for (let c = 0; c < headers.length; c++) {
      const meta = columns[c];
      if (!meta) continue; // ignorar colunas não reconhecidas

      const num = parseNumber(row[c]);
      if (num == null || num <= 0) continue; // ignora não numéricos/vazios

      if (meta.type === 'member') {
        accumulate(membrosRaw, meta.key, meta.display, num);
      } else if (meta.type === 'project') {
        accumulate(projetos, meta.key, meta.display, num);
      }
      globalSum += num;
      globalCount++;
    }
  }

  // === UNIFICAÇÃO ROBUSTA DE NOMES POR CHAVE CURTA ===
  const { merged: membros, mergesReport } = combineByShortKey(membrosRaw);

  // Escrever saídas
  writeMemberAverages(ss, membros);
  writeProjectAverages(ss, projetos);
  writeResumo(ss, globalSum, globalCount, mergesReport);

  toast('Médias atualizadas.');
}

/* ========== CLASSIFICAÇÃO DE COLUNAS ========== */
function classifyColumns(headers) {
  const out = new Array(headers.length).fill(null);

  for (let c = 0; c < headers.length; c++) {
    const raw = headers[c];
    const h = String(raw || '').trim();
    if (!h) continue;

    // Ignora metadados do Forms
    if (CFG.EXCLUDE_EXACT_HEADERS.some(x => eqIgnoreCase(x, h))) continue;

    // Ignora por palavra-chave textual
    const lower = h.toLowerCase();
    if (CFG.EXCLUDE_KEYWORDS.some(k => lower.includes(k))) continue;

    // MEMBER?
    for (const rx of CFG.MEMBER_COL_PATTERNS) {
      const m = h.match(rx);
      if (m && m[1]) {
        const nm = cleanDisplay(m[1]);
        const key = normalize(nm);
        if (key) { out[c] = { type: 'member', key, display: nm }; break; }
      }
    }
    if (out[c]) continue;

    // PROJECT?
    for (const rx of CFG.PROJECT_COL_PATTERNS) {
      const m = h.match(rx);
      if (m && m[1]) {
        const pj = cleanDisplay(m[1]);
        const key = normalize(pj);
        if (key) { out[c] = { type: 'project', key, display: pj }; break; }
      }
    }
    // Se não casou: permanece null (ignorar)
  }
  return out;
}

/* ========== ACÚMULO / UNIFICAÇÃO ========== */
function accumulate(map, key, display, val) {
  const cur = map.get(key) || { display, sum: 0, count: 0 };
  cur.sum += val; cur.count += 1;
  // mantém o display mais “completo” (maior)
  if ((display || '').length > (cur.display || '').length) cur.display = display;
  map.set(key, cur);
}

/**
 * Une entradas de membros por “shortKey”:
 * - remove artigos iniciais (a, o, as, os)
 * - remove acentos/pontuação
 * - remove stopwords (de, da, do, das, dos)
 * - usa PRIMEIRO + ÚLTIMO token como chave curta
 * Ex.: "a Manuella da Silva Padilha"  -> shortKey: "manuella padilha"
 *      "Manuella da Silva Padilha"    -> shortKey: "manuella padilha"  (mesma pessoa)
 *      "Thiago Athanasio"             -> shortKey: "thiago athanasio"
 *      "Thiago Athanasio Barreto ..." -> shortKey: "thiago athanasio"  (une)
 */
function combineByShortKey(mapIn) {
  const stop = new Set(['de','da','do','das','dos']);
  const groups = new Map(); // shortKey -> [{key,display,sum,count}]

  // 1) agrupa por shortKey
  for (const [key, obj] of mapIn.entries()) {
    const sk = deriveShortKey(obj.display, stop);
    const arr = groups.get(sk) || [];
    arr.push({ key, ...obj });
    groups.set(sk, arr);
  }

  // 2) mescla grupos
  const merged = new Map();              // newKey (shortKey) -> {display,sum,count}
  const mergesReport = [];               // para mostrar no Resumo
  for (const [shortKey, arr] of groups.entries()) {
    if (arr.length === 1) {
      const o = arr[0];
      merged.set(shortKey, { display: o.display, sum: o.sum, count: o.count });
    } else {
      // mescla vários aliases
      let sum = 0, count = 0, bestDisplay = '';
      const aliases = [];
      for (const o of arr) {
        sum += o.sum; count += o.count;
        if ((o.display || '').length > bestDisplay.length) bestDisplay = o.display;
        aliases.push(o.display);
      }
      merged.set(shortKey, { display: bestDisplay, sum, count });
      mergesReport.push({ shortKey, display: bestDisplay, aliases: aliases.sort() });
    }
  }
  return { merged, mergesReport };
}

function deriveShortKey(name, stopSet) {
  // tira artigos e acentos/pontuação
  let s = String(name || '')
    .replace(/^\s*(?:a|o|as|os)\s+/i, '')  // artigos no início
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[“”"’'`´]/g, '')
    .replace(/[^\w\s]/g, ' ')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();

  const tokens = s.split(' ').filter(t => t && !stopSet.has(t));
  if (tokens.length >= 2) return `${tokens[0]} ${tokens[tokens.length - 1]}`;
  return tokens.join(' '); // caso 1 token
}

/* ========== ESCRITA DE SAÍDA ========== */
function writeMemberAverages(ss, membros) {
  const sh = getOrCreateSheet(ss, CFG.SHEET_MEDIAS_MEMBROS);
  sh.clear();

  const rows = [['Nome Avaliado', 'Média', 'Qtd. notas']];
  const arr = [...membros.values()]
    .map(o => [o.display, safeAvg(o.sum, o.count), o.count])
    .sort((a, b) => b[1] - a[1]);
  rows.push(...(arr.length ? arr : [['—', 0, 0]]));

  const r = sh.getRange(1, 1, rows.length, rows[0].length);
  r.setValues(rows);
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('0.00');
  colorizeColumn(sh, 2, rows.length);
  autosize(sh);
  SpreadsheetApp.flush();
}

function writeProjectAverages(ss, projetos) {
  const sh = getOrCreateSheet(ss, CFG.SHEET_MEDIAS_PROJETOS);
  sh.clear();

  const rows = [['Projeto', 'Média', 'Qtd. notas']];
  const arr = [...projetos.values()]
    .map(o => [o.display, safeAvg(o.sum, o.count), o.count])
    .sort((a, b) => b[1] - a[1]);
  rows.push(...(arr.length ? arr : [['—', 0, 0]]));

  const r = sh.getRange(1, 1, rows.length, rows[0].length);
  r.setValues(rows);
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('0.00');
  colorizeColumn(sh, 2, rows.length);
  autosize(sh);
  SpreadsheetApp.flush();
}

function writeResumo(ss, globalSum, globalCount, mergesReport) {
  const sh = getOrCreateSheet(ss, CFG.SHEET_RESUMO);
  sh.clear();

  const mediaGlobal = safeAvg(globalSum, globalCount);

  const head = [
    ['Indicador', 'Valor'],
    ['Média global', mediaGlobal],
    ['Total de notas consideradas', globalCount],
    ['Atualizado em', Utilities.formatDate(new Date(), CFG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss')],
    ['','']
  ];

  // bloco de unificações detectadas
  const mergeHeader = [['Unificações de nomes (shortKey → display final → aliases)']];
  const mergeRows = (mergesReport && mergesReport.length)
    ? mergesReport.map(m => [ `${m.shortKey} → ${m.display} → ${m.aliases.join(' | ')}` ])
    : [['Nenhuma unificação necessária']];

  const all = head.concat(mergeHeader).concat(mergeRows);

  const r = sh.getRange(1, 1, all.length, 2);
  r.setValues(all.map(row => row.length === 1 ? [row[0], ''] : row));
  sh.getRange(2, 2, 1, 1).setNumberFormat('0.00');
  autosize(sh);
  SpreadsheetApp.flush();
}

/* ========== CORES / FORMATOS ========== */
function colorizeColumn(sh, colIdx, totalRows) {
  if (totalRows <= 1) return;
  const rg = sh.getRange(2, colIdx, totalRows - 1, 1);
  const vals = rg.getValues();
  for (let i = 0; i < vals.length; i++) {
    const v = Number(vals[i][0]);
    if (isNaN(v)) continue;
    rg.getCell(i + 1, 1).setBackground(pickColor(v));
  }
}
function pickColor(v) {
  for (const band of CFG.COLOR_SCALE) if (v < band.max) return band.color;
  return CFG.COLOR_SCALE[CFG.COLOR_SCALE.length - 1].color;
}

/* ========== UTILS ========== */
function toast(msg) { SpreadsheetApp.getActive().toast(msg, 'Avaliações', 5); }

function parseNumber(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).replace(/[^0-9,.\-]/g, '').replace(',', '.');
  if (!s) return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}
function normalize(str) {
  return String(str || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[“”"’'`´]/g, '')
    .replace(/\s+/g, ' ')
    .toLowerCase().trim();
}
// Remove aspas, espaços duplicados e pontuação final solta do display
function cleanDisplay(str) {
  return String(str || '')
    .replace(/[“”"’'`´]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/\s*[:;,.]\s*$/, '')
    .trim();
}

function eqIgnoreCase(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
function safeAvg(sum, cnt) { return cnt ? (sum / cnt) : 0; }

function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else if (sh.isSheetHidden && sh.isSheetHidden()) {
    sh.showSheet();
  }
  return sh;
}
function autosize(sh) {
  const n = Math.max(1, sh.getLastColumn());
  for (let c = 1; c <= n; c++) sh.autoResizeColumn(c);
}
