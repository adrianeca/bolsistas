// ============================================================
//  BRASAS Bolsistas — Code.gs
// ============================================================

const USERS_SHEET_ID          = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';
const BOLSISTAS_SHEET_ID      = '1G7RfHP_8j7-6VPqC8ReYvpScyWvBnVyRXu8QAdeL4bQ';
const FUNCIONARIOS_SHEET_ID   = '1BDiPjv0FqRJp5EwcvLdYXVvEAWesvwdEgbhYdnTlqPY';

// Logos: (1) Suba "Brasas logo.png" no Drive, copie o ID da URL e cole aqui.
// (2) Logos dos parceiros ficam na pasta abaixo com nome = Origem da Bolsa (qualquer extensão).
const BRASAS_LOGO_FILE_ID      = '1fqZbnxHJNyov_9NwhwDAg235FizcnTGQ';
const PARTNER_LOGOS_FOLDER_ID  = '1sKrz_-odKjx6YNVCSWpB5eW9v3atEvQD';

const EDIT_ROLES       = ['admin', 'secretaria', 'diretor', 'b2b'];
const EMAIL_ROLES      = ['admin', 'b2b'];
const ADMIN_ROLES      = ['admin'];

// ─── Entry point ────────────────────────────────────────────
function doGet(e) {
  const p    = (e && e.parameter) ? e.parameter : {};
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.sessionToken = p.s || '';
  return tmpl.evaluate()
    .setTitle('BRASAS Bolsistas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Auth ───────────────────────────────────────────────────
function validateSession(token) {
  try {
    const user = _getUser(token);
    return JSON.stringify({ ok: true, ...user });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

function _getUser(token) {
  if (!token) throw new Error('Sessão não encontrada. Acesse pelo hub.');

  // Cache do objeto de usuário por 10 min (evita ler SESSOES a cada chamada)
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'user_' + token;
  const cached   = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const ss        = SpreadsheetApp.openById(USERS_SHEET_ID);
  const sessSheet = ss.getSheetByName('SESSOES');
  if (!sessSheet) throw new Error('Sessão inválida.');

  const now  = new Date();
  const rows = sessSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(token)) continue;
    const expira = rows[i][6] ? new Date(rows[i][6]) : null;
    if (expira && expira < now) throw new Error('Sessão expirada. Acesse pelo hub novamente.');

    const role  = String(rows[i][3]).trim().toLowerCase();
    const email = String(rows[i][1]).trim().toLowerCase();

    if (!_hasAccess(ss, role, email)) {
      throw new Error('Sem permissão para acessar o painel de bolsistas.');
    }

    const unidade = _getUserUnidade(ss, email);

    const userObj = {
      email,
      nome:         String(rows[i][2]).trim(),
      role,
      unidade,
      canEdit:      EDIT_ROLES.includes(role),
      canSendEmail: EMAIL_ROLES.includes(role),
      isAdmin:      ADMIN_ROLES.includes(role),
    };

    try { cache.put(cacheKey, JSON.stringify(userObj), 600); } catch(e) {}
    return userObj;
  }

  throw new Error('Sessão não encontrada. Acesse pelo hub.');
}

function _getUserUnidade(ss, email) {
  const norm  = s => String(s || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const sheet = ss.getSheetByName('USUARIOS');
  if (!sheet) return '';
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (norm(d[i][0]) !== norm(email)) continue;
    return String(d[i][4] || '').trim(); // Col E = unidade
  }
  return '';
}

function _hasAccess(ss, role, email) {
  const norm = s => String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');

  const usuSheet = ss.getSheetByName('USUARIOS');
  if (!usuSheet) return false;

  const d = usuSheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (norm(d[i][0]) !== norm(email)) continue;
    // Colunas F (índice 5) e G (índice 6) = acessos_dashboards
    const acessos = [d[i][5], d[i][6]]
      .join(',')
      .split(',')
      .map(s => norm(s));
    return acessos.includes('bolsistas');
  }

  return false;
}

// ─── Init unificado (filtros + dados, com cache) ─────────────
function initApp(token) {
  try {
    const user  = _getUser(token); // usa cache de usuário
    const cache    = CacheService.getScriptCache();
    const dataKey  = 'appdata_v3_' + token;
    const cached   = cache.get(dataKey);
    if (cached) return cached; // JSON já formatado, retorna direto

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    if (!sheet) return JSON.stringify({ ok: false, error: 'Aba "Bolsistas App" não encontrada.' });

    const feriadosSheet = ss.getSheetByName('Feriados');
    const feriados = [];
    if (feriadosSheet) {
      const fData = feriadosSheet.getDataRange().getValues();
      for (let i = 1; i < fData.length; i++) {
        const d = fData[i][0];
        if (!d) continue;
        if (d instanceof Date) {
          const dd = String(d.getDate()).padStart(2, '0');
          const mm = String(d.getMonth() + 1).padStart(2, '0');
          feriados.push(`${dd}/${mm}/${d.getFullYear()}`);
        } else {
          const s = String(d).trim();
          if (s) feriados.push(s);
        }
      }
    }

    // ── Funcionários (RJ - UNIDADES1): unidades e nomes por unidade ──
    const funcionarios = { unidades: [], porUnidade: {} };
    const funcSheet = SpreadsheetApp.openById(FUNCIONARIOS_SHEET_ID).getSheetByName('RJ - UNIDADES');
    if (funcSheet) {
      const fRows = funcSheet.getDataRange().getValues();
      const unidadesSet = new Set();
      for (let i = 1; i < fRows.length; i++) {
        const fr      = fRows[i];
        const nome    = String(fr[2]  || '').trim();   // C
        const status  = String(fr[10] || '').trim();   // K
        const unit1   = String(fr[21] || '').trim();   // V
        const unit2   = String(fr[30] || '').trim();   // AE
        if (!nome || status.toLowerCase() !== 'ativo') continue;
        [unit1, unit2].filter(Boolean).forEach(u => {
          unidadesSet.add(u);
          if (!funcionarios.porUnidade[u]) funcionarios.porUnidade[u] = [];
          if (!funcionarios.porUnidade[u].includes(nome)) funcionarios.porUnidade[u].push(nome);
        });
      }
      funcionarios.unidades = [...unidadesSet].sort();
      Object.keys(funcionarios.porUnidade).forEach(u => funcionarios.porUnidade[u].sort());
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({
      ok: true, user, rows: [], anos: [], meses: [], unidades: [], origens: [],
    });

    const norm = s => String(s || '').trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');

    const userUnidades = (user.isAdmin || !user.unidade)
      ? null
      : user.unidade.split(/[,|]/).map(s => norm(s.trim())).filter(Boolean);

    const rows = [];
    const meses = new Set(), anos = new Set(), unidades = new Set(), origens = new Set();

    for (let i = 1; i < data.length; i++) {
      const r       = data[i];
      const unidade = String(r[2] || '').trim();
      const mes     = String(r[3] || '').trim();
      const ano     = String(r[4] || '').trim();

      if (!unidade && !mes) continue;
      if (userUnidades && !userUnidades.includes(norm(unidade))) continue;

      if (mes)     meses.add(mes);
      if (ano)     anos.add(ano);
      if (unidade) unidades.add(unidade);
      if (r[7])    origens.add(String(r[7]).trim());

      rows.push({
        rowIndex:          i + 1,
        timestamp:         _fmtDate(r[0]),
        emailSecretaria:   String(r[1]  || ''),
        unidade,
        mes,
        ano,
        nome:              String(r[5]  || ''),
        percentual:        r[6]  !== '' ? r[6]  : '',
        origemBolsa:       String(r[7]  || ''),
        data1aAula:        _fmtDate(r[8]),
        book:              String(r[9]  || ''),
        frequencia:        String(r[10] || ''),
        dataInicioBook:    _fmtDate(r[11]),
        dataPrevConclusao: _fmtDate(r[12]),
        horario:           String(r[13] || ''),
        diasPrevistos:     r[14] !== '' ? r[14] : '',
        diasAssistidos:    r[15] !== '' ? r[15] : '',
        valor:             r[16] !== '' ? r[16] : '',
        mt:                r[17] !== '' ? r[17] : '',
        wt:                r[18] !== '' ? r[18] : '',
        oc:                r[19] !== '' ? r[19] : '',
        ot:                r[20] !== '' ? r[20] : '',
        aproveitamento:    String(r[21] || ''),
        observacoes:       String(r[22] || ''),
        turma:             String(r[23] || ''),
        chavePDF:          String(r[24] || ''),
        dependenteUnidade: String(r[25] || ''),
        dependenteNome:    String(r[26] || ''),
        obsExtrasEmail:    String(r[27] || ''),
        statusAluno:       String(r[28] || ''),
        diasAula:          r[29] !== '' ? r[29] : '',
        teste1:            r[30] !== '' ? r[30] : '',
        teste2:            r[31] !== '' ? r[31] : '',
      });
    }

    const result = JSON.stringify({
      ok:       true,
      rows,
      feriados,
      funcionarios,
      anos:     [...anos].sort().reverse(),
      meses:    [...meses],
      unidades: [...unidades].sort(),
      origens:  [...origens].sort(),
    });

    // Cacheia por 5 min (ignora se payload for grande demais — limite 100KB)
    try { cache.put(dataKey, result, 300); } catch(e) {}

    return result;
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Leitura de dados ────────────────────────────────────────
function getBolsistasData(token, paramsJson) {
  try {
    const user   = _getUser(token);
    const params = JSON.parse(paramsJson || '{}');

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    if (!sheet) return JSON.stringify({ ok: false, error: 'Aba "Bolsistas App" não encontrada.' });

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ ok: true, rows: [] });

    const norm = s => String(s || '').trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');
    const _inArr = (v, a) => !a || !a.length || (Array.isArray(a) ? a : [a]).map(norm).includes(norm(String(v)));

    // Unidades permitidas para o usuário
    const userUnidades = (user.isAdmin || !user.unidade)
      ? null
      : user.unidade.split(/[,|]/).map(s => norm(s.trim())).filter(Boolean);

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r       = data[i];
      const unidade = String(r[2] || '').trim();
      const mes     = String(r[3] || '').trim();
      const ano     = String(r[4] || '').trim();

      if (!unidade && !mes) continue;

      if (userUnidades && !userUnidades.includes(norm(unidade))) continue;
      if (!_inArr(mes,               params.mes))         continue;
      if (!_inArr(ano,               params.ano))         continue;
      if (!_inArr(unidade,           params.unidade))     continue;
      if (!_inArr(String(r[7]||''), params.origemBolsa)) continue;
      if (params.busca) {
        const q = norm(params.busca);
        if (!norm(String(r[5]||'')).includes(q) && !norm(String(r[23]||'')).includes(q)) continue;
      }

      rows.push({
        rowIndex:          i + 1,
        timestamp:         _fmtDate(r[0]),
        emailSecretaria:   String(r[1]  || ''),
        unidade,
        mes,
        ano,
        nome:              String(r[5]  || ''),
        percentual:        r[6]  !== '' ? r[6]  : '',
        origemBolsa:       String(r[7]  || ''),
        data1aAula:        _fmtDate(r[8]),
        book:              String(r[9]  || ''),
        frequencia:        String(r[10] || ''),
        dataInicioBook:    _fmtDate(r[11]),
        dataPrevConclusao: _fmtDate(r[12]),
        horario:           String(r[13] || ''),
        diasPrevistos:     r[14] !== '' ? r[14] : '',
        diasAssistidos:    r[15] !== '' ? r[15] : '',
        valor:             r[16] !== '' ? r[16] : '',
        mt:                r[17] !== '' ? r[17] : '',
        wt:                r[18] !== '' ? r[18] : '',
        oc:                r[19] !== '' ? r[19] : '',
        ot:                r[20] !== '' ? r[20] : '',
        aproveitamento:    String(r[21] || ''),
        observacoes:       String(r[22] || ''),
        turma:             String(r[23] || ''),
        chavePDF:          String(r[24] || ''),
        dependenteUnidade: String(r[25] || ''),
        dependenteNome:    String(r[26] || ''),
        obsExtrasEmail:    String(r[27] || ''),
        statusAluno:       String(r[28] || ''),
        diasAula:          r[29] !== '' ? r[29] : '',
        teste1:            r[30] !== '' ? r[30] : '',
        teste2:            r[31] !== '' ? r[31] : '',
      });
    }

    return JSON.stringify({ ok: true, rows });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

function getFilterOptions(token) {
  try {
    const user = _getUser(token);
    const ss   = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    if (!sheet) return JSON.stringify({ ok: false, error: 'Aba não encontrada.' });

    const data = sheet.getDataRange().getValues();
    const norm = s => String(s || '').trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');

    const userUnidades = (user.isAdmin || !user.unidade)
      ? null
      : user.unidade.split(/[,|]/).map(s => norm(s.trim())).filter(Boolean);

    const meses = new Set(), anos = new Set(), unidades = new Set(), origens = new Set();

    for (let i = 1; i < data.length; i++) {
      const r       = data[i];
      const unidade = String(r[2] || '').trim();
      if (!unidade) continue;
      if (userUnidades && !userUnidades.includes(norm(unidade))) continue;

      if (r[3]) meses.add(String(r[3]).trim());
      if (r[4]) anos.add(String(r[4]).trim());
      unidades.add(unidade);
      if (r[7]) origens.add(String(r[7]).trim());
    }

    return JSON.stringify({
      ok:       true,
      meses:    [...meses],
      anos:     [...anos].sort().reverse(),
      unidades: [...unidades].sort(),
      origens:  [...origens].sort(),
    });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Edição de dados ─────────────────────────────────────────
function updateBolsista(token, payloadJson) {
  try {
    const user = _getUser(token);
    if (!user.canEdit) throw new Error('Sem permissão para editar.');

    const { rowIndex, field, value } = JSON.parse(payloadJson);

    const FIELD_COL = {
      unidade:           3,
      mes:               4,
      ano:               5,
      nome:              6,
      percentual:        7,
      origemBolsa:       8,
      data1aAula:        9,
      book:              10,
      frequencia:        11,
      dataInicioBook:    12,
      dataPrevConclusao: 13,
      horario:           14,
      diasPrevistos:     15,
      diasAssistidos:    16,
      valor:             17,
      mt:                18,
      wt:                19,
      oc:                20,
      ot:                21,
      aproveitamento:    22,
      observacoes:       23,
      turma:             24,
      dependenteUnidade: 26,
      dependenteNome:    27,
      obsExtrasEmail:    28,
      statusAluno:       29,
      diasAula:          30,
      teste1:            31,
      teste2:            32,
    };

    const col = FIELD_COL[field];
    if (!col) throw new Error('Campo não editável: ' + field);

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');

    // Valida acesso à unidade do row para não-admin
    if (!user.isAdmin && user.unidade) {
      const rowData = sheet.getRange(rowIndex, 3, 1, 1).getValue();
      const norm    = s => String(s || '').trim().toLowerCase()
        .normalize('NFD').replace(/[̀-ͯ]/g, '');
      const permitidas = user.unidade.split(/[,|]/).map(s => norm(s.trim()));
      if (!permitidas.includes(norm(String(rowData)))) {
        throw new Error('Sem acesso a esta unidade.');
      }
    }

    sheet.getRange(rowIndex, col).setValue(value);
    sheet.getRange(rowIndex, 1).setValue(new Date());
    sheet.getRange(rowIndex, 2).setValue(user.email);

    // Invalida cache de dados para que o próximo carregamento reflita a edição
    try { CacheService.getScriptCache().remove('appdata_v3_' + token); } catch(e) {}

    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Ciclo Mensal ────────────────────────────────────────────
function runCicloMensal(token) {
  try {
    const user = _getUser(token);
    if (!user.isAdmin) throw new Error('Apenas administradores podem executar o ciclo mensal.');

    const log = [];
    log.push(_backupMesAnterior());
    log.push(_calcularDiasUteis());
    log.push(_copiarConsolidado());
    log.push(_formatarHorarios());

    return JSON.stringify({ ok: true, log });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

function _backupMesAnterior() {
  const ss        = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
  const appSheet  = ss.getSheetByName('Bolsistas App');
  const destSheet = ss.getSheetByName('db_bolsistas_mes_anterior');
  if (!appSheet)  throw new Error('"Bolsistas App" não encontrada.');
  if (!destSheet) throw new Error('"db_bolsistas_mes_anterior" não encontrada.');

  const data = appSheet.getDataRange().getValues();
  if (data.length < 2) return 'Backup: nenhum dado para copiar.';

  // Pega o último mês/ano presente (última combinação distinta encontrada)
  let lastMes = '', lastAno = '';
  for (let i = 1; i < data.length; i++) {
    const mes = String(data[i][3] || '').trim();
    const ano = String(data[i][4] || '').trim();
    if (mes && ano) { lastMes = mes; lastAno = ano; }
  }
  if (!lastMes) return 'Backup: nenhum dado com mês/ano.';

  const rowsToBackup = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3] || '').trim() === lastMes && String(data[i][4] || '').trim() === lastAno) {
      rowsToBackup.push(data[i].slice(2, 23)); // cols C:W
    }
  }

  destSheet.clearContents();
  if (rowsToBackup.length > 0) {
    destSheet.getRange(1, 1, rowsToBackup.length, rowsToBackup[0].length).setValues(rowsToBackup);
  }

  return `Backup: ${rowsToBackup.length} linhas de ${lastMes} ${lastAno} → db_bolsistas_mes_anterior.`;
}

function _calcularDiasUteis() {
  const ss            = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
  const sheet         = ss.getSheetByName('bolsistas_consolidado');
  const feriadosSheet = ss.getSheetByName('Feriados');
  if (!sheet) throw new Error('"bolsistas_consolidado" não encontrada.');

  const feriados = feriadosSheet
    ? feriadosSheet.getRange('A1:A').getValues().flat().filter(d => d)
        .map(d => Utilities.formatDate(new Date(d), 'GMT-3', 'dd/MM'))
    : [];

  const data = sheet.getDataRange().getValues();
  const ano  = new Date().getFullYear();
  const diasSemana = {
    'domingo': 0, 'segunda': 1, 'terça': 2, 'quarta': 3,
    'quinta': 4, 'sexta': 5, 'sábado': 6,
  };

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][26] || '').trim() !== 'Sim') continue; // col AA

    const mesTexto    = String(data[i][1]  || '').trim(); // col B
    const diasAulaStr = String(data[i][24] || '').trim(); // col Y
    if (!mesTexto || !diasAulaStr) continue;

    const mesNum = parseInt(mesTexto.split(' ')[0], 10);
    if (isNaN(mesNum)) continue;

    const ultimoDia    = new Date(ano, mesNum, 0).getDate();
    const diasAulaNums = diasAulaStr.split(',')
      .map(s => diasSemana[s.toLowerCase().trim()])
      .filter(n => n !== undefined);

    let dias = 0;
    for (let d = 1; d <= ultimoDia; d++) {
      const dt  = new Date(ano, mesNum - 1, d);
      const fmt = Utilities.formatDate(dt, 'GMT-3', 'dd/MM');
      if (diasAulaNums.includes(dt.getDay()) && !feriados.includes(fmt)) dias++;
    }

    sheet.getRange(i + 1, 13).setValue(dias); // col M
    count++;
  }

  return `Dias úteis calculados: ${count} linhas atualizadas.`;
}

function _copiarConsolidado() {
  const ss        = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
  const consSheet = ss.getSheetByName('bolsistas_consolidado');
  const appSheet  = ss.getSheetByName('Bolsistas App');
  if (!consSheet) throw new Error('"bolsistas_consolidado" não encontrada.');
  if (!appSheet)  throw new Error('"Bolsistas App" não encontrada.');

  const data = consSheet.getDataRange().getValues();
  const now  = new Date();
  const rowsToCopy = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][26] || '').trim() === 'Sim') { // col AA
      rowsToCopy.push(data[i].slice(0, 22)); // cols A:V
    }
  }

  if (!rowsToCopy.length) return 'Copiar consolidado: nenhuma linha com AA="Sim".';

  const startRow = appSheet.getLastRow() + 1;
  const fullRows = rowsToCopy.map(r => {
    const row = new Array(24).fill('');
    row[0] = now;
    row[1] = '';
    for (let j = 0; j < 22; j++) row[j + 2] = r[j];
    return row;
  });

  appSheet.getRange(startRow, 1, fullRows.length, 24).setValues(fullRows);

  return `Copiar consolidado: ${rowsToCopy.length} linhas adicionadas à "Bolsistas App".`;
}

function _formatarHorarios() {
  const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
  const sheet = ss.getSheetByName('Bolsistas App');
  if (!sheet) throw new Error('"Bolsistas App" não encontrada.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'Formatar horários: nenhuma linha.';

  const range  = sheet.getRange(2, 14, lastRow - 1, 1); // col N a partir da linha 2
  const values = range.getValues();
  let count = 0;

  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (typeof v === 'string' && v.includes(' às ') && !/\d{1,2}h às \d{1,2}(:\d{2})?h/.test(v)) {
      values[i][0] = v.replace(/(\d{2}):(\d{2})/g, (_, h, m) => m === '00' ? h + 'h' : `${h}:${m}h`);
      count++;
    }
  }

  range.setValues(values);
  return `Formatar horários: ${count} células formatadas.`;
}

// ─── Diagnóstico: testa leitura da aba RJ-UNIDADES ───────────
function testeFuncionarios() {
  const fss = SpreadsheetApp.openById(FUNCIONARIOS_SHEET_ID);
  const sh  = fss.getSheetByName('RJ - UNIDADES');
  if (!sh) {
    Logger.log('ABA NÃO ENCONTRADA. Abas disponíveis: ' + fss.getSheets().map(s => s.getName()).join(', '));
    return;
  }
  const rows = sh.getDataRange().getValues();
  Logger.log('Total de linhas (sem cabeçalho): ' + (rows.length - 1));
  Logger.log('Cabeçalho: ' + rows[0].join(' | '));
  let ativos = 0;
  for (let i = 1; i < rows.length && i < 6; i++) {
    const nome = rows[i][2], status = rows[i][10], unit1 = rows[i][21], unit2 = rows[i][30];
    Logger.log('Linha ' + i + ' → C=' + nome + ' | K=' + status + ' | V=' + unit1 + ' | AE=' + unit2);
    if (String(status).trim().toLowerCase() === 'ativo') ativos++;
  }
  Logger.log('Ativos nas primeiras 5 linhas: ' + ativos);
}

// ─── Chaves PDF disponíveis ──────────────────────────────────
function getChavesPDF(token, mes, ano) {
  try {
    const user = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão.');

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    const data  = sheet.getDataRange().getValues();

    const chaves = new Set();
    for (let i = 1; i < data.length; i++) {
      if (mes && String(data[i][3] || '').trim() !== mes) continue;
      if (ano && String(data[i][4] || '').trim() !== ano) continue;
      const chave = String(data[i][24] || '').trim();
      if (chave) chaves.add(chave);
    }

    return JSON.stringify({ ok: true, chaves: [...chaves].sort() });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Preview de e-mail ───────────────────────────────────────
function getEmailPreview(token, chavePDF) {
  try {
    const user = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão.');

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    const data  = sheet.getDataRange().getValues();

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][24] || '').trim() === chavePDF.trim()) rows.push(data[i]);
    }
    if (!rows.length) throw new Error('Nenhum aluno encontrado para esta chave.');

    const origemBolsa = String(rows[0][7] || '').trim();
    const parcSheet   = ss.getSheetByName('Parceiros');
    let parceiro      = { origemBolsa, logo: '', to: '', cc: [] };

    if (parcSheet) {
      const pd = parcSheet.getDataRange().getValues();
      for (let i = 1; i < pd.length; i++) {
        if (String(pd[i][0] || '').trim() !== origemBolsa) continue;
        parceiro = {
          origemBolsa,
          logo:   String(pd[i][1] || '').trim(),
          to:     String(pd[i][2] || '').trim(),
          cc:     [pd[i][3], pd[i][4]].map(e => String(e || '').trim()).filter(Boolean),
        };
        break;
      }
    }

    const obsExtras = rows.map(r => String(r[27] || '').trim()).filter(Boolean);
    const parts = chavePDF.split(' - ');

    return JSON.stringify({
      ok: true,
      chavePDF,
      ano:         parts[0] || '',
      mes:         parts[1] || '',
      parceiro,
      totalAlunos: rows.length,
      obsExtras,
    });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Helper: logo do parceiro na pasta do Drive ──────────────
function _getPartnerLogoBlob(origemBolsa) {
  if (!PARTNER_LOGOS_FOLDER_ID || !origemBolsa) return null;
  try {
    const folder = DriveApp.getFolderById(PARTNER_LOGOS_FOLDER_ID);
    const files  = folder.getFiles();
    const nome   = origemBolsa.trim().toLowerCase();
    while (files.hasNext()) {
      const f  = files.next();
      const fn = f.getName();
      const base = fn.includes('.') ? fn.substring(0, fn.lastIndexOf('.')).toLowerCase() : fn.toLowerCase();
      if (base === nome) return f.getBlob();
    }
  } catch (e) {}
  return null;
}

// ─── Constrói seção de relatório dentro de um Doc ────────────
function _buildPDFSection(body, rows, origemBolsa, mes, ano, isFirst) {
  const NAVY       = '#1d3557';
  const STEEL      = '#334155';
  const LIGHT_BLUE = '#dce6f1';
  const WHITE      = '#ffffff';
  const GRAY       = '#64748b';
  const BORDER     = '#c9d4e0';

  if (!isFirst) body.appendPageBreak();

  // ── Cabeçalho: logo parceiro | título | logo BRASAS ──
  const hdrTbl = body.appendTable([['', '', '']]);
  hdrTbl.setBorderWidth(0);

  // Esquerda: logo do parceiro
  const cParc = hdrTbl.getCell(0, 0);
  const partnerBlob = _getPartnerLogoBlob(origemBolsa);
  if (partnerBlob) {
    try {
      cParc.editAsText().setText('');
      const img = cParc.getChild(0).asParagraph().appendInlineImage(partnerBlob);
      const w = img.getWidth(), h = img.getHeight();
      if (w > 140 || h > 50) { const s = Math.min(140/w, 50/h); img.setWidth(Math.round(w*s)).setHeight(Math.round(h*s)); }
    } catch(e2) {
      cParc.editAsText().setText(origemBolsa).setFontSize(12).setFontFamily('Arial').setBold(true).setForegroundColor(NAVY);
    }
  } else {
    cParc.editAsText().setText(origemBolsa).setFontSize(13).setFontFamily('Arial').setBold(true).setForegroundColor(NAVY);
  }
  // Centro: título
  const cTit = hdrTbl.getCell(0, 1);
  cTit.editAsText().setText('RELATÓRIO DE BOLSISTAS').setBold(true).setFontSize(26).setFontFamily('Arial').setForegroundColor(NAVY);
  const subPara = cTit.appendParagraph(`${origemBolsa}  ·  ${mes} / ${ano}`);
  subPara.editAsText().setFontSize(12).setFontFamily('Arial').setBold(false).setForegroundColor(STEEL);
  // Alinhamento central nos parágrafos do título
  for (let ci = 0; ci < cTit.getNumChildren(); ci++) {
    const ch = cTit.getChild(ci);
    if (ch.getType() === DocumentApp.ElementType.PARAGRAPH) {
      ch.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
  }

  // Direita: logo BRASAS (alinhado à direita via parágrafo)
  const cBrs = hdrTbl.getCell(0, 2);
  if (BRASAS_LOGO_FILE_ID) {
    try {
      cBrs.editAsText().setText('');
      const bPara = cBrs.getChild(0).asParagraph();
      bPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      const bImg = bPara.appendInlineImage(DriveApp.getFileById(BRASAS_LOGO_FILE_ID).getBlob());
      const w = bImg.getWidth(), h = bImg.getHeight();
      if (w > 500 || h > 170) { const s = Math.min(500/w, 170/h); bImg.setWidth(Math.round(w*s)).setHeight(Math.round(h*s)); }
    } catch(e2) {
      cBrs.editAsText().setText('BRASAS').setBold(true).setFontSize(15).setFontFamily('Arial').setForegroundColor(NAVY);
      cBrs.appendParagraph('ENGLISH COURSE').editAsText().setFontSize(9).setFontFamily('Arial').setBold(false).setForegroundColor(GRAY);
    }
  } else {
    cBrs.editAsText().setText('BRASAS').setBold(true).setFontSize(12).setFontFamily('Arial').setForegroundColor(NAVY);
    cBrs.appendParagraph('ENGLISH COURSE').editAsText().setFontSize(7).setFontFamily('Arial').setBold(false).setForegroundColor(GRAY);
  }

  // Linha em branco pequena
  body.appendParagraph('').editAsText().setFontSize(4);

  // ── Legenda (2 colunas) ──
  const legTbl = body.appendTable([['', '']]);
  legTbl.setBorderWidth(0);
  legTbl.getCell(0, 0).setWidth(860);
  legTbl.getCell(0, 1).setWidth(568);
  const cLeg = legTbl.getCell(0, 0);
  cLeg.editAsText().setText('Legenda:').setBold(true).setFontSize(12).setFontFamily('Arial').setForegroundColor(NAVY);
  ['MT - Midterm Test - Prova realizada na metade do módulo',
   'WT - Written Test - Prova escrita',
   'OC - Oral Comprehension - Prova de compreensão oral',
   'OT - Oral Test - Prova Oral'].forEach(function(t) {
    cLeg.appendParagraph(t).editAsText().setFontSize(11).setFontFamily('Arial').setBold(false).setForegroundColor(STEEL);
  });
  const cConc = legTbl.getCell(0, 1);
  cConc.editAsText().setText('Conceito:').setBold(true).setFontSize(12).setFontFamily('Arial').setForegroundColor(NAVY);
  ['E - Excelente','MB - Muito Bom','B - Bom','R - Regular','I - Insuficiente','N - Nulo'].forEach(function(t) {
    cConc.appendParagraph(t).editAsText().setFontSize(11).setFontFamily('Arial').setBold(false).setForegroundColor(STEEL);
  });

  body.appendParagraph('').editAsText().setFontSize(4);

  // ── Tabela de dados ──
  const hdrs = [
    'Aluno', 'Unidade', '% Bolsa', 'Módulo', 'Freq. / Horário',
    'Início\nMódulo', 'Prev.\nConclusão', 'Aulas\nPrev.', 'Aulas\nAssist.',
    'Faltas\n/Mês', 'Valor', 'T1', 'T2', 'MT', 'WT', 'OC', 'OT',
    'Média Final', 'Conceito', 'Observações',
  ];

  const tableData = [hdrs];
  rows.forEach(r => {
    const wt = parseFloat(r[18]), oc = parseFloat(r[19]), ot = parseFloat(r[20]);
    const mt = parseFloat(r[17]);
    let nota = '';
    if (!isNaN(wt) && !isNaN(oc) && !isNaN(ot)) nota = ((wt + oc + ot) / 3).toFixed(1);
    else if (!isNaN(mt)) nota = Number(mt).toFixed(1);
    const notaN = parseFloat(nota);
    const dias  = parseFloat(r[15]);
    let aprov   = String(r[21] || '');
    if (nota !== '') {
      if (!isNaN(dias) && dias <= 1) aprov = 'N';
      else if (notaN >= 90) aprov = 'E';
      else if (notaN >= 80) aprov = 'MB';
      else if (notaN >= 70) aprov = 'B';
      else if (notaN >= 60) aprov = 'R';
      else aprov = 'I';
    }
    const falta = (r[14] !== '' || r[15] !== '')
      ? String((parseFloat(r[14]) || 0) - (parseFloat(r[15]) || 0)) : '';

    tableData.push([
      String(r[5]  || ''),
      String(r[2]  || ''),
      r[6]  !== '' ? String(r[6]) + '%' : '',
      String(r[9]  || ''),
      `${r[10] || ''} / ${r[13] || ''}`,
      _fmtDate(r[11]),
      _fmtDate(r[12]),
      r[14] !== '' ? String(r[14]) : '',
      r[15] !== '' ? String(r[15]) : '',
      falta,
      r[16] !== '' ? 'R$ ' + r[16] : '',
      r[30] !== '' ? String(r[30]) : '',
      r[31] !== '' ? String(r[31]) : '',
      r[17] !== '' ? String(r[17]) : '',
      r[18] !== '' ? String(r[18]) : '',
      r[19] !== '' ? String(r[19]) : '',
      r[20] !== '' ? String(r[20]) : '',
      nota,
      aprov,
      String(r[22] || ''),
    ]);
  });

  const table = body.appendTable(tableData);
  table.setBorderWidth(0.3);

  // Larguras das colunas somando ~1428 pts (página 1500 - margens 36×2)
  const colWidths = [140,72,42,85,110,78,78,48,48,48,70,35,35,35,35,35,35,62,60,277];
  const hRow = table.getRow(0);
  for (let c = 0; c < hdrs.length; c++) {
    hRow.getCell(c).setBackgroundColor(NAVY)
      .editAsText().setForegroundColor(WHITE).setBold(true).setFontSize(12).setFontFamily('Arial');
    if (colWidths[c]) hRow.getCell(c).setWidth(colWidths[c]);
  }
  for (let ri = 1; ri < table.getNumRows(); ri++) {
    for (let c = 0; c < hdrs.length; c++) {
      const cell = table.getCell(ri, c);
      cell.editAsText().setFontSize(10).setFontFamily('Arial');
      if (ri % 2 === 0) cell.setBackgroundColor(LIGHT_BLUE);
    }
  }

  // ── Rodapé ──
  body.appendParagraph(
    `Total: ${rows.length} aluno${rows.length !== 1 ? 's' : ''}  ·  ` +
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') +
    '  ·  BRASAS English Course'
  ).editAsText().setFontSize(9).setFontFamily('Arial').setForegroundColor(GRAY);
}

// ─── Salva PDF no Drive (helper interno) ─────────────────────
function _savePDF(docId, docTitle, origemBolsa) {
  const docFile = DriveApp.getFileById(docId);
  const pdfBlob = docFile.getAs('application/pdf').setName(docTitle + '.pdf');

  const roots = DriveApp.getFoldersByName('Relatórios Bolsistas');
  const root  = roots.hasNext() ? roots.next() : DriveApp.createFolder('Relatórios Bolsistas');

  let folder = root;
  if (origemBolsa) {
    const subs = root.getFoldersByName(origemBolsa);
    folder = subs.hasNext() ? subs.next() : root.createFolder(origemBolsa);
  }

  const existing = folder.getFilesByName(docTitle + '.pdf');
  while (existing.hasNext()) existing.next().setTrashed(true);

  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  docFile.setTrashed(true);

  return JSON.stringify({
    ok:       true,
    fileId:   pdfFile.getId(),
    fileName: pdfFile.getName(),
    viewUrl:  pdfFile.getUrl(),
  });
}

// ─── Geração de PDF (parceiro único) ─────────────────────────
function gerarPDF(token, chavePDF, excludedRowsJson, customTitle) {
  try {
    const user     = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão para gerar PDF.');
    const excluded = excludedRowsJson ? new Set(JSON.parse(excludedRowsJson).map(Number)) : new Set();

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    const data  = sheet.getDataRange().getValues();

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][24] || '').trim() === chavePDF.trim() && !excluded.has(i + 1)) {
        rows.push(data[i]);
      }
    }
    if (!rows.length) throw new Error('Nenhum aluno selecionado para este relatório.');

    const origemBolsa = String(rows[0][7] || '').trim();
    const parts       = chavePDF.split(' - ');
    const ano         = parts[0] || '';
    const mes         = parts[1] || '';

    const docTitle = (customTitle || '').trim() || `Relatório Bolsistas — ${chavePDF}`;
    const doc  = DocumentApp.create(docTitle);
    const body = doc.getBody();
    body.clear();
    body.setPageWidth(1500).setPageHeight(850);
    body.setMarginTop(28).setMarginBottom(28).setMarginLeft(36).setMarginRight(36);

    _buildPDFSection(body, rows, origemBolsa, mes, ano, true);
    doc.saveAndClose();
    return _savePDF(doc.getId(), docTitle, origemBolsa);
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Geração de PDF mesclado (múltiplos parceiros) ───────────
function gerarPDFMesclado(token, chavesJSON, excludedRowsJson) {
  try {
    const user     = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão para gerar PDF.');
    const chaves   = JSON.parse(chavesJSON);
    if (!chaves.length) throw new Error('Nenhuma chave selecionada.');
    const excluded = excludedRowsJson ? new Set(JSON.parse(excludedRowsJson).map(Number)) : new Set();

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    const data  = sheet.getDataRange().getValues();

    const byChave = {};
    chaves.forEach(c => { byChave[c] = []; });
    for (let i = 1; i < data.length; i++) {
      const chave = String(data[i][24] || '').trim();
      if (byChave[chave] !== undefined && !excluded.has(i + 1)) byChave[chave].push(data[i]);
    }

    const docTitle = `Relatório Bolsistas — Mesclado (${chaves.length} parceiros)`;
    const doc  = DocumentApp.create(docTitle);
    const body = doc.getBody();
    body.clear();
    body.setPageWidth(1500).setPageHeight(850);
    body.setMarginTop(28).setMarginBottom(28).setMarginLeft(36).setMarginRight(36);

    let first = true;
    chaves.forEach(chave => {
      const rows = byChave[chave] || [];
      if (!rows.length) return;
      const origemBolsa = String(rows[0][7] || '').trim();
      const parts = chave.split(' - ');
      _buildPDFSection(body, rows, origemBolsa, parts[1] || '', parts[0] || '', first);
      first = false;
    });

    doc.saveAndClose();
    return _savePDF(doc.getId(), docTitle, 'Mesclado');
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Envio de e-mail ─────────────────────────────────────────
function enviarEmail(token, payloadJson) {
  try {
    const user = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão para enviar e-mails.');

    const { chavePDF, fileId, obsExtra } = JSON.parse(payloadJson);

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    const data  = sheet.getDataRange().getValues();

    const rowIndexes = [];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][24] || '').trim() === chavePDF.trim()) {
        rows.push(data[i]);
        rowIndexes.push(i + 1);
      }
    }
    if (!rows.length) throw new Error('Nenhum aluno encontrado.');

    const origemBolsa = String(rows[0][7] || '').trim();
    const parts       = chavePDF.split(' - ');
    const ano         = parts[0] || '';
    const mes         = parts[1] || '';

    // Destinatários
    const parcSheet = ss.getSheetByName('Parceiros');
    let to = '', cc = '';
    if (parcSheet) {
      const pd = parcSheet.getDataRange().getValues();
      for (let i = 1; i < pd.length; i++) {
        if (String(pd[i][0] || '').trim() !== origemBolsa) continue;
        to = String(pd[i][2] || '').trim();
        cc = [pd[i][3], pd[i][4]].map(e => String(e || '').trim()).filter(Boolean).join(',');
        break;
      }
    }
    if (!to) throw new Error('E-mail do parceiro não encontrado para: ' + origemBolsa);

    // Anexo PDF
    const attachments = [];
    if (fileId) {
      attachments.push(DriveApp.getFileById(fileId).getAs('application/pdf'));
    }

    const subject  = `Relatório de Frequência e Aproveitamento - BRASAS - ${origemBolsa}`;
    const htmlBody = _buildEmailHtml(origemBolsa, mes, ano, rows, obsExtra);

    GmailApp.sendEmail(to, subject, '', {
      htmlBody,
      cc:          cc || undefined,
      attachments,
      name:        'Administrativo BRASAS',
      replyTo:     'administrativo@brasas.com',
    });

    // Marca envio nas linhas (col AA = index 27)
    for (const ri of rowIndexes) {
      sheet.getRange(ri, 27).setValue(true);
    }

    return JSON.stringify({ ok: true, to, cc });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

function _buildEmailHtml(origemBolsa, mes, ano, rows, obsExtra) {
  const td  = (v, center, bold) =>
    `<td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;${center?'text-align:center;':''}${bold?'font-weight:700;':''}">${v}</td>`;

  const rowsHtml = rows.map((r, idx) => {
    const wt = parseFloat(r[18]), oc = parseFloat(r[19]), ot = parseFloat(r[20]);
    const mt = parseFloat(r[17]);
    let nota = '';
    if (!isNaN(wt) && !isNaN(oc) && !isNaN(ot)) nota = ((wt + oc + ot) / 3).toFixed(1);
    else if (!isNaN(mt)) nota = Number(mt).toFixed(1);
    const notaN = parseFloat(nota);
    const dias  = parseFloat(r[15]);
    let aprov   = String(r[21] || '');
    if (nota !== '') {
      if (!isNaN(dias) && dias <= 1) aprov = 'N';
      else if (notaN >= 90) aprov = 'E';
      else if (notaN >= 80) aprov = 'MB';
      else if (notaN >= 70) aprov = 'B';
      else if (notaN >= 60) aprov = 'R';
      else aprov = 'I';
    }
    const falta = (r[14] !== '' || r[15] !== '')
      ? String((parseFloat(r[14])||0) - (parseFloat(r[15])||0)) : '';

    return `<tr style="background:${idx % 2 === 0 ? '#fff' : '#f8fafc'}">
      ${td(`<b style="color:#1d3557">${r[2]}</b> — ${r[5]}`)}
      ${td(r[6] !== '' ? r[6] + '%' : '', true)}
      ${td(r[9] || '', true)}
      ${td((r[10]||'') + (r[13] ? '<br><span style="font-size:11px;color:#64748b">'+r[13]+'</span>' : ''), true)}
      ${td(r[14] !== '' ? String(r[14]) : '', true)}
      ${td(r[15] !== '' ? String(r[15]) : '', true)}
      ${td(falta, true)}
      ${td(r[30] !== '' ? String(r[30]) : '', true)}
      ${td(r[31] !== '' ? String(r[31]) : '', true)}
      ${td(r[17] !== '' ? String(r[17]) : '', true)}
      ${td(r[18] !== '' ? String(r[18]) : '', true)}
      ${td(r[19] !== '' ? String(r[19]) : '', true)}
      ${td(r[20] !== '' ? String(r[20]) : '', true)}
      ${td(nota, true)}
      ${td(aprov, true, true)}
      ${td(`<span style="font-size:11px;color:#475569">${r[22] || ''}</span>`)}
    </tr>`;
  }).join('');

  const obsBlock = obsExtra
    ? `<div style="margin:20px 0 0;padding:14px 18px;background:#f8fafc;border-left:3px solid #1d3557;border-radius:0 6px 6px 0;font-size:13px;color:#475569"><b>Observações:</b> ${obsExtra}</div>`
    : '';

  const th = txt => `<th style="padding:9px 8px;text-align:center;font-weight:600;white-space:nowrap">${txt}</th>`;

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;margin:0;padding:24px 16px;background:#f1f5f9">
  <div style="max-width:1020px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 16px rgba(0,0,0,.09)">

    <div style="background:#1d3557;padding:24px 32px">
      <div style="font-size:17px;font-weight:700;color:#fff;margin-bottom:5px">Relatório de Frequência e Aproveitamento</div>
      <div style="font-size:13px;color:rgba(255,255,255,.65)">BRASAS English Course &nbsp;·&nbsp; ${mes} / ${ano}</div>
    </div>

    <div style="padding:28px 32px 24px">
      <p style="margin:0 0 14px;font-size:14px;color:#334155">Olá!</p>
      <p style="margin:0;font-size:14px;color:#334155;line-height:1.65">
        Segue em anexo o <strong>Relatório de Frequência e Aproveitamento</strong> dos alunos bolsistas de
        <strong>${origemBolsa}</strong>, referente ao mês de <strong>${mes} / ${ano}</strong>.
      </p>
      ${obsBlock}
    </div>

    <div style="border-top:1px solid #e2e8f0;overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:#1d3557;color:#fff">
            <th style="padding:9px 12px;text-align:left;min-width:180px;font-weight:600">Unidade — Nome</th>
            ${th('% Bolsa')}${th('Módulo')}${th('Freq. / Horário')}
            ${th('Aulas Prev.')}${th('Aulas Assist.')}${th('Faltas/Mês')}
            ${th('T1')}${th('T2')}
            ${th('MT')}${th('WT')}${th('OC')}${th('OT')}
            ${th('Média Final')}${th('Conceito')}
            <th style="padding:9px 12px;text-align:left;font-weight:600">Observações</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>

    <div style="padding:16px 32px;font-size:11px;color:#94a3b8;border-top:1px solid #f1f5f9;text-align:right">
      BRASAS English Course &nbsp;·&nbsp; Gerado automaticamente
    </div>

  </div>
</body></html>`;
}

// ─── Helpers ─────────────────────────────────────────────────
function _fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(val);
}
