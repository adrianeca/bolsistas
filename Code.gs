// ============================================================
//  BRASAS Bolsistas — Code.gs
// ============================================================

const USERS_SHEET_ID          = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';
const BOLSISTAS_SHEET_ID      = '1G7RfHP_8j7-6VPqC8ReYvpScyWvBnVyRXu8QAdeL4bQ';
const FUNCIONARIOS_SHEET_ID   = '1BDiPjv0FqRJp5EwcvLdYXVvEAWesvwdEgbhYdnTlqPY';

// Logos: (1) Suba "Brasas logo.png" no Drive, copie o ID da URL e cole aqui.
// (2) Logos dos parceiros ficam na pasta abaixo com nome = Origem da Bolsa (qualquer extensão).
const BRASAS_LOGO_FILE_ID       = '1fqZbnxHJNyov_9NwhwDAg235FizcnTGQ'; // logo usada no PDF
const BRASAS_EMAIL_LOGO_FILE_ID = '1F-0NV036KUExaEGIMgLbUHqE-HtYyXvK'; // logo usada no corpo do e-mail
const PARTNER_LOGOS_FOLDER_ID   = '1sKrz_-odKjx6YNVCSWpB5eW9v3atEvQD';

const EDIT_ROLES       = ['admin', 'secretaria', 'diretor', 'b2b', 'operacional'];
const EMAIL_ROLES      = ['admin', 'b2b'];
const ADMIN_ROLES      = ['admin'];
const LOCK_EXEMPT_ROLES = ['admin', 'b2b', 'operacional'];

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

    const access = _hasAccess(ss, role, email);
    if (!access) throw new Error('Sem permissão para acessar o painel de bolsistas.');

    const unidade = _getUserUnidade(ss, email);

    const userObj = {
      email,
      nome:              String(rows[i][2]).trim(),
      role,
      unidade,
      canEdit:           EDIT_ROLES.includes(role),
      canSendEmail:      EMAIL_ROLES.includes(role),
      isAdmin:           ADMIN_ROLES.includes(role),
      canViewAlunos:     access.canViewAlunos,
      canViewRelatorios: access.canViewRelatorios,
      canViewHistorico:  access.canViewHistorico,
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
  if (!usuSheet) return null;

  const d = usuSheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (norm(d[i][0]) !== norm(email)) continue;
    // Colunas F (índice 5) e G (índice 6) = acessos_dashboards
    const acessos = [d[i][5], d[i][6]]
      .join(',')
      .split(',')
      .map(s => norm(s));

    const hasFull = acessos.includes('bolsistas');
    const canViewAlunos      = hasFull || acessos.includes('bolsistas_alunos');
    const canViewRelatorios  = hasFull || acessos.includes('bolsistas_relatorios');
    const canViewHistorico   = hasFull || acessos.includes('bolsistas_historico');

    if (!canViewAlunos && !canViewRelatorios) return null;
    return { canViewAlunos, canViewRelatorios, canViewHistorico };
  }

  return null;
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

    // ── Feriados (cache 6h — raramente mudam) ──
    let feriados = [];
    const ferCacheKey = 'feriados_v1';
    const ferCached = cache.get(ferCacheKey);
    if (ferCached) {
      feriados = JSON.parse(ferCached);
    } else {
      const feriadosSheet = ss.getSheetByName('Feriados');
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
      try { cache.put(ferCacheKey, JSON.stringify(feriados), 21600); } catch(e) {}
    }

    // ── Funcionários (cache 1h — planilha separada, custo alto) ──
    let funcionarios = { unidades: [], porUnidade: {} };
    const funcCacheKey = 'funcionarios_v1';
    const funcCached = cache.get(funcCacheKey);
    if (funcCached) {
      funcionarios = JSON.parse(funcCached);
    } else {
      const funcSheet = SpreadsheetApp.openById(FUNCIONARIOS_SHEET_ID).getSheetByName('RJ - UNIDADES');
      if (funcSheet) {
        const fRows = funcSheet.getDataRange().getValues();
        const unidadesSet = new Set();
        for (let i = 1; i < fRows.length; i++) {
          const fr    = fRows[i];
          const nome  = String(fr[2]  || '').trim();
          const status= String(fr[10] || '').trim();
          const unit1 = String(fr[21] || '').trim();
          const unit2 = String(fr[30] || '').trim();
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
      try { cache.put(funcCacheKey, JSON.stringify(funcionarios), 3600); } catch(e) {}
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
        editadoWebapp:     !!r[32],
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
        emailEnviado:      !!r[33],
      });
    }

    const isEditLocked = !LOCK_EXEMPT_ROLES.includes(user.role) && _isPastFifthBusinessDay();

    const result = JSON.stringify({
      ok:       true,
      rows,
      feriados,
      funcionarios,
      isEditLocked,
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

function reloadApp(token) {
  try { CacheService.getScriptCache().remove('appdata_v3_' + token); } catch(e) {}
  return initApp(token);
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
        editadoWebapp:     !!r[32],
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
        emailEnviado:      !!r[33],
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

    if (!LOCK_EXEMPT_ROLES.includes(user.role) && _isPastFifthBusinessDay()) {
      throw new Error('Edição bloqueada após o 5º dia útil do mês.');
    }

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

    if (field === 'dependenteUnidade' || field === 'dependenteNome') {
      const ss2   = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
      const sh2   = ss2.getSheetByName('Bolsistas App');
      const rowOrigem = String(sh2.getRange(rowIndex, 8, 1, 1).getValue() || '').toLowerCase();
      if (!rowOrigem.includes('dependente brasas')) {
        throw new Error('Este campo só pode ser editado para bolsistas Dependente BRASAS.');
      }
    }

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');

    // Lê contexto da linha (nome, unidade, mês, ano) + valor anterior em uma chamada
    const ctxRange = Math.max(col, 6);
    const ctx      = sheet.getRange(rowIndex, 1, 1, ctxRange).getValues()[0];
    const oldValue = ctx[col - 1];

    // Valida acesso à unidade do row para não-admin
    if (!user.isAdmin && user.unidade) {
      const norm    = s => String(s || '').trim().toLowerCase()
        .normalize('NFD').replace(/[̀-ͯ]/g, '');
      const permitidas = user.unidade.split(/[,|]/).map(s => norm(s.trim()));
      if (!permitidas.includes(norm(String(ctx[2])))) {
        throw new Error('Sem acesso a esta unidade.');
      }
    }

    sheet.getRange(rowIndex, col).setValue(value);
    sheet.getRange(rowIndex, 1).setValue(new Date());
    sheet.getRange(rowIndex, 33).setValue(true); // coluna dedicada ao flag "editado via webapp"

    // Log de edição
    _logEdicao(ss, ctx, rowIndex, field, oldValue, value, user.email);

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
    const parts   = chavePDF.split(' - ');
    const pdfList = _listPDFsNaPasta(origemBolsa);

    return JSON.stringify({
      ok: true,
      chavePDF,
      ano:         parts[0] || '',
      mes:         parts[1] || '',
      parceiro,
      totalAlunos: rows.length,
      obsExtras,
      pdfList,
    });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Helper: lista PDFs salvos para um parceiro no Drive ─────
function _listPDFsNaPasta(origemBolsa) {
  try {
    // Tenta localizar pela hierarquia de pastas primeiro
    const roots = DriveApp.getFoldersByName('Relatórios Bolsistas');
    if (roots.hasNext()) {
      const root = roots.next();
      const subs = root.getFoldersByName(origemBolsa);
      if (subs.hasNext()) {
        const folder = subs.next();
        const iter   = folder.getFilesByType('application/pdf');
        const result = [];
        while (iter.hasNext()) {
          const f = iter.next();
          result.push({ name: f.getName().replace(/\.pdf$/i, ''), fileId: f.getId(), viewUrl: f.getUrl() });
        }
        result.sort((a, b) => b.name.localeCompare(a.name));
        if (result.length) return result;
      }
    }
    // Fallback: busca por nome de arquivo contendo a origem da bolsa
    const safe  = origemBolsa.replace(/'/g, "\\'");
    const iter2 = DriveApp.searchFiles(
      "mimeType='application/pdf' and title contains '" + safe + "' and trashed=false"
    );
    const result2 = [];
    while (iter2.hasNext()) {
      const f = iter2.next();
      result2.push({ name: f.getName().replace(/\.pdf$/i, ''), fileId: f.getId(), viewUrl: f.getUrl() });
    }
    result2.sort((a, b) => b.name.localeCompare(a.name));
    return result2;
  } catch(e) { return []; }
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
function _buildPDFSection(body, rows, origemBolsa, mes, ano, isFirst, customOrigemBolsa) {
  const NAVY       = '#1d3557';
  const NAVY_B     = '#2a4d76';
  const STEEL      = '#334155';
  const LIGHT_BLUE = '#dce6f1';
  const WHITE      = '#ffffff';
  const GRAY       = '#64748b';
  if (!isFirst) body.appendPageBreak();

  // logos=500, title=428, legend=500 → centro do title = 714pt = centro da página ✓
  // dois logos lado a lado em 500pt: 244+8+244=496 < 500 ✓
  const LOGO_MAX_W = 244, LOGO_MAX_H = 140;

  // ── Faixa 1: logos (esq) | título+card (centro) | legenda+conceito (dir) ──
  // 500 + 428 + 500 = 1428
  const hdrTbl = body.appendTable([['', '', '']]);
  hdrTbl.setBorderWidth(0);
  hdrTbl.getRow(0).getCell(0).setWidth(500);
  hdrTbl.getRow(0).getCell(1).setWidth(428);
  hdrTbl.getRow(0).getCell(2).setWidth(500);

  // Esquerda: BRASAS logo + parceiro logo no mesmo parágrafo (lado a lado)
  const cLogos = hdrTbl.getCell(0, 0);
  cLogos.editAsText().setText('');
  const logoPara = cLogos.getChild(0).asParagraph();

  let brasasAdded = false;
  if (BRASAS_LOGO_FILE_ID) {
    try {
      const bImg = logoPara.appendInlineImage(DriveApp.getFileById(BRASAS_LOGO_FILE_ID).getBlob());
      const bw = bImg.getWidth(), bh = bImg.getHeight();
      const bs = Math.min(LOGO_MAX_W / bw, LOGO_MAX_H / bh);
      bImg.setWidth(Math.round(bw * bs)).setHeight(Math.round(bh * bs));
      brasasAdded = true;
    } catch(e2) {
      cLogos.editAsText().setText('BRASAS').setFontSize(14).setFontFamily('Arial').setBold(true).setForegroundColor(NAVY);
    }
  } else {
    cLogos.editAsText().setText('BRASAS').setFontSize(14).setFontFamily('Arial').setBold(true).setForegroundColor(NAVY);
  }

  // Parceiro logo: no mesmo parágrafo → lado a lado com BRASAS
  // Se não encontrar, deixa em branco (não adiciona texto)
  const partnerBlob = _getPartnerLogoBlob(origemBolsa);
  if (partnerBlob && brasasAdded) {
    try {
      logoPara.appendText('          ');
      const pImg = logoPara.appendInlineImage(partnerBlob);
      const pw = pImg.getWidth(), ph = pImg.getHeight();
      const ps = Math.min(LOGO_MAX_W / pw, LOGO_MAX_H / ph);
      pImg.setWidth(Math.round(pw * ps)).setHeight(Math.round(ph * ps));
    } catch(e2) { /* sem logo de parceiro — silencioso */ }
  }

  // Centro: título + card de alunos (centralizado, só borda)
  const cTit = hdrTbl.getCell(0, 1);
  const displayOrigem = (customOrigemBolsa || '').trim() || origemBolsa;
  cTit.editAsText().setText('RELATÓRIO DE BOLSISTAS').setBold(true).setFontSize(22).setFontFamily('Arial').setForegroundColor(NAVY);
  const subPara = cTit.appendParagraph(displayOrigem + '  ·  ' + mes + ' - ' + ano);
  subPara.editAsText().setFontSize(13).setFontFamily('Arial').setBold(false).setForegroundColor(STEEL);
  for (let ci = 0; ci < cTit.getNumChildren(); ci++) {
    const ch = cTit.getChild(ci);
    if (ch.getType() === DocumentApp.ElementType.PARAGRAPH)
      ch.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
  // Card de alunos: compacto e centralizado
  // Tabela espaçador 3-col (sem borda) → célula central tem tabela interna com borda
  // Valor total das bolsas (col Q = índice 16)
  const totalValor = rows.reduce((s, r) => s + (parseFloat(r[16]) || 0), 0);
  const fmtValor = totalValor > 0
    ? 'R$ ' + totalValor.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.')
    : '—';

  cTit.appendParagraph('').editAsText().setFontSize(4);
  // Dois cards lado a lado: [14|190 card1|10 gap|190 card2|24] = 428 ✓
  const cardsTbl = cTit.appendTable([['', '', '', '', '']]);
  cardsTbl.setBorderWidth(0);
  cardsTbl.getRow(0).getCell(0).setWidth(14);
  cardsTbl.getRow(0).getCell(1).setWidth(190);
  cardsTbl.getRow(0).getCell(2).setWidth(10);
  cardsTbl.getRow(0).getCell(3).setWidth(190);
  cardsTbl.getRow(0).getCell(4).setWidth(24);

  function _makeCard(cell, label, value, vSize) {
    cell.editAsText().setText('').setFontSize(1);
    const nc = cell.appendTable([['']]);
    nc.setBorderWidth(0.5);
    const cc = nc.getCell(0, 0);
    cc.editAsText().setText(label)
      .setFontSize(11).setFontFamily('Arial').setBold(false).setForegroundColor(GRAY);
    cc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const vp = cc.appendParagraph(String(value));
    vp.editAsText().setFontSize(vSize || 20).setFontFamily('Arial').setBold(true).setForegroundColor(NAVY);
    vp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
  _makeCard(cardsTbl.getRow(0).getCell(1), 'Total de Alunos',      rows.length, 20);
  _makeCard(cardsTbl.getRow(0).getCell(3), 'Valor Total em Bolsa', fmtValor,    20);

  // Direita: espaçador + legenda + conceito  (160 + 170 + 170 = 500pt)
  const cLeg = hdrTbl.getCell(0, 2);
  cLeg.editAsText().setText('').setFontSize(1);
  const legNested = cLeg.appendTable([['', '', '']]);
  legNested.setBorderWidth(0);
  legNested.getRow(0).getCell(0).setWidth(160); // espaçador
  legNested.getRow(0).getCell(1).setWidth(170);
  legNested.getRow(0).getCell(2).setWidth(170);

  const cAbbr = legNested.getRow(0).getCell(1);
  cAbbr.editAsText().setText('Legenda:').setBold(true).setFontSize(11).setFontFamily('Arial').setForegroundColor(NAVY);
  ['MT - Midterm Test', 'WT - Written Test', 'OC - Oral Comprehension', 'OT - Oral Test', 'T1 - Teste 1', 'T2 - Teste 2'].forEach(function(t) {
    cAbbr.appendParagraph(t).editAsText().setFontSize(10).setFontFamily('Arial').setBold(false).setForegroundColor(STEEL);
  });

  const cConc = legNested.getRow(0).getCell(2);
  cConc.editAsText().setText('Conceito:').setBold(true).setFontSize(11).setFontFamily('Arial').setForegroundColor(NAVY);
  ['E - Excelente', 'MB - Muito Bom', 'B - Bom', 'R - Regular', 'I - Insuficiente', 'N - Nulo'].forEach(function(t) {
    cConc.appendParagraph(t).editAsText().setFontSize(10).setFontFamily('Arial').setBold(false).setForegroundColor(STEEL);
  });

  body.appendParagraph('').editAsText().setFontSize(3);

  // ── Tabela de dados ──
  const hdrs = [
    'Aluno', 'Unidade', '% Bolsa', 'Módulo', 'Freq. / Horário\n/ Dias Aula',
    'Data\n1ª Aula', 'Início\nMódulo', 'Prev.\nConclusão',
    'Aulas\nPrev.', 'Aulas\nAssist.', 'Faltas\n/Mês',
    'Valor', 'T1', 'T2', 'MT', 'WT', 'OC', 'OT',
    'Média Final', 'Conceito', 'Observações',
  ];

  const tableData = [hdrs];
  rows.forEach(function(r) {
    const wt = parseFloat(r[18]), oc = parseFloat(r[19]), ot = parseFloat(r[20]);
    const mt = parseFloat(r[17]);
    let nota = '';
    if (!isNaN(wt) && !isNaN(oc) && !isNaN(ot)) nota = ((wt + oc + ot) / 3).toFixed(1);
    else if (!isNaN(mt)) nota = Number(mt).toFixed(1);
    const notaN = parseFloat(nota);
    const diasAssist = parseFloat(r[15]);
    let aprov = String(r[21] || '');
    if (nota !== '') {
      if (!isNaN(diasAssist) && diasAssist <= 1) aprov = 'N';
      else if (notaN >= 90) aprov = 'E';
      else if (notaN >= 80) aprov = 'MB';
      else if (notaN >= 70) aprov = 'B';
      else if (notaN >= 60) aprov = 'R';
      else aprov = 'I';
    }
    const falta = (r[14] !== '' || r[15] !== '')
      ? String((parseFloat(r[14]) || 0) - (parseFloat(r[15]) || 0)) : '';

    const diasAulaStr = String(r[29] || '').trim()
      .replace(/segunda(-feira)?/gi, 'Seg').replace(/ter[çc]a(-feira)?/gi, 'Ter')
      .replace(/quarta(-feira)?/gi, 'Qua').replace(/quinta(-feira)?/gi, 'Qui')
      .replace(/sexta(-feira)?/gi, 'Sex').replace(/s[áa]bado/gi, 'Sáb')
      .replace(/domingo/gi, 'Dom');
    const freqHorDias = (r[10] || '') + ' / ' + (r[13] || '') +
      (diasAulaStr ? '\n' + diasAulaStr : '');

    tableData.push([
      String(r[5]  || ''),
      String(r[2]  || ''),
      r[6]  !== '' ? String(r[6]) + '%' : '',
      String(r[9]  || ''),
      freqHorDias,
      _fmtDate(r[8]),
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

  // Larguras somando 1428 pts (página 1500 - margens 36×2)
  const colWidths = [135,68,38,80,120,62,75,75,46,46,46,68,33,33,33,33,33,33,60,58,253];
  const hRow = table.getRow(0);
  for (let c = 0; c < hdrs.length; c++) {
    const hCell = hRow.getCell(c);
    hCell.setBackgroundColor(NAVY)
      .editAsText().setForegroundColor(WHITE).setBold(true).setFontSize(10).setFontFamily('Arial');
    hCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    if (colWidths[c]) hCell.setWidth(colWidths[c]);
  }
  for (let ri = 1; ri < table.getNumRows(); ri++) {
    for (let c = 0; c < hdrs.length; c++) {
      const cell = table.getCell(ri, c);
      cell.editAsText().setFontSize(10).setFontFamily('Arial');
      if (ri % 2 === 0) cell.setBackgroundColor(LIGHT_BLUE);
    }
  }

}

// ─── URL da pasta raiz de relatórios ────────────────────────
function getRelatoriosFolderUrl(token) {
  try {
    _getUser(token);
    const folders = DriveApp.getFoldersByName('Relatórios Bolsistas');
    if (!folders.hasNext()) return JSON.stringify({ ok: false, error: 'Pasta não encontrada.' });
    const folder = folders.next();
    return JSON.stringify({ ok: true, url: folder.getUrl() });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Log de edições ──────────────────────────────────────────
const _FIELD_LABELS = {
  data1aAula:'Data 1ª Aula', frequencia:'Frequência', dataInicioBook:'Início do Módulo',
  dataPrevConclusao:'Prev. Conclusão', horario:'Horário', diasPrevistos:'Aulas Previstas',
  diasAssistidos:'Aulas Assistidas', valor:'Valor', mt:'MT', wt:'WT', oc:'OC', ot:'OT',
  aproveitamento:'Aproveitamento', observacoes:'Observações', turma:'Turma',
  dependenteUnidade:'Dep. Unidade', dependenteNome:'Dep. Nome',
  obsExtrasEmail:'Obs. E-mail', statusAluno:'Status', diasAula:'Dias de Aula',
  teste1:'Teste 1', teste2:'Teste 2',
};

function _logEdicao(ss, ctx, rowIndex, field, oldVal, newVal, email) {
  try {
    let log = ss.getSheetByName('log_edicoes');
    if (!log) {
      log = ss.insertSheet('log_edicoes');
      log.appendRow(['Timestamp','E-mail','Aluno','Unidade','Mês','Ano','Campo','Anterior','Novo','Linha']);
      log.setFrozenRows(1);
      log.getRange('E:F').setNumberFormat('@'); // evita auto-conversão de "Maio" → Date
    }
    const fmt = v => {
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      return v !== null && v !== undefined ? String(v) : '';
    };
    log.appendRow([
      new Date(), email,
      String(ctx[5] || ''), String(ctx[2] || ''), String(ctx[3] || ''), String(ctx[4] || ''),
      _FIELD_LABELS[field] || field, fmt(oldVal), fmt(newVal), rowIndex,
    ]);
  } catch(e) { /* silencioso: log não deve quebrar o save */ }
}

function getEditLog(token) {
  try {
    const user = _getUser(token);
    if (!user.isAdmin && !user.canViewHistorico) throw new Error('Sem permissão.');
    const ss  = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const log = ss.getSheetByName('log_edicoes');
    if (!log) return JSON.stringify({ ok: true, rows: [] });
    const data = log.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ ok: true, rows: [] });
    const tz = Session.getScriptTimeZone();
    const MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
    const fmtMes = v => v instanceof Date ? MESES_PT[v.getMonth()] : String(v || '');
    const fmtAno = v => v instanceof Date ? String(v.getFullYear())  : String(v || '');
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      rows.push({
        ts:       r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : String(r[0] || ''),
        email:    String(r[1] || ''),
        nome:     String(r[2] || ''),
        unidade:  String(r[3] || ''),
        mes:      fmtMes(r[4]),
        ano:      fmtAno(r[5]),
        campo:    String(r[6] || ''),
        anterior: String(r[7] || ''),
        novo:     String(r[8] || ''),
      });
    }
    rows.reverse(); // mais recentes primeiro
    return JSON.stringify({ ok: true, rows });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
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

// ─── Log de PDFs gerados ─────────────────────────────────────
function _logPDF(ss, tipo, origemBolsa, mes, ano, titulo, fileId, viewUrl, email, chavePDF) {
  try {
    let log = ss.getSheetByName('log_pdfs');
    if (!log) {
      log = ss.insertSheet('log_pdfs');
      log.appendRow(['Timestamp','GeradoPor','Tipo','OrigemBolsa','Mês','Ano','Título','FileId','ViewUrl','ChavePDF']);
      log.setFrozenRows(1);
      // Força colunas Mês (E) e Ano (F) como texto — evita auto-conversão de "Maio" → Date
      log.getRange('E:F').setNumberFormat('@');
    }
    const nextRow = log.getLastRow() + 1;
    // Pré-formata Mês(col E) e Ano(col F) como texto antes de escrever — evita auto-conversão
    const rng = log.getRange(nextRow, 1, 1, 10);
    rng.setNumberFormats([['dd/MM/yyyy HH:mm','@','@','@','@','@','@','@','@','@']]);
    rng.setValues([[new Date(), email, tipo, origemBolsa, mes, ano, titulo, fileId, viewUrl, chavePDF || '']]);
  } catch(e) { /* silencioso */ }
}

// ─── Geração de PDF (parceiro único) ─────────────────────────
function gerarPDF(token, chavePDF, excludedRowsJson, customTitle, customOrigemBolsa) {
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

    const docTitle = (customTitle || '').trim() || 'Relatório Bolsistas — ' + chavePDF;
    const doc  = DocumentApp.create(docTitle);
    const body = doc.getBody();
    body.clear();
    body.setPageWidth(1500).setPageHeight(850);
    body.setMarginTop(28).setMarginBottom(28).setMarginLeft(36).setMarginRight(36);

    _buildPDFSection(body, rows, origemBolsa, mes, ano, true, customOrigemBolsa || '');
    doc.saveAndClose();
    const result = JSON.parse(_savePDF(doc.getId(), docTitle, origemBolsa));
    if (result.ok) {
      const labelOrigem = (customOrigemBolsa || '').trim() || origemBolsa;
      _logPDF(ss, 'individual', labelOrigem, mes, ano, docTitle, result.fileId, result.viewUrl, user.email, chavePDF);
    }
    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Geração de PDF mesclado (múltiplos parceiros — tabela única) ───────────
function gerarPDFMesclado(token, chavesJSON, excludedRowsJson, customTitle, customOrigemBolsa, customFolderName) {
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

    // Reúne todas as linhas em uma só lista; coleta meses/anos distintos para subtítulo
    const allRows   = [];
    const seenMeses = [];
    const seenAnos  = [];
    chaves.forEach(chave => {
      const rows = byChave[chave] || [];
      if (!rows.length) return;
      const parts = chave.split(' - ');
      const m = parts[1] || '', a = parts[0] || '';
      if (m && !seenMeses.includes(m)) seenMeses.push(m);
      if (a && !seenAnos.includes(a))  seenAnos.push(a);
      allRows.push(...rows);
    });
    if (!allRows.length) throw new Error('Nenhum aluno selecionado.');

    const combinedMes = seenMeses.join(', ');
    const combinedAno = seenAnos.join(', ');

    const docTitle = (customTitle || '').trim() || `Relatório Bolsistas — Mesclado (${chaves.length} parceiros)`;
    const doc  = DocumentApp.create(docTitle);
    const body = doc.getBody();
    body.clear();
    body.setPageWidth(1500).setPageHeight(850);
    body.setMarginTop(28).setMarginBottom(28).setMarginLeft(36).setMarginRight(36);

    // Uma única seção com todos os alunos na mesma tabela
    _buildPDFSection(body, allRows, '', combinedMes, combinedAno, true, customOrigemBolsa || '');

    doc.saveAndClose();
    const folderName = (customFolderName || '').trim() || 'Mesclado';
    const result = JSON.parse(_savePDF(doc.getId(), docTitle, folderName));
    if (result.ok) {
      const labelOrigem = (customOrigemBolsa || '').trim() || ('Mesclado (' + chaves.length + ' parceiros)');
      _logPDF(ss, 'mesclado', labelOrigem, combinedMes, combinedAno, docTitle, result.fileId, result.viewUrl, user.email, '');
    }
    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Envio de e-mail ─────────────────────────────────────────
function enviarEmail(token, payloadJson) {
  try {
    const user = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão para enviar e-mails.');

    const { chavePDF, fileId, obsExtra, toOverride, ccOverride } = JSON.parse(payloadJson);

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

    // Destinatários: usa override do frontend se preenchido, senão busca na aba Parceiros
    let to = '', cc = '';
    if (toOverride && toOverride.trim()) {
      to = toOverride.trim();
      cc = (ccOverride || '').trim();
    } else {
      const parcSheet = ss.getSheetByName('Parceiros');
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
    }

    // CCs fixos sempre incluídos
    const fixedCC = ['adriane@brasas.com', 'administrativo@brasas.com'];
    const allCC   = [...new Set([...(cc ? cc.split(',').map(e=>e.trim()).filter(Boolean) : []), ...fixedCC])];
    cc = allCC.join(',');

    // Anexo PDF
    const attachments = [];
    if (fileId) {
      attachments.push(DriveApp.getFileById(fileId).getAs('application/pdf'));
    }

    // Logo BRASAS como imagem inline (cid:)
    const inlineImages = {};
    let logoSrc = '';
    try {
      const lb = DriveApp.getFileById(BRASAS_EMAIL_LOGO_FILE_ID).getBlob();
      lb.setName('brasasLogo');
      inlineImages.brasasLogo = lb;
      logoSrc = 'cid:brasasLogo';
    } catch(e) { /* sem logo — fallback para texto */ }

    const subject  = `Relatório de Frequência e Aproveitamento - BRASAS - ${origemBolsa}`;
    const htmlBody = _buildEmailHtml(origemBolsa, mes, ano, obsExtra, logoSrc);

    GmailApp.sendEmail(to, subject, '', {
      htmlBody,
      cc:          cc || undefined,
      attachments,
      name:        'Administrativo BRASAS',
      replyTo:     'administrativo@brasas.com',
      inlineImages: Object.keys(inlineImages).length ? inlineImages : undefined,
    });

    // Marca envio nas linhas (col AH = 34 — coluna dedicada, sem conflito com dados)
    for (const ri of rowIndexes) {
      sheet.getRange(ri, 34).setValue(true);
    }

    return JSON.stringify({ ok: true, to, cc });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── PDFs Gerados ────────────────────────────────────────────
function getPDFsGerados(token) {
  try {
    _getUser(token);
    const ss  = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const log = ss.getSheetByName('log_pdfs');
    if (!log) return JSON.stringify({ ok: true, rows: [] });
    const data = log.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ ok: true, rows: [] });
    const tz = Session.getScriptTimeZone();
    const MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
    const fmtMes = v => v instanceof Date ? MESES_PT[v.getMonth()] : String(v || '');
    const fmtAno = v => v instanceof Date ? String(v.getFullYear())  : String(v || '');
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[7]) continue; // sem fileId — linha inválida
      rows.push({
        ts:          r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : String(r[0]||''),
        geradoPor:   String(r[1]||''),
        tipo:        String(r[2]||''),
        origemBolsa: String(r[3]||''),
        mes:         fmtMes(r[4]),
        ano:         fmtAno(r[5]),
        titulo:      String(r[6]||''),
        fileId:      String(r[7]||''),
        viewUrl:     String(r[8]||''),
        chavePDF:    String(r[9]||''),
      });
    }
    rows.reverse();
    return JSON.stringify({ ok: true, rows });
  } catch(e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ─── Envio de e-mail (PDF mesclado — destinatário livre) ─────
function enviarEmailMesclado(token, payloadJson) {
  try {
    const user = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão para enviar e-mails.');
    const { fileId, to, cc, origemBolsa, mes, ano, obsExtra } = JSON.parse(payloadJson);
    if (!to)     throw new Error('Destinatário não informado.');
    if (!fileId) throw new Error('PDF não encontrado.');

    const fixedCC = ['adriane@brasas.com', 'administrativo@brasas.com'];
    const extraCC = cc ? cc.split(',').map(e => e.trim()).filter(Boolean) : [];
    const ccFinal = [...new Set([...extraCC, ...fixedCC])].join(',');

    const inlineImages = {};
    let logoSrc = '';
    try {
      const lb = DriveApp.getFileById(BRASAS_EMAIL_LOGO_FILE_ID).getBlob();
      lb.setName('brasasLogo');
      inlineImages.brasasLogo = lb;
      logoSrc = 'cid:brasasLogo';
    } catch(e) {}

    const subject  = `Relatório de Frequência e Aproveitamento - BRASAS - ${origemBolsa}`;
    const htmlBody = _buildEmailHtml(origemBolsa, mes, ano, obsExtra, logoSrc);

    GmailApp.sendEmail(to, subject, '', {
      htmlBody,
      cc:          ccFinal || undefined,
      attachments: [DriveApp.getFileById(fileId).getAs('application/pdf')],
      name:        'Administrativo BRASAS',
      replyTo:     'administrativo@brasas.com',
      inlineImages: Object.keys(inlineImages).length ? inlineImages : undefined,
    });

    return JSON.stringify({ ok: true, to });
  } catch(e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

function _buildEmailHtml(origemBolsa, mes, ano, obsExtra, logoSrc) {
  const logoHtml = logoSrc
    ? `<img src="${logoSrc}" alt="BRASAS" style="height:48px;display:block;margin-bottom:12px">`
    : `<div style="font-size:20px;font-weight:700;color:#fff;letter-spacing:.5px;margin-bottom:12px">BRASAS</div>`;

  const obsHtml = obsExtra
    ? `<p style="margin:18px 0 0;font-size:14px;color:#334155;line-height:1.65">${obsExtra}</p>`
    : '';

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;margin:0;padding:0;background:#fff">
  <div style="max-width:600px;margin:0 auto;padding:32px 16px">

    <div style="background:#1d3557;padding:28px 32px;border-radius:10px 10px 0 0">
      ${logoHtml}
      <div style="font-size:16px;font-weight:700;color:#fff;margin-bottom:4px">Relatório de Frequência e Aproveitamento</div>
      <div style="font-size:13px;color:rgba(255,255,255,.65)">BRASAS English Course &nbsp;·&nbsp; ${mes} - ${ano}</div>
    </div>

    <div style="padding:28px 32px 32px;background:#fff;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px">
      <p style="margin:0 0 14px;font-size:14px;color:#334155">Olá!</p>
      <p style="margin:0;font-size:14px;color:#334155;line-height:1.65">
        Segue em anexo o <strong>Relatório de Frequência e Aproveitamento</strong> dos alunos bolsistas de
        <strong>${origemBolsa}</strong>, referente ao mês de <strong>${mes} - ${ano}</strong>.
      </p>
      ${obsHtml}
    </div>

  </div>
</body></html>`;
}

// ─── Bloqueio após 5º dia útil ───────────────────────────────
function _isPastFifthBusinessDay() {
  const cache     = CacheService.getScriptCache();
  const ferCached = cache.get('feriados_v1');
  const feriados  = ferCached ? JSON.parse(ferCached) : [];

  const now   = new Date();
  const year  = now.getFullYear();
  const month = now.getMonth();
  const today = now.getDate();

  let bizDays = 0;
  for (let d = 1; d <= today; d++) {
    const dt  = new Date(year, month, d);
    if (dt.getDay() === 0 || dt.getDay() === 6) continue; // fim de semana
    const fmt = String(d).padStart(2,'0') + '/' + String(month+1).padStart(2,'0') + '/' + year;
    if (feriados.includes(fmt)) continue; // feriado
    if (++bizDays >= 5) return true;
  }
  return false;
}

// ─── Helpers ─────────────────────────────────────────────────
function _fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(val);
}
