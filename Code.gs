// ============================================================
//  BRASAS Bolsistas — Code.gs
// ============================================================

const USERS_SHEET_ID     = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';
const BOLSISTAS_SHEET_ID = '1G7RfHP_8j7-6VPqC8ReYvpScyWvBnVyRXu8QAdeL4bQ';

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

  const ss        = SpreadsheetApp.openById(USERS_SHEET_ID);
  const sessSheet = ss.getSheetByName('SESSOES');
  if (!sessSheet) throw new Error('Sessão inválida.');

  const now  = new Date();
  const rows = sessSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(token)) continue;
    const expira = rows[i][6] ? new Date(rows[i][6]) : null;
    if (expira && expira < now) throw new Error('Sessão expirada. Acesse pelo hub novamente.');

    const role = String(rows[i][3]).trim().toLowerCase();
    const email = String(rows[i][1]).trim().toLowerCase();

    if (!_hasAccess(ss, role, email)) {
      throw new Error('Sem permissão para acessar o painel de bolsistas.');
    }

    return {
      email,
      nome:         String(rows[i][2]).trim(),
      role,
      unidade:      String(rows[i][4]).trim(),
      canEdit:      EDIT_ROLES.includes(role),
      canSendEmail: EMAIL_ROLES.includes(role),
      isAdmin:      ADMIN_ROLES.includes(role),
    };
  }

  throw new Error('Sessão não encontrada. Acesse pelo hub.');
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
      if (params.mes        && norm(mes)               !== norm(params.mes))         continue;
      if (params.ano        && String(ano)              !== String(params.ano))        continue;
      if (params.unidade    && norm(unidade)            !== norm(params.unidade))      continue;
      if (params.origemBolsa && norm(String(r[7]||'')) !== norm(params.origemBolsa)) continue;
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
        obsExtrasEmail:    String(r[27] || ''),
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

    // col indexes são 1-based
    const FIELD_COL = {
      diasAssistidos: 16, // P
      mt:             18, // R
      wt:             19, // S
      oc:             20, // T
      ot:             21, // U
      aproveitamento: 22, // V
      observacoes:    23, // W
      obsExtrasEmail: 28, // AB
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

// ─── Geração de PDF ──────────────────────────────────────────
function gerarPDF(token, chavePDF) {
  try {
    const user = _getUser(token);
    if (!user.canSendEmail) throw new Error('Sem permissão para gerar PDF.');

    const ss    = SpreadsheetApp.openById(BOLSISTAS_SHEET_ID);
    const sheet = ss.getSheetByName('Bolsistas App');
    const data  = sheet.getDataRange().getValues();

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][24] || '').trim() === chavePDF.trim()) rows.push(data[i]);
    }
    if (!rows.length) throw new Error('Nenhum aluno encontrado para esta chave PDF.');

    const origemBolsa = String(rows[0][7] || '').trim();
    const parts       = chavePDF.split(' - ');
    const ano         = parts[0] || '';
    const mes         = parts[1] || '';

    // Parceiro
    const parcSheet = ss.getSheetByName('Parceiros');
    let logo = '';
    if (parcSheet) {
      const pd = parcSheet.getDataRange().getValues();
      for (let i = 1; i < pd.length; i++) {
        if (String(pd[i][0] || '').trim() === origemBolsa) { logo = String(pd[i][1] || '').trim(); break; }
      }
    }

    // Cria Google Doc
    const docTitle = `Relatório Bolsistas — ${chavePDF}`;
    const doc      = DocumentApp.create(docTitle);
    const body     = doc.getBody();
    body.clear();

    // Configuração de página
    body.setPageWidth(794).setPageHeight(1123); // A4 aproximado em pontos
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

    // Cabeçalho
    const titlePara = body.appendParagraph(`Relatório de Bolsistas — ${origemBolsa}`);
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    const subPara = body.appendParagraph(`${mes} — ${ano}`);
    subPara.setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph('');

    // Tabela
    const hdrs = [
      'Unidade — Nome', 'Origem da Bolsa', '% Bolsa', 'Book',
      'Frequência — Horário', 'Dias Prev.', 'Dias Assist.',
      'Data 1ª Aula', 'Início Book', 'Prev. Conclusão',
      'Valor', 'MT', 'WT', 'OC', 'OT', 'Aproveitamento', 'Observações',
    ];

    const tableData = [hdrs];
    rows.forEach(r => {
      tableData.push([
        `${r[2]} — ${r[5]}`,
        String(r[7]  || ''),
        r[6] !== '' ? r[6] + '%' : '',
        String(r[9]  || ''),
        `${r[10] || ''} — ${r[13] || ''}`,
        String(r[14] || ''),
        String(r[15] || ''),
        _fmtDate(r[8]),
        _fmtDate(r[11]),
        _fmtDate(r[12]),
        r[16] !== '' ? 'R$ ' + r[16] : '',
        String(r[17] || ''),
        String(r[18] || ''),
        String(r[19] || ''),
        String(r[20] || ''),
        String(r[21] || ''),
        String(r[22] || ''),
      ]);
    });

    const table = body.appendTable(tableData);
    table.setBorderWidth(0.5);
    const headerRow = table.getRow(0);
    for (let c = 0; c < hdrs.length; c++) {
      headerRow.getCell(c).setBackgroundColor('#0F2035')
        .editAsText().setForegroundColor('#FFFFFF').setBold(true).setFontSize(8);
    }
    for (let r = 1; r < table.getNumRows(); r++) {
      for (let c = 0; c < hdrs.length; c++) {
        table.getCell(r, c).editAsText().setFontSize(8);
        if (r % 2 === 0) table.getCell(r, c).setBackgroundColor('#F1F5F9');
      }
    }

    doc.saveAndClose();

    // Exporta como PDF
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs('application/pdf').setName(docTitle + '.pdf');

    let folder;
    const folders = DriveApp.getFoldersByName('Relatórios Bolsistas');
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Relatórios Bolsistas');

    // Remove versão anterior se existir
    const existing = folder.getFilesByName(docTitle + '.pdf');
    while (existing.hasNext()) existing.next().setTrashed(true);

    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    docFile.setTrashed(true); // descarta o Doc temporário

    return JSON.stringify({
      ok:          true,
      fileId:      pdfFile.getId(),
      fileName:    pdfFile.getName(),
      viewUrl:     pdfFile.getUrl(),
      totalAlunos: rows.length,
    });
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

    const subject  = `Relatório de Bolsistas — ${origemBolsa} — ${mes} ${ano}`;
    const htmlBody = _buildEmailHtml(origemBolsa, mes, ano, rows, obsExtra);

    GmailApp.sendEmail(to, subject, '', {
      htmlBody,
      cc:          cc || undefined,
      attachments,
      name:        'BRASAS Bolsistas',
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
  const rowsHtml = rows.map((r, idx) => `
    <tr style="background:${idx % 2 === 0 ? '#fff' : '#f8fafc'}">
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0"><b>${r[2]}</b><br><span style="color:#475569;font-size:11px">${r[5]}</span></td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[6] !== '' ? r[6] + '%' : ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[9] || ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[10] || ''}<br><span style="font-size:11px;color:#64748b">${r[13] || ''}</span></td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[14] !== '' ? r[14] : ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[15] !== '' ? r[15] : ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[17] || ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[18] || ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[19] || ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center">${r[20] || ''}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;text-align:center"><b>${r[21] || ''}</b></td>
      <td style="padding:7px 10px;border-bottom:1px solid #e2e8f0;font-size:11px;color:#475569">${r[22] || ''}</td>
    </tr>`).join('');

  return `<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f1f5f9">
  <div style="max-width:960px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08)">
    <div style="background:#0F2035;padding:28px 32px;display:flex;align-items:center;justify-content:space-between">
      <div>
        <h1 style="margin:0;color:#fff;font-size:20px">Relatório de Bolsistas</h1>
        <p style="margin:6px 0 0;color:rgba(255,255,255,.7);font-size:14px">${origemBolsa}</p>
      </div>
      <div style="color:rgba(255,255,255,.8);font-size:14px;text-align:right">${mes}<br>${ano}</div>
    </div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:#162D4A;color:#fff">
            <th style="padding:10px;text-align:left;min-width:160px">Unidade — Nome</th>
            <th style="padding:10px;text-align:center">% Bolsa</th>
            <th style="padding:10px;text-align:center">Book</th>
            <th style="padding:10px;text-align:center">Freq — Horário</th>
            <th style="padding:10px;text-align:center">Prev.</th>
            <th style="padding:10px;text-align:center">Assist.</th>
            <th style="padding:10px;text-align:center">MT</th>
            <th style="padding:10px;text-align:center">WT</th>
            <th style="padding:10px;text-align:center">OC</th>
            <th style="padding:10px;text-align:center">OT</th>
            <th style="padding:10px;text-align:center">Aprov.</th>
            <th style="padding:10px;text-align:left;min-width:200px">Observações</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
    ${obsExtra ? `<div style="padding:20px 32px;border-top:1px solid #e2e8f0;font-size:13px;color:#475569"><b>Observações gerais:</b> ${obsExtra}</div>` : ''}
    <div style="padding:16px 32px;background:#f8fafc;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0">
      Enviado por BRASAS Analytics &bull; ${new Date().toLocaleDateString('pt-BR')} &bull; Total: ${rows.length} aluno${rows.length !== 1 ? 's' : ''}
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
