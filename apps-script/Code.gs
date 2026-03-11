// ============================================================
// CONFIGURAÇÃO — atualizar no início de cada mês
// ============================================================

var SPRINT_CONFIG = [
  { name: 'S1', start: new Date('2026-03-01'), end: new Date('2026-03-08') },
  { name: 'S2', start: new Date('2026-03-09'), end: new Date('2026-03-16') },
  { name: 'S3', start: new Date('2026-03-17'), end: new Date('2026-03-24') },
  { name: 'S4', start: new Date('2026-03-25'), end: new Date('2026-04-01') }
];

var SPREADSHEET_ID = '1evy8peuLyilrhTndKHMbUgqDpWl55p8pCARax24TxGA';

// Abas que NÃO são campanhas (ignoradas)
var IGNORED_SHEETS = [
  'Métricas', 'CADENCIA', 'MENSAGEM', 'mat_analise>',
  'EngagementSample', 'CompaniesEngagement', 'table_import',
  'de-para', 'sales_pipe', 'clean_formula',
  'Companies CRM 20d9777bf857815ba546f7768e4ca2cc', 'View',
  'finserv', 'white_collars'
];

// ============================================================
// ENDPOINT PRINCIPAL
// ============================================================

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify(getData()))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// COLETA E PROCESSAMENTO
// ============================================================

function getData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var campaigns = [];

  for (var i = 0; i < sheets.length; i++) {
    var result = processCampaignSheet(sheets[i]);
    if (result) campaigns.push(result);
  }

  var totals = calculateTotals(campaigns);

  // Remove listas internas antes de retornar
  campaigns.forEach(function(c) { delete c._empresasCath; delete c._empresasMat; });

  return {
    lastUpdated: new Date().toISOString(),
    sprints: {
      S1: { start: '2026-03-01', end: '2026-03-08', label: '01/03 – 08/03' },
      S2: { start: '2026-03-09', end: '2026-03-16', label: '09/03 – 16/03' },
      S3: { start: '2026-03-17', end: '2026-03-24', label: '17/03 – 24/03' },
      S4: { start: '2026-03-25', end: '2026-04-01', label: '25/03 – 01/04' }
    },
    totals: totals,
    campaigns: campaigns
  };
}

function getSprint(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return null;
  for (var i = 0; i < SPRINT_CONFIG.length; i++) {
    var s = SPRINT_CONFIG[i];
    if (date >= s.start && date <= s.end) return s.name;
  }
  return null;
}

// Nomes alternativos aceitos para cada campo-chave
var COL_ALIASES = {
  nome:    ['nome', 'lead_nome'],
  empresa: ['empresa', 'company_name'],
  bdr:     ['bdr'],
  status:  ['status geral']
};

function findColIdx(headers, aliases) {
  var flat = headers.map(function(h) { return String(h).toLowerCase().trim(); });
  for (var i = 0; i < aliases.length; i++) {
    var idx = flat.indexOf(aliases[i]);
    if (idx > -1) return idx;
  }
  return -1;
}

function isCampaignSheet(name, headers) {
  if (IGNORED_SHEETS.indexOf(name) > -1) return false;
  return findColIdx(headers, COL_ALIASES.nome)    > -1 &&
         findColIdx(headers, COL_ALIASES.empresa) > -1 &&
         findColIdx(headers, COL_ALIASES.bdr)     > -1;
}

function emptySprintObj() {
  return { all: 0, S1: 0, S2: 0, S3: 0, S4: 0 };
}

function emptyBDRMetric() {
  return { total: emptySprintObj(), cath: emptySprintObj(), mat: emptySprintObj() };
}

function addCount(obj, bdrKey, sprint) {
  obj.total.all++;
  obj[bdrKey].all++;
  if (sprint && obj.total[sprint] !== undefined) {
    obj.total[sprint]++;
    obj[bdrKey][sprint]++;
  }
}

function processCampaignSheet(sheet) {
  var name = sheet.getName();
  var allValues = sheet.getDataRange().getValues();
  if (allValues.length < 3) return null;

  var headers = allValues[1].map(function(h) { return String(h).trim(); });
  if (!isCampaignSheet(name, headers)) return null;

  var nomeIdx    = findColIdx(headers, COL_ALIASES.nome);
  var empresaIdx = findColIdx(headers, COL_ALIASES.empresa);
  var bdrIdx     = findColIdx(headers, COL_ALIASES.bdr);
  var statusIdx  = findColIdx(headers, COL_ALIASES.status);

  // Todas as colunas "Data Envio"
  var dataEnvioIdxs = [];
  var headersLower = headers.map(function(h) { return String(h).toLowerCase().trim(); });
  for (var i = 0; i < headersLower.length; i++) {
    if (headersLower[i] === 'data envio') dataEnvioIdxs.push(i);
  }

  var empresasCath = {};
  var empresasMat  = {};
  var contatos     = { total: 0, cath: 0, mat: 0 };
  var mensagens    = emptyBDRMetric();
  var mqls         = emptyBDRMetric();

  for (var r = 2; r < allValues.length; r++) {
    var row = allValues[r];
    var nome = String(row[nomeIdx] || '').trim();
    if (!nome) continue;

    var empresa     = String(row[empresaIdx] || '').trim();
    var bdr         = String(row[bdrIdx]     || '').trim().toUpperCase();
    var statusGeral = String(statusIdx >= 0 ? row[statusIdx] : '').trim();
    if (!empresa) continue;

    var bdrKey = (bdr === 'CATH') ? 'cath' : 'mat';

    // Empresas únicas por BDR (dentro desta campanha)
    if (bdrKey === 'cath') empresasCath[empresa] = true;
    else                   empresasMat[empresa]  = true;

    // Contatos
    contatos.total++;
    contatos[bdrKey]++;

    // Mensagens
    for (var p = 0; p < dataEnvioIdxs.length; p++) {
      var dateVal = row[dataEnvioIdxs[p]];
      if (!dateVal) continue;
      var date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
      if (isNaN(date.getTime())) continue;
      addCount(mensagens, bdrKey, getSprint(date));
    }

    // MQLs
    if (statusGeral === 'MQL') {
      var mqlDate = null;
      for (var q = 0; q < dataEnvioIdxs.length; q++) {
        var dv = row[dataEnvioIdxs[q]];
        if (dv) { mqlDate = (dv instanceof Date) ? dv : new Date(dv); break; }
      }
      addCount(mqls, bdrKey, mqlDate ? getSprint(mqlDate) : null);
    }
  }

  var allEmpresas = {};
  Object.keys(empresasCath).forEach(function(k) { allEmpresas[k] = true; });
  Object.keys(empresasMat).forEach(function(k)  { allEmpresas[k] = true; });

  var empresas = {
    total: Object.keys(allEmpresas).length,
    cath:  Object.keys(empresasCath).length,
    mat:   Object.keys(empresasMat).length,
    list:  Object.keys(allEmpresas).sort()
  };

  var mqlCount          = mqls.total.all;
  var msgCount          = mensagens.total.all;
  var contatoPerEmpresa = empresas.total > 0 ? contatos.total / empresas.total : 0;

  return {
    name: name,
    // listas brutas para dedup global (removidas antes de retornar ao cliente)
    _empresasCath: empresasCath,
    _empresasMat:  empresasMat,
    empresas:  empresas,
    contatos:  contatos,
    mensagens: mensagens,
    mqls:      mqls,
    conversion: {
      pct_empresa:          empresas.total > 0 ? mqlCount / empresas.total : 0,
      msgs_per_mql:         mqlCount > 0       ? round2(msgCount / mqlCount) : 0,
      msgs_per_mql_contato: (mqlCount > 0 && contatoPerEmpresa > 0)
                              ? round2(msgCount / mqlCount / contatoPerEmpresa) : 0,
      contato_por_empresa:  round2(contatoPerEmpresa),
      pct_conv_stakeholder: contatos.total > 0 ? mqlCount / contatos.total  : 0,
      pct_conv_empresa:     empresas.total > 0 ? mqlCount / empresas.total  : 0
    }
  };
}

function calculateTotals(campaigns) {
  // Deduplicação global — mesma empresa em 2 campanhas = conta 1 vez
  var globalCath = {};
  var globalMat  = {};

  var contatos  = { total: 0, cath: 0, mat: 0 };
  var mensagens = emptyBDRMetric();
  var mqls      = emptyBDRMetric();

  campaigns.forEach(function(c) {
    Object.keys(c._empresasCath || {}).forEach(function(k) { globalCath[k] = true; });
    Object.keys(c._empresasMat  || {}).forEach(function(k) { globalMat[k]  = true; });

    contatos.total += c.contatos.total;
    contatos.cath  += c.contatos.cath;
    contatos.mat   += c.contatos.mat;

    ['total','cath','mat'].forEach(function(bdr) {
      ['all','S1','S2','S3','S4'].forEach(function(s) {
        mensagens[bdr][s] += c.mensagens[bdr][s];
        mqls[bdr][s]      += c.mqls[bdr][s];
      });
    });
  });

  var globalAll = {};
  Object.keys(globalCath).forEach(function(k) { globalAll[k] = true; });
  Object.keys(globalMat).forEach(function(k)  { globalAll[k] = true; });

  var empresas = {
    total: Object.keys(globalAll).length,
    cath:  Object.keys(globalCath).length,
    mat:   Object.keys(globalMat).length,
    list:  Object.keys(globalAll).sort()
  };

  var mqlCount          = mqls.total.all;
  var msgCount          = mensagens.total.all;
  var contatoPerEmpresa = empresas.total > 0 ? contatos.total / empresas.total : 0;

  return {
    empresas:  empresas,
    contatos:  contatos,
    mensagens: mensagens,
    mqls:      mqls,
    conversion: {
      pct_empresa:          empresas.total > 0 ? mqlCount / empresas.total : 0,
      msgs_per_mql:         mqlCount > 0       ? round2(msgCount / mqlCount) : 0,
      msgs_per_mql_contato: (mqlCount > 0 && contatoPerEmpresa > 0)
                              ? round2(msgCount / mqlCount / contatoPerEmpresa) : 0,
      contato_por_empresa:  round2(contatoPerEmpresa),
      pct_conv_stakeholder: contatos.total > 0 ? mqlCount / contatos.total  : 0,
      pct_conv_empresa:     empresas.total > 0 ? mqlCount / empresas.total  : 0
    }
  };
}

function round2(val) {
  return Math.round(val * 100) / 100;
}
