// ============================================================
// CONFIGURAÇÃO — atualizar no início de cada mês
// ============================================================

var MONTHS_CONFIG = [
  {
    id: 'FEB',
    label: 'Fevereiro',
    spreadsheetId: '11ygMu2JIO7XCeu3a0tAlNAfvH_Z8egHA_3kNwnYYB2A',
    allowedSheets: [
      'State of Comp in Tech',
      'Webinar Janeiro',
      'Engaged Comp LKD',
      'Maturidade em Remuneracao',
      'Webinar Open Demo 2025'
    ],
    sprints: [
      { name: 'S1', label: '01/02 – 07/02', start: new Date('2026-02-01'), end: new Date('2026-02-07') },
      { name: 'S2', label: '08/02 – 14/02', start: new Date('2026-02-08'), end: new Date('2026-02-14') },
      { name: 'S3', label: '15/02 – 21/02', start: new Date('2026-02-15'), end: new Date('2026-02-21') },
      { name: 'S4', label: '22/02 – 28/02', start: new Date('2026-02-22'), end: new Date('2026-02-28') }
    ]
  },
  {
    id: 'MAR',
    label: 'Março',
    spreadsheetId: '1evy8peuLyilrhTndKHMbUgqDpWl55p8pCARax24TxGA',
    allowedSheets: null,
    sprints: [
      { name: 'S1', label: '01/03 – 08/03', start: new Date('2026-03-01'), end: new Date('2026-03-08') },
      { name: 'S2', label: '09/03 – 16/03', start: new Date('2026-03-09'), end: new Date('2026-03-16') },
      { name: 'S3', label: '17/03 – 24/03', start: new Date('2026-03-17'), end: new Date('2026-03-24') },
      { name: 'S4', label: '25/03 – 01/04', start: new Date('2026-03-25'), end: new Date('2026-04-01') }
    ]
  }
];

// Abas que NÃO são campanhas (usadas apenas para Mar / auto-detect)
var IGNORED_SHEETS = [
  'Métricas', 'CADENCIA', 'CADENCIAS', 'MENSAGEM', 'mat_analise>',
  'EngagementSample', 'CompaniesEngagement', 'table_import',
  'de-para', 'sales_pipe', 'clean_formula',
  'Companies CRM 20d9777bf857815ba546f7768e4ca2cc', 'View',
  'finserv', 'white_collars'
];

// ============================================================
// ENDPOINT PRINCIPAL
// ============================================================

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  if (params.debug === 'mqls') {
    return ContentService
      .createTextOutput(JSON.stringify(getDebugMqls()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify(getData()))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// COLETA E PROCESSAMENTO
// ============================================================

function getData() {
  var campaigns = [];

  for (var m = 0; m < MONTHS_CONFIG.length; m++) {
    var monthCfg = MONTHS_CONFIG[m];
    var ss = SpreadsheetApp.openById(monthCfg.spreadsheetId);
    var sheets = ss.getSheets();

    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();

      if (monthCfg.allowedSheets !== null) {
        if (monthCfg.allowedSheets.indexOf(sheetName) === -1) continue;
      }

      var result = processCampaignSheet(sheet, monthCfg.id);
      if (result) { result.monthId = monthCfg.id; campaigns.push(result); }
    }
  }

  var totals = calculateTotals(campaigns);
  campaigns.forEach(function(c) { delete c._empresasCath; delete c._empresasMat; });

  // Metadata de meses para o frontend
  var monthsMeta = {};
  MONTHS_CONFIG.forEach(function(monthCfg) {
    monthsMeta[monthCfg.id] = { label: monthCfg.label, sprints: {} };
    monthCfg.sprints.forEach(function(s) {
      monthsMeta[monthCfg.id].sprints[s.name] = {
        start: s.start.toISOString().slice(0, 10),
        end:   s.end.toISOString().slice(0, 10),
        label: s.label
      };
    });
  });

  return {
    lastUpdated: new Date().toISOString(),
    months: monthsMeta,
    totals: totals,
    campaigns: campaigns
  };
}

function getMonthSprint(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return null;
  for (var m = 0; m < MONTHS_CONFIG.length; m++) {
    var month = MONTHS_CONFIG[m];
    for (var s = 0; s < month.sprints.length; s++) {
      var sprint = month.sprints[s];
      if (date >= sprint.start && date <= sprint.end) {
        return { month: month.id, sprint: sprint.name };
      }
    }
  }
  return null;
}

// ============================================================
// DETECÇÃO DE COLUNAS
// ============================================================

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

// ============================================================
// ESTRUTURA DE MÉTRICAS (por mês + sprint)
// ============================================================

function emptyMonthSprints() {
  var obj = { all: 0 };
  MONTHS_CONFIG.forEach(function(m) {
    obj[m.id] = { all: 0 };
    m.sprints.forEach(function(s) { obj[m.id][s.name] = 0; });
  });
  return obj;
}

function emptyBDRMetric() {
  return { total: emptyMonthSprints(), cath: emptyMonthSprints(), mat: emptyMonthSprints() };
}

function addCount(obj, bdrKey, monthSprint, fallbackMonth) {
  obj.total.all++;
  obj[bdrKey].all++;
  // If date falls in a different month than the campaign's month,
  // count in campaign's month without a specific sprint (overflow)
  var m, s;
  if (monthSprint && fallbackMonth && monthSprint.month !== fallbackMonth) {
    m = fallbackMonth;
    s = null;
  } else if (monthSprint) {
    m = monthSprint.month;
    s = monthSprint.sprint;
  } else {
    m = fallbackMonth;
    s = null;
  }
  if (m && obj.total[m] !== undefined) {
    obj.total[m].all++;
    if (s && obj.total[m][s] !== undefined) obj.total[m][s]++;
    obj[bdrKey][m].all++;
    if (s && obj[bdrKey][m][s] !== undefined) obj[bdrKey][m][s]++;
  }
}

// ============================================================
// PROCESSAMENTO POR ABA
// ============================================================

function processCampaignSheet(sheet, monthId) {
  var name = sheet.getName();
  var allValues = sheet.getDataRange().getValues();
  if (allValues.length < 3) return null;

  var headers = allValues[1].map(function(h) { return String(h).trim(); });
  if (!isCampaignSheet(name, headers)) return null;

  var nomeIdx    = findColIdx(headers, COL_ALIASES.nome);
  var empresaIdx = findColIdx(headers, COL_ALIASES.empresa);
  var bdrIdx     = findColIdx(headers, COL_ALIASES.bdr);
  var statusIdx  = findColIdx(headers, COL_ALIASES.status);

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

    if (bdrKey === 'cath') empresasCath[empresa] = true;
    else                   empresasMat[empresa]  = true;

    contatos.total++;
    contatos[bdrKey]++;

    for (var p = 0; p < dataEnvioIdxs.length; p++) {
      var dateVal = row[dataEnvioIdxs[p]];
      if (!dateVal) continue;
      var date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
      if (isNaN(date.getTime())) continue;
      addCount(mensagens, bdrKey, getMonthSprint(date), monthId);
    }

    if (statusGeral.toUpperCase() === 'MQL') {
      var mqlDate = null;
      for (var q = 0; q < dataEnvioIdxs.length; q++) {
        var dv = row[dataEnvioIdxs[q]];
        if (dv) { mqlDate = (dv instanceof Date) ? dv : new Date(dv); break; }
      }
      addCount(mqls, bdrKey, mqlDate ? getMonthSprint(mqlDate) : null, monthId);
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

// ============================================================
// TOTAIS CONSOLIDADOS
// ============================================================

function calculateTotals(campaigns) {
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
      mensagens[bdr].all += c.mensagens[bdr].all;
      mqls[bdr].all      += c.mqls[bdr].all;
      MONTHS_CONFIG.forEach(function(m) {
        mensagens[bdr][m.id].all += c.mensagens[bdr][m.id].all;
        mqls[bdr][m.id].all      += c.mqls[bdr][m.id].all;
        m.sprints.forEach(function(s) {
          mensagens[bdr][m.id][s.name] += c.mensagens[bdr][m.id][s.name];
          mqls[bdr][m.id][s.name]      += c.mqls[bdr][m.id][s.name];
        });
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

// ============================================================
// DEBUG
// ============================================================

function getDebugMqls() {
  var rows = [];
  MONTHS_CONFIG.forEach(function(monthCfg) {
    var ss = SpreadsheetApp.openById(monthCfg.spreadsheetId);
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var name = sheet.getName();
      if (monthCfg.allowedSheets !== null && monthCfg.allowedSheets.indexOf(name) === -1) continue;
      var allValues = sheet.getDataRange().getValues();
      if (allValues.length < 3) continue;
      var headers = allValues[1].map(function(h) { return String(h).trim(); });
      if (!isCampaignSheet(name, headers)) continue;
      var nomeIdx    = findColIdx(headers, COL_ALIASES.nome);
      var empresaIdx = findColIdx(headers, COL_ALIASES.empresa);
      var bdrIdx     = findColIdx(headers, COL_ALIASES.bdr);
      var statusIdx  = findColIdx(headers, COL_ALIASES.status);
      for (var r = 2; r < allValues.length; r++) {
        var row = allValues[r];
        var statusStr = String(statusIdx >= 0 ? row[statusIdx] : '').trim();
        if (statusStr.toUpperCase() === 'MQL') {
          var nome = String(row[nomeIdx] || '').trim();
          var empresa = String(row[empresaIdx] || '').trim();
          rows.push({
            month: monthCfg.id, campaign: name,
            nome: nome, empresa: empresa,
            bdr: String(row[bdrIdx] || '').trim(),
            statusRaw: statusStr, skipped: (!nome || !empresa)
          });
        }
      }
    }
  });
  return { count: rows.length, mqls: rows };
}
