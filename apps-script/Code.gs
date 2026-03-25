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
  'finserv', 'white_collars',
  'Edge Cases 1', 'Edge Cases 2'
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
  if (params.debug === 'msgs_month') {
    return ContentService
      .createTextOutput(JSON.stringify(getDebugMsgsByMonth(params.month || 'MAR')))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (params.debug === 'sheets') {
    return ContentService
      .createTextOutput(JSON.stringify(getDebugSheets()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify(getData()))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDebugSheets() {
  var result = [];
  MONTHS_CONFIG.forEach(function(monthCfg) {
    var ss = SpreadsheetApp.openById(monthCfg.spreadsheetId);
    ss.getSheets().forEach(function(sheet) {
      var name = sheet.getName();
      var vals = sheet.getDataRange().getValues();
      var row0 = vals.length > 0 ? vals[0].map(function(h){return String(h).trim();}).filter(Boolean) : [];
      var row1 = vals.length > 1 ? vals[1].map(function(h){return String(h).trim();}).filter(Boolean) : [];
      var ignored = IGNORED_SHEETS.indexOf(name) > -1;
      result.push({ month: monthCfg.id, name: name, ignored: ignored, row1: row0, row2: row1 });
    });
  });
  return result;
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
      if (result) { result.monthId = monthCfg.id; result.type = getCampaignType(result.name); campaigns.push(result); }
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

  var reach = computeAllReachStats();

  var attioMqls = null;
  try { attioMqls = getAttioMqls(); } catch(e) { attioMqls = null; }

  return {
    lastUpdated: new Date().toISOString(),
    months: monthsMeta,
    totals: totals,
    campaigns: campaigns,
    reach: reach,
    attioMqls: attioMqls
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
  nome:    ['nome', 'lead_nome', 'nomelead', 'full name', 'nome do participante'],
  empresa: ['empresa', 'company_name', 'nomeempresa', 'company', 'nome da empresa'],
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

function hasRequiredCols(headers) {
  return findColIdx(headers, COL_ALIASES.nome)    > -1 &&
         findColIdx(headers, COL_ALIASES.empresa) > -1 &&
         findColIdx(headers, COL_ALIASES.bdr)     > -1;
}

function isCampaignSheet(name, headers) {
  if (IGNORED_SHEETS.indexOf(name) > -1) return false;
  return hasRequiredCols(headers);
}

// Returns { headers, dataStart } trying row 0 first, then row 1
function detectHeaders(allValues) {
  if (allValues.length < 2) return null;
  var h0 = allValues[0].map(function(h) { return String(h).trim(); });
  var h1 = allValues[1].map(function(h) { return String(h).trim(); });
  if (hasRequiredCols(h0)) return { headers: h0, dataStart: 1 };
  if (hasRequiredCols(h1)) return { headers: h1, dataStart: 2 };
  return null;
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
  if (allValues.length < 2) return null;

  var detected = detectHeaders(allValues);
  if (!detected || !isCampaignSheet(name, detected.headers)) return null;
  var headers = detected.headers;
  var dataStart = detected.dataStart;

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

  for (var r = dataStart; r < allValues.length; r++) {
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

function getCampaignType(name) {
  if (name.toLowerCase().indexOf('outbound de outras') > -1) return 'outros';
  if (name.toLowerCase().indexOf('cold') > -1) return 'cold';
  return 'engaged';
}

// ============================================================
// REACH STATS (empresas/mensagens por mês calendario, cross-campanha)
// ============================================================

// Single pass over all sheets — computes reach for ALL months at once,
// broken down by source campaign month (e.g. FEB lists that messaged in MAR)
function computeAllReachStats() {
  var acc = {};
  MONTHS_CONFIG.forEach(function(m) {
    acc[m.id] = {
      rangeStart: m.sprints[0].start,
      rangeEnd:   m.sprints[m.sprints.length - 1].end,
      both: {}, cath: {}, mat: {},
      msgTotal: 0, msgCath: 0, msgMat: 0,
      by_source: {}
    };
    MONTHS_CONFIG.forEach(function(src) {
      acc[m.id].by_source[src.id] = { both: {}, msgTotal: 0 };
    });
  });

  MONTHS_CONFIG.forEach(function(monthCfg) {
    var srcId = monthCfg.id;
    var ss = SpreadsheetApp.openById(monthCfg.spreadsheetId);
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var name = sheet.getName();
      if (monthCfg.allowedSheets !== null && monthCfg.allowedSheets.indexOf(name) === -1) continue;
      var allValues = sheet.getDataRange().getValues();
      if (allValues.length < 2) continue;
      var detected = detectHeaders(allValues);
      if (!detected || !isCampaignSheet(name, detected.headers)) continue;
      var headers = detected.headers;
      var dataStart = detected.dataStart;

      var empresaIdx = findColIdx(headers, COL_ALIASES.empresa);
      var bdrIdx     = findColIdx(headers, COL_ALIASES.bdr);
      var headersLower = headers.map(function(h) { return String(h).toLowerCase().trim(); });
      var dataEnvioIdxs = [];
      for (var c = 0; c < headersLower.length; c++) {
        if (headersLower[c] === 'data envio') dataEnvioIdxs.push(c);
      }
      if (dataEnvioIdxs.length === 0) continue;

      for (var r = dataStart; r < allValues.length; r++) {
        var row = allValues[r];
        var empresa = String(row[empresaIdx] || '').trim();
        if (!empresa) continue;
        var bdr = String(row[bdrIdx] || '').trim().toUpperCase();
        var bdrKey = (bdr === 'CATH') ? 'cath' : 'mat';

        for (var p = 0; p < dataEnvioIdxs.length; p++) {
          var dv = row[dataEnvioIdxs[p]];
          if (!dv) continue;
          var d = (dv instanceof Date) ? dv : new Date(dv);
          if (isNaN(d.getTime())) continue;

          var mIds = Object.keys(acc);
          for (var mi = 0; mi < mIds.length; mi++) {
            var mId = mIds[mi];
            var a = acc[mId];
            if (d >= a.rangeStart && d <= a.rangeEnd) {
              a.both[empresa] = true;
              a[bdrKey][empresa] = true;
              a.msgTotal++;
              if (bdrKey === 'cath') a.msgCath++; else a.msgMat++;
              // track by source campaign month
              a.by_source[srcId].both[empresa] = true;
              a.by_source[srcId].msgTotal++;
            }
          }
        }
      }
    }
  });

  var result = {};
  MONTHS_CONFIG.forEach(function(m) {
    var a = acc[m.id];
    var bySource = {};
    MONTHS_CONFIG.forEach(function(src) {
      bySource[src.id] = {
        companies: Object.keys(a.by_source[src.id].both).length,
        messages:  a.by_source[src.id].msgTotal
      };
    });
    result[m.id] = {
      companies:      Object.keys(a.both).length,
      companies_cath: Object.keys(a.cath).length,
      companies_mat:  Object.keys(a.mat).length,
      messages:       a.msgTotal,
      messages_cath:  a.msgCath,
      messages_mat:   a.msgMat,
      by_source:      bySource
    };
  });
  return result;
}

// ============================================================
// ATTIO INTEGRATION
// ============================================================

var ATTIO_OUTBOUND_SOURCES = ['Outbound'];

function getAttioMqls() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ATTIO_API_KEY');
  if (!apiKey) return null;

  // Build year-month prefix → monthId map (e.g. '2026-03' → 'MAR')
  // Usa UTC para evitar bug de timezone no Apps Script
  var prefixToId = {};
  MONTHS_CONFIG.forEach(function(m) {
    var d = m.sprints[0].start;
    var prefix = d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2);
    prefixToId[prefix] = m.id;
  });

  var result = { all: { total: 0, cath: 0, mat: 0 } };
  MONTHS_CONFIG.forEach(function(m) { result[m.id] = { total: 0, cath: 0, mat: 0 }; });

  var offset = 0;
  var hasMore = true;

  while (hasMore) {
    var resp = UrlFetchApp.fetch('https://api.attio.com/v2/objects/deals/records/query', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ limit: 500, offset: offset }),
      muteHttpExceptions: true
    });

    var data = JSON.parse(resp.getContentText());
    var records = data.data || [];
    if (records.length === 0) { hasMore = false; break; }

    for (var i = 0; i < records.length; i++) {
      var vals = records[i].values || {};

      // Filter: must be outbound source
      var srcArr = vals.deal_source || [];
      var source = (srcArr[0] && srcArr[0].option) ? srcArr[0].option.title : null;
      if (ATTIO_OUTBOUND_SOURCES.indexOf(source) === -1) continue;

      // Must have date_mql (confirma que é um MQL)
      var dateMqlArr = vals.date_mql || [];
      if (!dateMqlArr[0] || !dateMqlArr[0].value) continue;

      // Agrupar por created_at do deal
      var prefix = String(records[i].created_at || '').slice(0, 7); // "2026-03"
      if (!prefix) continue;

      // BDR — excluir deals sem BDR
      var bdrArr = vals.bdr_associated || [];
      var bdrName = (bdrArr[0] && bdrArr[0].option) ? bdrArr[0].option.title : null;
      if (!bdrName) continue;
      var bdrKey = bdrName === 'Cath' ? 'cath' : (bdrName === 'Mateus Girardi' ? 'mat' : null);

      result.all.total++;
      if (bdrKey) result.all[bdrKey]++;

      var monthId = prefixToId[prefix];
      if (monthId && result[monthId]) {
        result[monthId].total++;
        if (bdrKey) result[monthId][bdrKey]++;
      }
    }

    offset += records.length;
    if (records.length < 500) hasMore = false;
  }

  return result;
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
      if (allValues.length < 2) continue;
      var detected = detectHeaders(allValues);
      if (!detected || !isCampaignSheet(name, detected.headers)) continue;
      var headers = detected.headers;
      var dataStart = detected.dataStart;
      var nomeIdx    = findColIdx(headers, COL_ALIASES.nome);
      var empresaIdx = findColIdx(headers, COL_ALIASES.empresa);
      var bdrIdx     = findColIdx(headers, COL_ALIASES.bdr);
      var statusIdx  = findColIdx(headers, COL_ALIASES.status);
      for (var r = dataStart; r < allValues.length; r++) {
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

// Returns unique companies messaged in a given calendar month (by date of send),
// across ALL spreadsheets — so FEB campaigns messaged in MAR are counted.
function getDebugMsgsByMonth(targetMonthId) {
  var targetCfg = null;
  for (var i = 0; i < MONTHS_CONFIG.length; i++) {
    if (MONTHS_CONFIG[i].id === targetMonthId) { targetCfg = MONTHS_CONFIG[i]; break; }
  }
  if (!targetCfg) return { error: 'Month not found: ' + targetMonthId };

  // Date range: first sprint start to last sprint end of target month
  var sprints = targetCfg.sprints;
  var rangeStart = sprints[0].start;
  var rangeEnd   = sprints[sprints.length - 1].end;

  var cathEmpresas = {};
  var matEmpresas  = {};
  var bothEmpresas = {};
  var details = [];
  var msgTotal = 0;
  var msgCath  = 0;
  var msgMat   = 0;

  MONTHS_CONFIG.forEach(function(monthCfg) {
    var ss = SpreadsheetApp.openById(monthCfg.spreadsheetId);
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var name = sheet.getName();
      if (monthCfg.allowedSheets !== null && monthCfg.allowedSheets.indexOf(name) === -1) continue;
      var allValues = sheet.getDataRange().getValues();
      if (allValues.length < 2) continue;
      var detected = detectHeaders(allValues);
      if (!detected || !isCampaignSheet(name, detected.headers)) continue;
      var headers = detected.headers;
      var dataStart = detected.dataStart;

      var empresaIdx = findColIdx(headers, COL_ALIASES.empresa);
      var bdrIdx     = findColIdx(headers, COL_ALIASES.bdr);
      var headersLower = headers.map(function(h) { return String(h).toLowerCase().trim(); });
      var dataEnvioIdxs = [];
      for (var c = 0; c < headersLower.length; c++) {
        if (headersLower[c] === 'data envio') dataEnvioIdxs.push(c);
      }
      if (dataEnvioIdxs.length === 0) continue;

      for (var r = dataStart; r < allValues.length; r++) {
        var row = allValues[r];
        var empresa = String(row[empresaIdx] || '').trim();
        if (!empresa) continue;
        var bdr = String(row[bdrIdx] || '').trim().toUpperCase();
        var bdrKey = (bdr === 'CATH') ? 'cath' : 'mat';

        var msgInMonth = false;
        for (var p = 0; p < dataEnvioIdxs.length; p++) {
          var dv = row[dataEnvioIdxs[p]];
          if (!dv) continue;
          var d = (dv instanceof Date) ? dv : new Date(dv);
          if (isNaN(d.getTime())) continue;
          if (d >= rangeStart && d <= rangeEnd) { msgInMonth = true; break; }
        }

        if (msgInMonth) {
          bothEmpresas[empresa] = true;
          if (bdrKey === 'cath') cathEmpresas[empresa] = true;
          else                   matEmpresas[empresa]  = true;

          // Count all messages sent in this month for this row
          for (var p2 = 0; p2 < dataEnvioIdxs.length; p2++) {
            var dv2 = row[dataEnvioIdxs[p2]];
            if (!dv2) continue;
            var d2 = (dv2 instanceof Date) ? dv2 : new Date(dv2);
            if (isNaN(d2.getTime())) continue;
            if (d2 >= rangeStart && d2 <= rangeEnd) {
              msgTotal++;
              if (bdrKey === 'cath') msgCath++; else msgMat++;
            }
          }
          details.push({ campaign: name, empresa: empresa, bdr: bdr });
        }
      }
    }
  });

  // Deduplicate details per empresa+bdr
  var seen = {};
  var dedupedDetails = [];
  details.forEach(function(d) {
    var key = d.empresa + '|' + d.bdr;
    if (!seen[key]) { seen[key] = true; dedupedDetails.push(d); }
  });

  return {
    month: targetMonthId,
    total:  Object.keys(bothEmpresas).length,
    cath:   Object.keys(cathEmpresas).length,
    mat:    Object.keys(matEmpresas).length,
    overlap: Object.keys(cathEmpresas).filter(function(k) { return matEmpresas[k]; }).length,
    msgs_total: msgTotal,
    msgs_cath:  msgCath,
    msgs_mat:   msgMat,
    empresas: Object.keys(bothEmpresas).sort(),
    cath_empresas: Object.keys(cathEmpresas).sort(),
    mat_empresas:  Object.keys(matEmpresas).sort(),
    details: dedupedDetails
  };
}
