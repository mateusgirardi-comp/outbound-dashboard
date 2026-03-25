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

// Planilha de prospecção outbound
var LEADS_SPREADSHEET_ID = '1OAGg5rDXi-uBO6SIPi3qldyyAPZvtZxuTz6XaPoFtoM';
var LEADS_SHEET_GID      = 487542244;

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
  if (params.clear_cache === '1') {
    CacheService.getScriptCache().remove('attio_search_fallback');
    return ContentService.createTextOutput('{"cache":"cleared"}').setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'get_leads') {
    return ContentService
      .createTextOutput(JSON.stringify({ leads: getLeadsData(), opts: getLeadsValidationOptions() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (params.action === 'update_lead') {
    var row = parseInt(params.row, 10);
    var col = parseInt(params.col, 10);
    var value = params.value !== undefined ? params.value : '';
    if (row > 0 && col > 0) updateLeadCell(row, col, value);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
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

  var contacts = getContactsData();
  var uniqueCompanies = contacts.map(function(c){ return c.empresa; }).filter(function(v,i,a){ return a.indexOf(v) === i; });

  var attioCompanyLinks = null;
  try { attioCompanyLinks = getAttioCompanyLinks(uniqueCompanies); } catch(e) { attioCompanyLinks = null; }

  return {
    lastUpdated: new Date().toISOString(),
    months: monthsMeta,
    totals: totals,
    campaigns: campaigns,
    reach: reach,
    attioMqls: attioMqls,
    contacts: contacts,
    attioCompanyLinks: attioCompanyLinks
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
      conBoth: {}, conCath: {}, conMat: {},
      by_source: {}
    };
    MONTHS_CONFIG.forEach(function(src) {
      acc[m.id].by_source[src.id] = { both: {}, msgTotal: 0, conBoth: {} };
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
      var nomeIdx    = findColIdx(headers, COL_ALIASES.nome);
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
        var nome = nomeIdx >= 0 ? String(row[nomeIdx] || '').trim() : '';
        var conKey = empresa + '|' + nome;
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
              a.conBoth[conKey] = true;
              if (bdrKey === 'cath') a.conCath[conKey] = true; else a.conMat[conKey] = true;
              // track by source campaign month
              a.by_source[srcId].both[empresa] = true;
              a.by_source[srcId].msgTotal++;
              a.by_source[srcId].conBoth[conKey] = true;
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
        messages:  a.by_source[src.id].msgTotal,
        contacts:  Object.keys(a.by_source[src.id].conBoth).length
      };
    });
    result[m.id] = {
      companies:      Object.keys(a.both).length,
      companies_cath: Object.keys(a.cath).length,
      companies_mat:  Object.keys(a.mat).length,
      messages:       a.msgTotal,
      messages_cath:  a.msgCath,
      messages_mat:   a.msgMat,
      contacts:       Object.keys(a.conBoth).length,
      contacts_cath:  Object.keys(a.conCath).length,
      contacts_mat:   Object.keys(a.conMat).length,
      by_source:      bySource
    };
  });
  return result;
}

// ============================================================
// CONTACTS VIEW
// ============================================================

function getContactsData() {
  var contacts = [];

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

      var allValues = sheet.getDataRange().getValues();
      if (allValues.length < 2) continue;

      var detected = detectHeaders(allValues);
      if (!detected || !isCampaignSheet(sheetName, detected.headers)) continue;

      var headers    = detected.headers;
      var dataStart  = detected.dataStart;
      var nomeIdx    = findColIdx(headers, COL_ALIASES.nome);
      var empresaIdx = findColIdx(headers, COL_ALIASES.empresa);
      var bdrIdx     = findColIdx(headers, COL_ALIASES.bdr);
      var statusIdx  = findColIdx(headers, COL_ALIASES.status);

      var headersLower = headers.map(function(h) { return String(h).toLowerCase().trim(); });
      var dataEnvioIdxs = [];
      for (var c = 0; c < headersLower.length; c++) {
        if (headersLower[c] === 'data envio') dataEnvioIdxs.push(c);
      }

      for (var r = dataStart; r < allValues.length; r++) {
        var row     = allValues[r];
        var nome    = String(row[nomeIdx]    || '').trim();
        var empresa = String(row[empresaIdx] || '').trim();
        if (!nome || !empresa) continue;

        var bdrRaw = String(row[bdrIdx] || '').trim().toUpperCase();
        var bdr    = bdrRaw === 'CATH' ? 'Cath' : (bdrRaw === 'MAT' ? 'Mat' : bdrRaw);
        var status = String(statusIdx >= 0 ? row[statusIdx] : '').trim();

        var nMensagens = 0;
        var ultimaMensagem = null;
        for (var p = 0; p < dataEnvioIdxs.length; p++) {
          var dv = row[dataEnvioIdxs[p]];
          if (!dv) continue;
          var d = (dv instanceof Date) ? dv : new Date(dv);
          if (isNaN(d.getTime())) continue;
          nMensagens++;
          if (!ultimaMensagem || d > ultimaMensagem) ultimaMensagem = d;
        }

        contacts.push({
          empresa:          empresa,
          contato:          nome,
          bdr:              bdr,
          status:           status,
          lista:            sheetName,
          mes:              monthCfg.id,
          n_mensagens:      nMensagens,
          ultima_mensagem:  ultimaMensagem ? ultimaMensagem.toISOString().slice(0, 10) : null
        });
      }
    }
  }

  contacts.sort(function(a, b) {
    var ea = a.empresa.toLowerCase(), eb = b.empresa.toLowerCase();
    if (ea !== eb) return ea < eb ? -1 : 1;
    return a.contato.toLowerCase() < b.contato.toLowerCase() ? -1 : 1;
  });

  return contacts;
}

// ============================================================
// ATTIO COMPANY LINKS
// ============================================================

function normalizeCompanyName(name) {
  return String(name)
    .split(/\s*\|\s*/)[0]  // strip " | description" suffix (e.g. "55PBX | Telecom e PABX na nuvem")
    .trim()
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // remove acentos
    .replace(/\b(s\.?a\.?|ltda\.?|me\.?|eireli\.?|epp\.?|inc\.?|llc\.?|corp\.?|group|grupo|tecnologia|technology|tech|solucoes|soluções|sistemas|system|systems|servicos|services|consultoria|consulting)\b/g, '')
    .replace(/[^a-z0-9\s]/g, '')  // remove pontuação
    .replace(/\s+/g, ' ')
    .trim();
}

function getAttioCompanyLinks(sheetCompanyNames) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ATTIO_API_KEY');
  if (!apiKey) return null;

  sheetCompanyNames = sheetCompanyNames || [];

  // Três mapas: exact, normalized e compact (sem espaços — para "ApexBrasil" → "apex brasil")
  var exactMap      = {}; // name.toLowerCase() → record_id
  var normalizedMap = {}; // normalizeCompanyName(name) → record_id
  var compactMap    = {}; // normalizeCompanyName(name).replace(/\s/g,'') → record_id
  var offset = 0;
  var hasMore = true;

  while (hasMore) {
    var resp = UrlFetchApp.fetch('https://api.attio.com/v2/objects/companies/records/query', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ limit: 500, offset: offset }),
      muteHttpExceptions: true
    });

    var data = JSON.parse(resp.getContentText());
    var records = data.data || [];
    if (records.length === 0) { hasMore = false; break; }

    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var nameArr = r.values && r.values.name ? r.values.name : [];
      var name = nameArr[0] && nameArr[0].value ? String(nameArr[0].value).trim() : null;
      if (!name) continue;
      var id = r.id.record_id;
      exactMap[name.toLowerCase()] = id;
      var norm = normalizeCompanyName(name);
      if (norm && !normalizedMap[norm]) normalizedMap[norm] = id;
      var compact = norm.replace(/\s/g, '');
      if (compact && !compactMap[compact]) compactMap[compact] = id;
    }

    offset += records.length;
    if (records.length < 500) hasMore = false;
  }

  // Para empresas sem match, buscar no Attio por nome ($contains) — resultado em cache 6h
  var unmatched = sheetCompanyNames.filter(function(name) {
    return !exactMap[name.toLowerCase()] && !normalizedMap[normalizeCompanyName(name)];
  });

  var cache = CacheService.getScriptCache();
  var cachedJson = cache.get('attio_search_fallback');
  var searchFallback = cachedJson ? JSON.parse(cachedJson) : {};

  var needsSearch = unmatched.filter(function(name) {
    return !(name.toLowerCase() in searchFallback);
  });

  for (var j = 0; j < needsSearch.length; j++) {
    var sheetName = needsSearch[j];
    // Tentar progressivamente: nome completo → sem parênteses/sufixo → primeiras 2 palavras → primeira palavra
    var base = sheetName.split(/\s*\|\s*/)[0].split(/\s*\(/)[0].split(/\s*-\s*/)[0].trim();
    var words = base.split(/\s+/);
    var terms = [base];
    if (words.length > 2) terms.push(words.slice(0, 2).join(' '));
    if (words.length > 1) terms.push(words[0]);
    // Remover duplicatas e termos muito curtos
    var seen = {};
    terms = terms.filter(function(t) {
      if (t.length < 2 || seen[t]) return false;
      seen[t] = true; return true;
    });
    var foundId = null;
    for (var ti = 0; ti < terms.length; ti++) {
      try {
        var sResp = UrlFetchApp.fetch('https://api.attio.com/v2/objects/companies/records/query', {
          method: 'post',
          headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
          payload: JSON.stringify({ filter: { name: { '$contains': terms[ti] } }, limit: 1 }),
          muteHttpExceptions: true
        });
        var sData = JSON.parse(sResp.getContentText());
        var sRecords = sData.data || [];
        if (sRecords.length > 0) { foundId = sRecords[0].id.record_id; break; }
      } catch(e) {}
    }
    searchFallback[sheetName.toLowerCase()] = foundId;
  }

  if (needsSearch.length > 0) {
    try { cache.put('attio_search_fallback', JSON.stringify(searchFallback), 21600); } catch(e) {}
  }

  // Mesclar resultados da busca no exactMap
  for (var k in searchFallback) {
    if (searchFallback[k] && !exactMap[k]) exactMap[k] = searchFallback[k];
  }

  return { exact: exactMap, normalized: normalizedMap, compact: compactMap };
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
  var list = [];

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

      // Deal name, stage, date_mql for list
      var dealName = (vals.name && vals.name[0]) ? String(vals.name[0].value) : '—';
      var stageArr = vals.stage || [];
      var stage = (stageArr[0] && stageArr[0].status) ? stageArr[0].status.title : '—';
      var dateMql = String(dateMqlArr[0].value).slice(0, 10);
      list.push({ name: dealName, date_mql: dateMql, stage: stage, bdr: bdrName, monthId: monthId || '' });
    }

    offset += records.length;
    if (records.length < 500) hasMore = false;
  }

  list.sort(function(a, b) { return a.date_mql < b.date_mql ? 1 : -1; }); // mais recente primeiro
  result.list = list;
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

// ============================================================
// PROSPECÇÃO OUTBOUND — leitura e escrita na planilha de leads
// ============================================================

function getLeadsSheet() {
  var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === LEADS_SHEET_GID) return sheets[i];
  }
  return sheets[0];
}

function formatDateVal(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'America/Sao_Paulo', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  // se vier como número serial (Google Sheets date serial)
  if (/^\d{5}$/.test(s)) {
    var d = new Date((parseInt(s) - 25569) * 86400000);
    return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
  }
  return s;
}

function getLeadsData() {
  var sheet = getLeadsSheet();
  var data  = sheet.getDataRange().getValues();
  var leads = [];
  // rows 0 e 1 são cabeçalhos; dados a partir do índice 2
  for (var i = 2; i < data.length; i++) {
    var r = data[i];
    if (!r[0] && !r[2]) continue; // pula linhas completamente vazias
    leads.push({
      row:       i + 1, // número da linha na planilha (1-based)
      company:   String(r[0]  || ''),
      employees: String(r[1]  || ''),
      name:      String(r[2]  || ''),
      title:     String(r[3]  || ''),
      location:  String(r[4]  || ''),
      domain:    String(r[5]  || ''),
      linkedin:  String(r[6]  || ''),
      phone1:    String(r[7]  || ''),
      phone2:    String(r[8]  || ''),
      email:     String(r[9]  || ''),
      status:    String(r[10] || ''),
      fups: [
        { status: String(r[11] || ''), date: formatDateVal(r[12]), answered: formatDateVal(r[13]) },
        { status: String(r[14] || ''), date: formatDateVal(r[15]), answered: formatDateVal(r[16]) },
        { status: String(r[17] || ''), date: formatDateVal(r[18]), answered: formatDateVal(r[19]) },
        { status: String(r[20] || ''), date: formatDateVal(r[21]), answered: formatDateVal(r[22]) },
        { status: String(r[23] || ''), date: formatDateVal(r[24]), answered: formatDateVal(r[25]) }
      ]
    });
  }
  return leads;
}

function updateLeadCell(row, col, value) {
  var sheet = getLeadsSheet();
  var cell  = sheet.getRange(row, col);
  if (value === '' || value === null || value === undefined) {
    cell.clearContent();
    return;
  }
  // Colunas de data (1-based): Data Envio=13,16,19,22,25 | Data Resp=14,17,20,23,26
  var dateCols = [13, 14, 16, 17, 19, 20, 22, 23, 25, 26];
  if (dateCols.indexOf(col) > -1) {
    var parts = String(value).split('-');
    if (parts.length === 3) {
      var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      if (!isNaN(d.getTime())) { cell.setValue(d); return; }
    }
  }
  cell.setValue(value);
}

function getLeadsValidationOptions() {
  var sheet = getLeadsSheet();
  function readValidation(row, col) {
    try {
      var v = sheet.getRange(row, col).getDataValidation();
      if (!v) return [];
      if (v.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        return v.getCriteriaValues()[0] || [];
      }
    } catch(e) {}
    return [];
  }
  return {
    status:    readValidation(3, 11),  // Status geral
    fupStatus: readValidation(3, 12),  // Status Fase1 (FUP)
    answered:  readValidation(3, 14)   // Respondeu?
  };
}
