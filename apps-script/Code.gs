// ============================================================
// CONFIGURAÇÃO — atualizar no início de cada mês
// ============================================================

var SPRINT_CONFIG = [
  { name: 'S1', start: new Date('2026-03-01'), end: new Date('2026-03-08') },
  { name: 'S2', start: new Date('2026-03-09'), end: new Date('2026-03-16') },
  { name: 'S3', start: new Date('2026-03-17'), end: new Date('2026-03-24') },
  { name: 'S4', start: new Date('2026-03-25'), end: new Date('2026-04-01') }
];

// Nome exato da aba de métricas (será ignorada no processamento)
var METRICS_SHEET_NAME = 'Métricas';

// ============================================================
// ENDPOINT PRINCIPAL
// ============================================================

function doGet(e) {
  var result = getData();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// COLETA E PROCESSAMENTO
// ============================================================

function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var campaigns = [];

  for (var i = 0; i < sheets.length; i++) {
    var result = processCampaignSheet(sheets[i]);
    if (result) campaigns.push(result);
  }

  return {
    lastUpdated: new Date().toISOString(),
    sprints: {
      S1: { start: '2026-03-01', end: '2026-03-08', label: '01/03 – 08/03' },
      S2: { start: '2026-03-09', end: '2026-03-16', label: '09/03 – 16/03' },
      S3: { start: '2026-03-17', end: '2026-03-24', label: '17/03 – 24/03' },
      S4: { start: '2026-03-25', end: '2026-04-01', label: '25/03 – 01/04' }
    },
    totals: calculateTotals(campaigns),
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

function isCampaignSheet(sheetName, headers) {
  if (sheetName === METRICS_SHEET_NAME) return false;
  var flat = headers.map(function(h) { return String(h).toLowerCase().trim(); });
  return flat.indexOf('nome') > -1 &&
         flat.indexOf('bdr') > -1 &&
         flat.indexOf('empresa') > -1;
}

function emptySprintObj() {
  return { all: 0, S1: 0, S2: 0, S3: 0, S4: 0 };
}

function emptyBDRMetric() {
  return {
    total: emptySprintObj(),
    cath:  emptySprintObj(),
    mat:   emptySprintObj()
  };
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

  // Row index 1 (0-based) has the real headers
  var headers = allValues[1].map(function(h) { return String(h).trim(); });
  if (!isCampaignSheet(name, headers)) return null;

  var nomeIdx        = headers.indexOf('Nome');
  var empresaIdx     = headers.indexOf('Empresa');
  var bdrIdx         = headers.indexOf('BDR');
  var statusGeralIdx = headers.indexOf('Status Geral');

  if (nomeIdx < 0 || empresaIdx < 0 || bdrIdx < 0) return null;

  // All "Data Envio" columns (one per phase)
  var dataEnvioIdxs = [];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'Data Envio') dataEnvioIdxs.push(i);
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
    var bdr         = String(row[bdrIdx] || '').trim().toUpperCase();
    var statusGeral = String(statusGeralIdx >= 0 ? row[statusGeralIdx] : '').trim();

    if (!empresa) continue;

    var bdrKey = (bdr === 'CATH') ? 'cath' : 'mat';

    // Unique companies per BDR
    if (bdrKey === 'cath') empresasCath[empresa] = true;
    else                   empresasMat[empresa]  = true;

    // Contatos
    contatos.total++;
    contatos[bdrKey]++;

    // Mensagens — one per phase that has a send date
    for (var p = 0; p < dataEnvioIdxs.length; p++) {
      var dateVal = row[dataEnvioIdxs[p]];
      if (!dateVal) continue;
      var date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
      if (isNaN(date.getTime())) continue;
      addCount(mensagens, bdrKey, getSprint(date));
    }

    // MQL — classified by first send date
    if (statusGeral === 'MQL') {
      var mqlDate = null;
      for (var q = 0; q < dataEnvioIdxs.length; q++) {
        var dv = row[dataEnvioIdxs[q]];
        if (dv) {
          mqlDate = (dv instanceof Date) ? dv : new Date(dv);
          break;
        }
      }
      addCount(mqls, bdrKey, mqlDate ? getSprint(mqlDate) : null);
    }
  }

  // Merge company sets
  var allEmpresas = {};
  Object.keys(empresasCath).forEach(function(k) { allEmpresas[k] = true; });
  Object.keys(empresasMat).forEach(function(k)  { allEmpresas[k] = true; });

  var empresas = {
    total: Object.keys(allEmpresas).length,
    cath:  Object.keys(empresasCath).length,
    mat:   Object.keys(empresasMat).length
  };

  var mqlCount         = mqls.total.all;
  var msgCount         = mensagens.total.all;
  var contatoPerEmpresa = empresas.total > 0 ? contatos.total / empresas.total : 0;

  return {
    name: name,
    empresas: empresas,
    contatos: contatos,
    mensagens: mensagens,
    mqls: mqls,
    conversion: {
      pct_empresa:          empresas.total > 0  ? mqlCount / empresas.total  : 0,
      msgs_per_mql:         mqlCount > 0        ? round2(msgCount / mqlCount) : 0,
      msgs_per_mql_contato: (mqlCount > 0 && contatoPerEmpresa > 0)
                              ? round2(msgCount / mqlCount / contatoPerEmpresa)
                              : 0,
      contato_por_empresa:  round2(contatoPerEmpresa),
      pct_conv_stakeholder: contatos.total > 0  ? mqlCount / contatos.total  : 0,
      pct_conv_empresa:     empresas.total > 0  ? mqlCount / empresas.total  : 0
    }
  };
}

function calculateTotals(campaigns) {
  var empresasTotal = 0, empresasCath = 0, empresasMat = 0;
  var contatos  = { total: 0, cath: 0, mat: 0 };
  var mensagens = emptyBDRMetric();
  var mqls      = emptyBDRMetric();

  campaigns.forEach(function(c) {
    empresasTotal += c.empresas.total;
    empresasCath  += c.empresas.cath;
    empresasMat   += c.empresas.mat;
    contatos.total += c.contatos.total;
    contatos.cath  += c.contatos.cath;
    contatos.mat   += c.contatos.mat;

    ['total', 'cath', 'mat'].forEach(function(bdr) {
      ['all', 'S1', 'S2', 'S3', 'S4'].forEach(function(s) {
        mensagens[bdr][s] += c.mensagens[bdr][s];
        mqls[bdr][s]      += c.mqls[bdr][s];
      });
    });
  });

  var mqlCount          = mqls.total.all;
  var msgCount          = mensagens.total.all;
  var contatoPerEmpresa = empresasTotal > 0 ? contatos.total / empresasTotal : 0;

  return {
    empresas: { total: empresasTotal, cath: empresasCath, mat: empresasMat },
    contatos: contatos,
    mensagens: mensagens,
    mqls: mqls,
    conversion: {
      pct_empresa:          empresasTotal > 0   ? mqlCount / empresasTotal   : 0,
      msgs_per_mql:         mqlCount > 0        ? round2(msgCount / mqlCount) : 0,
      msgs_per_mql_contato: (mqlCount > 0 && contatoPerEmpresa > 0)
                              ? round2(msgCount / mqlCount / contatoPerEmpresa)
                              : 0,
      contato_por_empresa:  round2(contatoPerEmpresa),
      pct_conv_stakeholder: contatos.total > 0  ? mqlCount / contatos.total  : 0,
      pct_conv_empresa:     empresasTotal > 0   ? mqlCount / empresasTotal   : 0
    }
  };
}

function round2(val) {
  return Math.round(val * 100) / 100;
}
