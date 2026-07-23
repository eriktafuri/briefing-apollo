/**
 * Backend do Briefing Apollo — recebe o POST do formulário HTML,
 * grava uma linha na aba "Respostas", salva os arquivos no Drive,
 * e serve o painel interno (lista de respostas + visualização de cada uma).
 *
 * Instruções completas de instalação: ver README.md.
 */

var SHEET_NAME = 'Respostas';
var DRIVE_FOLDER_NAME = 'Briefings Apollo — Arquivos';
var FILE_FIELDS = ['identidadeVisual', 'conteudoArquivos'];

// troque por uma senha sua antes de divulgar o painel
var PAINEL_SENHA = 'apollo123';

var HEADERS = [
  'timestamp',
  'respNome', 'respEmail', 'respWhatsapp',
  'empresaNome', 'empresaInicio', 'empresaAtividade', 'empresaProdutos',
  'objetivoPrincipal', 'objetivoSucesso',
  'publicoIdeal', 'publicoObjecoes',
  'concorrentes', 'diferenciais', 'concorrentesFrente',
  'personalidade', 'personalidadeTop3', 'linguagemEvitar', 'naoQuerVer',
  'identidadeVisual', 'referenciasVisuais', 'conteudoArquivos',
  'pastaDrive',
  'infoAdicional',
  'id' // sempre por último — identifica a resposta pro painel interno
];

/* ================= ROTEAMENTO ================= */
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'listar') return doGetListar_(e);
    if (action === 'resposta') return doGetResposta_(e);
    if (action === 'excluir') return doGetExcluir_(e);

    // sem action: diagnóstico de formatação (compatibilidade)
    var sheet = getSheet_();
    ensureHeader_(sheet);
    formatarPlanilha_(sheet);
    var bandings = sheet.getBandings();
    return jsonOut_({
      ok: true,
      linhas: sheet.getLastRow(),
      colunas: sheet.getLastColumn(),
      faixasAplicadas: bandings.length,
      linhaCongelada: sheet.getFrozenRows(),
      colunaCongelada: sheet.getFrozenColumns(),
      larguraColuna1: sheet.getColumnWidth(1)
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    ensureHeader_(sheet);

    var empresa = (data.empresaNome || data.respNome || 'sem-nome').toString();
    var timestamp = new Date();
    var id = Utilities.getUuid();

    // uma pasta só por envio — todos os arquivos (de qualquer campo) caem nela junto
    var hasFiles = FILE_FIELDS.some(function (key) {
      return data.files && data.files[key] && data.files[key].length;
    });
    var clientFolder = hasFiles ? createClientFolder_(empresa, timestamp) : null;

    var fileLinks = {};
    FILE_FIELDS.forEach(function (key) {
      fileLinks[key] = saveFiles_(clientFolder, data.files && data.files[key]);
    });

    var row = HEADERS.map(function (h) {
      if (h === 'timestamp') return timestamp;
      if (h === 'id') return id;
      if (h === 'pastaDrive') return clientFolder ? clientFolder.getUrl() : '';
      if (FILE_FIELDS.indexOf(h) > -1) return fileLinks[h];
      if (h === 'personalidade') return (data.personalidade || []).join(', ');
      return data[h] !== undefined ? data[h] : '';
    });
    sheet.appendRow(row);

    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

/* ================= PAINEL INTERNO (somente leitura) ================= */
function checarSenha_(e) {
  if (!e.parameter.senha || e.parameter.senha !== PAINEL_SENHA) {
    throw new Error('Senha incorreta.');
  }
}

// lista resumida — usada pelo painel.html
function doGetListar_(e) {
  checarSenha_(e);
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return jsonOut_({ ok: true, respostas: [] });

  var idCol = HEADERS.indexOf('id');
  var empresaCol = HEADERS.indexOf('empresaNome');
  var nomeCol = HEADERS.indexOf('respNome');
  var tsCol = HEADERS.indexOf('timestamp');

  var values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var lista = values.map(function (r) {
    return {
      id: r[idCol],
      empresaNome: r[empresaCol],
      respNome: r[nomeCol],
      timestamp: r[tsCol] instanceof Date ? r[tsCol].toISOString() : r[tsCol]
    };
  }).reverse(); // mais recentes primeiro

  return jsonOut_({ ok: true, respostas: lista });
}

// resposta completa (todos os campos) — usada pela resposta.html
function doGetResposta_(e) {
  checarSenha_(e);
  var id = (e.parameter.id || '').toString();
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return jsonOut_({ ok: false });

  var idCol = HEADERS.indexOf('id');
  var values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][idCol] === id) {
      var obj = {};
      HEADERS.forEach(function (h, j) {
        var v = values[i][j];
        obj[h] = v instanceof Date ? v.toISOString() : v;
      });
      return jsonOut_({ ok: true, resposta: obj });
    }
  }
  return jsonOut_({ ok: false });
}

// exclui uma resposta (linha inteira) pelo id — usada pelo painel.html
function doGetExcluir_(e) {
  checarSenha_(e);
  var id = (e.parameter.id || '').toString();
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return jsonOut_({ ok: false });

  var idCol = HEADERS.indexOf('id');
  var values = sheet.getRange(2, 1, last - 1, 1 + idCol).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][idCol] === id) {
      sheet.deleteRow(i + 2);
      return jsonOut_({ ok: true });
    }
  }
  return jsonOut_({ ok: false });
}

/* ---------- Sheet (Respostas) ---------- */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

// idempotente: cria o cabeçalho do zero, estende colunas novas no final
// (sem nunca deslocar colunas já existentes) e resincroniza os rótulos —
// seguro rodar sempre, mesmo que o schema tenha crescido desde a instalação.
function ensureHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    formatarPlanilha_(sheet);
    return;
  }
  var atuaisLen = sheet.getLastColumn();
  var cresceu = atuaisLen < HEADERS.length;
  if (cresceu) sheet.insertColumnsAfter(atuaisLen, HEADERS.length - atuaisLen);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (cresceu) formatarPlanilha_(sheet);
}

var COLUMN_WIDTHS = {
  timestamp: 140, respNome: 140, respEmail: 170, respWhatsapp: 130,
  empresaNome: 160, empresaInicio: 110, empresaAtividade: 260, empresaProdutos: 240,
  objetivoPrincipal: 240, objetivoSucesso: 220,
  publicoIdeal: 240, publicoObjecoes: 240,
  concorrentes: 220, diferenciais: 220, concorrentesFrente: 220,
  personalidade: 260, personalidadeTop3: 200, linguagemEvitar: 200, naoQuerVer: 200,
  identidadeVisual: 220, referenciasVisuais: 220, conteudoArquivos: 220,
  pastaDrive: 200, infoAdicional: 260, id: 220
};

// Visual "Apollo": faixas em laranja, cabeçalho fixo, colunas com largura pensada
// pro tipo de conteúdo e quebra de texto automática.
function formatarPlanilha_(sheet) {
  sheet.getBandings().forEach(function (b) { b.remove(); });
  var full = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 200), HEADERS.length);
  full.applyRowBanding(SpreadsheetApp.BandingTheme.ORANGE, true, false);

  HEADERS.forEach(function (h, i) {
    sheet.setColumnWidth(i + 1, COLUMN_WIDTHS[h] || 160);
  });

  var dataRange = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 200), HEADERS.length);
  dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  dataRange.setVerticalAlignment('top');

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function aplicarFormatacaoAgora() {
  formatarPlanilha_(getSheet_());
}

/* ---------- Drive ---------- */
function getRootFolder_() {
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function createClientFolder_(empresa, timestamp) {
  var root = getRootFolder_();
  var subName = empresa + ' — ' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd-MM-yyyy HH-mm');
  return root.createFolder(subName);
}

function saveFiles_(folder, files) {
  if (!folder || !files || !files.length) return '';
  var links = files.map(function (f) {
    var base64 = (f.dataUrl || '').split(',')[1] || '';
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, f.type || 'application/octet-stream', f.name || 'arquivo');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  });
  return links.join('\n');
}

/* ---------- Util ---------- */
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
