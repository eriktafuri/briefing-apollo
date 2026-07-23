/**
 * Backend do Briefing Apollo — recebe o POST do formulário HTML,
 * grava uma linha na aba "Respostas", salva os arquivos no Drive,
 * e gerencia os links personalizados por cliente (aba "Convites").
 *
 * Instruções completas de instalação: ver README.md.
 */

var SHEET_NAME = 'Respostas';
var CONVITES_SHEET_NAME = 'Convites';
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
  'slug' // sempre por último — não mexer na ordem das colunas já existentes
];

var CONVITES_HEADERS = ['slug', 'empresa', 'criadoEm', 'status', 'respondidoEm', 'link'];

/* ================= ROTEAMENTO ================= */
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'info') return doGetInfo_(e);
    if (action === 'criar') return doGetCriar_(e);
    if (action === 'listar') return doGetListar_(e);

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
      if (h === 'pastaDrive') return clientFolder ? clientFolder.getUrl() : '';
      if (h === 'slug') return data.slug || '';
      if (FILE_FIELDS.indexOf(h) > -1) return fileLinks[h];
      if (h === 'personalidade') return (data.personalidade || []).join(', ');
      return data[h] !== undefined ? data[h] : '';
    });
    sheet.appendRow(row);

    if (data.slug) marcarRespondido_(data.slug);

    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

/* ================= CONVITES (links personalizados) ================= */
function checarSenha_(e) {
  if (!e.parameter.senha || e.parameter.senha !== PAINEL_SENHA) {
    throw new Error('Senha incorreta.');
  }
}

function doGetInfo_(e) {
  var slug = (e.parameter.slug || '').toString();
  var convite = buscarConvitePorSlug_(slug);
  if (!convite) return jsonOut_({ ok: false });
  return jsonOut_({ ok: true, empresa: convite.empresa, status: convite.status });
}

function doGetCriar_(e) {
  checarSenha_(e);
  var empresa = (e.parameter.empresa || '').toString().trim();
  if (!empresa) throw new Error('Informe o nome da empresa.');
  var link = (e.parameter.link || '').toString();
  var criado = criarConvite_(empresa, link);
  return jsonOut_({ ok: true, slug: criado.slug });
}

function doGetListar_(e) {
  checarSenha_(e);
  return jsonOut_({ ok: true, convites: listarConvites_() });
}

function getConvitesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONVITES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONVITES_SHEET_NAME);
    sheet.appendRow(CONVITES_HEADERS);
    sheet.getRange(1, 1, 1, CONVITES_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    var widths = { slug: 180, empresa: 200, criadoEm: 150, status: 110, respondidoEm: 150, link: 320 };
    CONVITES_HEADERS.forEach(function (h, i) { sheet.setColumnWidth(i + 1, widths[h] || 160); });
    sheet.getBandings().forEach(function (b) { b.remove(); });
    sheet.getRange(1, 1, 200, CONVITES_HEADERS.length).applyRowBanding(SpreadsheetApp.BandingTheme.ORANGE, true, false);
  }
  return sheet;
}

function slugify_(text) {
  var s = (text || '').toString().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 48);
  return s || 'briefing';
}

function criarConvite_(empresa, link) {
  var sheet = getConvitesSheet_();
  var base = slugify_(empresa);
  var last = sheet.getLastRow();
  var existentes = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  var slug = base;
  var i = 2;
  while (existentes.indexOf(slug) > -1) { slug = base + '-' + i; i++; }
  sheet.appendRow([slug, empresa, new Date(), 'Enviado', '', link || '']);
  return { slug: slug };
}

function listarConvites_() {
  var sheet = getConvitesSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, CONVITES_HEADERS.length).getValues();
  return values.map(function (r) {
    var o = {};
    CONVITES_HEADERS.forEach(function (h, i) { o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
    return o;
  }).reverse();
}

function buscarConvitePorSlug_(slug) {
  if (!slug) return null;
  var sheet = getConvitesSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var values = sheet.getRange(2, 1, last - 1, CONVITES_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === slug) return { row: i + 2, empresa: values[i][1], status: values[i][3] };
  }
  return null;
}

function marcarRespondido_(slug) {
  var found = buscarConvitePorSlug_(slug);
  if (!found) return;
  var sheet = getConvitesSheet_();
  sheet.getRange(found.row, 4).setValue('Respondido');
  sheet.getRange(found.row, 5).setValue(new Date());
}

/* ---------- Sheet (Respostas) ---------- */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

// idempotente: cria o cabeçalho do zero, ou completa colunas novas no final
// (sem nunca deslocar colunas já existentes) se o schema crescer com o tempo.
function ensureHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    formatarPlanilha_(sheet);
    return;
  }
  var atuais = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (atuais.length < HEADERS.length) {
    var faltando = HEADERS.slice(atuais.length);
    sheet.getRange(1, atuais.length + 1, 1, faltando.length).setValues([faltando]);
    formatarPlanilha_(sheet);
  }
}

var COLUMN_WIDTHS = {
  timestamp: 140, respNome: 140, respEmail: 170, respWhatsapp: 130,
  empresaNome: 160, empresaInicio: 110, empresaAtividade: 260, empresaProdutos: 240,
  objetivoPrincipal: 240, objetivoSucesso: 220,
  publicoIdeal: 240, publicoObjecoes: 240,
  concorrentes: 220, diferenciais: 220, concorrentesFrente: 220,
  personalidade: 260, personalidadeTop3: 200, linguagemEvitar: 200, naoQuerVer: 200,
  identidadeVisual: 220, referenciasVisuais: 220, conteudoArquivos: 220,
  pastaDrive: 200, infoAdicional: 260, slug: 160
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
