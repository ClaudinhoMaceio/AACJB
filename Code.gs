/**
 * Backend do Censo (Google Apps Script + Google Drive JSON).
 *
 * Objetivo:
 * - Receber respostas do sistema (POST action=append)
 * - Entregar respostas ao sistema (GET action=getAll)
 * - Criar automaticamente o primeiro JSON caso nao exista
 * - Funcionar com JSONP e com iframe/postMessage para GitHub Pages
 */

var ROOT_FOLDER_ID = '1kz9CGNas03tiQzUUu34vraQFhimWkPgJ';
var SCRIPT_PROP_ROOT_OVERRIDE = 'CENSO_ROOT_FOLDER_ID';
var PROP_BASE = 'CENSO_JSON_BASE_NAME';
var DEFAULT_BASE_NAME = 'censo_culinaria_japonesa_respostas';
var MAX_JSON_BYTES = 1024 * 1024 * 1024;
/** Arquivo na mesma pasta do censo: lista de administradores (nome + senha em texto — restrinja acesso à pasta). */
var ADMINS_FILE_NAME = 'censo_culinaria_japonesa_admins.json';
/**
 * Manter igual ao ADMIN_PASSWORD em index.html (para reset remoto funcionar).
 * Opcionalmente defina só no Apps Script com a propriedade CENSO_RESET_SENHA.
 */
var CENSO_RESET_SENHA_PADRAO = 'Marcello';

function validateResetCredential_(provided) {
  var s = String(provided || '').trim();
  if (!s) {
    return false;
  }
  var configured = String(getPropsStore_().getProperty('CENSO_RESET_SENHA') || CENSO_RESET_SENHA_PADRAO || '').trim();
  return configured === s;
}

function getPropsStore_() {
  var props = null;
  try {
    props = PropertiesService.getDocumentProperties();
  } catch (_) {}
  if (!props) {
    props = PropertiesService.getScriptProperties();
  }
  return props;
}

function getBaseName_() {
  return getPropsStore_().getProperty(PROP_BASE) || DEFAULT_BASE_NAME;
}

function getConfiguredRootFolderId_() {
  var propKey =
    typeof SCRIPT_PROP_ROOT_OVERRIDE !== 'undefined' && SCRIPT_PROP_ROOT_OVERRIDE
      ? String(SCRIPT_PROP_ROOT_OVERRIDE)
      : 'CENSO_ROOT_FOLDER_ID';
  var overrideId = getPropsStore_().getProperty(propKey);
  if (overrideId && String(overrideId).trim()) {
    return String(overrideId).trim();
  }
  return String(ROOT_FOLDER_ID || '').trim();
}

function getTargetFolder_() {
  var folderId = getConfiguredRootFolderId_();
  if (!folderId) {
    throw new Error('Pasta nao configurada. Execute configurarNaPasta("ID_DA_PASTA").');
  }
  return DriveApp.getFolderById(folderId);
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }
  var raw = String(e.postData.contents || '');
  if (!raw) {
    return {};
  }
  return JSON.parse(raw);
}

function parseQueryString_(qs) {
  var out = {};
  if (!qs) {
    return out;
  }
  var parts = String(qs).split('&');
  for (var i = 0; i < parts.length; i++) {
    var pair = parts[i].split('=');
    var key = decodeURIComponent((pair[0] || '').replace(/\+/g, ' '));
    var val = decodeURIComponent((pair.slice(1).join('=') || '').replace(/\+/g, ' '));
    if (key) {
      out[key] = val;
    }
  }
  return out;
}

function getRequestParams_(e) {
  var out = {};
  if (!e) {
    return out;
  }
  var k;
  if (e.parameter) {
    for (k in e.parameter) {
      out[k] = e.parameter[k];
    }
  }
  if (e.parameters) {
    for (k in e.parameters) {
      if (e.parameters[k] && e.parameters[k].length) {
        out[k] = e.parameters[k][0];
      }
    }
  }
  if (e.queryString) {
    var parsed = parseQueryString_(e.queryString);
    for (k in parsed) {
      if (out[k] === undefined) {
        out[k] = parsed[k];
      }
    }
  }
  return out;
}

function jsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function sanitizeJsonpCallback_(raw) {
  var s = String(raw || '_cb').replace(/[^\w$]/g, '');
  if (!s.length) {
    s = '_cb';
  }
  return s.substring(0, 64);
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function padPart_(n) {
  var s = String(n);
  return s.length >= 3 ? s : ('000' + s).slice(-3);
}

function readJsonArrayFromFile_(file) {
  var raw = file.getBlob().getDataAsString('UTF-8');
  if (!raw || !raw.trim()) {
    return [];
  }
  try {
    var data = JSON.parse(raw);
    return Array.isArray(data) ? data : [];
  } catch (_) {
    return [];
  }
}

function listShardFiles_(folder) {
  var base = getBaseName_();
  var re = new RegExp('^' + escapeRegex_(base) + '(?:_part(\\d+))?\\.json$', 'i');
  var it = folder.getFiles();
  var list = [];
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    var m = name.match(re);
    if (!m) {
      continue;
    }
    list.push({ file: f, part: m[1] ? parseInt(m[1], 10) : 1 });
  }
  list.sort(function (a, b) {
    return a.part - b.part;
  });
  return list.map(function (x) {
    return x.file;
  });
}

function ensureFirstShard_(folder) {
  var files = listShardFiles_(folder);
  if (files.length > 0) {
    return files[0];
  }
  return folder.createFile(getBaseName_() + '.json', '[]', MimeType.PLAIN_TEXT);
}

function getNextPartNumber_(folder) {
  var base = getBaseName_();
  var re = new RegExp('^' + escapeRegex_(base) + '(?:_part(\\d+))?\\.json$', 'i');
  var it = folder.getFiles();
  var max = 0;
  while (it.hasNext()) {
    var name = it.next().getName();
    var m = name.match(re);
    if (!m) {
      continue;
    }
    var part = m[1] ? parseInt(m[1], 10) : 1;
    if (part > max) {
      max = part;
    }
  }
  return max + 1;
}

function getAdminsFile_(folder) {
  var it = folder.getFilesByName(ADMINS_FILE_NAME);
  if (it.hasNext()) {
    return it.next();
  }
  return null;
}

function ensureAdminsFile_(folder) {
  var f = getAdminsFile_(folder);
  if (f) {
    return f;
  }
  return folder.createFile(ADMINS_FILE_NAME, '[]', MimeType.PLAIN_TEXT);
}

/**
 * @returns {Array<{id:string,nome:string,senha:string}>}
 */
function getAdmins() {
  var folder = getTargetFolder_();
  var f = getAdminsFile_(folder);
  if (!f) {
    return [];
  }
  var raw = f.getBlob().getDataAsString('UTF-8');
  if (!raw || !raw.trim()) {
    return [];
  }
  try {
    var data = JSON.parse(raw);
    return Array.isArray(data) ? data : [];
  } catch (_) {
    return [];
  }
}

/**
 * Substitui o arquivo de administradores pelo array enviado (validação mínima no cliente).
 */
function saveAdmins(admins) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var folder = getTargetFolder_();
    var f = ensureAdminsFile_(folder);
    var arr = Array.isArray(admins) ? admins : [];
    f.setContent(JSON.stringify(arr));
    return { ok: true, fileName: ADMINS_FILE_NAME };
  } finally {
    lock.releaseLock();
  }
}

/** Limpa respostas; exige a mesma senha do ADMIN_PASSWORD no front (ou CENSO_RESET_SENHA nas propriedades). */
function resetCensusResponses(resetSecret) {
  if (!validateResetCredential_(resetSecret)) {
    throw new Error('Credencial de reset invalida ou ausente.');
  }
  return limparBancoCensoAgora();
}

function getAllResponses() {
  var folder = getTargetFolder_();
  ensureFirstShard_(folder);
  var files = listShardFiles_(folder);
  var all = [];
  for (var i = 0; i < files.length; i++) {
    var rows = readJsonArrayFromFile_(files[i]);
    for (var j = 0; j < rows.length; j++) {
      all.push(rows[j]);
    }
  }
  return all;
}

function appendResponse(newResponse) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var folder = getTargetFolder_();
    var files = listShardFiles_(folder);
    var activeFile = files.length ? files[files.length - 1] : ensureFirstShard_(folder);
    var arr = readJsonArrayFromFile_(activeFile);
    arr.push(newResponse);
    var text = JSON.stringify(arr);
    if (Utilities.newBlob(text, 'UTF-8').getBytes().length <= MAX_JSON_BYTES) {
      activeFile.setContent(text);
      return;
    }
    var last = arr.pop();
    activeFile.setContent(JSON.stringify(arr));
    var nextPart = getNextPartNumber_(folder);
    var nextName = getBaseName_() + '_part' + padPart_(nextPart) + '.json';
    folder.createFile(nextName, JSON.stringify([last]), MimeType.PLAIN_TEXT);
  } finally {
    lock.releaseLock();
  }
}

function normalizePrizeKey_(prize) {
  var p = String(prize || '').toLowerCase();
  if (p.indexOf('fitness') !== -1) {
    return 'fitness';
  }
  if (p.indexOf('sushi') !== -1 || p.indexOf('sashimi') !== -1) {
    return 'sushi';
  }
  return '';
}

function attachmentBlobFromPayload_(obj) {
  if (!obj || !obj.base64) {
    return null;
  }
  var bytes = Utilities.base64Decode(String(obj.base64));
  var name = String(obj.name || 'ebook.pdf');
  var mimeType = String(obj.mimeType || 'application/octet-stream');
  return Utilities.newBlob(bytes, mimeType, name);
}

function sendGiftEmails(recipients, attachments, mailOptions) {
  var list = Array.isArray(recipients) ? recipients : [];
  var atts = attachments || {};
  var opts = mailOptions || {};
  var senderCompany = String(opts.senderCompany || 'AACJ - Associação dos Adeptos da Culinária Japonesa').trim();
  var senderDisplayName = String(opts.senderDisplayName || 'Equipe AACJ').trim();
  var fitnessBlob = attachmentBlobFromPayload_(atts.fitness);
  var sushiBlob = attachmentBlobFromPayload_(atts.sushi);
  var sent = 0;
  var skipped = 0;
  var details = [];

  for (var i = 0; i < list.length; i++) {
    var r = list[i] || {};
    var email = String(r.email || '').trim();
    if (!email || email.indexOf('@') === -1) {
      skipped++;
      details.push({ ok: false, email: email, reason: 'email-invalido' });
      continue;
    }
    var prizeKey = normalizePrizeKey_(r.premioDesejado);
    var att = prizeKey === 'fitness' ? fitnessBlob : prizeKey === 'sushi' ? sushiBlob : null;
    if (!att) {
      skipped++;
      details.push({ ok: false, email: email, reason: 'anexo-ausente' });
      continue;
    }
    var nome = String(r.nome || 'Participante').trim() || 'Participante';
    var premioNome = String(r.premioDesejado || '').trim() || 'E-book especial AACJ';
    var subject = 'Seu presente gratuito da AACJ chegou';
    var bodyText =
      'Olá, ' +
      nome +
      '!\n\n' +
      'Arigatou por participar do Censo Nacional da AACJ.\n' +
      'Estamos muito felizes com seu cadastro.\n\n' +
      'Seu presente selecionado:\n' +
      premioNome +
      '\n\n' +
      'O arquivo segue em anexo neste e-mail.\n\n' +
      'Com carinho,\n' +
      senderCompany;
    var bodyHtml =
      '<div style="font-family:Arial,sans-serif;background:#f8fafc;padding:24px;color:#0f172a">' +
      '<div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #e2e8f0;border-radius:16px;overflow:hidden">' +
      '<div style="background:linear-gradient(90deg,#0f172a,#1e3a8a);padding:16px 20px;color:#fff">' +
      '<div style="font-size:18px;font-weight:bold">AACJ • Presente Especial</div>' +
      '<div style="font-size:12px;opacity:.9">Censo Nacional dos Adeptos da Culinária Japonesa</div>' +
      '</div>' +
      '<div style="padding:20px">' +
      '<p style="margin:0 0 10px 0">Olá, <strong>' +
      nome +
      '</strong>!</p>' +
      '<p style="margin:0 0 12px 0;line-height:1.6">Muito obrigado pelo seu cadastro no nosso censo.<br>Conforme sua escolha, enviamos em anexo o seu presente digital.</p>' +
      '<div style="background:#eff6ff;border:1px solid #bfdbfe;padding:12px;border-radius:10px;margin:12px 0">' +
      '<div style="font-size:12px;color:#475569">Presente selecionado</div>' +
      '<div style="font-size:14px;font-weight:bold;color:#1e3a8a">' +
      premioNome +
      '</div>' +
      '</div>' +
      '<p style="margin:14px 0 0 0;line-height:1.6">Desejamos uma excelente experiência com a culinária japonesa.<br>Arigatou gozaimasu!</p>' +
      '<p style="margin:16px 0 0 0;color:#334155"><strong>' +
      senderCompany +
      '</strong></p>' +
      '</div>' +
      '</div>' +
      '</div>';
    GmailApp.sendEmail(email, subject, bodyText, {
      name: senderDisplayName || 'AACJ',
      htmlBody: bodyHtml,
      attachments: [att]
    });
    sent++;
    details.push({ ok: true, email: email, prizeKey: prizeKey });
  }

  return { ok: true, sent: sent, skipped: skipped, total: list.length, details: details };
}

function responseGetAllIframe_(p) {
  var targetOrigin = p.origin ? String(p.origin) : '*';
  var payloadObj;
  try {
    payloadObj = { source: 'censo-drive', payload: getAllResponses() };
  } catch (err) {
    payloadObj = { source: 'censo-drive', error: String(err) };
  }
  var inner = JSON.stringify(payloadObj);
  var b64 = Utilities.base64Encode(Utilities.newBlob(inner, 'UTF-8').getBytes());
  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="robots" content="noindex"></head><body>' +
    '<script>(function(){try{window.parent.postMessage({source:"censo-drive",b64:' +
    JSON.stringify(b64) +
    '},' +
    JSON.stringify(targetOrigin) +
    ');}catch(err){window.parent.postMessage({source:"censo-drive",error:String(err)},"*");}})();</script>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).setTitle(' ');
}

function doGet(e) {
  try {
    var p = getRequestParams_(e);
    var action = String(p.action || '').trim();

    if (!action || action === 'ping' || action === 'health') {
      return jsonOutput_({
        ok: true,
        service: 'censo-drive-backend',
        time: new Date().toISOString(),
        folderId: getConfiguredRootFolderId_(),
        baseName: getBaseName_()
      });
    }

    if (action === 'getAll' && String(p.fmt) === 'iframe') {
      return responseGetAllIframe_(p);
    }

    if (action === 'getAll' && (p.callback || p.cb)) {
      var cb = sanitizeJsonpCallback_(p.callback || p.cb);
      try {
        return ContentService.createTextOutput(cb + '(' + JSON.stringify(getAllResponses()) + ');').setMimeType(
          ContentService.MimeType.JAVASCRIPT
        );
      } catch (errJsonp) {
        return ContentService.createTextOutput(cb + '(' + JSON.stringify({ error: String(errJsonp) }) + ');').setMimeType(
          ContentService.MimeType.JAVASCRIPT
        );
      }
    }

    if (action === 'getAll') {
      return jsonOutput_({ ok: true, data: getAllResponses() });
    }

    if (action === 'getAdmins' && (p.callback || p.cb)) {
      var cbAdm = sanitizeJsonpCallback_(p.callback || p.cb);
      try {
        var admList = getAdmins();
        return ContentService.createTextOutput(cbAdm + '(' + JSON.stringify(admList) + ');').setMimeType(
          ContentService.MimeType.JAVASCRIPT
        );
      } catch (errAdmJsonp) {
        return ContentService.createTextOutput(
          cbAdm + '(' + JSON.stringify([]) + ');'
        ).setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
    }

    if (action === 'getAdmins') {
      return jsonOutput_({ ok: true, data: getAdmins() });
    }

    if (action === 'resetResponses') {
      try {
        var secGet = String(p.resetSecret || p.token || '').trim();
        var outReset = resetCensusResponses(secGet);
        return jsonOutput_(outReset);
      } catch (errR) {
        return jsonOutput_({ ok: false, error: String(errR && errR.message ? errR.message : errR) });
      }
    }

    return jsonOutput_({ ok: false, error: 'Acao GET invalida.' });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doPost(e) {
  try {
    var payload = parseBody_(e);
    var action = String(payload.action || '').trim();

    if (action === 'append') {
      if (!payload.payload || typeof payload.payload !== 'object') {
        return jsonOutput_({ ok: false, error: 'Payload invalido.' });
      }
      appendResponse(payload.payload);
      return jsonOutput_({ ok: true });
    }

    if (action === 'getAll') {
      return jsonOutput_({ ok: true, data: getAllResponses() });
    }

    if (action === 'sendGiftEmails') {
      return jsonOutput_(sendGiftEmails(payload.recipients, payload.attachments, payload.mailOptions));
    }

    if (action === 'saveAdmins') {
      return jsonOutput_(saveAdmins(payload.admins));
    }

    if (action === 'resetResponses') {
      try {
        var secPost = String(payload.resetSecret || payload.adminPassword || '').trim();
        return jsonOutput_(resetCensusResponses(secPost));
      } catch (errP) {
        return jsonOutput_({ ok: false, error: String(errP && errP.message ? errP.message : errP) });
      }
    }

    return jsonOutput_({ ok: false, error: 'Acao POST invalida.' });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function configurarPrimeiraVez() {
  var folder = getTargetFolder_();
  var props = getPropsStore_();
  var propKey =
    typeof SCRIPT_PROP_ROOT_OVERRIDE !== 'undefined' && SCRIPT_PROP_ROOT_OVERRIDE
      ? String(SCRIPT_PROP_ROOT_OVERRIDE)
      : 'CENSO_ROOT_FOLDER_ID';
  if (!props.getProperty(PROP_BASE)) {
    props.setProperty(PROP_BASE, DEFAULT_BASE_NAME);
  }
  props.setProperty(propKey, folder.getId());
  ensureFirstShard_(folder);
  Logger.log('OK. Pasta ID: ' + folder.getId());
}

function configurarNaPasta(pastaId) {
  if (!pastaId || !String(pastaId).trim()) {
    throw new Error('Passe o ID da pasta.');
  }
  var folder = DriveApp.getFolderById(String(pastaId).trim());
  var props = getPropsStore_();
  var propKey =
    typeof SCRIPT_PROP_ROOT_OVERRIDE !== 'undefined' && SCRIPT_PROP_ROOT_OVERRIDE
      ? String(SCRIPT_PROP_ROOT_OVERRIDE)
      : 'CENSO_ROOT_FOLDER_ID';
  props.setProperty(propKey, folder.getId());
  if (!props.getProperty(PROP_BASE)) {
    props.setProperty(PROP_BASE, DEFAULT_BASE_NAME);
  }
  ensureFirstShard_(folder);
  Logger.log('OK. Pasta: ' + folder.getName() + ' ID: ' + folder.getId());
}

function diagnostico() {
  var folder = getTargetFolder_();
  ensureFirstShard_(folder);
  var files = listShardFiles_(folder);
  var names = files.map(function (f) {
    return f.getName();
  });
  return {
    ok: true,
    folderId: folder.getId(),
    folderName: folder.getName(),
    baseName: getBaseName_(),
    files: names,
    totalRespostas: getAllResponses().length
  };
}

/**
 * Limpa todo o banco JSON do censo e deixa pronto para recomeçar.
 * - Mantém somente o arquivo base principal
 * - Remove shards extras (_partXXX.json)
 * - Define o conteúdo final como []
 */
function limparBancoCensoAgora() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var folder = getTargetFolder_();
    var baseName = getBaseName_();
    var files = listShardFiles_(folder);
    /** Envia todos os fragmentos (.json base e _partXXX) para lixeira e recria um arquivo novo vazio ([ ]). Evita ficar só limpando o “primeiro” shard e manter dados noutros. */
    var i;
    for (i = 0; i < files.length; i++) {
      try {
        files[i].setTrashed(true);
      } catch (_) {}
    }
    var novo = folder.createFile(baseName + '.json', '[]', MimeType.PLAIN_TEXT);
    return {
      ok: true,
      message: 'Banco do censo limpo com sucesso (arquivos antigos para a lixeira; arquivo novo criado).',
      fileName: novo.getName(),
      folderId: folder.getId()
    };
  } finally {
    lock.releaseLock();
  }
}
