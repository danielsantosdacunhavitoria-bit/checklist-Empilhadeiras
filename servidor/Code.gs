/*  SERVIDOR DO CHECKLIST – Google Apps Script
    1) Edite a lista DESTINATARIOS abaixo.
    2) Execute a função "autorizar" uma vez (menu Executar) e aceite as permissões.
    3) Implantar > Nova implantação > Aplicativo da Web
       (Executar como: Eu | Quem tem acesso: Qualquer pessoa) e copie a URL.            */
var DESTINATARIOS = ['danielsantosdacunhavitoria@gmail.com'];   // e-mails que recebem os checklists
var NOME_REMETENTE = 'Desenvolvimento Operacional';

function autorizar() {
  sheet_(); folder_();
  Logger.log('Autorizado. Cota de e-mails restante hoje: ' + MailApp.getRemainingDailyQuota());
}
function sheet_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SHEET_ID'), ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create('Checklists – índice');
    p.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].appendRow(['id', 'equipamento', 'operador', 'equipe', 'turno', 'fim_ms', 'ts_ms', 'nao_conforme', 'arquivo_json', 'arquivo_pdf']);
  }
  return ss.getSheets()[0];
}
function folder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);
  var f = DriveApp.createFolder('Checklists'); p.setProperty('FOLDER_ID', f.getId()); return f;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

// Consulta (QR Code): só devolve checklists cuja escala ainda não terminou
function doGet(e) {
  try {
    var a = e.parameter.action, v = sheet_().getDataRange().getValues(), now = Date.now(), i;
    if (a === 'list') {
      var eq = String(e.parameter.eq || '').toUpperCase(), L = [];
      for (i = 1; i < v.length; i++)
        if (String(v[i][1]).toUpperCase() === eq && Number(v[i][5]) > now)
          L.push({ id: v[i][0], op: v[i][2], team: v[i][3], tn: v[i][4], fim: Number(v[i][5]), ts: Number(v[i][6]), bad: Number(v[i][7]) === 1 });
      L.sort(function (x, y) { return y.ts - x.ts; });
      return out_({ ok: 1, list: L });
    }
    if (a === 'get') {
      for (i = 1; i < v.length; i++)
        if (v[i][0] === e.parameter.id) {
          if (Number(v[i][5]) <= now) return out_({ ok: 0, err: 'expirado' });
          return out_({ ok: 1, rec: JSON.parse(DriveApp.getFileById(v[i][8]).getBlob().getDataAsString()) });
        }
      return out_({ ok: 0, err: 'não encontrado' });
    }
    return out_({ ok: 0, err: 'ação inválida' });
  } catch (err) { return out_({ ok: 0, err: String(err) }); }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var b = JSON.parse(e.postData.contents);
    if (b.action === 'submit') return out_(submit_(b));
    if (b.action === 'maint') return out_(maint_(b));
    return out_({ ok: 0, err: 'ação inválida' });
  } catch (err) { return out_({ ok: 0, err: String(err) }); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}

function submit_(b) {
  var r = b.rec; r.id = Utilities.getUuid().slice(0, 8);
  var fo = folder_(), f = fo.createFile(r.id + '.json', JSON.stringify(r), 'application/json'),
      pdf = fo.createFile(Utilities.newBlob(Utilities.base64Decode(b.pdf), 'application/pdf', b.pdfName));
  sheet_().appendRow([r.id, String(r.eq).toUpperCase(), r.op, r.team, r.tn, r.fim, r.ts, b.mail.bad ? 1 : 0, f.getId(), pdf.getId()]);
  mail_(b, r);
  return { ok: 1, id: r.id };
}

function maint_(b) {
  var sh = sheet_(), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === b.id) {
    if (Number(v[i][5]) <= Date.now()) return { ok: 0, err: 'expirado' };
    var f = DriveApp.getFileById(v[i][8]), r = JSON.parse(f.getBlob().getDataAsString());
    r.extra = r.extra || []; r.extra.push(b.item);
    f.setContent(JSON.stringify(r));
    sh.getRange(i + 1, 8).setValue(1);
    folder_().createFile(Utilities.newBlob(Utilities.base64Decode(b.pdf), 'application/pdf', b.pdfName));
    mail_(b, r);
    return { ok: 1, rec: r };
  }
  return { ok: 0, err: 'não encontrado' };
}

function mail_(b, r) {
  var m = b.mail, att = [Utilities.newBlob(Utilities.base64Decode(b.pdf), 'application/pdf', b.pdfName)];
  function fotos(a, pre) {
    (a || []).forEach(function (d, i) {
      att.push(Utilities.newBlob(Utilities.base64Decode(d.split(',')[1]), 'image/jpeg', pre + '_' + (i + 1) + '.jpg'));
    });
  }
  fotos(r.ph, 'avaria');
  (r.extra || []).forEach(function (x, i) { fotos(x.ph, 'manutencao' + (i + 1)); });

  var h = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:620px">' +
          '<p>Olá prezados!</p><p>Segue checklist do equipamento <b>' + esc_(r.eq) + '</b> preenchido pelo operador <b>' + esc_(r.op) + '</b>.</p>';
  var t = 'Olá prezados!\n\nSegue checklist do equipamento ' + r.eq + ' preenchido pelo operador ' + r.op + '.\n\n';
  if (m.lines && m.lines.length) {
    h += '<div style="border:2px solid #c62828;background:#fdeaea;color:#8e0000;padding:10px 12px;border-radius:6px;margin:12px 0">';
    m.lines.forEach(function (l) { h += '<div style="margin-bottom:6px"><b>ATENÇÃO:</b> ' + esc_(l) + '</div>'; t += 'ATENÇÃO: ' + l + '\n'; });
    h += '<div><b>Manutenção gentileza informar o número da O.S.</b></div></div>';
    t += 'Manutenção gentileza informar o número da O.S.\n\n';
  }
  h += '<p>Att;<br>Desenvolvimento Operacional.</p></div>';
  t += 'Att;\nDesenvolvimento Operacional.';
  MailApp.sendEmail({ to: DESTINATARIOS.join(','), subject: m.subject, htmlBody: h, body: t, name: NOME_REMETENTE, attachments: att });
}
