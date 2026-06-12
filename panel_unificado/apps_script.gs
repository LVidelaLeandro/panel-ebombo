// ═══════════════════════════════════════════════════════
//  Panel SDR eBombo — Apps Script
//  Maneja registros diarios + datos de Mi Panel por SDR
// ═══════════════════════════════════════════════════════

const SHEET_REGISTROS = 'Registros';
const SHEET_MI_PANEL  = 'Mi Panel';

function doGet(e) {
  const action   = e.parameter.action   || 'get';
  const callback = e.parameter.callback || '';
  let result;
  try {
    if      (action === 'get')    result = getRegistros();
    else if (action === 'getMi')  result = getMiPanel(e.parameter.sdr, e.parameter.period);
    else if (action === 'status') result = { ok: true };
    else result = { ok: false, error: 'Acción no reconocida' };
  } catch(err) {
    result = { ok: false, error: err.message };
  }
  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    let result;
    if      (body.action === 'add')       result = addRegistro(body.record);
    else if (body.action === 'saveMi')    result = saveMiPanel(body.sdr, body.period, body.out, body.ib);
    else if (body.action === 'addCamp')   result = addCampRegistro(body.camp);
    else if (body.action === 'addReu')    result = addReuRegistro(body.reu);
    else if (body.action === 'sendEmail') result = sendReporteEmail(body.recipients, body.subject, body.htmlBody);
    else result = { ok: false, error: 'Acción no reconocida' };
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── ENVÍO DE REPORTES POR EMAIL ──────────────────────
function sendReporteEmail(recipients, subject, htmlBody) {
  if (!recipients || !recipients.length) return { ok: false, error: 'Sin destinatarios' };
  if (!subject || !htmlBody) return { ok: false, error: 'Faltan datos del email' };
  var errors = [];
  recipients.forEach(function(email) {
    try {
      MailApp.sendEmail({
        to:       email,
        subject:  subject,
        htmlBody: htmlBody,
      });
    } catch(err) {
      errors.push(email + ': ' + err.message);
    }
  });
  if (errors.length) return { ok: false, error: errors.join(' | ') };
  return { ok: true, sent: recipients.length };
}

// ─── REGISTROS ────────────────────────────────────────
function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0F172A')
      .setFontColor('#34D399');
  }
  return sheet;
}

function getRegistros() {
  const sheet = getSheet(SHEET_REGISTROS,
    ['id','fecha','sdr','pais','canal','contactados','respuestas','positivas','reuniones','notas']);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const records = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  return { ok: true, records: records };
}

function addRegistro(rec) {
  const sheet = getSheet(SHEET_REGISTROS,
    ['id','fecha','sdr','pais','canal','contactados','respuestas','positivas','reuniones','notas']);
  sheet.appendRow([
    rec.id || Date.now(),
    rec.fecha    || '',
    rec.sdr      || '',
    rec.pais     || '',
    rec.canal    || '',
    rec.contactados || 0,
    rec.respuestas  || 0,
    rec.positivas   || 0,
    rec.reuniones   || 0,
    rec.notas    || '',
  ]);
  return { ok: true };
}

// ─── MI PANEL (datos Outbound / Inbound por SDR) ──────
function getMiPanel(sdr, period) {
  if (!sdr || !period) return { ok: false, error: 'Faltan parámetros' };
  const sheet = getSheet(SHEET_MI_PANEL,
    ['key','sdr','period','tipo','datos','updated']);
  const data = sheet.getDataRange().getValues();

  var out = null, ib = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === sdr && data[i][2] === period) {
      if (data[i][3] === 'out') { try { out = JSON.parse(data[i][4]); } catch(e){} }
      if (data[i][3] === 'ib')  { try { ib  = JSON.parse(data[i][4]); } catch(e){} }
    }
  }
  return { ok: true, out: out, ib: ib };
}

function saveMiPanel(sdr, period, out, ib) {
  const sheet = getSheet(SHEET_MI_PANEL,
    ['key','sdr','period','tipo','datos','updated']);
  const data = sheet.getDataRange().getValues();
  const ts   = new Date().toLocaleString('es-ES');

  ['out', 'ib'].forEach(function(tipo) {
    const payload = tipo === 'out' ? out : ib;
    if (!payload) return;
    const key = sdr + '_' + period + '_' + tipo;
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 5).setValue(JSON.stringify(payload));
        sheet.getRange(i + 1, 6).setValue(ts);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, sdr, period, tipo, JSON.stringify(payload), ts]);
    }
  });
  return { ok: true };
}

// ─── CAMPAÑAS ─────────────────────────────────────────
function addCampRegistro(camp) {
  const sheet = getSheet('Campañas',
    ['id','sdr','pais','fecha','nombre','contactos','respuestas','reuniones','updated']);
  // Upsert: buscar si ya existe
  const data = sheet.getDataRange().getValues();
  const ts   = new Date().toLocaleString('es-ES');
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === camp.id) {
      sheet.getRange(i+1, 1, 1, 9).setValues([[
        camp.id, camp.sdr||'', camp.pais||'', camp.fecha||'',
        camp.nombre||'', camp.contactos||0, camp.respuestas||0,
        camp.reuniones||0, ts
      ]]);
      return { ok: true };
    }
  }
  sheet.appendRow([
    camp.id, camp.sdr||'', camp.pais||'', camp.fecha||'',
    camp.nombre||'', camp.contactos||0, camp.respuestas||0,
    camp.reuniones||0, ts
  ]);
  return { ok: true };
}

// ─── REUNIONES EXTRA ──────────────────────────────────
function addReuRegistro(reu) {
  const sheet = getSheet('Reuniones Extra',
    ['_id','mes','semana','sdr','empresa','realizada','pais','canal','cotizacion','estado','monto','updated']);
  const ts = new Date().toLocaleString('es-ES');
  // Upsert
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === reu._id) {
      sheet.getRange(i+1, 1, 1, 12).setValues([[
        reu._id, reu.mes||0, reu.semana||0, reu.sdr||'',
        reu.empresa||'', reu.realizada||'', reu.pais||'', reu.canal||'',
        reu.cotizacion||'', reu.estado||'', reu.monto||0, ts
      ]]);
      return { ok: true };
    }
  }
  sheet.appendRow([
    reu._id, reu.mes||0, reu.semana||0, reu.sdr||'',
    reu.empresa||'', reu.realizada||'', reu.pais||'', reu.canal||'',
    reu.cotizacion||'', reu.estado||'', reu.monto||0, ts
  ]);
  return { ok: true };
}
