/**
 * MIC LifeOS 1.1 — Google Apps Script
 * Incollare questo codice in Estensioni > Apps Script del Foglio Google.
 * Cambiare SECRET_KEY prima della distribuzione.
 */
const SECRET_KEY = 'CAMBIA_QUESTA_CHIAVE';
const PROP_DATA = 'MIC_LIFEOS_DATA';

function doPost(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    if (p.key !== SECRET_KEY) return text_({ok:false,error:'Chiave non valida'});
    if (p.action !== 'push') return text_({ok:false,error:'Azione non valida'});
    const incoming = JSON.parse(p.payload || '{}');
    const merged = merge_(load_(), incoming);
    PropertiesService.getScriptProperties().setProperty(PROP_DATA, JSON.stringify(merged));
    writeSheets_(merged);
    return text_({ok:true,updatedAt:new Date().toISOString()});
  } catch (err) {
    return text_({ok:false,error:String(err)});
  }
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const callback = String(p.callback || '').replace(/[^a-zA-Z0-9_$]/g, '');
  let result;
  if (p.key !== SECRET_KEY) result = {ok:false,error:'Chiave non valida'};
  else if (p.action !== 'pull') result = {ok:false,error:'Azione non valida'};
  else result = {ok:true,data:load_()};
  const json = JSON.stringify(result);
  if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function load_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_DATA);
  if (!raw) return {version:'1.1',days:{},eurekas:[]};
  try { return JSON.parse(raw); } catch (e) { return {version:'1.1',days:{},eurekas:[]}; }
}

function merge_(local, incoming) {
  local = local || {days:{},eurekas:[]};
  incoming = incoming || {days:{},eurekas:[]};
  const days = Object.assign({}, local.days || {});
  Object.keys(incoming.days || {}).forEach(k => {
    const a = days[k], b = incoming.days[k];
    if (!a || new Date(b.updatedAt || 0) >= new Date(a.updatedAt || 0)) days[k] = b;
  });
  const map = {};
  (local.eurekas || []).forEach(x => { if (x.id) map[x.id] = x; });
  (incoming.eurekas || []).forEach(x => {
    if (!x.id) return;
    const a = map[x.id];
    if (!a || new Date(x.updatedAt || x.at || 0) >= new Date(a.updatedAt || a.at || 0)) map[x.id] = x;
  });
  return {version:'1.1',deviceUpdatedAt:incoming.deviceUpdatedAt || new Date().toISOString(),days:days,eurekas:Object.values(map)};
}

function writeSheets_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const giorni = sheet_(ss, 'Giorni', ['Data','Passi','Corsa km','Bici km','Lavoro manuale min','Peso kg','Sonno ore','Energia','Recupero','Dolore','MIC Score','Riepilogo serale','Aggiornato']);
  const eureka = sheet_(ss, 'Eureka', ['ID','Creata','Aggiornata','Stato','Testo']);

  const dayRows = Object.keys(data.days || {}).sort().map(k => {
    const d = data.days[k] || {};
    return [k,d.steps||0,d.run||0,d.bike||0,d.manual||0,d.weight||'',d.sleep||'',d.energy||'',d.recovery||'',d.pain||'',d.score||'',d.evening||'',d.updatedAt||''];
  });
  rewrite_(giorni, dayRows, 13);

  const eurekaRows = (data.eurekas || []).sort((a,b)=>new Date(a.at)-new Date(b.at)).map(x => [x.id||'',x.at||'',x.updatedAt||'',x.status||'Idea',x.text||'']);
  rewrite_(eureka, eurekaRows, 5);
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function rewrite_(sh, rows, cols) {
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2,1,last-1,Math.max(cols,sh.getLastColumn())).clearContent();
  if (rows.length) sh.getRange(2,1,rows.length,cols).setValues(rows);
  sh.autoResizeColumns(1,cols);
}

function text_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
