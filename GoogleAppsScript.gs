/** MIC LifeOS 1.2 STABLE — Google Apps Script */
const SECRET_KEY = 'MICSilvioLifeOS2026'; // deve coincidere con la chiave nella PWA
const PROP_DATA = 'MIC_LIFEOS_DATA_V12';
const MAX_BACKUPS = 50;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = e && e.parameter ? e.parameter : {};
    if (p.key !== SECRET_KEY) return text_({ok:false,error:'Chiave non valida'});
    if (p.action !== 'push') return text_({ok:false,error:'Azione non valida'});
    const incoming = normalizeData_(JSON.parse(p.payload || '{}'));
    const current = normalizeData_(load_());
    const merged = merge_(current, incoming);
    save_(merged, 'push da ' + (incoming.deviceId || 'dispositivo'));
    return text_({ok:true,updatedAt:new Date().toISOString(),eurekaCount:(merged.eurekas||[]).length});
  } catch (err) {
    return text_({ok:false,error:String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const callback = String(p.callback || '').replace(/[^a-zA-Z0-9_$]/g, '');
  let result;
  if (p.key !== SECRET_KEY) result = {ok:false,error:'Chiave non valida'};
  else if (p.action === 'pull') result = {ok:true,data:load_()};
  else if (p.action === 'health') result = {ok:true,version:'1.2',at:new Date().toISOString()};
  else result = {ok:false,error:'Azione non valida'};
  const json = JSON.stringify(result);
  if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function load_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_DATA);
  if (!raw) return {version:'1.2',days:{},eurekas:[],meta:{updatedAt:new Date().toISOString()}};
  try { return normalizeData_(JSON.parse(raw)); } catch (_) { return {version:'1.2',days:{},eurekas:[],meta:{updatedAt:new Date().toISOString()}}; }
}

function save_(data, reason) {
  data.version = '1.2';
  data.meta = Object.assign({}, data.meta || {}, {updatedAt:new Date().toISOString()});
  const json = JSON.stringify(data);
  PropertiesService.getScriptProperties().setProperty(PROP_DATA, json);
  writeSheets_(data);
  appendBackup_(json, reason);
}

function merge_(server, incoming) {
  server = normalizeData_(server || {days:{},eurekas:[]});
  incoming = normalizeData_(incoming || {days:{},eurekas:[]});
  const days = Object.assign({}, server.days || {});
  Object.keys(incoming.days || {}).forEach(k => {
    const a = days[k], b = incoming.days[k];
    if (!a || new Date(b.updatedAt || 0) >= new Date(a.updatedAt || 0)) days[k] = b;
  });
  const map = {};
  (server.eurekas || []).forEach(x => { if (x.id) map[x.id] = x; });
  (incoming.eurekas || []).forEach(x => {
    const a = map[x.id];
    if (eurekaWins_(x, a)) map[x.id] = x;
  });
  return {version:'1.2',days:days,eurekas:Object.values(map),meta:{updatedAt:new Date().toISOString(),lastDevice:incoming.deviceId||''}};
}

function legacyEurekaId_(x) {
  const source = [x.at || '', x.text || '', x.status || 'Idea'].join('|');
  let hash = 2166136261;
  for (let i = 0; i < source.length; i++) {
    hash ^= source.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return 'legacy-' + (hash >>> 0).toString(16).padStart(8, '0');
}

function normalizeData_(data) {
  data = data || {};
  data.days = data.days || {};
  data.eurekas = Array.isArray(data.eurekas) ? data.eurekas : [];
  data.eurekas.forEach(x => {
    if (!x.id) x.id = legacyEurekaId_(x);
    if (!x.updatedAt) x.updatedAt = x.at || new Date().toISOString();
    const revision = Number(x.revision);
    x.revision = isFinite(revision) && revision >= 1 ? revision : 1;
    if (typeof x.deleted !== 'boolean') x.deleted = false;
  });
  return data;
}

function eurekaWins_(candidate, current) {
  if (!current) return true;
  const candidateRevision = Number(candidate.revision) || 1;
  const currentRevision = Number(current.revision) || 1;
  if (candidateRevision !== currentRevision) return candidateRevision > currentRevision;
  const candidateTime = new Date(candidate.updatedAt || candidate.at || 0).getTime() || 0;
  const currentTime = new Date(current.updatedAt || current.at || 0).getTime() || 0;
  if (candidateTime !== currentTime) return candidateTime > currentTime;
  return candidate.deleted === true && current.deleted !== true;
}

function writeSheets_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const giorni = sheet_(ss,'Giorni',['Data','Passi','Corsa km','Bici km','Lavoro manuale min','Peso kg','Sonno ore','Energia','Recupero','Dolore','MIC Score','Riepilogo serale','Aggiornato']);
  const eureka = sheet_(ss,'Eureka',['ID','Creata','Aggiornata','Stato','Testo','Eliminata','Revisione']);
  const dayRows = Object.keys(data.days||{}).sort().map(k=>{const d=data.days[k]||{};return[k,d.steps||0,d.run||0,d.bike||0,d.manual||0,d.weight||'',d.sleep||'',d.energy||'',d.recovery||'',d.pain||'',d.score||'',d.evening||'',d.updatedAt||''];});
  rewrite_(giorni,dayRows,13);
  const rows=(data.eurekas||[]).sort((a,b)=>new Date(a.at||0)-new Date(b.at||0)).map(x=>[x.id||'',x.at||'',x.updatedAt||'',x.status||'Idea',x.text||'',x.deleted?'SI':'NO',x.revision||1]);
  rewrite_(eureka,rows,7);
}

function appendBackup_(json, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = sheet_(ss,'Backup',['Data','Motivo','JSON completo']);
  sh.appendRow([new Date(),reason||'',json]);
  const extra = sh.getLastRow() - 1 - MAX_BACKUPS;
  if (extra > 0) sh.deleteRows(2, extra);
  sh.hideSheet();
}

function sheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');sh.setFrozenRows(1);return sh;}
function rewrite_(sh,rows,cols){const last=sh.getLastRow();if(last>1)sh.getRange(2,1,last-1,Math.max(cols,sh.getLastColumn())).clearContent();if(rows.length)sh.getRange(2,1,rows.length,cols).setValues(rows);sh.autoResizeColumns(1,cols);}
function text_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
