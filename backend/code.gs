/**
 * Most Likely To — Apps Script backend
 * Storage: Google Sheet with tabs Settings, Names, Nominations, Votes
 * Update: allow multiple submissions per device; block duplicate (device_id,name_id) pairs.
 * Admin (ADMIN_KEY) can bypass and submit anything anytime.
 */

const SHEET = {
  settings: 'Settings',
  names: 'Names',
  nominations: 'Nominations',
  votes: 'Votes',
};

function getSheet(name){ return SpreadsheetApp.getActive().getSheetByName(name); }
function getProp(key){ return PropertiesService.getScriptProperties().getProperty(key); }

function _ok(data){
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function _err(msg, code){
  return ContentService.createTextOutput(JSON.stringify({ error: msg, code: code||400 }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    const body = e.postData ? JSON.parse(e.postData.contents) : {};
    const action = body.action;
    if(!action) return _err('No action');
    switch(action){
      case 'init': return handleInit(body);
      case 'submit': return handleSubmit(body);
      case 'list': return handleList();
      case 'vote': return handleVote(body);
      case 'adminDelete': return handleAdminDelete(body);
      default: return _err('Unknown action');
    }
  }catch(ex){
    return _err('Server error: '+ex, 500);
  }
}

function handleInit(p){
  const namesSh = getSheet(SHEET.names);
  const data = namesSh.getDataRange().getValues();
  const header = data.shift();
  const idIdx = header.indexOf('id');
  const nameIdx = header.indexOf('name');
  const names = data.filter(r => r[idIdx] !== '').map(r => ({ id: Number(r[idIdx]), name: String(r[nameIdx]) }));
  // No “alreadySubmitted” check anymore — frontend always shows the form
  return _ok({ names: names });
}

function handleSubmit(p){
  const device_id = String(p.device_id||'');
  const adminKey = String(p.admin_key||'');
  const isAdmin = !!adminKey && adminKey === getProp('ADMIN_KEY');
  if(!device_id && !isAdmin) return _err('Missing device_id');

  const entries = p.entries || [];
  const soft = String(p.soft||'');

  const nomSh = getSheet(SHEET.nominations);
  const nomData = nomSh.getDataRange().getValues();
  const nomHead = nomData.shift();
  const devIdx = nomHead.indexOf('device_id');
  const nameIdx = nomHead.indexOf('name_id');

  // Build a set of existing (device_id, name_id) to prevent duplicate-to-same-person from same device
  const seen = {};
  nomData.forEach(r => {
    const d = String(r[devIdx]||'').split('#')[0]; // ignore soft fingerprint suffix
    const n = String(r[nameIdx]||'');
    if(d && n) seen[`${d}|${n}`] = true;
  });

  const ts = new Date();
  const rows = [];
  entries.forEach(e => {
    const text = String(e.text||'').trim();
    const name_id = Number(e.name_id);
    if(!text || !name_id) return;

    if(isAdmin){
      rows.push([ts, (device_id||'admin') + (soft? ('#'+soft) : ''), name_id, text]);
    } else {
      const key = `${device_id}|${name_id}`;
      if(seen[key]) return; // already nominated this person from this device
      seen[key] = true;
      rows.push([ts, device_id + (soft? ('#'+soft) : ''), name_id, text]);
    }
  });

  if(rows.length){
    nomSh.getRange(nomSh.getLastRow()+1, 1, rows.length, 4).setValues(rows);
  }

  return handleList();
}

function handleList(){
  const namesSh = getSheet(SHEET.names);
  const nomSh = getSheet(SHEET.nominations);
  const voteSh = getSheet(SHEET.votes);

  const namesData = namesSh.getDataRange().getValues();
  const nh = namesData.shift();
  const idIdx = nh.indexOf('id');
  const nmIdx = nh.indexOf('name');
  const names = namesData.map(r => ({ id:Number(r[idIdx]), name:String(r[nmIdx]) }));

  const nomData = nomSh.getDataRange().getValues();
  const nomHead = nomData.shift();
  const nomNameId = nomHead.indexOf('name_id');
  const nomText = nomHead.indexOf('text');

  const nominations = [];
  for (var i=0; i<nomData.length; i++){
    const r = nomData[i];
    const text = String(r[nomText]||'').trim();
    const name_id = Number(r[nomNameId]||0);
    if(!text || !name_id) continue;
    const id = i+2; // row id (1-based + header)
    nominations.push({ id:id, name_id:name_id, text:text });
  }

  const voteData = voteSh.getDataRange().getValues();
  const vh = voteData.shift();
  const vNom = vh.indexOf('nomination_id');
  const voteCount = {};
  voteData.forEach(r => {
    const nid = Number(r[vNom]||0);
    if(!nid) return;
    voteCount[nid] = (voteCount[nid]||0)+1;
  });

  const byName = {};
  names.forEach(n => byName[n.id] = { name_id: n.id, name: n.name, items: [] });
  nominations.forEach(n => {
    const g = byName[n.name_id] || (byName[n.name_id] = { name_id:n.name_id, name:`#${n.name_id}`, items: [] });
    g.items.push({ id: n.id, text: n.text, votes: voteCount[n.id]||0 });
  });

  const responses = Object.values(byName).map(g => {
    g.items.sort((a,b) => (b.votes - a.votes) || a.text.localeCompare(b.text));
    return g;
  }).sort((a,b) => a.name.localeCompare(b.name));

  return _ok({ responses });
}

function handleVote(p){
  const device_id = String(p.device_id||'');
  const nomination_id = Number(p.nomination_id||0);
  if(!device_id || !nomination_id) return _err('Missing vote fields');

  const voteSh = getSheet(SHEET.votes);
  const data = voteSh.getDataRange().getValues();
  const head = data.shift();
  const devIdx = head.indexOf('device_id');
  const nidIdx = head.indexOf('nomination_id');

  const already = data.some(r => String(r[devIdx])===device_id && Number(r[nidIdx])===nomination_id);
  if(already) return _ok({ ok:true, duplicate:true });

  voteSh.appendRow([new Date(), device_id + (p.soft? ('#'+String(p.soft)) : ''), nomination_id]);
  return _ok({ ok:true });
}

function handleAdminDelete(p){
  const key = String(p.admin_key||'');
  const propKey = getProp('ADMIN_KEY');
  if(!propKey || key !== propKey) return _err('Unauthorized', 401);

  const id = Number(p.nomination_id||0);
  if(!id) return _err('Missing nomination_id');

  const nomSh = getSheet(SHEET.nominations);
  // soft delete: blank the row (keeps timestamp/audit)
  nomSh.getRange(id, 1, 1, 4).setValues([[new Date(), '', '', '']]);
  return _ok({ ok:true });
}
