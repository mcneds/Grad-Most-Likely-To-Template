/**
 * Most Likely To — Apps Script backend (with moderation)
 * Sheets required:
 *   - Settings
 *   - Names         (headers: id, name)
 *   - Nominations   (headers: timestamp, device_id, name_id, text, status, mod_by, mod_ts, mod_note)
 *   - Votes         (headers: timestamp, device_id, nomination_id)
 *
 * Behavior
 *  - Multiple submissions per device
 *  - Blocks duplicate (device_id, name_id) pairs
 *  - New user nominations => status="pending" (admin auto-approve)
 *  - Public list shows only approved nominations
 *  - Admin Delete => status="rejected" and clears text, preserves device_id/name_id for re-submit blocking
 */

const SHEET = {
  settings: 'Settings',
  names: 'Names',
  nominations: 'Nominations',
  votes: 'Votes',
};

// ---------- helpers ----------
function getSheet(name){ return SpreadsheetApp.getActive().getSheetByName(name); }
function getProp(key){ return PropertiesService.getScriptProperties().getProperty(key); }
function _ok(data){ return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function _err(msg, code){ return ContentService.createTextOutput(JSON.stringify({ error: msg, code: code||400 })).setMimeType(ContentService.MimeType.JSON); }

// header utils (tolerant to case/whitespace/order)
function getHeaderMap(sh){
  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  const map = {};
  head.forEach((h,i)=>{ if(h) map[h] = i+1; });
  return map;
}
function needCols(map, cols){
  for (var i=0;i<cols.length;i++){
    if(!(cols[i] in map)) throw new Error('Missing header: '+cols[i]);
  }
  return true;
}

// ---------- router ----------
function doPost(e){
  try{
    const body = e.postData ? JSON.parse(e.postData.contents) : {};
    const action = body.action;
    if(!action) return _err('No action');

    switch (action) {
      case 'init':             return handleInit(body);
      case 'submit':           return handleSubmit(body);
      case 'list':             return handleList();
      case 'vote':             return handleVote(body);
      case 'adminDelete':      return handleAdminDelete(body);

      // moderation
      case 'adminListPending': return handleAdminListPending(body);
      case 'adminModerate':    return handleAdminModerate(body);

      default: return _err('Unknown action');
    }
  } catch(ex){
    return _err('Server error: '+ex, 500);
  }
}


// ---------- actions ----------
function handleInit(p){
  const sh = getSheet(SHEET.names);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return _ok({ names: [] });

  const head = values.shift().map(h => String(h).trim().toLowerCase());
  const iId   = head.indexOf('id');
  const iName = head.indexOf('name');
  if (iId < 0 || iName < 0) return _err('Names sheet missing headers id/name.', 500);

  const names = values
    .filter(r => r[iId] !== '')
    .map(r => ({ id: Number(r[iId]), name: String(r[iName]) }));

  return _ok({ names });
}

function handleSubmit(p){
  const device_id = String(p.device_id||'');
  const adminKey  = String(p.admin_key||'');
  const isAdmin   = !!adminKey && adminKey === getProp('ADMIN_KEY');
  if(!device_id && !isAdmin) return _err('Missing device_id');

  const entries = p.entries || [];
  const soft = String(p.soft||'');

  const sh = getSheet(SHEET.nominations);
  const map = getHeaderMap(sh);
  needCols(map, ['timestamp','device_id','name_id','text','status','mod_by','mod_ts','mod_note']);

  // Build duplicate set (device_id,name_id)
  const data = sh.getDataRange().getValues();
  data.shift(); // header
  const devCol  = map['device_id'] - 1;
  const nameCol = map['name_id'] - 1;
  const seen = {};
  for (var i=0;i<data.length;i++){
    const d = String(data[i][devCol]||'').split('#')[0];
    const n = String(data[i][nameCol]||'');
    if(d && n) seen[`${d}|${n}`] = true;
  }

  const ts = new Date();
  const rows = [];
  entries.forEach(e => {
    const text = String(e.text||'').trim();
    const name_id = Number(e.name_id||0);
    if(!text || !name_id) return;

    const status   = isAdmin ? 'approved' : 'pending';
    const mod_by   = isAdmin ? 'admin' : '';
    const mod_ts   = isAdmin ? ts : '';
    const mod_note = '';

    if(isAdmin){
      rows.push([ts, (device_id||'admin') + (soft? ('#'+soft) : ''), name_id, text, status, mod_by, mod_ts, mod_note]);
    } else {
      const key = `${device_id}|${name_id}`;
      if(seen[key]) return; // already nominated this person from this device
      seen[key] = true;
      rows.push([ts, device_id + (soft? ('#'+soft) : ''), name_id, text, status, mod_by, mod_ts, mod_note]);
    }
  });

  if(rows.length){
    sh.getRange(sh.getLastRow()+1, 1, rows.length, 8).setValues(rows);
  }

  return handleList(); // public list
}

function handleList(){
  const namesSh = getSheet(SHEET.names);
  const nomSh   = getSheet(SHEET.nominations);
  const voteSh  = getSheet(SHEET.votes);

  // Names
  const namesVals = namesSh.getDataRange().getValues();
  if (namesVals.length < 2) return _ok({ responses: [] });
  const namesHead = namesVals.shift().map(h => String(h).trim().toLowerCase());
  const nId   = namesHead.indexOf('id');
  const nName = namesHead.indexOf('name');
  const names = namesVals.map(r => ({ id:Number(r[nId]), name:String(r[nName]) }));

  // Nominations (approved only)
  const nomMap = getHeaderMap(nomSh);
  needCols(nomMap, ['name_id','text']); // status may be missing on very old sheets
  const nomVals = nomSh.getDataRange().getValues();
  const nomHead = nomVals.shift().map(h => String(h).trim().toLowerCase());
  const iNameId = nomHead.indexOf('name_id');
  const iText   = nomHead.indexOf('text');
  const iStatus = nomHead.indexOf('status'); // might be -1

  const nominations = [];
  for (var i=0;i<nomVals.length;i++){
    const r = nomVals[i];
    const text = String(r[iText]||'').trim();
    const name_id = Number(r[iNameId]||0);
    const statusCell = (iStatus === -1) ? 'approved' : String(r[iStatus]||'approved').toLowerCase();
    if(!text || !name_id) continue;
    if(statusCell !== 'approved') continue;
    const id = i+2; // 1-based row id (+ header)
    nominations.push({ id, name_id, text });
  }

  // Votes
  const voteVals = voteSh.getDataRange().getValues();
  if (voteVals.length === 0) return _ok({ responses: groupByName(names, nominations, {}) });
  const voteHead = voteVals.shift().map(h => String(h).trim().toLowerCase());
  const iNomId   = voteHead.indexOf('nomination_id');
  const voteCount = {};
  voteVals.forEach(r => {
    const nid = Number(r[iNomId]||0);
    if(nid) voteCount[nid] = (voteCount[nid]||0)+1;
  });

  return _ok({ responses: groupByName(names, nominations, voteCount) });
}

function groupByName(names, nominations, voteCount){
  const byName = {};
  names.forEach(n => byName[n.id] = { name_id: n.id, name: n.name, items: [] });
  nominations.forEach(n => {
    const g = byName[n.name_id] || (byName[n.name_id] = { name_id:n.name_id, name:`#${n.name_id}`, items: [] });
    g.items.push({ id: n.id, text: n.text, votes: voteCount[n.id]||0 });
  });
  return Object.values(byName).map(g => {
    g.items.sort((a,b) => (b.votes - a.votes) || a.text.localeCompare(b.text));
    return g;
  }).sort((a,b) => a.name.localeCompare(b.name));
}

function handleVote(p){
  const device_id = String(p.device_id||'');
  const nomination_id = Number(p.nomination_id||0);
  if(!device_id || !nomination_id) return _err('Missing vote fields');

  const sh = getSheet(SHEET.votes);
  const values = sh.getDataRange().getValues();
  if (values.length > 0){
    const head = values.shift().map(h => String(h).trim().toLowerCase());
    const iDev = head.indexOf('device_id');
    const iNid = head.indexOf('nomination_id');
    const dup = values.some(r => String(r[iDev])===device_id && Number(r[iNid])===nomination_id);
    if (dup) return _ok({ ok:true, duplicate:true });
  }
  sh.appendRow([new Date(), device_id + (p.soft? ('#'+String(p.soft)) : ''), nomination_id]);
  return _ok({ ok:true });
}

// Delete = reject + hide text, preserve device_id/name_id for duplicate blocking
function handleAdminDelete(p){
  const key = String(p.admin_key||'');
  if (key !== getProp('ADMIN_KEY')) return _err('Unauthorized', 401);

  const id = Number(p.nomination_id||0);
  if(!id) return _err('Missing nomination_id');

  const sh = getSheet(SHEET.nominations);
  const map = getHeaderMap(sh);
  needCols(map, ['text','status','mod_by','mod_ts','mod_note']);

  const who = Session.getActiveUser().getEmail() || 'admin';
  const now = new Date();

  sh.getRange(id, map['text']).setValue('');
  sh.getRange(id, map['status']).setValue('rejected');
  sh.getRange(id, map['mod_by']).setValue(who);
  sh.getRange(id, map['mod_ts']).setValue(now);
  sh.getRange(id, map['mod_note']).setValue('deleted-by-admin');

  return _ok({ ok:true });
}

// ---------- moderation ----------
function handleAdminListPending(p){
  const key = String(p.admin_key||'');
  if (key !== getProp('ADMIN_KEY')) return _err('Unauthorized', 401);

  const sh = getSheet(SHEET.nominations);
  const data = sh.getDataRange().getValues();
  const head = data.shift();
  const idxText   = head.indexOf('text');
  const idxNameId = head.indexOf('name_id');
  const idxStatus = head.indexOf('status');

  const pending = [];
  for (var i=0; i<data.length; i++){
    const r = data[i];
    const st = String(r[idxStatus]||'').trim().toLowerCase();
    if (st === '' || st === 'pending') {
      pending.push({ id: i+2, name_id: Number(r[idxNameId]||0), text: String(r[idxText]||'') });
    }
  }
  return _ok({ pending });
}

function handleAdminModerate(p){
  const key = String(p.admin_key || '');
  if (key !== getProp('ADMIN_KEY')) return _err('Unauthorized', 401);

  const id = Number(p.nomination_id || 0);      // sheet row (>=2)
  // Read decision from mod_action (preferred) or fall back to action
  const raw = String(p.mod_action || p.action || '').toLowerCase().trim();
  const decision = raw.startsWith('approve') ? 'approve'
                 : raw.startsWith('reject')  ? 'reject'
                 : '';
  const note = String(p.note || '');

  if (!id || id < 2)     return _err('Bad request: nomination_id must be a sheet row (>=2).');
  if (!decision)         return _err('Bad request: action must be approve|reject');

  const sh = getSheet(SHEET.nominations);
  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]
                 .map(h => String(h).trim().toLowerCase());
  const cStatus  = head.indexOf('status')   + 1;
  const cModBy   = head.indexOf('mod_by')   + 1;
  const cModTs   = head.indexOf('mod_ts')   + 1;
  const cModNote = head.indexOf('mod_note') + 1;
  if (cStatus <= 0 || cModBy <= 0 || cModTs <= 0 || cModNote <= 0)
    return _err('Nominations sheet missing moderation columns (status/mod_by/mod_ts/mod_note).', 500);

  const who = Session.getActiveUser().getEmail() || 'admin';
  const now = new Date();
  const statusVal = (decision === 'approve') ? 'approved' : 'rejected';

  // atomic write of the 4 moderation fields
  sh.getRange(id, cStatus, 1, 4).setValues([[statusVal, who, now, note]]);
  SpreadsheetApp.flush();

  return _ok({ ok:true, id, action: decision, status: statusVal });
}





