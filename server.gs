/** Utilities (no globals duplicated here) **/
function sh_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}
function ymd_(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function addDays_(d, n) { var x = new Date(d); x.setDate(x.getDate()+n); return x; }
function weekRange_(mondayDate) {
  var start = new Date(mondayDate);
  var end = addDays_(start, 6);
  return {
    startIso: ymd_(start),
    endIso: ymd_(end),
    // simple human label: dd/MM – dd/MM
    label: Utilities.formatDate(start, Session.getScriptTimeZone(), "dd/MM") + " – " +
           Utilities.formatDate(end,   Session.getScriptTimeZone(), "dd/MM")
  };
}

/** Web App entry */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Chấm công')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Build current-week structures once (parse slot_id; ignore RESPONSES.week_start lag) */
function _buildWeekModel_() {
  var wkDate = getSelectedWeek_();
  var wkStr = ymd_(wkDate);
  var range = weekRange_(wkDate);

  // PEOPLE (fixed)
  var pVals = sh_(TABS.PEOPLE).getDataRange().getValues();
  var pHdr = pVals.shift();
  var iPid = pHdr.indexOf('person_id');
  var iPname = pHdr.indexOf('pname');
  var people = [];
  var pmap = {}; // person_id -> name
  for (var i=0;i<pVals.length;i++) {
    var r = pVals[i]; if (!r[iPid]) continue;
    var pid = String(r[iPid]);
    var nm  = r[iPname];
    people.push({ person_id: pid, pname: nm });
    pmap[pid] = nm;
  }

  // SLOTS (only current week)
  var sVals = sh_(TABS.SLOTS).getDataRange().getValues();
  var sHdr = sVals.shift();
  var iSlot = sHdr.indexOf('slot_id');
  var iWk   = sHdr.indexOf('week_start');
  var iDay  = sHdr.indexOf('day');
  var iBlk  = sHdr.indexOf('block');
  var slots = [];
  for (var j=0;j<sVals.length;j++) {
    var s = sVals[j];
    if (!s[iWk] || ymd_(s[iWk]) !== wkStr) continue;
    slots.push({ slot_id: String(s[iSlot]), day: s[iDay], block: s[iBlk] });
  }

  // RESPONSES for current week (parse slot_id yyyy-mm-dd|DAY|BLOCK)
  var rVals = sh_(TABS.RESPONSES).getDataRange().getValues();
  var rHdr = rVals.shift();
  var irPid = rHdr.indexOf('person_id');
  var irSlot = rHdr.indexOf('slot_id');
  var irStatus = rHdr.indexOf('status');

  var responsesByPerson = {}; // { person_id: { slot_id: status } }
  var countsBySlot = {};      // { slot_id: {Free,Class,x,total} }
  var rosterBySlot = {};      // { slot_id: {Free:[{pid,name}], Class:[...], x:[...]} }

  function ensureCount(sid) {
    if (!countsBySlot[sid]) countsBySlot[sid] = { Free:0, Class:0, x:0, total:0 };
    return countsBySlot[sid];
  }
  function ensureRoster(sid) {
    if (!rosterBySlot[sid]) rosterBySlot[sid] = { Free:[], Class:[], x:[] };
    return rosterBySlot[sid];
  }

  for (var k=0;k<rVals.length;k++) {
    var row = rVals[k];
    var pid0 = row[irPid];
    var sid0 = row[irSlot];
    var st0  = row[irStatus] || '';
    if (!pid0 || !sid0) continue;

    var sid = String(sid0);
    var parts = sid.split('|');         // yyyy-mm-dd | DAY | BLOCK
    if (parts.length < 3) continue;
    if (parts[0] !== wkStr) continue;   // not this week

    var pid = String(pid0);
    if (!responsesByPerson[pid]) responsesByPerson[pid] = {};
    responsesByPerson[pid][sid] = st0;

    if (st0 === 'Free' || st0 === 'Class' || st0 === 'x') {
      var c = ensureCount(sid); c[st0]++; c.total++;
      var r = ensureRoster(sid); r[st0].push({ pid: pid, name: pmap[pid] || pid });
    }
  }

  return {
    weekStart: wkStr,
    weekRange: range, // {startIso,endIso,label}
    people: people,
    slots: slots,
    responsesByPerson: responsesByPerson,
    countsBySlot: countsBySlot,
    rosterBySlot: rosterBySlot,
    days: DAYS, blocks: BLOCKS, statuses: STATUS_VALUES
  };
}

/** Initial payload for UI */
function getInitData() { return _buildWeekModel_(); }

/** Lightweight refresh (after save) */
function getState() {
  var m = _buildWeekModel_();
  return {
    responsesByPerson: m.responsesByPerson,
    countsBySlot: m.countsBySlot,
    rosterBySlot: m.rosterBySlot
  };
}

/** Change week by +/- N weeks (e.g., -1 for previous, +1 for next) and return fresh model */
function changeWeek(deltaWeeks) {
  var sh = sh_(TABS.SETTINGS);
  var cur = getSelectedWeek_();
  var next = addDays_(cur, 7 * (deltaWeeks || 0));
  sh.getRange(SETTINGS_RANGE).setValue(next);
  return _buildWeekModel_();
}

/** Upsert rows submitted from UI (A–E; F–G computed by sheet) */
function saveResponses(rows) {
  if (!Array.isArray(rows) || !rows.length) return {updated:0, inserted:0};

  var sh = sh_(TABS.RESPONSES);
  var data = sh.getDataRange().getValues();
  var hdr  = data[0];
  var body = data.slice(1);

  var iRespondId = hdr.indexOf('respond_id');
  var iPid   = hdr.indexOf('person_id');
  var iSlot  = hdr.indexOf('slot_id');
  var iStatus= hdr.indexOf('status');
  var iTs    = hdr.indexOf('timestamp');

  var idx = new Map(); // person|slot -> row#
  body.forEach(function(r, i){
    var key = String(r[iPid]) + '|' + String(r[iSlot]);
    idx.set(key, i+2);
  });

  var now = new Date();
  var updated = 0, inserted = 0;

  rows.forEach(function(entry){
    var pid = String(entry.person_id || '');
    var sid = String(entry.slot_id || '');
    var st  = entry.status || '';
    if (!pid || !sid) return;

    var key = pid + '|' + sid;
    var rowNo = idx.get(key);
    if (rowNo) {
      if (st === '') { sh.deleteRow(rowNo); }
      else {
        sh.getRange(rowNo, iStatus+1).setValue(st);
        if (iTs >= 0) sh.getRange(rowNo, iTs+1).setValue(now);
      }
      updated++;
    } else {
      if (st === '') return;
      var rowArray = new Array(hdr.length).fill('');
      rowArray[iRespondId] = Utilities.getUuid();
      rowArray[iPid]       = pid;
      rowArray[iSlot]      = sid;
      rowArray[iStatus]    = st;
      if (iTs >= 0) rowArray[iTs] = now;
      sh.appendRow(rowArray);
      inserted++;
    }
  });

  return {updated:updated, inserted:inserted};
}
