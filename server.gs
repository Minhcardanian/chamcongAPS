
/***** CORE HELPERS (idempotent) *****/
function _ss() { return SpreadsheetApp.getActive(); }
function _sheet(name) {
  var sh = _ss().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}
function _ymd(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function _addDays(d, n) { var x = new Date(d); x.setDate(x.getDate()+n); return x; }

/** Normalize any date to the Monday of its week (00:00), using SETTINGS!B1 as source */
function _getSelectedWeekDate() {
  var sh = _sheet(TABS.SETTINGS);
  var v = sh.getRange(SETTINGS_RANGE).getValue();
  if (!v) throw new Error('SETTINGS!'+SETTINGS_RANGE+' empty');
  var d = new Date(v);
  var day = d.getDay();            // Sun=0, Mon=1
  var delta = (day === 0 ? -6 : 1 - day);
  d.setDate(d.getDate() + delta);
  d.setHours(0,0,0,0);
  return d;
}

/** Week label helper: dd/MM – dd/MM (and start/end ISO) */
function _weekRangeLabel(mondayDate) {
  var start = new Date(mondayDate);
  var end = _addDays(start, 6);
  return {
    startIso: _ymd(start),
    endIso: _ymd(end),
    label:
      Utilities.formatDate(start, Session.getScriptTimeZone(), "dd/MM") +
      " – " +
      Utilities.formatDate(end,   Session.getScriptTimeZone(), "dd/MM")
  };
}

/***** WEB APP ENTRY *****/
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Chấm công')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** MODEL BUILD (current week) *****/
function _buildWeekModel_() {
  var monday = _getSelectedWeekDate();
  var wkStr = _ymd(monday);
  var range = _weekRangeLabel(monday);

  // PEOPLE
  var pVals = _sheet(TABS.PEOPLE).getDataRange().getValues();
  var pHdr = pVals.shift();
  var iPid = pHdr.indexOf('person_id');
  var iPname = pHdr.indexOf('pname');
  var people = [];
  var pmap = {}; // pid -> name
  for (var i=0;i<pVals.length;i++) {
    var r = pVals[i]; if (!r[iPid]) continue;
    var pid = String(r[iPid]);
    var nm  = r[iPname];
    people.push({ person_id: pid, pname: nm });
    pmap[pid] = nm;
  }

  // SLOTS (optional: restricted to this week if SLOTS has week_start)
  var sVals = _sheet(TABS.SLOTS).getDataRange().getValues();
  var sHdr = sVals.shift();
  var iSlot = sHdr.indexOf('slot_id');
  var iWk   = sHdr.indexOf('week_start');
  var iDay  = sHdr.indexOf('day');
  var iBlk  = sHdr.indexOf('block');
  var slots = [];
  for (var j=0;j<sVals.length;j++) {
    var s = sVals[j];
    if (iWk >= 0) {
      if (!s[iWk] || _ymd(s[iWk]) !== wkStr) continue;
    }
    slots.push({ slot_id: String(s[iSlot]), day: s[iDay], block: s[iBlk] });
  }

  // RESPONSES (derive week from slot_id "yyyy-mm-dd|DAY|BLOCK" if no week_start col)
  var rVals = _sheet(TABS.RESPONSES).getDataRange().getValues();
  var rHdr = rVals.shift();
  var irPid = rHdr.indexOf('person_id');
  var irSlot = rHdr.indexOf('slot_id');
  var irStatus = rHdr.indexOf('status');
  var irWeekStart = rHdr.indexOf('week_start');

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
    var wkFromCol = (irWeekStart >= 0 && row[irWeekStart]) ? _ymd(row[irWeekStart]) : null;
    var wkFromSid = (!wkFromCol) ? (String(sid).split('|')[0] || '') : null;
    var wk = wkFromCol || wkFromSid || '';
    if (wk !== wkStr) continue; // skip other weeks

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
    weekRange: range, // {startIso, endIso, label}
    people: people,
    slots: slots,
    responsesByPerson: responsesByPerson,
    countsBySlot: countsBySlot,
    rosterBySlot: rosterBySlot,
    days: DAYS, blocks: BLOCKS, statuses: STATUS_VALUES
  };
}

/***** INIT + REFRESH *****/
function getInitData() { return _buildWeekModel_(); }

function getState() {
  var m = _buildWeekModel_();
  return {
    responsesByPerson: m.responsesByPerson,
    countsBySlot: m.countsBySlot,
    rosterBySlot: m.rosterBySlot
  };
}

/***** WEEK NAV *****/
function changeWeek(deltaWeeks) {
  deltaWeeks = deltaWeeks || 0;
  var sh = _sheet(TABS.SETTINGS);
  var cur = _getSelectedWeekDate();
  var next = _addDays(cur, 7 * deltaWeeks);
  sh.getRange(SETTINGS_RANGE).setValue(next);
  return _buildWeekModel_();
}

/***** UPSERT RESPONSES *****/
function saveResponses(rows) {
  if (!Array.isArray(rows) || !rows.length) return {updated:0, inserted:0};

  var sh = _sheet(TABS.RESPONSES);
  var data = sh.getDataRange().getValues();
  var hdr  = data[0];
  var body = data.slice(1);

  var iRespondId = hdr.indexOf('respond_id');
  var iPid   = hdr.indexOf('person_id');
  var iSlot  = hdr.indexOf('slot_id');
  var iStatus= hdr.indexOf('status');
  var iTs    = hdr.indexOf('timestamp');

  // Build index once (person|slot -> row number)
  var idx = new Map();
  for (var i=0;i<body.length;i++) {
    var r = body[i];
    idx.set(String(r[iPid]) + '|' + String(r[iSlot]), i+2);
  }

  var now = new Date();
  var updated = 0, inserted = 0;
  // IMPORTANT: when deleting rows, do it after loop to avoid shifting. We’ll track deletions.
  var toDelete = [];

  rows.forEach(function(entry){
    var pid = String(entry.person_id || '');
    var sid = String(entry.slot_id || '');
    var st  = entry.status || '';
    if (!pid || !sid) return;

    var key = pid + '|' + sid;
    var rowNo = idx.get(key);
    if (rowNo) {
      if (st === '') {
        toDelete.push(rowNo);
      } else {
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

  // Delete rows bottom-up to avoid shifting
  toDelete.sort(function(a,b){ return b-a; }).forEach(function(rowNo){
    sh.deleteRow(rowNo);
  });

  return {updated:updated, inserted:inserted};
}

/***** WEEKLY SUMMARY (optional; counts Free by default) *****/
function getWeeklySummary(opts) {
  var statusesToCount = (opts && opts.statusesToCount) || ['Free'];

  var monday = _getSelectedWeekDate();
  var weekStr = _ymd(monday);

  var people = new Map(); // pid -> {pid, name}
  var pplSh = _sheet(TABS.PEOPLE);
  if (pplSh && pplSh.getLastRow() > 1) {
    var pdata = pplSh.getRange(2,1, pplSh.getLastRow()-1, Math.min(4,pplSh.getLastColumn())).getValues();
    pdata.forEach(function(r){
      var pid = String(r[0] || '');
      if (pid) people.set(pid, { pid: pid, name: String(r[1] || '') });
    });
  }

  var sh = _sheet(TABS.RESPONSES);
  if (!sh || sh.getLastRow() < 2) return { weekStart: weekStr, rows: [] };

  var data = sh.getDataRange().getValues();
  var header = data[0];
  var rows = data.slice(1);

  var idx = {
    person_id:  header.indexOf('person_id'),
    slot_id:    header.indexOf('slot_id'),
    status:     header.indexOf('status'),
    week_start: header.indexOf('week_start')
  };

  var byPerson = new Map(); // pid -> counts
  rows.forEach(function(r){
    if (!r || !r.length) return;
    var wk = (idx.week_start >= 0 && r[idx.week_start]) ? _ymd(r[idx.week_start])
            : (String(r[idx.slot_id]||'').split('|')[0]||'');
    if (wk !== weekStr) return;

    var pid = String(r[idx.person_id] || '');
    var st  = String(r[idx.status] || '');
    if (!pid) return;

    var base = byPerson.get(pid) || {
      pid: pid,
      name: (people.get(pid) && people.get(pid).name) || '',
      free: 0, class: 0, x: 0, total: 0, measure: 0
    };

    if (st === 'Free')  base.free++;
    if (st === 'Class') base.class++;
    if (st === 'x')     base.x++;
    if (st === 'Free' || st === 'Class' || st === 'x') base.total++;
    if (statusesToCount.indexOf(st) >= 0) base.measure++;

    byPerson.set(pid, base);
  });

  var out = Array.from(byPerson.values()).sort(function(a,b){ return a.name.localeCompare(b.name); });
  return { weekStart: weekStr, rows: out };
}

/***** SETTLEMENT (legacy week helpers) *****/
function settleWeekFromSettings() {
  var setSh = _sheet(TABS.SETTINGS);
  var min = Number(setSh.getRange('B2').getValue() || 0); // MinShifts
  return settleWeek(min);
}
function settleWeek(minShifts) {
  if (!(minShifts >= 0)) minShifts = 0;
  var summary = getWeeklySummary({ statusesToCount: ['Free'] }); // legacy behavior: only Free counts
  var weekStr = summary.weekStart;

  var ss = _ss();
  var name = 'SETTLEMENT';
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,10).setValues([[
      'week_start','person_id','name','free','class','x','total',
      'min_required','met_min','timestamp'
    ]]);
  }

  var rows = summary.rows.map(function(r){
    return [
      weekStr,
      r.pid, r.name, r.free, r.class, r.x, r.total,
      minShifts,
      r.measure >= minShifts,
      new Date()
    ];
  });

  if (rows.length) sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
  return { weekStart: weekStr, min: minShifts, written: rows.length };
}

/***** KẾT SỔ by period (only count 'x' and strict pass thresholds) *****/
/**
 * args: { type: 'week'|'month'|'term', start?: 'yyyy-mm-dd', end?: 'yyyy-mm-dd' }
 * Returns: { range:{start,end,label}, rows:[{pid,name,x,passed}] }
 * Rules: pass if x > 3 (week), > 15 (month), > 60 (term) — strictly greater
 */
function settlePeriod(args) {
  if (!args || !args.type) throw new Error('settlePeriod: missing args.type');
  var type  = String(args.type);
  var start = args.start ? new Date(args.start) : null;
  var end   = args.end   ? new Date(args.end)   : null;

  // If range not provided, derive from current selected week (SETTINGS!B1)
  if (!start || !end) {
    var monday = _getSelectedWeekDate();
    if (type === 'week') {
      start = new Date(monday);
      end   = _addDays(monday, 6);
    } else if (type === 'month') {
      start = new Date(monday.getFullYear(), monday.getMonth(), 1);
      end   = new Date(monday.getFullYear(), monday.getMonth()+1, 0);
    } else if (type === 'term') {
      var m = monday.getMonth(); var isH1 = (m < 6);
      start = new Date(monday.getFullYear(), isH1 ? 0 : 6, 1);
      end   = new Date(monday.getFullYear(), isH1 ? 6 : 12, 0);
    } else {
      throw new Error('settlePeriod: invalid type '+type);
    }
  }

  // Normalize inclusive range
  start = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0,0,0,0);
  end   = new Date(end.getFullYear(),   end.getMonth(),   end.getDate(),   23,59,59,999);

  // Thresholds: strictly greater than
  var THRESH = { week: 3, month: 15, term: 60 };
  if (!(type in THRESH)) throw new Error('settlePeriod: invalid type '+type);
  var thr = THRESH[type];

  // People map
  var pmap = {};
  var pplSh = _sheet(TABS.PEOPLE);
  if (pplSh && pplSh.getLastRow() > 1) {
    var pdata = pplSh.getRange(2,1, pplSh.getLastRow()-1, Math.min(4, pplSh.getLastColumn())).getValues();
    pdata.forEach(function(r){
      if (r && r[0] != null) pmap[String(r[0])] = String(r[1] || '');
    });
  }

  var respSh = _sheet(TABS.RESPONSES);
  var data = respSh.getDataRange().getValues();
  if (!data || data.length < 2) return { range: _labelRange_(type, start, end), rows: [] };

  var header = data[0], rows = data.slice(1);
  var idx = {
    person_id:  header.indexOf('person_id'),
    status:     header.indexOf('status'),
    week_start: header.indexOf('week_start'),
    date:       header.indexOf('date'),
    slot_id:    header.indexOf('slot_id')
  };
  if (idx.person_id < 0 || idx.status < 0) {
    throw new Error('RESPONSES missing person_id/status columns');
  }

  // Count only 'x' within [start, end]
  var xByPid = {}; // { pid: count }
  rows.forEach(function(r){
    if (!r || !r.length) return;
    var pid = String(r[idx.person_id] || '');
    var st  = String(r[idx.status]    || '');
    if (!pid || st !== 'x') return;

    var dCell = (idx.date >= 0 && r[idx.date]) ? new Date(r[idx.date])
              : (idx.week_start >= 0 && r[idx.week_start]) ? new Date(r[idx.week_start])
              : (idx.slot_id >= 0 && r[idx.slot_id]) ? (function() {
                    var iso = String(r[idx.slot_id]).split('|')[0]||'';
                    return iso ? new Date(iso) : null;
                })()
              : null;
    if (!dCell) return;
    var dt = new Date(dCell);
    if (dt < start || dt > end) return;

    xByPid[pid] = (xByPid[pid] || 0) + 1;
  });

  // Build output rows (passed = x > threshold)
  var out = Object.keys(xByPid).map(function(pid){
    var x = xByPid[pid] || 0;
    return { pid: pid, name: pmap[pid] || '', x: x, passed: (x > thr) };
  });

  // Sort: x desc, then name asc
  out.sort(function(a,b){ return (b.x - a.x) || String(a.name).localeCompare(String(b.name)); });

  return { range: _labelRange_(type, start, end), rows: out };

  function _labelRange_(t, s, e){
    var tz = Session.getScriptTimeZone();
    var fmt = function(d){ return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
    var label;
    if (t === 'week')       label = 'Tuần ' + fmt(s) + ' → ' + fmt(e);
    else if (t === 'month') label = 'Tháng ' + (s.getMonth()+1) + '/' + s.getFullYear();
    else                    label = 'Kỳ ' + (s.getMonth() < 6 ? '1' : '2') + ' / ' + s.getFullYear();
    return { start: fmt(s), end: fmt(e), label: label };
  }
}
