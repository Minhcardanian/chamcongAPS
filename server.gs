/** ====================================================================
 *  server.gs — HTTP endpoints, router, auth, and action handlers
 *  HtmlService client MUST call: google.script.run.route({ action, ... })
 *  External callers (webhooks) may POST JSON to doPost.
 *
 *  Auth options:
 *    A) Classic login: user_key = pname (or person_id), password = person_id
 *    B) Google login: loginWithGoogle() — email must match /^\d+@student.vgu.edu.vn$/
 *       and person_id must exist in PEOPLE/Users
 *
 *  Modes:
 *    - Admin Edit Mode (ADMIN_EDIT_OVERRIDE): if ON, everyone can edit anyone and deadline ignored
 *    - Matrix Read-Only (MATRIX_READ_ONLY): if ON, members cannot edit; admins can
 * ==================================================================== */

/** Internal property keys (uses getProperties_() from config.js) */
const PROP_MATRIX_RO = 'MATRIX_READ_ONLY';

/** Serve the web UI */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ChamCong / Shift Registration');
}

/** HtmlService-safe entrypoint (use this for google.script.run calls) */
function route(payload) {
  try {
    return dispatch_(payload || {});
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/** Optional external HTTP endpoint (JSON POST). Not used by the UI by default. */
function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    const out = dispatch_(payload || {});
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: String((err && err.message) || err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/** Central dispatcher for all actions */
function dispatch_(payload) {
  const action = payload && payload.action;
  switch (action) {
    case 'login':
      return { ok: true, data: login_(payload.user_key, payload.password) };

    case 'loginWithGoogle':
      return { ok: true, data: loginWithGoogle_() };

    case 'getWeek':
      return { ok: true, data: handleGetWeek_(payload) };

    case 'setStatus':
      return { ok: true, data: handleSetStatus_(payload) };

    case 'setStatusBatch':
      return { ok: true, data: handleSetStatusBatch_(payload) };

    case 'toggleAdminEdit':
      return { ok: true, data: handleToggleAdminEdit_(payload) };

    case 'setMatrixRO':
      return { ok: true, data: handleSetMatrixRO_(payload) };

    case 'sendMonthlyReminders':
      return { ok: true, data: emailMonthlyReminders(payload.year, payload.month) };

    case 'getAudit':
      return { ok: true, data: handleGetAudit_(payload) };

    default:
      throw new Error('Unknown action: ' + action);
  }
}

/* ========================= AUTH ============================== */

function _norm(s) {
  return String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _listFromPeople_() {
  const ss = getSS();
  const sh = ss.getSheetByName('PEOPLE');
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  const map = {};
  header.forEach((h, i) => map[String(h).toLowerCase()] = i);

  function col() {
    for (let k = 0; k < arguments.length; k++) {
      const idx = map[String(arguments[k]).toLowerCase()];
      if (idx !== undefined) return idx;
    }
    return -1;
  }

  const idCol   = col('person_id', 'user_id', 'id');
  const nameCol = col('pname', 'name', 'full_name', 'fullname');
  const mailCol = col('email', 'mail');
  if (idCol < 0 || nameCol < 0) return [];

  const out = [];
  rows.forEach(r => {
    const uid = r[idCol], nm = r[nameCol];
    if (!uid || !nm) return;
    out.push({
      user_id: String(uid),
      name: String(nm),
      email: mailCol >= 0 ? String(r[mailCol] || '') : ''
    });
  });
  return out;
}

function _listUsersCombined_() {
  let users = [];
  try { users = listUsers_(); } catch (e) {}
  if (users && users.length) return users;
  return _listFromPeople_();
}

function _withCommonSession_(me) {
  return {
    user: { user_id: me.user_id, name: me.name, email: me.email || '' },
    is_admin: isAdminId_(me.user_id),
    admin_edit_override: isAdminOverride_(),
    read_only: isMatrixRO_()
  };
}

/** Classic login: user_key=pname (or person_id), password=person_id */
function login_(user_key, password) {
  if (!user_key || !password) throw new Error('Missing credentials');

  const users = _listUsersCombined_();
  if (!users.length) throw new Error('No users found. Please set up Users or PEOPLE sheet.');

  const keyNorm = _norm(user_key);
  let me = users.find(u => _norm(u.name) === keyNorm);
  if (!me) me = users.find(u => _norm(u.user_id) === keyNorm);
  if (!me) throw new Error('Invalid credentials');

  // password = person_id by default, with optional Settings override
  let passOk = (String(password) === String(me.user_id));
  if (!passOk) {
    const sh = getOrCreateSheet_(SHEET_SETTINGS);
    const rows = sh.getDataRange().getValues();
    const header = rows.shift() || [];
    const idx = Object.fromEntries(header.map((h, i) => [h, i]));
    let stored = null;
    rows.forEach(r => { if (r[idx.key] === `pwd:${me.user_id}`) stored = r[idx.value]; });
    passOk = !!stored && String(stored) === String(password);
  }
  if (!passOk) throw new Error('Invalid credentials');

  return _withCommonSession_(me);
}

/** Google login: email must be like 10423075@student.vgu.edu.vn and exist as person_id */
function loginWithGoogle_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw new Error('No Google identity. Ask admin to deploy as "User accessing the web app".');

  const m = email.match(/^(\d+)@student\.vgu\.edu\.vn$/);
  if (!m) throw new Error('Email not allowed (must be person_id@student.vgu.edu.vn).');
  const personId = m[1];

  const users = _listUsersCombined_();
  const me = users.find(u => String(u.user_id) === personId);
  if (!me) throw new Error('Your person_id is not in the PEOPLE/Users list.');

  return _withCommonSession_(me);
}

/* ====================== ACTION HANDLERS ======================= */

function handleGetWeek_(p) {
  const weekStartISO = p && p.weekStartISO ? p.weekStartISO : isoDate_(weekStartMonday_(now_()));
  const matrix = loadWeekMatrix(weekStartISO);
  // include flags so client reflects accurate state
  return Object.assign({}, matrix, {
    admin_edit_override: isAdminOverride_(),
    read_only: isMatrixRO_()
  });
}

/** Centralized permission with Admin Edit Mode + Matrix Read-Only */
function _enforceEditPermission_(actor_id, target_user_id, dateISO) {
  const actorIsAdmin = isAdminId_(String(actor_id));

  // Hard lock: Matrix Read-Only blocks members, but not admins
  if (isMatrixRO_() && !actorIsAdmin) {
    throw new Error('Matrix is in read-only mode.');
  }

  // If Admin Edit Mode is ON: everyone can edit anyone, and deadline ignored
  if (isAdminOverride_()) return;

  if (!actorIsAdmin && String(actor_id) !== String(target_user_id)) {
    throw new Error('Not allowed to edit others.');
  }
  if (!actorIsAdmin && !isEditableForMember_(dateISO)) {
    throw new Error('Registration for next week is locked (passed Sunday 23:59).');
  }
}

function handleSetStatus_(p) {
  const { actor_id, target_user_id, dateISO, status, subtype } = p || {};
  if (!actor_id || !target_user_id || !dateISO) throw new Error('Missing fields');

  _enforceEditPermission_(actor_id, target_user_id, dateISO);

  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    lock.waitLock(5000); // throws if not acquired in time
    acquired = true;
    return saveStatus('', dateISO, String(target_user_id), String(status), String(subtype || ''), String(actor_id));
  } catch (e) {
    const msg = String(e || '');
    if (msg.includes('Service invoked too many times') || msg.includes('Could not acquire')) {
      throw new Error('System is busy. Please try again in a few seconds.');
    }
    throw e;
  } finally {
    if (acquired) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}

/** Batch save: { actor_id, changes:[ { target_user_id, dateISO, status, subtype } ] } */
function handleSetStatusBatch_(p) {
  const { actor_id } = p || {};
  const changes = ((p && p.changes) || []).map(c => ({
    target_user_id: String(c.target_user_id),
    dateISO: String(c.dateISO),
    status: String(c.status || ''),
    subtype: String(c.subtype || '')
  }));

  if (!actor_id) throw new Error('Missing actor_id');
  if (!changes.length) return { saved: 0 };

  // Permission pre-check for each change
  changes.forEach(c => _enforceEditPermission_(actor_id, c.target_user_id, c.dateISO));

  const lock = LockService.getScriptLock();
  let acquired = false, saved = 0;
  try {
    lock.waitLock(30000); // throws if not acquired
    acquired = true;

    changes.forEach(c => {
      saveStatus('', c.dateISO, c.target_user_id, c.status, c.subtype, String(actor_id));
      saved++;
    });

    return { saved };
  } catch (e) {
    const msg = String(e || '');
    if (!saved && (msg.includes('Service invoked too many times') || msg.includes('Could not acquire'))) {
      throw new Error('System is busy. Please try again shortly.');
    }
    throw e;
  } finally {
    if (acquired) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}

function handleToggleAdminEdit_(p) {
  const { actor_id, on } = p || {};
  if (!isAdminId_(String(actor_id))) throw new Error('Admin only');
  setAdminOverride_(!!on);
  return { admin_edit_override: isAdminOverride_() };
}

/** Matrix Read-Only handlers */
function isMatrixRO_() {
  return getProperties_().getProperty(PROP_MATRIX_RO) === 'true';
}

function setMatrixRO_(val) {
  getProperties_().setProperty(PROP_MATRIX_RO, val ? 'true' : 'false');
}

function handleSetMatrixRO_(p) {
  const { actor_id, on } = p || {};
  if (!isAdminId_(String(actor_id))) throw new Error('Admin only');
  setMatrixRO_(!!on);
  return { read_only: isMatrixRO_() };
}

function handleGetAudit_(p) {
  const sh = getOrCreateSheet_(SHEET_AUDIT);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  const idx = Object.fromEntries(header.map((h, i) => [h, i]));
  const out = rows.reverse().slice(0, 200).map(r => ({
    ts: r[idx.ts],
    actor_id: r[idx.actor_id],
    user_id: r[idx.user_id],
    date: r[idx.date],
    prev_status: r[idx.prev_status],
    prev_subtype: r[idx.prev_subtype],
    new_status: r[idx.new_status],
    new_subtype: r[idx.new_subtype],
  }));
  return { items: out };
}

/* ===================== ADMIN HELPERS ========================= */

function adminSetPassword(user_id, password) {
  const sh = getOrCreateSheet_(SHEET_SETTINGS);
  const range = sh.getDataRange();
  const values = range.getValues();
  let found = false;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === `pwd:${user_id}`) {
      sh.getRange(i + 1, 2).setValue(password);
      found = true; break;
    }
  }
  if (!found) sh.appendRow([`pwd:${user_id}`, password]);
}

function installMonthlyReminderTrigger() {
  ScriptApp.newTrigger('cronMonthlyReminder_')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .inTimezone(TZ)
    .create();
}

function cronMonthlyReminder_() {
  emailMonthlyReminders();
}
