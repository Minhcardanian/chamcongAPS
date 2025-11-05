/** Core business logic for shift registration (single cell per day) */

function bootstrapSheets() {
  // Build (or reset) schema as requested => start fresh data
  const ss = getSS();

  // USERS
  {
    const sh = getOrCreateSheet_(SHEET_USERS);
    sh.clear();
    sh.getRange(1,1,1,4).setValues([['user_id','name','email','role']]);
    // NOTE: Leave population to admin (no demo users for privacy)
  }

  // RESPONSES
  {
    const sh = getOrCreateSheet_(SHEET_RESPONSES);
    sh.clear();
    sh.getRange(1,1,1,6)
      .setValues([['date','user_id','status','subtype','updated_by','updated_ts']]);
  }

  // AUDIT
  {
    const sh = getOrCreateSheet_(SHEET_AUDIT);
    sh.clear();
    sh.getRange(1,1,1,8).setValues([[
      'ts','actor_id','user_id','date','prev_status','prev_subtype','new_status','new_subtype'
    ]]);
  }

  // SETTINGS
  {
    const sh = getOrCreateSheet_(SHEET_SETTINGS);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,2).setValues([['key','value']]);
    }
  }

  // Turn OFF admin override by default
  setAdminOverride_(false);
}

/** Read from Users tab (if any) */
function listUsersFromUsers_() {
  const sh = getOrCreateSheet_(SHEET_USERS);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  if (!header.length || rows.length === 0) return [];
  const idx = Object.fromEntries(header.map((h,i)=>[String(h).toLowerCase(),i]));
  return rows
    .filter(r => r[idx['user_id']] )
    .map(r => ({
      user_id: String(r[idx['user_id']]),
      name: r[idx['name']] || '',
      email: r[idx['email']] || '',
      role: r[idx['role']] || 'active'
    }));
}

/** Fallback: read from PEOPLE tab when Users is empty */
function listUsersFromPeople_() {
  const ss = getSS();
  const sh = ss.getSheetByName('PEOPLE');
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  if (!header.length || rows.length === 0) return [];

  const map = {};
  header.forEach((h,i) => map[String(h).toLowerCase()] = i);

  function col() {
    for (let k=0; k<arguments.length; k++) {
      const key = String(arguments[k]).toLowerCase();
      if (map.hasOwnProperty(key)) return map[key];
    }
    return -1;
  }

  const idCol   = col('person_id','user_id','id');
  const nameCol = col('pname','name','full_name','fullname');
  const mailCol = col('email','mail');

  if (idCol < 0 || nameCol < 0) return [];

  const out = [];
  rows.forEach(r => {
    const uid = r[idCol];
    const nm  = r[nameCol];
    if (!uid || !nm) return;
    out.push({
      user_id: String(uid),
      name: String(nm),
      email: mailCol >= 0 ? String(r[mailCol] || '') : '',
      role: 'active'
    });
  });
  return out;
}

/** Combined list; prefer Users, fallback to PEOPLE */
function listUsers_() {
  let users = [];
  try { users = listUsersFromUsers_(); } catch (e) {}
  if (users && users.length) return users;
  return listUsersFromPeople_();
}

function loadWeekMatrix(weekStartISO) {
  // Returns: { weekStartISO, days: [iso...], users:[{user_id,name}], data: {user_id:{iso:{status,subtype}}}, colors:{...} }
  const users = listUsers_();
  const start = parseISO_(weekStartISO);
  const days = Array.from({length:7}, (_,i)=> {
    const d = new Date(start.getTime());
    d.setDate(d.getDate()+i);
    return isoDate_(d);
  });

  const data = {};
  users.forEach(u => { data[u.user_id] = {}; });

  // Pull responses for those dates
  const sh = getOrCreateSheet_(SHEET_RESPONSES);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  const idx = Object.fromEntries(header.map((h,i)=>[h,i]));

  const daySet = new Set(days);
  const uidSet = new Set(users.map(u=>u.user_id));

  rows.forEach(r => {
    const date = r[idx.date];
    const uid = String(r[idx.user_id] || '');
    if (!daySet.has(date) || !uidSet.has(uid)) return;
    const status = (r[idx.status] || '').toString().toUpperCase();
    const subtype = (r[idx.subtype] || '').toString().toUpperCase();
    data[uid][date] = normalizeStatus_(status, subtype);
  });

  // Fill blanks (must not be blank → default OTHER)
  users.forEach(u=>{
    days.forEach(d=>{
      if (!data[u.user_id][d]) data[u.user_id][d] = {status: STATUS_OTHER, subtype: ''};
    });
  });

  return {
    weekStartISO,
    days,
    users: users.map(({user_id, name})=>({user_id, name})),
    data,
    colors: {
      REGISTERED: COLOR_REGISTERED,
      BUSY: COLOR_BUSY,
      BUSY_INROOM: COLOR_BUSY_INROOM,
      OTHER: COLOR_OTHER
    }
  };
}

function normalizeStatus_(status, subtype) {
  let s = (status || '').toUpperCase();
  let sub = (subtype || '').toUpperCase();

  if (s === STATUS_REGISTERED) return {status: STATUS_REGISTERED, subtype: ''};
  if (s === STATUS_BUSY) {
    if (!BUSY_SUBTYPES.includes(sub)) sub = 'OTHER';
    return {status: STATUS_BUSY, subtype: sub};
  }
  // Anything else -> OTHER
  return {status: STATUS_OTHER, subtype: ''};
}

function saveStatus(weekStartISO, dateISO, targetUserId, status, subtype, actorId) {
  // Permission & lock checks are done in server router; here we just persist and audit.
  const normalized = normalizeStatus_(status, subtype);

  const sh = getOrCreateSheet_(SHEET_RESPONSES);
  const rows = sh.getDataRange().getValues();
  let header = rows.shift() || [];
  if (header.length === 0) {
    header = ['date','user_id','status','subtype','updated_by','updated_ts'];
    sh.getRange(1,1,1,header.length).setValues([header]);
  }
  const idx = Object.fromEntries(header.map((h,i)=>[h,i]));

  // Find existing row by (date,user_id)
  let foundRow = null;
  rows.forEach((r, i) => {
    if (String(r[idx.user_id]) === String(targetUserId) && r[idx.date] === dateISO) {
      foundRow = {arr: r, rowNumber: i+2}; // +2 because header offset
    }
  });

  const prev = foundRow ? {
    status: String(foundRow.arr[idx.status] || ''),
    subtype: String(foundRow.arr[idx.subtype] || '')
  } : {status: '', subtype: ''};

  const nowTs = Utilities.formatDate(now_(), TZ, "yyyy-MM-dd' 'HH:mm:ss");

  if (foundRow) {
    sh.getRange(foundRow.rowNumber, idx.status+1).setValue(normalized.status);
    sh.getRange(foundRow.rowNumber, idx.subtype+1).setValue(normalized.subtype);
    sh.getRange(foundRow.rowNumber, idx.updated_by+1).setValue(String(actorId));
    sh.getRange(foundRow.rowNumber, idx.updated_ts+1).setValue(nowTs);
  } else {
    sh.appendRow([dateISO, String(targetUserId), normalized.status, normalized.subtype, String(actorId), nowTs]);
  }

  // Audit trail
  const audit = getOrCreateSheet_(SHEET_AUDIT);
  audit.appendRow([
    nowTs, String(actorId), String(targetUserId), dateISO,
    prev.status, prev.subtype,
    normalized.status, normalized.subtype
  ]);

  return {ok: true};
}

function computeMonthlyCounts_(year, month /*1-12*/) {
  // returns { user_id: {count, byStatus:{REGISTERED:n, BUSY_INROOM:n, OTHER:n}} }
  const sh = getOrCreateSheet_(SHEET_RESPONSES);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  const idx = Object.fromEntries(header.map((h,i)=>[h,i]));
  const res = {};

  rows.forEach(r=>{
    const dateISO = r[idx.date];
    if (!dateISO) return;
    const d = parseISO_(dateISO);
    if ((d.getFullYear() !== year) || ((d.getMonth()+1) !== month)) return;

    const uid = String(r[idx.user_id] || '');
    const status = (r[idx.status] || '').toString().toUpperCase();
    const subtype = (r[idx.subtype] || '').toString().toUpperCase();

    if (!res[uid]) res[uid] = { count: 0, byStatus:{REGISTERED:0, BUSY_INROOM:0, OTHER:0} };

    if (status === STATUS_REGISTERED) {
      res[uid].count += 1;
      res[uid].byStatus.REGISTERED++;
    } else if (status === STATUS_BUSY && subtype === 'IN-ROOM') {
      res[uid].count += 1;
      res[uid].byStatus.BUSY_INROOM++;
    } else {
      res[uid].byStatus.OTHER++;
    }
  });

  return res;
}

function emailMonthlyReminders(year, month) {
  // If omitted, default to previous month (run on 1st)
  const now = now_();
  if (!year || !month) {
    const d = new Date(now.getTime());
    d.setMonth(d.getMonth() - 1);
    year = d.getFullYear();
    month = d.getMonth() + 1;
  }

  const users = listUsers_();
  const counts = computeMonthlyCounts_(year, month);

  const behind = [];
  users.forEach(u=>{
    const c = counts[u.user_id]?.count || 0;
    if (c < MONTHLY_MIN_SHIFTS) behind.push({user: u, count: c});
  });

  // Helper: email fallback builder
  const emailOf = (u) => {
    const mail = (u && u.email) ? String(u.email).trim() : '';
    if (mail && mail.indexOf('@') > -1) return mail;
    const pid = (u && u.user_id) ? String(u.user_id).trim() : '';
    return pid ? `${pid}@student.vgu.edu.vn` : '';
  };

  if (behind.length === 0) {
    // Send a single summary email to admins
    ADMINS.forEach(adminId=>{
      const admin = users.find(u=>u.user_id === adminId);
      const to = emailOf(admin) || Session.getActiveUser().getEmail();
      if (to) {
        MailApp.sendEmail({
          to,
          subject: `[ChamCong] Monthly summary ${year}-${(''+month).padStart(2,'0')}`,
          htmlBody: `<p>Congratulation no lates this month</p>`
        });
      }
    });
    return {ok:true, message:'All passed. Admins notified.'};
  }

  // Notify each who is behind
  behind.forEach(({user, count})=>{
    const diff = MONTHLY_MIN_SHIFTS - count;
    const to = emailOf(user);
    if (!to) return;
    MailApp.sendEmail({
      to,
      subject: `[ChamCong] Monthly reminder ${year}-${(''+month).padStart(2,'0')}`,
      htmlBody: `<p>Hi ${user.name},</p>
<p>You currently have <b>${count}</b> shifts, which is <b>${diff}</b> behind the standard minimum of <b>${MONTHLY_MIN_SHIFTS}</b> this month.</p>
<p>Please coordinate with admins for upcoming schedules.</p>`
    });
  });

  return {ok:true, message:`Notified ${behind.length} users`};
}

// === WEEK/DEADLINE ===

// Return Monday of the week containing given date
function weekStartMonday_(d) {
  const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const day = dd.getDay(); // 0=Sun..6=Sat
  const diff = (day === 0 ? -6 : (1 - day)); // shift to Monday
  dd.setDate(dd.getDate() + diff);
  return dd;
}

// Next week's Monday (relative to "now")
function nextWeekMonday_() {
  const n = now_();
  const thisMonday = weekStartMonday_(n);
  const nextMon = new Date(thisMonday.getTime());
  nextMon.setDate(thisMonday.getDate()+7);
  return nextMon;
}

// Deadline is Sunday 23:59 local, right before nextMon
function nextWeekDeadline_() {
  const nextMon = nextWeekMonday_();
  const dl = new Date(nextMon.getTime());
  dl.setMinutes(-1); // 23:59 of the previous day (Sunday)
  return dl;
}

function isEditableForMember_(targetDateISO) {
  // Members can only edit the *next week* grid until Sunday 23:59
  if (isAdminOverride_()) return true; // admin override (and, per server, also everyone-can-edit)
  const now = now_();
  const target = parseISO_(targetDateISO);
  const targetWeekStart = weekStartMonday_(target);
  const nextMon = nextWeekMonday_();
  if (targetWeekStart.getTime() !== nextMon.getTime()) return false; // only next week editable
  const deadline = nextWeekDeadline_();
  return now.getTime() <= deadline.getTime();
}
