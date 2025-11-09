/** Core business logic for shift registration (single cell per day) */

/* ====================== BOOTSTRAP / SCHEMA ====================== */

function bootstrapSheets() {
  const ss = getSS();

  // USERS
  {
    const sh = getOrCreateSheet_(SHEET_USERS);
    sh.clear();
    sh.getRange(1, 1, 1, 4).setValues([['user_id', 'name', 'email', 'role']]);
  }

  // RESPONSES — canonical schema
  {
    const sh = getOrCreateSheet_(SHEET_RESPONSES);
    sh.clear();
    sh.getRange(1, 1, 1, 6)
      .setValues([['date', 'user_id', 'status', 'subtype', 'updated_by', 'updated_ts']]);
    // Ensure "date" (col A) stays plain text for many future rows
    sh.getRange(2, 1, Math.max(10000, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  }

  // AUDIT
  {
    const sh = getOrCreateSheet_(SHEET_AUDIT);
    sh.clear();
    sh.getRange(1, 1, 1, 8).setValues([[
      'ts', 'actor_id', 'user_id', 'date', 'prev_status', 'prev_subtype', 'new_status', 'new_subtype'
    ]]);
  }

  // SETTINGS
  {
    const sh = getOrCreateSheet_(SHEET_SETTINGS);
    sh.clear();
    sh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }

  // Turn OFF admin override by default
  setAdminOverride_(false);
}

/* ====================== USERS ====================== */

function listUsersFromUsers_() {
  const sh = getOrCreateSheet_(SHEET_USERS);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  if (!header.length || rows.length === 0) return [];
  const idx = Object.fromEntries(header.map((h, i) => [String(h).toLowerCase(), i]));
  return rows
    .filter(r => r[idx['user_id']])
    .map(r => ({
      user_id: String(r[idx['user_id']]),
      name: r[idx['name']] || '',
      email: r[idx['email']] || '',
      role: r[idx['role']] || 'active'
    }));
}

function listUsersFromPeople_() {
  const ss = getSS();
  const sh = ss.getSheetByName('PEOPLE');
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];
  if (!header.length || rows.length === 0) return [];

  const map = {};
  header.forEach((h, i) => (map[String(h).toLowerCase()] = i));
  const col = (...names) => {
    for (let n of names) {
      const k = String(n).toLowerCase();
      if (map.hasOwnProperty(k)) return map[k];
    }
    return -1;
  };

  const idCol = col('person_id', 'user_id', 'id');
  const nameCol = col('pname', 'name', 'full_name', 'fullname');
  const mailCol = col('email', 'mail');

  if (idCol < 0 || nameCol < 0) return [];

  const out = [];
  rows.forEach(r => {
    const uid = r[idCol];
    const nm = r[nameCol];
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

function listUsers_() {
  try {
    const u = listUsersFromUsers_();
    if (u && u.length) return u;
  } catch (e) {}
  return listUsersFromPeople_();
}

/* ====================== MATRIX LOAD ====================== */

function loadWeekMatrix(weekStartISO) {
  const users = listUsers_();
  const start = parseISO_(weekStartISO);
  const days = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(start.getTime());
    d.setDate(d.getDate() + i);
    return isoDate_(d);
  });

  const data = {};
  users.forEach(u => (data[u.user_id] = {}));

  const sh = getOrCreateSheet_(SHEET_RESPONSES);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];

  // Case-insensitive, alias-friendly header map
  const hmap = {};
  header.forEach((h, i) => (hmap[String(h).trim().toLowerCase()] = i));
  const col = (...names) => {
    for (let n of names) {
      const k = String(n).trim().toLowerCase();
      if (hmap.hasOwnProperty(k)) return hmap[k];
    }
    return -1;
  };
  const cDate = col('date', 'ngay', 'day');
  const cUserId = col('user_id', 'userid', 'person_id', 'student_id', 'uid', 'id');
  const cStatus = col('status', 'trangthai');
  const cSubtype = col('subtype', 'sub_type', 'reason', 'note');

  const daySet = new Set(days);
  const uidSet = new Set(users.map(u => u.user_id));

  rows.forEach(r => {
    const raw = r[cDate];
    const dateISO = raw instanceof Date ? isoDate_(raw) : String(raw);
    const uid = String(r[cUserId] || '');
    if (!daySet.has(dateISO) || !uidSet.has(uid)) return;

    const status = (r[cStatus] || '').toString().toUpperCase();
    const subtype = (r[cSubtype] || '').toString().toUpperCase();
    data[uid][dateISO] = normalizeStatus_(status, subtype);
  });

  // Fill blanks (default OTHER)
  users.forEach(u => {
    days.forEach(d => {
      if (!data[u.user_id][d]) data[u.user_id][d] = { status: STATUS_OTHER, subtype: '' };
    });
  });

  return {
    weekStartISO,
    days,
    users: users.map(({ user_id, name }) => ({ user_id, name })),
    data,
    colors: {
      REGISTERED: COLOR_REGISTERED,
      BUSY: COLOR_BUSY,
      BUSY_INROOM: COLOR_BUSY_INROOM,
      OTHER: COLOR_OTHER
    }
  };
}

/* ====================== SAVE / NORMALIZE ====================== */

function normalizeStatus_(status, subtype) {
  let s = (status || '').toUpperCase();
  let sub = (subtype || '').toUpperCase();

  if (s === STATUS_REGISTERED) return { status: STATUS_REGISTERED, subtype: '' };
  if (s === STATUS_BUSY) {
    if (!BUSY_SUBTYPES.includes(sub)) sub = 'OTHER';
    return { status: STATUS_BUSY, subtype: sub };
  }
  return { status: STATUS_OTHER, subtype: '' };
}

function saveStatus(weekStartISO, dateISO, targetUserId, status, subtype, actorId) {
  const normalized = normalizeStatus_(status, subtype);

  const sh = getOrCreateSheet_(SHEET_RESPONSES);
  const rows = sh.getDataRange().getValues();
  let header = rows.shift() || [];
  if (header.length === 0) {
    header = ['date', 'user_id', 'status', 'subtype', 'updated_by', 'updated_ts'];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }

  // case-insensitive header map + aliases
  const hmap = {};
  header.forEach((h, i) => (hmap[String(h).trim().toLowerCase()] = i));
  const col = (...names) => {
    for (let n of names) {
      const k = String(n).trim().toLowerCase();
      if (hmap.hasOwnProperty(k)) return hmap[k];
    }
    return -1;
  };
  const cDate = col('date', 'ngay', 'day');
  const cUserId = col('user_id', 'userid', 'person_id', 'student_id', 'uid', 'id');
  const cStatus = col('status', 'trangthai');
  const cSubtype = col('subtype', 'sub_type', 'reason', 'note');
  const cBy = col('updated_by', 'by', 'actor', 'updatedby');
  const cTs = col('updated_ts', 'updated_at', 'timestamp', 'ts');

  const nowTs = Utilities.formatDate(now_(), TZ, "yyyy-MM-dd' 'HH:mm:ss");

  // Find existing row by (date,user_id)
  let foundRow = null;
  rows.forEach((r, i) => {
    const raw = r[cDate];
    const dISO = raw instanceof Date ? isoDate_(raw) : String(raw);
    const uid = String(r[cUserId] || '');
    if (uid === String(targetUserId) && dISO === dateISO) {
      foundRow = { rowNumber: i + 2, arr: r };
    }
  });

  if (foundRow) {
    if (cStatus >= 0) sh.getRange(foundRow.rowNumber, cStatus + 1).setValue(normalized.status);
    if (cSubtype >= 0) sh.getRange(foundRow.rowNumber, cSubtype + 1).setValue(normalized.subtype);
    if (cBy >= 0) sh.getRange(foundRow.rowNumber, cBy + 1).setValue(String(actorId));
    if (cTs >= 0) sh.getRange(foundRow.rowNumber, cTs + 1).setValue(nowTs);
  } else {
    // Compose row in detected order
    const maxIdx = Math.max(cDate, cUserId, cStatus, cSubtype, cBy, cTs);
    const row = new Array(maxIdx + 1).fill('');
    if (cDate >= 0) row[cDate] = dateISO;
    if (cUserId >= 0) row[cUserId] = String(targetUserId);
    if (cStatus >= 0) row[cStatus] = normalized.status;
    if (cSubtype >= 0) row[cSubtype] = normalized.subtype;
    if (cBy >= 0) row[cBy] = String(actorId);
    if (cTs >= 0) row[cTs] = nowTs;
    sh.appendRow(row);
    // Keep the appended date cell as plain text
    if (cDate >= 0) sh.getRange(sh.getLastRow(), cDate + 1).setNumberFormat('@');
  }

  // Audit trail
  const audit = getOrCreateSheet_(SHEET_AUDIT);
  audit.appendRow([
    nowTs,
    String(actorId),
    String(targetUserId),
    dateISO,
    '', // prev_status (not tracked in this simplified write path)
    '',
    normalized.status,
    normalized.subtype
  ]);

  return { ok: true };
}

/* ====================== MONTHLY COUNTS & EMAIL ====================== */

function computeMonthlyCounts_(year, month /*1-12*/) {
  const sh = getOrCreateSheet_(SHEET_RESPONSES);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift() || [];

  const hmap = {};
  header.forEach((h, i) => (hmap[String(h).trim().toLowerCase()] = i));
  const col = (...names) => {
    for (let n of names) {
      const k = String(n).trim().toLowerCase();
      if (hmap.hasOwnProperty(k)) return hmap[k];
    }
    return -1;
  };
  const cDate = col('date', 'ngay', 'day');
  const cUserId = col('user_id', 'userid', 'person_id', 'student_id', 'uid', 'id');
  const cStatus = col('status', 'trangthai');
  const cSubtype = col('subtype', 'sub_type', 'reason', 'note');

  const res = {};
  rows.forEach(r => {
    const raw = r[cDate];
    const dateISO = raw instanceof Date ? isoDate_(raw) : String(raw);
    if (!dateISO) return;
    const d = parseISO_(dateISO);
    if (d.getFullYear() !== year || d.getMonth() + 1 !== month) return;

    const uid = String(r[cUserId] || '');
    const status = (r[cStatus] || '').toString().toUpperCase();
    const subtype = (r[cSubtype] || '').toString().toUpperCase();

    if (!res[uid]) res[uid] = { count: 0, byStatus: { REGISTERED: 0, BUSY_INROOM: 0, OTHER: 0 } };
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
  users.forEach(u => {
    const c = counts[u.user_id]?.count || 0;
    if (c < MONTHLY_MIN_SHIFTS) behind.push({ user: u, count: c });
  });

  const emailOf = u => {
    const mail = (u && u.email) ? String(u.email).trim() : '';
    if (mail && mail.indexOf('@') > -1) return mail;
    const pid = (u && u.user_id) ? String(u.user_id).trim() : '';
    return pid ? `${pid}@student.vgu.edu.vn` : '';
  };

  if (behind.length === 0) {
    ADMINS.forEach(adminId => {
      const admin = users.find(u => u.user_id === adminId);
      const to = emailOf(admin) || Session.getActiveUser().getEmail();
      if (to) {
        MailApp.sendEmail({
          to,
          subject: `[ChamCong] Monthly summary ${year}-${('' + month).padStart(2, '0')}`,
          htmlBody: `<p>Congratulation no lates this month</p>`
        });
      }
    });
    return { ok: true, message: 'All passed. Admins notified.' };
  }

  behind.forEach(({ user, count }) => {
    const diff = MONTHLY_MIN_SHIFTS - count;
    const to = emailOf(user);
    if (!to) return;
    MailApp.sendEmail({
      to,
      subject: `[ChamCong] Monthly reminder ${year}-${('' + month).padStart(2, '0')}`,
      htmlBody: `<p>Hi ${user.name},</p>
<p>You currently have <b>${count}</b> shifts, which is <b>${diff}</b> behind the standard minimum of <b>${MONTHLY_MIN_SHIFTS}</b> this month.</p>
<p>Please coordinate with admins for upcoming schedules.</p>`
    });
  });

  return { ok: true, message: `Notified ${behind.length} users` };
}

/* ====================== WEEK/DEADLINE HELPERS ====================== */

function weekStartMonday_(d) {
  const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const day = dd.getDay(); // 0=Sun..6=Sat
  const diff = (day === 0 ? -6 : (1 - day));
  dd.setDate(dd.getDate() + diff);
  return dd;
}

function nextWeekMonday_() {
  const n = now_();
  const thisMonday = weekStartMonday_(n);
  const nextMon = new Date(thisMonday.getTime());
  nextMon.setDate(thisMonday.getDate() + 7);
  return nextMon;
}

function nextWeekDeadline_() {
  const nextMon = nextWeekMonday_();
  const dl = new Date(nextMon.getTime());
  dl.setMinutes(-1); // 23:59 Sunday
  return dl;
}

function isEditableForMember_(targetDateISO) {
  if (isAdminOverride_()) return true; // everyone can edit; deadline ignored
  const now = now_();
  const target = parseISO_(targetDateISO);
  const targetWeekStart = weekStartMonday_(target);
  const nextMon = nextWeekMonday_();
  if (targetWeekStart.getTime() !== nextMon.getTime()) return false; // only next week editable
  const deadline = nextWeekDeadline_();
  return now.getTime() <= deadline.getTime();
}
