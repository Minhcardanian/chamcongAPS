/** @OnlyCurrentDoc
 * CONFIG / CONSTANTS
 * Timezone is important for weekly deadline rules
 */
const TZ = 'Asia/Ho_Chi_Minh';

// === ADMIN & AUTH ===
const ADMINS = new Set([
  '10423075', // Bùi Quang Minh
  '10622006', // Lê Ngọc Ánh
]);

// === SHEET NAMES ===
const SHEET_USERS = 'Users';           // user_id, name, email, role(active)
const SHEET_RESPONSES = 'Responses';   // date, user_id, status, subtype, updated_by, updated_ts
const SHEET_AUDIT = 'AuditLog';        // ts, actor_id, user_id, date, prev_status, prev_subtype, new_status, new_subtype
const SHEET_SETTINGS = 'Settings';     // key, value

// === STATUS / SUBTYPES ===
const STATUS_REGISTERED = 'REGISTERED';
const STATUS_BUSY = 'BUSY';
const STATUS_OTHER = 'OTHER';  // anything not REGISTERED/BUSY becomes OTHER
const STATUS_VALUES = [STATUS_REGISTERED, STATUS_BUSY, STATUS_OTHER];

const BUSY_SUBTYPES = ['EXAM', 'HOME', 'SICK', 'IN-ROOM'];

// === COLORS (for UI hints) ===
const COLOR_REGISTERED = '#16a34a'; // green-600
const COLOR_BUSY = '#dc2626';       // red-600
const COLOR_BUSY_INROOM = '#eab308';// yellow-500
const COLOR_OTHER = '#9ca3af';      // gray-400

// === MONTHLY RULE ===
const MONTHLY_MIN_SHIFTS = 8;

// === DEADLINE & EDIT MODE ===
// Sunday 23:59 local time to lock the NEXT WEEK plan for members.
// Admin can toggle override on/off (ignores deadlines).
const PROP_ADMIN_EDIT_OVERRIDE = 'ADMIN_EDIT_OVERRIDE'; // "true"/"false"

// === HELPERS ===
function getSS() { return SpreadsheetApp.getActive(); }

function getOrCreateSheet_(name) {
  const ss = getSS();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function isoDate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function now_() {
  return new Date();
}

function parseISO_(iso) {
  // Expect 'yyyy-MM-dd'
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function isAdminId_(userId) {
  return ADMINS.has(String(userId));
}

function getProperties_() {
  return PropertiesService.getScriptProperties();
}

function setAdminOverride_(val) {
  getProperties_().setProperty(PROP_ADMIN_EDIT_OVERRIDE, val ? 'true' : 'false');
}

function isAdminOverride_() {
  return getProperties_().getProperty(PROP_ADMIN_EDIT_OVERRIDE) === 'true';
}

