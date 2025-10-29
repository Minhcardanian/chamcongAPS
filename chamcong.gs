/** ===== Global config (single source of truth) ===== */
const TABS = {
  PEOPLE: 'PEOPLE',                // fixed roster
  SLOTS: 'SLOTS',                  // slot_id | week_start | day | block
  RESPONSES: 'RESPONSES',          // respond_id | person_id | slot_id | status | timestamp | week_start | date
  SETTINGS: 'SETTINGS',            // A1: SelectedWeek, B1: <Monday date>
  ARCHIVE: 'RESPONSES_Archive',    // same headers as RESPONSES
};
const SETTINGS_RANGE = 'B1';       // SelectedWeek (a Monday)
const DAYS   = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
const BLOCKS = ['9-12','13-16','19-21'];
const STATUS_VALUES = ['', 'Free', 'Class', 'x']; // cycle order & UI legend

/** ===== Helpers (shared) ===== */
function getSelectedWeek_() {
  const ws = SpreadsheetApp.getActive().getSheetByName(TABS.SETTINGS).getRange(SETTINGS_RANGE).getValue();
  if (!ws) throw new Error(`${TABS.SETTINGS}!${SETTINGS_RANGE} is empty (SelectedWeek). Put a Monday date.`);
  return new Date(ws);
}
function ymd_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** ===== Spreadsheet menu ===== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Chấm công')
    .addItem('1) Tạo 21 slot cho tuần ở SETTINGS!B1', 'seedWeekSlots')
    .addItem('2) Lưu trữ & reset tuần hiện tại', 'archiveAndResetWeek')
    .addSeparator()
    .addItem('3) Nhảy sang tuần kế & tạo slot', 'goToNextWeekAndSeed')
    .addToUi();
}

/** 1) Append 21 SLOTS for the selected week (skip existing) */
function seedWeekSlots() {
  const ss = SpreadsheetApp.getActive();
  const slots = ss.getSheetByName(TABS.SLOTS);
  const weekStart = getSelectedWeek_();

  const existing = new Set(
    (slots.getRange(2,1,Math.max(slots.getLastRow()-1,0),1).getValues()||[])
      .flat().filter(Boolean)
  );

  const wkStr = ymd_(weekStart);
  const rows = [];
  for (const block of BLOCKS) {
    for (const day of DAYS) {
      const id = `${wkStr}|${day}|${block}`;
      if (existing.has(id)) continue;
      rows.push([id, weekStart, day, block]); // slot_id, week_start, day, block
    }
  }
  if (!rows.length) {
    SpreadsheetApp.getUi().alert('No new slots to add (all 21 exist).');
    return;
  }
  slots.getRange(slots.getLastRow()+1, 1, rows.length, 4).setValues(rows);
  SpreadsheetApp.getUi().alert(`Added ${rows.length} slot(s) for week ${wkStr}.`);
}

/** 2) Archive this week’s RESPONSES into RESPONSES_Archive and remove them from RESPONSES */
function archiveAndResetWeek() {
  const ss = SpreadsheetApp.getActive();
  const weekStart = ymd_(getSelectedWeek_());

  const responses = ss.getSheetByName(TABS.RESPONSES);
  const archive = ss.getSheetByName(TABS.ARCHIVE) || ss.insertSheet(TABS.ARCHIVE);

  const data = responses.getDataRange().getValues();
  if (!data.length) return;
  const header = data[0];
  const rows = data.slice(1);

  const weekIdx = header.indexOf('week_start');
  if (weekIdx === -1) throw new Error('RESPONSES needs a "week_start" column (computed).');

  const keep = [header];
  const move = [header];

  for (const row of rows) {
    const wk = row[weekIdx];
    const wkStr = wk ? ymd_(wk) : '';
    if (wkStr === weekStart) move.push(row);
    else keep.push(row);
  }

  if (move.length > 1) {
    archive.getRange(archive.getLastRow()+1,1,move.length,move[0].length).setValues(move);
  }
  responses.clearContents();
  responses.getRange(1,1,keep.length,keep[0].length).setValues(keep);

  SpreadsheetApp.getUi().alert(`Archived ${move.length-1} row(s) for week ${weekStart}.`);
}

/** 3) Advance SETTINGS!B1 to next Monday and seed the 21 slots */
function goToNextWeekAndSeed() {
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName(TABS.SETTINGS);
  const cur = getSelectedWeek_();
  const next = new Date(cur); next.setDate(cur.getDate() + 7);
  settings.getRange(SETTINGS_RANGE).setValue(next);
  seedWeekSlots();
}
