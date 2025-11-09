# ChamCong APS Function Reference

This repository implements a Google Apps Script web app for managing weekly shift registrations. The codebase is split between backend `.gs` scripts (executed on Apps Script) and a single-page HTML frontend. This document summarizes every function provided in the project and explains its role.

## Backend Scripts (Apps Script)

### `config.gs`

Utility constants and helpers shared across the backend modules.

- `getSS()` – Opens the target spreadsheet, preferring `SPREADSHEET_ID` when configured; falls back to the bound spreadsheet.
- `getOrCreateSheet_(name)` – Returns a sheet with the provided name, creating it if missing.
- `isoDate_(d)` – Formats a `Date` object into a `yyyy-MM-dd` string in the configured timezone.
- `now_()` – Convenience wrapper returning the current `Date`.
- `parseISO_(iso)` – Parses an ISO date string (`yyyy-MM-dd`) into a `Date` object.
- `isAdminId_(userId)` – Checks whether a user ID belongs to the administrator set.
- `getProperties_()` – Returns the script properties store used for feature flags.
- `setAdminOverride_(val)` – Enables or disables the admin edit override flag in script properties.
- `isAdminOverride_()` – Reads the admin edit override flag.

### `chamcong.gs`

Core scheduling logic working against spreadsheet data.

- `bootstrapSheets()` – Initializes or resets the schema for Users, Responses, AuditLog, and Settings sheets, and turns off the admin override.
- `listUsersFromUsers_()` – Reads and normalizes user records from the `Users` sheet.
- `listUsersFromPeople_()` – Falls back to the `PEOPLE` sheet, mapping common header aliases to user fields.
- `listUsers_()` – Attempts to load users from `Users`; if unavailable, uses `PEOPLE`.
- `loadWeekMatrix(weekStartISO)` – Builds the seven-day schedule matrix for the requested week, including status colors for each user/day cell.
- `normalizeStatus_(status, subtype)` – Standardizes status/subtype combinations, constraining busy subtypes.
- `saveStatus(weekStartISO, dateISO, targetUserId, status, subtype, actorId)` – Upserts a single user/day status row, maintaining audit logs and timestamps.
- `computeMonthlyCounts_(year, month)` – Aggregates monthly participation counts per user, including subtype breakdowns.
- `emailMonthlyReminders(year, month)` – Sends summary or reminder emails based on monthly minimum shift requirements.
- `weekStartMonday_(d)` – Returns the Monday of the week for a given date.
- `nextWeekMonday_()` – Finds the Monday for the upcoming week relative to “now”.
- `nextWeekDeadline_()` – Calculates the Sunday 23:59 deadline preceding the next week.
- `isEditableForMember_(targetDateISO)` – Determines whether a member (non-admin) may edit a specific date, honoring deadlines and overrides.

### `server.gs`

HTTP endpoints, authentication, and action handlers for frontend calls.

- `doGet(e)` – Serves the `index.html` UI for web deployments.
- `route(payload)` – Wrapper for HtmlService calls, delegating to `dispatch_` and catching errors.
- `doPost(e)` – Accepts JSON POST payloads (e.g., external integrations) and routes them through the dispatcher.
- `dispatch_(payload)` – Central action router handling login, matrix retrieval, updates, admin toggles, and reporting tasks.
- `_norm(s)` – Normalizes strings for case/diacritic-insensitive comparisons.
- `_listFromPeople_()` – Loads user records from the `PEOPLE` sheet with alias-aware headers.
- `_listUsersCombined_()` – Combines canonical and PEOPLE sheet sourcing, mirroring `listUsers_()` from `chamcong.gs`.
- `_withCommonSession_(me)` – Produces the session payload shared with the client (user identity and flags).
- `login_(user_key, password)` – Implements classic username/password authentication (name or ID + person ID).
- `loginWithGoogle_()` – Authenticates via Google account, enforcing the `student.vgu.edu.vn` email pattern.
- `handleGetWeek_(p)` – Retrieves the week matrix and current feature flags for the frontend.
- `_enforceEditPermission_(actor_id, target_user_id, dateISO)` – Central permission gate respecting admin overrides, read-only mode, and weekly deadlines.
- `handleSetStatus_(p)` – Saves a single status change after permission checks and locking.
- `handleSetStatusBatch_(p)` – Persists multiple status changes atomically while enforcing permissions.
- `handleToggleAdminEdit_(p)` – Admin endpoint to toggle the global admin edit override flag.
- `isMatrixRO_()` – Reads the matrix read-only flag from script properties.
- `setMatrixRO_(val)` – Writes the matrix read-only flag.
- `handleSetMatrixRO_(p)` – Admin endpoint to toggle read-only mode.
- `handleGetAudit_(p)` – Returns the most recent audit log entries (up to 200 rows).
- `adminSetPassword(user_id, password)` – Stores or updates a manual password override for a user in the Settings sheet.
- `installMonthlyReminderTrigger()` – Installs a time-based trigger that runs on the first of each month at 08:00.
- `cronMonthlyReminder_()` – Trigger target that dispatches monthly reminder emails.

## Frontend Script (`index.html`)

Client-side logic managing the login flow, weekly grid interactions, and UI state.

- `hasDirty()` – Indicates whether there are staged (unsaved) cell edits.
- `toast(msg, ms)` / `toastRO()` – Lightweight notification helpers, including a specialized read-only warning.
- `toMondayISO(iso)` – Normalizes any date string to its corresponding Monday.
- `saveAllThen(cb)` – Saves all staged changes via `setStatusBatch` and reloads the matrix before invoking a callback.
- `keyFor(u, d)` – Generates a unique key for a user/date combination in the staged map.
- `updateSavebar()` – Reflects the staged-change count, enables/disables the save button, and shows or hides the save bar.
- `setLoginState(isLoading, msg, isError)` – Updates login button and message states during authentication attempts.
- `login()` – Handles classic credential submission through the backend `login` action.
- `loginWithGoogle()` – Initiates Google sign-in via the backend `loginWithGoogle` action.
- `finishLogin(d)` – Finalizes the login sequence, storing session data and initializing UI controls.
- `logout()` – Resets session state and returns to the login view.
- `setGridLoading(on)` – Shows or hides a loading indicator within the grid container.
- `_performLoadWeek(weekStartISO)` – Requests week data from the server, updates session flags, and renders the grid.
- `loadWeek()` – Public week loader that respects unsaved-change guards and optional autosave.
- `colorFor(status, subtype)` – Resolves a cell background color from the status/subtype combination.
- `labelFor(status, subtype)` – Produces the text label shown in each grid cell.
- `renderGrid()` – Renders the weekly matrix table based on current data and staged edits, applying permission-aware states.
- `onCellClick(ev, user_id, dateISO)` – Handles cell clicks to cycle statuses and open the subtype picker.
- `openPicker(x, y)` – Positions and shows the BUSY subtype picker near the cursor.
- `hidePicker()` – Hides the subtype picker and clears its context.
- `saveAll()` – Commits staged changes through `saveAllThen` and restores the save button state.
- `discardAll()` – Clears staged edits and rerenders the grid.
- `toggleAdminEdit(on)` – Calls the backend toggle for admin edit mode and syncs session state.
- `toggleMatrixRO(on)` – Calls the backend toggle for matrix read-only mode and refreshes UI locks.
- `sendMonthly()` – Prompts for a year-month and invokes the backend monthly reminder action.
- `reflectReadOnlyUI()` – Updates read-only badges and save-bar availability based on session flags.
- `(function () { ... })` – Immediately-invoked helper that wires up the “Show full matrix” checkbox label with accessible text.

Event listeners in the script also guard against unsaved navigation (`beforeunload`), close the picker on outside interactions, and apply picker selections.

## Scheduled/Trigger Functions

- `installMonthlyReminderTrigger()` (Apps Script) – Sets up the cron-style trigger for monthly reminders.
- `cronMonthlyReminder_()` (Apps Script) – Executes monthly reminders when triggered.

These functions, together with the constants defined in each file, form the complete behavior of the ChamCong APS scheduling tool.
