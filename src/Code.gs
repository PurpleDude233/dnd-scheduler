// Code.gs

const SHEET_PLAYERS = 'Players';
const SHEET_CAMPAIGNS = 'Campaigns';
const SHEET_SLOTS = 'Slots';
const SHEET_AVAIL = 'Availability';
const SHEET_RESULTS = 'Results';
const SHEET_CAMPTIMES = 'CampaignTimes';
const SHEET_LOCKED = 'LockedTimes';
const CALENDAR_ID = 'primary'; // or a specific calendar id/email
const SEND_CALENDAR_INVITES = true;
const DISPLAY_TZ = Session.getScriptTimeZone();
const ADMIN_LOCKER_ID = 'P04';
// ---- Rolling window settings ----
const ROLL_WEEKS_AHEAD = 6;     // keep slots for next N weeks
const SLOT_MINUTES = 30;       // must match your setup
const START_HOUR = 16;         // local time (script timezone)
const END_HOUR = 23;           // local time (script timezone)
const SHEET_TRACKER = 'Tracker';
const MIN_VISIBLE_SLOT_MINUTES = 16 * 60; // remove slots before 16:00

// If this script is not bound to a spreadsheet, set SPREADSHEET_ID in
// Apps Script Properties (Project Settings -> Script properties).
const SPREADSHEET_ID_PROP = 'SPREADSHEET_ID';
const SHEET_CACHE_TTL_SECONDS = 60;
const VERSIONED_SHEET_CACHE_TTL_SECONDS = 120;
const CACHED_SHEET_NAMES = Object.freeze([SHEET_PLAYERS, SHEET_CAMPAIGNS, SHEET_SLOTS]);
const VERSIONED_CACHED_SHEET_NAMES = Object.freeze([SHEET_AVAIL, SHEET_LOCKED, SHEET_TRACKER]);
const DASHBOARD_DATA_VERSION_PROP = 'DASHBOARD_DATA_VERSION';
const DERIVED_CACHE_TTL_SECONDS = 300;
const SERVER_PERF_LOGGING = true;
const SCHEDULER_TZ_REPAIR_PROP = 'SCHEDULER_TZ_REPAIR_V1';
const SCHEDULER_TZ_ROLLBACK_PROP = 'SCHEDULER_TZ_REPAIR_ROLLBACK_V1';
const SLOTS_DATA_KEY_PROP = 'SLOTS_DATA_KEY_V1';
const CLIENT_GRID_CACHE_SCHEMA_VERSION = '2';
const INIT_DATA_CACHE_SCHEMA_VERSION = '2';
let spreadsheetExecutionCache_ = null;
let spreadsheetExecutionCacheKey_ = '';
const sheetExecutionCacheBySpreadsheet_ = Object.create(null);

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = String(props.getProperty(SPREADSHEET_ID_PROP) || '').trim();
  const cacheKey = id ? `id:${id}` : 'active';

  if (spreadsheetExecutionCache_ && spreadsheetExecutionCacheKey_ === cacheKey) {
    return spreadsheetExecutionCache_;
  }

  if (id) {
    try {
      spreadsheetExecutionCache_ = SpreadsheetApp.openById(id);
      spreadsheetExecutionCacheKey_ = cacheKey;
      return spreadsheetExecutionCache_;
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      throw new Error(
        `Cannot access spreadsheet by id (${id}). ` +
        `Check sharing permissions for the script's user. Details: ${msg}`
      );
    }
  }

  // Fallback for bound scripts.
  spreadsheetExecutionCache_ = SpreadsheetApp.getActive();
  spreadsheetExecutionCacheKey_ = cacheKey;
  return spreadsheetExecutionCache_;
}

function getSheetByNameCached_(ss, sheetName) {
  const spreadsheet = ss || getSpreadsheet_();
  const name = String(sheetName || '').trim();
  if (!spreadsheet || !name) return null;

  const spreadsheetId = spreadsheet.getId ? spreadsheet.getId() : 'default';
  if (!sheetExecutionCacheBySpreadsheet_[spreadsheetId]) {
    sheetExecutionCacheBySpreadsheet_[spreadsheetId] = Object.create(null);
  }

  const cache = sheetExecutionCacheBySpreadsheet_[spreadsheetId];
  if (cache[name]) return cache[name];

  const sheet = spreadsheet.getSheetByName(name);
  if (sheet) cache[name] = sheet;
  return sheet;
}

function shouldUseCachedTable_(sheet) {
  return !!sheet && CACHED_SHEET_NAMES.indexOf(String(sheet.getName() || '')) !== -1;
}

function shouldUseVersionedCachedTable_(sheet) {
  return !!sheet && VERSIONED_CACHED_SHEET_NAMES.indexOf(String(sheet.getName() || '')) !== -1;
}

function getCachedTableKey_(sheet) {
  if (!sheet) return '';
  const parent = sheet.getParent && sheet.getParent();
  const spreadsheetId = parent && parent.getId ? parent.getId() : 'default';
  return `table:${spreadsheetId}:${sheet.getName()}`;
}

function getVersionedCachedTableKey_(sheet) {
  if (!sheet) return '';
  const parent = sheet.getParent && sheet.getParent();
  const spreadsheetId = parent && parent.getId ? parent.getId() : 'default';
  return `tablev:${getDashboardDataVersion_()}:${spreadsheetId}:${sheet.getName()}`;
}

function getOptimizedTableCacheKey_(sheet) {
  if (!sheet) return '';
  if (shouldUseCachedTable_(sheet)) return getCachedTableKey_(sheet);
  if (shouldUseVersionedCachedTable_(sheet)) return getVersionedCachedTableKey_(sheet);
  return '';
}

function getOptimizedTableCacheTtlSeconds_(sheet) {
  if (shouldUseVersionedCachedTable_(sheet)) return VERSIONED_SHEET_CACHE_TTL_SECONDS;
  return SHEET_CACHE_TTL_SECONDS;
}

function invalidateCachedTableByName_(ss, sheetName) {
  const spreadsheet = ss || getSpreadsheet_();
  const name = String(sheetName || '').trim();
  if (!spreadsheet || !name) return;

  const spreadsheetId = spreadsheet.getId ? spreadsheet.getId() : 'default';
  const cache = CacheService.getScriptCache();
  try {
    cache.remove(`table:${spreadsheetId}:${name}`);
  } catch (e) {}
  try {
    cache.remove(`tablev:${getDashboardDataVersion_()}:${spreadsheetId}:${name}`);
  } catch (e) {}
}

function getSlotRowValue_(row, key) {
  if (Array.isArray(row)) {
    if (key === 'slot_id') return row[0];
    if (key === 'start_utc') return row[1];
    if (key === 'end_utc') return row[2];
    return '';
  }
  return row && typeof row === 'object' ? row[key] : '';
}

function buildSlotsDataKeyFromRows_(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return 'slots:0';

  const first = list[0];
  const last = list[list.length - 1];
  return [
    list.length,
    Number(getSlotRowValue_(first, 'slot_id')) || 0,
    toIsoString_(getSlotRowValue_(first, 'start_utc')),
    Number(getSlotRowValue_(last, 'slot_id')) || 0,
    toIsoString_(getSlotRowValue_(last, 'end_utc'))
  ].join('|');
}

function setSlotsDataKey_(value) {
  const key = String(value || 'slots:0').trim() || 'slots:0';
  PropertiesService.getScriptProperties().setProperty(SLOTS_DATA_KEY_PROP, key);
  return key;
}

function refreshSlotsDataKeyFromRows_(rows) {
  return setSlotsDataKey_(buildSlotsDataKeyFromRows_(rows));
}

function refreshSlotsCacheMetadata_(ss, rows) {
  const spreadsheet = ss || getSpreadsheet_();
  invalidateCachedTableByName_(spreadsheet, SHEET_SLOTS);
  return refreshSlotsDataKeyFromRows_(rows);
}

function readOptimizedTable_(sheet) {
  if (!sheet) return [];
  const cacheKey = getOptimizedTableCacheKey_(sheet);
  if (!cacheKey) return readTable(sheet);

  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const rows = readTable(sheet);
  try {
    cache.put(cacheKey, JSON.stringify(rows), getOptimizedTableCacheTtlSeconds_(sheet));
  } catch (e) {}
  return rows;
}

function getDashboardDataVersion_() {
  const props = PropertiesService.getScriptProperties();
  return String(props.getProperty(DASHBOARD_DATA_VERSION_PROP) || '1').trim() || '1';
}

function bumpDashboardDataVersion_() {
  const props = PropertiesService.getScriptProperties();
  const current = Number(props.getProperty(DASHBOARD_DATA_VERSION_PROP) || '1');
  const next = Number.isFinite(current) ? current + 1 : Date.now();
  props.setProperty(DASHBOARD_DATA_VERSION_PROP, String(next));
  return String(next);
}

function getVersionedCacheKey_(name, extra) {
  const suffix = extra ? `:${Utilities.base64EncodeWebSafe(String(extra))}` : '';
  return `dashboard:${getDashboardDataVersion_()}:${name}${suffix}`;
}

function getCachedJson_(key) {
  if (!key) return null;
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function putCachedJson_(key, value, ttlSeconds) {
  if (!key) return value;
  try {
    CacheService.getScriptCache().put(
      key,
      JSON.stringify(value),
      ttlSeconds || DERIVED_CACHE_TTL_SECONDS
    );
  } catch (e) {}
  return value;
}

function getOrBuildCachedJson_(key, builder, ttlSeconds) {
  const cached = getCachedJson_(key);
  if (cached !== null) return cached;
  return putCachedJson_(key, builder(), ttlSeconds);
}

function setSingleRowValues_(sheet, rowIndex, rowValues) {
  if (!sheet || !Number.isFinite(rowIndex) || rowIndex < 1) return;
  const values = Array.isArray(rowValues) ? rowValues : [];
  if (!values.length) return;
  sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
}

function appendRowsWithSetValues_(sheet, rows) {
  const values = Array.isArray(rows)
    ? rows.filter(row => Array.isArray(row) && row.length)
    : [];
  if (!sheet || !values.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
}

function getActorEmail_() {
  const effectiveEmail = (Session.getEffectiveUser && Session.getEffectiveUser().getEmail)
    ? (Session.getEffectiveUser().getEmail() || '')
    : '';
  const activeEmail = (Session.getActiveUser && Session.getActiveUser().getEmail)
    ? (Session.getActiveUser().getEmail() || '')
    : '';
  return effectiveEmail || activeEmail || '';
}

function logServerPerf_(label, startedAt, meta) {
  if (!SERVER_PERF_LOGGING) return;
  const durationMs = Date.now() - startedAt;
  const suffix = meta ? ` ${meta}` : '';
  console.log(`[perf] ${label}: ${durationMs}ms${suffix}`);
}

function runTimed_(label, fn, meta) {
  const startedAt = Date.now();
  try {
    return fn();
  } finally {
    logServerPerf_(label, startedAt, meta);
  }
}

// ---- Availability storage (compact) ----
// Sheet columns: player_id | slot_ids | updated_at | updated_by
function ensureAvailabilitySheetInSpreadsheet_(ss) {
  const sh = ss.getSheetByName(SHEET_AVAIL) || ss.insertSheet(SHEET_AVAIL);
  const wanted = ['player_id','slot_ids','updated_at','updated_by'];

  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,wanted.length).setValues([wanted]);
    sh.getRange(1,1,1,wanted.length).setFontWeight('bold');
    return sh;
  }

  const data = sh.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());

  // Old format detection: player_id + slot_id rows
  if (header.includes('slot_id') && header.includes('player_id') && !header.includes('slot_ids')) {
    migrateAvailabilityToCompact_(sh, data, header);
    return sh;
  }

  // If header already matches, keep.
  const missing = wanted.some(h => !header.includes(h));
  if (!missing) return sh;

  // Best-effort remap to new header (preserve existing values)
  const idx = {};
  header.forEach((h,i)=>{ idx[h]=i; });
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const out = wanted.map(h => (idx[h] !== undefined ? row[idx[h]] : ''));
    const any = out.some(v => String(v).trim() !== '');
    if (any) rows.push(out);
  }

  sh.clearContents();
  sh.getRange(1,1,1,wanted.length).setValues([wanted]);
  sh.getRange(1,1,1,wanted.length).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2,1,rows.length,wanted.length).setValues(rows);
  }
  bumpDashboardDataVersion_();
  return sh;
}

function ensureAvailabilitySheet_() {
  return ensureAvailabilitySheetInSpreadsheet_(getSpreadsheet_());
}

function migrateAvailabilityToCompact_(sh, data, header) {
  const idxUpdatedAt = header.indexOf('updated_at');
  const idxPid = header.indexOf('player_id');
  const idxSlotId = header.indexOf('slot_id');

  const byPlayer = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pid = String(row[idxPid] || '').trim();
    const slotId = Number(row[idxSlotId]);
    if (!pid || !Number.isFinite(slotId)) continue;

    if (!byPlayer.has(pid)) {
      byPlayer.set(pid, { slots: new Set(), updatedAt: '' });
    }
    const entry = byPlayer.get(pid);
    entry.slots.add(slotId);

    if (idxUpdatedAt !== -1) {
      const ts = String(row[idxUpdatedAt] || '').trim();
      if (ts) {
        const cur = entry.updatedAt ? new Date(entry.updatedAt).getTime() : -1;
        const next = new Date(ts).getTime();
        if (Number.isFinite(next) && next > cur) entry.updatedAt = new Date(next).toISOString();
      }
    }
  }

  const wanted = ['player_id','slot_ids','updated_at','updated_by'];
  const rows = [];
  for (const [pid, entry] of byPlayer.entries()) {
    const slotCsv = Array.from(entry.slots).sort((a,b)=>a-b).join(',');
    rows.push([pid, slotCsv, entry.updatedAt || '', '']);
  }

  sh.clearContents();
  sh.getRange(1,1,1,wanted.length).setValues([wanted]);
  sh.getRange(1,1,1,wanted.length).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2,1,rows.length,wanted.length).setValues(rows);
  }
  bumpDashboardDataVersion_();
}

function parseSlotIds_(csv) {
  return String(csv || '')
    .split(',')
    .map(s => Number(String(s).trim()))
    .filter(n => Number.isFinite(n) && n > 0);
}

function getPlayerSlotIdsFromRows_(rows, playerId) {
  const pid = String(playerId || '').trim();
  for (const r of (rows || [])) {
    if (String(r.player_id || '').trim() !== pid) continue;
    return parseSlotIds_(r.slot_ids);
  }
  return [];
}

function getPlayerSlotIds_(playerId) {
  const sh = ensureAvailabilitySheet_();
  return getPlayerSlotIdsFromRows_(readOptimizedTable_(sh), playerId);
}

function buildAvailabilityMapFromRows_(rows) {
  const availMap = new Map();
  for (const r of (rows || [])) {
    const pid = String(r.player_id || '').trim();
    if (!pid) continue;
    const slotIds = parseSlotIds_(r.slot_ids);
    for (const slotId of slotIds) {
      if (!availMap.has(slotId)) availMap.set(slotId, new Set());
      availMap.get(slotId).add(pid);
    }
  }
  return availMap;
}

function buildAvailabilityMap_() {
  const sh = ensureAvailabilitySheet_();
  return buildAvailabilityMapFromRows_(readOptimizedTable_(sh));
}

function buildLockedSlotIdsFromRows_(lockedRows) {
  const set = new Set();
  for (const r of (lockedRows || [])) {
    const slotIds = String(r.slot_ids || '')
      .split(',')
      .map(s => Number(String(s).trim()))
      .filter(n => Number.isFinite(n) && n > 0);
    slotIds.forEach(id => set.add(id));
  }
  return set;
}

function buildLockedSlotCampaignMapFromRows_(lockedRows) {
  const map = new Map();
  for (const r of (lockedRows || [])) {
    const campaign = String(r.campaign || '').trim();
    if (!campaign) continue;
    const slotIds = String(r.slot_ids || '')
      .split(',')
      .map(s => Number(String(s).trim()))
      .filter(n => Number.isFinite(n) && n > 0);
    for (const id of slotIds) {
      if (!map.has(id)) map.set(id, campaign);
    }
  }
  return map;
}

function ensureTrackerSheetInSpreadsheet_(ss) {
  const sh = ss.getSheetByName(SHEET_TRACKER) || ss.insertSheet(SHEET_TRACKER);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([['player_id','last_saved_at','last_saved_by']]);
    sh.getRange(1,1,1,3).setFontWeight('bold');
  }
  return sh;
}

function normalizeDashboardSnapshotOptions_(options) {
  const opts = options && typeof options === 'object' ? options : {};
  return {
    includePlayers: !!opts.includePlayers,
    includeCampaigns: !!opts.includeCampaigns,
    includeSlots: !!opts.includeSlots,
    includeAvailability: !!opts.includeAvailability,
    includeLocked: !!opts.includeLocked,
    includeTracker: !!opts.includeTracker
  };
}

function getDashboardSnapshot_(options) {
  const opts = normalizeDashboardSnapshotOptions_(options);
  const cacheKey = getVersionedCacheKey_('snapshot', JSON.stringify(opts));

  return getOrBuildCachedJson_(cacheKey, () => {
    const ss = getSpreadsheet_();
    const snapshot = {};

    if (opts.includePlayers) {
      snapshot.players = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_PLAYERS));
    }
    if (opts.includeCampaigns) {
      snapshot.campaigns = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_CAMPAIGNS));
    }
    if (opts.includeSlots) {
      snapshot.slots = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_SLOTS));
    }
    if (opts.includeAvailability) {
      snapshot.availability = readOptimizedTable_(ensureAvailabilitySheetInSpreadsheet_(ss));
    }
    if (opts.includeLocked) {
      snapshot.locked = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_LOCKED));
    }
    if (opts.includeTracker) {
      snapshot.tracker = readOptimizedTable_(ensureTrackerSheetInSpreadsheet_(ss));
    }

    return snapshot;
  }, SHEET_CACHE_TTL_SECONDS);
}

function toClientPlayers_(rows) {
  return (rows || []).map(r => ({
    player_id: String(r.player_id || ''),
    name: String(r.name || ''),
    usual_times: String(r.usual_times || '')
  }));
}

function parseClientGridDayValue_(dayLabel) {
  const match = String(dayLabel || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return Number.MAX_SAFE_INTEGER;
  return Number(`${match[3]}${match[2]}${match[1]}`);
}

function buildClientGridData_(slotRows) {
  const byDay = new Map();
  const timeSet = new Set();

  (slotRows || []).forEach(r => {
    const slotId = Number(r.slot_id);
    if (!Number.isFinite(slotId) || slotId <= 0) return;

    const localStart = getLocalPartsFromIso_(r.start_utc, DISPLAY_TZ);
    if (!localStart) return;

    const mins = localStart.hour * 60 + localStart.minute;
    if (mins < MIN_VISIBLE_SLOT_MINUTES) return;

    const day = localStart.dayLabel;
    if (!byDay.has(day)) {
      byDay.set(day, {
        day,
        weekday: localStart.weekday,
        slots: []
      });
    }

    byDay.get(day).slots.push({ id: slotId, mins });
    timeSet.add(mins);
  });

  const timeMins = Array.from(timeSet).sort((a, b) => a - b);
  const rows = Array.from(byDay.values())
    .sort((a, b) => parseClientGridDayValue_(a.day) - parseClientGridDayValue_(b.day))
    .map(row => ({
      day: row.day,
      weekday: Number(row.weekday) || 0,
      slots: (row.slots || []).sort((a, b) => a.mins - b.mins)
    }));

  return { timeMins, rows };
}

function isValidClientGridData_(payload) {
  return !!payload &&
    typeof payload === 'object' &&
    Array.isArray(payload.timeMins) &&
    Array.isArray(payload.rows);
}

function getCachedClientGridData_(slotRows, slotDataKey) {
  const build = () => buildClientGridData_(slotRows);
  if (!slotDataKey) return build();

  const cacheKey = getVersionedCacheKey_(
    'clientGrid',
    `${CLIENT_GRID_CACHE_SCHEMA_VERSION}:${slotDataKey}`
  );
  const cached = getCachedJson_(cacheKey);
  if (isValidClientGridData_(cached)) return cached;

  return putCachedJson_(cacheKey, build());
}

function getInitDataFromSnapshot_(snapshot, playerId, slotDataKey) {
  const rawSlots = snapshot && Array.isArray(snapshot.slots) ? snapshot.slots : [];

  const selectedSlotIds = new Set(getPlayerSlotIdsFromRows_(snapshot.availability, playerId));
  const lockedSlotIds = Array.from(buildLockedSlotIdsFromRows_(snapshot.locked));
  const lockedSlotCampaigns = Array.from(buildLockedSlotCampaignMapFromRows_(snapshot.locked).entries())
    .map(([slot_id, campaign]) => ({ slot_id, campaign }));

  return {
    players: toClientPlayers_(snapshot.players),
    gridData: getCachedClientGridData_(rawSlots, slotDataKey),
    selectedSlotIds: Array.from(selectedSlotIds),
    lockedSlotIds,
    lockedSlotCampaigns
  };
}

function buildAvailabilitySnapshotForSlots_(slotIds) {
  const target = new Set((slotIds || []).map(n => Number(n)).filter(n => Number.isFinite(n) && n > 0));
  const snapshot = {};
  if (!target.size) return snapshot;

  const rows = readOptimizedTable_(ensureAvailabilitySheet_());
  for (const r of rows) {
    const pid = String(r.player_id || '').trim();
    if (!pid) continue;
    const playerSlotIds = parseSlotIds_(r.slot_ids);
    for (const slotId of playerSlotIds) {
      if (!target.has(slotId)) continue;
      if (!snapshot[slotId]) snapshot[slotId] = [];
      snapshot[slotId].push(pid);
    }
  }

  Object.keys(snapshot).forEach(slotId => snapshot[slotId].sort());
  return snapshot;
}

function restoreAvailabilitySnapshot_(snapshot) {
  const snap = snapshot && typeof snapshot === 'object' ? snapshot : {};
  const playerToSlots = new Map();

  Object.keys(snap).forEach(slotIdRaw => {
    const slotId = Number(slotIdRaw);
    if (!Number.isFinite(slotId) || slotId <= 0) return;
    const players = Array.isArray(snap[slotIdRaw]) ? snap[slotIdRaw] : [];
    players.forEach(pidRaw => {
      const pid = String(pidRaw || '').trim();
      if (!pid) return;
      if (!playerToSlots.has(pid)) playerToSlots.set(pid, new Set());
      playerToSlots.get(pid).add(slotId);
    });
  });

  if (!playerToSlots.size) return false;

  const sh = ensureAvailabilitySheet_();
  const data = sh.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());
  const idxPid = header.indexOf('player_id');
  const idxSlots = header.indexOf('slot_ids');
  const idxUpdatedAt = header.indexOf('updated_at');
  const idxUpdatedBy = header.indexOf('updated_by');
  if (idxPid === -1 || idxSlots === -1) return false;

  const nowIso = new Date().toISOString();
  const who = getActorEmail_();

  const rowIndexByPid = new Map();
  for (let i = 1; i < data.length; i++) {
    const pid = String(data[i][idxPid] || '').trim();
    if (pid) rowIndexByPid.set(pid, i + 1);
  }

  for (const [pid, slotSet] of playerToSlots.entries()) {
    const rowIndex = rowIndexByPid.get(pid);
    if (rowIndex) {
      const row = data[rowIndex - 1].slice();
      const current = new Set(parseSlotIds_(row[idxSlots]));
      slotSet.forEach(slotId => current.add(slotId));
      row[idxSlots] = Array.from(current).sort((a,b)=>a-b).join(',');
      if (idxUpdatedAt !== -1) row[idxUpdatedAt] = nowIso;
      if (idxUpdatedBy !== -1) row[idxUpdatedBy] = who;
      data[rowIndex - 1] = row;
    } else {
      const newRow = new Array(header.length).fill('');
      newRow[idxPid] = pid;
      newRow[idxSlots] = Array.from(slotSet).sort((a,b)=>a-b).join(',');
      if (idxUpdatedAt !== -1) newRow[idxUpdatedAt] = nowIso;
      if (idxUpdatedBy !== -1) newRow[idxUpdatedBy] = who;
      data.push(newRow);
      rowIndexByPid.set(pid, data.length);
    }
  }

  sh.getRange(1,1,data.length,data[0].length).setValues(data);

  return true;
}



/**
 * Main daily maintenance:
 * - removes past Slots
 * - removes Availability rows referring to removed slots
 * - removes past LockedTimes
 * - adds new Slots up to the rolling horizon
 * - recomputes CampaignTimes & Results
 */

function getCampaignRowFromRows_(campaigns, campaignName) {
  const wanted = String(campaignName || '').trim();
  if (!wanted) return null;
  const byName = (campaigns || []).find(c =>
    String(c.campaign_name || '').trim() === wanted
  );
  if (byName) return byName;
  return (campaigns || []).find(c =>
    String(c.campaign_id || '').trim() === wanted
  ) || null;
}

function getCampaignRequiredPlayerIdsFromRow_(campaignRow) {
  return String((campaignRow && campaignRow.required_players) || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

function getCampaignRequiredPlayerIds_(campaignName) {
  const ss = getSpreadsheet_();
  const campaigns = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_CAMPAIGNS));
  return getCampaignRequiredPlayerIdsFromRow_(getCampaignRowFromRows_(campaigns, campaignName));
}

function getCampaignByName_(campaignName) {
  const ss = getSpreadsheet_();
  const campaigns = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_CAMPAIGNS));
  return getCampaignRowFromRows_(campaigns, campaignName);
}

function isOneShotCampaign_(campaignRow) {
  const id = String((campaignRow && campaignRow.campaign_id) || '').trim().toLowerCase();
  const name = String((campaignRow && campaignRow.campaign_name) || '').trim().toLowerCase();
  return /(^|[^a-z])one[\s-]?shot([^a-z]|$)/i.test(id) || /(^|[^a-z])one[\s-]?shot([^a-z]|$)/i.test(name);
}

function getCampaignPlayableThreshold_(campaignRow, requiredIds) {
  const reqCount = Array.isArray(requiredIds) ? requiredIds.length : 0;
  if (reqCount <= 0) return 0;
  if (isOneShotCampaign_(campaignRow)) return Math.min(4, reqCount);
  return reqCount;
}

function buildCampaignMeta_(campaigns) {
  return (campaigns || []).map(c => {
    const req = String(c.required_players || '').split(',').map(s => s.trim()).filter(Boolean);
    const playableThreshold = getCampaignPlayableThreshold_(c, req);
    return {
      id: String(c.campaign_id || '').trim(),
      name: String(c.campaign_name || c.campaign_id || 'Unnamed').trim(),
      req,
      required_count: req.length,
      playable_threshold: playableThreshold,
      is_oneshot: isOneShotCampaign_(c)
    };
  }).filter(c => c.req.length);
}

function getTimeZoneOffsetMinutes_(date, timeZone) {
  const value = Utilities.formatDate(date, timeZone || DISPLAY_TZ, 'Z');
  const match = String(value || '').match(/([+-])(\d{2})(\d{2})/);
  if (!match) return 0;
  const sign = match[1] === '-' ? -1 : 1;
  return sign * ((Number(match[2]) * 60) + Number(match[3]));
}

function buildUtcDateFromLocalParts_(year, monthIndex, day, hour, minute, second, millisecond, timeZone) {
  const tz = timeZone || DISPLAY_TZ;
  const baseUtcMs = Date.UTC(
    Number(year) || 0,
    Number(monthIndex) || 0,
    Number(day) || 1,
    Number(hour) || 0,
    Number(minute) || 0,
    Number(second) || 0,
    Number(millisecond) || 0
  );

  let candidate = new Date(baseUtcMs);
  let offsetMinutes = getTimeZoneOffsetMinutes_(candidate, tz);
  let resolvedUtcMs = baseUtcMs - (offsetMinutes * 60000);
  candidate = new Date(resolvedUtcMs);

  const finalOffsetMinutes = getTimeZoneOffsetMinutes_(candidate, tz);
  if (finalOffsetMinutes !== offsetMinutes) {
    resolvedUtcMs = baseUtcMs - (finalOffsetMinutes * 60000);
    candidate = new Date(resolvedUtcMs);
  }

  return candidate;
}

function parseStoredIsoWallParts_(iso) {
  const raw = String(iso || '').trim();
  const match = raw.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(?:Z)?$/
  );
  if (!match) return null;
  const msRaw = String(match[7] || '0');
  return {
    year: Number(match[1]),
    monthIndex: Number(match[2]) - 1,
    day: Number(match[3]),
    hour: Number(match[4]),
    minute: Number(match[5]),
    second: Number(match[6] || 0),
    millisecond: Number(msRaw.padEnd(3, '0').slice(0, 3))
  };
}

function getLocalPartsFromIso_(iso, timeZone) {
  const d = new Date(String(iso || ''));
  if (isNaN(d.getTime())) return null;
  const tz = timeZone || DISPLAY_TZ;
  const weekday = Number(Utilities.formatDate(d, tz, 'u'));
  return {
    year: Number(Utilities.formatDate(d, tz, 'yyyy')),
    month: Number(Utilities.formatDate(d, tz, 'M')),
    day: Number(Utilities.formatDate(d, tz, 'd')),
    weekday: Number.isFinite(weekday) ? (weekday % 7) : 0,
    hour: Number(Utilities.formatDate(d, tz, 'H')),
    minute: Number(Utilities.formatDate(d, tz, 'm')),
    dayLabel: Utilities.formatDate(d, tz, 'dd/MM/yyyy'),
    dayKey: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
    timeLabel: Utilities.formatDate(d, tz, 'HH:mm')
  };
}

function convertStoredLocalIsoToTrueUtc_(iso, timeZone) {
  const parts = parseStoredIsoWallParts_(iso);
  if (!parts) return '';
  return buildUtcDateFromLocalParts_(
    parts.year,
    parts.monthIndex,
    parts.day,
    parts.hour,
    parts.minute,
    parts.second,
    parts.millisecond,
    timeZone || DISPLAY_TZ
  ).toISOString();
}

function convertTrueUtcToStoredLocalIso_(iso, timeZone) {
  const d = new Date(String(iso || ''));
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(
    d,
    timeZone || DISPLAY_TZ,
    "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'"
  );
}

function getLocalSlotKeyFromIso_(iso, timeZone) {
  const parts = getLocalPartsFromIso_(iso, timeZone);
  if (!parts) return '';
  return `${parts.dayKey}T${parts.timeLabel}`;
}

function formatLocalSlotKeyFromDate_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatString(
    '%04d-%02d-%02dT%02d:%02d',
    date.getFullYear(),
    date.getMonth() + 1,
    date.getDate(),
    date.getHours(),
    date.getMinutes()
  );
}

function toUtcIso_(localDate) {
  if (!(localDate instanceof Date)) return '';
  return buildUtcDateFromLocalParts_(
    localDate.getFullYear(),
    localDate.getMonth(),
    localDate.getDate(),
    localDate.getHours(),
    localDate.getMinutes(),
    localDate.getSeconds(),
    localDate.getMilliseconds(),
    DISPLAY_TZ
  ).toISOString();
}

function fromIsoToLocal_(iso) {
  const parts = getLocalPartsFromIso_(iso, DISPLAY_TZ);
  if (!parts) return null;
  return new Date(Date.UTC(
    parts.year,
    parts.month - 1,
    parts.day,
    parts.hour,
    parts.minute,
    0,
    0
  ));
}

function shouldDropSlotByLocalMinute_(iso) {
  const parts = getLocalPartsFromIso_(iso, DISPLAY_TZ);
  if (!parts) return false;
  return (parts.hour * 60 + parts.minute) < MIN_VISIBLE_SLOT_MINUTES;
}

function getPlayerEmailsByIdsFromRows_(players, playerIds) {
  const emailById = new Map();
  for (const p of players) {
    const id = String(p.player_id || '').trim();
    const email = String(p.email || '').trim();
    if (id && email) emailById.set(id, email);
  }

  return (playerIds || []).map(id => emailById.get(String(id))).filter(Boolean);
}

function getPlayerEmailsByIds_(playerIds) {
  const ss = getSpreadsheet_();
  const players = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_PLAYERS));
  return getPlayerEmailsByIdsFromRows_(players, playerIds);
}

function getLocalWindowLabels_(startMs, endMs, timeZone) {
  if (!Number.isFinite(startMs) || !Number.isFinite(endMs)) {
    return { date_local: '', time_local: '' };
  }
  const tz = timeZone || DISPLAY_TZ;
  const start = new Date(startMs);
  const end = new Date(endMs);
  return {
    date_local: Utilities.formatDate(start, tz, 'dd.MM.yyyy (EEE)'),
    time_local: `${Utilities.formatDate(start, tz, 'HH:mm')}-${Utilities.formatDate(end, tz, 'HH:mm')}`
  };
}

function updateLockedLocalWindowFields_(row, idxStart, idxEnd, idxDateLocal, idxTimeLocal) {
  if (!row || (idxDateLocal === -1 && idxTimeLocal === -1)) return false;

  const startMs = new Date(String(idxStart !== -1 ? row[idxStart] : '')).getTime();
  const endMs = new Date(String(idxEnd !== -1 ? row[idxEnd] : '')).getTime();
  const labels = getLocalWindowLabels_(startMs, endMs, DISPLAY_TZ);

  let changed = false;
  if (idxDateLocal !== -1 && String(row[idxDateLocal] || '') !== labels.date_local) {
    row[idxDateLocal] = labels.date_local;
    changed = true;
  }
  if (idxTimeLocal !== -1 && String(row[idxTimeLocal] || '') !== labels.time_local) {
    row[idxTimeLocal] = labels.time_local;
    changed = true;
  }
  return changed;
}

function fillLockedLocalWindowColumns_(sh) {
  if (!sh || sh.getLastRow() < 2) return 0;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return 0;

  const header = data[0].map(h => String(h).trim());
  const idxStart = header.indexOf('start_utc');
  const idxEnd = header.indexOf('end_utc');
  const idxDateLocal = header.indexOf('date_local');
  const idxTimeLocal = header.indexOf('time_local');
  if (idxStart === -1 || idxEnd === -1 || (idxDateLocal === -1 && idxTimeLocal === -1)) return 0;

  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    if (updateLockedLocalWindowFields_(data[i], idxStart, idxEnd, idxDateLocal, idxTimeLocal)) {
      updated += 1;
    }
  }

  if (updated > 0) {
    sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
  return updated;
}

/**
 * Creates a calendar event and invites guests.
 * Returns eventId (string) or '' if not created.
 */
function createCalendarEventForLock_(campaignName, startMs, endMs, guestEmails) {
  if (!SEND_CALENDAR_INVITES) return '';

  const start = new Date(startMs);
  const end = new Date(endMs);

  const title = `${campaignName}`;
  const desc =
    `Session locked in by ${ADMIN_LOCKER_ID}.\n` +
    `Campaign: ${campaignName}\n`;

  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('Calendar not found. Check CALENDAR_ID.');

  const guestsCsv = (guestEmails || []).join(',');
  // Create event and invite guests
  const ev = cal.createEvent(title, start, end, {
    description: desc,
    guests: guestsCsv,
    sendInvites: true
  });

  // Ensure guests are added (some calendars are picky with the guests option)
  (guestEmails || []).forEach(email => {
    if (!email) return;
    try { ev.addGuest(email); } catch (e) {}
  });

  return ev.getId();
}

function deleteCalendarEventForLock_(eventId) {
  const id = String(eventId || '').trim();
  if (!id) return false;

  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('Calendar not found. Check CALENDAR_ID.');

  const ev = cal.getEventById(id);
  if (!ev) return false;
  ev.deleteEvent();
  return true;
}

function getLockedSlotIds_() {
  const ss = getSpreadsheet_();
  const shLocked = getSheetByNameCached_(ss, SHEET_LOCKED);
  return buildLockedSlotIdsFromRows_(readOptimizedTable_(shLocked));
}

function getLockedSlotCampaignMap_() {
  const ss = getSpreadsheet_();
  const shLocked = getSheetByNameCached_(ss, SHEET_LOCKED);
  return buildLockedSlotCampaignMapFromRows_(readOptimizedTable_(shLocked));
}

function ensureLockedSheetInSpreadsheet_(ss) {
  const spreadsheet = ss || getSpreadsheet_();
  const sh = spreadsheet.getSheetByName(SHEET_LOCKED) || spreadsheet.insertSheet(SHEET_LOCKED);
  const wanted = [
    'locked_at',
    'campaign_id',
    'campaign',
    'start_utc',
    'end_utc',
    'slot_ids',
    'locked_by',
    'event_id',
    'guest_emails',
    'availability_snapshot',
    'date_local',
    'time_local'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,wanted.length).setValues([wanted]);
    sh.getRange(1,1,1,wanted.length).setFontWeight('bold');
    return sh;
  }

  const data = sh.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());
  const missing = wanted.some(h => !header.includes(h));
  if (!missing) return sh;

  // Best-effort migration: remap existing columns into new header
  const idx = {};
  header.forEach((h,i)=>{ idx[h]=i; });
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const out = wanted.map(h => (idx[h] !== undefined ? row[idx[h]] : ''));
    const any = out.some(v => String(v).trim() !== '');
    if (any) rows.push(out);
  }

  sh.clearContents();
  sh.getRange(1,1,1,wanted.length).setValues([wanted]);
  sh.getRange(1,1,1,wanted.length).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2,1,rows.length,wanted.length).setValues(rows);
  }
  return sh;
}

function ensureLockedSheet_() {
  return ensureLockedSheetInSpreadsheet_(getSpreadsheet_());
}

function maintainRollingWindow() {
  const ss = getSpreadsheet_();
  const shSlots = getSheetByNameCached_(ss, SHEET_SLOTS);
  const shAvail = getSheetByNameCached_(ss, SHEET_AVAIL);
  const shLocked = ensureLockedSheet_();

  if (!shSlots) throw new Error('Slots sheet missing');

  const now = new Date();
  const nowMs = now.getTime();

  // ---------- Read Slots ----------
  const slotsTable = readTable(shSlots); // expects slot_id,start_utc,end_utc
  const slotRows = slotsTable.map(r => {
    const id = Number(r.slot_id);
    const start = new Date(String(r.start_utc));
    const end = new Date(String(r.end_utc));
    return {
      slot_id: id,
      start_utc: toIsoString_(r.start_utc),
      end_utc: toIsoString_(r.end_utc),
      startMs: start.getTime(),
      endMs: end.getTime(),
    };
  }).filter(s => Number.isFinite(s.slot_id) && Number.isFinite(s.endMs));

  // Keep only slots whose end is still in the future and are not hidden minute buckets.
  const keptSlots = slotRows.filter(s => s.endMs >= nowMs && !shouldDropSlotByLocalMinute_(s.start_utc));

  // Create a set of kept slot IDs
  const keptIdSet = new Set(keptSlots.map(s => s.slot_id));

  // ---------- Rewrite Slots (header + kept) ----------
  // Ensure header exists
  shSlots.clearContents();
  shSlots.getRange(1,1,1,3).setValues([['slot_id','start_utc','end_utc']]);

  // Sort kept slots by time
  keptSlots.sort((a,b) => a.startMs - b.startMs);

  if (keptSlots.length) {
    const keptValues = keptSlots.map(s => [s.slot_id, s.start_utc, s.end_utc]);
    shSlots.getRange(2,1,keptValues.length,3).setValues(keptValues);
  }

  // ---------- Clean Availability (remove missing slot_ids) ----------
  if (shAvail) {
    const shA = ensureAvailabilitySheet_();
    const data = shA.getDataRange().getValues();
    if (data.length > 0) {
      const header = data[0].map(h => String(h).trim());
      const idxPid = header.indexOf('player_id');
      const idxSlots = header.indexOf('slot_ids');
      const idxUpdatedAt = header.indexOf('updated_at');
      const idxUpdatedBy = header.indexOf('updated_by');

      if (idxPid !== -1 && idxSlots !== -1) {
        const out = [data[0]];
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const pid = String(row[idxPid] || '').trim();
          if (!pid) continue;
          const slotIds = parseSlotIds_(row[idxSlots]);
          const kept = slotIds.filter(id => keptIdSet.has(id));
          const newRow = row.slice();
          newRow[idxSlots] = kept.join(',');
          // keep updated_at/by unchanged
          if (idxUpdatedAt !== -1 && !newRow[idxUpdatedAt]) newRow[idxUpdatedAt] = row[idxUpdatedAt] || '';
          if (idxUpdatedBy !== -1 && !newRow[idxUpdatedBy]) newRow[idxUpdatedBy] = row[idxUpdatedBy] || '';
          out.push(newRow);
        }
        shA.clearContents();
        shA.getRange(1,1,out.length,header.length).setValues(out);
      }
    }
  }

  // ---------- Clean LockedTimes (remove locks whose end_utc passed) ----------
  if (shLocked) {
    const data = shLocked.getDataRange().getValues();
    const header = data[0].map(h => String(h).trim());
    const out = [data[0]];

    const idxEnd = header.indexOf('end_utc');
    const idxSlotIds = header.indexOf('slot_ids');

    // If header is unexpected, keep everything (fail-safe)
    if (idxEnd === -1 || idxSlotIds === -1) {
      // do nothing
    } else {
      // Keep only locks whose end is still >= now, AND whose slot_ids still exist
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const end = new Date(String(row[idxEnd] || ''));
        const endMs = end.getTime();
        if (!Number.isFinite(endMs)) continue;
        if (endMs < nowMs) continue;

        const slotIds = String(row[idxSlotIds] || '')
          .split(',')
          .map(s => Number(String(s).trim()))
          .filter(n => Number.isFinite(n) && n > 0);

        // If any slot in the lock no longer exists, drop the lock entry
        // (prevents “ghost locks” after pruning)
        const allExist = slotIds.every(id => keptIdSet.has(id));
        if (!allExist) continue;

        out.push(row);
      }

      shLocked.clearContents();
      shLocked.getRange(1,1,out.length,out[0].length).setValues(out);
    }
  }

  // ---------- Add new slots up to rolling horizon ----------
  // Determine horizon end = today + ROLL_WEEKS_AHEAD
  const horizon = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
  horizon.setDate(horizon.getDate() + ROLL_WEEKS_AHEAD * 7);
  const horizonMs = horizon.getTime();

  // Find max slot_id to continue incrementing
  let maxId = 0;
  for (const s of keptSlots) maxId = Math.max(maxId, s.slot_id);

  // Build a set of existing slot start times (local) to avoid duplicates
  const existingStart = new Set();
  keptSlots.forEach(s => {
    const localStartKey = getLocalSlotKeyFromIso_(s.start_utc, DISPLAY_TZ);
    if (!localStartKey) return;
    existingStart.add(localStartKey);
  });

  const newRows = [];
  function addDaySlots(dateObj) {
    // dateObj at midnight
    for (let mins = START_HOUR * 60; mins < END_HOUR * 60; mins += SLOT_MINUTES) {
      const localStart = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), 0, mins, 0, 0);
      const localEnd = new Date(localStart.getTime() + SLOT_MINUTES * 60 * 1000);

      if (existingStart.has(formatLocalSlotKeyFromDate_(localStart))) continue;

      // store UTC ISO correctly
      const startUtc = toUtcIso_(localStart);
      const endUtc = toUtcIso_(localEnd);

      maxId += 1;
      newRows.push([maxId, startUtc, endUtc]);
      keptIdSet.add(maxId);
    }
  }

  // Generate by day from today to horizon (deterministic)
  const startDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0,0,0,0);
  for (let d = new Date(startDay); d.getTime() <= horizonMs; d.setDate(d.getDate() + 1)) {
    addDaySlots(d);
  }

  // Append new rows to Slots
  if (newRows.length) {
    const lastRow = shSlots.getLastRow();
    shSlots.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
  }

  refreshSlotsCacheMetadata_(ss, keptSlots.concat(
    newRows.map(row => ({
      slot_id: row[0],
      start_utc: row[1],
      end_utc: row[2]
    }))
  ));

  // ---------- Recompute derived sheets ----------
  recomputeDerivedSheets_();
  bumpDashboardDataVersion_();
}

/**
 * One-time helper to install a daily trigger (runs in your script timezone).
 */
function installDailyMaintenanceTrigger() {
  // Remove existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'maintainRollingWindow') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('maintainRollingWindow')
    .timeBased()
    .everyDays(1)
    .atHour(3) // pick a quiet hour
    .create();
}

/**
 * One-time helper: recreate all required sheets and headers.
 * Safe to run after clearing the spreadsheet.
 */
function resetAllSheets() {
  const ss = getSpreadsheet_();

  function ensureSheet(name, headers) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return sh;
  }

  ensureSheet(SHEET_PLAYERS, ['player_id','name','email','usual_times']);
  ensureSheet(SHEET_CAMPAIGNS, ['campaign_id','campaign_name','required_players']);
  ensureSheet(SHEET_SLOTS, ['slot_id','start_utc','end_utc']);
  ensureSheet(SHEET_AVAIL, ['player_id','slot_ids','updated_at','updated_by']);
  ensureSheet(SHEET_RESULTS, ['campaign','date','time','required_players']);
  ensureSheet(SHEET_CAMPTIMES, ['campaign_id','campaign','date','time','start_utc','end_utc','status','missing_players','available_players','required_players','slot_ids_merged']);
  ensureSheet(SHEET_LOCKED, ['locked_at','campaign_id','campaign','start_utc','end_utc','slot_ids','locked_by','event_id','guest_emails','availability_snapshot','date_local','time_local']);
  ensureSheet(SHEET_TRACKER, ['player_id','last_saved_at','last_saved_by']);
  invalidateCachedTableByName_(ss, SHEET_PLAYERS);
  invalidateCachedTableByName_(ss, SHEET_CAMPAIGNS);
  invalidateCachedTableByName_(ss, SHEET_SLOTS);

  // Generate slots and derived sheets
  setupSlotsNextWeeks();
  recomputeDerivedSheets_();
  bumpDashboardDataVersion_();
}


function doGet() {
  return runTimed_('doGet', () => {
    const template = HtmlService.createTemplateFromFile('Index');
    template.bootstrapData = JSON.stringify(getBootstrapData_());
    return template.evaluate()
      .setTitle('DnD Scheduler')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  });
}

function getBootstrapData_() {
  const ss = getSpreadsheet_();
  const initialLang = getDefaultClientLang_();
  return {
    initialLang,
    timeZone: DISPLAY_TZ,
    translations: getClientTranslations(initialLang),
    players: toClientPlayers_(readOptimizedTable_(getSheetByNameCached_(ss, SHEET_PLAYERS)))
  };
}

function getSlotsDataKey_() {
  const props = PropertiesService.getScriptProperties();
  const cached = String(props.getProperty(SLOTS_DATA_KEY_PROP) || '').trim();
  if (cached) return cached;

  const ss = getSpreadsheet_();
  const rows = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_SLOTS));
  return refreshSlotsDataKeyFromRows_(rows);
}

/**
 * One-time helper: generates slots for next N weeks.
 * Stores start/end in UTC ISO strings.
 */
function setupSlotsNextWeeks() {
  const weeks = ROLL_WEEKS_AHEAD;
  const slotMinutes = SLOT_MINUTES;
  const startHour = START_HOUR;
  const endHour = END_HOUR;
  const days = weeks * 7;

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_SLOTS);

  sh.clearContents();
  sh.getRange(1,1,1,3).setValues([['slot_id','start_utc','end_utc']]);

  const now = new Date();
  const startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), startHour, 0, 0, 0);

  const rows = [];
  let slotId = 1;

  for (let d = 0; d < days; d++) {
    for (let mins = startHour * 60; mins < endHour * 60; mins += slotMinutes) {
      const localStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + d, 0, mins, 0, 0);
      const localEnd = new Date(localStart.getTime() + slotMinutes * 60 * 1000);

      const startUtc = toUtcIso_(localStart);
      const endUtc = toUtcIso_(localEnd);

      rows.push([slotId, startUtc, endUtc]);
      slotId++;
    }
  }

  sh.getRange(2,1,rows.length,3).setValues(rows);
  refreshSlotsCacheMetadata_(ss, rows);
  bumpDashboardDataVersion_();
}

function updateCalendarEventTimeForLock_(eventId, startMs, endMs) {
  const id = String(eventId || '').trim();
  if (!id || !Number.isFinite(startMs) || !Number.isFinite(endMs)) return false;
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) return false;
  const ev = cal.getEventById(id);
  if (!ev) return false;
  ev.setTime(new Date(startMs), new Date(endMs));
  return true;
}

function getSchedulerTimezoneRepairStatus_() {
  const ss = getSpreadsheet_();
  const rows = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_SLOTS));
  const expectedSlotsPerDay = Math.floor(((END_HOUR - START_HOUR) * 60) / SLOT_MINUTES);
  const expectedFirstLocalTime = Utilities.formatString('%02d:00', START_HOUR);
  const byDay = new Map();

  (rows || []).forEach(r => {
    const iso = toIsoString_(r.start_utc);
    const local = getLocalPartsFromIso_(iso, DISPLAY_TZ);
    if (!local) return;
    if (!byDay.has(local.dayKey)) {
      byDay.set(local.dayKey, {
        dayKey: local.dayKey,
        firstLocalTime: local.timeLabel,
        slotCount: 0
      });
    }
    const entry = byDay.get(local.dayKey);
    entry.slotCount += 1;
    if (local.timeLabel < entry.firstLocalTime) entry.firstLocalTime = local.timeLabel;
  });

  const days = Array.from(byDay.values()).sort((a, b) => String(a.dayKey || '').localeCompare(String(b.dayKey || '')));
  const probe = days.find(day => day.slotCount >= expectedSlotsPerDay) || days[0] || null;
  return {
    needsRepair: !!(probe && probe.firstLocalTime !== expectedFirstLocalTime),
    sampleDay: probe ? probe.dayKey : '',
    sampleLocalStart: probe ? probe.firstLocalTime : '',
    expectedLocalStart: expectedFirstLocalTime,
    sampleSlotCount: probe ? probe.slotCount : 0,
    repairedAt: String(PropertiesService.getScriptProperties().getProperty(SCHEDULER_TZ_REPAIR_PROP) || '')
  };
}

function getSchedulerTimezoneRepairStatus() {
  return getSchedulerTimezoneRepairStatus_();
}

function repairSchedulerTimezoneData(force) {
  const forceRun = force === true || String(force || '').toLowerCase() === 'true';
  const before = getSchedulerTimezoneRepairStatus_();

  const ss = getSpreadsheet_();
  const shSlots = getSheetByNameCached_(ss, SHEET_SLOTS);
  const shLocked = ensureLockedSheet_();
  if (!shSlots) throw new Error('Slots sheet missing.');

  if (!before.needsRepair && !forceRun) {
    const lockedLocalLabelsUpdated = fillLockedLocalWindowColumns_(shLocked);
    return {
      ok: true,
      skipped: true,
      reason: lockedLocalLabelsUpdated > 0
        ? 'No timezone repair needed. LockedTimes local display columns were refreshed.'
        : 'No timezone repair needed.',
      before,
      after: getSchedulerTimezoneRepairStatus_(),
      lockedLocalLabelsUpdated
    };
  }

  const slotRows = readTable(shSlots);
  const correctedSlotValues = [];
  const correctedSlotMap = new Map();
  let slotsUpdated = 0;

  for (const row of slotRows) {
    const slotId = Number(row.slot_id);
    if (!Number.isFinite(slotId) || slotId <= 0) continue;

    const currentStart = toIsoString_(row.start_utc);
    const currentEnd = toIsoString_(row.end_utc);
    const nextStart = convertStoredLocalIsoToTrueUtc_(currentStart, DISPLAY_TZ) || currentStart;
    const nextEnd = convertStoredLocalIsoToTrueUtc_(currentEnd, DISPLAY_TZ) || currentEnd;

    if (currentStart !== nextStart || currentEnd !== nextEnd) slotsUpdated += 1;

    correctedSlotValues.push([slotId, nextStart, nextEnd]);
    correctedSlotMap.set(slotId, {
      start_utc: nextStart,
      end_utc: nextEnd,
      startMs: new Date(nextStart).getTime(),
      endMs: new Date(nextEnd).getTime()
    });
  }

  shSlots.clearContents();
  shSlots.getRange(1, 1, 1, 3).setValues([['slot_id','start_utc','end_utc']]);
  if (correctedSlotValues.length) {
    shSlots.getRange(2, 1, correctedSlotValues.length, 3).setValues(correctedSlotValues);
  }
  refreshSlotsCacheMetadata_(ss, correctedSlotValues);

  let locksUpdated = 0;
  let lockedLocalLabelsUpdated = 0;
  let calendarEventsUpdated = 0;
  let calendarEventsFailed = 0;
  const lockedData = shLocked.getDataRange().getValues();
  if (lockedData.length >= 2) {
    const header = lockedData[0].map(h => String(h).trim());
    const idxStart = header.indexOf('start_utc');
    const idxEnd = header.indexOf('end_utc');
    const idxSlotIds = header.indexOf('slot_ids');
    const idxEventId = header.indexOf('event_id');
    const idxDateLocal = header.indexOf('date_local');
    const idxTimeLocal = header.indexOf('time_local');

    for (let i = 1; i < lockedData.length; i++) {
      const row = lockedData[i];
      let nextStart = idxStart !== -1 ? toIsoString_(row[idxStart]) : '';
      let nextEnd = idxEnd !== -1 ? toIsoString_(row[idxEnd]) : '';
      let nextStartMs = Number.isFinite(new Date(nextStart).getTime()) ? new Date(nextStart).getTime() : null;
      let nextEndMs = Number.isFinite(new Date(nextEnd).getTime()) ? new Date(nextEnd).getTime() : null;

      if (idxSlotIds !== -1) {
        const slotIds = String(row[idxSlotIds] || '')
          .split(',')
          .map(s => Number(String(s).trim()))
          .filter(n => Number.isFinite(n) && n > 0);

        let minStart = null;
        let maxEnd = null;
        slotIds.forEach(id => {
          const slot = correctedSlotMap.get(id);
          if (!slot) return;
          if (minStart === null || slot.startMs < minStart) minStart = slot.startMs;
          if (maxEnd === null || slot.endMs > maxEnd) maxEnd = slot.endMs;
        });

        if (minStart !== null && maxEnd !== null) {
          nextStart = new Date(minStart).toISOString();
          nextEnd = new Date(maxEnd).toISOString();
          nextStartMs = minStart;
          nextEndMs = maxEnd;

          if (idxEventId !== -1) {
            const eventId = String(row[idxEventId] || '').trim();
            if (eventId) {
              try {
                if (updateCalendarEventTimeForLock_(eventId, minStart, maxEnd)) {
                  calendarEventsUpdated += 1;
                }
              } catch (e) {
                calendarEventsFailed += 1;
              }
            }
          }
        }
      }

      if (idxStart !== -1 && idxEnd !== -1) {
        if (String(row[idxStart] || '') !== nextStart || String(row[idxEnd] || '') !== nextEnd) {
          locksUpdated += 1;
        }
        row[idxStart] = nextStart;
        row[idxEnd] = nextEnd;
      }
      if (updateLockedLocalWindowFields_(row, idxStart, idxEnd, idxDateLocal, idxTimeLocal)) {
        lockedLocalLabelsUpdated += 1;
      }
    }

    shLocked.clearContents();
    shLocked.getRange(1, 1, lockedData.length, lockedData[0].length).setValues(lockedData);
  }

  recomputeDerivedSheets_();
  bumpDashboardDataVersion_();

  const repairedAt = new Date().toISOString();
  const props = PropertiesService.getScriptProperties();
  props.setProperty(SCHEDULER_TZ_REPAIR_PROP, repairedAt);

  return {
    ok: true,
    before,
    after: getSchedulerTimezoneRepairStatus_(),
    repairedAt,
    slotsUpdated,
    locksUpdated,
    lockedLocalLabelsUpdated,
    calendarEventsUpdated,
    calendarEventsFailed
  };
}

function rollbackSchedulerTimezoneRepair() {
  const ss = getSpreadsheet_();
  const shSlots = getSheetByNameCached_(ss, SHEET_SLOTS);
  const shLocked = ensureLockedSheet_();
  if (!shSlots) throw new Error('Slots sheet missing.');

  const slotRows = readTable(shSlots);
  const correctedSlotValues = [];
  const correctedSlotMap = new Map();
  let slotsUpdated = 0;

  for (const row of slotRows) {
    const slotId = Number(row.slot_id);
    if (!Number.isFinite(slotId) || slotId <= 0) continue;

    const currentStart = toIsoString_(row.start_utc);
    const currentEnd = toIsoString_(row.end_utc);
    const nextStart = convertTrueUtcToStoredLocalIso_(currentStart, DISPLAY_TZ) || currentStart;
    const nextEnd = convertTrueUtcToStoredLocalIso_(currentEnd, DISPLAY_TZ) || currentEnd;

    if (currentStart !== nextStart || currentEnd !== nextEnd) slotsUpdated += 1;

    correctedSlotValues.push([slotId, nextStart, nextEnd]);
    correctedSlotMap.set(slotId, {
      start_utc: nextStart,
      end_utc: nextEnd,
      startMs: new Date(nextStart).getTime(),
      endMs: new Date(nextEnd).getTime()
    });
  }

  shSlots.clearContents();
  shSlots.getRange(1, 1, 1, 3).setValues([['slot_id','start_utc','end_utc']]);
  if (correctedSlotValues.length) {
    shSlots.getRange(2, 1, correctedSlotValues.length, 3).setValues(correctedSlotValues);
  }
  refreshSlotsCacheMetadata_(ss, correctedSlotValues);

  let locksUpdated = 0;
  let lockedLocalLabelsUpdated = 0;
  let calendarEventsUpdated = 0;
  let calendarEventsFailed = 0;
  const lockedData = shLocked.getDataRange().getValues();
  if (lockedData.length >= 2) {
    const header = lockedData[0].map(h => String(h).trim());
    const idxStart = header.indexOf('start_utc');
    const idxEnd = header.indexOf('end_utc');
    const idxSlotIds = header.indexOf('slot_ids');
    const idxEventId = header.indexOf('event_id');
    const idxDateLocal = header.indexOf('date_local');
    const idxTimeLocal = header.indexOf('time_local');

    for (let i = 1; i < lockedData.length; i++) {
      const row = lockedData[i];
      let nextStart = idxStart !== -1 ? toIsoString_(row[idxStart]) : '';
      let nextEnd = idxEnd !== -1 ? toIsoString_(row[idxEnd]) : '';

      if (idxSlotIds !== -1) {
        const slotIds = String(row[idxSlotIds] || '')
          .split(',')
          .map(s => Number(String(s).trim()))
          .filter(n => Number.isFinite(n) && n > 0);

        let minStart = null;
        let maxEnd = null;
        slotIds.forEach(id => {
          const slot = correctedSlotMap.get(id);
          if (!slot) return;
          if (minStart === null || slot.startMs < minStart) minStart = slot.startMs;
          if (maxEnd === null || slot.endMs > maxEnd) maxEnd = slot.endMs;
        });

        if (minStart !== null && maxEnd !== null) {
          nextStart = new Date(minStart).toISOString();
          nextEnd = new Date(maxEnd).toISOString();

          if (idxEventId !== -1) {
            const eventId = String(row[idxEventId] || '').trim();
            if (eventId) {
              try {
                if (updateCalendarEventTimeForLock_(eventId, minStart, maxEnd)) {
                  calendarEventsUpdated += 1;
                }
              } catch (e) {
                calendarEventsFailed += 1;
              }
            }
          }
        }
      }

      if (idxStart !== -1 && idxEnd !== -1) {
        if (String(row[idxStart] || '') !== nextStart || String(row[idxEnd] || '') !== nextEnd) {
          locksUpdated += 1;
        }
        row[idxStart] = nextStart;
        row[idxEnd] = nextEnd;
      }
      if (updateLockedLocalWindowFields_(row, idxStart, idxEnd, idxDateLocal, idxTimeLocal)) {
        lockedLocalLabelsUpdated += 1;
      }
    }

    shLocked.clearContents();
    shLocked.getRange(1, 1, lockedData.length, lockedData[0].length).setValues(lockedData);
  }

  recomputeDerivedSheets_();
  bumpDashboardDataVersion_();

  const rolledBackAt = new Date().toISOString();
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(SCHEDULER_TZ_REPAIR_PROP);
  props.setProperty(SCHEDULER_TZ_ROLLBACK_PROP, rolledBackAt);

  return {
    ok: true,
    rolledBackAt,
    slotsUpdated,
    locksUpdated,
    lockedLocalLabelsUpdated,
    calendarEventsUpdated,
    calendarEventsFailed
  };
}

function getDashboardSnapshotOptionsForPanels_(options) {
  const opts = options && typeof options === 'object' ? options : {};
  const needsCampaigns = !!opts.includeCampaigns;
  const needsTracker = !!opts.includeTracker;
  const needsWho = !!opts.includeWho;

  return {
    includePlayers: needsCampaigns || needsTracker || needsWho,
    includeCampaigns: needsCampaigns || needsTracker || needsWho,
    includeSlots: needsCampaigns || needsTracker || needsWho,
    includeAvailability: needsCampaigns || needsTracker || needsWho,
    includeLocked: needsCampaigns || needsWho,
    includeTracker: needsTracker
  };
}

function isValidInitDataPayload_(payload, includeSlots) {
  if (!payload || typeof payload !== 'object') return false;
  if (!Array.isArray(payload.selectedSlotIds)) return false;
  if (!Array.isArray(payload.lockedSlotIds)) return false;
  if (!Array.isArray(payload.lockedSlotCampaigns)) return false;
  if (includeSlots && !isValidClientGridData_(payload.gridData)) return false;
  return true;
}

/** Load slot grid state for a single player. */
function getInitData(playerId, options) {
  const pid = String(playerId || '').trim();
  const opts = options && typeof options === 'object' ? options : {};
  if (!pid) {
    return {
      players: [],
      gridData: null,
      selectedSlotIds: [],
      lockedSlotIds: [],
      lockedSlotCampaigns: [],
      slotDataKey: ''
    };
  }

  return runTimed_('getInitData', () => {
    const slotDataKey = getSlotsDataKey_();
    const includeSlots = String(opts.slotDataKey || '').trim() !== slotDataKey;
    const cacheKey = getVersionedCacheKey_('init', JSON.stringify({
      v: INIT_DATA_CACHE_SCHEMA_VERSION,
      pid,
      includeSlots,
      slotDataKey
    }));

    const cached = getCachedJson_(cacheKey);
    if (isValidInitDataPayload_(cached, includeSlots)) return cached;

    const data = getInitDataFromSnapshot_(getDashboardSnapshot_({
      includeSlots,
      includeAvailability: true,
      includeLocked: true
    }), pid, slotDataKey);
    if (!includeSlots) data.gridData = null;
    data.slotDataKey = slotDataKey;
    return putCachedJson_(cacheKey, data);
  }, `player:${pid}`);
}

function getDashboardData(playerId, options) {
  const pid = String(playerId || '').trim();
  const opts = options && typeof options === 'object' ? options : {};
  const snapshotOptions = getDashboardSnapshotOptionsForPanels_(opts);
  snapshotOptions.includePlayers = true;
  if (pid) {
    snapshotOptions.includeSlots = true;
    snapshotOptions.includeAvailability = true;
    snapshotOptions.includeLocked = true;
  }

  const snapshot = getDashboardSnapshot_(snapshotOptions);
  const panels = buildDashboardPanelsFromSnapshot_(snapshot, opts);
  const slotDataKey = pid ? getSlotsDataKey_() : '';
  return {
    init: pid
      ? Object.assign(
          getInitDataFromSnapshot_(snapshot, pid, slotDataKey),
          { slotDataKey }
        )
      : { players: snapshot.players || [] },
    campaignGroups: panels.campaignGroups,
    trackerRows: panels.trackerRows,
    whoGrid: panels.whoGrid
  };
}

function buildDashboardPanelsFromSnapshot_(snapshot, options) {
  const opts = options && typeof options === 'object' ? options : {};
  return {
    campaignGroups: opts.includeCampaigns ? getCampaignSessionOptionsFromSnapshot_(snapshot) : null,
    trackerRows: opts.includeTracker ? getTrackerStatusFromSnapshot_(snapshot) : null,
    whoGrid: opts.includeWho ? getWhoCanPlayGridFromSnapshot_(snapshot, String(opts.whoCampaignRef || '')) : null
  };
}

function buildDashboardPanels_(options) {
  return buildDashboardPanelsFromSnapshot_(
    getDashboardSnapshot_(getDashboardSnapshotOptionsForPanels_(options)),
    options
  );
}

function putDashboardPanelsCaches_(options, panels) {
  const opts = options && typeof options === 'object' ? options : {};
  const data = panels && typeof panels === 'object' ? panels : {};
  if (opts.includeCampaigns && data.campaignGroups !== null && data.campaignGroups !== undefined) {
    putCachedJson_(getVersionedCacheKey_('campaignOptions'), data.campaignGroups);
  }
  if (opts.includeTracker && data.trackerRows !== null && data.trackerRows !== undefined) {
    putCachedJson_(getVersionedCacheKey_('tracker'), data.trackerRows);
  }
  if (opts.includeWho && data.whoGrid !== null && data.whoGrid !== undefined) {
    putCachedJson_(getVersionedCacheKey_('whoGrid', String(opts.whoCampaignRef || '')), data.whoGrid);
  }
}

function getCachedDashboardPanelsFromPanelCaches_(options) {
  const opts = options && typeof options === 'object' ? options : {};
  const ref = String(opts.whoCampaignRef || '');
  const data = {
    campaignGroups: null,
    trackerRows: null,
    whoGrid: null
  };
  const missing = {
    includeCampaigns: false,
    includeTracker: false,
    includeWho: false,
    whoCampaignRef: ref
  };

  if (opts.includeCampaigns) {
    const cachedCampaignGroups = getCachedJson_(getVersionedCacheKey_('campaignOptions'));
    if (cachedCampaignGroups !== null) data.campaignGroups = cachedCampaignGroups;
    else missing.includeCampaigns = true;
  }

  if (opts.includeTracker) {
    const cachedTrackerRows = getCachedJson_(getVersionedCacheKey_('tracker'));
    if (cachedTrackerRows !== null) data.trackerRows = cachedTrackerRows;
    else missing.includeTracker = true;
  }

  if (opts.includeWho) {
    const cachedWhoGrid = getCachedJson_(getVersionedCacheKey_('whoGrid', ref));
    if (cachedWhoGrid !== null) data.whoGrid = cachedWhoGrid;
    else missing.includeWho = true;
  }

  return { data, missing };
}

function hasMissingDashboardPanels_(options) {
  const opts = options && typeof options === 'object' ? options : {};
  return !!opts.includeCampaigns || !!opts.includeTracker || !!opts.includeWho;
}

function getDashboardPanelsData(options) {
  const opts = options && typeof options === 'object' ? options : {};
  return runTimed_('getDashboardPanelsData', () => {
    const key = getVersionedCacheKey_('panels', JSON.stringify({
      includeCampaigns: !!opts.includeCampaigns,
      includeTracker: !!opts.includeTracker,
      includeWho: !!opts.includeWho,
      whoCampaignRef: String(opts.whoCampaignRef || '')
    }));
    const cached = getCachedJson_(key);
    if (cached !== null) return cached;

    const panelState = getCachedDashboardPanelsFromPanelCaches_(opts);
    const data = panelState.data;

    if (hasMissingDashboardPanels_(panelState.missing)) {
      const built = buildDashboardPanels_(panelState.missing);
      if (panelState.missing.includeCampaigns) data.campaignGroups = built.campaignGroups;
      if (panelState.missing.includeTracker) data.trackerRows = built.trackerRows;
      if (panelState.missing.includeWho) data.whoGrid = built.whoGrid;
      putDashboardPanelsCaches_(panelState.missing, built);
    }

    return putCachedJson_(key, data);
  }, [
    `campaigns:${!!opts.includeCampaigns}`,
    `tracker:${!!opts.includeTracker}`,
    `who:${!!opts.includeWho}`
  ].join(' '));
}

function getTrackerStatusFromSnapshot_(snapshot) {
  const players = snapshot.players || [];
  const campaigns = snapshot.campaigns || [];
  const slots = snapshot.slots || [];
  const availability = snapshot.availability || [];
  const trackerRows = snapshot.tracker || [];

  const nowMs = Date.now();

  const members = new Set();
  for (const c of campaigns) {
    const req = String(c.required_players || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);
    req.forEach(pid => members.add(pid));
  }

  const lastById = new Map();
  for (const r of trackerRows) {
    const pid = String(r.player_id || '').trim();
    const last = String(r.last_saved_at || '').trim();
    const savedBy = String(r.last_saved_by || '').trim();
    if (!pid || !last) continue;
    const ms = new Date(last).getTime();
    if (Number.isFinite(ms)) lastById.set(pid, { ms, savedBy });
  }

  const slotEndById = new Map();
  for (const s of slots) {
    const slotId = Number(s.slot_id);
    const endMs = new Date(String(s.end_utc || '')).getTime();
    if (!Number.isFinite(slotId) || !Number.isFinite(endMs)) continue;
    slotEndById.set(slotId, endMs);
  }

  const furthestEndByPlayer = new Map();
  for (const r of availability) {
    const pid = String(r.player_id || '').trim();
    if (!pid) continue;

    const slotIds = parseSlotIds_(r.slot_ids);
    let furthest = null;
    for (const id of slotIds) {
      const endMs = slotEndById.get(id);
      if (!Number.isFinite(endMs)) continue;
      if (furthest === null || endMs > furthest) furthest = endMs;
    }
    if (furthest !== null) furthestEndByPlayer.set(pid, furthest);
  }

  return players
    .map(p => {
      const pid = String(p.player_id || '').trim();
      const name = String(p.name || '').trim() || pid;
      return { pid, name };
    })
    .filter(p => p.pid && members.has(p.pid))
    .map(p => {
      const tracker = lastById.get(p.pid);
      const ms = tracker && tracker.ms;
      const savedBy = tracker && tracker.savedBy ? tracker.savedBy : '';
      const furthestEndMs = furthestEndByPlayer.get(p.pid);
      const furthestSlotEndAt = Number.isFinite(furthestEndMs)
        ? new Date(furthestEndMs).toISOString()
        : '';

      if (!ms) {
        return {
          player_id: p.pid,
          name: p.name,
          status: 'NEEDS UPDATE',
          last_saved_at: '',
          last_saved_by: '',
          days_ago: null,
          furthest_slot_end_at: furthestSlotEndAt
        };
      }

      const ageMs = nowMs - ms;
      const totalHours = Math.floor(ageMs / (60*60*1000));
      const days = Math.floor(totalHours / 24);
      const hours = totalHours % 24;
      const age_display = days === 0 ? `${hours}h` : `${days}d ${hours}h`;

      let status = 'UPDATED';
      if (!Number.isFinite(furthestEndMs)) {
        status = 'NEEDS UPDATE';
      } else {
        const maxAllowedAgeMs = Math.max(0, furthestEndMs - nowMs);
        status = ageMs > maxAllowedAgeMs ? 'NEEDS UPDATE' : 'UPDATED';
      }

      return {
        player_id: p.pid,
        name: p.name,
        status,
        last_saved_at: new Date(ms).toISOString(),
        last_saved_by: savedBy,
        age_display,
        furthest_slot_end_at: furthestSlotEndAt
      };
    });
}

function buildCampaignTimesRowsFromSnapshot_(snapshot) {
  const slots = snapshot.slots || [];
  const campaigns = snapshot.campaigns || [];
  const availMap = buildAvailabilityMapFromRows_(snapshot.availability || []);
  const players = snapshot.players || [];
  const locked = snapshot.locked || [];

  const tz = Session.getScriptTimeZone();

  const playerNameById = new Map();
  for (const p of players) {
    const id = String(p.player_id || '').trim();
    const name = String(p.name || '').trim();
    if (id) playerNameById.set(id, name || id);
  }

  const slotList = [];
  for (const s of slots) {
    const slotId = Number(s.slot_id);
    const start = new Date(String(s.start_utc));
    const end = new Date(String(s.end_utc));
    if (!slotId || isNaN(start.getTime()) || isNaN(end.getTime())) continue;
    slotList.push({ slotId, startMs: start.getTime(), endMs: end.getTime() });
  }
  slotList.sort((a,b)=>a.startMs-b.startMs);

  const lockedSlotToCampaign = new Map();
  const lockedByCampaign = new Map();

  for (const r of locked) {
    const campId = String(r.campaign_id || '').trim();
    const campName = String(r.campaign || '').trim();
    const camp = campId || campName;
    const slotIds = String(r.slot_ids || '')
      .split(',')
      .map(s => Number(String(s).trim()))
      .filter(n => Number.isFinite(n) && n > 0);

    if (!camp || !slotIds.length) continue;

    if (!lockedByCampaign.has(camp)) lockedByCampaign.set(camp, new Set());
    const set = lockedByCampaign.get(camp);

    for (const id of slotIds) {
      lockedSlotToCampaign.set(id, camp);
      set.add(id);
    }
  }

  const campParsed = buildCampaignMeta_(campaigns);
  const rows = [];

  function emitMerged(campaignId, campaignName, reqIds, items) {
    const byDay = new Map();
    for (const it of items) {
      const dayKey = Utilities.formatDate(new Date(it.startMs), tz, 'yyyy-MM-dd');
      if (!byDay.has(dayKey)) byDay.set(dayKey, []);
      byDay.get(dayKey).push(it);
    }

    for (const arr of byDay.values()) {
      arr.sort((a,b)=>a.startMs-b.startMs);

      const merged = [];
      for (const it of arr) {
        if (!merged.length) {
          merged.push({ startMs: it.startMs, endMs: it.endMs, status: it.status, groupKey: it.groupKey, missingNames: it.missingNames || [], availableNames: it.availableNames || [], slotIds:[it.slotId] });
          continue;
        }
        const last = merged[merged.length-1];
        const contiguous = it.startMs === last.endMs;
        const sameGroup = it.status === last.status && it.groupKey === last.groupKey;
        if (contiguous && sameGroup) {
          last.endMs = it.endMs;
          last.slotIds.push(it.slotId);
        } else {
          merged.push({ startMs: it.startMs, endMs: it.endMs, status: it.status, groupKey: it.groupKey, missingNames: it.missingNames || [], availableNames: it.availableNames || [], slotIds:[it.slotId] });
        }
      }

      for (const m of merged) {
        const start = new Date(m.startMs);
        const end = new Date(m.endMs);
        rows.push({
          campaign_id: campaignId,
          campaign: campaignName,
          date: Utilities.formatDate(start, tz, 'dd.MM.yyyy (EEE)'),
          time: `${Utilities.formatDate(start, tz, 'HH:mm')}-${Utilities.formatDate(end, tz, 'HH:mm')}`,
          start_utc: toIsoString_(start),
          end_utc: toIsoString_(end),
          status: m.status,
          missing_players: (m.missingNames || []).join(', '),
          available_players: (m.availableNames || []).join(', '),
          required_players: reqIds.join(', '),
          slot_ids_merged: m.slotIds.join(',')
        });
      }
    }
  }

  for (const camp of campParsed) {
    const items = [];
    const campKey = camp.id || camp.name;
    const lockedSet = lockedByCampaign.get(campKey) || new Set();

    for (const sl of slotList) {
      if (!lockedSet.has(sl.slotId)) continue;
      items.push({ slotId: sl.slotId, startMs: sl.startMs, endMs: sl.endMs, status:'LOCKED', groupKey:'LOCKED', missingNames:[] });
    }

    for (const sl of slotList) {
      const lockedCamp = lockedSlotToCampaign.get(sl.slotId);
      if (lockedCamp) continue;

      const present = availMap.get(sl.slotId) || new Set();
      const availableIds = camp.req.filter(pid => present.has(pid));
      if (!availableIds.length) continue;

      const missingIds = camp.req.filter(pid => !present.has(pid));
      const availableNames = availableIds.map(pid => playerNameById.get(pid) || pid);
      const missingNames = missingIds.map(pid => playerNameById.get(pid) || pid);
      const status = availableIds.length >= camp.playable_threshold ? 'PLAYABLE' : 'MISSING';
      const groupKey = status === 'PLAYABLE'
        ? `PLAYABLE:${missingIds.join(',')}`
        : `MISSING:${missingIds.join(',')}`;

      items.push({ slotId: sl.slotId, startMs: sl.startMs, endMs: sl.endMs, status, groupKey, missingNames, availableNames });
    }

    emitMerged(camp.id, camp.name, camp.req, items);
  }

  const rank = s => s === 'LOCKED' ? 0 : (s === 'PLAYABLE' ? 1 : 2);
  rows.sort((a,b)=>{
    if (a.campaign !== b.campaign) return String(a.campaign || '').localeCompare(String(b.campaign || ''));
    if (a.date !== b.date) return String(a.date || '').localeCompare(String(b.date || ''));
    if (a.time !== b.time) return String(a.time || '').localeCompare(String(b.time || ''));
    return rank(a.status) - rank(b.status);
  });

  return rows;
}

function getCachedCampaignTimesRows_(snapshot) {
  const cacheKey = getVersionedCacheKey_('campaignTimesRows');
  return getOrBuildCachedJson_(cacheKey, () =>
    buildCampaignTimesRowsFromSnapshot_(
      snapshot || getDashboardSnapshot_({
        includePlayers: true,
        includeCampaigns: true,
        includeSlots: true,
        includeAvailability: true,
        includeLocked: true
      })
    )
  );
}

function getCampaignSessionOptionsFromSnapshot_(snapshot) {
  const slots = snapshot.slots || [];
  const campaigns = snapshot.campaigns || [];
  const players = snapshot.players || [];
  const locked = snapshot.locked || [];
  const availMap = buildAvailabilityMapFromRows_(snapshot.availability || []);
  const campParsed = buildCampaignMeta_(campaigns);
  if (!campParsed.length) return [];

  const nowMs = Date.now();
  const minPlayableMs = 60 * 60 * 1000;
  const statusRank = (s) => s === 'LOCKED' ? 0 : (s === 'PLAYABLE' ? 1 : (s === 'ALMOST' ? 2 : 3));
  const playerNameById = new Map();
  for (const p of players) {
    const id = String(p.player_id || '').trim();
    const name = String(p.name || '').trim();
    if (id) playerNameById.set(id, name || id);
  }

  const slotList = [];
  for (const s of slots) {
    const slotId = Number(s.slot_id);
    const start = new Date(String(s.start_utc || ''));
    const end = new Date(String(s.end_utc || ''));
    if (!Number.isFinite(slotId) || isNaN(start.getTime()) || isNaN(end.getTime())) continue;
    const startMs = start.getTime();
    if (startMs < nowMs) continue;
    slotList.push({
      slotId,
      startMs,
      endMs: end.getTime(),
      dayKey: Utilities.formatDate(start, DISPLAY_TZ, 'yyyy-MM-dd')
    });
  }
  if (!slotList.length) return [];
  slotList.sort((a, b) => a.startMs - b.startMs);

  const lockedSlotToCampaign = new Map();
  const lockedByCampaign = new Map();
  for (const r of locked) {
    const campId = String(r.campaign_id || '').trim();
    const campName = String(r.campaign || '').trim();
    const camp = campId || campName;
    const slotIds = String(r.slot_ids || '')
      .split(',')
      .map(s => Number(String(s).trim()))
      .filter(n => Number.isFinite(n) && n > 0);
    if (!camp || !slotIds.length) continue;
    if (!lockedByCampaign.has(camp)) lockedByCampaign.set(camp, new Set());
    const set = lockedByCampaign.get(camp);
    for (const id of slotIds) {
      lockedSlotToCampaign.set(id, camp);
      set.add(id);
    }
  }

  function limitPerStatus(items) {
    const lockedItems = items
      .filter(it => it.status === 'LOCKED')
      .sort((a, b) => a.startMs - b.startMs);

    const playable = items
      .filter(it => it.status === 'PLAYABLE' && (it.endMs - it.startMs) > minPlayableMs)
      .sort((a, b) => a.startMs - b.startMs);

    const almost = items
      .filter(it => it.status === 'ALMOST')
      .sort((a, b) => a.startMs - b.startMs)
      .slice(0, 3);

    const missing = items
      .filter(it => it.status === 'MISSING')
      .sort((a, b) => {
        const ma = Number.isFinite(a.missingCount) ? a.missingCount : 999;
        const mb = Number.isFinite(b.missingCount) ? b.missingCount : 999;
        if (ma !== mb) return ma - mb;
        return a.startMs - b.startMs;
      })
      .slice(0, 3);

    return [...lockedItems, ...playable, ...almost, ...missing]
      .sort((a, b) => {
        const rankDiff = (statusRank(a.status) - statusRank(b.status));
        return rankDiff || (a.startMs - b.startMs);
      });
  }

  const result = [];

  for (const camp of campParsed) {
    const mergedByDay = new Map();
    const campKey = camp.id || camp.name;
    const lockedSet = lockedByCampaign.get(campKey) || new Set();

    for (const sl of slotList) {
      let item = null;

      if (lockedSet.has(sl.slotId)) {
        item = {
          slotId: sl.slotId,
          startMs: sl.startMs,
          endMs: sl.endMs,
          status: 'LOCKED',
          groupKey: 'LOCKED',
          missingNames: [],
          availableNames: []
        };
      } else if (!lockedSlotToCampaign.has(sl.slotId)) {
        const present = availMap.get(sl.slotId) || new Set();
        const availableIds = [];
        const missingIds = [];

        for (const pid of camp.req) {
          if (present.has(pid)) availableIds.push(pid);
          else missingIds.push(pid);
        }

        if (!availableIds.length) continue;

        const availableNames = availableIds.map(pid => playerNameById.get(pid) || pid);
        const missingNames = missingIds.map(pid => playerNameById.get(pid) || pid);
        const status = availableIds.length >= camp.playable_threshold ? 'PLAYABLE' : 'MISSING';
        item = {
          slotId: sl.slotId,
          startMs: sl.startMs,
          endMs: sl.endMs,
          status,
          groupKey: `${status}:${missingIds.join(',')}`,
          missingNames,
          availableNames
        };
      }

      if (!item) continue;

      if (!mergedByDay.has(sl.dayKey)) mergedByDay.set(sl.dayKey, []);
      const merged = mergedByDay.get(sl.dayKey);
      const last = merged[merged.length - 1];

      if (
        last &&
        item.startMs === last.endMs &&
        item.status === last.status &&
        item.groupKey === last.groupKey
      ) {
        last.endMs = item.endMs;
        last.slotIds.push(item.slotId);
      } else {
        merged.push({
          startMs: item.startMs,
          endMs: item.endMs,
          status: item.status,
          groupKey: item.groupKey,
          missingNames: item.missingNames || [],
          availableNames: item.availableNames || [],
          slotIds: [item.slotId]
        });
      }
    }

    const items = [];
    for (const merged of mergedByDay.values()) {
      for (const entry of merged) {
        const startLocal = new Date(entry.startMs);
        const endLocal = new Date(entry.endMs);
        items.push({
          campaign_id: camp.id,
          campaign: camp.name,
          is_oneshot: camp.is_oneshot,
          date: Utilities.formatDate(startLocal, DISPLAY_TZ, 'dd.MM.yyyy (EEE)'),
          time: `${Utilities.formatDate(startLocal, DISPLAY_TZ, 'HH:mm')}-${Utilities.formatDate(endLocal, DISPLAY_TZ, 'HH:mm')}`,
          status: entry.status,
          missing_players: (entry.missingNames || []).join(', '),
          available_players: (entry.availableNames || []).join(', '),
          required_players: camp.req.join(', '),
          slot_ids_merged: entry.slotIds.join(','),
          startMs: entry.startMs,
          endMs: entry.endMs,
          missingCount: (entry.missingNames || []).length,
          _dayKey: Utilities.formatDate(startLocal, DISPLAY_TZ, 'yyyy-MM-dd')
        });
      }
    }

    const missingByDay = new Map();
    items.forEach(it => {
      if (it.status !== 'MISSING' || it.missingCount !== 1) return;
      if (!missingByDay.has(it._dayKey)) missingByDay.set(it._dayKey, []);
      missingByDay.get(it._dayKey).push(it);
    });
    for (const list of missingByDay.values()) {
      list.sort((a, b) => a.startMs - b.startMs);
      const promoteCount = Math.min(3, list.length);
      for (let i = 0; i < promoteCount; i++) {
        list[i].status = 'ALMOST';
      }
    }

    const limited = limitPerStatus(items);
    if (!limited.length) continue;
    limited.forEach(it => {
      delete it.missingCount;
      delete it._dayKey;
    });
    result.push({
      campaign: camp.name,
      is_oneshot: camp.is_oneshot,
      items: limited
    });
  }

  result.sort((a,b) => (a.items[0]?.startMs ?? 0) - (b.items[0]?.startMs ?? 0));
  return result;
}

function getWhoCanPlayGridFromSnapshot_(snapshot, campaignRef) {
  const campaigns = snapshot.campaigns || [];
  const slots = snapshot.slots || [];
  const players = snapshot.players || [];
  const locked = snapshot.locked || [];
  const availMap = buildAvailabilityMapFromRows_(snapshot.availability || []);

  const playerNameById = new Map();
  for (const p of players) {
    const pid = String(p.player_id || '').trim();
    const name = String(p.name || '').trim();
    if (pid) playerNameById.set(pid, name || pid);
  }

  const campParsed = buildCampaignMeta_(campaigns).map(c => ({
    campaign_id: c.id,
    campaign: c.name,
    required_ids: c.req,
    required_count: c.required_count,
    playable_threshold: c.playable_threshold,
    is_oneshot: c.is_oneshot
  })).filter(c => c.campaign && c.required_count > 0);

  if (!campParsed.length) {
    return { campaigns: [], selected_campaign: null, slots: [] };
  }

  const ref = String(campaignRef || '').trim();
  const selected = campParsed.find(c =>
    c.campaign_id === ref || c.campaign === ref
  ) || campParsed[0];

  const requiredPlayers = selected.required_ids.map(pid => ({
    player_id: pid,
    name: playerNameById.get(pid) || pid
  }));

  const lockedSlotToCampaign = new Map();
  for (const r of locked) {
    const camp = String(r.campaign || r.campaign_id || '').trim();
    const ids = String(r.slot_ids || '')
      .split(',')
      .map(s => Number(String(s).trim()))
      .filter(n => Number.isFinite(n) && n > 0);
    if (!camp || !ids.length) continue;
    ids.forEach(id => lockedSlotToCampaign.set(id, camp));
  }

  const nowMs = Date.now();
  const slotRows = slots
    .map(s => {
      const slotId = Number(s.slot_id);
      const start = new Date(String(s.start_utc));
      const end = new Date(String(s.end_utc));
      if (!Number.isFinite(slotId) || isNaN(start.getTime()) || isNaN(end.getTime())) return null;
      return {
        slot_id: slotId,
        start_utc: toIsoString_(s.start_utc),
        end_utc: toIsoString_(s.end_utc),
        startMs: start.getTime()
      };
    })
    .filter(Boolean)
    .filter(s => s.startMs >= nowMs)
    .sort((a,b) => a.startMs - b.startMs);

  const outSlots = slotRows.map(s => {
    const present = availMap.get(s.slot_id) || new Set();
    const availableIds = selected.required_ids.filter(pid => present.has(pid));
    const missingIds = selected.required_ids.filter(pid => !present.has(pid));
    const lockedCamp = lockedSlotToCampaign.get(s.slot_id) || '';

    return {
      slot_id: s.slot_id,
      start_utc: s.start_utc,
      end_utc: s.end_utc,
      available_count: availableIds.length,
      required_count: selected.required_count,
      playable_threshold: selected.playable_threshold,
      available_players: availableIds.map(pid => playerNameById.get(pid) || pid),
      missing_players: missingIds.map(pid => playerNameById.get(pid) || pid),
      is_locked: !!lockedCamp,
      locked_campaign: lockedCamp
    };
  });

  return {
    campaigns: campParsed.map(c => ({
      campaign_id: c.campaign_id,
      campaign: c.campaign,
      required_count: c.required_count,
      playable_threshold: c.playable_threshold,
      is_oneshot: c.is_oneshot
    })),
    selected_campaign: {
      campaign_id: selected.campaign_id,
      campaign: selected.campaign,
      required_count: selected.required_count,
      playable_threshold: selected.playable_threshold,
      is_oneshot: selected.is_oneshot,
      required_players: requiredPlayers
    },
    slots: outSlots
  };
}

function ensureTrackerSheet_() {
  return ensureTrackerSheetInSpreadsheet_(getSpreadsheet_());
}

function updateTrackerOnSave_(playerId) {
  const sh = ensureTrackerSheet_();
  const nowIso = new Date().toISOString();
  const who = getActorEmail_();

  const data = sh.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());
  const idxPid = header.indexOf('player_id');
  const idxLast = header.indexOf('last_saved_at');
  const idxWho = header.indexOf('last_saved_by');

  // Find row
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxPid]) === String(playerId)) { rowIndex = i + 1; break; }
  }

  if (rowIndex === -1) {
    appendRowsWithSetValues_(sh, [[String(playerId), nowIso, who]]);
  } else {
    const row = data[rowIndex - 1].slice();
    if (idxLast !== -1) row[idxLast] = nowIso;
    if (idxWho !== -1) row[idxWho] = who;
    setSingleRowValues_(sh, rowIndex, row);
  }
}

function toIsoString_(v) {
  if (v instanceof Date) return v.toISOString();
  const s = String(v || '').trim();
  if (!s) return '';
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString();
  return s;
}

/** Save selected slots for a player (overwrite their previous entries). */
function saveAvailability(playerId, selectedSlotIds, options) {
  const sh = ensureAvailabilitySheet_();

  const data = sh.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());
  const idxPid = header.indexOf('player_id');
  const idxSlots = header.indexOf('slot_ids');
  const idxUpdatedAt = header.indexOf('updated_at');
  const idxUpdatedBy = header.indexOf('updated_by');

  if (idxPid === -1 || idxSlots === -1) {
    throw new Error('Availability sheet header is invalid.');
  }

  // Filter selection (exclude locked slots to prevent stale client writes)
  const locked = getLockedSlotIds_();
  const slotIds = Array.from(new Set((selectedSlotIds || [])
    .map(id => Number(id))
    .filter(n => Number.isFinite(n) && !locked.has(n))))
    .sort((a,b)=>a-b);

  const nowIso = new Date().toISOString();
  const who = getActorEmail_();

  // Find row
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxPid]) === String(playerId)) { rowIndex = i + 1; break; }
  }

  const slotCsv = slotIds.join(',');
  if (rowIndex === -1) {
    const newRow = new Array(header.length).fill('');
    newRow[idxPid] = String(playerId);
    newRow[idxSlots] = slotCsv;
    if (idxUpdatedAt !== -1) newRow[idxUpdatedAt] = nowIso;
    if (idxUpdatedBy !== -1) newRow[idxUpdatedBy] = who;
    appendRowsWithSetValues_(sh, [newRow]);
  } else {
    const row = data[rowIndex - 1].slice();
    row[idxSlots] = slotCsv;
    if (idxUpdatedAt !== -1) row[idxUpdatedAt] = nowIso;
    if (idxUpdatedBy !== -1) row[idxUpdatedBy] = who;
    setSingleRowValues_(sh, rowIndex, row);
  }

  updateTrackerOnSave_(playerId);
  bumpDashboardDataVersion_();
  const panels = buildDashboardPanels_(options);
  const opts = options && typeof options === 'object' ? options : {};
  putDashboardPanelsCaches_(opts, panels);

  return {
    ok: true,
    campaignGroups: panels.campaignGroups,
    trackerRows: panels.trackerRows,
    whoGrid: panels.whoGrid
  };
}

function getFreshDerivedSheetsSnapshot_(ss) {
  const spreadsheet = ss || getSpreadsheet_();
  return {
    players: readTable(getSheetByNameCached_(spreadsheet, SHEET_PLAYERS)),
    campaigns: readTable(getSheetByNameCached_(spreadsheet, SHEET_CAMPAIGNS)),
    slots: readTable(getSheetByNameCached_(spreadsheet, SHEET_SLOTS)),
    availability: readTable(ensureAvailabilitySheetInSpreadsheet_(spreadsheet)),
    locked: readTable(ensureLockedSheetInSpreadsheet_(spreadsheet))
  };
}

function writeTableSheet_(sheet, rows, options) {
  const opts = options && typeof options === 'object' ? options : {};
  if (!sheet || !rows || !rows.length || !rows[0] || !rows[0].length) return;
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  if (opts.autoResize) {
    sheet.autoResizeColumns(1, rows[0].length);
  }
}

function buildResultsSheetRowsFromSnapshot_(snapshot) {
  const slots = snapshot && Array.isArray(snapshot.slots) ? snapshot.slots : [];
  const campaigns = snapshot && Array.isArray(snapshot.campaigns) ? snapshot.campaigns : [];
  const availability = snapshot && Array.isArray(snapshot.availability) ? snapshot.availability : [];
  const availMap = buildAvailabilityMapFromRows_(availability);
  const tz = Session.getScriptTimeZone();

  const slotMap = new Map();
  for (const s of slots) {
    const start = new Date(String(s.start_utc));
    const end = new Date(String(s.end_utc));
    const slotId = Number(s.slot_id);
    if (!slotId || isNaN(start) || isNaN(end)) continue;
    slotMap.set(slotId, { startMs: start.getTime(), endMs: end.getTime() });
  }

  const campParsed = campaigns.map(c => {
    const req = String(c.required_players || '').split(',').map(s => s.trim()).filter(Boolean);
    return { name: c.campaign_name || c.campaign_id, req };
  }).filter(c => c.req.length);

  const out = [['campaign', 'date', 'time', 'required_players']];
  const EPS = 60 * 1000;

  for (const camp of campParsed) {
    const intervals = [];
    for (const [slotId, present] of availMap.entries()) {
      let ok = true;
      for (const pid of camp.req) {
        if (!present.has(pid)) { ok = false; break; }
      }
      if (!ok) continue;
      const slot = slotMap.get(slotId);
      if (!slot) continue;
      intervals.push({ startMs: slot.startMs, endMs: slot.endMs });
    }

    intervals.sort((a,b)=>a.startMs-b.startMs);

    const byDate = new Map();
    for (const it of intervals) {
      const dKey = Utilities.formatDate(new Date(it.startMs), tz, 'yyyy-MM-dd');
      if (!byDate.has(dKey)) byDate.set(dKey, []);
      byDate.get(dKey).push(it);
    }

    for (const dayIntervals of byDate.values()) {
      dayIntervals.sort((a,b)=>a.startMs-b.startMs);
      const merged = [];
      for (const it of dayIntervals) {
        if (!merged.length) { merged.push({ ...it }); continue; }
        const last = merged[merged.length - 1];
        if (it.startMs <= last.endMs + EPS) last.endMs = Math.max(last.endMs, it.endMs);
        else merged.push({ ...it });
      }

      for (const interval of merged) {
        const start = new Date(interval.startMs);
        const end = new Date(interval.endMs);
        const dateLabel = Utilities.formatDate(start, tz, 'dd.MM.yyyy (EEE)');
        const timeLabel = `${Utilities.formatDate(start, tz, 'HH:mm')}-${Utilities.formatDate(end, tz, 'HH:mm')}`;
        out.push([camp.name, dateLabel, timeLabel, camp.req.join(', ')]);
      }
    }
  }

  return out;
}

function buildCampaignTimesSheetRowsFromSnapshot_(snapshot) {
  const rows = buildCampaignTimesRowsFromSnapshot_(snapshot || {});
  const out = [['campaign_id','campaign','date','time','start_utc','end_utc','status','missing_players','available_players','required_players','slot_ids_merged']];
  rows.forEach(r => out.push([
    r.campaign_id,
    r.campaign,
    r.date,
    r.time,
    r.start_utc,
    r.end_utc,
    r.status,
    r.missing_players,
    r.available_players,
    r.required_players,
    r.slot_ids_merged
  ]));
  return out;
}

function recomputeResultsFromSnapshot_(ss, snapshot) {
  const spreadsheet = ss || getSpreadsheet_();
  const sh = getSheetByNameCached_(spreadsheet, SHEET_RESULTS) || spreadsheet.insertSheet(SHEET_RESULTS);
  writeTableSheet_(sh, buildResultsSheetRowsFromSnapshot_(snapshot));
}

function recomputeCampaignTimesFromSnapshot_(ss, snapshot) {
  const spreadsheet = ss || getSpreadsheet_();
  const shOut = getSheetByNameCached_(spreadsheet, SHEET_CAMPTIMES) || spreadsheet.insertSheet(SHEET_CAMPTIMES);
  writeTableSheet_(shOut, buildCampaignTimesSheetRowsFromSnapshot_(snapshot));
}

function recomputeDerivedSheets_() {
  const ss = getSpreadsheet_();
  const snapshot = getFreshDerivedSheetsSnapshot_(ss);
  recomputeResultsFromSnapshot_(ss, snapshot);
  recomputeCampaignTimesFromSnapshot_(ss, snapshot);
}

/** Build Results sheet: per campaign, merged playable windows (PLAYABLE only). */
function recomputeResults() {
  const ss = getSpreadsheet_();
  recomputeResultsFromSnapshot_(ss, getFreshDerivedSheetsSnapshot_(ss));
}

function recomputeCampaignTimes() {
  const ss = getSpreadsheet_();
  recomputeCampaignTimesFromSnapshot_(ss, getFreshDerivedSheetsSnapshot_(ss));
}

/** Lock (P04 only). Prevent double-lock. Writes to LockedTimes. */
function lockCampaignTime(adminPlayerId, campaignName, slotIdsCsv) {
  if (String(adminPlayerId) !== ADMIN_LOCKER_ID) {
    throw new Error('Not authorized to lock times.');
  }

  const ss = getSpreadsheet_();
  const shLocked = ensureLockedSheetInSpreadsheet_(ss);
  const shSlots = getSheetByNameCached_(ss, SHEET_SLOTS);
  const campaignNameText = String(campaignName || '').trim();
  const campaigns = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_CAMPAIGNS));
  const players = readOptimizedTable_(getSheetByNameCached_(ss, SHEET_PLAYERS));

  const slotIds = String(slotIdsCsv || '')
    .split(',')
    .map(s => Number(String(s).trim()))
    .filter(n => Number.isFinite(n) && n > 0);
  const availabilitySnapshot = JSON.stringify(buildAvailabilitySnapshotForSlots_(slotIds));

  if (slotIds.length === 0) throw new Error('No slot IDs provided.');

  // Prevent double-lock
  const existing = shLocked.getDataRange().getValues();
  if (existing.length >= 2) {
    const header = existing[0].map(h => String(h).trim());
    const idxSlotIds = header.indexOf('slot_ids');
    if (idxSlotIds !== -1) {
      const lockedSet = new Set();
      for (let i = 1; i < existing.length; i++) {
        const csv = String(existing[i][idxSlotIds] || '');
        csv.split(',').forEach(x => {
          const n = Number(String(x).trim());
          if (Number.isFinite(n) && n > 0) lockedSet.add(n);
        });
      }
      const dup = slotIds.find(id => lockedSet.has(id));
      if (dup) throw new Error('This time was already locked. Refresh and choose another.');
    }
  }

  // Compute start/end for stored record
  const slots = readTable(shSlots);
  const slotMap = new Map();
  for (const s of slots) {
    slotMap.set(Number(s.slot_id), {
      start: new Date(String(s.start_utc)).getTime(),
      end: new Date(String(s.end_utc)).getTime()
    });
  }

  let minStart = null;
  let maxEnd = null;
  for (const id of slotIds) {
    const it = slotMap.get(id);
    if (!it) continue;
    if (minStart === null || it.start < minStart) minStart = it.start;
    if (maxEnd === null || it.end > maxEnd) maxEnd = it.end;
  }
  if (minStart === null || maxEnd === null) throw new Error('Could not resolve slot times.');

  // Resolve campaign id and required players -> emails
  const campRow = getCampaignRowFromRows_(campaigns, campaignNameText);
  const campaignId = campRow ? String(campRow.campaign_id || '').trim() : '';
  let eventId = '';
  let guestEmails = [];
  let inviteError = '';
  try {
    const reqIds = getCampaignRequiredPlayerIdsFromRow_(campRow);
    guestEmails = getPlayerEmailsByIdsFromRows_(players, reqIds);
    // Create calendar event (invites sent)
    eventId = createCalendarEventForLock_(campaignNameText, minStart, maxEnd, guestEmails);
  } catch (e) {
    // Do not block locking if calendar invite fails
    inviteError = e && e.message ? e.message : String(e);
    eventId = '';
  }

  const localWindow = getLocalWindowLabels_(minStart, maxEnd, DISPLAY_TZ);

  appendRowsWithSetValues_(shLocked, [[
    new Date().toISOString(),
    campaignId,
    campaignNameText,
    new Date(minStart).toISOString(),
    new Date(maxEnd).toISOString(),
    slotIds.join(','),
    ADMIN_LOCKER_ID,
    eventId,
    guestEmails.join(','),
    availabilitySnapshot,
    localWindow.date_local,
    localWindow.time_local
  ]]);


  recomputeDerivedSheets_();
  bumpDashboardDataVersion_();

  return { ok: true, eventId, guests: guestEmails, invite_error: inviteError };
}

/** Unlock (P04 only). Removes a lock record so slots become available again. */
function unlockCampaignTime(adminPlayerId, campaignName, slotIdsCsv) {
  if (String(adminPlayerId) !== ADMIN_LOCKER_ID) {
    throw new Error('Not authorized to unlock times.');
  }

  const shLocked = ensureLockedSheet_();
  const data = shLocked.getDataRange().getValues();
  if (data.length < 2) throw new Error('No locked times found.');

  const header = data[0].map(h => String(h).trim());
  const idxCampaign = header.indexOf('campaign');
  const idxCampaignId = header.indexOf('campaign_id');
  const idxSlotIds = header.indexOf('slot_ids');
  const idxEventId = header.indexOf('event_id');
  const idxSnapshot = header.indexOf('availability_snapshot');

  if (idxSlotIds === -1) {
    throw new Error('LockedTimes sheet header is invalid.');
  }

  const wantedCampaign = String(campaignName || '').trim();
  const wantedSlotCsv = String(slotIdsCsv || '')
    .split(',')
    .map(s => Number(String(s).trim()))
    .filter(n => Number.isFinite(n) && n > 0)
    .sort((a,b)=>a-b)
    .join(',');

  if (!wantedSlotCsv) throw new Error('No slot IDs provided.');

  const wantedCampaignRow = getCampaignByName_(wantedCampaign);
  const wantedCampaignId = wantedCampaignRow ? String(wantedCampaignRow.campaign_id || '').trim() : '';

  let removed = false;
  let removedEventId = '';
  let removedCalendarEvent = false;
  let removedSnapshot = null;
  const out = [data[0]];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowCampaign = idxCampaign !== -1 ? String(row[idxCampaign] || '').trim() : '';
    const rowCampaignId = idxCampaignId !== -1 ? String(row[idxCampaignId] || '').trim() : '';
    const rowSlotCsv = String(row[idxSlotIds] || '')
      .split(',')
      .map(s => Number(String(s).trim()))
      .filter(n => Number.isFinite(n) && n > 0)
      .sort((a,b)=>a-b)
      .join(',');

    const campaignMatch = (rowCampaign && rowCampaign === wantedCampaign) ||
      (wantedCampaignId && rowCampaignId === wantedCampaignId);
    const slotMatch = rowSlotCsv === wantedSlotCsv;

    if (!removed && campaignMatch && slotMatch) {
      removed = true;
      removedEventId = idxEventId !== -1 ? String(row[idxEventId] || '').trim() : '';
      if (idxSnapshot !== -1) {
        const rawSnapshot = String(row[idxSnapshot] || '').trim();
        if (rawSnapshot) {
          try {
            removedSnapshot = JSON.parse(rawSnapshot);
          } catch (e) {
            removedSnapshot = null;
          }
        }
      }
      continue;
    }

    out.push(row);
  }

  if (!removed) {
    throw new Error('Lock not found. Refresh and try again.');
  }

  shLocked.clearContents();
  shLocked.getRange(1,1,out.length,out[0].length).setValues(out);

  if (removedEventId) {
    try {
      removedCalendarEvent = deleteCalendarEventForLock_(removedEventId);
    } catch (e) {}
  }

  restoreAvailabilitySnapshot_(removedSnapshot);

  recomputeDerivedSheets_();
  bumpDashboardDataVersion_();

  return { ok: true, eventId: removedEventId, calendar_event_removed: removedCalendarEvent };
}

function compactAvailabilitySheet() {
  const ss = getSpreadsheet_();
  const sh = getSheetByNameCached_(ss, SHEET_AVAIL) || ss.insertSheet(SHEET_AVAIL);
  const beforeRows = Math.max(0, sh.getLastRow() - 1);
  const beforeHeader = sh.getLastRow()
    ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim())
    : [];

  const normalized = ensureAvailabilitySheetInSpreadsheet_(ss);
  const afterRows = Math.max(0, normalized.getLastRow() - 1);
  const afterHeader = normalized.getLastRow()
    ? normalized.getRange(1, 1, 1, normalized.getLastColumn()).getValues()[0].map(h => String(h).trim())
    : [];

  bumpDashboardDataVersion_();
  return { beforeRows, afterRows, beforeHeader, afterHeader };
}

/** Utility: read sheet into array of objects using header row. */
function readTable(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = values[i][j];
    const any = Object.values(obj).some(v => String(v).trim() !== '');
    if (any) rows.push(obj);
  }
  return rows;
}

function getTrackerStatus() {
  return getOrBuildCachedJson_(getVersionedCacheKey_('tracker'), () =>
    getTrackerStatusFromSnapshot_(getDashboardSnapshot_({
      includePlayers: true,
      includeCampaigns: true,
      includeSlots: true,
      includeAvailability: true,
      includeTracker: true
    }))
  );
}

function buildCampaignTimesRows_() {
  return buildCampaignTimesRowsFromSnapshot_(getDashboardSnapshot_({
    includePlayers: true,
    includeCampaigns: true,
    includeSlots: true,
    includeAvailability: true,
    includeLocked: true
  }));
}

function getCampaignSessionOptions() {
  return getOrBuildCachedJson_(getVersionedCacheKey_('campaignOptions'), () =>
    getCampaignSessionOptionsFromSnapshot_(getDashboardSnapshot_({
      includePlayers: true,
      includeCampaigns: true,
      includeSlots: true,
      includeAvailability: true,
      includeLocked: true
    }))
  );
}

function getWhoCanPlayGrid(campaignRef) {
  const ref = String(campaignRef || '').trim();
  return getOrBuildCachedJson_(getVersionedCacheKey_('whoGrid', ref), () =>
    getWhoCanPlayGridFromSnapshot_(getDashboardSnapshot_({
      includePlayers: true,
      includeCampaigns: true,
      includeSlots: true,
      includeAvailability: true,
      includeLocked: true
    }), ref)
  );
}
