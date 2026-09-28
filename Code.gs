// ============================================================================
// SUPPORT HUB - PHASE 13H.6.1 TASK PARENT CLARITY RELEASE
// Google Sheets + Apps Script internal workflow for Questions, Answered Tickets,
// Tasks, Peer Questions, notifications, reporting, audit history, and Team Setup.
// Existing sheet columns are preserved positionally; newer features use additive
// columns/sheets and idempotent schema checks so historical data remains intact.
// ============================================================================

// ==========================================
// GLOBALS & CONSTANTS
// ==========================================
const SHEET_QUESTIONS  = "Questions Tracker";
const SHEET_ANSWERED   = "Answered";
const SHEET_TEAM       = "Team Setup";
const SHEET_AUDIT      = "Audit Log";
const SHEET_DELETED    = "Deleted Records";
const SHEET_NOTIFICATIONS = "Notification Preferences";
const SHEET_MENTION_NOTIFICATIONS = "Mention Notifications";
const SHEET_TASK_NOTIFICATIONS = "Task Notifications";
const SHEET_TASKS = "Tasks";
const SHEET_PEER_QUESTIONS = "Peer Questions";
const SHEET_WORK_CATEGORIES = "Work Categories";
const SHEET_CATEGORY_SUGGESTIONS = "Category Suggestions";
const SHEET_PERFORMANCE = "Performance Log";
const APP_RELEASE = "13H.6.8-PEER";
const PERFORMANCE_SLOW_MS = 5000;

// Question duplicates are checked by normalized Question + Event + Case Link
// across both open and answered tickets. This small time constant is retained
// only for older callers/tests that reference it; question safety no longer
// depends on a short retry window.
const DUPLICATE_WINDOW_MS = 10 * 60 * 1000;

// How long a mutation will wait to acquire the write lock before giving up.
const LOCK_WAIT_MS = 60000;
const SLOW_LOCK_LOG_MS = 1500;

// Lightweight runtime caches. These only avoid repeated Sheet/schema reads;
// all mutations still use the same ScriptLock and server-side validation.
const SETUP_CACHE_KEY = 'SUPPORT_HUB_SCHEMA_READY_ATTENTION_TODAY_V1';
const SETUP_CACHE_SECONDS = 300;
// A successful migration is durable across Apps Script cold starts. The old
// cache-only guard expired every five minutes and made a regular user recheck
// every sheet header. Bump this value only when a future release changes the
// spreadsheet schema, then run runManualSetup() once before deployment.
const SCHEMA_VERSION_KEY = 'SUPPORT_HUB_SCHEMA_VERSION';
const SCHEMA_VERSION = 'PROFESSIONAL_PERFORMANCE_RC1_TEAM_FLAG_V1_TASK_ATTENTION_V1';
// Team routing must never reuse a roster cached by an older release. The
// generation suffix also prevents an in-flight read from restoring stale
// Primary/Backup values after Admin saves a newer roster.
const TEAM_CACHE_KEY = 'SUPPORT_HUB_TEAM_ROSTER_V2_ROUTING';
const TEAM_CACHE_VERSION_KEY = 'SUPPORT_HUB_TEAM_CACHE_GENERATION_V1';
const TEAM_CACHE_SECONDS = 60;
const WORK_CATEGORY_CACHE_KEY = 'SUPPORT_HUB_WORK_CATEGORIES_V1';
const WORK_CATEGORY_CACHE_SECONDS = 120;
const ROW_INDEX_CACHE_SECONDS = 21600;
const DUPLICATE_MATCH_CACHE_SECONDS = 21600;

// ---- Column maps (1-indexed, matches getRange) ----
const Q_COL = {
  QUESTION: 1, EVENT: 2, ASKED_BY: 3, CREATED: 4, DUE: 5, PRIORITY: 6,
  HOURS_LEFT: 7, LINK: 8, ANSWER: 9, STATUS: 10, ASSIGNED: 11,
  TICKET_ID: 12, HOLD_REASON: 13, HOLD_SINCE: 14, HOLD_ACCUM: 15, ASKED_BY_EMAIL: 16,
  FOLLOWUP: 17, ASSIGN_NOTES: 18, STATUS_CHANGED_AT: 19, STATUS_CHANGED_BY: 20,
  WORK_CATEGORY: 21, EVENT_THIS_WEEKEND: 22, EVENT_DATE: 23,
  TALENT_ROLE: 24, ROUTED_SUPPORT_EMAIL: 25,
  ATTENTION_TODAY: 26, ATTENTION_REASON: 27, ATTENTION_SET_AT: 28,
  ATTENTION_SET_BY: 29, ATTENTION_UNTIL: 30,
  ATTENTION_CLEARED_AT: 31, ATTENTION_CLEARED_BY: 32,
  ATTENTION_FIRST_REVIEW_AT: 33, ATTENTION_FIRST_REVIEW_BY: 34,
  REQUESTER_FLAGGED: 35, REQUESTER_FLAGGED_AT: 36, REQUESTER_FLAGGED_BY: 37,
  REQUESTER_FLAG_REASON: 38, REQUESTER_FLAG_ACK_AT: 39, REQUESTER_FLAG_ACK_BY: 40
};
const Q_WIDTH = 40;

const A_COL = {
  QUESTION: 1, EVENT: 2, ASKED_BY: 3, CREATED: 4, ANSWERED: 5, PRIORITY: 6,
  TURNAROUND: 7, LINK: 8, ANSWER: 9, STATUS: 10, ANSWERED_BY: 11, READ_BY: 12,
  TICKET_ID: 13, HOLD_HOURS: 14, ASKED_BY_EMAIL: 15, FIRST_READ_AT: 16, LAST_UPDATED_AT: 17,
  WORK_CATEGORY: 18, EVENT_THIS_WEEKEND: 19, EVENT_DATE: 20,
  TALENT_ROLE: 21, ROUTED_SUPPORT_EMAIL: 22, ANSWERED_BY_EMAIL: 23,
  ATTENTION_REQUESTED: 24, ATTENTION_REASON: 25, ATTENTION_SET_AT: 26,
  ATTENTION_SET_BY: 27, ATTENTION_UNTIL: 28,
  ATTENTION_CLEARED_AT: 29, ATTENTION_CLEARED_BY: 30,
  ATTENTION_FIRST_REVIEW_AT: 31, ATTENTION_FIRST_REVIEW_BY: 32
};
const A_WIDTH = 32;

const MAX_ATTENTION_REASON_CHARS = 500;

const EVENT_ROLE_OPTIONS = [
  'Lead Photographer',
  'Associate Photographer',
  'Lead Videographer',
  'Associate Videographer',
  'Client 1',
  'Client 2'
];

const TEAM_COL = {
  NAME: 1, TITLE: 2, STATUS: 3, CATEGORY: 4, EMAIL: 5,
  PASSWORD_HASH: 6, PASSWORD_SALT: 7, PRIMARY_SUPPORT_EMAIL: 8, BACKUP_SUPPORT_EMAIL: 9
};
const TEAM_WIDTH = 9;

// ---- Task columns (Phase 1: independent task workflow) ----
const T_COL = {
  TASK_ID: 1, TITLE: 2, INSTRUCTIONS: 3, ASSIGNED_TO: 4, ASSIGNED_TO_EMAIL: 5,
  CREATED_BY: 6, CREATED_BY_EMAIL: 7, CREATED_AT: 8, DUE_DATE: 9, PRIORITY: 10,
  STATUS: 11, COMPLETED_AT: 12, COMPLETION_NOTE: 13, SOURCE: 14,
  RELATED_TICKET_ID: 15, RELATED_EVENT: 16, UPDATED_AT: 17, UPDATED_BY: 18,
  WORK_CATEGORY: 19, RELATED_CLIENT_TALENT: 20, RELATED_LINK: 21,
  PARENT_TYPE: 22, PARENT_ID: 23, SUPPORT_OWNER_EMAIL: 24,
  COMPLETED_BY_EMAIL: 25, COMPLETED_BY: 26, RELATED_ENTITY_TYPE: 27,
  ATTENTION_TODAY: 28, ATTENTION_SET_AT: 29, ATTENTION_SET_BY: 30, ATTENTION_UNTIL: 31
};
const T_WIDTH = 31;
const TASK_STATUS_PENDING = 'Pending';
const TASK_STATUS_IN_PROGRESS = 'In Progress';
const TASK_STATUS_COMPLETED = 'Completed';
const TASK_STATUS_CANCELLED = 'Cancelled';
const CONFIG_KEY_TASK_DATA_VERSION = 'TASK_DATA_VERSION_COUNTER';
const TASK_DUPLICATE_WINDOW_MS = 15000;
// Google Sheets cells are limited to 50,000 characters. Leave a little room
// for future history labels instead of allowing a late, unclear append error.
const MAX_SHEET_CELL_CHARS = 48000;
const ATTACHMENT_FOLDER_NAME = 'Support Hub Attachments';
// V2 uses one project-wide folder. The earlier UserProperties key created a
// separate folder context for each signed-in user and then shared every image
// individually, which generated Drive "shared with you" notifications.
const ATTACHMENT_FOLDER_SCRIPT_PROP = 'SUPPORT_HUB_ATTACHMENT_FOLDER_ID_V2';
const ATTACHMENT_ACCESS_SIGNATURE_PROP = 'SUPPORT_HUB_ATTACHMENT_ACCESS_V2';
const ATTACHMENT_FOLDER_USER_PROP = 'SUPPORT_HUB_ATTACHMENT_FOLDER_ID';

function requireSheetCellLength(value, label) {
  if (String(value == null ? '' : value).length > MAX_SHEET_CELL_CHARS) {
    throw new Error((label || 'This entry') + ' is too large to save. Please remove an image or shorten the text and try again.');
  }
}

function rejectEmbeddedBase64Image(value, label) {
  if (/data:image\/(?:png|jpe?g|gif|webp);base64,/i.test(String(value || ''))) {
    throw new Error((label || 'This entry') + ' contains an embedded image that was not uploaded. Paste or attach the image again and wait for the upload to finish.');
  }
}

function safeAttachmentFileName(name, mimeType) {
  let clean = String(name || 'screenshot').trim().replace(/[\\/:*?"<>|\x00-\x1F]/g, '_').replace(/\s+/g, ' ');
  if (!clean) clean = 'screenshot';
  if (clean.length > 120) clean = clean.slice(0, 120);
  const extensions = { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/gif': '.gif', 'image/webp': '.webp' };
  const wanted = extensions[mimeType] || '';
  if (wanted && !/\.(png|jpe?g|gif|webp)$/i.test(clean)) clean += wanted;
  return clean;
}

function getSupportHubAttachmentFolder_() {
  const scriptProps = PropertiesService.getScriptProperties();
  const savedId = String(scriptProps.getProperty(ATTACHMENT_FOLDER_SCRIPT_PROP) || '').trim();
  if (savedId) {
    try {
      const saved = DriveApp.getFolderById(savedId);
      if (!saved.isTrashed()) return saved;
    } catch (e) {}
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const recheckId = String(scriptProps.getProperty(ATTACHMENT_FOLDER_SCRIPT_PROP) || '').trim();
    if (recheckId) {
      try {
        const recheck = DriveApp.getFolderById(recheckId);
        if (!recheck.isTrashed()) return recheck;
      } catch (e) {}
    }

    // Reuse a valid legacy folder where possible so existing administrators do
    // not end up with another folder of the same name after this upgrade.
    const legacyProps = PropertiesService.getUserProperties();
    const legacyId = String(legacyProps.getProperty(ATTACHMENT_FOLDER_USER_PROP) || '').trim();
    if (legacyId) {
      try {
        const legacyFolder = DriveApp.getFolderById(legacyId);
        if (!legacyFolder.isTrashed()) {
          scriptProps.setProperty(ATTACHMENT_FOLDER_SCRIPT_PROP, legacyFolder.getId());
          return legacyFolder;
        }
      } catch (e) {}
    }

    const matches = DriveApp.getFoldersByName(ATTACHMENT_FOLDER_NAME);
    if (matches.hasNext()) {
      const existing = matches.next();
      scriptProps.setProperty(ATTACHMENT_FOLDER_SCRIPT_PROP, existing.getId());
      return existing;
    }
    const folder = DriveApp.createFolder(ATTACHMENT_FOLDER_NAME);
    scriptProps.setProperty(ATTACHMENT_FOLDER_SCRIPT_PROP, folder.getId());
    return folder;
  } finally {
    lock.releaseLock();
  }
}

function attachmentFolderPermissionRole_() {
  // If the web app executes as its owner, team members only need inherited read
  // access. If it executes as the accessing user, they need writer access to
  // place uploads in the common folder. This keeps both deployment modes usable.
  try {
    const active = normalizeEmail(Session.getActiveUser().getEmail());
    const effective = normalizeEmail(Session.getEffectiveUser().getEmail());
    if (active && effective && emailsRepresentSameWorkspaceIdentity(active, effective)) return 'writer';
  } catch (e) {}
  return 'reader';
}

function drivePermissionRequest_(fileId, email, role) {
  const targetId = String(fileId || '').trim();
  const targetEmail = normalizeEmail(email);
  if (!targetId || !targetEmail) return null;
  return {
    url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(targetId) +
      '/permissions?supportsAllDrives=true&sendNotificationEmail=false',
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ type: 'user', role: role || 'reader', emailAddress: targetEmail }),
    muteHttpExceptions: true
  };
}

function drivePermissionResponseSucceeded_(response) {
  const code = Number(response.getResponseCode()) || 0;
  const body = String(response.getContentText() || '');
  // Drive may report an existing inherited/direct permission as a conflict.
  // That still means the person already has access.
  return (code >= 200 && code < 300) || (code === 400 && /already|existing|duplicate/i.test(body));
}

function createDrivePermissionSilently_(fileId, email, role) {
  const request = drivePermissionRequest_(fileId, email, role);
  if (!request) return { success: false, reason: 'Missing file or user.' };
  const params = Object.assign({}, request);
  delete params.url;
  const response = UrlFetchApp.fetch(request.url, params);
  return drivePermissionResponseSucceeded_(response)
    ? { success: true }
    : { success: false, reason: 'Drive permission response ' + response.getResponseCode() };
}

function ensureSupportHubAttachmentFolderAccess_(folder) {
  const activeEmails = Array.from(new Set(_getTeamMembersInternal()
    .filter(row => String(row.status || '').toLowerCase() === 'active')
    .map(row => normalizeEmail(row.email))
    .filter(Boolean))).sort();
  const role = attachmentFolderPermissionRole_();
  const signature = role + '|' + activeEmails.join('|');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(ATTACHMENT_ACCESS_SIGNATURE_PROP) === signature) {
    return { failedEmails: [] };
  }

  const failed = [];
  const effectiveEmail = (() => {
    try { return normalizeEmail(Session.getEffectiveUser().getEmail()); } catch (e) { return ''; }
  })();
  const targets = activeEmails.filter(email => !effectiveEmail || !emailsRepresentSameWorkspaceIdentity(email, effectiveEmail));
  const requests = targets.map(email => drivePermissionRequest_(folder.getId(), email, role));
  if (requests.length) {
    try {
      const responses = UrlFetchApp.fetchAll(requests);
      responses.forEach((response, index) => {
        if (!drivePermissionResponseSucceeded_(response)) failed.push(targets[index]);
      });
    } catch (e) {
      targets.forEach(email => failed.push(email));
    }
  }
  if (!failed.length) props.setProperty(ATTACHMENT_ACCESS_SIGNATURE_PROP, signature);
  return { failedEmails: failed };
}

// Screenshots are stored in Drive, not inside a Sheets cell. The editor stores
// only a short Drive link and preview URL, so the source image size does
// not consume the 50,000-character cell limit.
function uploadRichTextImage(dataUrl, fileName, declaredSize, context, recordId, requestingEmail) {
  const member = requireAuthenticatedMember(requestingEmail);
  const raw = String(dataUrl || '');
  const match = raw.match(/^data:(image\/(?:png|jpeg|gif|webp));base64,([A-Za-z0-9+/=\r\n]+)$/i);
  if (!match) throw new Error('Use a PNG, JPG, GIF, or WebP image.');
  const mimeType = match[1].toLowerCase() === 'image/jpg' ? 'image/jpeg' : match[1].toLowerCase();

  let bytes;
  try { bytes = Utilities.base64Decode(match[2].replace(/\s+/g, '')); }
  catch (e) { throw new Error('The image could not be read. Please try attaching it again.'); }
  if (!bytes || !bytes.length) throw new Error('The selected image is empty.');

  // Task attachments may only be added by somebody who can already see the task.
  if (/^Task/i.test(String(context || '')) && recordId) {
    const target = requireTaskRow(recordId);
    const task = taskRowToObject(target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0]);
    const email = normalizeEmail(member.email);
    if (!isSupportMember(member) && task.assignedToEmail !== email && task.createdByEmail !== email) {
      throw new Error('Access denied: you cannot attach an image to this task.');
    }
  }

  const safeName = safeAttachmentFileName(fileName, mimeType);
  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  let file;
  const folder = getSupportHubAttachmentFolder_();
  let folderAccess = { failedEmails: [] };
  try { folderAccess = ensureSupportHubAttachmentFolderAccess_(folder); }
  catch (e) {
    // Preserve access even if the one-time folder sync is blocked by a
    // Workspace policy. The post-create fallback below will retry silently on
    // the individual file without generating sharing email.
    folderAccess = {
      failedEmails: Array.from(new Set(_getTeamMembersInternal()
        .filter(row => String(row.status || '').toLowerCase() === 'active')
        .map(row => normalizeEmail(row.email)).filter(Boolean)))
    };
  }
  try {
    file = folder.createFile(blob);
  } catch (e) {
    throw new Error('The image could not be saved to Google Drive. Please confirm Drive access and try again.');
  }
  file.setDescription('Support Hub image uploaded by ' + member.name + (recordId ? ' for ' + String(recordId) : '') + '.');

  // New files inherit access from the central folder. If a specific folder
  // permission could not be established because of a Workspace policy, retry
  // that one permission silently on the file. Never use DriveApp.addViewer(s),
  // because it can generate a separate sharing notification for every image.
  let sharingWarning = '';
  const stillFailed = [];
  (folderAccess.failedEmails || []).forEach(email => {
    try {
      const retry = createDrivePermissionSilently_(file.getId(), email, 'reader');
      if (!retry.success) stillFailed.push(email);
    } catch (e) { stillFailed.push(email); }
  });
  if (stillFailed.length) {
    sharingWarning = 'The image uploaded, but access could not be confirmed for ' + stillFailed.length + ' team member(s).';
  }

  const fileId = file.getId();
  const result = {
    success: true,
    fileId: fileId,
    fileName: safeName,
    fileSize: bytes.length,
    mimeType: mimeType,
    viewUrl: file.getUrl(),
    previewUrl: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1600',
    warning: sharingWarning
  };
  logAudit('IMAGE_ATTACHMENT_UPLOAD', member.email, member.name, String(recordId || ''), {
    fileId: fileId, fileName: safeName, fileSize: bytes.length,
    context: String(context || '').slice(0, 100), sharingWarning: sharingWarning
  });
  return result;
}

function isExactIsoCalendarDate(value) {
  const m = String(value || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return false;
  const year = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
  const d = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
  return d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day;
}


// Phase 10A: operational categories. These are separate from Team Setup
// Category, which continues to control permissions (Coordinator/Support/Admin).
const BASE_WORK_CATEGORY_CONFIGS = [
  ['Cancellation', ['cancel', 'cancelled', 'canceled', 'cancellation', 'considering cancelling', 'terminate'], 10],
  ['Billing / Payment', ['payment', 'payments', 'pay overdue', 'balance', 'invoice', 'refund', 'charge', 'billing', 'credit card'], 20],
  ['Album / Print', ['album', 'albums', 'print', 'prints', 'canvas', 'wall art', 'book design'], 30],
  ['Video', ['video', 'videos', 'ver', 'revision', 'revisions', 'film', 'footage', 'song', 'timecode'], 40],
  ['Talent', ['talent', 'photographer', 'photographers', 'videographer', 'videographers', 'lead photographer', 'second shooter', 'vendor', 'vendors', 'upload issue'], 50],
  ['Photo', ['photo', 'photos', 'gallery', 'galleries', 'image', 'images', 'raw', 'photograph', 'photographs', 'editing', 'retouch'], 60],
  ['Event Plan', ['event plan', 'event role', 'event hub', 'coverage', 'timeline', 'planning'], 70],
  ['Scheduling', ['schedule', 'scheduled', 'reschedule', 'rescheduled', 'availability', 'booking', 'rebook', 'session'], 80],
  ['System / Salesforce', ['salesforce', 'salesforce record', 'portal', 'login', 'technical', 'bug', 'system error', 'error message', 'lightning'], 90],
  ['Client Communication', ['client', 'email', 'reach out', 'follow up', 'communication', 'call'], 100],
  ['General / Other', [], 999]
];

function defaultWorkCategoryObjects() {
  return BASE_WORK_CATEGORY_CONFIGS.map(row => ({ category: row[0], keywords: row[1].slice(), active: true, sortOrder: row[2] }));
}

function getWorkCategoryConfigs(includeInactive) {
  const cache = CacheService.getScriptCache();
  let configs = null;
  try {
    const cached = cache.get(WORK_CATEGORY_CACHE_KEY);
    if (cached) configs = JSON.parse(cached);
  } catch (e) {}
  if (!Array.isArray(configs)) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_WORK_CATEGORIES);
    if (!sheet || sheet.getLastRow() <= 1) {
      configs = defaultWorkCategoryObjects();
    } else {
      configs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().map((row, index) => ({
        category: String(row[0] || '').trim(),
        keywords: String(row[1] || '').split(',').map(k => k.trim()).filter(Boolean),
        active: row[2] === true || String(row[2] || '').trim().toLowerCase() === 'true' || String(row[2] || '').trim().toLowerCase() === 'active',
        sortOrder: Number(row[3]) || (500 + index),
        updatedAt: row[4] ? safeIsoDate(row[4]) : '',
        updatedBy: String(row[5] || '').trim()
      })).filter(row => row.category);
      if (!configs.some(row => row.category.toLowerCase() === 'general / other')) {
        configs.push({ category: 'General / Other', keywords: [], active: true, sortOrder: 999 });
      }
    }
    try { cache.put(WORK_CATEGORY_CACHE_KEY, JSON.stringify(configs), WORK_CATEGORY_CACHE_SECONDS); } catch (e) {}
  }
  configs.sort((a, b) => a.sortOrder - b.sortOrder || a.category.localeCompare(b.category));
  return includeInactive ? configs : configs.filter(row => row.active);
}

function clearWorkCategoryCache() {
  try { CacheService.getScriptCache().remove(WORK_CATEGORY_CACHE_KEY); } catch (e) {}
}

function normalizeWorkCategory(value) {
  const raw = String(value || '').trim();
  if (!raw) return 'General / Other';
  const found = getWorkCategoryConfigs(true).find(row => row.category.toLowerCase() === raw.toLowerCase());
  return found ? found.category : 'General / Other';
}

function normalizeEventDate(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_TIMEZONE, 'yyyy-MM-dd');
  }
  const raw = String(value || '').trim();
  if (!raw) return '';
  return isExactIsoCalendarDate(raw) ? raw : '';
}

function normalizeEventRole(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const match = EVENT_ROLE_OPTIONS.find(role => role.toLowerCase() === raw.toLowerCase());
  if (!match) throw new Error('Please choose a valid Event Role.');
  return match;
}

function isEventDateCurrentOrFuture(value) {
  const normalized = normalizeEventDate(value);
  if (!normalized) return false;
  const today = Utilities.formatDate(new Date(), APP_TIMEZONE, 'yyyy-MM-dd');
  return normalized >= today;
}

function isActiveEventThisWeekend(flag, eventDate) {
  return !!flag && isEventDateCurrentOrFuture(eventDate);
}

function todayCSTDateString() {
  return Utilities.formatDate(new Date(), APP_TIMEZONE, 'yyyy-MM-dd');
}

function normalizeAttentionUntil(value) {
  const normalized = normalizeEventDate(value) || todayCSTDateString();
  if (normalized < todayCSTDateString()) {
    throw new Error('Attention Today cannot end in the past.');
  }
  return normalized;
}

function isAttentionActive(flag, untilDate) {
  if (flag !== true) return false;
  const until = normalizeEventDate(untilDate);
  return !!until && until >= todayCSTDateString();
}

function attentionWasRequested(flag, setAt) {
  return flag === true || !!setAt;
}

function cleanAttentionReason(value) {
  const reason = String(value || '').trim();
  if (!reason) throw new Error('Please add a short reason for marking this question as Needs Attention Today.');
  if (reason.length > MAX_ATTENTION_REASON_CHARS) {
    throw new Error('The Attention Today reason must be ' + MAX_ATTENTION_REASON_CHARS + ' characters or fewer.');
  }
  return reason;
}

function workCategoryKeywordMatches(text, keyword) {
  const haystack = String(text || '').toLowerCase();
  const needle = String(keyword || '').toLowerCase().trim();
  if (!needle) return false;
  // Multi-word phrases can safely use direct matching. Single-word tokens use
  // boundaries so short terms such as "ver" do not accidentally match words
  // like "cover", and "photo" does not classify "photographer" as Photo.
  if (/[^a-z0-9]/i.test(needle)) return haystack.indexOf(needle) !== -1;
  return new RegExp('(^|[^a-z0-9])' + needle + '(?=$|[^a-z0-9])', 'i').test(haystack);
}

function inferWorkCategory(question, eventName) {
  const text = (String(question || '').replace(/<[^>]*>/g, ' ') + ' ' + String(eventName || '')).toLowerCase();
  const configs = getWorkCategoryConfigs(false);
  for (let i = 0; i < configs.length; i++) {
    if (configs[i].category === 'General / Other') continue;
    if (configs[i].keywords.some(k => workCategoryKeywordMatches(text, k))) return configs[i].category;
  }
  return 'General / Other';
}


// Canonical status values — never write anything else into the Status columns.
const STATUS_OPEN = "Open";

function isFollowUpValue(value) {
  if (value === true || value === 1) return true;
  return ['true', 'yes', '1', 'follow-up', 'follow up'].includes(String(value || '').trim().toLowerCase());
}
const STATUS_HOLD = "Hold";
const STATUS_ANSWERED = "Answered";

// App-wide timezone: every date shown to users follows Central Time (CST/CDT).
const APP_TIMEZONE = 'America/Chicago';

// Phase 8: Google Workspace authentication. Deploy this web app with
// access limited to approved company Workspace domains and execute as the user
// accessing the web app. Identity is then taken from Google's authenticated
// session, never from a user-selectable client value.
const WORKSPACE_DOMAINS = ['orionphotogroup.com', 'georgestreetphoto.com', 'lilyandlime.com'];

function isApprovedWorkspaceEmail(email) {
  const normalized = normalizeEmail(email);
  if (!normalized || !normalized.includes('@')) return false;
  return WORKSPACE_DOMAINS.includes(normalized.split('@')[1] || '');
}

function workspaceLocalPart(email) {
  const normalized = normalizeEmail(email);
  return normalized && normalized.includes('@') ? normalized.split('@')[0] : '';
}

// Employees may legitimately have the same mailbox identity across our approved
// Workspace domains (for example user@orionphotogroup.com and
// user@georgestreetphoto.com). Exact email remains preferred, but matching local
// parts across the three company-controlled domains are accepted as the same
// Workspace identity. Team Setup remains the source of truth for the user's
// Support Hub profile and permissions.
function emailsRepresentSameWorkspaceIdentity(a, b) {
  const left = normalizeEmail(a);
  const right = normalizeEmail(b);
  if (!left || !right) return false;
  if (left === right) return true;
  if (!isApprovedWorkspaceEmail(left) || !isApprovedWorkspaceEmail(right)) return false;
  return workspaceLocalPart(left) === workspaceLocalPart(right);
}

const ACTIVE_WORKSPACE_PROFILE_PROP = 'SUPPORT_HUB_ACTIVE_WORKSPACE_PROFILE_V1';

function memberProfileKey(member) {
  if (!member) return '';
  return [normalizeEmail(member.email), String(member.category || '').trim(), String(member.name || '').trim()].join('|');
}

function withProfileKey(member) {
  return member ? Object.assign({}, member, { profileKey: memberProfileKey(member) }) : null;
}

function getWorkspaceIdentityProfiles(authenticatedEmail) {
  const auth = normalizeEmail(authenticatedEmail);
  if (!auth) return [];
  const active = _getTeamMembersInternal().filter(m => String(m.status || '').trim().toLowerCase() === 'active');

  // Exact mailbox matches are always preferred. This intentionally supports
  // one Google Workspace identity having both a Support and an Admin profile.
  const exact = active.filter(m => normalizeEmail(m.email) === auth);
  if (exact.length) return exact.map(withProfileKey);

  if (!isApprovedWorkspaceEmail(auth)) return [];
  return active.filter(m => emailsRepresentSameWorkspaceIdentity(m.email, auth)).map(withProfileKey);
}

function getStoredWorkspaceProfile(authenticatedEmail) {
  const profiles = getWorkspaceIdentityProfiles(authenticatedEmail);
  if (!profiles.length) return null;
  let key = '';
  try { key = PropertiesService.getUserProperties().getProperty(ACTIVE_WORKSPACE_PROFILE_PROP) || ''; } catch (e) {}
  if (!key) return profiles.length === 1 ? profiles[0] : null;
  return profiles.find(p => p.profileKey === key) || (profiles.length === 1 ? profiles[0] : null);
}

function clearActiveWorkspaceProfile() {
  try { PropertiesService.getUserProperties().deleteProperty(ACTIVE_WORKSPACE_PROFILE_PROP); } catch (e) {}
  return { success: true };
}

function findTeamMemberForWorkspaceIdentity(authenticatedEmail) {
  return getStoredWorkspaceProfile(authenticatedEmail);
}

function requireAuthenticatedWorkspaceMember() {
  const authenticatedEmail = getAuthenticatedEmail();
  const profiles = getWorkspaceIdentityProfiles(authenticatedEmail);
  if (!profiles.length) {
    throw new Error('Access denied: your signed-in Google Workspace account is not linked to an active Team Setup user.');
  }
  const member = getStoredWorkspaceProfile(authenticatedEmail);
  if (!member) {
    throw new Error('Choose which Support Hub profile you want to use before continuing.');
  }
  return member;
}

// Phase 9.7: server-side role recovery for shared-mailbox identities.
// A Workspace account may intentionally have both an operational profile
// (Coordinator/Support) and a separate Admin profile using the same email.
// UserProperties normally remembers the profile selected at login, but a
// server call must not fail just because that transient selection is missing.
// Sensitive actions therefore resolve the REQUIRED role directly from the
// authenticated Workspace identity. This does not let a user elevate access:
// the matching active Team Setup profile must already exist for that Google
// account.
function getAuthenticatedProfileForRole(role) {
  const authenticatedEmail = getAuthenticatedEmail();
  const profiles = getWorkspaceIdentityProfiles(authenticatedEmail);
  if (!profiles.length) {
    throw new Error('Access denied: your signed-in Google Workspace account is not linked to an active Team Setup user.');
  }

  const stored = getStoredWorkspaceProfile(authenticatedEmail);
  if (stored) {
    if (role === 'Support' && isSupportMember(stored) && !isAdminMember(stored)) return stored;
    if (role === 'Admin' && isAdminMember(stored)) return stored;
    if (role === 'Operational' && !isAdminMember(stored)) return stored;
  }

  let matches = [];
  if (role === 'Support') {
    // Admin is deliberately excluded. Admin is a maintenance/development
    // profile and is never used to answer Supervisor Desk questions.
    matches = profiles.filter(p => isSupportMember(p) && !isAdminMember(p));
  } else if (role === 'Admin') {
    matches = profiles.filter(p => isAdminMember(p));
  } else {
    matches = profiles.filter(p => !isAdminMember(p));
  }

  if (matches.length === 1) return matches[0];
  if (!matches.length) {
    if (role === 'Support') throw new Error('Access denied: this action requires your active Support profile.');
    if (role === 'Admin') throw new Error('Access denied: this action requires your active Admin profile.');
    throw new Error('Access denied: no active operational Support Hub profile is linked to your Google account.');
  }
  throw new Error('More than one matching Support Hub profile exists for your Google account. Please ask an Admin to correct Team Setup.');
}

function getAuthenticatedEmail() {
  const email = normalizeEmail(Session.getActiveUser().getEmail());
  if (!email) {
    throw new Error('Google Workspace identity is unavailable. Redeploy the web app as User accessing the web app and restrict access to your organization.');
  }
  const domain = email.split('@')[1] || '';
  if (!WORKSPACE_DOMAINS.includes(domain)) {
    throw new Error('Access denied: please open Support Hub using an approved company Google Workspace account.');
  }
  return email;
}

function requireAuthenticatedMember(claimedEmail) {
  const authenticatedEmail = getAuthenticatedEmail();
  let authenticatedMember = getStoredWorkspaceProfile(authenticatedEmail);
  const claimed = normalizeEmail(claimedEmail);

  // If the browser is operating as the same authenticated person but the
  // selected-profile UserProperty is missing, recover the person's unique
  // non-Admin operational profile. This is especially important for users who
  // share one Workspace email between Support and Admin profiles.
  if (!authenticatedMember) {
    const profiles = getWorkspaceIdentityProfiles(authenticatedEmail);
    const operational = profiles.filter(p => !isAdminMember(p));
    if (operational.length === 1 && (!claimed || emailsRepresentSameWorkspaceIdentity(claimed, authenticatedEmail))) {
      authenticatedMember = operational[0];
    } else {
      throw new Error('Choose which Support Hub profile you want to use before continuing.');
    }
  }

  if (!claimed || claimed === normalizeEmail(authenticatedMember.email) || emailsRepresentSameWorkspaceIdentity(claimed, authenticatedEmail)) {
    return authenticatedMember;
  }

  // Admin Test Mode is available only when the real Workspace identity has
  // explicitly selected its Admin profile and passed the Admin password.
  if (isAdminMember(authenticatedMember)) {
    const effectiveMember = requireTeamMember(claimed);
    return Object.assign({}, effectiveMember, {
      _adminTestMode: true,
      _authenticatedAdminEmail: authenticatedMember.email,
      _authenticatedAdminName: authenticatedMember.name
    });
  }

  throw new Error('Security check failed: the selected user does not match your signed-in Google account. Refresh Support Hub and try again.');
}

function requireAuthenticatedSupport(claimedEmail) {
  const authenticatedEmail = getAuthenticatedEmail();
  const claimed = normalizeEmail(claimedEmail);

  // Admin Test Mode keeps the existing explicit Admin-profile path. In normal
  // use, answering/holding/assigning always resolves to the real operational
  // Support profile and never to the Admin profile.
  const stored = getStoredWorkspaceProfile(authenticatedEmail);
  if (stored && isAdminMember(stored) && claimed && !emailsRepresentSameWorkspaceIdentity(claimed, authenticatedEmail)) {
    const effective = requireAuthenticatedMember(claimedEmail);
    if (!isSupportMember(effective)) throw new Error('Access denied: this action requires a Supervisor / Support role.');
    return effective;
  }

  if (claimed && !emailsRepresentSameWorkspaceIdentity(claimed, authenticatedEmail)) {
    throw new Error('Security check failed: the selected Support user does not match your signed-in Google account.');
  }

  return getAuthenticatedProfileForRole('Support');
}

function requireAuthenticatedAdmin(claimedEmail) {
  const authenticatedEmail = getAuthenticatedEmail();
  const claimed = normalizeEmail(claimedEmail);
  if (claimed && !emailsRepresentSameWorkspaceIdentity(claimed, authenticatedEmail)) {
    throw new Error('Security check failed: the selected Admin user does not match your signed-in Google account.');
  }
  return getAuthenticatedProfileForRole('Admin');
}

function getWorkspaceBootstrap() {
  ensureSheetsExist();
  const member = requireAuthenticatedMember();
  return {
    success: true,
    authenticatedEmail: getAuthenticatedEmail(),
    member: member,
    team: _getTeamMembersInternal(),
    requiresAdminPassword: isAdminMember(member)
  };
}

// Fresh-session bootstrap. If the signed-in Workspace identity has more than
// one active Team Setup profile (for example Support + Admin using the same
// mailbox), no profile is selected automatically; the user chooses explicitly.
function getWorkspaceProfileBootstrap() {
  ensureSheetsExist();
  const authenticatedEmail = getAuthenticatedEmail();
  const profiles = getWorkspaceIdentityProfiles(authenticatedEmail);
  return {
    success: true,
    authenticatedEmail: authenticatedEmail,
    profiles: profiles,
    team: profiles.length ? _getTeamMembersInternal() : []
  };
}

// Backward-compatible name retained for older frontends, but dual-profile
// identities are never auto-logged in as Admin.
function getAdminAutoLoginBootstrap() {
  const res = getWorkspaceProfileBootstrap();
  if (res.profiles.length !== 1 || !isAdminMember(res.profiles[0])) {
    return { success: true, isAdmin: false, authenticatedEmail: res.authenticatedEmail, profiles: res.profiles };
  }
  return { success: true, isAdmin: false, authenticatedEmail: res.authenticatedEmail, profiles: res.profiles };
}

function getRawTeamMembersByEmail(email) {
  email = normalizeEmail(email);
  if (!email) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (normalizeEmail(data[i][4]) !== email) continue;
    out.push({
      rowIndex: i + 1,
      name: stripNoraPrefix(String(data[i][0] || '').trim()),
      title: String(data[i][1] || '').trim(),
      status: String(data[i][2] || '').trim(),
      category: String(data[i][3] || '').trim(),
      email: email,
      passwordHash: String(data[i][5] || '').trim(),
      passwordSalt: String(data[i][6] || '').trim()
    });
  }
  return out;
}

function activateWorkspaceProfile(profileKey, passwordAttempt) {
  ensureSheetsExist();
  const authenticatedEmail = getAuthenticatedEmail();
  const profiles = getWorkspaceIdentityProfiles(authenticatedEmail);
  const selected = profiles.find(p => p.profileKey === String(profileKey || ''));
  if (!selected) throw new Error('The selected login profile is not available for your signed-in Google account.');

  if (isAdminMember(selected)) {
    const raws = getRawTeamMembersByEmail(selected.email);
    const raw = raws.find(r => isAdminCategoryValue(r.category) && r.name === selected.name) || raws.find(r => isAdminCategoryValue(r.category));
    if (!raw || String(raw.status || '').toLowerCase() !== 'active') throw new Error('This Admin profile is inactive or unavailable.');
    if (!raw.passwordHash) throw new Error('This Admin account has no password set yet. Ask another Admin to set one in Team Setup.');
    const attemptHash = hashPassword(passwordAttempt, raw.passwordSalt);
    if (attemptHash !== raw.passwordHash) {
      logAudit('ADMIN_LOGIN_FAILED', authenticatedEmail, selected.name, '', 'Incorrect password attempt');
      logAudit('AUTH_FAILED', authenticatedEmail, selected.name, '', { effectiveEmail: selected.email, reason: 'Incorrect Admin password' });
      throw new Error('Incorrect password.');
    }
    logAudit('ADMIN_LOGIN', authenticatedEmail, selected.name, '', 'Admin password verified');
  }

  try { PropertiesService.getUserProperties().setProperty(ACTIVE_WORKSPACE_PROFILE_PROP, selected.profileKey); } catch (e) {
    throw new Error('Could not save your selected Support Hub profile. Please try again.');
  }
  logAudit('AUTH_SUCCESS', authenticatedEmail, selected.name, '', {
    method: isAdminMember(selected) ? 'profile_choice_admin_password' : 'profile_choice',
    effectiveEmail: selected.email,
    effectiveCategory: selected.category,
    effectiveTitle: selected.title
  });
  return { success: true, member: selected, team: _getTeamMembersInternal(), authenticatedEmail: authenticatedEmail };
}

function recordWorkspaceLoginSuccess(claimedEmail, method) {
  const authenticatedEmail = getAuthenticatedEmail();
  const member = requireAuthenticatedWorkspaceMember();
  const effective = normalizeEmail(claimedEmail) || member.email;
  logAudit('AUTH_SUCCESS', authenticatedEmail, member.name, '', {
    method: String(method || 'manual'),
    effectiveEmail: effective,
    effectiveCategory: member.category,
    effectiveTitle: member.title
  });
  return { success: true };
}

function recordWorkspaceLogout(claimedEmail) {
  const authenticatedEmail = getAuthenticatedEmail();
  let member = null;
  try { member = requireAuthenticatedWorkspaceMember(); } catch (e) {}
  logAudit('AUTH_LOGOUT', authenticatedEmail, member ? member.name : '', '', {
    effectiveEmail: normalizeEmail(claimedEmail) || (member ? member.email : authenticatedEmail),
    effectiveCategory: member ? member.category : ''
  });
  clearActiveWorkspaceProfile();
  return { success: true };
}

function startAdminTestMode(targetEmail, requestingEmail) {
  const admin = requireAuthenticatedAdmin(requestingEmail);
  const target = requireTeamMember(targetEmail);
  logAudit('ADMIN_TEST_START', admin.email, admin.name, '', {
    targetEmail: target.email,
    targetName: target.name,
    targetCategory: target.category,
    targetTitle: target.title
  });
  return { success: true, admin: admin, target: target, team: _getTeamMembersInternal() };
}

function stopAdminTestMode(targetEmail, requestingEmail) {
  const admin = requireAuthenticatedAdmin(requestingEmail);
  const target = findTeamMemberByEmail(targetEmail);
  logAudit('ADMIN_TEST_STOP', admin.email, admin.name, '', {
    targetEmail: normalizeEmail(targetEmail),
    targetName: target ? target.name : ''
  });
  return { success: true, admin: admin, team: _getTeamMembersInternal() };
}

function getLoginAccessMonitor(requestingEmail) {
  const admin = requireAuthenticatedAdmin(requestingEmail);
  const team = _getTeamMembersInternal();
  const sheet = ensureAuditSheet(SpreadsheetApp.getActiveSpreadsheet());
  const lastRow = sheet.getLastRow();
  const state = {};
  team.forEach(m => {
    const key = memberProfileKey(m);
    state[key] = {
      name: m.name, title: m.title, category: m.category, status: m.status, email: m.email,
      lastSuccessAt: '', lastSuccessMethod: '', lastFailureAt: '', lastFailureReason: '', lastLogoutAt: ''
    };
  });

  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      const stamp = rows[i][0];
      const actorEmail = normalizeEmail(rows[i][1]);
      const action = String(rows[i][3] || '');
      let details = {};
      try { details = JSON.parse(String(rows[i][5] || '{}')); } catch (e) { details = { reason: String(rows[i][5] || '') }; }
      const effectiveEmail = normalizeEmail(details.effectiveEmail || details.enteredEmail || actorEmail);
      const category = String(details.effectiveCategory || '').trim();
      let rec = null;
      if (category) rec = state[[effectiveEmail, category, Object.keys(state).map(k=>state[k]).find(r=>r.email===effectiveEmail&&r.category===category)?.name || ''].join('|')];
      if (!rec) rec = Object.values(state).find(r => r.email === effectiveEmail) || Object.values(state).find(r => r.email === actorEmail);
      if (!rec) continue;
      if (action === 'AUTH_SUCCESS' && !rec.lastSuccessAt) {
        rec.lastSuccessAt = stamp instanceof Date ? stamp.toISOString() : safeIsoDate(stamp);
        rec.lastSuccessMethod = String(details.method || 'manual');
      } else if (action === 'AUTH_FAILED' && !rec.lastFailureAt) {
        rec.lastFailureAt = stamp instanceof Date ? stamp.toISOString() : safeIsoDate(stamp);
        rec.lastFailureReason = String(details.reason || 'Login verification failed');
      } else if (action === 'AUTH_LOGOUT' && !rec.lastLogoutAt) {
        rec.lastLogoutAt = stamp instanceof Date ? stamp.toISOString() : safeIsoDate(stamp);
      }
    }
  }
  return { success: true, requestedBy: admin.email, rows: Object.values(state) };
}

// ==========================================
// PHASE 13G - PERFORMANCE DIAGNOSTICS
// ==========================================
// Telemetry is deliberately sparse. Each browser sends one normal sample per
// core feed, then only slow or failed calls. These writes run after the user
// already has a response and skip themselves when the write lock is busy.
const PERFORMANCE_HEADERS = [
  'Timestamp', 'Metric', 'Status', 'User Email', 'User Name', 'Role',
  'Round Trip Ms', 'Server Ms', 'Auth Ms', 'Data Ms', 'Error', 'Release', 'Sample Type'
];
const PERFORMANCE_METRICS = ['supervisor','history','tasks','peer','team-roster','search-server'];

function ensurePerformanceLogSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PERFORMANCE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PERFORMANCE);
    sheet.appendRow(PERFORMANCE_HEADERS);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const width = PERFORMANCE_HEADERS.length;
  const headers = sheet.getRange(1, 1, 1, Math.max(width, sheet.getLastColumn())).getValues()[0];
  for (let i = 0; i < width; i++) {
    if (String(headers[i] || '').trim() !== PERFORMANCE_HEADERS[i]) sheet.getRange(1, i + 1).setValue(PERFORMANCE_HEADERS[i]);
  }
  return sheet;
}

function performanceMilliseconds_(value) {
  const number = Number(value);
  if (!isFinite(number)) return 0;
  return Math.max(0, Math.min(300000, Math.round(number)));
}

function recordPerformanceEvent(payload, requestingEmail) {
  const actor = requireAuthenticatedMember(requestingEmail);
  payload = payload || {};
  const metric = String(payload.metric || '').trim().toLowerCase();
  if (PERFORMANCE_METRICS.indexOf(metric) === -1) return { recorded:false, reason:'unsupported_metric' };
  const status = String(payload.status || 'success').trim().toLowerCase() === 'failure' ? 'failure' : 'success';
  const roundTripMs = performanceMilliseconds_(payload.roundTripMs);
  const sampleType = status === 'failure' ? 'failure' : (roundTripMs >= PERFORMANCE_SLOW_MS ? 'slow' : 'baseline');
  if (sampleType === 'baseline' && payload.baseline !== true) return { recorded:false, reason:'not_sampled' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return { recorded:false, reason:'busy' };
  try {
    const sheet = ensurePerformanceLogSheet_(SpreadsheetApp.getActiveSpreadsheet());
    sheet.appendRow([
      new Date(), metric, status, normalizeEmail(actor.email), String(actor.name || ''), String(actor.category || ''),
      roundTripMs, performanceMilliseconds_(payload.serverMs), performanceMilliseconds_(payload.authMs),
      performanceMilliseconds_(payload.dataMs), String(payload.error || '').slice(0, 300), APP_RELEASE, sampleType
    ]);
    return { recorded:true };
  } catch (e) {
    Logger.log('Performance telemetry skipped: ' + e);
    return { recorded:false, reason:'write_failed' };
  } finally {
    lock.releaseLock();
  }
}

function performancePercentile_(values, percentile) {
  if (!values || !values.length) return 0;
  const sorted = values.slice().sort(function(a, b) { return a - b; });
  const index = Math.max(0, Math.min(sorted.length - 1, Math.ceil(sorted.length * percentile) - 1));
  return sorted[index];
}

function getPerformanceDiagnostics(requestingEmail) {
  const admin = requireAuthenticatedAdmin(requestingEmail);
  const sheet = ensurePerformanceLogSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success:true, requestedBy:admin.email, release:APP_RELEASE, total:0, slow:0, failures:0, metrics:[], recent:[] };
  const startRow = Math.max(2, lastRow - 499);
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, PERFORMANCE_HEADERS.length).getValues();
  const groups = {};
  let slow = 0;
  let failures = 0;
  const records = rows.map(function(row) {
    const record = {
      timestamp: row[0] instanceof Date ? row[0].toISOString() : safeIsoDate(row[0]),
      metric: String(row[1] || ''), status: String(row[2] || ''), userEmail: normalizeEmail(row[3]),
      userName: String(row[4] || ''), role: String(row[5] || ''), roundTripMs: performanceMilliseconds_(row[6]),
      serverMs: performanceMilliseconds_(row[7]), authMs: performanceMilliseconds_(row[8]), dataMs: performanceMilliseconds_(row[9]),
      error: String(row[10] || ''), release: String(row[11] || ''), sampleType: String(row[12] || '')
    };
    if (record.status === 'failure') failures++;
    if (record.roundTripMs >= PERFORMANCE_SLOW_MS) slow++;
    if (!groups[record.metric]) groups[record.metric] = { metric:record.metric, values:[], failures:0, lastAt:'', lastUser:'' };
    const group = groups[record.metric];
    group.values.push(record.roundTripMs);
    if (record.status === 'failure') group.failures++;
    if (!group.lastAt || record.timestamp > group.lastAt) { group.lastAt = record.timestamp; group.lastUser = record.userName || record.userEmail; }
    return record;
  });
  const metrics = Object.keys(groups).map(function(key) {
    const group = groups[key];
    const total = group.values.reduce(function(sum, value) { return sum + value; }, 0);
    return {
      metric:group.metric, samples:group.values.length, failures:group.failures,
      averageMs:group.values.length ? Math.round(total / group.values.length) : 0,
      p95Ms:performancePercentile_(group.values, 0.95), maxMs:Math.max.apply(null, group.values),
      lastAt:group.lastAt, lastUser:group.lastUser
    };
  }).sort(function(a, b) { return b.p95Ms - a.p95Ms || a.metric.localeCompare(b.metric); });
  return {
    success:true, requestedBy:admin.email, release:APP_RELEASE, total:records.length, slow:slow, failures:failures,
    metrics:metrics,
    recent:records.slice().reverse().slice(0, 25)
  };
}

function verifyWorkspaceEmail(enteredEmail) {
  ensureSheetsExist();
  const entered = normalizeEmail(enteredEmail);
  let authenticatedEmail = '';
  try {
    if (!entered) throw new Error('Enter your company email address.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(entered)) throw new Error('Enter a valid email address.');
    const domain = entered.split('@')[1] || '';
    if (!WORKSPACE_DOMAINS.includes(domain)) {
      throw new Error('Please use an approved company email address: @orionphotogroup.com, @georgestreetphoto.com, or @lilyandlime.com.');
    }

    authenticatedEmail = getAuthenticatedEmail();
    if (!emailsRepresentSameWorkspaceIdentity(entered, authenticatedEmail)) {
      throw new Error('This email does not match the Google Workspace account currently signed in.');
    }

    const profiles = _getTeamMembersInternal()
      .filter(m => normalizeEmail(m.email) === entered && String(m.status || '').trim().toLowerCase() === 'active')
      .map(withProfileKey);
    if (!profiles.length) throw new Error('This email is not linked to an active Team Setup user.');

    return {
      success: true,
      verifiedEmail: entered,
      profiles: profiles,
      member: profiles.length === 1 ? profiles[0] : null,
      requiresProfileSelection: profiles.length > 1,
      team: _getTeamMembersInternal()
    };
  } catch (err) {
    try {
      logAudit('AUTH_FAILED', authenticatedEmail || entered, '', '', {
        enteredEmail: entered, effectiveEmail: entered, reason: String(err && err.message ? err.message : err)
      });
    } catch (logErr) {}
    throw err;
  }
}

function getTeamMembersForCurrentUser() {
  requireAuthenticatedMember();
  // Team Setup is a small sheet and is refreshed far less often than ticket
  // polling. Reading it authoritatively here ensures every role sees routing
  // changes immediately, while normal authentication/submission calls keep
  // using the fast versioned roster cache.
  return _getTeamMembersInternal(true);
}

// ==========================================
// ROLE MODEL
// ==========================================
// PERMISSIONS ARE DRIVEN BY CATEGORY ONLY. Category is one of exactly three
// fixed values: "Coordinator" (limited access), "Support" (elevated access -
// answer/assign/hold tickets, Reporting, Team Setup), "Admin" (everything
// Support has, plus editing/deleting records, roster management, and the
// Salesforce/Titles config panels). These three values are wired directly
// into isSupportMember()/isAdminMember() below and are NOT admin-addable -
// a 4th category would need matching permission logic added throughout the
// rest of this file, so it can't be safely introduced from a config screen.
//
// TITLE (Coordinator, Supervisor, Manager, "Project Manager", etc.) is a
// display label ONLY. It has no effect on what a person can do - that is
// intentional, so that adding a fancy-sounding custom title can never
// accidentally grant elevated access. The set of available titles is
// admin-manageable (see BASE_TITLES / getAllTitles() / addCustomTitle() /
// deleteCustomTitle() below) so Admins can add roles like "Project Manager"
// without a code change, while access still always comes from Category.
//
// NOTE: an earlier version of this app granted Support access to anyone
// whose TITLE was Manager/Assistant Manager/Supervisor/Escalation Supervisor,
// regardless of Category. That was a hidden bypass (a Coordinator-category
// member with one of those titles silently got Support access) and has been
// removed - Category is now the only thing that determines access.
const BASE_TITLES = ["Coordinator", "Sr. Coordinator", "Supervisor", "Escalation Supervisor", "Assistant Manager", "Manager"];
const CONFIG_KEY_CUSTOM_TITLES = 'CUSTOM_TITLES_LIST';

function doGet() {
  ensureSheetsExist();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Support Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Optional convenience menu inside the bound Sheet.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Support Hub')
    .addItem('Run setup / migration', 'runManualSetup')
    .addToUi();
}
function runManualSetup() {
  ensureSheetsExist(true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Setup complete - sheets are up to date.', 'Support Hub', 5);
}

// ==========================================
// CONCURRENCY: the single write-lock wrapper
// ==========================================
function withLock(fn, options) {
  const lock = LockService.getScriptLock();
  const requestedAt = Date.now();
  try {
    lock.waitLock(LOCK_WAIT_MS);
  } catch (e) {
    throw new Error("The system is busy handling another change right now. Please try again in a few seconds.");
  }
  const acquiredAt = Date.now();
  try {
    const result = fn();
    // Ticket/version refreshes are only needed for mutations that can change
    // queue/answered data. Preference/config-only writes and explicit no-op
    // results opt out so they don't make every user reload the full data set.
    const skipVersion = (options && options.bumpDataVersion === false) || !!(result && result._skipDataVersion);
    if (!skipVersion) {
      // The business write has already completed at this point. A temporary
      // PropertiesService failure must not turn that successful write into a
      // browser-visible error, because the user may retry and create another
      // record. Tell the client to force-refresh if the version counter could
      // not be advanced.
      try {
        const nextVersion = bumpDataVersion();
        if (result && typeof result === 'object') result.dataVersion = nextVersion;
      } catch (versionError) {
        Logger.log('Data version update failed after a successful write: ' + versionError);
        if (result && typeof result === 'object') result.forceRefresh = true;
      }
    }
    if (result && Object.prototype.hasOwnProperty.call(result, '_skipDataVersion')) delete result._skipDataVersion;
    return result;
  } finally {
    const completedAt = Date.now();
    const waitMs = Math.max(0, acquiredAt - requestedAt);
    const heldMs = Math.max(0, completedAt - acquiredAt);
    if (waitMs >= SLOW_LOCK_LOG_MS || heldMs >= SLOW_LOCK_LOG_MS) {
      try {
        Logger.log('Support Hub slow write lock: ' + JSON.stringify({
          operation: String(options && options.operation || 'unspecified'),
          waitMs: waitMs,
          heldMs: heldMs
        }));
      } catch (timingLogError) {}
    }
    lock.releaseLock();
  }
}

// ==========================================
// SHEET SETUP & MIGRATIONS (idempotent)
// ==========================================
function ensureSheetsExist(force) {
  const cache = CacheService.getScriptCache();
  const properties = PropertiesService.getScriptProperties();
  if (!force && properties.getProperty(SCHEMA_VERSION_KEY) === SCHEMA_VERSION) {
    try { cache.put(SETUP_CACHE_KEY, '1', SETUP_CACHE_SECONDS); } catch (e) {}
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureWorkCategoriesSheet(ss);
  ensureCategorySuggestionsSheet(ss);

  let qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!qSheet) {
    qSheet = ss.insertSheet(SHEET_QUESTIONS);
    qSheet.appendRow(["Question", "Event/Wedding Name", "Asked By", "Created", "Due Date & Time", "Auto Priority", "Hours Left", "Case / Event Link", "Answer", "Status", "Assigned To", "Ticket ID", "Hold Reason", "Hold Since", "Cumulative Hold Hours", "Asked By Email", "Is Follow Up", "Assignment Notes", "Status Changed At", "Status Changed By", "Work Category", "Event This Weekend", "Event Date", "Event Role", "Routed Support Email", "Needs Attention Today", "Attention Reason", "Attention Set At", "Attention Set By", "Attention Until", "Attention Cleared At", "Attention Cleared By", "Attention First Review At", "Attention First Review By", "Requester Flagged", "Requester Flagged At", "Requester Flagged By", "Requester Flag Reason", "Requester Flag Acknowledged At", "Requester Flag Acknowledged By"]);
  }
  ensureQuestionsSchema(qSheet);

  let aSheet = ss.getSheetByName(SHEET_ANSWERED);
  if (!aSheet) {
    aSheet = ss.insertSheet(SHEET_ANSWERED);
    aSheet.appendRow(["Question", "Event/Wedding Name", "Asked By", "Created", "Answered Date", "Auto Priority", "Turnaround Hours", "Case / Event Link", "Answer", "Status", "Answered By", "Read By", "Ticket ID", "Hold Hours Excluded", "Asked By Email", "First Read At", "Last Updated At", "Work Category", "Event This Weekend", "Event Date", "Event Role", "Routed Support Email", "Answered By Email", "Attention Requested", "Attention Reason", "Attention Set At", "Attention Set By", "Attention Until", "Attention Cleared At", "Attention Cleared By", "Attention First Review At", "Attention First Review By"]);
  }
  ensureAnsweredSchema(aSheet);

  if (!ss.getSheetByName(SHEET_TEAM)) {
    const s = ss.insertSheet(SHEET_TEAM);
    s.appendRow(["Name", "Title", "Status", "Category", "Email", "PasswordHash", "PasswordSalt", "Primary Support Email", "Backup Support Email"]);
  }
  ensureTeamSchema(ss.getSheetByName(SHEET_TEAM));

  ensureAuditSheet(ss);
  ensureDeletedSheet(ss);
  ensureNotificationPreferencesSheet(ss);
  ensureMentionNotificationsSheet(ss);
  ensureTaskNotificationsSheet(ss);
  ensureTasksSheet(ss);
  ensurePeerQuestionsSheet(ss);

  // All required sheets/columns now exist. Cache this short-lived fact so
  // high-frequency polling does not re-read every header on every request.
  properties.setProperty(SCHEMA_VERSION_KEY, SCHEMA_VERSION);
  cache.put(SETUP_CACHE_KEY, '1', SETUP_CACHE_SECONDS);
}

function ensureWorkCategoriesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_WORK_CATEGORIES);
  const headers = ['Category', 'Keywords', 'Active', 'Sort Order', 'Updated At', 'Updated By'];
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WORK_CATEGORIES);
    sheet.appendRow(headers);
    const now = new Date();
    const rows = defaultWorkCategoryObjects().map(row => [row.category, row.keywords.join(', '), true, row.sortOrder, now, 'System Setup']);
    if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.setFrozenRows(1);
    clearWorkCategoryCache();
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach((header, index) => { if (String(current[index] || '').trim() !== header) sheet.getRange(1, index + 1).setValue(header); });
  if (sheet.getLastRow() <= 1) {
    const now = new Date();
    const rows = defaultWorkCategoryObjects().map(row => [row.category, row.keywords.join(', '), true, row.sortOrder, now, 'System Setup']);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    clearWorkCategoryCache();
  }
  return sheet;
}

function ensureCategorySuggestionsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_CATEGORY_SUGGESTIONS);
  const headers = ['Suggestion ID', 'Suggested Category', 'Suggested By', 'Suggested By Email', 'Context', 'Record ID', 'Created At', 'Status', 'Reviewed By', 'Reviewed At', 'Admin Note'];
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CATEGORY_SUGGESTIONS);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach((header, index) => { if (String(current[index] || '').trim() !== header) sheet.getRange(1, index + 1).setValue(header); });
  return sheet;
}

function ensureAuditSheet(ss) {
  let s = ss.getSheetByName(SHEET_AUDIT);
  if (!s) {
    s = ss.insertSheet(SHEET_AUDIT);
    s.appendRow(["Timestamp", "Actor Email", "Actor Name", "Action", "Ticket ID", "Details"]);
    s.setFrozenRows(1);
  }
  return s;
}

function ensureDeletedSheet(ss) {
  let s = ss.getSheetByName(SHEET_DELETED);
  if (!s) {
    s = ss.insertSheet(SHEET_DELETED);
    s.appendRow(["Deleted At", "Deleted By", "Source Sheet", "Ticket ID", "Question", "Event", "Asked By", "Full Row Data (JSON)"]);
    s.setFrozenRows(1);
  }
  return s;
}

function ensureTasksSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_TASKS);
  const expected = [
    'Task ID', 'Task Title', 'Instructions', 'Assigned To', 'Assigned To Email',
    'Created By', 'Created By Email', 'Created At', 'Due Date', 'Priority',
    'Status', 'Completed At', 'Completion Note', 'Source', 'Related Ticket ID',
    'Related Event', 'Updated At', 'Updated By', 'Work Category',
    'Related Event / Client / Talent', 'Related Link', 'Parent Type', 'Parent ID',
    'Support Owner Email', 'Completed By Email', 'Completed By', 'Related Entity Type',
    'Need Attention Today', 'Attention Set At', 'Attention Set By', 'Attention Until'
  ];

  if (!s) {
    s = ss.insertSheet(SHEET_TASKS);
    s.appendRow(expected);
    s.setFrozenRows(1);
    return s;
  }

  const lastCol = Math.max(s.getLastColumn(), 1);
  const headers = s.getRange(1, 1, 1, Math.max(lastCol, T_WIDTH)).getValues()[0];
  const workCategoryWasMissing = String(headers[T_COL.WORK_CATEGORY - 1] || '').trim() !== 'Work Category';
  const relatedNameHeader = String(headers[T_COL.RELATED_CLIENT_TALENT - 1] || '').trim();
  const relatedNameWasMissing = !relatedNameHeader;
  const parentTypeWasMissing = String(headers[T_COL.PARENT_TYPE - 1] || '').trim() !== 'Parent Type';
  const parentIdWasMissing = String(headers[T_COL.PARENT_ID - 1] || '').trim() !== 'Parent ID';

  // Keep the fixed positional schema, but do not rewrite task data on every poll.
  // The only row-level migration below runs when a column is genuinely introduced.
  for (let i = 0; i < expected.length; i++) {
    const current = String(headers[i] || '').trim();
    const wanted = expected[i];
    // Column 20 had an earlier label; rename it without treating the existing data as missing.
    const compatibleOldRelatedLabel = i === T_COL.RELATED_CLIENT_TALENT - 1 && current === 'Related Client / Talent';
    if (current !== wanted) s.getRange(1, i + 1).setValue(wanted);
    if (compatibleOldRelatedLabel) {
      // Label-only migration; existing values are left untouched.
    }
  }

  const lastRow = s.getLastRow();
  if (lastRow <= 1) return s;
  if (!workCategoryWasMissing && !relatedNameWasMissing && !parentTypeWasMissing && !parentIdWasMissing) return s;

  const rows = s.getRange(2, 1, lastRow - 1, T_WIDTH).getValues();

  if (workCategoryWasMissing) {
    s.getRange(2, T_COL.WORK_CATEGORY, rows.length, 1).setValues(rows.map(row => [
      normalizeWorkCategory(row[T_COL.WORK_CATEGORY - 1] || inferWorkCategory(
        String(row[T_COL.TITLE - 1] || '') + ' ' + String(row[T_COL.INSTRUCTIONS - 1] || ''),
        row[T_COL.RELATED_EVENT - 1]
      ))
    ]));
  }

  // Only backfill the new related-name column when that column did not exist at all.
  // Once present, an intentionally blank value must stay blank after edits/refreshes.
  if (relatedNameWasMissing) {
    s.getRange(2, T_COL.RELATED_CLIENT_TALENT, rows.length, 1).setValues(rows.map(row => [
      String(row[T_COL.RELATED_EVENT - 1] || '').trim()
    ]));
  }

  if (parentTypeWasMissing) {
    s.getRange(2, T_COL.PARENT_TYPE, rows.length, 1).setValues(rows.map(row => {
      const src = String(row[T_COL.SOURCE - 1] || '').trim().toLowerCase();
      const relatedTicketId = String(row[T_COL.RELATED_TICKET_ID - 1] || '').trim();
      return [(src === 'answer follow-up' && relatedTicketId) ? 'Question' : ''];
    }));
  }

  if (parentIdWasMissing) {
    s.getRange(2, T_COL.PARENT_ID, rows.length, 1).setValues(rows.map(row => {
      const src = String(row[T_COL.SOURCE - 1] || '').trim().toLowerCase();
      const relatedTicketId = String(row[T_COL.RELATED_TICKET_ID - 1] || '').trim();
      return [(src === 'answer follow-up' && relatedTicketId) ? relatedTicketId : ''];
    }));
  }

  return s;
}

function ensureNotificationPreferencesSheet(ss) {
  let s = ss.getSheetByName(SHEET_NOTIFICATIONS);
  const expected = ["Email", "In App Toast", "Notification Bell", "Browser Alert", "Sound", "Email Alert", "Updated At", "Cleared Read Tickets", "Task Assigned", "Ticket New", "Ticket On Hold", "Ticket Resumed", "Ticket Answered"];
  if (!s) {
    s = ss.insertSheet(SHEET_NOTIFICATIONS);
    s.appendRow(expected);
    s.setFrozenRows(1);
  } else {
    // Safe schema extension for Phase 8.8. Existing notification preferences
    // remain untouched; the new column only stores personally-cleared read
    // notification ticket IDs.
    const headers = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), expected.length)).getValues()[0];
    for (let i = 0; i < expected.length; i++) {
      if (String(headers[i] || '').trim() !== expected[i]) s.getRange(1, i + 1).setValue(expected[i]);
    }
  }
  return s;
}


function ensureMentionNotificationsSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_MENTION_NOTIFICATIONS);
  const expected = [
    'Notification ID', 'Recipient Email', 'Recipient Name', 'Ticket ID',
    'Event / Client / Talent', 'Preview', 'Mentioned By', 'Mentioned By Email',
    'Created At', 'Read At', 'Source', 'Cleared At'
  ];
  if (!s) {
    s = ss.insertSheet(SHEET_MENTION_NOTIFICATIONS);
    s.appendRow(expected);
    s.setFrozenRows(1);
  } else {
    const headers = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), expected.length)).getValues()[0];
    for (let i = 0; i < expected.length; i++) {
      if (String(headers[i] || '').trim() !== expected[i]) s.getRange(1, i + 1).setValue(expected[i]);
    }
  }
  return s;
}

function ensureTaskNotificationsSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_TASK_NOTIFICATIONS);
  const expected = [
    'Notification ID', 'Recipient Email', 'Recipient Name', 'Task ID',
    'Task Title', 'Event Type', 'Actor Name', 'Actor Email',
    'Created At', 'Read At', 'Cleared At', 'Dedupe Key'
  ];
  if (!s) {
    s = ss.insertSheet(SHEET_TASK_NOTIFICATIONS);
    s.appendRow(expected);
    s.setFrozenRows(1);
  } else {
    const headers = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), expected.length)).getValues()[0];
    for (let i = 0; i < expected.length; i++) {
      if (String(headers[i] || '').trim() !== expected[i]) s.getRange(1, i + 1).setValue(expected[i]);
    }
  }
  return s;
}

function stripHtmlToText(value) {
  return String(value || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/\s+/g, ' ')
    .trim();
}

function compactRichContentForFeed(value) {
  const original = String(value || '');
  if (!original) return { html: '', deferred: false };
  // Table/search feeds do not need the image itself. Removing IMG tags is
  // especially important for historical rows that may still contain a large
  // base64 screenshot from before Drive-backed attachments were introduced.
  const compact = original.replace(/<img\b[^>]*>/gi, '<span>[Image attached]</span>');
  return { html: compact, deferred: compact !== original };
}

function activeMentionableMembers() {
  const byEmail = {};
  _getTeamMembersInternal()
    .filter(m => String(m.status || '').trim().toLowerCase() === 'active')
    .filter(m => !isAdminMember(m))
    .forEach(m => {
      const email = normalizeEmail(m.email);
      if (!email) return;
      if (!byEmail[email] || isSupportMember(m)) byEmail[email] = m;
    });
  return Object.keys(byEmail).map(k => byEmail[k]);
}

function extractMentionedMembers(value) {
  const text = stripHtmlToText(value);
  if (!text) return [];
  const matches = [];
  const seen = {};
  activeMentionableMembers().forEach(member => {
    const name = String(member.name || '').trim();
    if (!name) return;
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp('(^|\\s)@' + escaped + '(?=\\s|$|[.,!?;:])', 'i');
    if (pattern.test(text)) {
      const email = normalizeEmail(member.email);
      if (email && !seen[email]) {
        seen[email] = true;
        matches.push(member);
      }
    }
  });
  return matches;
}

function createMentionNotifications(ticketId, eventName, content, actor, source, excludeEmails) {
  const actorEmail = normalizeEmail(actor && actor.email);
  const excluded = {};
  (excludeEmails || []).forEach(e => { const n = normalizeEmail(e); if (n) excluded[n] = true; });
  if (actorEmail) excluded[actorEmail] = true;

  const mentioned = extractMentionedMembers(content).filter(m => !excluded[normalizeEmail(m.email)]);
  if (!mentioned.length) return [];

  const sheet = ensureMentionNotificationsSheet(SpreadsheetApp.getActiveSpreadsheet());
  const now = new Date();
  const preview = stripHtmlToText(content).slice(0, 240);
  const rows = mentioned.map(member => [
    Utilities.getUuid(), normalizeEmail(member.email), member.name, String(ticketId || ''),
    String(eventName || ''), preview, actor ? actor.name : '', actorEmail,
    now, '', String(source || 'Answer')
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  mentioned.forEach(member => {
    logAudit('MENTION_NOTIFY', actorEmail, actor ? actor.name : '', ticketId, {
      recipientEmail: normalizeEmail(member.email),
      recipientName: member.name,
      source: String(source || 'Answer')
    });
  });
  return mentioned.map(m => ({ email: normalizeEmail(m.email), name: m.name }));
}

function createDirectTicketNotification(ticketId, eventName, previewText, recipient, actor, source) {
  if (!recipient || !normalizeEmail(recipient.email) || !ticketId) return null;
  if (emailsRepresentSameWorkspaceIdentity(recipient.email, actor && actor.email)) return null;
  const eventKey = notificationEventKeyForTicketSource_(source);
  if (!eventKey || !notificationEventEnabledForEmail_(recipient.email, eventKey)) return null;
  const sheet = ensureMentionNotificationsSheet(SpreadsheetApp.getActiveSpreadsheet());
  const notificationId = Utilities.getUuid();
  sheet.appendRow([
    notificationId, normalizeEmail(recipient.email), recipient.name || '', String(ticketId),
    String(eventName || ''), stripHtmlToText(previewText).slice(0, 240),
    actor ? actor.name : '', normalizeEmail(actor && actor.email), new Date(), '',
    String(source || 'Question Assignment'), ''
  ]);
  logAudit('QUESTION_ASSIGNMENT_NOTIFY', actor && actor.email, actor && actor.name, ticketId, {
    recipientEmail: normalizeEmail(recipient.email), source: String(source || 'Question Assignment')
  });
  return notificationId;
}

function isOperationalNotificationProfile(member) {
  // An Admin profile is maintenance-only. A dual-profile Workspace email may
  // still own a separate Support profile, but notifications must be exposed
  // only while that operational profile is selected. Admin Test Mode is also
  // kept quiet so testing does not read or clear the target user's alerts.
  return !!member && !isAdminMember(member) && !member._adminTestMode;
}

function requireOperationalNotificationProfile(requestingEmail) {
  const member = requireAuthenticatedMember(requestingEmail);
  if (!isOperationalNotificationProfile(member)) {
    throw new Error('Ticket notifications are available only from an operational Support Hub profile.');
  }
  return member;
}

function getMentionNotifications(requestingEmail) {
  ensureSheetsExist();
  const member = requireAuthenticatedMember(requestingEmail);
  if (!isOperationalNotificationProfile(member)) return [];
  const email = normalizeEmail(member.email);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MENTION_NOTIFICATIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  return data
    .filter(r => normalizeEmail(r[1]) === email && !r[11])
    .map(r => ({
      notificationId: String(r[0] || ''),
      recipientEmail: normalizeEmail(r[1]),
      recipientName: String(r[2] || ''),
      ticketId: String(r[3] || ''),
      eventName: String(r[4] || ''),
      preview: String(r[5] || ''),
      mentionedBy: String(r[6] || ''),
      mentionedByEmail: normalizeEmail(r[7]),
      createdAt: r[8] instanceof Date ? r[8].toISOString() : String(r[8] || ''),
      readAt: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || ''),
      source: String(r[10] || 'Answer'),
      clearedAt: r[11] instanceof Date ? r[11].toISOString() : String(r[11] || '')
    }))
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
    .slice(0, 100);
}

function markMentionNotificationRead(notificationId, requestingEmail) {
  return withLock(() => {
    const member = requireOperationalNotificationProfile(requestingEmail);
    const email = normalizeEmail(member.email);
    const sheet = ensureMentionNotificationsSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (sheet.getLastRow() < 2) return { success: false, message: 'Notification not found.', _skipDataVersion: true };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '') !== String(notificationId || '')) continue;
      if (normalizeEmail(data[i][1]) !== email) throw new Error('Access denied: this notification belongs to another user.');
      const existing = data[i][9];
      const readAt = existing instanceof Date ? existing : new Date();
      if (!existing) {
        sheet.getRange(i + 2, 10).setValue(readAt);
        logAudit('MENTION_NOTIFICATION_READ', member.email, member.name, String(data[i][3] || ''), {
          notificationId: String(notificationId || ''),
          readAt: readAt.toISOString(),
          source: String(data[i][10] || 'Answer')
        });
      }
      return { success: true, readAt: readAt.toISOString(), _skipDataVersion: true };
    }
    return { success: false, message: 'Notification not found.', _skipDataVersion: true };
  }, { bumpDataVersion: false, operation: 'markMentionNotificationRead' });
}

function createTaskNotification(task, recipientEmail, recipientName, eventType, actorName, actorEmail, dedupeKey) {
  const recipient = normalizeEmail(recipientEmail);
  const actor = normalizeEmail(actorEmail);
  const taskId = String(task && task.taskId || '').trim();
  const type = String(eventType || 'assigned').trim().toLowerCase();
  const key = String(dedupeKey || (type + ':' + taskId + ':' + recipient)).trim();
  if (!recipient || !taskId || !key) return null;
  // Task notifications are intentionally assignment-only. Status, progress,
  // completion and cancellation remain in Task activity/history without alerts.
  if (!['assigned', 'reassigned'].includes(type)) return null;
  if (!notificationEventEnabledForEmail_(recipient, 'taskAssigned')) return null;
  if (emailsRepresentSameWorkspaceIdentity(recipient, actor)) return null;

  const sheet = ensureTaskNotificationsSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (sheet.getLastRow() >= 2) {
    const keys = sheet.getRange(2, 12, sheet.getLastRow() - 1, 1).getValues();
    if (keys.some(r => String(r[0] || '').trim() === key)) return null;
  }

  const notificationId = Utilities.getUuid();
  sheet.appendRow([
    notificationId, recipient, String(recipientName || ''), taskId,
    String(task.title || 'Task'), type, String(actorName || ''), actor,
    new Date(), '', '', key
  ]);
  logAudit('TASK_NOTIFICATION_CREATE', actor, String(actorName || ''), taskId, {
    notificationId: notificationId, recipientEmail: recipient, eventType: type
  });
  return notificationId;
}

function safeCreateTaskNotification(task, recipientEmail, recipientName, eventType, actorName, actorEmail, dedupeKey) {
  try {
    return createTaskNotification(task, recipientEmail, recipientName, eventType, actorName, actorEmail, dedupeKey);
  } catch (err) {
    try { Logger.log('Task notification persistence failed for ' + String(task && task.taskId || '') + ': ' + err); } catch (ignore) {}
    return null;
  }
}

function createTaskCreatorProgressNotification(task, eventType, actor, dedupeKey) {
  const creatorEmail = normalizeEmail(task && task.createdByEmail);
  const actorEmail = normalizeEmail(actor && actor.email);
  if (!creatorEmail || emailsRepresentSameWorkspaceIdentity(creatorEmail, actorEmail)) return null;
  safeCreateTaskNotification(
    task, creatorEmail, task.createdBy, eventType,
    actor && actor.name, actorEmail, dedupeKey
  );
  return {
    recipientEmail: creatorEmail,
    recipientName: task.createdBy,
    eventType: eventType,
    actorName: actor && actor.name,
    actorEmail: actorEmail,
    task: task
  };
}

function backfillTaskNotificationsForMember(member) {
  const email = normalizeEmail(member && member.email);
  if (!email) return;
  const marker = 'TASK_NOTIFICATION_BACKFILL_V1_' + Utilities.base64EncodeWebSafe(email).replace(/=+$/g, '');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(marker) === '1') return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    if (props.getProperty(marker) === '1') return;

    const taskSheet = ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
    const notificationSheet = ensureTaskNotificationsSheet(SpreadsheetApp.getActiveSpreadsheet());
    const existing = {};
    if (notificationSheet.getLastRow() >= 2) {
      const rows = notificationSheet.getRange(2, 1, notificationSheet.getLastRow() - 1, 12).getValues();
      rows.forEach(r => {
        if (normalizeEmail(r[1]) === email && ['assigned', 'reassigned'].includes(String(r[5] || '').toLowerCase())) {
          existing[String(r[3] || '').trim()] = true;
        }
      });
    }

    const additions = [];
    if (taskSheet.getLastRow() >= 2) {
      const taskRows = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, T_WIDTH).getValues();
      taskRows.forEach(row => {
        const task = taskRowToObject(row);
        if (task.assignedToEmail !== email || existing[task.taskId]) return;
        if (task.status === TASK_STATUS_COMPLETED || task.status === TASK_STATUS_CANCELLED) return;
        if (emailsRepresentSameWorkspaceIdentity(task.createdByEmail, email)) return;
        additions.push([
          Utilities.getUuid(), email, task.assignedTo || member.name || '', task.taskId,
          task.title || 'Task', 'assigned', task.createdBy || 'Team member', task.createdByEmail || '',
          task.createdAt ? new Date(task.createdAt) : new Date(), '', '',
          'backfill-assigned:' + task.taskId + ':' + email
        ]);
      });
    }
    if (additions.length) {
      notificationSheet.getRange(notificationSheet.getLastRow() + 1, 1, additions.length, 12).setValues(additions);
    }
    props.setProperty(marker, '1');
  } finally {
    lock.releaseLock();
  }
}

function getTaskNotifications(requestingEmail) {
  ensureSheetsExist();
  const member = requireAuthenticatedMember(requestingEmail);
  if (!isOperationalNotificationProfile(member)) return [];
  backfillTaskNotificationsForMember(member);
  const email = normalizeEmail(member.email);
  const sheet = ensureTaskNotificationsSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues()
    .filter(r => normalizeEmail(r[1]) === email && !r[10])
    .map(r => ({
      notificationId: String(r[0] || ''), recipientEmail: normalizeEmail(r[1]), recipientName: String(r[2] || ''),
      taskId: String(r[3] || ''), taskTitle: String(r[4] || ''), eventType: String(r[5] || 'assigned').toLowerCase(),
      actorName: String(r[6] || ''), actorEmail: normalizeEmail(r[7]),
      createdAt: r[8] instanceof Date ? r[8].toISOString() : String(r[8] || ''),
      readAt: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || ''),
      clearedAt: r[10] instanceof Date ? r[10].toISOString() : String(r[10] || '')
    }))
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
    .slice(0, 100);
}

function getPersistentNotifications(requestingEmail) {
  return {
    mentions: getMentionNotifications(requestingEmail),
    tasks: getTaskNotifications(requestingEmail)
  };
}

function markTaskNotificationRead(notificationId, requestingEmail) {
  return withLock(() => {
    const member = requireOperationalNotificationProfile(requestingEmail);
    const email = normalizeEmail(member.email);
    const sheet = ensureTaskNotificationsSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (sheet.getLastRow() < 2) return { success: false, message: 'Notification not found.', _skipDataVersion: true };
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0] || '') !== String(notificationId || '')) continue;
      if (normalizeEmail(rows[i][1]) !== email) throw new Error('Access denied: this notification belongs to another user.');
      const existing = rows[i][9];
      const readAt = existing instanceof Date ? existing : new Date();
      if (!existing) {
        sheet.getRange(i + 2, 10).setValue(readAt);
        logAudit('TASK_NOTIFICATION_READ', member.email, member.name, String(rows[i][3] || ''), {
          notificationId: String(notificationId || ''), eventType: String(rows[i][5] || '')
        });
      }
      return { success: true, readAt: readAt.toISOString(), _skipDataVersion: true };
    }
    return { success: false, message: 'Notification not found.', _skipDataVersion: true };
  }, { bumpDataVersion: false });
}

function sendMentionEmailIfEnabled(recipientEmail, recipientName, eventName, preview, actorName, actorEmail, sourceLabel, ticketId) {
  try {
    const recipient = normalizeEmail(recipientEmail);
    if (!recipient) return;
    const prefs = readNotificationPreferencesForEmail(recipient);
    if (!prefs.email || prefs.ticketNew === false) return;

    // Email notifications are assignment-only. @mentions, notes, answers, and
    // status/progress updates remain in-app and do not send email.
    const isAssignment = /question assignment|routed question|automatic question routing/i.test(String(sourceLabel || ''));
    if (!isAssignment) return;

    // Admin Test Mode is identified by the real Workspace account differing
    // from the profile it is acting as.
    try {
      const authenticatedEmail = normalizeEmail(Session.getActiveUser().getEmail());
      const effectiveActorEmail = normalizeEmail(actorEmail);
      if (authenticatedEmail && effectiveActorEmail && !emailsRepresentSameWorkspaceIdentity(authenticatedEmail, effectiveActorEmail)) return;
    } catch (identityError) {}

    const name = stripNoraPrefix(String(recipientName || '').trim()) || 'there';
    const actor = stripNoraPrefix(String(actorName || 'A team member').trim()) || 'A team member';
    const ticket = ticketId ? ticketEmailSnapshot_(ticketId) : null;
    const eventLabel = String((ticket && ticket.eventName) || eventName || 'Support Hub ticket').trim();
    const questionText = String((ticket && ticket.question) || preview || '');
    const directUrl = ticketId ? supportHubDeepLink_('ticket', ticketId) : '';
    const subject = 'Support Hub: Ticket assigned to you - ' + eventLabel;
    const body =
      'Hi ' + name + ',\n\n' + actor + ' assigned a Support Hub ticket to you for ' + eventLabel + '.\n\n' +
      'Question / Instructions: ' + stripHtmlToText(questionText) + '\n' +
      (ticket && ticket.category ? 'Work Category: ' + ticket.category + '\n' : '') +
      (ticket && ticket.eventDate ? 'Event Date: ' + ticket.eventDate + '\n' : '') +
      (ticket && ticket.talentRole ? 'Event Role: ' + ticket.talentRole + '\n' : '') +
      (ticket && ticket.caseLink ? 'Case / Event Link: ' + ticket.caseLink + '\n' : '') +
      (directUrl ? '\nOpen Ticket: ' + directUrl + '\n' : '') +
      '\nSupport Hub';
    const htmlBody = ticketAssignmentEmailHtml_(name, actor, ticket || { eventName: eventLabel, question: questionText }, directUrl);
    sendSupportHubEmail(recipient, subject, body, actorEmail, actor, htmlBody);
  } catch (err) {
    Logger.log('Ticket assignment email failed for ' + normalizeEmail(recipientEmail) + ': ' + err);
  }
}


function defaultNotificationEventPreferences() {
  return { taskAssigned: true, ticketNew: true, ticketOnHold: true, ticketResumed: true, ticketAnswered: true };
}

function defaultNotificationPreferences() {
  return Object.assign({ toast: true, bell: true, browser: false, sound: false, email: false }, defaultNotificationEventPreferences());
}

function boolPref(value) {
  if (value === true || value === 1) return true;
  return ["true", "yes", "1", "on"].includes(String(value || "").trim().toLowerCase());
}

function readNotificationPreferencesForEmail(email) {
  ensureSheetsExist();
  const normalized = normalizeEmail(email);
  const defaults = defaultNotificationPreferences();
  if (!normalized) return defaults;
  const s = ensureNotificationPreferencesSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (!s || s.getLastRow() < 2) return defaults;
  const width = Math.max(13, s.getLastColumn());
  const data = s.getRange(2, 1, s.getLastRow() - 1, width).getValues();
  for (let i = 0; i < data.length; i++) {
    if (normalizeEmail(data[i][0]) === normalized) {
      return {
        toast: boolPref(data[i][1]), bell: boolPref(data[i][2]), browser: boolPref(data[i][3]),
        sound: boolPref(data[i][4]), email: boolPref(data[i][5]),
        taskAssigned: data[i][8] === '' || data[i][8] == null ? defaults.taskAssigned : boolPref(data[i][8]),
        ticketNew: data[i][9] === '' || data[i][9] == null ? defaults.ticketNew : boolPref(data[i][9]),
        ticketOnHold: data[i][10] === '' || data[i][10] == null ? defaults.ticketOnHold : boolPref(data[i][10]),
        ticketResumed: data[i][11] === '' || data[i][11] == null ? defaults.ticketResumed : boolPref(data[i][11]),
        ticketAnswered: data[i][12] === '' || data[i][12] == null ? defaults.ticketAnswered : boolPref(data[i][12])
      };
    }
  }
  return defaults;
}

function notificationEventKeyForTicketSource_(source) {
  const value = String(source || '').trim().toLowerCase();
  if (/question assignment|automatic question routing|reopened follow-up/.test(value)) return 'ticketNew';
  if (/ticket on hold|put on hold/.test(value)) return 'ticketOnHold';
  if (/ticket resumed|resume from hold|resumed from hold/.test(value)) return 'ticketResumed';
  return '';
}

function notificationEventEnabledForEmail_(email, eventKey) {
  if (!eventKey) return false;
  const prefs = readNotificationPreferencesForEmail(email);
  return prefs[eventKey] !== false;
}

function operationalMemberForTicketRequester_(askedByEmail, askedByName) {
  const email = normalizeEmail(askedByEmail);
  const name = stripNoraPrefix(String(askedByName || '')).trim().toLowerCase();
  const members = _getTeamMembersInternal().filter(function(m) {
    return String(m.status || '').trim().toLowerCase() === 'active' && !isAdminMember(m);
  });
  if (email) {
    const matches = members.filter(function(m) { return normalizeEmail(m.email) === email; });
    if (matches.length) return matches.find(isSupportMember) || matches[0];
  }
  if (name) {
    const matches = members.filter(function(m) { return stripNoraPrefix(String(m.name || '')).trim().toLowerCase() === name; });
    if (matches.length === 1) return matches[0];
  }
  return null;
}

function parseClearedReadTicketIds(value) {
  const raw = String(value || '').trim();
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed.map(v => String(v || '').trim()).filter(Boolean);
  } catch (e) {}
  // Backward-compatible fallback if the cell was ever manually populated.
  return raw.split(/[|,\n]+/).map(v => String(v || '').trim()).filter(Boolean);
}

function readClearedReadTicketIdsForEmail(email) {
  ensureSheetsExist();
  const normalized = normalizeEmail(email);
  if (!normalized) return [];
  const s = ensureNotificationPreferencesSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (s.getLastRow() < 2) return [];
  const data = s.getRange(2, 1, s.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < data.length; i++) {
    if (normalizeEmail(data[i][0]) === normalized) return parseClearedReadTicketIds(data[i][7]);
  }
  return [];
}

function getNotificationPreferences(requestingEmail) {
  const member = requireAuthenticatedMember(requestingEmail);
  if (!isOperationalNotificationProfile(member)) {
    return {
      success: true,
      notificationsSuppressed: true,
      preferences: { toast: false, bell: false, browser: false, sound: false, email: false, taskAssigned: false, ticketNew: false, ticketOnHold: false, ticketResumed: false, ticketAnswered: false },
      minimumRequired: 0,
      clearedReadTicketIds: []
    };
  }
  const prefs = readNotificationPreferencesForEmail(member.email);
  return {
    success: true,
    preferences: prefs,
    minimumRequired: 2,
    clearedReadTicketIds: readClearedReadTicketIdsForEmail(member.email)
  };
}

function clearReadNotifications(requestingEmail) {
  return withLock(() => {
    const member = requireOperationalNotificationProfile(requestingEmail);
    const email = normalizeEmail(member.email);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aSheet = ss.getSheetByName(SHEET_ANSWERED);
    const prefsSheet = ensureNotificationPreferencesSheet(ss);

    // Determine eligible IDs server-side. Clearing a bell notification never
    // deletes an Answered Ticket and never changes its Read/Unread state.
    const eligible = [];
    if (aSheet && aSheet.getLastRow() >= 2) {
      const rows = aSheet.getRange(2, 1, aSheet.getLastRow() - 1, A_WIDTH).getValues();
      const memberName = stripNoraPrefix(String(member.name || '')).trim().toLowerCase();
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const askerEmail = normalizeEmail(row[A_COL.ASKED_BY_EMAIL - 1]);
        const askerName = stripNoraPrefix(String(row[A_COL.ASKED_BY - 1] || '')).trim().toLowerCase();
        const isOwn = askerEmail ? askerEmail === email : (!!askerName && askerName === memberName);
        if (!isOwn) continue;

        const readBy = parseReadByList(row[A_COL.READ_BY - 1]);
        const isRead = readBy.includes('*') || readBy.map(normalizeEmail).includes(email);
        if (!isRead) continue;

        const ticketId = String(row[A_COL.TICKET_ID - 1] || '').trim();
        if (ticketId) eligible.push(ticketId);
      }
    }

    // Read @mention notifications are cleared from the bell too, while their
    // historical row/read timestamp remains intact for audit purposes.
    let clearedMentionCount = 0;
    const mentionSheet = ensureMentionNotificationsSheet(ss);
    if (mentionSheet && mentionSheet.getLastRow() >= 2) {
      const mentionRows = mentionSheet.getRange(2, 1, mentionSheet.getLastRow() - 1, 12).getValues();
      let mentionChanged = false;
      const clearedAt = new Date();
      for (let i = 0; i < mentionRows.length; i++) {
        if (normalizeEmail(mentionRows[i][1]) !== email) continue;
        if (!mentionRows[i][9] || mentionRows[i][11]) continue;
        mentionRows[i][11] = clearedAt;
        clearedMentionCount++;
        mentionChanged = true;
      }
      if (mentionChanged) {
        mentionSheet.getRange(2, 12, mentionRows.length, 1).setValues(mentionRows.map(r => [r[11] || '']));
      }
    }

    // Task notification rows are retained for history; clearing only removes
    // already-read items from this user's bell.
    let clearedTaskCount = 0;
    const taskNotificationSheet = ensureTaskNotificationsSheet(ss);
    if (taskNotificationSheet && taskNotificationSheet.getLastRow() >= 2) {
      const taskNotificationRows = taskNotificationSheet.getRange(2, 1, taskNotificationSheet.getLastRow() - 1, 12).getValues();
      let taskNotificationChanged = false;
      const clearedAt = new Date();
      for (let i = 0; i < taskNotificationRows.length; i++) {
        if (normalizeEmail(taskNotificationRows[i][1]) !== email) continue;
        if (!taskNotificationRows[i][9] || taskNotificationRows[i][10]) continue;
        taskNotificationRows[i][10] = clearedAt;
        clearedTaskCount++;
        taskNotificationChanged = true;
      }
      if (taskNotificationChanged) {
        taskNotificationSheet.getRange(2, 11, taskNotificationRows.length, 1).setValues(taskNotificationRows.map(r => [r[10] || '']));
      }
    }

    const existing = new Set(readClearedReadTicketIdsForEmail(email));
    eligible.forEach(id => existing.add(id));
    // Keep a generous bounded history so the preference cell cannot grow
    // indefinitely over years of use. Most-recent eligible IDs are retained.
    const merged = Array.from(existing);
    const kept = merged.slice(Math.max(0, merged.length - 1000));

    let targetRow = -1;
    const lastRow = prefsSheet.getLastRow();
    if (lastRow >= 2) {
      const emails = prefsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) {
        if (normalizeEmail(emails[i][0]) === email) { targetRow = i + 2; break; }
      }
    }
    if (targetRow === -1) {
      const d = defaultNotificationPreferences();
      prefsSheet.appendRow([email, d.toast, d.bell, d.browser, d.sound, d.email, new Date(), JSON.stringify(kept), d.taskAssigned, d.ticketNew, d.ticketOnHold, d.ticketResumed, d.ticketAnswered]);
    } else {
      prefsSheet.getRange(targetRow, 8).setValue(JSON.stringify(kept));
      prefsSheet.getRange(targetRow, 7).setValue(new Date());
    }

    const totalCleared = eligible.length + clearedMentionCount + clearedTaskCount;
    logAudit('NOTIFICATIONS_CLEAR_READ', member.email, member.name, '', { clearedAnsweredCount: eligible.length, clearedMentionCount: clearedMentionCount, clearedTaskCount: clearedTaskCount, clearedCount: totalCleared });
    return { success: true, clearedCount: totalCleared, clearedAnsweredCount: eligible.length, clearedMentionCount: clearedMentionCount, clearedTaskCount: clearedTaskCount, clearedReadTicketIds: kept };
  }, { bumpDataVersion: false });
}

function saveNotificationPreferences(requestingEmail, preferences) {
  return withLock(() => {
    const member = requireOperationalNotificationProfile(requestingEmail);
    const p = preferences || {};
    const clean = { toast: !!p.toast, bell: !!p.bell, browser: !!p.browser, sound: !!p.sound, email: !!p.email };
    const enabledCount = Object.keys(clean).reduce((n, k) => n + (clean[k] ? 1 : 0), 0);
    if (enabledCount < 2) throw new Error("Please keep at least two notification methods enabled.");

    const existingPrefs = readNotificationPreferencesForEmail(member.email);
    const events = {
      taskAssigned: existingPrefs.taskAssigned !== false,
      ticketNew: existingPrefs.ticketNew !== false,
      ticketOnHold: existingPrefs.ticketOnHold !== false,
      ticketResumed: existingPrefs.ticketResumed !== false,
      ticketAnswered: existingPrefs.ticketAnswered !== false
    };
    const s = ensureNotificationPreferencesSheet(SpreadsheetApp.getActiveSpreadsheet());
    const lastRow = s.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const emails = s.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) if (normalizeEmail(emails[i][0]) === member.email) { targetRow = i + 2; break; }
    }
    const cleared = targetRow === -1 ? '' : s.getRange(targetRow, 8).getValue();
    const row = [member.email, clean.toast, clean.bell, clean.browser, clean.sound, clean.email, new Date(), cleared,
      events.taskAssigned, events.ticketNew, events.ticketOnHold, events.ticketResumed, events.ticketAnswered];
    if (targetRow === -1) s.appendRow(row);
    else s.getRange(targetRow, 1, 1, row.length).setValues([row]);

    logAudit('NOTIFICATION_PREFS', member.email, member.name, '', Object.assign({}, clean, events));
    return { success: true, preferences: Object.assign({}, clean, events), minimumRequired: 2 };
  }, { bumpDataVersion: false });
}

function requireNotificationAdmin_(requestingEmail) {
  const member = requireAuthenticatedMember(requestingEmail);
  if (!isAdminMember(member)) throw new Error('Admin access is required to manage notification controls.');
  return member;
}

function getAdminNotificationControlData(requestingEmail) {
  const admin = requireNotificationAdmin_(requestingEmail);
  const byEmail = {};
  _getTeamMembersInternal().filter(function(m) {
    return String(m.status || '').trim().toLowerCase() === 'active' && !isAdminMember(m) && normalizeEmail(m.email);
  }).forEach(function(m) {
    const email = normalizeEmail(m.email);
    // A dual profile should appear once. Prefer its operational Support profile.
    if (!byEmail[email] || isSupportMember(m)) byEmail[email] = m;
  });
  const members = Object.keys(byEmail).map(function(email) {
    const m = byEmail[email];
    const prefs = readNotificationPreferencesForEmail(email);
    return {
      email: email, name: m.name || email, title: m.title || '', category: m.category || '',
      events: {
        taskAssigned: prefs.taskAssigned !== false, ticketNew: prefs.ticketNew !== false,
        ticketOnHold: prefs.ticketOnHold !== false, ticketResumed: prefs.ticketResumed !== false,
        ticketAnswered: prefs.ticketAnswered !== false
      }
    };
  }).sort(function(a,b) { return String(a.name).localeCompare(String(b.name)); });
  return { success: true, members: members, titles: Array.from(new Set(members.map(function(m){return m.title;}).filter(Boolean))).sort(), admin: admin.name };
}

function saveAdminNotificationControls(requestingEmail, targetEmails, eventPreferences) {
  return withLock(() => {
    const admin = requireNotificationAdmin_(requestingEmail);
    const requested = Array.isArray(targetEmails) ? targetEmails.map(normalizeEmail).filter(Boolean) : [];
    const unique = Array.from(new Set(requested));
    if (!unique.length) throw new Error('Select at least one team member.');
    const eventDefaults = defaultNotificationEventPreferences();
    const input = eventPreferences || {};
    const events = {
      taskAssigned: input.taskAssigned === undefined ? eventDefaults.taskAssigned : !!input.taskAssigned,
      ticketNew: input.ticketNew === undefined ? eventDefaults.ticketNew : !!input.ticketNew,
      ticketOnHold: input.ticketOnHold === undefined ? eventDefaults.ticketOnHold : !!input.ticketOnHold,
      ticketResumed: input.ticketResumed === undefined ? eventDefaults.ticketResumed : !!input.ticketResumed,
      ticketAnswered: input.ticketAnswered === undefined ? eventDefaults.ticketAnswered : !!input.ticketAnswered
    };
    const allowed = {};
    _getTeamMembersInternal().forEach(function(m) {
      if (String(m.status || '').trim().toLowerCase() === 'active' && !isAdminMember(m)) allowed[normalizeEmail(m.email)] = true;
    });
    const targets = unique.filter(function(email) { return allowed[email]; });
    if (!targets.length) throw new Error('No active operational team members were selected.');
    const sheet = ensureNotificationPreferencesSheet(SpreadsheetApp.getActiveSpreadsheet());
    const rows = sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(13, sheet.getLastColumn())).getValues() : [];
    const rowByEmail = {};
    rows.forEach(function(row, i) { rowByEmail[normalizeEmail(row[0])] = { row: row, index: i + 2 }; });
    targets.forEach(function(email) {
      const current = readNotificationPreferencesForEmail(email);
      const existing = rowByEmail[email];
      const cleared = existing ? existing.row[7] : '';
      const row = [email, current.toast, current.bell, current.browser, current.sound, current.email, new Date(), cleared,
        events.taskAssigned, events.ticketNew, events.ticketOnHold, events.ticketResumed, events.ticketAnswered];
      if (existing) sheet.getRange(existing.index, 1, 1, row.length).setValues([row]);
      else sheet.appendRow(row);
    });
    logAudit('ADMIN_NOTIFICATION_CONTROLS', admin.email, admin.name, '', { targetEmails: targets, events: events, count: targets.length });
    return { success: true, updatedCount: targets.length, events: events };
  }, { bumpDataVersion: false });
}

function sendSupportHubEmail(recipient, subject, body, actorEmail, actorName, htmlBody) {
  const cleanRecipient = normalizeEmail(recipient);
  if (!cleanRecipient) return;

  const senderEmail = normalizeEmail(actorEmail);
  const senderName = stripNoraPrefix(String(actorName || 'Support Hub').trim()) || 'Support Hub';
  const options = {
    name: senderName + ' via Support Hub'
  };
  if (senderEmail) options.replyTo = senderEmail;
  if (htmlBody) options.htmlBody = htmlBody;

  // Apps Script can only use a real From address when that address is an
  // approved Gmail alias of the account executing the web app. When the
  // logged-in Support Hub user's email is such an alias, send from it.
  // Otherwise preserve the actual sender account, but show the logged-in
  // user's name and route replies directly to their email.
  try {
    const aliases = GmailApp.getAliases().map(normalizeEmail);
    if (senderEmail && aliases.indexOf(senderEmail) !== -1) {
      options.from = senderEmail;
    }
    GmailApp.sendEmail(cleanRecipient, subject, body, options);
  } catch (gmailErr) {
    // Fallback keeps notifications working even if Gmail scope/alias lookup
    // is unavailable in an older deployment.
    const mailOptions = {
      to: cleanRecipient,
      subject: subject,
      body: body,
      name: options.name
    };
    if (senderEmail) mailOptions.replyTo = senderEmail;
    if (htmlBody) mailOptions.htmlBody = htmlBody;
    MailApp.sendEmail(mailOptions);
  }
}

function emailEscapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function emailTextHtml_(value) {
  return emailEscapeHtml_(stripHtmlToText(String(value || ''))).replace(/\n/g, '<br>');
}

function supportHubDeepLink_(kind, id) {
  try {
    const base = ScriptApp.getService().getUrl();
    if (!base || !id) return '';
    return base + '#' + String(kind || '').toLowerCase() + '=' + encodeURIComponent(String(id));
  } catch (e) { return ''; }
}

function emailDetailRow_(label, value, isHtml) {
  if (value === '' || value === null || value === undefined || value === false) return '';
  const rendered = isHtml ? String(value) : emailTextHtml_(value);
  return '<tr><td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;width:180px;vertical-align:top;font-weight:700;color:#475569">' + emailEscapeHtml_(label) + '</td>' +
    '<td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;color:#0f172a;vertical-align:top">' + rendered + '</td></tr>';
}

function ticketEmailSnapshot_(ticketId) {
  try {
    const target = requireQuestionRow(ticketId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    return {
      ticketId: String(row[Q_COL.TICKET_ID - 1] || ticketId || ''),
      eventName: String(row[Q_COL.EVENT - 1] || ''),
      question: String(row[Q_COL.QUESTION - 1] || ''),
      askedBy: stripNoraPrefix(String(row[Q_COL.ASKED_BY - 1] || '')),
      askedByEmail: normalizeEmail(row[Q_COL.ASKED_BY_EMAIL - 1]),
      created: row[Q_COL.CREATED - 1] ? formatCSTStamp(row[Q_COL.CREATED - 1]) + ' CST' : '',
      caseLink: String(row[Q_COL.LINK - 1] || ''),
      status: String(row[Q_COL.STATUS - 1] || STATUS_OPEN),
      assignedToEmail: normalizeEmail(row[Q_COL.ASSIGNED - 1]),
      holdReason: String(row[Q_COL.HOLD_REASON - 1] || ''),
      followUp: isFollowUpValue(row[Q_COL.FOLLOWUP - 1]),
      assignmentNotes: String(row[Q_COL.ASSIGN_NOTES - 1] || ''),
      category: normalizeWorkCategory(row[Q_COL.WORK_CATEGORY - 1] || inferWorkCategory(row[Q_COL.QUESTION - 1], row[Q_COL.EVENT - 1])),
      eventThisWeekend: row[Q_COL.EVENT_THIS_WEEKEND - 1] === true,
      eventDate: normalizeEventDate(row[Q_COL.EVENT_DATE - 1]),
      talentRole: String(row[Q_COL.TALENT_ROLE - 1] || ''),
      routedSupportEmail: normalizeEmail(row[Q_COL.ROUTED_SUPPORT_EMAIL - 1]),
      attentionToday: isAttentionActive(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_UNTIL - 1]),
      attentionReason: String(row[Q_COL.ATTENTION_REASON - 1] || ''),
      attentionUntil: normalizeEventDate(row[Q_COL.ATTENTION_UNTIL - 1])
    };
  } catch (e) { return null; }
}

function ticketAssignmentEmailHtml_(recipientName, actorName, ticket, directUrl) {
  ticket = ticket || {};
  let rows = '';
  rows += emailDetailRow_('Event / Client / Talent', ticket.eventName);
  rows += emailDetailRow_('Question / Instructions', ticket.question);
  rows += emailDetailRow_('Asked By', ticket.askedBy + (ticket.askedByEmail ? ' (' + ticket.askedByEmail + ')' : ''));
  rows += emailDetailRow_('Submitted', ticket.created);
  rows += emailDetailRow_('Work Category', ticket.category);
  rows += emailDetailRow_('Event Role', ticket.talentRole);
  rows += emailDetailRow_('Event Date', ticket.eventDate);
  rows += emailDetailRow_('Status', ticket.status);
  rows += emailDetailRow_('Follow-Up', ticket.followUp ? 'Yes' : '');
  rows += emailDetailRow_('Event This Weekend', ticket.eventThisWeekend ? 'Yes' : '');
  rows += emailDetailRow_('Attention Needed Today', ticket.attentionToday ? 'Yes' : '');
  rows += emailDetailRow_('Attention Reason', ticket.attentionReason);
  rows += emailDetailRow_('Attention Until', ticket.attentionUntil);
  rows += emailDetailRow_('On Hold Reason', ticket.holdReason);
  rows += emailDetailRow_('Assignment Notes', ticket.assignmentNotes);
  if (ticket.caseLink) {
    const href = /^https?:\/\//i.test(ticket.caseLink) ? ticket.caseLink : '';
    rows += emailDetailRow_('Case / Event Link', href ? '<a href="' + emailEscapeHtml_(href) + '" style="color:#2563eb">Open related record</a>' : ticket.caseLink, !!href);
  }
  return '<div style="font-family:Arial,sans-serif;max-width:760px;color:#0f172a">' +
    '<p>Hi ' + emailEscapeHtml_(recipientName || 'there') + ',</p>' +
    '<p><strong>' + emailEscapeHtml_(actorName || 'A team member') + '</strong> assigned a Support Hub ticket to you.</p>' +
    '<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:8px">' + rows + '</table>' +
    (directUrl ? '<p style="margin-top:20px"><a href="' + emailEscapeHtml_(directUrl) + '" style="display:inline-block;background:#0f172a;color:#fff;text-decoration:none;padding:10px 16px;border-radius:8px;font-weight:700">Open Ticket in Support Hub</a></p>' : '') +
    '<p style="color:#64748b;font-size:12px">This email is sent for assignment only. Status updates and notes stay inside Support Hub.</p></div>';
}

function taskAssignmentEmailHtml_(recipientName, actorName, task, directUrl, intro) {
  let rows = '';
  rows += emailDetailRow_('Task', task.title);
  rows += emailDetailRow_('Instructions', task.instructions);
  rows += emailDetailRow_('Assigned To', (task.assignedTo || '') + (task.assignedToEmail ? ' (' + task.assignedToEmail + ')' : ''));
  rows += emailDetailRow_('Created By', (task.createdBy || '') + (task.createdByEmail ? ' (' + task.createdByEmail + ')' : ''));
  rows += emailDetailRow_('Created', task.createdAt);
  rows += emailDetailRow_('Due Date', task.dueDate);
  rows += emailDetailRow_('Status', task.status);
  rows += emailDetailRow_('Work Category', normalizeWorkCategory(task.category));
  rows += emailDetailRow_('Related Event / Client / Talent', task.relatedClientTalent || task.relatedEvent);
  rows += emailDetailRow_('Source', task.source);
  rows += emailDetailRow_('Support Owner', task.supportOwnerEmail);
  if (task.relatedLink) {
    const href = /^https?:\/\//i.test(task.relatedLink) ? task.relatedLink : '';
    rows += emailDetailRow_('Related Link', href ? '<a href="' + emailEscapeHtml_(href) + '" style="color:#2563eb">Open related record</a>' : task.relatedLink, !!href);
  }
  if (task.relatedTicketId) {
    const ticketUrl = supportHubDeepLink_('ticket', task.relatedTicketId);
    rows += emailDetailRow_('Parent Ticket', ticketUrl ? '<a href="' + emailEscapeHtml_(ticketUrl) + '" style="color:#2563eb">Open parent ticket</a>' : 'Connected ticket', !!ticketUrl);
  }
  if (String(task.parentType || '').toLowerCase() === 'task' && task.parentId) {
    const parentTaskUrl = supportHubDeepLink_('task', task.parentId);
    rows += emailDetailRow_('Parent Task', parentTaskUrl ? '<a href="' + emailEscapeHtml_(parentTaskUrl) + '" style="color:#2563eb">Open parent task</a>' : 'Connected task', !!parentTaskUrl);
  }
  return '<div style="font-family:Arial,sans-serif;max-width:760px;color:#0f172a">' +
    '<p>Hi ' + emailEscapeHtml_(recipientName || 'there') + ',</p>' +
    '<p>' + emailEscapeHtml_(intro || ((actorName || 'A team member') + ' assigned a task to you.')) + '</p>' +
    '<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:8px">' + rows + '</table>' +
    (directUrl ? '<p style="margin-top:20px"><a href="' + emailEscapeHtml_(directUrl) + '" style="display:inline-block;background:#0f172a;color:#fff;text-decoration:none;padding:10px 16px;border-radius:8px;font-weight:700">Open Task in Support Hub</a></p>' : '') +
    '<p style="color:#64748b;font-size:12px">This email is sent for assignment only. Status updates and notes stay inside Support Hub.</p></div>';
}

function sendAnswerEmailIfEnabled(askedByEmail, askedByName, eventName, answerText, answeredByName, answeredByEmail) {
  // Email is intentionally assignment-only. Answer/status/note activity stays
  // in the in-app notification center and the ticket history.
  return;
}


function sendTaskEmailIfEnabled(recipientEmail, recipientName, eventType, task, actorName, actorEmail) {
  // Email is assignment-only. Progress notes, status changes, completion, and
  // cancellation remain visible in Support Hub without generating email.
  if (eventType !== 'assigned' && eventType !== 'reassigned') return;

  // Admin Test Mode never sends external email as the impersonated user.
  try {
    const authEmail = normalizeEmail(Session.getActiveUser().getEmail());
    if (authEmail && normalizeEmail(actorEmail) && !emailsRepresentSameWorkspaceIdentity(authEmail, actorEmail)) return;
  } catch (e) {}
  const recipient = normalizeEmail(recipientEmail);
  if (!recipient || !task) return;
  try {
    const prefs = readNotificationPreferencesForEmail(recipient);
    if (!prefs.email || prefs.taskAssigned === false) return;

    const name = stripNoraPrefix(String(recipientName || '').trim()) || 'there';
    const actor = stripNoraPrefix(String(actorName || 'a team member').trim());
    const title = String(task.title || 'Task').trim();
    const due = String(task.dueDate || '').trim();
    const reassigned = eventType === 'reassigned';
    const subject = reassigned ? 'Support Hub: Task reassigned to you - ' + title : 'Support Hub: New task assigned - ' + title;
    const intro = reassigned ? ('A task has been reassigned to you by ' + actor + '.') : ('A new task has been assigned to you by ' + actor + '.');
    const directUrl = supportHubDeepLink_('task', task.taskId);

    const body =
      'Hi ' + name + ',\n\n' + intro + '\n\n' +
      'Task: ' + title + '\n' +
      'Instructions: ' + stripHtmlToText(task.instructions || '') + '\n' +
      'Category: ' + normalizeWorkCategory(task.category) + '\n' +
      (due ? 'Due Date: ' + due + '\n' : '') +
      ((task.relatedClientTalent || task.relatedEvent) ? 'Related Event / Client / Talent: ' + String(task.relatedClientTalent || task.relatedEvent) + '\n' : '') +
      (task.relatedLink ? 'Related Link: ' + String(task.relatedLink) + '\n' : '') +
      (directUrl ? '\nOpen Task: ' + directUrl + '\n' : '') +
      '\nSupport Hub';

    const htmlBody = taskAssignmentEmailHtml_(name, actor, task, directUrl, intro);
    sendSupportHubEmail(recipient, subject, body, actorEmail, actor, htmlBody);
  } catch (err) {
    Logger.log('Task assignment email failed for ' + recipient + ': ' + err);
  }
}


function ensureQuestionsSchema(sheet) {
  let lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headers[Q_COL.TICKET_ID - 1] !== 'Ticket ID') {
    sheet.getRange(1, Q_COL.TICKET_ID, 1, 5).setValues([
      ['Ticket ID', 'Hold Reason', 'Hold Since', 'Cumulative Hold Hours', 'Asked By Email']
    ]);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const idRange = sheet.getRange(2, Q_COL.TICKET_ID, lastRow - 1, 1);
      idRange.setValues(idRange.getValues().map(r => [r[0] ? r[0] : Utilities.getUuid()]));
      const accumRange = sheet.getRange(2, Q_COL.HOLD_ACCUM, lastRow - 1, 1);
      accumRange.setValues(accumRange.getValues().map(r => [(r[0] === '' || r[0] === null || isNaN(r[0])) ? 0 : r[0]]));
    }
  }

  lastCol = Math.max(sheet.getLastColumn(), 1);
  headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers[Q_COL.FOLLOWUP - 1] !== 'Is Follow Up') {
    sheet.getRange(1, Q_COL.FOLLOWUP).setValue('Is Follow Up');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, Q_COL.FOLLOWUP, lastRow - 1, 1);
      range.setValues(range.getValues().map(r => [r[0] === true]));
    }
  }

  lastCol = Math.max(sheet.getLastColumn(), 1);
  headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers[Q_COL.ASSIGN_NOTES - 1] !== 'Assignment Notes') {
    sheet.getRange(1, Q_COL.ASSIGN_NOTES).setValue('Assignment Notes');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, Q_COL.ASSIGN_NOTES, lastRow - 1, 1);
      range.setValues(range.getValues().map(r => [r[0] || '']));
    }
  }

  lastCol = Math.max(sheet.getLastColumn(), 1);
  headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers[Q_COL.STATUS_CHANGED_AT - 1] !== 'Status Changed At') {
    sheet.getRange(1, Q_COL.STATUS_CHANGED_AT, 1, 2).setValues([['Status Changed At', 'Status Changed By']]);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, Q_WIDTH).getValues();
      sheet.getRange(2, Q_COL.STATUS_CHANGED_AT, lastRow - 1, 2).setValues(rows.map(row => [row[Q_COL.CREATED - 1] || '', row[Q_COL.ASKED_BY - 1] || '']));
    }
  }

  const workCategoryWasMissing = String(sheet.getRange(1, Q_COL.WORK_CATEGORY).getValue() || '').trim() !== 'Work Category';
  const weekendWasMissing = String(sheet.getRange(1, Q_COL.EVENT_THIS_WEEKEND).getValue() || '').trim() !== 'Event This Weekend';
  const eventDateWasMissing = String(sheet.getRange(1, Q_COL.EVENT_DATE).getValue() || '').trim() !== 'Event Date';
  if (workCategoryWasMissing || weekendWasMissing || eventDateWasMissing) {
    sheet.getRange(1, Q_COL.WORK_CATEGORY, 1, 3).setValues([['Work Category', 'Event This Weekend', 'Event Date']]);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, Q_WIDTH).getValues();
      if (workCategoryWasMissing) {
        sheet.getRange(2, Q_COL.WORK_CATEGORY, lastRow - 1, 1).setValues(rows.map(row => [
          normalizeWorkCategory(row[Q_COL.WORK_CATEGORY - 1] || inferWorkCategory(row[Q_COL.QUESTION - 1], row[Q_COL.EVENT - 1]))
        ]));
      }
      if (weekendWasMissing) {
        sheet.getRange(2, Q_COL.EVENT_THIS_WEEKEND, lastRow - 1, 1).setValues(rows.map(row => [row[Q_COL.EVENT_THIS_WEEKEND - 1] === true]));
      }
    }
  }

  const phase11Headers = ['Event Role', 'Routed Support Email'];
  phase11Headers.forEach((header, index) => {
    const column = Q_COL.TALENT_ROLE + index;
    if (String(sheet.getRange(1, column).getValue() || '').trim() !== header) sheet.getRange(1, column).setValue(header);
  });

  const attentionHeaders = [
    'Needs Attention Today', 'Attention Reason', 'Attention Set At',
    'Attention Set By', 'Attention Until', 'Attention Cleared At',
    'Attention Cleared By', 'Attention First Review At', 'Attention First Review By'
  ];
  attentionHeaders.forEach((header, index) => {
    const column = Q_COL.ATTENTION_TODAY + index;
    if (String(sheet.getRange(1, column).getValue() || '').trim() !== header) sheet.getRange(1, column).setValue(header);
  });

  // Coordinator/Sr Coordinator/Project Manager escalation flag. This is
  // intentionally separate from Support's "Attention Today" state so the
  // asker can request a review without changing ownership or ticket status.
  const requesterFlagHeaders = [
    'Requester Flagged', 'Requester Flagged At', 'Requester Flagged By',
    'Requester Flag Reason', 'Requester Flag Acknowledged At', 'Requester Flag Acknowledged By'
  ];
  requesterFlagHeaders.forEach((header, index) => {
    const column = Q_COL.REQUESTER_FLAGGED + index;
    if (String(sheet.getRange(1, column).getValue() || '').trim() !== header) sheet.getRange(1, column).setValue(header);
  });
}

function ensureAnsweredSchema(sheet) {
  let lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headers[A_COL.TICKET_ID - 1] !== 'Ticket ID') {
    sheet.getRange(1, A_COL.TICKET_ID, 1, 2).setValues([['Ticket ID', 'Hold Hours Excluded']]);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const idRange = sheet.getRange(2, A_COL.TICKET_ID, lastRow - 1, 1);
      idRange.setValues(idRange.getValues().map(r => [r[0] ? r[0] : Utilities.getUuid()]));
      const holdRange = sheet.getRange(2, A_COL.HOLD_HOURS, lastRow - 1, 1);
      holdRange.setValues(holdRange.getValues().map(r => [(r[0] === '' || r[0] === null || isNaN(r[0])) ? 0 : r[0]]));
    }
  }

  lastCol = Math.max(sheet.getLastColumn(), 1);
  headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers[A_COL.ASKED_BY_EMAIL - 1] !== 'Asked By Email') {
    sheet.getRange(1, A_COL.ASKED_BY_EMAIL).setValue('Asked By Email');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, A_COL.ASKED_BY_EMAIL, lastRow - 1, 1);
      range.setValues(range.getValues().map(r => [r[0] || '']));
    }
  }

  if (String(sheet.getRange(1, A_COL.FIRST_READ_AT).getValue() || '').trim() !== 'First Read At') {
    sheet.getRange(1, A_COL.FIRST_READ_AT).setValue('First Read At');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, A_COL.FIRST_READ_AT, lastRow - 1, 1).setValues(Array.from({ length: lastRow - 1 }, () => ['']));
    }
  }

  if (String(sheet.getRange(1, A_COL.LAST_UPDATED_AT).getValue() || '').trim() !== 'Last Updated At') {
    sheet.getRange(1, A_COL.LAST_UPDATED_AT).setValue('Last Updated At');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const answeredDates = sheet.getRange(2, A_COL.ANSWERED, lastRow - 1, 1).getValues();
      sheet.getRange(2, A_COL.LAST_UPDATED_AT, lastRow - 1, 1).setValues(answeredDates.map(r => [r[0] || '']));
    }
  }

  const workCategoryWasMissing = String(sheet.getRange(1, A_COL.WORK_CATEGORY).getValue() || '').trim() !== 'Work Category';
  const weekendWasMissing = String(sheet.getRange(1, A_COL.EVENT_THIS_WEEKEND).getValue() || '').trim() !== 'Event This Weekend';
  const eventDateWasMissing = String(sheet.getRange(1, A_COL.EVENT_DATE).getValue() || '').trim() !== 'Event Date';
  if (workCategoryWasMissing || weekendWasMissing || eventDateWasMissing) {
    sheet.getRange(1, A_COL.WORK_CATEGORY, 1, 3).setValues([['Work Category', 'Event This Weekend', 'Event Date']]);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, A_WIDTH).getValues();
      if (workCategoryWasMissing) {
        sheet.getRange(2, A_COL.WORK_CATEGORY, lastRow - 1, 1).setValues(rows.map(row => [
          normalizeWorkCategory(row[A_COL.WORK_CATEGORY - 1] || inferWorkCategory(row[A_COL.QUESTION - 1], row[A_COL.EVENT - 1]))
        ]));
      }
      if (weekendWasMissing) {
        sheet.getRange(2, A_COL.EVENT_THIS_WEEKEND, lastRow - 1, 1).setValues(rows.map(row => [row[A_COL.EVENT_THIS_WEEKEND - 1] === true]));
      }
    }
  }

  const phase11Headers = ['Event Role', 'Routed Support Email', 'Answered By Email'];
  phase11Headers.forEach((header, index) => {
    const column = A_COL.TALENT_ROLE + index;
    if (String(sheet.getRange(1, column).getValue() || '').trim() !== header) sheet.getRange(1, column).setValue(header);
  });

  const attentionHeaders = [
    'Attention Requested', 'Attention Reason', 'Attention Set At',
    'Attention Set By', 'Attention Until', 'Attention Cleared At',
    'Attention Cleared By', 'Attention First Review At', 'Attention First Review By'
  ];
  attentionHeaders.forEach((header, index) => {
    const column = A_COL.ATTENTION_REQUESTED + index;
    if (String(sheet.getRange(1, column).getValue() || '').trim() !== header) sheet.getRange(1, column).setValue(header);
  });
}

function ensureTeamSchema(sheet) {
  if (!sheet) return;
  const expected = ['Name', 'Title', 'Status', 'Category', 'Email', 'PasswordHash', 'PasswordSalt', 'Primary Support Email', 'Backup Support Email'];
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), TEAM_WIDTH)).getValues()[0];
  expected.forEach((header, index) => {
    if (String(headers[index] || '').trim() !== header) sheet.getRange(1, index + 1).setValue(header);
  });
}

// ==========================================
// AUDIT LOG & SOFT-DELETE ARCHIVE
// ==========================================
function logAudit(action, actorEmail, actorName, ticketId, details) {
  try {
    let detailValue = details;
    try {
      const authenticatedEmail = normalizeEmail(Session.getActiveUser().getEmail());
      const effectiveEmail = normalizeEmail(actorEmail);
      if (authenticatedEmail && effectiveEmail && authenticatedEmail !== effectiveEmail) {
        const authenticatedMember = findTeamMemberForWorkspaceIdentity(authenticatedEmail);
        if (authenticatedMember && isAdminMember(authenticatedMember)) {
          let obj;
          if (detailValue && typeof detailValue === 'object' && !Array.isArray(detailValue)) obj = Object.assign({}, detailValue);
          else obj = { originalDetails: String(detailValue || '') };
          obj.adminTestMode = true;
          obj.authenticatedAdminEmail = authenticatedMember.email;
          obj.authenticatedAdminName = authenticatedMember.name;
          obj.actingAsEmail = effectiveEmail;
          obj.actingAsName = String(actorName || '');
          detailValue = obj;
        }
      }
    } catch (identityErr) {}

    const s = ensureAuditSheet(SpreadsheetApp.getActiveSpreadsheet());
    s.appendRow([
      new Date(),
      String(actorEmail || '').trim().toLowerCase(),
      String(actorName || '').trim(),
      String(action || ''),
      String(ticketId || ''),
      typeof detailValue === 'string' ? detailValue : JSON.stringify(detailValue || {})
    ]);
  } catch (e) {
    Logger.log('Audit log failed: ' + e);
  }
}

function archiveDeletedRow(sourceSheetName, rowValues, ticketId, deletedByEmail) {
  const s = ensureDeletedSheet(SpreadsheetApp.getActiveSpreadsheet());
  s.appendRow([
    new Date(),
    String(deletedByEmail || '').trim().toLowerCase(),
    sourceSheetName,
    String(ticketId || ''),
    String(rowValues[0] || ''),
    String(rowValues[1] || ''),
    String(rowValues[2] || ''),
    JSON.stringify(rowValues)
  ]);
}

// ==========================================
// ROLE / AUTH HELPERS
// ==========================================
// Access is derived from CATEGORY ONLY - see the "ROLE MODEL" comment block
// near the top of this file for why Title is never part of this check.
function isAdminCategoryValue(value) {
  return String(value || '').trim().toLowerCase() === 'admin';
}

function isSupportMember(member) {
  if (!member) return false;
  if (String(member.status).toLowerCase() !== 'active') return false;
  const category = String(member.category || '').trim().toLowerCase();
  return category === 'support' || category === 'admin';
}

function isAdminMember(member) {
  if (!member) return false;
  if (String(member.status).toLowerCase() !== 'active') return false;
  return isAdminCategoryValue(member.category);
}

function findTeamMemberByEmail(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return null;
  const team = _getTeamMembersInternal();
  const matches = team.filter(m => m.email === email);
  if (!matches.length) return null;
  if (matches.length === 1) return matches[0];
  // When one person has both an operational Support profile and a separate
  // Admin profile on the same mailbox, normal ticket/task assignment should
  // resolve to the non-Admin profile. Admin authentication is resolved through
  // the explicit profile-selection functions above, not this generic lookup.
  return matches.find(m => !isAdminMember(m)) || matches[0];
}

function requireTeamMember(email) {
  const member = findTeamMemberByEmail(email);
  if (!member || String(member.status || '').trim().toLowerCase() !== 'active') {
    throw new Error("Access denied: unrecognized or inactive user.");
  }
  return member;
}

function requireSupportRole(email) {
  const member = requireTeamMember(email);
  if (!isSupportMember(member)) {
    throw new Error("Access denied: this action requires a Supervisor / Support role.");
  }
  return member;
}

// Question ownership is intentionally narrower than Support Hub permissions.
// Admin can maintain/test the application, but questions must NEVER be assigned
// to the Admin profile. A question owner must be an active Category=Support
// profile. This also keeps shared-email users (Support + Admin) operationally
// attached to their Support identity.
function requireQuestionAssignee(email) {
  const normalized = normalizeEmail(email);
  if (!normalized) throw new Error("Choose a Support user to assign this question to.");
  const matches = _getTeamMembersInternal().filter(m =>
    normalizeEmail(m.email) === normalized &&
    String(m.status || '').trim().toLowerCase() === 'active' &&
    String(m.category || '').trim().toLowerCase() === 'support'
  );
  if (matches.length === 1) return matches[0];
  if (!matches.length) throw new Error("Questions can only be assigned to an active Support profile. Admin profiles cannot own questions.");
  throw new Error("More than one active Support profile uses this email. Please correct Team Setup before assigning the question.");
}

function resolveAutomaticQuestionAssignee(member) {
  if (!member || isAdminMember(member)) return null;
  const candidates = [member.primarySupportEmail, member.backupSupportEmail].map(normalizeEmail).filter(Boolean);
  for (let i = 0; i < candidates.length; i++) {
    try { return requireQuestionAssignee(candidates[i]); } catch (e) {}
  }
  return null;
}

// Phase 11B keeps the working assignee and the Support owner separate. A task
// assigned to Support is owned by that Support profile. A task assigned to a
// Coordinator uses the Coordinator's active Primary Support, then Backup.
function resolveTaskSupportOwnerEmail(assignee) {
  if (!assignee || String(assignee.status || '').trim().toLowerCase() !== 'active') return '';
  const category = String(assignee.category || '').trim().toLowerCase();
  if (category === 'support') {
    try { return requireQuestionAssignee(assignee.email).email; } catch (e) { return ''; }
  }
  if (category === 'coordinator') {
    const owner = resolveAutomaticQuestionAssignee(assignee);
    return owner ? owner.email : '';
  }
  return '';
}

function requireTaskAssignee(email) {
  const normalized = normalizeEmail(email);
  if (!normalized) throw new Error('Choose a team member to assign this task to.');
  const matches = _getTeamMembersInternal().filter(m =>
    normalizeEmail(m.email) === normalized &&
    String(m.status || '').trim().toLowerCase() === 'active' &&
    !isAdminMember(m)
  );
  if (matches.length === 1) return matches[0];
  if (!matches.length) throw new Error('Tasks can only be assigned to an active operational team member. Admin-only profiles cannot own tasks.');
  throw new Error('More than one active operational profile uses this email. Please correct Team Setup before assigning the task.');
}

function requireAdminRole(email) {
  const member = requireTeamMember(email);
  if (!isAdminMember(member)) {
    throw new Error("Access denied: this action requires the Admin role.");
  }
  return member;
}

function getTeamCacheGeneration_() {
  try {
    return String(PropertiesService.getScriptProperties().getProperty(TEAM_CACHE_VERSION_KEY) || '0');
  } catch (e) {
    return '0';
  }
}

function getTeamCacheKey_() {
  return TEAM_CACHE_KEY + ':' + getTeamCacheGeneration_();
}

function invalidateTeamCache() {
  const oldKey = getTeamCacheKey_();
  try {
    const nextGeneration = Utilities.getUuid ? Utilities.getUuid() : String(new Date().getTime());
    PropertiesService.getScriptProperties().setProperty(TEAM_CACHE_VERSION_KEY, nextGeneration);
  } catch (e) {}
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(oldKey);
    // Remove the legacy unversioned keys as well. This is safe and makes the
    // first deployment of this fix deterministic even during a warm runtime.
    cache.remove(TEAM_CACHE_KEY);
    cache.remove('SUPPORT_HUB_TEAM_ROSTER_V1');
  } catch (e) {}
}

function cachedTeamHasRoutingShape_(team) {
  return Array.isArray(team) && team.every(member => member &&
    Object.prototype.hasOwnProperty.call(member, 'primarySupportEmail') &&
    Object.prototype.hasOwnProperty.call(member, 'backupSupportEmail'));
}

function _getTeamMembersInternal(forceFresh) {
  ensureSheetsExist();

  const cache = CacheService.getScriptCache();
  const cacheKey = getTeamCacheKey_();
  const cached = forceFresh ? null : cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (cachedTeamHasRoutingShape_(parsed)) return parsed;
    } catch (e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let team = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).trim() !== "") {
      team.push({
        name: stripNoraPrefix(String(data[i][0]).trim()),
        title: String(data[i][1]).trim(),
        status: String(data[i][2]).trim(),
        category: String(data[i][3]).trim(),
        email: String(data[i][4]).trim().toLowerCase(),
        primarySupportEmail: normalizeEmail(data[i][TEAM_COL.PRIMARY_SUPPORT_EMAIL - 1]),
        backupSupportEmail: normalizeEmail(data[i][TEAM_COL.BACKUP_SUPPORT_EMAIL - 1])
      });
    }
  }
  try { cache.put(cacheKey, JSON.stringify(team), TEAM_CACHE_SECONDS); } catch (e) {}
  return team;
}


function getTeamMembers() {
  requireAuthenticatedMember();
  return _getTeamMembersInternal();
}

function stripNoraPrefix(name) {
  const raw = String(name || '').trim();
  const cleaned = raw.replace(/^Nora[\s\-_:/]*/i, '').trim();
  return cleaned || raw;
}

// ---- Password hashing (Admin accounts only) ----
function hashPassword(password, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + '::' + String(salt));
  return Utilities.base64Encode(digest);
}

function countWorkingAdmins() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const category = String(data[i][3] || '').trim();
    const status = String(data[i][2] || '').trim().toLowerCase();
    const passwordHash = String(data[i][5] || '').trim();
    if (isAdminCategoryValue(category) && status === 'active' && passwordHash) count++;
  }
  return count;
}

function getTeamMemberRawByEmail(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  let adminMatch = null;
  let nonAdminMatch = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim().toLowerCase() !== email) continue;
    const rec = {
      rowIndex: i + 1,
      name: String(data[i][0]).trim(),
      title: String(data[i][1]).trim(),
      status: String(data[i][2]).trim(),
      category: String(data[i][3]).trim(),
      email: email,
      passwordHash: String(data[i][5] || '').trim(),
      passwordSalt: String(data[i][6] || '').trim()
    };
    if (isAdminCategoryValue(rec.category)) adminMatch = rec;
    else if (!nonAdminMatch) nonAdminMatch = rec;
  }
  // Operational lookups should resolve the person's normal profile, not the
  // generic Admin profile, when both intentionally share the same mailbox.
  return nonAdminMatch || adminMatch;
}

function getTeamMemberRawByProfileKey(profileKey) {
  const wanted = String(profileKey || '').trim();
  if (!wanted) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEAM);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const member = {
      name: stripNoraPrefix(String(data[i][0] || '').trim()),
      title: String(data[i][1] || '').trim(),
      status: String(data[i][2] || '').trim(),
      category: String(data[i][3] || '').trim(),
      email: normalizeEmail(data[i][4])
    };
    if (memberProfileKey(member) !== wanted) continue;
    return {
      rowIndex: i + 1,
      name: String(data[i][0] || '').trim(),
      title: String(data[i][1] || '').trim(),
      status: String(data[i][2] || '').trim(),
      category: String(data[i][3] || '').trim(),
      email: normalizeEmail(data[i][4]),
      passwordHash: String(data[i][5] || '').trim(),
      passwordSalt: String(data[i][6] || '').trim()
    };
  }
  return null;
}

function verifyAdminPassword(email, passwordAttempt) {
  const authenticatedEmail = getAuthenticatedEmail();
  const workspaceMember = requireAuthenticatedWorkspaceMember();
  if (normalizeEmail(email) !== normalizeEmail(workspaceMember.email)) {
    throw new Error('Security check failed: Admin account does not match the signed-in Google Workspace identity.');
  }
  const member = getTeamMemberRawByProfileKey(workspaceMember.profileKey) || getTeamMemberRawByEmail(workspaceMember.email);
  if (!member || String(member.status).toLowerCase() !== 'active') {
    throw new Error("Access denied: unrecognized or inactive user.");
  }
  if (!isAdminCategoryValue(member.category)) {
    throw new Error("This account is not an Admin account.");
  }
  if (!member.passwordHash) {
    throw new Error("This Admin account has no password set yet. Ask another Admin to set one in Team Setup.");
  }
  const attemptHash = hashPassword(passwordAttempt, member.passwordSalt);
  if (attemptHash !== member.passwordHash) {
    logAudit('ADMIN_LOGIN_FAILED', email, member.name, '', 'Incorrect password attempt');
    logAudit('AUTH_FAILED', email, member.name, '', { effectiveEmail: email, reason: 'Incorrect Admin password' });
    throw new Error("Incorrect password.");
  }
  logAudit('ADMIN_LOGIN', email, member.name, '', 'Admin password verified');
  return { success: true };
}

// ==========================================
// TITLES CONFIG (Admin-manageable, display-only, never affects permissions)
// ==========================================
function getCustomTitles() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(CONFIG_KEY_CUSTOM_TITLES);
    const list = raw ? JSON.parse(raw) : [];
    return Array.isArray(list) ? list.filter(t => typeof t === 'string' && t.trim()) : [];
  } catch (e) {
    return [];
  }
}

// Base titles + admin-added custom titles, in that order. This is the full
// list of options the Team Setup "Title" dropdown should offer.
function getAllTitles() {
  return BASE_TITLES.concat(getCustomTitles());
}

// Admin-only. Adds a new title option. This ONLY changes what's offered in
// the Title dropdown - it never grants any access by itself (see ROLE MODEL
// comment above). Rejects duplicates (case-insensitive) against the
// combined base + custom list.
function addCustomTitle(title, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const clean = String(title || '').trim();
    if (!clean) throw new Error("Enter a title name.");
    if (clean.length > 60) throw new Error("Title is too long - keep it under 60 characters.");

    const existing = getAllTitles();
    if (existing.some(t => t.toLowerCase() === clean.toLowerCase())) {
      throw new Error("That title already exists.");
    }

    const custom = getCustomTitles();
    custom.push(clean);
    PropertiesService.getScriptProperties().setProperty(CONFIG_KEY_CUSTOM_TITLES, JSON.stringify(custom));
    logAudit('ADD_TITLE', admin.email, admin.name, '', { title: clean });
    return { success: true, titles: getAllTitles() };
  }, { bumpDataVersion: false });
}

// Admin-only. Removes a previously-added custom title. Base titles can never
// be removed. Blocked if any roster row (active or inactive) still uses this
// title, so nobody's roster entry is left pointing at a title that no longer
// exists in the dropdown.
function deleteCustomTitle(title, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const clean = String(title || '').trim();
    const custom = getCustomTitles();

    if (BASE_TITLES.some(t => t.toLowerCase() === clean.toLowerCase())) {
      throw new Error("Base titles can't be removed.");
    }
    if (!custom.some(t => t.toLowerCase() === clean.toLowerCase())) {
      throw new Error("That title doesn't exist.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TEAM);
    const data = sheet.getDataRange().getValues();
    const inUse = data.slice(1).some(row => String(row[1] || '').trim().toLowerCase() === clean.toLowerCase());
    if (inUse) {
      throw new Error("This title is still assigned to one or more team members. Reassign them to a different title first.");
    }

    const updated = custom.filter(t => t.toLowerCase() !== clean.toLowerCase());
    PropertiesService.getScriptProperties().setProperty(CONFIG_KEY_CUSTOM_TITLES, JSON.stringify(updated));
    logAudit('DELETE_TITLE', admin.email, admin.name, '', { title: clean });
    return { success: true, titles: getAllTitles() };
  }, { bumpDataVersion: false });
}

// ==========================================
// DATE / MISC HELPERS
// ==========================================
function safeIsoDate(val) {
  if (!val) return new Date().toISOString();
  if (val instanceof Date && !isNaN(val.getTime())) {
    return val.toISOString();
  }
  try {
    const parsed = new Date(val);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }
  } catch (e) {}
  return new Date().toISOString();
}

function formatCSTStamp(date) {
  return Utilities.formatDate(date || new Date(), APP_TIMEZONE, "MMM d, h:mm a");
}

// ==========================================
// VALIDATION HELPERS (server-side, always enforced)
// ==========================================
const EVENT_NAME_INVALID_MSG = "That doesn't look like a name - please re-enter the Event/Client/Talent name as text, or put NA if no event or contact is available/connected.";
function isValidEventName(text) {
  var value = String(text || '').trim();
  if (!value) return false;
  var isUrl = /^(https?:\/\/|www\.)/i.test(value) || /:\/\//.test(value) || /^(ftp|file):\/\//i.test(value);
  if (isUrl) return false;
  return true;
}

function isValidCaseLink(text) {
  const value = String(text || '').trim();
  if (!value) return true;
  // Support Hub historically stores either a full Salesforce URL or the raw
  // 15/18-character Salesforce record id. Admin edits must preserve both.
  return /^https?:\/\/\S+$/i.test(value) || /^[A-Za-z0-9]{15}(?:[A-Za-z0-9]{3})?$/.test(value);
}

function normalizeEmail(email) {
  return String(email || '').trim().toLowerCase();
}

// ==========================================
// APP CONFIG: Salesforce base URL + Titles
// ==========================================
const CONFIG_KEY_SF_BASE = 'SALESFORCE_BASE_URL';

function getSalesforceBaseUrl() {
  return PropertiesService.getScriptProperties().getProperty(CONFIG_KEY_SF_BASE) || '';
}

function setSalesforceBaseUrl(url, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const clean = String(url || '').trim().replace(/\/+$/, '');
    if (clean && !/^https:\/\/[^\/]+\.(force\.com|salesforce\.com)$/i.test(clean)) {
      throw new Error("Enter just the base URL of your Salesforce instance, e.g. https://yourcompany.lightning.force.com (no trailing path).");
    }
    PropertiesService.getScriptProperties().setProperty(CONFIG_KEY_SF_BASE, clean);
    logAudit('CONFIG_SET_SF_BASE', admin.email, admin.name, '', { url: clean });
    return { success: true, salesforceBaseUrl: clean };
  }, { bumpDataVersion: false });
}

function autoDetectSalesforceBaseUrl(fullUrl, requestingEmail) {
  return withLock(() => {
    requireAuthenticatedMember(requestingEmail);
    if (getSalesforceBaseUrl()) return { success: true, changed: false };

    const hostMatch = String(fullUrl || '').match(/^https?:\/\/([^\/]+)/i);
    if (!hostMatch) return { success: true, changed: false };
    const host = hostMatch[1].toLowerCase();
    if (!/\.(force\.com|salesforce\.com)$/i.test(host)) return { success: true, changed: false };

    const base = 'https://' + host;
    PropertiesService.getScriptProperties().setProperty(CONFIG_KEY_SF_BASE, base);
    logAudit('CONFIG_AUTO_DETECT_SF_BASE', requestingEmail, '', '', { url: base });
    return { success: true, changed: true, base: base };
  }, { bumpDataVersion: false });
}

// Single call the client makes once at boot to pick up server-side config,
// including the current full list of Title options.
function getAppConfig() {
  return {
    salesforceBaseUrl: getSalesforceBaseUrl(),
    titles: getAllTitles(),
    workCategories: getWorkCategoryConfigs(false).map(row => row.category)
  };
}

function normalizeWorkCategoryKeywords(value) {
  const raw = Array.isArray(value) ? value : String(value || '').split(/[\n,]+/);
  const seen = {};
  const out = [];
  raw.forEach(item => {
    const clean = String(item || '').trim().replace(/\s+/g, ' ');
    const key = clean.toLowerCase();
    if (!clean || seen[key]) return;
    if (clean.length > 80) throw new Error('Each keyword or phrase must be 80 characters or fewer.');
    seen[key] = true;
    out.push(clean);
  });
  if (out.length > 50) throw new Error('Use no more than 50 keywords or phrases for one category.');
  return out;
}

function getWorkCategoryAdminData(requestingEmail) {
  const admin = requireAuthenticatedAdmin(requestingEmail);
  ensureSheetsExist();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const suggestionsSheet = ensureCategorySuggestionsSheet(ss);
  let suggestions = [];
  if (suggestionsSheet.getLastRow() > 1) {
    suggestions = suggestionsSheet.getRange(2, 1, suggestionsSheet.getLastRow() - 1, 11).getValues().map(row => ({
      id: String(row[0] || ''), suggestion: String(row[1] || ''), suggestedBy: stripNoraPrefix(String(row[2] || '')),
      suggestedByEmail: normalizeEmail(row[3]), context: String(row[4] || ''), recordId: String(row[5] || ''),
      createdAt: row[6] ? safeIsoDate(row[6]) : '', status: String(row[7] || 'Pending'),
      reviewedBy: stripNoraPrefix(String(row[8] || '')), reviewedAt: row[9] ? safeIsoDate(row[9]) : '', adminNote: String(row[10] || '')
    })).filter(row => row.id).sort((a, b) => {
      if (a.status === 'Pending' && b.status !== 'Pending') return -1;
      if (a.status !== 'Pending' && b.status === 'Pending') return 1;
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    }).slice(0, 100);
  }
  return { success: true, requestedBy: admin.email, categories: getWorkCategoryConfigs(true), suggestions: suggestions };
}

function saveWorkCategoryConfig(payload, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    payload = payload || {};
    const name = String(payload.category || '').trim().replace(/\s+/g, ' ');
    const original = String(payload.originalCategory || '').trim();
    const active = payload.active !== false;
    const keywords = normalizeWorkCategoryKeywords(payload.keywords);
    if (!name) throw new Error('Enter a category name.');
    if (name.length > 60) throw new Error('Category name must be 60 characters or fewer.');
    if (/[<>]/.test(name)) throw new Error('Category name contains unsupported characters.');
    if (original && original.toLowerCase() !== name.toLowerCase()) throw new Error('Category names cannot be renamed because historical tickets may use them. Create a new category instead.');
    if (name.toLowerCase() !== 'general / other' && !keywords.length) throw new Error('Add at least one keyword or phrase for automatic categorization.');

    ensureSheetsExist();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureWorkCategoriesSheet(ss);
    const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 6).getValues();
    let targetIndex = -1;
    data.forEach((row, index) => {
      if (String(row[0] || '').trim().toLowerCase() === (original || name).toLowerCase()) targetIndex = index;
    });
    if (!original && targetIndex !== -1) throw new Error('That category already exists.');
    if (original && targetIndex === -1) throw new Error('That category no longer exists. Refresh and try again.');

    const otherConfigs = data.filter((row, index) => index !== targetIndex && (row[2] === true || String(row[2] || '').toLowerCase() === 'true'));
    keywords.forEach(keyword => {
      const duplicate = otherConfigs.find(row => String(row[1] || '').split(',').map(k => k.trim().toLowerCase()).includes(keyword.toLowerCase()));
      if (duplicate) throw new Error('The keyword "' + keyword + '" is already used by ' + String(duplicate[0] || 'another category') + '.');
    });

    const now = new Date();
    const isGeneral = name.toLowerCase() === 'general / other';
    if (targetIndex === -1) {
      sheet.appendRow([name, keywords.join(', '), true, 5, now, admin.name]);
    } else {
      const existingOrder = Number(data[targetIndex][3]) || (isGeneral ? 999 : 500 + targetIndex);
      sheet.getRange(targetIndex + 2, 1, 1, 6).setValues([[name, keywords.join(', '), isGeneral ? true : active, existingOrder, now, admin.name]]);
    }
    clearWorkCategoryCache();

    const suggestionId = String(payload.suggestionId || '').trim();
    if (suggestionId) updateCategorySuggestionStatus_(suggestionId, 'Approved', admin, 'Created category: ' + name);
    logAudit(targetIndex === -1 ? 'WORK_CATEGORY_ADD' : 'WORK_CATEGORY_EDIT', admin.email, admin.name, '', { category: name, active: isGeneral ? true : active, keywords: keywords });
    return getWorkCategoryAdminData(admin.email);
  }, { bumpDataVersion: false });
}

function submitWorkCategorySuggestion(suggestion, context, recordId, requestingEmail) {
  return withLock(() => {
    const member = requireAuthenticatedMember(requestingEmail);
    const clean = String(suggestion || '').trim().replace(/\s+/g, ' ');
    if (clean.length < 3) throw new Error('Enter a category suggestion with at least 3 characters.');
    if (clean.length > 60) throw new Error('Category suggestion must be 60 characters or fewer.');
    if (/[<>]/.test(clean)) throw new Error('Category suggestion contains unsupported characters.');
    ensureSheetsExist();
    if (getWorkCategoryConfigs(true).some(row => row.category.toLowerCase() === clean.toLowerCase())) {
      throw new Error('That category already exists. Select it from the category list.');
    }
    const sheet = ensureCategorySuggestionsSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (sheet.getLastRow() > 1) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
      const existing = rows.find(row => String(row[1] || '').trim().toLowerCase() === clean.toLowerCase() && String(row[7] || 'Pending') === 'Pending');
      if (existing) return { success: true, duplicate: true, message: 'This category has already been suggested and is waiting for Admin review.', _skipDataVersion: true };
    }
    const id = Utilities.getUuid();
    sheet.appendRow([id, clean, member.name, member.email, String(context || '').slice(0, 120), String(recordId || '').slice(0, 100), new Date(), 'Pending', '', '', '']);
    logAudit('WORK_CATEGORY_SUGGEST', member.email, member.name, String(recordId || ''), { suggestion: clean, context: String(context || '') });
    return { success: true, id: id, message: 'Category suggestion sent to Admin.' };
  }, { bumpDataVersion: false });
}

function updateCategorySuggestionStatus_(suggestionId, status, admin, note) {
  const sheet = ensureCategorySuggestionsSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (sheet.getLastRow() <= 1) throw new Error('Category suggestion not found.');
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const index = ids.findIndex(row => String(row[0] || '') === String(suggestionId || ''));
  if (index === -1) throw new Error('Category suggestion not found.');
  const rowNumber = index + 2;
  sheet.getRange(rowNumber, 8, 1, 4).setValues([[status, admin.name, new Date(), String(note || '')]]);
}

function dismissWorkCategorySuggestion(suggestionId, note, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    updateCategorySuggestionStatus_(suggestionId, 'Declined', admin, String(note || 'Not added'));
    logAudit('WORK_CATEGORY_SUGGESTION_DECLINE', admin.email, admin.name, '', { suggestionId: suggestionId, note: String(note || '') });
    return getWorkCategoryAdminData(admin.email);
  }, { bumpDataVersion: false });
}

// ==========================================
// URL OPTIMIZATION: Extract Case ID from Salesforce URLs
// ==========================================
function extractCaseId(fullUrl) {
  const url = String(fullUrl || '').trim();
  if (!url) return '';
  if (!url.includes('://')) return url;

  const hostMatch = url.match(/^https?:\/\/([^\/]+)/i);
  const host = hostMatch ? hostMatch[1].toLowerCase() : '';
  const isSalesforceHost = /\.force\.com$/.test(host) || /\.salesforce\.com$/.test(host);
  if (!isSalesforceHost) return url;

  let match = url.match(/\/[Cc]ase\/([a-zA-Z0-9]{15,18})(?:[\/?]|$)/);
  if (match) return match[1];

  match = url.match(/\/([5][a-zA-Z0-9]{14,17})(?:[\/?]|$)/);
  if (match) return match[1];

  return url;
}

function reconstructCaseUrl(caseId, baseUrl) {
  if (!caseId) return '';
  if (caseId.includes('://')) return caseId;

  const base = (baseUrl || getSalesforceBaseUrl() || '').replace(/\/+$/, '');
  if (!base) return caseId;

  return base + '/lightning/r/Case/' + caseId + '/view';
}

function stableCacheHash(value) {
  const text = String(value || '');
  let h1 = 0x811c9dc5;
  let h2 = 0x9e3779b9;
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 16777619);
    h2 = Math.imul(h2 ^ c, 2246822519);
  }
  return (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36);
}

function safeScriptCacheGet(key) {
  try { return CacheService.getScriptCache().get(key); }
  catch (e) { return null; }
}

function safeScriptCachePut(key, value, seconds) {
  try { CacheService.getScriptCache().put(key, String(value), Number(seconds) || 300); }
  catch (e) { Logger.log('Non-fatal cache write failed: ' + e); }
}

function safeScriptCacheRemove(key) {
  try { CacheService.getScriptCache().remove(key); }
  catch (e) {}
}

// ==========================================
// TICKET-ID ROW RESOLUTION
// ==========================================
function findRowIndexByTicketId(sheet, ticketId, ticketIdColumn) {
  const id = String(ticketId || '').trim();
  if (!id) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const sheetName = sheet.getName ? String(sheet.getName() || '') : '';
  const cacheKey = 'ROW_INDEX_V1_' + stableCacheHash(sheetName + '|' + String(ticketIdColumn) + '|' + id);
  const cachedRow = Number(safeScriptCacheGet(cacheKey));
  if (cachedRow >= 2 && cachedRow <= lastRow) {
    try {
      const cachedId = String(sheet.getRange(cachedRow, ticketIdColumn).getValue() || '').trim();
      if (cachedId === id) return cachedRow;
    } catch (e) {}
    safeScriptCacheRemove(cacheKey);
  }
  const ids = sheet.getRange(2, ticketIdColumn, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) {
      const rowIndex = i + 2;
      safeScriptCachePut(cacheKey, rowIndex, ROW_INDEX_CACHE_SECONDS);
      return rowIndex;
    }
  }
  return -1;
}

function requireQuestionRow(ticketId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_QUESTIONS);
  const rowIndex = findRowIndexByTicketId(sheet, ticketId, Q_COL.TICKET_ID);
  if (rowIndex === -1) {
    throw new Error("This ticket is no longer open - it may have just been answered or removed by someone else. Refresh to see its current state.");
  }
  return { sheet: sheet, rowIndex: rowIndex };
}

function requireAnsweredRow(ticketId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ANSWERED);
  const rowIndex = findRowIndexByTicketId(sheet, ticketId, A_COL.TICKET_ID);
  if (rowIndex === -1) {
    throw new Error("This ticket record no longer exists - it may have just been edited or deleted by someone else. Refresh to see its current state.");
  }
  return { sheet: sheet, rowIndex: rowIndex };
}

// ==========================================
// HOLD-AWARE ELAPSED-TIME MATH
// ==========================================
function calculateWorkingHours(created, holdSinceVal, cumulativeHoldHours) {
  const now = new Date();
  let totalElapsed = Math.max(0, (now - created) / (1000 * 60 * 60));
  let holdHours = Number(cumulativeHoldHours) || 0;

  if (holdSinceVal) {
    const holdSince = (holdSinceVal instanceof Date) ? holdSinceVal : new Date(holdSinceVal);
    if (!isNaN(holdSince.getTime())) {
      holdHours += Math.max(0, (now - holdSince) / (1000 * 60 * 60));
    }
  }
  return Math.max(0, totalElapsed - holdHours);
}

function calculateElapsedHours(createdVal, holdSinceVal, cumulativeHoldHours) {
  if (!createdVal) return "0.0";
  const created = (createdVal instanceof Date) ? createdVal : new Date(createdVal);
  if (isNaN(created.getTime())) return "0.0";
  return calculateWorkingHours(created, holdSinceVal, cumulativeHoldHours).toFixed(1);
}

// ---- Per-user "Read By" helpers ----
function parseReadByList(raw) {
  if (raw === true) return ['*'];
  if (!raw) return [];
  return String(raw).split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}
function isReadByUser(raw, email) {
  const list = parseReadByList(raw);
  if (list.indexOf('*') !== -1) return true;
  return list.indexOf(String(email || '').trim().toLowerCase()) !== -1;
}

function parseFirstReadMap(raw) {
  if (!raw) return {};
  if (typeof raw === 'object' && !(raw instanceof Date)) return raw;
  try {
    const parsed = JSON.parse(String(raw));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch (e) {
    return {};
  }
}

function serializeFirstReadMap(map) {
  const clean = {};
  Object.keys(map || {}).forEach(key => {
    const email = normalizeEmail(key);
    const value = String(map[key] || '').trim();
    if (email && value) clean[email] = value;
  });
  return Object.keys(clean).length ? JSON.stringify(clean) : '';
}

function getFirstReadAtForUser(raw, email) {
  const map = parseFirstReadMap(raw);
  return map[normalizeEmail(email)] || '';
}

// ==========================================
// READ: main data feed
// ==========================================
function getQuestionsData(requestingEmail) {
  const requestingMember = requireAuthenticatedMember(requestingEmail);
  return getQuestionsDataForMember(requestingEmail, requestingMember);
}

function getQuestionsDataForMember(requestingEmail, requestingMember) {
  const requesterIsSupport = isSupportMember(requestingMember);
  const requesterEmail = normalizeEmail(requestingEmail);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);

  let records = [];

  try {
    if (qSheet && qSheet.getLastRow() > 1) {
      // Read only the application schema columns and data rows. getDataRange()
      // can become much larger when a Sheet has formatting or stray content
      // beyond the schema, which slows every browser refresh.
      const qData = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, Q_WIDTH).getValues();
      for (let i = 0; i < qData.length; i++) {
        const row = qData[i];
        const qText = String(row[0] || "").trim();
        if (!qText) continue;

        const ticketId = String(row[Q_COL.TICKET_ID - 1] || "").trim();
        const holdSinceRaw = row[Q_COL.HOLD_SINCE - 1];
        const cumulativeHold = Number(row[Q_COL.HOLD_ACCUM - 1]) || 0;
        const createdIso = safeIsoDate(row[3]);
        const hoursElapsed = calculateElapsedHours(row[3], holdSinceRaw, cumulativeHold);

        const compactQuestion = compactRichContentForFeed(qText);
        records.push({
          id: "Q_" + (ticketId || (i + 2)),
          ticketId: ticketId,
          sheet: SHEET_QUESTIONS,
          question: compactQuestion.html,
          eventName: String(row[1] || "").trim(),
          askedBy: stripNoraPrefix(String(row[2] || "").trim()),
          askedByEmail: String(row[Q_COL.ASKED_BY_EMAIL - 1] || "").trim().toLowerCase(),
          created: createdIso,
          caseLink: extractCaseId(String(row[7] || "").trim()),
          answer: String(row[Q_COL.ANSWER - 1] || '').trim(),
          contentDeferred: compactQuestion.deferred,
          status: String(row[9] || STATUS_OPEN).trim(),
          assignedTo: String(row[10] || "").trim().toLowerCase(),
          hoursElapsed: hoursElapsed,
          holdReason: String(row[Q_COL.HOLD_REASON - 1] || "").trim(),
          holdSince: holdSinceRaw ? safeIsoDate(holdSinceRaw) : "",
          isFollowUp: isFollowUpValue(row[Q_COL.FOLLOWUP - 1]),
          assignNotes: String(row[Q_COL.ASSIGN_NOTES - 1] || "").trim(),
          statusChangedAt: row[Q_COL.STATUS_CHANGED_AT - 1] ? safeIsoDate(row[Q_COL.STATUS_CHANGED_AT - 1]) : "",
          statusChangedBy: stripNoraPrefix(String(row[Q_COL.STATUS_CHANGED_BY - 1] || '').trim()),
          category: normalizeWorkCategory(row[Q_COL.WORK_CATEGORY - 1] || inferWorkCategory(qText, row[Q_COL.EVENT - 1])),
          eventDate: normalizeEventDate(row[Q_COL.EVENT_DATE - 1]),
          talentRole: String(row[Q_COL.TALENT_ROLE - 1] || '').trim(),
          routedSupportEmail: normalizeEmail(row[Q_COL.ROUTED_SUPPORT_EMAIL - 1]),
          eventThisWeekend: isActiveEventThisWeekend(row[Q_COL.EVENT_THIS_WEEKEND - 1] === true, row[Q_COL.EVENT_DATE - 1]),
          eventThisWeekendFlag: row[Q_COL.EVENT_THIS_WEEKEND - 1] === true,
          attentionRequested: attentionWasRequested(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_SET_AT - 1]),
          attentionToday: isAttentionActive(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_UNTIL - 1]),
          attentionReason: String(row[Q_COL.ATTENTION_REASON - 1] || '').trim(),
          attentionSetAt: row[Q_COL.ATTENTION_SET_AT - 1] ? safeIsoDate(row[Q_COL.ATTENTION_SET_AT - 1]) : '',
          attentionSetBy: stripNoraPrefix(String(row[Q_COL.ATTENTION_SET_BY - 1] || '').trim()),
          attentionUntil: normalizeEventDate(row[Q_COL.ATTENTION_UNTIL - 1]),
          attentionClearedAt: row[Q_COL.ATTENTION_CLEARED_AT - 1] ? safeIsoDate(row[Q_COL.ATTENTION_CLEARED_AT - 1]) : '',
          attentionClearedBy: stripNoraPrefix(String(row[Q_COL.ATTENTION_CLEARED_BY - 1] || '').trim()),
          attentionFirstReviewAt: row[Q_COL.ATTENTION_FIRST_REVIEW_AT - 1] ? safeIsoDate(row[Q_COL.ATTENTION_FIRST_REVIEW_AT - 1]) : '',
          attentionFirstReviewBy: stripNoraPrefix(String(row[Q_COL.ATTENTION_FIRST_REVIEW_BY - 1] || '').trim()),
          requesterFlagged: row[Q_COL.REQUESTER_FLAGGED - 1] === true || !!row[Q_COL.REQUESTER_FLAGGED_AT - 1],
          requesterFlaggedAt: row[Q_COL.REQUESTER_FLAGGED_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAGGED_AT - 1]) : '',
          requesterFlaggedBy: stripNoraPrefix(String(row[Q_COL.REQUESTER_FLAGGED_BY - 1] || '').trim()),
          requesterFlagReason: String(row[Q_COL.REQUESTER_FLAG_REASON - 1] || '').trim(),
          requesterFlagAcknowledgedAt: row[Q_COL.REQUESTER_FLAG_ACK_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAG_ACK_AT - 1]) : '',
          requesterFlagAcknowledgedBy: stripNoraPrefix(String(row[Q_COL.REQUESTER_FLAG_ACK_BY - 1] || '').trim()),
          requesterFlagActive: (row[Q_COL.REQUESTER_FLAGGED - 1] === true || !!row[Q_COL.REQUESTER_FLAGGED_AT - 1]) && !row[Q_COL.REQUESTER_FLAG_ACK_AT - 1],
          isRead: true
        });
      }
    }

    if (aSheet && aSheet.getLastRow() > 1) {
      const aData = aSheet.getRange(2, 1, aSheet.getLastRow() - 1, A_WIDTH).getValues();
      for (let i = 0; i < aData.length; i++) {
        const row = aData[i];
        const qText = String(row[0] || "").trim();
        if (!qText) continue;

        const ticketId = String(row[A_COL.TICKET_ID - 1] || "").trim();
        const createdIso = safeIsoDate(row[3]);
        const answeredIso = safeIsoDate(row[4]);

        let turnaround = row[6];
        if (turnaround === "" || turnaround === undefined || isNaN(turnaround) || Number(turnaround) < 0) {
          const cDate = new Date(createdIso);
          const aDate = new Date(answeredIso);
          const diff = (aDate - cDate) / (1000 * 60 * 60);
          turnaround = diff > 0 ? diff.toFixed(1) : "0.5";
        } else {
          turnaround = Math.abs(Number(turnaround)).toFixed(1);
        }

        const compactQuestion = compactRichContentForFeed(qText);
        const compactAnswer = compactRichContentForFeed(row[A_COL.ANSWER - 1]);
        const readByList = parseReadByList(row[A_COL.READ_BY - 1]);
        const firstReadMap = parseFirstReadMap(row[A_COL.FIRST_READ_AT - 1]);
        const requesterHasRead = readByList.indexOf('*') !== -1 || readByList.indexOf(requesterEmail) !== -1;
        const lastUpdatedIso = safeIsoDate(row[A_COL.LAST_UPDATED_AT - 1] || row[A_COL.ANSWERED - 1]);
        records.push({
          id: "A_" + (ticketId || (i + 2)),
          ticketId: ticketId,
          sheet: SHEET_ANSWERED,
          question: compactQuestion.html,
          eventName: String(row[1] || "").trim(),
          askedBy: stripNoraPrefix(String(row[2] || "").trim()),
          askedByEmail: String(row[A_COL.ASKED_BY_EMAIL - 1] || "").trim().toLowerCase(),
          created: createdIso,
          answeredDate: answeredIso,
          turnaroundHours: turnaround,
          holdHoursExcluded: Number(row[A_COL.HOLD_HOURS - 1]) || 0,
          caseLink: extractCaseId(String(row[7] || "").trim()),
          answer: compactAnswer.html,
          contentDeferred: compactQuestion.deferred || compactAnswer.deferred,
          status: String(row[9] || STATUS_ANSWERED).trim(),
          answeredBy: stripNoraPrefix(String(row[10] || "Supervisor").trim()),
          hoursElapsed: 0,
          isRead: requesterHasRead,
          firstReadAt: firstReadMap[requesterEmail] || '',
          lastUpdatedAt: lastUpdatedIso,
          hasUnreadUpdate: !requesterHasRead && !!row[A_COL.LAST_UPDATED_AT - 1] &&
            new Date(lastUpdatedIso).getTime() > new Date(answeredIso).getTime(),
          category: normalizeWorkCategory(row[A_COL.WORK_CATEGORY - 1] || inferWorkCategory(qText, row[A_COL.EVENT - 1])),
          eventDate: normalizeEventDate(row[A_COL.EVENT_DATE - 1]),
          talentRole: String(row[A_COL.TALENT_ROLE - 1] || '').trim(),
          routedSupportEmail: normalizeEmail(row[A_COL.ROUTED_SUPPORT_EMAIL - 1]),
          answeredByEmail: normalizeEmail(row[A_COL.ANSWERED_BY_EMAIL - 1]),
          eventThisWeekend: isActiveEventThisWeekend(row[A_COL.EVENT_THIS_WEEKEND - 1] === true, row[A_COL.EVENT_DATE - 1]),
          eventThisWeekendFlag: row[A_COL.EVENT_THIS_WEEKEND - 1] === true,
          attentionRequested: row[A_COL.ATTENTION_REQUESTED - 1] === true || !!row[A_COL.ATTENTION_SET_AT - 1],
          attentionToday: false,
          attentionReason: String(row[A_COL.ATTENTION_REASON - 1] || '').trim(),
          attentionSetAt: row[A_COL.ATTENTION_SET_AT - 1] ? safeIsoDate(row[A_COL.ATTENTION_SET_AT - 1]) : '',
          attentionSetBy: stripNoraPrefix(String(row[A_COL.ATTENTION_SET_BY - 1] || '').trim()),
          attentionUntil: normalizeEventDate(row[A_COL.ATTENTION_UNTIL - 1]),
          attentionClearedAt: row[A_COL.ATTENTION_CLEARED_AT - 1] ? safeIsoDate(row[A_COL.ATTENTION_CLEARED_AT - 1]) : '',
          attentionClearedBy: stripNoraPrefix(String(row[A_COL.ATTENTION_CLEARED_BY - 1] || '').trim()),
          attentionFirstReviewAt: row[A_COL.ATTENTION_FIRST_REVIEW_AT - 1] ? safeIsoDate(row[A_COL.ATTENTION_FIRST_REVIEW_AT - 1]) : '',
          attentionFirstReviewBy: stripNoraPrefix(String(row[A_COL.ATTENTION_FIRST_REVIEW_BY - 1] || '').trim()),
          // Support/Admin reporting needs the full review state to calculate each
          // asker's progress. Coordinators only receive their own derived state.
          readByEmails: requesterIsSupport ? readByList.filter(e => e !== '*') : [],
          readByAll: requesterIsSupport ? readByList.indexOf('*') !== -1 : false,
          firstReadBy: requesterIsSupport ? firstReadMap : {}
        });
      }
    }
  } catch (err) {
    Logger.log("Error in getQuestionsData: " + err.toString());
    // Never stamp an incomplete history as the current successful version.
    throw new Error('Ticket history could not be loaded completely. Please retry.');
  }

  return records;
}

// ==========================================
// OPTIMIZED DATA LOADING WITH CHANGE DETECTION
// ==========================================
const CONFIG_KEY_DATA_VERSION = 'DATA_VERSION_COUNTER';
const CONFIG_KEY_LEGACY_ROUTING_SCAN_VERSION = 'LEGACY_ROUTING_SCAN_DATA_VERSION_V1';
const LEGACY_ROUTING_LOOKBACK_DAYS = 14;

function bumpDataVersion() {
  const props = PropertiesService.getScriptProperties();
  const current = Number(props.getProperty(CONFIG_KEY_DATA_VERSION)) || 0;
  const next = String(current + 1);
  props.setProperty(CONFIG_KEY_DATA_VERSION, next);
  return next;
}

function getDataVersion() {
  return PropertiesService.getScriptProperties().getProperty(CONFIG_KEY_DATA_VERSION) || '0';
}

// Older active deployments share this spreadsheet but do not know about
// Primary/Backup routing. They can therefore append a valid open question with
// both assignment columns blank. Repair only recent, never-routed questions;
// an explicit Support unassignment is detected from Audit and preserved.
function reconcileLegacyDeploymentQuestionRouting(observedDataVersion, prefetchedSheet, prefetchedRows) {
  return withLock(() => {
    const properties = PropertiesService.getScriptProperties();
    const currentVersion = String(getDataVersion());
    const scannedVersion = String(properties.getProperty(CONFIG_KEY_LEGACY_ROUTING_SCAN_VERSION) || '');
    if (scannedVersion === currentVersion) {
      return { updated: 0, version: currentVersion, emailNotifications: [] };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = prefetchedSheet || ss.getSheetByName(SHEET_QUESTIONS);
    if (!sheet || sheet.getLastRow() < 2) {
      properties.setProperty(CONFIG_KEY_LEGACY_ROUTING_SCAN_VERSION, currentVersion);
      return { updated: 0, version: currentVersion, emailNotifications: [] };
    }

    // Supervisor Desk already reads these rows for display. Reuse that feed
    // whenever available so compatibility routing does not double the main
    // Questions Tracker read.
    const rows = Array.isArray(prefetchedRows)
      ? prefetchedRows
      : sheet.getRange(2, 1, sheet.getLastRow() - 1, Q_WIDTH).getValues();
    const team = _getTeamMembersInternal(true);
    const activeCoordinators = team.filter(member =>
      String(member.status || '').trim().toLowerCase() === 'active' &&
      String(member.category || '').trim().toLowerCase() === 'coordinator'
    );
    const activeSupport = team.filter(member =>
      String(member.status || '').trim().toLowerCase() === 'active' &&
      String(member.category || '').trim().toLowerCase() === 'support'
    );
    const cutoff = new Date().getTime() - LEGACY_ROUTING_LOOKBACK_DAYS * 24 * 60 * 60 * 1000;
    const candidates = [];

    rows.forEach((row, index) => {
      const ticketId = String(row[Q_COL.TICKET_ID - 1] || '').trim();
      const status = String(row[Q_COL.STATUS - 1] || STATUS_OPEN).trim();
      const assigned = normalizeEmail(row[Q_COL.ASSIGNED - 1]);
      const routed = normalizeEmail(row[Q_COL.ROUTED_SUPPORT_EMAIL - 1]);
      const created = row[Q_COL.CREATED - 1] instanceof Date
        ? row[Q_COL.CREATED - 1]
        : new Date(row[Q_COL.CREATED - 1]);
      if (!ticketId || status !== STATUS_OPEN || assigned || routed || isNaN(created.getTime()) || created.getTime() < cutoff) return;

      const askerEmail = normalizeEmail(row[Q_COL.ASKED_BY_EMAIL - 1]);
      const askerName = stripNoraPrefix(String(row[Q_COL.ASKED_BY - 1] || '')).trim().toLowerCase();
      let coordinator = askerEmail
        ? activeCoordinators.find(member => normalizeEmail(member.email) === askerEmail)
        : null;
      if (!coordinator && askerName) {
        const nameMatches = activeCoordinators.filter(member => stripNoraPrefix(member.name).trim().toLowerCase() === askerName);
        if (nameMatches.length === 1) coordinator = nameMatches[0];
      }
      if (!coordinator) return;

      const routeEmails = [coordinator.primarySupportEmail, coordinator.backupSupportEmail]
        .map(normalizeEmail).filter(Boolean);
      let owner = null;
      for (let i = 0; i < routeEmails.length && !owner; i++) {
        const matches = activeSupport.filter(member => normalizeEmail(member.email) === routeEmails[i]);
        if (matches.length === 1) owner = matches[0];
      }
      if (!owner) return;
      candidates.push({ rowIndex: index + 2, row: row, ticketId: ticketId, coordinator: coordinator, owner: owner });
    });

    // Only recent audit entries are needed because candidate questions are
    // themselves limited to the recent compatibility window.
    const manuallyUnassigned = {};
    if (candidates.length) {
      const audit = ensureAuditSheet(ss);
      const auditLastRow = audit.getLastRow();
      if (auditLastRow > 1) {
        const startRow = Math.max(2, auditLastRow - 999);
        const auditRows = audit.getRange(startRow, 4, auditLastRow - startRow + 1, 2).getValues();
        auditRows.forEach(entry => {
          if (String(entry[0] || '').trim() === 'UNASSIGN') manuallyUnassigned[String(entry[1] || '').trim()] = true;
        });
      }
    }

    let updated = 0;
    const emailNotifications = [];
    candidates.forEach(candidate => {
      if (manuallyUnassigned[candidate.ticketId]) return;
      sheet.getRange(candidate.rowIndex, Q_COL.ASSIGNED).setValue(candidate.owner.email);
      sheet.getRange(candidate.rowIndex, Q_COL.ROUTED_SUPPORT_EMAIL).setValue(candidate.owner.email);
      candidate.row[Q_COL.ASSIGNED - 1] = candidate.owner.email;
      candidate.row[Q_COL.ROUTED_SUPPORT_EMAIL - 1] = candidate.owner.email;
      logAudit('AUTO_ROUTE_LEGACY_SUBMISSION', candidate.coordinator.email, candidate.coordinator.name, candidate.ticketId, {
        assignedTo: candidate.owner.email,
        source: 'Older active Support Hub deployment'
      });
      try {
        createDirectTicketNotification(
          candidate.ticketId,
          candidate.row[Q_COL.EVENT - 1],
          candidate.row[Q_COL.QUESTION - 1],
          candidate.owner,
          candidate.coordinator,
          'Automatic Question Routing'
        );
      } catch (notificationError) {
        Logger.log('Legacy question routing notification failed for ' + candidate.ticketId + ': ' + notificationError);
      }
      if (!emailsRepresentSameWorkspaceIdentity(candidate.owner.email, candidate.coordinator.email)) {
        emailNotifications.push({
          recipientEmail: candidate.owner.email,
          recipientName: candidate.owner.name,
          eventName: String(candidate.row[Q_COL.EVENT - 1] || ''),
          preview: String(candidate.row[Q_COL.QUESTION - 1] || ''),
          actorName: candidate.coordinator.name,
          actorEmail: candidate.coordinator.email, ticketId: candidate.ticketId
        });
      }
      updated++;
    });

    const resultingVersion = updated ? bumpDataVersion() : currentVersion;
    properties.setProperty(CONFIG_KEY_LEGACY_ROUTING_SCAN_VERSION, String(resultingVersion));
    return { updated: updated, version: String(resultingVersion), emailNotifications: emailNotifications };
  }, { bumpDataVersion: false });
}

function getQuestionsDataIfChanged(requestingEmail, clientVersion) {
  const startedAt = Date.now();
  // Authorization stays server-side on every poll, but the roster itself is
  // cached briefly. This avoids re-reading Team Setup every 30-60 seconds.
  const member = requireAuthenticatedMember(requestingEmail);
  const authenticatedAt = Date.now();
  const serverVersion = getDataVersion();

  if (clientVersion && clientVersion === serverVersion) {
    return { unchanged: true, version: serverVersion, serverTiming: { totalMs: Date.now() - startedAt, authMs: authenticatedAt - startedAt, dataMs: 0 } };
  }

  // Reuse the member we just verified instead of running setup/auth twice.
  const records = getQuestionsDataForMember(requestingEmail, member);
  // Persistent notifications use their own client refresh path. Returning
  // them here duplicated sheet reads whenever ticket data changed.
  return { unchanged: false, data: records, version: serverVersion, serverTiming: { totalMs: Date.now() - startedAt, authMs: authenticatedAt - startedAt, dataMs: Date.now() - authenticatedAt } };
}

// Fast feed for Supervisor Desk. The desk only displays open questions, so it
// should not read and transmit the entire Answered history on every queue
// change. Other tabs continue using getQuestionsDataIfChanged and receive the
// complete data set when they are opened.
function getSupervisorQuestionsDataIfChanged(requestingEmail, clientVersion) {
  const startedAt = Date.now();
  const member = requireAuthenticatedMember(requestingEmail);
  const authenticatedAt = Date.now();
  if (!isSupportMember(member)) {
    throw new Error('Access denied: this action requires a Supervisor / Support role.');
  }

  let serverVersion = getDataVersion();
  if (clientVersion !== null && clientVersion !== undefined && String(clientVersion) === String(serverVersion)) {
    return { unchanged: true, version: serverVersion, serverTiming: { totalMs: Date.now() - startedAt, authMs: authenticatedAt - startedAt, dataMs: 0 } };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { unchanged: false, data: [], version: serverVersion, serverTiming: { totalMs: Date.now() - startedAt, authMs: authenticatedAt - startedAt, dataMs: Date.now() - authenticatedAt } };
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Q_WIDTH).getValues();
  // Compatibility bridge for questions submitted through an older deployment
  // that is still active against the same spreadsheet. Reuse the rows already
  // loaded for Supervisor Desk to preserve the fast one-read feed.
  const routingRepair = reconcileLegacyDeploymentQuestionRouting(serverVersion, sheet, data);
  serverVersion = String((routingRepair && routingRepair.version) || getDataVersion());
  if (routingRepair && Array.isArray(routingRepair.emailNotifications)) {
    routingRepair.emailNotifications.forEach(n => {
      sendMentionEmailIfEnabled(
        n.recipientEmail, n.recipientName, n.eventName, n.preview,
        n.actorName, n.actorEmail, 'automatically routed question', n.ticketId
      );
    });
  }
  const records = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const qText = String(row[Q_COL.QUESTION - 1] || '').trim();
    if (!qText) continue;

    const ticketId = String(row[Q_COL.TICKET_ID - 1] || '').trim();
    const holdSinceRaw = row[Q_COL.HOLD_SINCE - 1];
    const cumulativeHold = Number(row[Q_COL.HOLD_ACCUM - 1]) || 0;
    const compactQuestion = compactRichContentForFeed(qText);
    records.push({
      id: 'Q_' + (ticketId || (i + 2)),
      ticketId: ticketId,
      sheet: SHEET_QUESTIONS,
      question: compactQuestion.html,
      eventName: String(row[Q_COL.EVENT - 1] || '').trim(),
      askedBy: stripNoraPrefix(String(row[Q_COL.ASKED_BY - 1] || '').trim()),
      askedByEmail: normalizeEmail(row[Q_COL.ASKED_BY_EMAIL - 1]),
      created: safeIsoDate(row[Q_COL.CREATED - 1]),
      caseLink: extractCaseId(String(row[Q_COL.LINK - 1] || '').trim()),
      answer: String(row[Q_COL.ANSWER - 1] || '').trim(),
      contentDeferred: compactQuestion.deferred,
      status: String(row[Q_COL.STATUS - 1] || STATUS_OPEN).trim(),
      assignedTo: normalizeEmail(row[Q_COL.ASSIGNED - 1]),
      hoursElapsed: calculateElapsedHours(row[Q_COL.CREATED - 1], holdSinceRaw, cumulativeHold),
      holdReason: String(row[Q_COL.HOLD_REASON - 1] || '').trim(),
      holdSince: holdSinceRaw ? safeIsoDate(holdSinceRaw) : '',
      isFollowUp: isFollowUpValue(row[Q_COL.FOLLOWUP - 1]),
      assignNotes: String(row[Q_COL.ASSIGN_NOTES - 1] || '').trim(),
      statusChangedAt: row[Q_COL.STATUS_CHANGED_AT - 1] ? safeIsoDate(row[Q_COL.STATUS_CHANGED_AT - 1]) : '',
      statusChangedBy: stripNoraPrefix(String(row[Q_COL.STATUS_CHANGED_BY - 1] || '').trim()),
      category: normalizeWorkCategory(row[Q_COL.WORK_CATEGORY - 1] || inferWorkCategory(qText, row[Q_COL.EVENT - 1])),
      eventDate: normalizeEventDate(row[Q_COL.EVENT_DATE - 1]),
      talentRole: String(row[Q_COL.TALENT_ROLE - 1] || '').trim(),
      routedSupportEmail: normalizeEmail(row[Q_COL.ROUTED_SUPPORT_EMAIL - 1]),
      eventThisWeekend: isActiveEventThisWeekend(row[Q_COL.EVENT_THIS_WEEKEND - 1] === true, row[Q_COL.EVENT_DATE - 1]),
      eventThisWeekendFlag: row[Q_COL.EVENT_THIS_WEEKEND - 1] === true,
      attentionRequested: attentionWasRequested(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_SET_AT - 1]),
      attentionToday: isAttentionActive(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_UNTIL - 1]),
      attentionReason: String(row[Q_COL.ATTENTION_REASON - 1] || '').trim(),
      attentionSetAt: row[Q_COL.ATTENTION_SET_AT - 1] ? safeIsoDate(row[Q_COL.ATTENTION_SET_AT - 1]) : '',
      attentionSetBy: stripNoraPrefix(String(row[Q_COL.ATTENTION_SET_BY - 1] || '').trim()),
      attentionUntil: normalizeEventDate(row[Q_COL.ATTENTION_UNTIL - 1]),
      attentionClearedAt: row[Q_COL.ATTENTION_CLEARED_AT - 1] ? safeIsoDate(row[Q_COL.ATTENTION_CLEARED_AT - 1]) : '',
      attentionClearedBy: stripNoraPrefix(String(row[Q_COL.ATTENTION_CLEARED_BY - 1] || '').trim()),
      attentionFirstReviewAt: row[Q_COL.ATTENTION_FIRST_REVIEW_AT - 1] ? safeIsoDate(row[Q_COL.ATTENTION_FIRST_REVIEW_AT - 1]) : '',
      attentionFirstReviewBy: stripNoraPrefix(String(row[Q_COL.ATTENTION_FIRST_REVIEW_BY - 1] || '').trim()),
      requesterFlagged: row[Q_COL.REQUESTER_FLAGGED - 1] === true || !!row[Q_COL.REQUESTER_FLAGGED_AT - 1],
      requesterFlaggedAt: row[Q_COL.REQUESTER_FLAGGED_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAGGED_AT - 1]) : '',
      requesterFlaggedBy: stripNoraPrefix(String(row[Q_COL.REQUESTER_FLAGGED_BY - 1] || '').trim()),
      requesterFlagReason: String(row[Q_COL.REQUESTER_FLAG_REASON - 1] || '').trim(),
      requesterFlagAcknowledgedAt: row[Q_COL.REQUESTER_FLAG_ACK_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAG_ACK_AT - 1]) : '',
      requesterFlagAcknowledgedBy: stripNoraPrefix(String(row[Q_COL.REQUESTER_FLAG_ACK_BY - 1] || '').trim()),
      requesterFlagActive: (row[Q_COL.REQUESTER_FLAGGED - 1] === true || !!row[Q_COL.REQUESTER_FLAGGED_AT - 1]) && !row[Q_COL.REQUESTER_FLAG_ACK_AT - 1],
      isRead: true
    });
  }

  return { unchanged: false, data: records, version: serverVersion, serverTiming: { totalMs: Date.now() - startedAt, authMs: authenticatedAt - startedAt, dataMs: Date.now() - authenticatedAt } };
}

function getTicketRichContent(ticketId, sourceSheet, requestingEmail) {
  requireAuthenticatedMember(requestingEmail);
  const sheetName = String(sourceSheet || '').trim();
  let target;
  let row;
  if (sheetName === SHEET_ANSWERED) {
    target = requireAnsweredRow(ticketId);
    row = target.sheet.getRange(target.rowIndex, 1, 1, A_WIDTH).getValues()[0];
    return {
      success: true,
      ticketId: String(ticketId || ''),
      sheet: SHEET_ANSWERED,
      question: String(row[A_COL.QUESTION - 1] || ''),
      answer: String(row[A_COL.ANSWER - 1] || '')
    };
  }
  if (sheetName === SHEET_QUESTIONS) {
    target = requireQuestionRow(ticketId);
    row = target.sheet.getRange(target.rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    return {
      success: true,
      ticketId: String(ticketId || ''),
      sheet: SHEET_QUESTIONS,
      question: String(row[Q_COL.QUESTION - 1] || ''),
      answer: String(row[Q_COL.ANSWER - 1] || '')
    };
  }
  throw new Error('This ticket source is not valid.');
}

// Targeted relationship reads keep ticket/task navigation independent from the
// large Supervisor, Answered, and Reporting feeds. Any authenticated active
// team member may follow an existing relationship, but this does not grant
// permission to edit, reassign, complete, or delete the returned record.
function getRelatedTasksForTicket(ticketId, requestingEmail) {
  requireAuthenticatedMember(requestingEmail);
  const id = String(ticketId || '').trim();
  if (!id) return [];
  const idKey = id.toLowerCase();
  const sheet = ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, T_WIDTH).getValues();
  return rows.map(taskRowToObject).filter(function(task) {
    const relatedId = String(task.relatedTicketId || '').trim().toLowerCase();
    const parentId = String(task.parentId || '').trim().toLowerCase();
    const parentType = String(task.parentType || '').trim().toLowerCase();
    return relatedId === idKey || (parentType === 'question' && parentId === idKey);
  });
}

function getTaskForNavigation(taskId, requestingEmail) {
  requireAuthenticatedMember(requestingEmail);
  const target = requireTaskRow(taskId);
  return taskRowToObject(target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0]);
}

// Returns the connected task family in one sheet read so the client can draw
// the parent/follow-up chain without issuing a request for every node.
function getTaskChain(taskId, requestingEmail) {
  requireAuthenticatedMember(requestingEmail);
  const wantedId = String(taskId || '').trim();
  if (!wantedId) throw new Error('Task ID is required.');
  const sheet = ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (sheet.getLastRow() < 2) throw new Error('This task no longer exists.');
  const tasks = sheet.getRange(2, 1, sheet.getLastRow() - 1, T_WIDTH).getValues()
    .map(taskRowToObject).filter(function(task) { return !!task.taskId; });
  const byId = {};
  tasks.forEach(function(task) { byId[task.taskId] = task; });
  if (!byId[wantedId]) throw new Error('This task no longer exists.');

  let root = byId[wantedId];
  const ancestorIds = new Set([root.taskId]);
  while (String(root.parentType || '').toLowerCase() === 'task' && root.parentId && byId[root.parentId] && !ancestorIds.has(root.parentId)) {
    root = byId[root.parentId];
    ancestorIds.add(root.taskId);
  }

  const childrenByParent = {};
  tasks.forEach(function(task) {
    if (String(task.parentType || '').toLowerCase() !== 'task' || !task.parentId) return;
    if (!childrenByParent[task.parentId]) childrenByParent[task.parentId] = [];
    childrenByParent[task.parentId].push(task);
  });
  Object.keys(childrenByParent).forEach(function(parentId) {
    childrenByParent[parentId].sort(function(a, b) { return new Date(a.createdAt || 0) - new Date(b.createdAt || 0); });
  });

  const connected = [];
  const seen = new Set();
  const visit = function(task) {
    if (!task || seen.has(task.taskId) || connected.length >= 200) return;
    seen.add(task.taskId);
    connected.push(task);
    (childrenByParent[task.taskId] || []).forEach(visit);
  };
  visit(root);
  return { rootTaskId: root.taskId, tasks: connected, truncated: connected.length >= 200 };
}

function navigationQuestionFromRow(row, sheetName, requestingEmail) {
  const requesterEmail = normalizeEmail(requestingEmail);
  const answered = sheetName === SHEET_ANSWERED;
  const C = answered ? A_COL : Q_COL;
  const ticketId = String(row[C.TICKET_ID - 1] || '').trim();
  const qText = String(row[C.QUESTION - 1] || '').trim();
  const record = {
    id: (answered ? 'A_' : 'Q_') + ticketId,
    ticketId: ticketId,
    sheet: sheetName,
    question: qText,
    eventName: String(row[C.EVENT - 1] || '').trim(),
    askedBy: stripNoraPrefix(String(row[C.ASKED_BY - 1] || '').trim()),
    askedByEmail: normalizeEmail(row[C.ASKED_BY_EMAIL - 1]),
    created: safeIsoDate(row[C.CREATED - 1]),
    caseLink: extractCaseId(String(row[C.LINK - 1] || '').trim()),
    answer: String(row[C.ANSWER - 1] || '').trim(),
    contentDeferred: false,
    status: String(row[C.STATUS - 1] || (answered ? STATUS_ANSWERED : STATUS_OPEN)).trim(),
    category: normalizeWorkCategory(row[C.WORK_CATEGORY - 1] || inferWorkCategory(qText, row[C.EVENT - 1])),
    eventDate: normalizeEventDate(row[C.EVENT_DATE - 1]),
    talentRole: String(row[C.TALENT_ROLE - 1] || '').trim(),
    routedSupportEmail: normalizeEmail(row[C.ROUTED_SUPPORT_EMAIL - 1]),
    eventThisWeekend: isActiveEventThisWeekend(row[C.EVENT_THIS_WEEKEND - 1] === true, row[C.EVENT_DATE - 1]),
    eventThisWeekendFlag: row[C.EVENT_THIS_WEEKEND - 1] === true
  };
  if (answered) {
    const readByList = parseReadByList(row[A_COL.READ_BY - 1]);
    const firstReadMap = parseFirstReadMap(row[A_COL.FIRST_READ_AT - 1]);
    let turnaround = row[A_COL.TURNAROUND - 1];
    if (turnaround === '' || turnaround === undefined || isNaN(turnaround) || Number(turnaround) < 0) {
      turnaround = Math.max(0, (new Date(row[A_COL.ANSWERED - 1]) - new Date(row[A_COL.CREATED - 1])) / 3600000).toFixed(1);
    } else turnaround = Math.abs(Number(turnaround)).toFixed(1);
    record.answeredDate = safeIsoDate(row[A_COL.ANSWERED - 1]);
    record.turnaroundHours = turnaround;
    record.holdHoursExcluded = Number(row[A_COL.HOLD_HOURS - 1]) || 0;
    record.answeredBy = stripNoraPrefix(String(row[A_COL.ANSWERED_BY - 1] || 'Supervisor').trim());
    record.answeredByEmail = normalizeEmail(row[A_COL.ANSWERED_BY_EMAIL - 1]);
    record.isRead = readByList.indexOf('*') !== -1 || readByList.indexOf(requesterEmail) !== -1;
    record.firstReadAt = firstReadMap[requesterEmail] || '';
    record.lastUpdatedAt = safeIsoDate(row[A_COL.LAST_UPDATED_AT - 1] || row[A_COL.ANSWERED - 1]);
    record.hasUnreadUpdate = !record.isRead && !!row[A_COL.LAST_UPDATED_AT - 1] &&
      new Date(record.lastUpdatedAt).getTime() > new Date(record.answeredDate).getTime();
    record.attentionRequested = row[A_COL.ATTENTION_REQUESTED - 1] === true || !!row[A_COL.ATTENTION_SET_AT - 1];
    record.attentionToday = false;
    record.attentionReason = String(row[A_COL.ATTENTION_REASON - 1] || '').trim();
    record.attentionSetAt = row[A_COL.ATTENTION_SET_AT - 1] ? safeIsoDate(row[A_COL.ATTENTION_SET_AT - 1]) : '';
    record.attentionSetBy = stripNoraPrefix(String(row[A_COL.ATTENTION_SET_BY - 1] || '').trim());
    record.attentionUntil = normalizeEventDate(row[A_COL.ATTENTION_UNTIL - 1]);
  } else {
    const holdSinceRaw = row[Q_COL.HOLD_SINCE - 1];
    record.assignedTo = normalizeEmail(row[Q_COL.ASSIGNED - 1]);
    record.hoursElapsed = calculateElapsedHours(row[Q_COL.CREATED - 1], holdSinceRaw, Number(row[Q_COL.HOLD_ACCUM - 1]) || 0);
    record.holdReason = String(row[Q_COL.HOLD_REASON - 1] || '').trim();
    record.holdSince = holdSinceRaw ? safeIsoDate(holdSinceRaw) : '';
    record.isFollowUp = isFollowUpValue(row[Q_COL.FOLLOWUP - 1]);
    record.assignNotes = String(row[Q_COL.ASSIGN_NOTES - 1] || '').trim();
    record.statusChangedAt = row[Q_COL.STATUS_CHANGED_AT - 1] ? safeIsoDate(row[Q_COL.STATUS_CHANGED_AT - 1]) : '';
    record.statusChangedBy = stripNoraPrefix(String(row[Q_COL.STATUS_CHANGED_BY - 1] || '').trim());
    record.attentionRequested = attentionWasRequested(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_SET_AT - 1]);
    record.attentionToday = isAttentionActive(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_UNTIL - 1]);
    record.attentionReason = String(row[Q_COL.ATTENTION_REASON - 1] || '').trim();
    record.attentionSetAt = row[Q_COL.ATTENTION_SET_AT - 1] ? safeIsoDate(row[Q_COL.ATTENTION_SET_AT - 1]) : '';
    record.attentionSetBy = stripNoraPrefix(String(row[Q_COL.ATTENTION_SET_BY - 1] || '').trim());
    record.attentionUntil = normalizeEventDate(row[Q_COL.ATTENTION_UNTIL - 1]);
    record.requesterFlagged = row[Q_COL.REQUESTER_FLAGGED - 1] === true || !!row[Q_COL.REQUESTER_FLAGGED_AT - 1];
    record.requesterFlaggedAt = row[Q_COL.REQUESTER_FLAGGED_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAGGED_AT - 1]) : '';
    record.requesterFlaggedBy = stripNoraPrefix(String(row[Q_COL.REQUESTER_FLAGGED_BY - 1] || '').trim());
    record.requesterFlagReason = String(row[Q_COL.REQUESTER_FLAG_REASON - 1] || '').trim();
    record.requesterFlagAcknowledgedAt = row[Q_COL.REQUESTER_FLAG_ACK_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAG_ACK_AT - 1]) : '';
    record.requesterFlagAcknowledgedBy = stripNoraPrefix(String(row[Q_COL.REQUESTER_FLAG_ACK_BY - 1] || '').trim());
    record.requesterFlagActive = record.requesterFlagged && !record.requesterFlagAcknowledgedAt;
    record.isRead = true;
  }
  return record;
}

function getTicketForNavigation(ticketId, requestingEmail) {
  requireAuthenticatedMember(requestingEmail);
  const id = String(ticketId || '').trim();
  if (!id) throw new Error('The linked ticket ID is missing.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (qSheet) {
    const qRow = findRowIndexByTicketId(qSheet, id, Q_COL.TICKET_ID);
    if (qRow !== -1) return navigationQuestionFromRow(qSheet.getRange(qRow, 1, 1, Q_WIDTH).getValues()[0], SHEET_QUESTIONS, requestingEmail);
  }
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);
  if (aSheet) {
    const aRow = findRowIndexByTicketId(aSheet, id, A_COL.TICKET_ID);
    if (aRow !== -1) return navigationQuestionFromRow(aSheet.getRange(aRow, 1, 1, A_WIDTH).getValues()[0], SHEET_ANSWERED, requestingEmail);
  }
  throw new Error('The linked ticket no longer exists.');
}

// ==========================================
// PHASE 13F - SERVER SEARCH + GROUP PAGINATION
// ==========================================
// Search stays read-only and outside the global write lock. TextFinder first
// identifies candidate rows, then only those rows are read and permission
// filtered. This keeps the browser from retaining and rescanning the complete
// ticket/task history on every keystroke.
const GLOBAL_SEARCH_MAX_QUERY_CHARS = 160;
const GLOBAL_SEARCH_MAX_CANDIDATE_ROWS = 1200;
const GLOBAL_SEARCH_DEFAULT_GROUP_SIZE = 6;
const GLOBAL_SEARCH_MAX_GROUP_SIZE = 12;

function normalizeServerSearchText_(value) {
  return stripHtmlToText(value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function serverSearchHaystack_(values) {
  return (values || []).map(normalizeServerSearchText_).join(' | ');
}

function serverSearchSeed_(query) {
  const tokens = String(query || '').split(/\s+/).filter(Boolean);
  if (!tokens.length) return '';
  tokens.sort(function(a, b) { return b.length - a.length; });
  return tokens[0];
}

function findServerSearchRows_(sheet, query, maxRows) {
  if (!sheet || sheet.getLastRow() < 2) return { rows: [], capped: false, matches: 0 };
  const seed = serverSearchSeed_(query);
  if (!seed) return { rows: [], capped: false, matches: 0 };
  const found = sheet.createTextFinder(seed)
    .matchCase(false)
    .matchEntireCell(false)
    .useRegularExpression(false)
    .findAll();
  const unique = {};
  (found || []).forEach(function(range) {
    const row = Number(range.getRow());
    if (row >= 2) unique[row] = true;
  });
  const allRows = Object.keys(unique).map(Number).sort(function(a, b) { return b - a; });
  const limit = Math.max(1, Number(maxRows) || GLOBAL_SEARCH_MAX_CANDIDATE_ROWS);
  return { rows: allRows.slice(0, limit), capped: allRows.length > limit, matches: allRows.length };
}

// Consecutive candidate rows are read in one range call. A 20-result query
// normally needs only one or two Sheet reads instead of one read per result.
function readServerSearchRows_(sheet, rowNumbers, width) {
  const sorted = (rowNumbers || []).slice().sort(function(a, b) { return a - b; });
  if (!sheet || !sorted.length) return [];
  const groups = [];
  let start = sorted[0];
  let previous = sorted[0];
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === previous + 1) {
      previous = sorted[i];
      continue;
    }
    groups.push({ start: start, count: previous - start + 1 });
    start = previous = sorted[i];
  }
  groups.push({ start: start, count: previous - start + 1 });
  const out = [];
  groups.forEach(function(group) {
    const values = sheet.getRange(group.start, 1, group.count, width).getValues();
    values.forEach(function(row, index) { out.push({ rowNumber: group.start + index, row: row }); });
  });
  return out;
}

function serverSearchScore_(query, id, title, haystack) {
  const q = normalizeServerSearchText_(query);
  const cleanId = normalizeServerSearchText_(id);
  const cleanTitle = normalizeServerSearchText_(title);
  if (cleanId && cleanId === q) return 140;
  if (cleanTitle && cleanTitle === q) return 120;
  if (cleanId && cleanId.indexOf(q) === 0) return 105;
  if (cleanTitle && cleanTitle.indexOf(q) === 0) return 90;
  if (cleanTitle && cleanTitle.indexOf(q) !== -1) return 70;
  return String(haystack || '').indexOf(q) !== -1 ? 40 : 0;
}

function serverSearchTime_(value) {
  const date = value instanceof Date ? value : new Date(value || 0);
  return isNaN(date.getTime()) ? 0 : date.getTime();
}

function sortServerSearchResults_(items) {
  return (items || []).sort(function(a, b) {
    return (Number(b._score) - Number(a._score)) || (Number(b._sortAt) - Number(a._sortAt)) || String(a.title || '').localeCompare(String(b.title || ''));
  });
}

function buildServerSearchTeamMaps_(team) {
  const byEmail = {};
  const byName = {};
  function prefer(existing, candidate) {
    if (!existing) return candidate;
    if (isAdminMember(existing) && !isAdminMember(candidate)) return candidate;
    return existing;
  }
  (team || []).forEach(function(member) {
    const email = normalizeEmail(member.email);
    const name = normalizeServerSearchText_(stripNoraPrefix(member.name));
    if (email) byEmail[email] = prefer(byEmail[email], member);
    if (name) byName[name] = prefer(byName[name], member);
  });
  return { byEmail: byEmail, byName: byName };
}

function serverSearchDisplayName_(emailOrName, maps) {
  const email = normalizeEmail(emailOrName);
  const member = email && maps && maps.byEmail ? maps.byEmail[email] : null;
  return member && member.name ? member.name : stripNoraPrefix(String(emailOrName || '').trim());
}

function serverSearchTicketVisible_(row, C, requesterIsSupport, maps) {
  if (requesterIsSupport) return true;
  const email = normalizeEmail(row[C.ASKED_BY_EMAIL - 1]);
  const name = normalizeServerSearchText_(stripNoraPrefix(row[C.ASKED_BY - 1]));
  const asker = (email && maps.byEmail[email]) || (name && maps.byName[name]) || null;
  return !(asker && (isSupportMember(asker) || isAdminMember(asker)));
}

function buildServerTicketSearchResult_(entry, sheetName, query, requesterIsSupport, maps) {
  const row = entry.row;
  const answered = sheetName === SHEET_ANSWERED;
  const C = answered ? A_COL : Q_COL;
  if (!serverSearchTicketVisible_(row, C, requesterIsSupport, maps)) return null;
  const ticketId = String(row[C.TICKET_ID - 1] || '').trim();
  if (!ticketId) return null;
  const question = stripHtmlToText(row[C.QUESTION - 1]);
  const answer = stripHtmlToText(row[C.ANSWER - 1]);
  const eventName = String(row[C.EVENT - 1] || '').trim();
  const askedBy = stripNoraPrefix(String(row[C.ASKED_BY - 1] || '').trim());
  const category = normalizeWorkCategory(row[C.WORK_CATEGORY - 1] || inferWorkCategory(question, eventName));
  const assignedEmail = answered ? normalizeEmail(row[A_COL.ANSWERED_BY_EMAIL - 1]) : normalizeEmail(row[Q_COL.ASSIGNED - 1]);
  const assignedName = answered
    ? stripNoraPrefix(String(row[A_COL.ANSWERED_BY - 1] || '').trim())
    : serverSearchDisplayName_(assignedEmail, maps);
  const status = answered ? 'Answered' : (String(row[Q_COL.STATUS - 1] || STATUS_OPEN).trim() === STATUS_HOLD ? 'On Hold' : (isFollowUpValue(row[Q_COL.FOLLOWUP - 1]) ? 'Follow-Up' : 'Open'));
  const values = answered
    ? [ticketId, question, answer, eventName, askedBy, row[A_COL.ASKED_BY_EMAIL - 1], assignedName, assignedEmail, row[A_COL.STATUS - 1], category, row[A_COL.LINK - 1], row[A_COL.TALENT_ROLE - 1], row[A_COL.EVENT_DATE - 1]]
    : [ticketId, question, answer, eventName, askedBy, row[Q_COL.ASKED_BY_EMAIL - 1], assignedName, assignedEmail, row[Q_COL.STATUS - 1], category, row[Q_COL.LINK - 1], row[Q_COL.ASSIGN_NOTES - 1], row[Q_COL.TALENT_ROLE - 1], row[Q_COL.EVENT_DATE - 1]];
  const haystack = serverSearchHaystack_(values);
  const score = serverSearchScore_(query, ticketId, eventName || question, haystack);
  if (!score) return null;
  const dateValue = answered ? (row[A_COL.LAST_UPDATED_AT - 1] || row[A_COL.ANSWERED - 1] || row[A_COL.CREATED - 1]) : row[Q_COL.CREATED - 1];
  return {
    type: 'ticket',
    id: ticketId,
    ticketId: ticketId,
    sheet: sheetName,
    title: eventName || question.slice(0, 90) || ticketId,
    meta: status + ' · ' + category + ' · Asked by ' + (askedBy || 'Unknown') + (assignedName ? ' · ' + (answered ? 'Answered by ' : 'Assigned to ') + assignedName : ''),
    detail: question.slice(0, 150),
    _score: score,
    _sortAt: serverSearchTime_(dateValue)
  };
}

function buildServerTaskSearchResult_(entry, query, actor, maps) {
  const task = taskRowToObject(entry.row);
  if (!task.taskId || !task.title) return null;
  const requesterEmail = normalizeEmail(actor.email);
  if (!isSupportMember(actor) && task.assignedToEmail !== requesterEmail && task.createdByEmail !== requesterEmail) return null;
  const values = [task.taskId, task.title, task.instructions, task.assignedTo, task.assignedToEmail, task.createdBy, task.createdByEmail, task.status, task.category, task.relatedClientTalent, task.relatedLink, task.dueDate, task.source, task.relatedTicketId, task.relatedEvent, task.parentType, task.parentId, task.supportOwnerEmail, task.completedBy, task.completedByEmail, task.completionNote];
  const haystack = serverSearchHaystack_(values);
  const score = serverSearchScore_(query, task.taskId, task.title, haystack);
  if (!score) return null;
  const assignee = task.assignedTo || serverSearchDisplayName_(task.assignedToEmail, maps) || 'Unassigned';
  return {
    type: 'task',
    id: task.taskId,
    taskId: task.taskId,
    title: task.title,
    meta: (task.status || TASK_STATUS_PENDING) + ' · ' + (task.category || 'General / Other') + ' · Assigned to ' + assignee,
    detail: (task.dueDate ? 'Due ' + task.dueDate + ' · ' : '') + stripHtmlToText(task.instructions).slice(0, 130),
    _score: score,
    _sortAt: serverSearchTime_(task.updatedAt || task.createdAt)
  };
}

function buildServerPeerSearchResult_(entry, query, actor) {
  const peer = peerRowToObject(entry.row);
  if (!peer.id) return null;
  const requesterEmail = normalizeEmail(actor.email);
  const canSeeAll = isSupportMember(actor) || isAdminMember(actor);
  if (!canSeeAll && peer.askedByEmail !== requesterEmail && peer.assignedToEmail !== requesterEmail) return null;
  const threadText = (peer.thread || []).map(function(item) { return item && item.text ? item.text : ''; }).join(' ');
  const values = [peer.id, peer.question, peer.eventName, peer.askedBy, peer.askedByEmail, peer.assignedTo, peer.assignedToEmail, peer.status, threadText, peer.caseLink];
  const haystack = serverSearchHaystack_(values);
  const score = serverSearchScore_(query, peer.id, peer.eventName || peer.question, haystack);
  if (!score) return null;
  return {
    type: 'peer',
    id: peer.id,
    title: peer.eventName || 'Peer Question',
    meta: (peer.status || PEER_STATUS_OPEN) + ' · ' + (peer.askedBy || 'Unknown') + ' → ' + (peer.assignedTo || 'Unassigned'),
    detail: stripHtmlToText(peer.question).slice(0, 140),
    record: peer,
    _score: score,
    _sortAt: serverSearchTime_(peer.updated || peer.created)
  };
}

function buildServerTeamSearchResults_(team, query) {
  return (team || []).map(function(member) {
    const values = [member.name, member.title, member.status, member.category, member.email];
    const haystack = serverSearchHaystack_(values);
    const score = serverSearchScore_(query, member.email, member.name, haystack);
    if (!score) return null;
    return {
      type: 'team',
      id: normalizeEmail(member.email) + '|' + normalizeServerSearchText_(member.category),
      title: member.name || member.email || 'Team member',
      meta: (member.title || member.category || 'Team') + ' · ' + (member.status || 'Unknown'),
      detail: member.email || '',
      _score: score,
      _sortAt: 0
    };
  }).filter(Boolean);
}

function stripServerSearchPrivateFields_(item) {
  const clean = {};
  Object.keys(item || {}).forEach(function(key) {
    if (key.charAt(0) !== '_') clean[key] = item[key];
  });
  return clean;
}

function searchSupportHub(requestingEmail, rawQuery, page, groupPageSize) {
  const startedAt = Date.now();
  const actor = requireAuthenticatedMember(requestingEmail);
  const authenticatedAt = Date.now();
  const query = normalizeServerSearchText_(String(rawQuery || '').slice(0, GLOBAL_SEARCH_MAX_QUERY_CHARS));
  const currentPage = Math.max(1, Math.floor(Number(page) || 1));
  const pageSize = Math.max(1, Math.min(GLOBAL_SEARCH_MAX_GROUP_SIZE, Math.floor(Number(groupPageSize) || GLOBAL_SEARCH_DEFAULT_GROUP_SIZE)));
  if (query.length < 2 || query.charAt(0) === '@') {
    return { query: query, page: currentPage, pageSize: pageSize, total: 0, counts: { ticket:0, task:0, peer:0, team:0 }, results: [], hasMore:false, hasPrevious:currentPage>1, capped:false, serverTiming:{ totalMs:Date.now()-startedAt, authMs:authenticatedAt-startedAt, searchMs:Date.now()-authenticatedAt } };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const team = _getTeamMembersInternal();
  const maps = buildServerSearchTeamMaps_(team);
  const requesterIsSupport = isSupportMember(actor);
  const groups = { ticket: [], peer: [], task: [], team: [] };
  let capped = false;
  const candidateCounts = { questions:0, answered:0, tasks:0, peers:0 };

  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  const qMatches = findServerSearchRows_(qSheet, query, GLOBAL_SEARCH_MAX_CANDIDATE_ROWS);
  candidateCounts.questions = qMatches.matches;
  capped = capped || qMatches.capped;
  readServerSearchRows_(qSheet, qMatches.rows, Q_WIDTH).forEach(function(entry) {
    const result = buildServerTicketSearchResult_(entry, SHEET_QUESTIONS, query, requesterIsSupport, maps);
    if (result) groups.ticket.push(result);
  });

  const aSheet = ss.getSheetByName(SHEET_ANSWERED);
  const aMatches = findServerSearchRows_(aSheet, query, GLOBAL_SEARCH_MAX_CANDIDATE_ROWS);
  candidateCounts.answered = aMatches.matches;
  capped = capped || aMatches.capped;
  readServerSearchRows_(aSheet, aMatches.rows, A_WIDTH).forEach(function(entry) {
    const result = buildServerTicketSearchResult_(entry, SHEET_ANSWERED, query, requesterIsSupport, maps);
    if (result) groups.ticket.push(result);
  });

  const taskSheet = ss.getSheetByName(SHEET_TASKS);
  const taskMatches = findServerSearchRows_(taskSheet, query, GLOBAL_SEARCH_MAX_CANDIDATE_ROWS);
  candidateCounts.tasks = taskMatches.matches;
  capped = capped || taskMatches.capped;
  readServerSearchRows_(taskSheet, taskMatches.rows, T_WIDTH).forEach(function(entry) {
    const result = buildServerTaskSearchResult_(entry, query, actor, maps);
    if (result) groups.task.push(result);
  });

  const peerSheet = ss.getSheetByName(SHEET_PEER_QUESTIONS);
  const peerMatches = findServerSearchRows_(peerSheet, query, GLOBAL_SEARCH_MAX_CANDIDATE_ROWS);
  candidateCounts.peers = peerMatches.matches;
  capped = capped || peerMatches.capped;
  readServerSearchRows_(peerSheet, peerMatches.rows, P_WIDTH).forEach(function(entry) {
    const result = buildServerPeerSearchResult_(entry, query, actor);
    if (result) groups.peer.push(result);
  });

  groups.team = buildServerTeamSearchResults_(team, query);
  Object.keys(groups).forEach(function(type) { sortServerSearchResults_(groups[type]); });

  const counts = {};
  const results = [];
  const offset = (currentPage - 1) * pageSize;
  ['ticket','peer','task','team'].forEach(function(type) {
    counts[type] = groups[type].length;
    groups[type].slice(offset, offset + pageSize).forEach(function(item) { results.push(stripServerSearchPrivateFields_(item)); });
  });
  const total = Object.keys(counts).reduce(function(sum, type) { return sum + counts[type]; }, 0);
  const hasMore = Object.keys(groups).some(function(type) { return groups[type].length > offset + pageSize; });
  const finishedAt = Date.now();
  return {
    query: query,
    page: currentPage,
    pageSize: pageSize,
    total: total,
    counts: counts,
    results: results,
    hasMore: hasMore,
    hasPrevious: currentPage > 1,
    capped: capped,
    candidateCounts: candidateCounts,
    serverTiming: { totalMs: finishedAt - startedAt, authMs: authenticatedAt - startedAt, searchMs: finishedAt - authenticatedAt }
  };
}

// Task relationship picker: ticket-only search across both active and answered
// Support Hub tickets. This reuses the Phase 13F server-search visibility rules,
// but avoids searching tasks, peers, or team members when the user only needs a
// Related Ticket.
function searchRelatedTickets(requestingEmail, rawQuery, limit) {
  const actor = requireAuthenticatedMember(requestingEmail);
  const query = normalizeServerSearchText_(String(rawQuery || '').slice(0, GLOBAL_SEARCH_MAX_QUERY_CHARS));
  const maxResults = Math.max(1, Math.min(20, Math.floor(Number(limit) || 10)));
  if (query.length < 2 || query.charAt(0) === '@') return { query: query, total: 0, results: [], capped: false };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const team = _getTeamMembersInternal();
  const maps = buildServerSearchTeamMaps_(team);
  const requesterIsSupport = isSupportMember(actor);
  const items = [];
  let capped = false;

  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  const qMatches = findServerSearchRows_(qSheet, query, GLOBAL_SEARCH_MAX_CANDIDATE_ROWS);
  capped = capped || qMatches.capped;
  readServerSearchRows_(qSheet, qMatches.rows, Q_WIDTH).forEach(function(entry) {
    const result = buildServerTicketSearchResult_(entry, SHEET_QUESTIONS, query, requesterIsSupport, maps);
    if (result) items.push(result);
  });

  const aSheet = ss.getSheetByName(SHEET_ANSWERED);
  const aMatches = findServerSearchRows_(aSheet, query, GLOBAL_SEARCH_MAX_CANDIDATE_ROWS);
  capped = capped || aMatches.capped;
  readServerSearchRows_(aSheet, aMatches.rows, A_WIDTH).forEach(function(entry) {
    const result = buildServerTicketSearchResult_(entry, SHEET_ANSWERED, query, requesterIsSupport, maps);
    if (result) items.push(result);
  });

  sortServerSearchResults_(items);
  const seen = {};
  const unique = [];
  items.forEach(function(item) {
    const key = String(item.ticketId || '').toLowerCase();
    if (!key || seen[key]) return;
    seen[key] = true;
    unique.push(stripServerSearchPrivateFields_(item));
  });

  return { query: query, total: unique.length, results: unique.slice(0, maxResults), capped: capped || unique.length > maxResults };
}

function normalizeQuestionDuplicateText(value) {
  return stripHtmlToText(value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function normalizeQuestionDuplicateLink(value) {
  // Treat a stored Salesforce record id and a full Salesforce URL containing
  // that id as the same Case. For other links, ignore casing and trailing
  // slashes but retain the full path/query so unrelated records never match.
  const extracted = extractCaseId(String(value || '').trim());
  return String(extracted || '').trim().replace(/\/+$/, '').toLowerCase();
}

function questionDuplicateCacheKey(normalizedQuestion, normalizedEvent, normalizedLink) {
  return 'QUESTION_DUP_V1_' + stableCacheHash(normalizedQuestion + '\u001f' + normalizedEvent + '\u001f' + normalizedLink);
}

function cacheDuplicateQuestionMatch_(normalizedQuestion, normalizedEvent, normalizedLink, match) {
  if (!match || !match.ticketId || !match.sheet) return;
  safeScriptCachePut(
    questionDuplicateCacheKey(normalizedQuestion, normalizedEvent, normalizedLink),
    JSON.stringify(match),
    DUPLICATE_MATCH_CACHE_SECONDS
  );
}

function readVerifiedCachedDuplicateQuestion_(spreadsheet, normalizedQuestion, normalizedEvent, normalizedLink) {
  const key = questionDuplicateCacheKey(normalizedQuestion, normalizedEvent, normalizedLink);
  const raw = safeScriptCacheGet(key);
  if (!raw) return null;
  try {
    const cached = JSON.parse(raw);
    const isAnswered = cached.sheet === SHEET_ANSWERED;
    const C = isAnswered ? A_COL : Q_COL;
    const width = Math.max(C.QUESTION, C.EVENT, C.LINK, C.STATUS, C.TICKET_ID, C.ASKED_BY, C.CREATED);
    const sheet = spreadsheet.getSheetByName(cached.sheet);
    if (!sheet) throw new Error('Cached sheet is unavailable.');
    const rowIndex = findRowIndexByTicketId(sheet, cached.ticketId, C.TICKET_ID);
    if (rowIndex < 2) throw new Error('Cached ticket moved or was removed.');
    const row = sheet.getRange(rowIndex, 1, 1, width).getValues()[0];
    if (normalizeQuestionDuplicateText(row[C.QUESTION - 1]) !== normalizedQuestion ||
        normalizeQuestionDuplicateText(row[C.EVENT - 1]) !== normalizedEvent ||
        normalizeQuestionDuplicateLink(row[C.LINK - 1]) !== normalizedLink) {
      throw new Error('Cached ticket details changed.');
    }
    return {
      ticketId: String(row[C.TICKET_ID - 1] || ''),
      sheet: cached.sheet,
      status: isAnswered ? 'Answered' : String(row[C.STATUS - 1] || STATUS_OPEN),
      askedBy: String(row[C.ASKED_BY - 1] || ''),
      createdAt: row[C.CREATED - 1] instanceof Date && !isNaN(row[C.CREATED - 1].getTime())
        ? row[C.CREATED - 1].toISOString() : String(row[C.CREATED - 1] || '')
    };
  } catch (e) {
    safeScriptCacheRemove(key);
    return null;
  }
}

function findDuplicateQuestionInSheet_(sheet, sheetName, normalizedQuestion, normalizedEvent, normalizedLink) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const isAnswered = sheetName === SHEET_ANSWERED;
  const C = isAnswered ? A_COL : Q_COL;
  const width = Math.max(C.QUESTION, C.EVENT, C.LINK, C.STATUS, C.TICKET_ID, C.ASKED_BY, C.CREATED);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();

  // Search newest first so, if historical duplicates already exist, the user
  // is directed to the latest matching ticket while no new row is created.
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (normalizeQuestionDuplicateText(row[C.QUESTION - 1]) !== normalizedQuestion) continue;
    if (normalizeQuestionDuplicateText(row[C.EVENT - 1]) !== normalizedEvent) continue;
    if (normalizeQuestionDuplicateLink(row[C.LINK - 1]) !== normalizedLink) continue;
    const created = row[C.CREATED - 1];
    return {
      ticketId: String(row[C.TICKET_ID - 1] || ''),
      sheet: sheetName,
      status: isAnswered ? 'Answered' : String(row[C.STATUS - 1] || STATUS_OPEN),
      askedBy: String(row[C.ASKED_BY - 1] || ''),
      createdAt: created instanceof Date && !isNaN(created.getTime()) ? created.toISOString() : String(created || '')
    };
  }
  return null;
}

function findExistingDuplicateQuestion(question, eventName, caseLink, spreadsheet) {
  const normalizedQuestion = normalizeQuestionDuplicateText(question);
  const normalizedEvent = normalizeQuestionDuplicateText(eventName);
  const normalizedLink = normalizeQuestionDuplicateLink(caseLink);
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

  const cached = readVerifiedCachedDuplicateQuestion_(ss, normalizedQuestion, normalizedEvent, normalizedLink);
  if (cached) return cached;

  // Open matches take precedence because that is the ticket the team should
  // continue working. If none exists, also prevent recreating an answered
  // ticket with identical Question + Event + Case details.
  const match = findDuplicateQuestionInSheet_(
    ss.getSheetByName(SHEET_QUESTIONS), SHEET_QUESTIONS,
    normalizedQuestion, normalizedEvent, normalizedLink
  ) || findDuplicateQuestionInSheet_(
    ss.getSheetByName(SHEET_ANSWERED), SHEET_ANSWERED,
    normalizedQuestion, normalizedEvent, normalizedLink
  );
  if (match) cacheDuplicateQuestionMatch_(normalizedQuestion, normalizedEvent, normalizedLink, match);
  return match;
}

// Backward-compatible boolean helper retained for any older tests or internal
// calls. The new rule is global and does not depend on user or elapsed time.
function isDuplicateSubmission(sheet, question, eventName, askedByName, askedByEmail, caseLink) {
  const sheetName = sheet && sheet.getName ? sheet.getName() : SHEET_QUESTIONS;
  return !!findDuplicateQuestionInSheet_(
    sheet, sheetName,
    normalizeQuestionDuplicateText(question),
    normalizeQuestionDuplicateText(eventName),
    normalizeQuestionDuplicateLink(caseLink)
  );
}

// ==========================================
// MUTATIONS (all wrapped in withLock)
// ==========================================

function submitQuestion(payload) {
  // Validate and resolve identity before taking the shared write lock. The
  // lock is reserved for the duplicate check and Questions row append, which
  // keeps concurrent submissions moving while preserving one-at-a-time writes.
  ensureSheetsExist();
  payload = payload || {};

  // Highest-priority attention is an operational Support decision made only
  // after a question exists in Supervisor Desk. Reject stale or handcrafted
  // submission payloads so Coordinators cannot bypass that workflow.
  if (payload.needsAttentionToday) {
    throw new Error('Submit the question normally. Support can raise it to Attention Today from Supervisor Desk.');
  }

  const member = requireAuthenticatedMember(payload.askedByEmail);
  const question = String(payload.question || '').trim();
  const eventName = String(payload.eventName || '').trim();
  const caseLink = String(payload.caseLink || '').trim();
  const requestedAssignedToEmail = String(payload.assignedToEmail || '').trim().toLowerCase();
  const talentRole = normalizeEventRole(payload.talentRole);
  // Questions may only be owned by an active Support profile. Admin can
  // administer the app but can never be the question assignee.
  const automaticAssignee = requestedAssignedToEmail ? null : resolveAutomaticQuestionAssignee(member);
  const selectedAssignee = requestedAssignedToEmail ? requireQuestionAssignee(requestedAssignedToEmail) : automaticAssignee;
  const assignedToEmail = selectedAssignee ? selectedAssignee.email : '';

  if (!question) throw new Error("Please describe your question before submitting.");
  requireSheetCellLength(question, 'The question');
  rejectEmbeddedBase64Image(question, 'The question');
  if (!isValidEventName(eventName)) throw new Error(EVENT_NAME_INVALID_MSG);
  if (!isValidCaseLink(caseLink)) throw new Error("Please provide a valid Case / Event URL.");

  const now = new Date();
  const ticketId = Utilities.getUuid();
  const eventThisWeekend = !!payload.isEventThisWeekend;
  const eventDate = normalizeEventDate(payload.eventDate);
  if (eventThisWeekend && !eventDate) throw new Error('Please add the event date when selecting Event this weekend.');
  if (eventThisWeekend && !isEventDateCurrentOrFuture(eventDate)) throw new Error('Event this weekend must use today or a future event date.');
  const workCategory = inferWorkCategory(question, eventName);

  const result = withLock(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = findExistingDuplicateQuestion(question, eventName, caseLink, ss);
    if (existing) {
      return {
        success: false,
        duplicate: true,
        message: 'This ticket already exists and was not submitted again.',
        existingTicket: existing,
        _skipDataVersion: true
      };
    }

    const sheet = ss.getSheetByName(SHEET_QUESTIONS);
    sheet.appendRow([
      question,
      eventName,
      member.name,
      now,
      "",
      "", // Legacy priority column intentionally unused.
      "",
      caseLink,
      "",
      STATUS_OPEN,
      assignedToEmail,
      ticketId,
      "",
      "",
      0,
      member.email,
      false,
      "",
      now,
      member.name,
      workCategory,
      eventThisWeekend,
      eventDate,
      talentRole,
      assignedToEmail,
      false,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      false,
      '',
      '',
      '',
      '',
      ''
    ]);
    cacheDuplicateQuestionMatch_(
      normalizeQuestionDuplicateText(question),
      normalizeQuestionDuplicateText(eventName),
      normalizeQuestionDuplicateLink(caseLink),
      { ticketId: ticketId, sheet: SHEET_QUESTIONS, status: STATUS_OPEN, askedBy: member.name, createdAt: now.toISOString() }
    );
    return {
      success: true, ticketId: ticketId, assignedToEmail: assignedToEmail,
      automaticAssignment: !requestedAssignedToEmail && !!assignedToEmail,
      needsAttentionToday: false,
      assignmentEmailNotification: selectedAssignee && !emailsRepresentSameWorkspaceIdentity(selectedAssignee.email, member.email) ? {
        recipientEmail: selectedAssignee.email, recipientName: selectedAssignee.name,
        eventName: eventName, preview: question, actorName: member.name, actorEmail: member.email,
        sourceLabel: requestedAssignedToEmail ? 'question assignment' : 'automatically routed question', ticketId: ticketId
      } : null
    };
  }, { operation: 'submitQuestion' });

  if (!result || result.duplicate) return result;

  // Audit and notification work is intentionally outside the ticket write
  // lock. Either may fail independently without changing a successful ticket
  // save into an error that encourages the user to submit again.
  logAudit('SUBMIT', member.email, member.name, ticketId, {
    event: eventName,
    eventThisWeekend: eventThisWeekend,
    eventDate: eventDate,
    category: workCategory,
    talentRole: talentRole,
    assignedTo: assignedToEmail,
    routing: requestedAssignedToEmail ? 'Manual' : (assignedToEmail ? 'Automatic' : 'Unassigned')
  });

  if (selectedAssignee) {
    try {
      createDirectTicketNotification(
        ticketId, eventName, question, selectedAssignee, member,
        requestedAssignedToEmail ? 'Question Assignment' : 'Automatic Question Routing'
      );
    } catch (notificationError) {
      Logger.log('Question routing notification failed for ' + ticketId + ': ' + notificationError);
    }
  }

  if (result && result.assignmentEmailNotification) {
    const n = result.assignmentEmailNotification;
    sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail, n.sourceLabel, n.ticketId);
    delete result.assignmentEmailNotification;
  }
  return result;
}

function updateOpenQuestionCategory(ticketId, category, requestingEmail) {
  return withLock(() => {
    const member = requireAuthenticatedSupport(requestingEmail);
    const target = requireQuestionRow(ticketId);
    const normalized = normalizeWorkCategory(category);
    target.sheet.getRange(target.rowIndex, Q_COL.WORK_CATEGORY).setValue(normalized);
    logAudit('QUESTION_CATEGORY_UPDATE', member.email, member.name, ticketId, { category: normalized });
    return { success: true, category: normalized };
  });
}

function answerQuestion(ticketId, answerText, supervisorEmail, supervisorName, followUpTask) {
  const result = withLock(() => {
    const supervisor = requireAuthenticatedSupport(supervisorEmail);
    const answer = String(answerText || '').trim();
    if (!answer) throw new Error('Please write an answer before submitting.');
    requireSheetCellLength(answer, 'The answer');
    rejectEmbeddedBase64Image(answer, 'The answer');

    const target = requireQuestionRow(ticketId);
    const qSheet = target.sheet;
    const rowIndex = target.rowIndex;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aSheet = ss.getSheetByName(SHEET_ANSWERED);
    const taskSheet = ensureTasksSheet(ss);
    const rowData = qSheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];

    const env = followUpTask || {};
    const requestedTasks = Array.isArray(env)
      ? env
      : (Array.isArray(env.tasks) ? env.tasks : (env.required ? [env] : []));
    const selectedCategory = normalizeWorkCategory(
      env.questionCategory ||
      (requestedTasks[0] && requestedTasks[0].questionCategory) ||
      rowData[Q_COL.WORK_CATEGORY - 1] ||
      inferWorkCategory(rowData[Q_COL.QUESTION - 1], rowData[Q_COL.EVENT - 1])
    );

    const eventThisWeekend = rowData[Q_COL.EVENT_THIS_WEEKEND - 1] === true;
    const now = new Date();
    const createdDate = rowData[Q_COL.CREATED - 1] ? new Date(rowData[Q_COL.CREATED - 1]) : now;
    const askerEmail = normalizeEmail(rowData[Q_COL.ASKED_BY_EMAIL - 1]);

    const taskInfos = requestedTasks.map((raw, i) => {
      raw = raw || {};
      const requestedAssignee = normalizeEmail(raw.assignedToEmail) || askerEmail;
      if (!requestedAssignee) throw new Error('Task ' + (i + 1) + ' needs an assignee.');
      const assignee = requireTaskAssignee(requestedAssignee);
      const title = String(raw.title || '').trim();
      const instructions = String(raw.instructions || '').trim();
      if (!title) throw new Error('Task ' + (i + 1) + ' needs a title.');
      if (title.length > 140) throw new Error('Task ' + (i + 1) + ' title is too long.');
      if (!stripHtmlToText(instructions)) throw new Error('Task ' + (i + 1) + ' needs instructions.');
      requireSheetCellLength(instructions, 'Task ' + (i + 1) + ' instructions');
      rejectEmbeddedBase64Image(instructions, 'Task ' + (i + 1) + ' instructions');

      const relatedClientTalent = Object.prototype.hasOwnProperty.call(raw, 'relatedClientTalent')
        ? String(raw.relatedClientTalent || '').trim()
        : String(rowData[Q_COL.EVENT - 1] || '').trim();
      const relatedLink = Object.prototype.hasOwnProperty.call(raw, 'relatedLink')
        ? String(raw.relatedLink || '').trim()
        : String(rowData[Q_COL.LINK - 1] || '').trim();

      return {
        taskId: Utilities.getUuid(),
        title: title,
        instructions: instructions,
        assignee: assignee,
        dueDate: normalizeTaskDueDate(raw.dueDate) || getNextBusinessTaskDueDate(),
        priority: '', // Legacy priority column retained only for sheet compatibility.
        category: normalizeWorkCategory(raw.category || selectedCategory),
        relatedClientTalent: relatedClientTalent,
        relatedLink: relatedLink,
        needAttentionToday: raw.needAttentionToday === true || raw.attentionToday === true,
        supportOwnerEmail: resolveTaskSupportOwnerEmail(assignee)
      };
    });

    let cumulativeHold = Number(rowData[Q_COL.HOLD_ACCUM - 1]) || 0;
    const holdSinceValue = rowData[Q_COL.HOLD_SINCE - 1];
    if (holdSinceValue) {
      const d = holdSinceValue instanceof Date ? holdSinceValue : new Date(holdSinceValue);
      if (!isNaN(d.getTime())) cumulativeHold += Math.max(0, (now - d) / (1000 * 60 * 60));
    }
    const turnaroundHours = Math.max(0.1, Math.max(0, (now - createdDate) / (1000 * 60 * 60)) - cumulativeHold).toFixed(1);

    const appendedTaskIds = [];
    let answerAppended = false;
    try {
      taskInfos.forEach(t => {
        taskSheet.appendRow([
          t.taskId, t.title, t.instructions, t.assignee.name, t.assignee.email,
          supervisor.name, supervisor.email, now, t.dueDate, t.priority,
          TASK_STATUS_PENDING, '', '', 'Answer Follow-up', ticketId,
          t.relatedClientTalent, now, supervisor.name, t.category,
          t.relatedClientTalent, t.relatedLink, 'Question', ticketId,
          t.supportOwnerEmail, '', '', '',
          t.needAttentionToday === true, t.needAttentionToday ? now : '', t.needAttentionToday ? supervisor.name : '', t.needAttentionToday ? todayCSTDateString() : ''
        ]);
        appendedTaskIds.push(t.taskId);
      });

      aSheet.appendRow([
        rowData[Q_COL.QUESTION - 1], rowData[Q_COL.EVENT - 1], rowData[Q_COL.ASKED_BY - 1],
        rowData[Q_COL.CREATED - 1], now, '', turnaroundHours, rowData[Q_COL.LINK - 1],
        answer, STATUS_ANSWERED, supervisor.name, '', rowData[Q_COL.TICKET_ID - 1],
        cumulativeHold.toFixed(2), rowData[Q_COL.ASKED_BY_EMAIL - 1] || '', '', now,
        selectedCategory, eventThisWeekend, normalizeEventDate(rowData[Q_COL.EVENT_DATE - 1]),
        String(rowData[Q_COL.TALENT_ROLE - 1] || '').trim(),
        normalizeEmail(rowData[Q_COL.ROUTED_SUPPORT_EMAIL - 1] || rowData[Q_COL.ASSIGNED - 1]),
        supervisor.email,
        attentionWasRequested(rowData[Q_COL.ATTENTION_TODAY - 1], rowData[Q_COL.ATTENTION_SET_AT - 1]),
        String(rowData[Q_COL.ATTENTION_REASON - 1] || '').trim(),
        rowData[Q_COL.ATTENTION_SET_AT - 1] || '',
        String(rowData[Q_COL.ATTENTION_SET_BY - 1] || '').trim(),
        normalizeEventDate(rowData[Q_COL.ATTENTION_UNTIL - 1]),
        rowData[Q_COL.ATTENTION_CLEARED_AT - 1] || (attentionWasRequested(rowData[Q_COL.ATTENTION_TODAY - 1], rowData[Q_COL.ATTENTION_SET_AT - 1]) ? now : ''),
        String(rowData[Q_COL.ATTENTION_CLEARED_BY - 1] || '').trim() || (attentionWasRequested(rowData[Q_COL.ATTENTION_TODAY - 1], rowData[Q_COL.ATTENTION_SET_AT - 1]) ? supervisor.name : ''),
        rowData[Q_COL.ATTENTION_FIRST_REVIEW_AT - 1] || (attentionWasRequested(rowData[Q_COL.ATTENTION_TODAY - 1], rowData[Q_COL.ATTENTION_SET_AT - 1]) ? now : ''),
        String(rowData[Q_COL.ATTENTION_FIRST_REVIEW_BY - 1] || '').trim() || (attentionWasRequested(rowData[Q_COL.ATTENTION_TODAY - 1], rowData[Q_COL.ATTENTION_SET_AT - 1]) ? supervisor.name : '')
      ]);
      answerAppended = true;
      qSheet.deleteRow(rowIndex);
    } catch (e) {
      if (answerAppended) {
        try {
          const r = findRowIndexByTicketId(aSheet, ticketId, A_COL.TICKET_ID);
          if (r !== -1) aSheet.deleteRow(r);
        } catch (rollbackAnswerErr) {}
      }
      appendedTaskIds.forEach(id => {
        try {
          const r = findTaskRowById(taskSheet, id);
          if (r !== -1) taskSheet.deleteRow(r);
        } catch (rollbackTaskErr) {}
      });
      throw e;
    }

    if (taskInfos.length) bumpTaskDataVersion();

    const tasks = taskInfos.map(t => ({
      taskId: t.taskId,
      title: t.title,
      status: TASK_STATUS_PENDING,
      assignedTo: t.assignee.name,
      assignedToEmail: t.assignee.email,
      dueDate: t.dueDate,
      source: 'Answer Follow-up',
      relatedTicketId: ticketId,
      relatedEvent: t.relatedClientTalent,
      category: t.category,
      relatedClientTalent: t.relatedClientTalent,
      relatedLink: t.relatedLink,
      attentionToday: t.needAttentionToday === true,
      parentType: 'Question',
      parentId: ticketId
    }));

    return {
      success: true,
      task: tasks[0] || null,
      tasks: tasks,
      // Notification and audit sheet writes are post-commit work. Keeping
      // them outside the ScriptLock materially shortens the critical section
      // without weakening the Questions -> Answered + related Tasks
      // transaction above.
      _postCommit: {
        ticketId: ticketId,
        eventName: String(rowData[Q_COL.EVENT - 1] || ''),
        answer: answer,
        supervisor: supervisor,
        askerEmail: rowData[Q_COL.ASKED_BY_EMAIL - 1] || '',
        turnaroundHours: turnaroundHours,
        taskInfos: taskInfos,
        tasks: tasks,
        selectedCategory: selectedCategory,
        attentionRequested: attentionWasRequested(rowData[Q_COL.ATTENTION_TODAY - 1], rowData[Q_COL.ATTENTION_SET_AT - 1]),
        attentionUntil: normalizeEventDate(rowData[Q_COL.ATTENTION_UNTIL - 1])
      },
      emailNotification: {
        askedByEmail: rowData[Q_COL.ASKED_BY_EMAIL - 1] || '',
        askedByName: rowData[Q_COL.ASKED_BY - 1] || '',
        eventName: rowData[Q_COL.EVENT - 1] || '',
        answerText: answer,
        answeredByName: supervisor.name,
        answeredByEmail: supervisor.email
      },
      taskEmailNotifications: tasks
        .filter(t => normalizeEmail(t.assignedToEmail) !== normalizeEmail(supervisor.email))
        .map(t => ({
          recipientEmail: t.assignedToEmail, recipientName: t.assignedTo,
          eventType: 'assigned', actorName: supervisor.name, actorEmail: supervisor.email, task: t
        }))
    };
  }, { operation: 'answerQuestion' });

  if (result && result._postCommit) {
    const post = result._postCommit;
    let mentions = [];
    try {
      mentions = createMentionNotifications(
        post.ticketId, post.eventName, post.answer, post.supervisor,
        'Resolved Answer', [post.askerEmail]
      );
    } catch (notificationError) {
      Logger.log('Resolved-answer mention persistence failed for ' + post.ticketId + ': ' + notificationError);
    }

    logAudit('ANSWER', post.supervisor.email, post.supervisor.name, post.ticketId, {
      turnaroundHours: post.turnaroundHours,
      followUpTaskIds: post.taskInfos.map(t => t.taskId),
      followUpTaskCount: post.taskInfos.length,
      followUpRequired: post.taskInfos.length > 0,
      mentionedUsers: mentions.map(m => m.email),
      category: post.selectedCategory,
      attentionRequested: post.attentionRequested,
      attentionUntil: post.attentionUntil
    });

    post.taskInfos.forEach(t => logAudit('TASK_CREATE_FROM_ANSWER', post.supervisor.email, post.supervisor.name, t.taskId, {
      assignedTo: t.assignee.email,
      dueDate: t.dueDate,
      relatedTicketId: post.ticketId,
      parentType: 'Question',
      parentId: post.ticketId,
      category: t.category,
      relatedClientTalent: t.relatedClientTalent,
      entity: 'Task'
    }));

    post.tasks.forEach(t => safeCreateTaskNotification(
      t, t.assignedToEmail, t.assignedTo, 'assigned', post.supervisor.name, post.supervisor.email,
      'assigned:' + t.taskId + ':' + normalizeEmail(t.assignedToEmail)
    ));

    result.mentionEmailNotifications = mentions.map(m => ({
      recipientEmail: m.email, recipientName: m.name,
      eventName: post.eventName, preview: post.answer,
      actorName: post.supervisor.name, actorEmail: post.supervisor.email
    }));
    delete result._postCommit;
  }

  if (result && result.emailNotification) {
    const n = result.emailNotification;
    sendAnswerEmailIfEnabled(n.askedByEmail, n.askedByName, n.eventName, n.answerText, n.answeredByName, n.answeredByEmail);
    delete result.emailNotification;
  }
  if (result && Array.isArray(result.mentionEmailNotifications)) {
    result.mentionEmailNotifications.forEach(n => sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail));
    delete result.mentionEmailNotifications;
  }
  if (result && Array.isArray(result.taskEmailNotifications)) {
    result.taskEmailNotifications.forEach(n => sendTaskEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventType, n.task, n.actorName, n.actorEmail));
    delete result.taskEmailNotifications;
  }
  return result;
}

function holdQuestion(ticketId, reason, requestingEmail) {
  return withLock(() => {
    const member = requireAuthenticatedSupport(requestingEmail);
    const target = requireQuestionRow(ticketId);
    const sheet = target.sheet;
    const rowIndex = target.rowIndex;

    const current = sheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const alreadyOnHold = !!current[Q_COL.HOLD_SINCE - 1];
    const cleanReason = String(reason || '').trim();

    const stamp = formatCSTStamp(new Date());
    const entry = '[' + (alreadyOnHold ? 'Hold reason updated' : 'Put on hold') + ' by ' + member.name + ' · ' + stamp + ' CST]: ' + cleanReason;
    const existingReason = String(current[Q_COL.HOLD_REASON - 1] || '').trim();
    const newReasonValue = (alreadyOnHold && existingReason) ? existingReason + '\n\n' + entry : entry;

    sheet.getRange(rowIndex, Q_COL.HOLD_REASON).setValue(newReasonValue);
    sheet.getRange(rowIndex, Q_COL.FOLLOWUP).setValue(false);
    if (!alreadyOnHold) {
      sheet.getRange(rowIndex, Q_COL.STATUS).setValue(STATUS_HOLD);
      sheet.getRange(rowIndex, Q_COL.HOLD_SINCE).setValue(new Date());
      sheet.getRange(rowIndex, Q_COL.STATUS_CHANGED_AT).setValue(new Date());
      sheet.getRange(rowIndex, Q_COL.STATUS_CHANGED_BY).setValue(member.name);
    }
    logAudit(alreadyOnHold ? 'HOLD_REASON_UPDATED' : 'HOLD', member.email, member.name, ticketId, { reason: cleanReason });
    if (!alreadyOnHold) {
      const requester = operationalMemberForTicketRequester_(current[Q_COL.ASKED_BY_EMAIL - 1], current[Q_COL.ASKED_BY - 1]);
      if (requester) {
        try { createDirectTicketNotification(ticketId, current[Q_COL.EVENT - 1], cleanReason || current[Q_COL.QUESTION - 1], requester, member, 'Ticket On Hold'); }
        catch (notificationError) { Logger.log('Ticket hold notification failed for ' + ticketId + ': ' + notificationError); }
      }
    }
    return { success: true };
  });
}

function resumeQuestion(ticketId, requestingEmail) {
  return withLock(() => {
    const member = requireAuthenticatedSupport(requestingEmail);
    const target = requireQuestionRow(ticketId);
    const sheet = target.sheet;
    const rowIndex = target.rowIndex;

    const current = sheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const holdSinceVal = current[Q_COL.HOLD_SINCE - 1];
    if (!holdSinceVal) return { success: true };

    const holdSince = (holdSinceVal instanceof Date) ? holdSinceVal : new Date(holdSinceVal);
    const now = new Date();
    const elapsedHours = isNaN(holdSince.getTime()) ? 0 : Math.max(0, (now - holdSince) / (1000 * 60 * 60));
    const newCumulative = (Number(current[Q_COL.HOLD_ACCUM - 1]) || 0) + elapsedHours;

    sheet.getRange(rowIndex, Q_COL.STATUS).setValue(STATUS_OPEN);
    sheet.getRange(rowIndex, Q_COL.HOLD_REASON).setValue("");
    sheet.getRange(rowIndex, Q_COL.HOLD_SINCE).setValue("");
    sheet.getRange(rowIndex, Q_COL.HOLD_ACCUM).setValue(newCumulative);
    sheet.getRange(rowIndex, Q_COL.FOLLOWUP).setValue(false);
    sheet.getRange(rowIndex, Q_COL.STATUS_CHANGED_AT).setValue(new Date());
    sheet.getRange(rowIndex, Q_COL.STATUS_CHANGED_BY).setValue(member.name);
    logAudit('RESUME', member.email, member.name, ticketId, { holdHoursAdded: elapsedHours.toFixed(2) });
    const requester = operationalMemberForTicketRequester_(current[Q_COL.ASKED_BY_EMAIL - 1], current[Q_COL.ASKED_BY - 1]);
    if (requester) {
      try { createDirectTicketNotification(ticketId, current[Q_COL.EVENT - 1], current[Q_COL.QUESTION - 1], requester, member, 'Ticket Resumed From Hold'); }
      catch (notificationError) { Logger.log('Ticket resume notification failed for ' + ticketId + ': ' + notificationError); }
    }
    return { success: true };
  });
}

function addQuestionUpdate(ticketId, updateText, requestingEmail) {
  return withLock(() => {
    const member = requireAuthenticatedMember(requestingEmail);
    const text = String(updateText || '').trim();
    if (!text) throw new Error("Please enter some update text.");

    const target = requireQuestionRow(ticketId);
    const sheet = target.sheet;
    const rowIndex = target.rowIndex;

    const current = sheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const askedByName = String(current[Q_COL.ASKED_BY - 1] || '').trim();
    const askedByEmail = String(current[Q_COL.ASKED_BY_EMAIL - 1] || '').trim().toLowerCase();
    const reqEmail = normalizeEmail(requestingEmail);

    const isOwner = askedByEmail
      ? emailsRepresentSameWorkspaceIdentity(askedByEmail, reqEmail)
      : stripNoraPrefix(askedByName).toLowerCase() === stripNoraPrefix(String(member.name || '').trim()).toLowerCase();

    if (!isOwner && !isSupportMember(member)) {
      throw new Error("Access denied: you can only add updates to your own tickets.");
    }

    const stamp = formatCSTStamp(new Date());
    const newQuestion = String(current[Q_COL.QUESTION - 1] || '') + "\n\n[Update " + stamp + " CST by " + member.name + "]: " + text;
    sheet.getRange(rowIndex, Q_COL.QUESTION).setValue(newQuestion);

    const wasOnHold = String(current[Q_COL.STATUS - 1] || '').trim() === STATUS_HOLD;
    if (wasOnHold) {
      const holdSinceVal = current[Q_COL.HOLD_SINCE - 1];
      const holdSince = holdSinceVal instanceof Date ? holdSinceVal : new Date(holdSinceVal);
      const holdHours = holdSinceVal && !isNaN(holdSince.getTime())
        ? Math.max(0, (new Date() - holdSince) / (1000 * 60 * 60))
        : 0;
      const cumulativeHold = (Number(current[Q_COL.HOLD_ACCUM - 1]) || 0) + holdHours;

      sheet.getRange(rowIndex, Q_COL.STATUS).setValue(STATUS_OPEN);
      sheet.getRange(rowIndex, Q_COL.HOLD_REASON).setValue('');
      sheet.getRange(rowIndex, Q_COL.HOLD_SINCE).setValue('');
      sheet.getRange(rowIndex, Q_COL.HOLD_ACCUM).setValue(cumulativeHold);
      sheet.getRange(rowIndex, Q_COL.FOLLOWUP).setValue(true);
      sheet.getRange(rowIndex, Q_COL.STATUS_CHANGED_AT).setValue(new Date());
      sheet.getRange(rowIndex, Q_COL.STATUS_CHANGED_BY).setValue(member.name);
    }

    logAudit(wasOnHold ? 'ADD_UPDATE_FOLLOWUP' : 'ADD_UPDATE', member.email, member.name, ticketId, { textLength: text.length });
    return { success: true, followUp: wasOnHold };
  });
}


function markAnsweredTicketChangedForAsker(sheet, rowIndex, row, now) {
  const askerEmail = normalizeEmail(row[A_COL.ASKED_BY_EMAIL - 1]);
  if (askerEmail) {
    const readList = parseReadByList(row[A_COL.READ_BY - 1]).filter(e => e !== '*' && normalizeEmail(e) !== askerEmail);
    sheet.getRange(rowIndex, A_COL.READ_BY).setValue(readList.join(','));
  }
  sheet.getRange(rowIndex, A_COL.LAST_UPDATED_AT).setValue(now || new Date());
}

function sendTicketChangeEmailIfEnabled(recipientEmail, recipientName, eventName, changeLabel, detailText, actorName, actorEmail) {
  // Email is assignment-only. Ticket notes, answers, reopen/status changes,
  // and other informational updates remain available in Support Hub.
  return;
}


function addAnsweredTicketUpdate(ticketId, updateText, requestingEmail) {
  const result = withLock(() => {
    const member = requireAuthenticatedSupport(requestingEmail);
    const text = String(updateText || '').trim();
    if (!text) throw new Error('Please write an update before saving.');
    rejectEmbeddedBase64Image(text, 'The additional answer');
    const target = requireAnsweredRow(ticketId);
    const sheet = target.sheet;
    const rowIndex = target.rowIndex;
    const row = sheet.getRange(rowIndex, 1, 1, A_WIDTH).getValues()[0];
    const now = new Date();
    const stamp = formatCSTStamp(now);
    const previous = String(row[A_COL.ANSWER - 1] || '').trim();
    const entry = '[Additional Answer ' + stamp + ' CST by ' + member.name + ']: ' + text;
    const combinedAnswer = previous ? previous + '\n\n' + entry : entry;
    requireSheetCellLength(combinedAnswer, 'The answer history');
    sheet.getRange(rowIndex, A_COL.ANSWER).setValue(combinedAnswer);
    markAnsweredTicketChangedForAsker(sheet, rowIndex, row, now);
    const mentionRecipients = createMentionNotifications(
      ticketId,
      String(row[A_COL.EVENT - 1] || ''),
      text,
      member,
      'Answered Ticket Update',
      [row[A_COL.ASKED_BY_EMAIL - 1] || '']
    );
    logAudit('ANSWER_UPDATE', member.email, member.name, ticketId, {
      textLength: text.length,
      statusUnchanged: true,
      mentionedUsers: mentionRecipients.map(m => m.email)
    });
    return {
      success: true,
      emailNotification: {
        askedByEmail: row[A_COL.ASKED_BY_EMAIL - 1] || '', askedByName: row[A_COL.ASKED_BY - 1] || '',
        eventName: row[A_COL.EVENT - 1] || '', changeLabel: 'Additional answer', detailText: text,
        actorName: member.name, actorEmail: member.email
      },
      mentionEmailNotifications: mentionRecipients.map(m => ({
        recipientEmail: m.email, recipientName: m.name,
        eventName: String(row[A_COL.EVENT - 1] || ''), preview: text,
        actorName: member.name, actorEmail: member.email
      }))
    };
  });
  if (result && result.emailNotification) {
    const n = result.emailNotification;
    sendTicketChangeEmailIfEnabled(n.askedByEmail, n.askedByName, n.eventName, n.changeLabel, n.detailText, n.actorName, n.actorEmail);
    delete result.emailNotification;
  }
  if (result && Array.isArray(result.mentionEmailNotifications)) {
    result.mentionEmailNotifications.forEach(n => {
      sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail);
    });
    delete result.mentionEmailNotifications;
  }
  return result;
}

// One browser action handles both answered-ticket update choices. Keeping the
// existing public functions intact preserves compatibility with older deployed
// clients while the current UI gets a single, explicit workflow.
function postAnsweredTicketUpdate(ticketId, updateText, needsAnotherLook, requestingEmail) {
  requireAuthenticatedSupport(requestingEmail);
  if (needsAnotherLook === true) {
    return reopenAnsweredTicket(ticketId, updateText, requestingEmail, STATUS_OPEN);
  }
  return addAnsweredTicketUpdate(ticketId, updateText, requestingEmail);
}

function createFollowUpTaskFromAnswered(ticketId, payload, requestingEmail) {
  const result = withLock(() => {
    const member = requireAuthenticatedSupport(requestingEmail);
    payload = payload || {};
    const target = requireAnsweredRow(ticketId);
    const aSheet = target.sheet;
    const rowIndex = target.rowIndex;
    const row = aSheet.getRange(rowIndex, 1, 1, A_WIDTH).getValues()[0];
    const requestedAssignee = payload.assignedToEmail || row[A_COL.ASKED_BY_EMAIL - 1];
    const assignee = requireTaskAssignee(requestedAssignee);
    const title = String(payload.title || '').trim();
    const instructions = String(payload.instructions || '').trim();
    const dueDate = normalizeTaskDueDate(payload.dueDate) || getNextBusinessTaskDueDate();
    const priority = ''; // Legacy priority column retained only for sheet compatibility.
    const taskCategory = normalizeWorkCategory(payload.category || row[A_COL.WORK_CATEGORY - 1]);
    const relatedClientTalent = Object.prototype.hasOwnProperty.call(payload, 'relatedClientTalent')
      ? String(payload.relatedClientTalent || '').trim()
      : String(row[A_COL.EVENT - 1] || '').trim();
    const relatedLink = Object.prototype.hasOwnProperty.call(payload, 'relatedLink')
      ? String(payload.relatedLink || '').trim()
      : String(row[A_COL.LINK - 1] || '').trim();
    const relatedEntityType = String(payload.relatedEntityType || (relatedClientTalent ? 'Event' : '')).trim();
    const attentionToday = payload.needAttentionToday === true || payload.attentionToday === true;
    const attentionUntil = attentionToday ? todayCSTDateString() : '';
    if (!title) throw new Error('Task title is required.');
    if (title.length > 140) throw new Error('Task title is too long. Please keep it under 140 characters.');
    if (!stripHtmlToText(instructions)) throw new Error('Task instructions are required.');
    requireSheetCellLength(instructions, 'Task instructions');
    rejectEmbeddedBase64Image(instructions, 'Task instructions');

    const taskSheet = ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
    const now = new Date();
    const taskId = Utilities.getUuid();
    const supportOwnerEmail = resolveTaskSupportOwnerEmail(assignee);
    taskSheet.appendRow([
      taskId, title, instructions, assignee.name, assignee.email,
      member.name, member.email, now, dueDate, priority,
      TASK_STATUS_PENDING, '', '', 'Answer Follow-up', ticketId,
      relatedClientTalent, now, member.name, taskCategory, relatedClientTalent, relatedLink, 'Question', ticketId,
      supportOwnerEmail, '', '', relatedEntityType,
      attentionToday, attentionToday ? now : '', attentionToday ? member.name : '', attentionUntil
    ]);
    bumpTaskDataVersion();
    markAnsweredTicketChangedForAsker(aSheet, rowIndex, row, now);
    logAudit('TASK_CREATE_FROM_ANSWERED', member.email, member.name, taskId, {
      assignedTo: assignee.email, relatedTicketId: ticketId,
      category: taskCategory, relatedClientTalent: relatedClientTalent,
      parentType: 'Question', parentId: ticketId, supportOwnerEmail: supportOwnerEmail, entity: 'Task'
    });
    const task = taskRowToObject(taskSheet.getRange(taskSheet.getLastRow(), 1, 1, T_WIDTH).getValues()[0]);
    safeCreateTaskNotification(
      task, assignee.email, assignee.name, 'assigned', member.name, member.email,
      'assigned:' + task.taskId + ':' + normalizeEmail(assignee.email)
    );
    return {
      success: true,
      task: task,
      taskEmailNotification: assignee.email !== member.email ? {
        recipientEmail: assignee.email, recipientName: assignee.name,
        eventType: 'assigned', actorName: member.name, actorEmail: member.email, task: task
      } : null,
      askerEmailNotification: {
        askedByEmail: row[A_COL.ASKED_BY_EMAIL - 1] || '',
        askedByName: row[A_COL.ASKED_BY - 1] || '',
        eventName: row[A_COL.EVENT - 1] || '',
        changeLabel: 'Follow-up task created', detailText: title,
        actorName: member.name, actorEmail: member.email
      }
    };
  });
  if (result && result.taskEmailNotification) {
    const n = result.taskEmailNotification;
    sendTaskEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventType, n.task, n.actorName, n.actorEmail);
    delete result.taskEmailNotification;
  }
  if (result && result.askerEmailNotification) {
    const n = result.askerEmailNotification;
    sendTicketChangeEmailIfEnabled(n.askedByEmail, n.askedByName, n.eventName, n.changeLabel, n.detailText, n.actorName, n.actorEmail);
    delete result.askerEmailNotification;
  }
  return result;
}

function reopenAnsweredTicket(ticketId, updateText, requestingEmail, requestedStatus) {
  const result = withLock(() => {
    const member = requireAuthenticatedMember(requestingEmail);
    const text = String(updateText || '').trim();
    if (!text) throw new Error("Please describe the follow-up before reopening this ticket.");

    const target = requireAnsweredRow(ticketId);
    const aSheet = target.sheet;
    const rowIndex = target.rowIndex;
    const current = aSheet.getRange(rowIndex, 1, 1, A_WIDTH).getValues()[0];

    const askedByEmail = String(current[A_COL.ASKED_BY_EMAIL - 1] || '').trim().toLowerCase();
    // Any authenticated, active Support Hub team member may reopen an answered
    // ticket. The action and note are stamped and audited below, so ownership
    // is visible without preventing cross-team follow-up.

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const qSheet = ss.getSheetByName(SHEET_QUESTIONS);

    const stamp = formatCSTStamp(new Date());
    const previousAnswer = String(current[A_COL.ANSWER - 1] || '').trim();
    const previousAnsweredBy = String(current[A_COL.ANSWERED_BY - 1] || 'Supervisor').trim();
    const combinedQuestion =
      String(current[A_COL.QUESTION - 1] || '').trim() +
      "\n\n[Previous Answer by " + previousAnsweredBy + "]: " + previousAnswer +
      "\n\n[Follow-Up " + stamp + " CST by " + member.name + "]: " + text;
    requireSheetCellLength(combinedQuestion, 'The reopened ticket history');
    const reopenHistoryEntry = '[Reopened as Follow-Up by ' + member.name + ' · ' + stamp + ' CST]: ' + text;
    requireSheetCellLength(reopenHistoryEntry, 'The follow-up note');

    const now = new Date();
    // Reopening moves the existing ticket back to Questions. Keep its original
    // ID so related tasks, notifications, reports, and cross-navigation remain
    // connected to the same ticket.
    const newTicketId = String(current[A_COL.TICKET_ID - 1] || ticketId || '').trim() || Utilities.getUuid();
    const reopenStatus = String(requestedStatus || STATUS_OPEN).trim() === STATUS_HOLD ? STATUS_HOLD : STATUS_OPEN;
    // A Coordinator follow-up returns to the Support person who actually
    // answered the ticket. For legacy rows without Answered By Email, resolve
    // the stored responder name. Only then fall back to the original route.
    let assignedSupport = null;
    const answeredByEmail = normalizeEmail(current[A_COL.ANSWERED_BY_EMAIL - 1]);
    try { assignedSupport = answeredByEmail ? requireQuestionAssignee(answeredByEmail) : null; } catch (e) { assignedSupport = null; }
    if (!assignedSupport) {
      const answeredByName = stripNoraPrefix(String(current[A_COL.ANSWERED_BY - 1] || '')).trim().toLowerCase();
      if (answeredByName) {
        assignedSupport = _getTeamMembersInternal().find(row =>
          String(row.status || '').trim().toLowerCase() === 'active' &&
          String(row.category || '').trim().toLowerCase() === 'support' &&
          stripNoraPrefix(String(row.name || '')).trim().toLowerCase() === answeredByName
        ) || null;
      }
    }

    let routedSupportEmail = normalizeEmail(current[A_COL.ROUTED_SUPPORT_EMAIL - 1]);
    try { routedSupportEmail = routedSupportEmail ? requireQuestionAssignee(routedSupportEmail).email : ''; } catch (e) { routedSupportEmail = ''; }
    if (!assignedSupport && routedSupportEmail) {
      try { assignedSupport = requireQuestionAssignee(routedSupportEmail); } catch (e) { assignedSupport = null; }
    }
    if (!assignedSupport) {
      const askerMember = findTeamMemberByEmail(askedByEmail);
      const fallbackRoute = resolveAutomaticQuestionAssignee(askerMember);
      assignedSupport = fallbackRoute || null;
    }
    const assignedSupportEmail = assignedSupport ? normalizeEmail(assignedSupport.email) : '';
    if (!routedSupportEmail) routedSupportEmail = assignedSupportEmail;

    let appendedOpenRow = false;
    try {
      qSheet.appendRow([
        combinedQuestion,
        current[A_COL.EVENT - 1],
        current[A_COL.ASKED_BY - 1],
        now,
        "",
        "",
        "",
        current[A_COL.LINK - 1],
        "",
        reopenStatus,
        assignedSupportEmail,
        newTicketId,
        reopenStatus === STATUS_HOLD ? text : "",
        reopenStatus === STATUS_HOLD ? now : "",
        0,
        askedByEmail || "",
        true,
        reopenHistoryEntry,
        now,
        member.name,
        normalizeWorkCategory(current[A_COL.WORK_CATEGORY - 1] || inferWorkCategory(current[A_COL.QUESTION - 1], current[A_COL.EVENT - 1])),
        current[A_COL.EVENT_THIS_WEEKEND - 1] === true,
        normalizeEventDate(current[A_COL.EVENT_DATE - 1]),
        String(current[A_COL.TALENT_ROLE - 1] || '').trim(),
        routedSupportEmail,
        false,
        String(current[A_COL.ATTENTION_REASON - 1] || '').trim(),
        current[A_COL.ATTENTION_SET_AT - 1] || '',
        String(current[A_COL.ATTENTION_SET_BY - 1] || '').trim(),
        normalizeEventDate(current[A_COL.ATTENTION_UNTIL - 1]),
        current[A_COL.ATTENTION_CLEARED_AT - 1] || (current[A_COL.ATTENTION_REQUESTED - 1] === true ? now : ''),
        String(current[A_COL.ATTENTION_CLEARED_BY - 1] || '').trim() || (current[A_COL.ATTENTION_REQUESTED - 1] === true ? member.name : ''),
        current[A_COL.ATTENTION_FIRST_REVIEW_AT - 1] || '',
        String(current[A_COL.ATTENTION_FIRST_REVIEW_BY - 1] || '').trim(),
        false,
        '',
        '',
        '',
        '',
        ''
      ]);
      appendedOpenRow = true;
      aSheet.deleteRow(rowIndex);
    } catch (e) {
      if (appendedOpenRow) {
        try {
          const rollbackRow = findRowIndexByTicketId(qSheet, newTicketId, Q_COL.TICKET_ID);
          if (rollbackRow !== -1) qSheet.deleteRow(rollbackRow);
        } catch (rollbackError) {}
      }
      throw new Error('The ticket could not be reopened safely. No changes were saved. ' + (e && e.message ? e.message : e));
    }

    if (assignedSupport) {
      try {
        createDirectTicketNotification(newTicketId, current[A_COL.EVENT - 1], text, assignedSupport, member, 'Reopened Follow-Up');
      } catch (notificationError) {
        Logger.log('Reopened follow-up notification failed for ' + newTicketId + ': ' + notificationError);
      }
    }

    logAudit('REOPEN', member.email, member.name, newTicketId, {
      previousTicketId: String(ticketId || ''), status: reopenStatus,
      assignedTo: assignedSupportEmail,
      returnedToPreviousAnswerer: !!answeredByEmail || !!current[A_COL.ANSWERED_BY - 1],
      followUpNote: text,
      historyEntry: reopenHistoryEntry
    });
    return {
      success: true,
      ticketId: newTicketId,
      assignedToEmail: assignedSupportEmail,
      assignedToName: assignedSupport ? assignedSupport.name : '',
      historyEntry: reopenHistoryEntry,
      assignmentEmailNotification: assignedSupport && !emailsRepresentSameWorkspaceIdentity(assignedSupport.email, member.email) ? {
        recipientEmail: assignedSupport.email,
        recipientName: assignedSupport.name,
        eventName: String(current[A_COL.EVENT - 1] || ''),
        preview: text,
        actorName: member.name,
        actorEmail: member.email,
        ticketId: newTicketId
      } : null
    };
  });
  if (result && result.assignmentEmailNotification) {
    const n = result.assignmentEmailNotification;
    sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail, 'question assignment', n.ticketId);
    delete result.assignmentEmailNotification;
  }
  return result;
}

// Coordinator-level escalation after a question has waited more than 24 wall-clock hours.
// This does not change status, assignment, or Support's separate Attention Today flag.
function requestSupportAttention(ticketId, reason, requestingEmail) {
  const result = withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    if (String(actor.status || '').trim().toLowerCase() !== 'active' || String(actor.category || '').trim().toLowerCase() !== 'coordinator') {
      throw new Error('Only Coordinator, Sr Coordinator, or Project Manager profiles can request Support attention from My Active Submitted Questions.');
    }
    const target = requireQuestionRow(ticketId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const askedByEmail = normalizeEmail(row[Q_COL.ASKED_BY_EMAIL - 1]);
    if (!emailsRepresentSameWorkspaceIdentity(askedByEmail, actor.email)) {
      throw new Error('You can only flag a question that you submitted.');
    }
    const created = row[Q_COL.CREATED - 1] instanceof Date ? row[Q_COL.CREATED - 1] : new Date(row[Q_COL.CREATED - 1]);
    if (!created || isNaN(created.getTime()) || (Date.now() - created.getTime()) <= 24 * 60 * 60 * 1000) {
      throw new Error('This question can be flagged after it has been waiting for more than 24 hours.');
    }
    const alreadyFlagged = row[Q_COL.REQUESTER_FLAGGED - 1] === true || !!row[Q_COL.REQUESTER_FLAGGED_AT - 1];
    if (alreadyFlagged) {
      const ackAt = row[Q_COL.REQUESTER_FLAG_ACK_AT - 1];
      return {
        success: true,
        alreadyFlagged: true,
        acknowledged: !!ackAt,
        requesterFlaggedAt: row[Q_COL.REQUESTER_FLAGGED_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAGGED_AT - 1]) : '',
        _skipDataVersion: true
      };
    }
    const cleanReason = String(reason || '').trim().slice(0, 500);
    const now = new Date();
    target.sheet.getRange(target.rowIndex, Q_COL.REQUESTER_FLAGGED, 1, 6).setValues([[
      true, now, actor.name, cleanReason, '', ''
    ]]);
    logAudit('REQUESTER_SUPPORT_FLAG_SET', actor.email, actor.name, ticketId, {
      reason: cleanReason,
      assignedTo: normalizeEmail(row[Q_COL.ASSIGNED - 1]),
      event: String(row[Q_COL.EVENT - 1] || '')
    });

    const recipients = [];
    const seen = {};
    const addRecipient = function(member) {
      if (!member || !normalizeEmail(member.email) || isAdminMember(member)) return;
      const email = normalizeEmail(member.email);
      if (seen[email] || emailsRepresentSameWorkspaceIdentity(email, actor.email)) return;
      seen[email] = true;
      recipients.push(member);
    };
    const assignedEmail = normalizeEmail(row[Q_COL.ASSIGNED - 1]);
    if (assignedEmail) {
      try { addRecipient(requireQuestionAssignee(assignedEmail)); } catch (e) {}
    }
    // Assistant Managers receive oversight notifications even when a Supervisor owns the ticket.
    _getTeamMembersInternal().forEach(function(member) {
      if (String(member.status || '').trim().toLowerCase() !== 'active') return;
      if (String(member.category || '').trim().toLowerCase() !== 'support') return;
      if (String(member.title || '').trim().toLowerCase() !== 'assistant manager') return;
      addRecipient(member);
    });
    // If a legacy/unassigned ticket has no owner, include its stored route so the flag is not stranded.
    if (!assignedEmail) {
      const routed = normalizeEmail(row[Q_COL.ROUTED_SUPPORT_EMAIL - 1]);
      if (routed) {
        try { addRecipient(requireQuestionAssignee(routed)); } catch (e) {}
      }
    }

    const preview = (cleanReason ? cleanReason + ': ' : '') + stripHtmlToText(row[Q_COL.QUESTION - 1]);
    recipients.forEach(function(recipient) {
      try {
        createDirectTicketNotification(
          ticketId,
          String(row[Q_COL.EVENT - 1] || ''),
          preview,
          recipient,
          actor,
          'Team Flag - Support Attention Requested'
        );
      } catch (notificationError) {
        Logger.log('Requester team-flag notification failed for ' + ticketId + ': ' + notificationError);
      }
    });
    return {
      success: true,
      requesterFlagActive: true,
      requesterFlaggedAt: now.toISOString(),
      requesterFlaggedBy: actor.name,
      requesterFlagReason: cleanReason,
      notified: recipients.map(function(m) { return normalizeEmail(m.email); })
    };
  }, { operation: 'requestSupportAttention' });
  return result;
}

function acknowledgeRequesterSupportFlag(ticketId, requestingEmail) {
  return withLock(() => {
    const actor = requireAuthenticatedSupport(requestingEmail);
    if (isAdminMember(actor)) throw new Error('Switch to an operational Support profile to acknowledge team flags.');
    const target = requireQuestionRow(ticketId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const wasFlagged = row[Q_COL.REQUESTER_FLAGGED - 1] === true || !!row[Q_COL.REQUESTER_FLAGGED_AT - 1];
    if (!wasFlagged) return { success: true, acknowledged: false, _skipDataVersion: true };
    if (row[Q_COL.REQUESTER_FLAG_ACK_AT - 1]) {
      return { success: true, acknowledged: true, alreadyAcknowledged: true, _skipDataVersion: true };
    }
    const now = new Date();
    target.sheet.getRange(target.rowIndex, Q_COL.REQUESTER_FLAG_ACK_AT, 1, 2).setValues([[now, actor.name]]);
    logAudit('REQUESTER_SUPPORT_FLAG_ACKNOWLEDGED', actor.email, actor.name, ticketId, {
      flaggedAt: row[Q_COL.REQUESTER_FLAGGED_AT - 1] ? safeIsoDate(row[Q_COL.REQUESTER_FLAGGED_AT - 1]) : '',
      flaggedBy: String(row[Q_COL.REQUESTER_FLAGGED_BY - 1] || ''),
      reason: String(row[Q_COL.REQUESTER_FLAG_REASON - 1] || '')
    });
    return { success: true, acknowledged: true, acknowledgedAt: now.toISOString(), acknowledgedBy: actor.name };
  }, { operation: 'acknowledgeRequesterSupportFlag' });
}

function markQuestionAttentionReviewed(ticketId, requestingEmail) {
  return withLock(() => {
    const actor = requireAuthenticatedSupport(requestingEmail);
    const target = requireQuestionRow(ticketId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    if (!isAttentionActive(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_UNTIL - 1]) || row[Q_COL.ATTENTION_FIRST_REVIEW_AT - 1]) {
      return { success: true, recorded: false, _skipDataVersion: true };
    }
    const now = new Date();
    target.sheet.getRange(target.rowIndex, Q_COL.ATTENTION_FIRST_REVIEW_AT, 1, 2).setValues([[now, actor.name]]);
    logAudit('ATTENTION_TODAY_FIRST_REVIEW', actor.email, actor.name, ticketId, { reviewedAt: now.toISOString() });
    return { success: true, recorded: true, reviewedAt: now.toISOString(), reviewedBy: actor.name };
  });
}

function setQuestionAttention(ticketId, enabled, reason, untilDate, requestingEmail, optionalAssigneeEmail) {
  const result = withLock(() => {
    const actor = requireAuthenticatedSupport(requestingEmail);
    const target = requireQuestionRow(ticketId);
    const sheet = target.sheet;
    const rowIndex = target.rowIndex;
    const row = sheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const turnOn = enabled === true;
    const now = new Date();

    if (turnOn) {
      const existingAssignedEmail = normalizeEmail(row[Q_COL.ASSIGNED - 1]);
      const requestedAssigneeEmail = normalizeEmail(optionalAssigneeEmail);
      // Assignment is optional. If Support chooses somebody in the final
      // attention step, validate and apply that owner atomically. Otherwise
      // preserve the current owner or keep the question unassigned.
      const assignee = requestedAssigneeEmail
        ? requireQuestionAssignee(requestedAssigneeEmail)
        : (existingAssignedEmail ? requireQuestionAssignee(existingAssignedEmail) : null);
      const effectiveAssignedEmail = assignee ? assignee.email : '';
      const cleanReason = cleanAttentionReason(reason);
      const normalizedUntil = normalizeAttentionUntil(untilDate);
      const alreadySame = isAttentionActive(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_UNTIL - 1])
        && String(row[Q_COL.ATTENTION_REASON - 1] || '').trim() === cleanReason
        && normalizeEventDate(row[Q_COL.ATTENTION_UNTIL - 1]) === normalizedUntil
        && existingAssignedEmail === effectiveAssignedEmail;
      if (alreadySame) return { success: true, attentionToday: true, _skipDataVersion: true };

      if (requestedAssigneeEmail && existingAssignedEmail !== effectiveAssignedEmail) {
        sheet.getRange(rowIndex, Q_COL.ASSIGNED).setValue(effectiveAssignedEmail);
      }
      sheet.getRange(rowIndex, Q_COL.ATTENTION_TODAY, 1, 9).setValues([[
        true, cleanReason, now, actor.name, normalizedUntil, '', '', '', ''
      ]]);
      logAudit('ATTENTION_TODAY_SET', actor.email, actor.name, ticketId, {
        reason: cleanReason, attentionUntil: normalizedUntil,
        assignedTo: effectiveAssignedEmail, assignmentOptional: true
      });
      if (assignee) {
        try {
          createDirectTicketNotification(
            ticketId,
            String(row[Q_COL.EVENT - 1] || ''),
            cleanReason + ': ' + stripHtmlToText(row[Q_COL.QUESTION - 1]),
            assignee,
            actor,
            'Needs Attention Today'
          );
        } catch (notificationError) {
          Logger.log('Attention Today notification failed for ' + ticketId + ': ' + notificationError);
        }
      }
      return {
        success: true,
        attentionToday: true,
        attentionUntil: normalizedUntil,
        assignedToEmail: effectiveAssignedEmail,
        emailNotification: requestedAssigneeEmail && existingAssignedEmail !== effectiveAssignedEmail && assignee && !emailsRepresentSameWorkspaceIdentity(assignee.email, actor.email) ? {
          recipientEmail: assignee.email,
          recipientName: assignee.name,
          eventName: String(row[Q_COL.EVENT - 1] || ''),
          preview: cleanReason + ': ' + stripHtmlToText(row[Q_COL.QUESTION - 1]),
          actorName: actor.name,
          actorEmail: actor.email,
          ticketId: ticketId
        } : null
      };
    }

    if (!attentionWasRequested(row[Q_COL.ATTENTION_TODAY - 1], row[Q_COL.ATTENTION_SET_AT - 1])) {
      return { success: true, attentionToday: false, _skipDataVersion: true };
    }
    sheet.getRange(rowIndex, Q_COL.ATTENTION_TODAY).setValue(false);
    sheet.getRange(rowIndex, Q_COL.ATTENTION_CLEARED_AT, 1, 2).setValues([[now, actor.name]]);
    logAudit('ATTENTION_TODAY_CLEARED', actor.email, actor.name, ticketId, {
      previousUntil: normalizeEventDate(row[Q_COL.ATTENTION_UNTIL - 1]),
      previousReason: String(row[Q_COL.ATTENTION_REASON - 1] || '').trim()
    });
    return { success: true, attentionToday: false };
  });

  if (result && result.emailNotification) {
    const n = result.emailNotification;
    sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail, 'question assignment', n.ticketId);
    delete result.emailNotification;
  }
  return result;
}

function assignQuestion(ticketId, assigneeEmail, requestingEmail, notes) {
  const result = withLock(() => {
    const requester = requireAuthenticatedSupport(requestingEmail);
    const assignee = requireQuestionAssignee(assigneeEmail);
    const target = requireQuestionRow(ticketId);
    const sheet = target.sheet;
    const rowIndex = target.rowIndex;
    const row = sheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const previousAssignee = normalizeEmail(row[Q_COL.ASSIGNED - 1]);

    sheet.getRange(rowIndex, Q_COL.ASSIGNED).setValue(assignee.email);
    // Remember the first owner so a later deliberate unassignment is not
    // mistaken for a never-routed legacy submission.
    if (!normalizeEmail(row[Q_COL.ROUTED_SUPPORT_EMAIL - 1])) {
      sheet.getRange(rowIndex, Q_COL.ROUTED_SUPPORT_EMAIL).setValue(assignee.email);
    }

    const note = String(notes || '').trim();
    if (note) {
      const stamp = formatCSTStamp(new Date());
      const entry = "[Assign to " + assignee.name + " · " + stamp + " CST by " + requester.name + "]: " + note;
      const currentNotes = String(sheet.getRange(rowIndex, Q_COL.ASSIGN_NOTES).getValue() || '').trim();
      sheet.getRange(rowIndex, Q_COL.ASSIGN_NOTES).setValue(currentNotes ? currentNotes + "\n\n" + entry : entry);
    }

    logAudit('ASSIGN', requester.email, requester.name, ticketId, { assignedTo: assignee.email, hasNotes: !!note });
    if (previousAssignee !== assignee.email) {
      try { createDirectTicketNotification(ticketId, row[Q_COL.EVENT - 1], row[Q_COL.QUESTION - 1], assignee, requester, 'Question Assignment'); }
      catch (notificationError) { Logger.log('Question assignment notification failed for ' + ticketId + ': ' + notificationError); }
    }
    return {
      success: true,
      assignmentEmailNotification: previousAssignee !== assignee.email && !emailsRepresentSameWorkspaceIdentity(assignee.email, requester.email) ? {
        recipientEmail: assignee.email, recipientName: assignee.name, eventName: String(row[Q_COL.EVENT - 1] || ''),
        preview: String(row[Q_COL.QUESTION - 1] || ''), actorName: requester.name, actorEmail: requester.email, ticketId: ticketId
      } : null
    };
  });
  if (result && result.assignmentEmailNotification) {
    const n = result.assignmentEmailNotification;
    sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail, 'question assignment', n.ticketId);
    delete result.assignmentEmailNotification;
  }
  return result;
}

function assignQuestionsBulk(ticketIds, assigneeEmail, requestingEmail, notes) {
  const result = withLock(() => {
    const requester = requireAuthenticatedSupport(requestingEmail);
    const assignee = requireQuestionAssignee(assigneeEmail);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
    let updated = 0;
    const assignmentEmailNotifications = [];
    const note = String(notes || '').trim();
    const stamp = formatCSTStamp(new Date());

    (ticketIds || []).forEach(id => {
      const rowIndex = findRowIndexByTicketId(qSheet, id, Q_COL.TICKET_ID);
      if (rowIndex !== -1) {
        const row = qSheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];
        const previousAssignee = normalizeEmail(row[Q_COL.ASSIGNED - 1]);
        qSheet.getRange(rowIndex, Q_COL.ASSIGNED).setValue(assignee.email);
        if (!normalizeEmail(row[Q_COL.ROUTED_SUPPORT_EMAIL - 1])) {
          qSheet.getRange(rowIndex, Q_COL.ROUTED_SUPPORT_EMAIL).setValue(assignee.email);
        }
        if (note) {
          const entry = "[Assign to " + assignee.name + " · " + stamp + " CST by " + requester.name + "]: " + note;
          const currentNotes = String(qSheet.getRange(rowIndex, Q_COL.ASSIGN_NOTES).getValue() || '').trim();
          qSheet.getRange(rowIndex, Q_COL.ASSIGN_NOTES).setValue(currentNotes ? currentNotes + "\n\n" + entry : entry);
        }
        if (previousAssignee !== assignee.email) {
          try { createDirectTicketNotification(id, row[Q_COL.EVENT - 1], row[Q_COL.QUESTION - 1], assignee, requester, 'Question Assignment'); }
          catch (notificationError) { Logger.log('Bulk question assignment notification failed for ' + id + ': ' + notificationError); }
          if (!emailsRepresentSameWorkspaceIdentity(assignee.email, requester.email)) {
            assignmentEmailNotifications.push({
              recipientEmail: assignee.email, recipientName: assignee.name, eventName: String(row[Q_COL.EVENT - 1] || ''),
              preview: String(row[Q_COL.QUESTION - 1] || ''), actorName: requester.name, actorEmail: requester.email, ticketId: id
            });
          }
        }
        updated++;
      }
    });

    logAudit('ASSIGN_BULK', requester.email, requester.name, '', { assignedTo: assignee.email, count: updated, hasNotes: !!note });
    return { success: true, updated: updated, assignmentEmailNotifications: assignmentEmailNotifications };
  });
  if (result && Array.isArray(result.assignmentEmailNotifications)) {
    result.assignmentEmailNotifications.forEach(n => sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail, 'question assignment', n.ticketId));
    delete result.assignmentEmailNotifications;
  }
  return result;
}

function unassignQuestion(ticketId, requestingEmail) {
  return withLock(() => {
    const requester = requireAuthenticatedSupport(requestingEmail);
    const target = requireQuestionRow(ticketId);
    const current = target.sheet.getRange(target.rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    const previousAssignee = normalizeEmail(current[Q_COL.ASSIGNED - 1]);
    if (previousAssignee && !normalizeEmail(current[Q_COL.ROUTED_SUPPORT_EMAIL - 1])) {
      target.sheet.getRange(target.rowIndex, Q_COL.ROUTED_SUPPORT_EMAIL).setValue(previousAssignee);
    }
    target.sheet.getRange(target.rowIndex, Q_COL.ASSIGNED).setValue("");
    logAudit('UNASSIGN', requester.email, requester.name, ticketId, { attentionTodayPreserved: true });
    return { success: true };
  });
}

function toggleReadStatus(ticketId, isRead, requestingEmail) {
  return withLock(() => {
    const member = requireAuthenticatedMember(requestingEmail);
    const email = normalizeEmail(requestingEmail);
    const target = requireAnsweredRow(ticketId);

    // Server-side isolation: Coordinators may only change read state on answers
    // that belong to them. Support/Admin can review any answer as themselves.
    if (!isSupportMember(member)) {
      const row = target.sheet.getRange(target.rowIndex, 1, 1, A_WIDTH).getValues()[0];
      const askerEmail = normalizeEmail(row[A_COL.ASKED_BY_EMAIL - 1]);
      const askerName = stripNoraPrefix(String(row[A_COL.ASKED_BY - 1] || '')).trim().toLowerCase();
      const isOwn = askerEmail
        ? askerEmail === email
        : askerName === stripNoraPrefix(String(member.name || '')).trim().toLowerCase();
      if (!isOwn) throw new Error("Access denied: you can only change read status on your own answered tickets.");
    }

    const readCell = target.sheet.getRange(target.rowIndex, A_COL.READ_BY);
    let list = parseReadByList(readCell.getValue()).filter(e => e !== '*');

    const firstReadCell = target.sheet.getRange(target.rowIndex, A_COL.FIRST_READ_AT);
    const firstReadMap = parseFirstReadMap(firstReadCell.getValue());
    let firstReadAt = firstReadMap[email] || '';

    if (isRead) {
      if (list.indexOf(email) === -1) list.push(email);
      if (!firstReadAt) {
        firstReadAt = new Date().toISOString();
        firstReadMap[email] = firstReadAt;
        firstReadCell.setValue(serializeFirstReadMap(firstReadMap));
        logAudit('FIRST_REVIEW', member.email, member.name, ticketId, { firstReadAt: firstReadAt });
      }
    } else {
      // Current state can go back to unread, but firstReadMap is intentionally
      // preserved so reporting keeps the original first-review speed.
      list = list.filter(e => e !== email);
    }

    readCell.setValue(list.join(', '));
    return { success: true, isRead: !!isRead, firstReadAt: firstReadAt };
  });
}

function markAllAnsweredRead(requestingEmail) {
  return withLock(() => {
    const member = requireAuthenticatedMember(requestingEmail);
    const email = normalizeEmail(requestingEmail);
    const aSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ANSWERED);
    const lastRow = aSheet.getLastRow();
    if (lastRow < 2) return { success: true, updated: 0, _skipDataVersion: true };

    const data = aSheet.getRange(2, 1, lastRow - 1, A_WIDTH).getValues();
    // "Mark All Read" is intentionally personal for everyone. Support users do
    // not bulk-mark other people's answered tickets as reviewed.
    let updated = 0;
    let firstReviewStamped = 0;
    const nowIso = new Date().toISOString();
    const newReadVals = [];
    const newFirstReadVals = [];

    for (let i = 0; i < data.length; i++) {
      const rowAskerEmail = String(data[i][A_COL.ASKED_BY_EMAIL - 1] || '').trim().toLowerCase();
      const rowAskerName = stripNoraPrefix(String(data[i][A_COL.ASKED_BY - 1] || '').trim()).toLowerCase();
      const isOwn = rowAskerEmail === email || (rowAskerName && rowAskerName === stripNoraPrefix(String(member.name)).trim().toLowerCase());
      let list = parseReadByList(data[i][A_COL.READ_BY - 1]).filter(e => e !== '*');
      const firstReadMap = parseFirstReadMap(data[i][A_COL.FIRST_READ_AT - 1]);

      if (isOwn && list.indexOf(email) === -1) {
        list.push(email);
        updated++;
        if (!firstReadMap[email]) {
          firstReadMap[email] = nowIso;
          firstReviewStamped++;
        }
      }
      newReadVals.push([list.join(', ')]);
      newFirstReadVals.push([serializeFirstReadMap(firstReadMap)]);
    }

    if (updated > 0) {
      aSheet.getRange(2, A_COL.READ_BY, lastRow - 1, 1).setValues(newReadVals);
      aSheet.getRange(2, A_COL.FIRST_READ_AT, lastRow - 1, 1).setValues(newFirstReadVals);
    }
    logAudit('MARK_ALL_READ', member.email, member.name, '', { updated: updated, firstReviewStamped: firstReviewStamped });
    return { success: true, updated: updated, firstReviewStamped: firstReviewStamped, _skipDataVersion: updated === 0 };
  });
}

// ==========================================
// PHASE 1 TASKS — independent from Questions / Answered
// ==========================================
function bumpTaskDataVersion() {
  try {
    const props = PropertiesService.getScriptProperties();
    const current = Number(props.getProperty(CONFIG_KEY_TASK_DATA_VERSION)) || 0;
    const next = String(current + 1);
    props.setProperty(CONFIG_KEY_TASK_DATA_VERSION, next);
    return next;
  } catch (versionError) {
    // The task Sheet mutation is the source of truth. A transient
    // PropertiesService failure after that write must not be surfaced as if
    // the task action failed, otherwise a retry can duplicate the operation.
    Logger.log('Task data version update failed after a successful write: ' + versionError);
    return '';
  }
}

function getTaskDataVersion() {
  return PropertiesService.getScriptProperties().getProperty(CONFIG_KEY_TASK_DATA_VERSION) || '0';
}

function normalizeTaskDueDate(value) {
  const clean = String(value || '').trim();
  if (!clean) return '';
  if (!isExactIsoCalendarDate(clean)) throw new Error('Due date is invalid. Please choose a date from the calendar.');
  return clean;
}

function getNextBusinessTaskDueDate() {
  const today = Utilities.formatDate(new Date(), APP_TIMEZONE, 'yyyy-MM-dd');
  const parts = today.split('-').map(Number);
  let d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 12, 0, 0));
  do {
    d.setUTCDate(d.getUTCDate() + 1);
  } while (d.getUTCDay() === 0 || d.getUTCDay() === 6);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function findTaskRowById(sheet, taskId) {
  const id = String(taskId || '').trim();
  if (!id) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, T_COL.TASK_ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === id) return i + 2;
  }
  return -1;
}

function requireTaskRow(taskId) {
  const sheet = ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
  const rowIndex = findTaskRowById(sheet, taskId);
  if (rowIndex === -1) throw new Error('This task no longer exists. Refresh to see its current state.');
  return { sheet: sheet, rowIndex: rowIndex };
}

function taskDueDateString(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, APP_TIMEZONE, 'yyyy-MM-dd');
  const clean = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(clean)) return clean;
  const d = new Date(clean);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, APP_TIMEZONE, 'yyyy-MM-dd');
}

function taskRowToObject(row) {
  return {
    taskId: String(row[T_COL.TASK_ID - 1] || '').trim(),
    title: String(row[T_COL.TITLE - 1] || '').trim(),
    instructions: String(row[T_COL.INSTRUCTIONS - 1] || '').trim(),
    assignedTo: stripNoraPrefix(String(row[T_COL.ASSIGNED_TO - 1] || '').trim()),
    assignedToEmail: normalizeEmail(row[T_COL.ASSIGNED_TO_EMAIL - 1]),
    createdBy: stripNoraPrefix(String(row[T_COL.CREATED_BY - 1] || '').trim()),
    createdByEmail: normalizeEmail(row[T_COL.CREATED_BY_EMAIL - 1]),
    createdAt: row[T_COL.CREATED_AT - 1] ? safeIsoDate(row[T_COL.CREATED_AT - 1]) : '',
    dueDate: taskDueDateString(row[T_COL.DUE_DATE - 1]),
    status: String(row[T_COL.STATUS - 1] || TASK_STATUS_PENDING).trim(),
    completedAt: row[T_COL.COMPLETED_AT - 1] ? safeIsoDate(row[T_COL.COMPLETED_AT - 1]) : '',
    completionNote: String(row[T_COL.COMPLETION_NOTE - 1] || '').trim(),
    source: String(row[T_COL.SOURCE - 1] || 'Manual').trim(),
    relatedTicketId: String(row[T_COL.RELATED_TICKET_ID - 1] || '').trim(),
    relatedEvent: String(row[T_COL.RELATED_EVENT - 1] || '').trim(),
    updatedAt: row[T_COL.UPDATED_AT - 1] ? safeIsoDate(row[T_COL.UPDATED_AT - 1]) : '',
    updatedBy: stripNoraPrefix(String(row[T_COL.UPDATED_BY - 1] || '').trim()),
    category: String(row[T_COL.WORK_CATEGORY - 1] || '').trim().toLowerCase() === 'none' ? 'None' : normalizeWorkCategory(row[T_COL.WORK_CATEGORY - 1]),
    relatedClientTalent: String(row[T_COL.RELATED_CLIENT_TALENT - 1] || '').trim(),
    relatedLink: String(row[T_COL.RELATED_LINK - 1] || '').trim(),
    relatedEntityType: String(row[T_COL.RELATED_ENTITY_TYPE - 1] || '').trim(),
    parentType: String(row[T_COL.PARENT_TYPE - 1] || '').trim(),
    parentId: String(row[T_COL.PARENT_ID - 1] || '').trim(),
    supportOwnerEmail: normalizeEmail(row[T_COL.SUPPORT_OWNER_EMAIL - 1]),
    completedByEmail: normalizeEmail(row[T_COL.COMPLETED_BY_EMAIL - 1]),
    completedBy: stripNoraPrefix(String(row[T_COL.COMPLETED_BY - 1] || '').trim()),
    attentionToday: isAttentionActive(row[T_COL.ATTENTION_TODAY - 1] === true, row[T_COL.ATTENTION_UNTIL - 1]),
    attentionSetAt: row[T_COL.ATTENTION_SET_AT - 1] ? safeIsoDate(row[T_COL.ATTENTION_SET_AT - 1]) : '',
    attentionSetBy: stripNoraPrefix(String(row[T_COL.ATTENTION_SET_BY - 1] || '').trim()),
    attentionUntil: normalizeEventDate(row[T_COL.ATTENTION_UNTIL - 1])
  };
}

function getTasksData(requestingEmail) {
  const member = requireAuthenticatedMember(requestingEmail);
  const requesterIsSupport = isSupportMember(member);
  const email = normalizeEmail(member.email);
  const sheet = ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, T_WIDTH).getValues();
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const task = taskRowToObject(data[i]);
    if (!task.taskId || !task.title) continue;
    if (!task.supportOwnerEmail && task.assignedToEmail) {
      const assignee = findTeamMemberByEmail(task.assignedToEmail);
      task.supportOwnerEmail = resolveTaskSupportOwnerEmail(assignee);
      task.supportOwnerDerived = !!task.supportOwnerEmail;
    }
    if (requesterIsSupport || task.assignedToEmail === email || task.createdByEmail === email) out.push(task);
  }
  return out;
}

function getTasksDataIfChanged(requestingEmail, clientVersion) {
  const startedAt = Date.now();
  requireAuthenticatedMember(requestingEmail);
  const authenticatedAt = Date.now();
  ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
  const version = getTaskDataVersion();
  if (clientVersion !== null && clientVersion !== undefined && String(clientVersion) === String(version)) {
    return { unchanged: true, version: version, serverTiming: { totalMs: Date.now() - startedAt, authMs: authenticatedAt - startedAt, dataMs: Date.now() - authenticatedAt } };
  }
  const data = getTasksData(requestingEmail);
  return { unchanged: false, version: version, data: data, serverTiming: { totalMs: Date.now() - startedAt, authMs: authenticatedAt - startedAt, dataMs: Date.now() - authenticatedAt } };
}

function isDuplicateTaskSubmission(sheet, title, instructions, assigneeEmail, creatorEmail) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const startRow = Math.max(2, lastRow - 49);
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, T_WIDTH).getValues();
  const now = new Date();
  const normTitle = String(title || '').trim().toLowerCase();
  const normInstructions = String(instructions || '').trim().toLowerCase();
  const assignee = normalizeEmail(assigneeEmail);
  const creator = normalizeEmail(creatorEmail);
  for (let i = 0; i < rows.length; i++) {
    const created = rows[i][T_COL.CREATED_AT - 1];
    if (!(created instanceof Date)) continue;
    const same = String(rows[i][T_COL.TITLE - 1] || '').trim().toLowerCase() === normTitle &&
      String(rows[i][T_COL.INSTRUCTIONS - 1] || '').trim().toLowerCase() === normInstructions &&
      normalizeEmail(rows[i][T_COL.ASSIGNED_TO_EMAIL - 1]) === assignee &&
      normalizeEmail(rows[i][T_COL.CREATED_BY_EMAIL - 1]) === creator;
    if (same && now - created >= 0 && now - created < TASK_DUPLICATE_WINDOW_MS) return true;
  }
  return false;
}

function normalizeTaskParentType(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (raw === 'question') return 'Question';
  if (raw === 'task') return 'Task';
  return '';
}

function findQuestionOrAnsweredByTicketId(ticketId) {
  const id = String(ticketId || '').trim();
  if (!id) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (qSheet) {
    const r = findRowIndexByTicketId(qSheet, id, Q_COL.TICKET_ID);
    if (r !== -1) {
      const row = qSheet.getRange(r, 1, 1, Q_WIDTH).getValues()[0];
      return {
        sheet: SHEET_QUESTIONS,
        event: String(row[Q_COL.EVENT - 1] || '').trim(),
        link: String(row[Q_COL.LINK - 1] || '').trim(),
        category: normalizeWorkCategory(row[Q_COL.WORK_CATEGORY - 1]),
        askedByEmail: normalizeEmail(row[Q_COL.ASKED_BY_EMAIL - 1])
      };
    }
  }
  const aSheet = ss.getSheetByName(SHEET_ANSWERED);
  if (aSheet) {
    const r = findRowIndexByTicketId(aSheet, id, A_COL.TICKET_ID);
    if (r !== -1) {
      const row = aSheet.getRange(r, 1, 1, A_WIDTH).getValues()[0];
      return {
        sheet: SHEET_ANSWERED,
        event: String(row[A_COL.EVENT - 1] || '').trim(),
        link: String(row[A_COL.LINK - 1] || '').trim(),
        category: normalizeWorkCategory(row[A_COL.WORK_CATEGORY - 1]),
        askedByEmail: normalizeEmail(row[A_COL.ASKED_BY_EMAIL - 1])
      };
    }
  }
  return null;
}

function resolveTaskParentContext(parentType, parentId, actor) {
  const type = normalizeTaskParentType(parentType);
  const id = String(parentId || '').trim();
  if (!type || !id) return { parentType: '', parentId: '', relatedTicketId: '', event: '', link: '', category: '' };

  if (type === 'Question') {
    const q = findQuestionOrAnsweredByTicketId(id);
    if (!q) throw new Error('The parent question could not be found. Refresh and try again.');
    const actorEmail = normalizeEmail(actor && actor.email);
    if (!isSupportMember(actor) && q.askedByEmail !== actorEmail) {
      throw new Error('You do not have access to create a related task from this question.');
    }
    return {
      parentType: 'Question', parentId: id, relatedTicketId: id,
      event: q.event, link: q.link, category: q.category
    };
  }

  const target = requireTaskRow(id);
  const pt = taskRowToObject(target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0]);
  const actorEmail = normalizeEmail(actor && actor.email);
  if (!isSupportMember(actor) && pt.createdByEmail !== actorEmail && pt.assignedToEmail !== actorEmail) {
    throw new Error('You do not have access to create a related task from this parent task.');
  }
  return {
    parentType: 'Task', parentId: id, relatedTicketId: pt.relatedTicketId || '',
    event: pt.relatedClientTalent || '', link: pt.relatedLink || '', category: pt.category || 'General / Other'
  };
}

function normalizeTaskCategoryValue_(value, fallbackValue) {
  const raw = String(value !== undefined ? value : (fallbackValue || '')).trim();
  if (!raw || raw.toLowerCase() === 'none') return 'None';
  return normalizeWorkCategory(raw);
}

function normalizeTaskRelatedEntityType_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const allowed = ['Event','Lead Photographer','Associate Photographer','Lead Videographer','Associate Videographer','Client 1','Client 2','CR','TR','Vendor'];
  const match = allowed.find(function(item) { return item.toLowerCase() === raw.toLowerCase(); });
  if (!match) throw new Error('Please choose a valid relationship type for the related Event / Client / Talent / CR / TR.');
  return match;
}

function validateTaskRelatedContext_(relatedClientTalent, relatedEntityType, relatedLink) {
  const name = String(relatedClientTalent || '').trim();
  let type = normalizeTaskRelatedEntityType_(relatedEntityType);
  const link = String(relatedLink || '').trim();
  if (!name) type = '';
  if (name && !type) throw new Error('Please choose the relationship type for the related Event / Client / Talent / CR / TR.');
  if (name && !link) throw new Error((type || 'Related') + ' Link is required when related Event / Client / Talent / CR / TR is added.');
  if (link && !/^https?:\/\/[^\s]+$/i.test(link)) throw new Error('Please enter a valid Related Link starting with http:// or https://.');
  return { name: name, type: type, link: link };
}

function createTask(payload) {
  const result = withLock(() => {
    payload = payload || {};
    const creator = requireAuthenticatedMember(payload.createdByEmail);
    const assignee = requireTaskAssignee(payload.assignedToEmail);
    const title = String(payload.title || '').trim();
    const instructions = String(payload.instructions || '').trim();
    const dueDate = normalizeTaskDueDate(payload.dueDate) || getNextBusinessTaskDueDate();
    const priority = ''; // Legacy priority column retained only for sheet compatibility.
    const parent = resolveTaskParentContext(payload.parentType, payload.parentId, creator);
    const requestedTaskCategory = Object.prototype.hasOwnProperty.call(payload, 'category') ? String(payload.category || '').trim() : 'None';

    // Ticket -> + Add Task is a clean Create Task workflow. The server must
    // enforce that rule too, because the browser is not the source of truth.
    // The Ticket remains the parent connection, but its Event / Client /
    // Talent / CR / TR and Related Link must not be inherited. A separately
    // selected Related Ticket continues to use the normal inheritance behavior.
    const cleanParentTicketContext = payload.cleanParentTicketContext === true && parent.parentType === 'Question';
    const relatedClientTalent = cleanParentTicketContext
      ? ''
      : (Object.prototype.hasOwnProperty.call(payload, 'relatedClientTalent')
        ? String(payload.relatedClientTalent || '').trim()
        : String(parent.event || '').trim());
    const relatedLink = cleanParentTicketContext
      ? ''
      : (Object.prototype.hasOwnProperty.call(payload, 'relatedLink')
        ? String(payload.relatedLink || '').trim()
        : String(parent.link || '').trim());

    // Backward compatibility: older deployed Create Task forms did not send
    // relatedEntityType. Keep those clients working while the newer UI can
    // explicitly choose Event / Client / Talent / CR / TR relationship types.
    const hasRelatedEntityType = Object.prototype.hasOwnProperty.call(payload, 'relatedEntityType');
    const inheritedType = cleanParentTicketContext
      ? ''
      : ((!hasRelatedEntityType && parent.parentType === 'Question' && relatedClientTalent) ? 'Event' : '');
    const legacyType = cleanParentTicketContext
      ? ''
      : ((!hasRelatedEntityType && relatedClientTalent && relatedLink && !inheritedType) ? 'Event' : inheritedType);
    const relatedContext = validateTaskRelatedContext_(relatedClientTalent, hasRelatedEntityType ? payload.relatedEntityType : legacyType, relatedLink);
    const relatedEntityType = relatedContext.type;
    // Category behaves like ticket intake: an explicit category wins. If the
    // creator leaves Category as None, infer it from the Task content and
    // related Event/Client/Talent/CR/TR context using the same keyword map.
    const taskCategory = (!requestedTaskCategory || requestedTaskCategory.toLowerCase() === 'none')
      ? inferWorkCategory(title + ' ' + stripHtmlToText(instructions), relatedClientTalent || parent.event || '')
      : normalizeTaskCategoryValue_(requestedTaskCategory, parent.category);

    if (!title) throw new Error('Task title is required.');
    if (title.length > 140) throw new Error('Task title is too long. Please keep it under 140 characters.');
    if (!stripHtmlToText(instructions)) throw new Error('Task instructions are required.');
    requireSheetCellLength(instructions, 'Task instructions');
    rejectEmbeddedBase64Image(instructions, 'Task instructions');

    const sheet = ensureTasksSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (isDuplicateTaskSubmission(sheet, title, instructions, assignee.email, creator.email)) {
      return { success: false, duplicate: true, message: 'This looks like the same task was just created. Please check the task list.', _skipDataVersion: true };
    }

    const now = new Date();
    const taskId = Utilities.getUuid();
    const supportOwnerEmail = resolveTaskSupportOwnerEmail(assignee);
    const attentionToday = payload.needAttentionToday === true || payload.attentionToday === true;
    const attentionUntil = attentionToday ? todayCSTDateString() : '';
    sheet.appendRow([
      taskId, title, instructions, assignee.name, assignee.email,
      creator.name, creator.email, now, dueDate, priority,
      TASK_STATUS_PENDING, '', '', parent.parentType ? 'Related Task' : 'Manual', parent.relatedTicketId || '', relatedClientTalent,
      now, creator.name, taskCategory, relatedClientTalent, relatedLink, parent.parentType, parent.parentId,
      supportOwnerEmail, '', '', relatedEntityType,
      attentionToday, attentionToday ? now : '', attentionToday ? creator.name : '', attentionUntil
    ]);
    bumpTaskDataVersion();
    const task = taskRowToObject(sheet.getRange(sheet.getLastRow(), 1, 1, T_WIDTH).getValues()[0]);
    return {
      success: true,
      taskId: taskId,
      task: task,
      _postCommit: {
        creator: creator,
        assignee: assignee,
        dueDate: dueDate,
        taskCategory: taskCategory,
        relatedClientTalent: relatedClientTalent,
        relatedLink: relatedLink,
        relatedEntityType: relatedEntityType,
        parent: parent,
        supportOwnerEmail: supportOwnerEmail
      },
      taskEmailNotification: assignee.email !== creator.email ? {
        recipientEmail: assignee.email,
        recipientName: assignee.name,
        eventType: 'assigned',
        actorName: creator.name,
        actorEmail: creator.email,
        task: task
      } : null
    };
  }, { bumpDataVersion: false, operation: 'createTask' });

  if (result && result.success && result._postCommit) {
    const post = result._postCommit;
    logAudit('TASK_CREATE', post.creator.email, post.creator.name, result.taskId, {
      assignedTo: post.assignee.email, dueDate: post.dueDate, category: post.taskCategory,
      relatedClientTalent: post.relatedClientTalent, relatedLink: post.relatedLink, relatedEntityType: post.relatedEntityType,
      parentType: post.parent.parentType, parentId: post.parent.parentId,
      supportOwnerEmail: post.supportOwnerEmail, entity: 'Task'
    });
    if (String(post.parent.parentType || '').toLowerCase() === 'task' && post.parent.parentId) {
      logAudit('TASK_CHILD_CREATE', post.creator.email, post.creator.name, post.parent.parentId, {
        entity: 'Task', childTaskId: result.taskId, childTitle: result.task.title,
        assignedTo: post.assignee.email, dueDate: post.dueDate
      });
    }
    safeCreateTaskNotification(
      result.task, post.assignee.email, post.assignee.name, 'assigned', post.creator.name, post.creator.email,
      'assigned:' + result.task.taskId + ':' + normalizeEmail(post.assignee.email)
    );
    delete result._postCommit;
  }

  const notification = result && result.taskEmailNotification;
  if (notification) {
    sendTaskEmailIfEnabled(notification.recipientEmail, notification.recipientName, notification.eventType, notification.task, notification.actorName, notification.actorEmail);
    delete result.taskEmailNotification;
  }
  return result;
}


function setTaskAttentionToday(taskId, enabled, requestingEmail) {
  return withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    const target = requireTaskRow(taskId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const task = taskRowToObject(row);
    const actorEmail = normalizeEmail(actor.email);
    const canChange = isAdminMember(actor) || isSupportMember(actor) || task.createdByEmail === actorEmail || task.assignedToEmail === actorEmail;
    if (!canChange) throw new Error('Only the task creator, assignee, Support, or Admin can change Task attention.');
    if (task.status === TASK_STATUS_COMPLETED || task.status === TASK_STATUS_CANCELLED) throw new Error('Completed or cancelled tasks cannot be marked Need Attention Today.');
    const active = enabled === true || String(enabled).toLowerCase() === 'true';
    const now = new Date();
    target.sheet.getRange(target.rowIndex, T_COL.ATTENTION_TODAY).setValue(active);
    target.sheet.getRange(target.rowIndex, T_COL.ATTENTION_SET_AT).setValue(active ? now : '');
    target.sheet.getRange(target.rowIndex, T_COL.ATTENTION_SET_BY).setValue(active ? actor.name : '');
    target.sheet.getRange(target.rowIndex, T_COL.ATTENTION_UNTIL).setValue(active ? todayCSTDateString() : '');
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_AT).setValue(now);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_BY).setValue(actor.name);
    bumpTaskDataVersion();
    logAudit(active ? 'TASK_ATTENTION_TODAY_SET' : 'TASK_ATTENTION_TODAY_CLEARED', actor.email, actor.name, taskId, { attentionToday: active, entity: 'Task' });
    return { success: true, task: taskRowToObject(target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0]) };
  }, { operation: 'setTaskAttentionToday' });
}

function setTaskRelatedTicket(taskId, ticketId, requestingEmail, expected) {
  const result = withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    const target = requireTaskRow(taskId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const task = taskRowToObject(row);
    const actorEmail = normalizeEmail(actor.email);
    const isCreator = task.createdByEmail === actorEmail;
    const support = isSupportMember(actor);

    if (!isCreator && !support) {
      throw new Error('Only the task creator, Support, or Admin can connect a parent ticket to this task.');
    }

    expected = expected || {};
    if (expected.status && String(expected.status) !== task.status) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest status.');
    }
    if (expected.updatedAt && task.updatedAt && String(expected.updatedAt) !== String(task.updatedAt)) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest details.');
    }

    const cleanTicketId = String(ticketId || '').trim();
    const existingTicketId = String(task.relatedTicketId || '').trim();
    if (existingTicketId && cleanTicketId === existingTicketId) {
      return { success: true, task: task, _skipDataVersion: true };
    }
    if (existingTicketId) {
      throw new Error('This task already has a parent ticket. A task journey can only have one parent ticket.');
    }
    if (!cleanTicketId) {
      throw new Error('A parent ticket cannot be removed after it is connected.');
    }
    const ticket = findQuestionOrAnsweredByTicketId(cleanTicketId);
    if (!ticket) throw new Error('That ticket could not be found in open or answered tickets.');

    const now = new Date();
    const taskRows = target.sheet.getLastRow() < 2 ? [] : target.sheet.getRange(2, 1, target.sheet.getLastRow() - 1, T_WIDTH).getValues();
    const records = taskRows.map(function(taskRow, index) {
      return { task: taskRowToObject(taskRow), rowIndex: index + 2 };
    }).filter(function(record) { return !!record.task.taskId; });
    const byId = {};
    records.forEach(function(record) { byId[record.task.taskId] = record; });

    let root = byId[task.taskId] || { task: task, rowIndex: target.rowIndex };
    const ancestorIds = new Set([root.task.taskId]);
    while (String(root.task.parentType || '').toLowerCase() === 'task' && root.task.parentId && byId[root.task.parentId] && !ancestorIds.has(root.task.parentId)) {
      root = byId[root.task.parentId];
      ancestorIds.add(root.task.taskId);
    }

    const family = [];
    const queued = [root.task.taskId];
    const seen = new Set();
    while (queued.length) {
      const currentId = queued.shift();
      if (!currentId || seen.has(currentId)) continue;
      seen.add(currentId);
      const current = byId[currentId];
      if (current) family.push(current);
      records.forEach(function(record) {
        if (String(record.task.parentType || '').toLowerCase() === 'task' && String(record.task.parentId || '') === String(currentId)) queued.push(record.task.taskId);
      });
    }

    const conflict = family.find(function(record) {
      const familyTicketId = String(record.task.relatedTicketId || '').trim();
      return familyTicketId && familyTicketId !== cleanTicketId;
    });
    if (conflict) {
      throw new Error('This task journey is already connected to a different parent ticket.');
    }

    const affected = family.filter(function(record) { return String(record.task.relatedTicketId || '').trim() !== cleanTicketId; });
    affected.forEach(function(record) {
      target.sheet.getRange(record.rowIndex, T_COL.RELATED_TICKET_ID).setValue(cleanTicketId);
      target.sheet.getRange(record.rowIndex, T_COL.UPDATED_AT).setValue(now);
      target.sheet.getRange(record.rowIndex, T_COL.UPDATED_BY).setValue(actor.name);
    });
    bumpTaskDataVersion();

    logAudit('TASK_TICKET_LINK', actor.email, actor.name, taskId, {
      entity: 'Task',
      relatedTicketId: cleanTicketId,
      rootTaskId: root.task.taskId,
      affectedTasks: affected.length,
      parentType: task.parentType,
      parentId: task.parentId
    });

    const updated = taskRowToObject(target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0]);
    return { success: true, task: updated };
  }, { bumpDataVersion: false, operation: 'setTaskRelatedTicket' });

  return result;
}

function updateTaskDueDate(taskId, dueDate, requestingEmail, expected) {
  return withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    const target = requireTaskRow(taskId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const task = taskRowToObject(row);
    const actorEmail = normalizeEmail(actor.email);
    const canViewTask = isSupportMember(actor) || task.assignedToEmail === actorEmail || task.createdByEmail === actorEmail;
    if (!canViewTask) throw new Error('Access denied: you cannot update this task.');
    const cleanDueDate = normalizeTaskDueDate(dueDate);
    if (!cleanDueDate) throw new Error('Please choose a due date.');

    expected = expected || {};
    if (expected.updatedAt && task.updatedAt && String(expected.updatedAt) !== String(task.updatedAt)) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest details.');
    }

    if (cleanDueDate === task.dueDate) {
      return { success: true, task: task, _skipDataVersion: true };
    }

    const now = new Date();
    target.sheet.getRange(target.rowIndex, T_COL.DUE_DATE).setValue(cleanDueDate);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_AT).setValue(now);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_BY).setValue(actor.name);
    bumpTaskDataVersion();

    logAudit('TASK_DUE_DATE', actor.email, actor.name, taskId, {
      entity: 'Task',
      dueDateFrom: task.dueDate || '',
      dueDateTo: cleanDueDate
    });

    const updatedRow = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    return { success: true, task: taskRowToObject(updatedRow) };
  }, { bumpDataVersion: false });
}

// Phase 4: task-detail editing/reassignment. This stays completely separate
// from Questions / Answered records. Only active tasks can be edited, and
// every write rechecks the task's last-known status/updatedAt to avoid one
// user's stale modal overwriting another user's newer change.
function updateTaskDetails(taskId, updates, requestingEmail, expected) {
  const result = withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    const target = requireTaskRow(taskId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const task = taskRowToObject(row);
    const actorEmail = normalizeEmail(actor.email);
    const isCreator = task.createdByEmail === actorEmail;
    const support = isSupportMember(actor);
    const admin = isAdminMember(actor);

    if (!isCreator && !support) {
      throw new Error('Only the task creator, Support, or Admin can edit task details or reassign it.');
    }
    // Completed work is an immutable historical record. Additional work must
    // be captured in a follow-up child task instead of rewriting the original.
    if (task.status === TASK_STATUS_COMPLETED) {
      throw new Error('Completed tasks cannot be edited. Create a follow-up task for additional work.');
    }
    // Admin may still correct cancelled task records.
    if (!admin && task.status === TASK_STATUS_CANCELLED) {
      throw new Error('Closed tasks cannot be edited.');
    }

    expected = expected || {};
    if (expected.status && String(expected.status) !== task.status) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest status.');
    }
    if (expected.updatedAt && task.updatedAt && String(expected.updatedAt) !== String(task.updatedAt)) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest details.');
    }

    updates = updates || {};
    const title = String(updates.title !== undefined ? updates.title : task.title).trim();
    const instructions = String(updates.instructions !== undefined ? updates.instructions : task.instructions).trim();
    const dueDate = normalizeTaskDueDate(updates.dueDate !== undefined ? updates.dueDate : task.dueDate);
    const priority = ''; // Legacy priority column retained only for sheet compatibility.
    const taskCategory = normalizeTaskCategoryValue_(updates.category !== undefined ? updates.category : task.category, task.category);
    const relatedClientTalent = String(updates.relatedClientTalent !== undefined ? updates.relatedClientTalent : task.relatedClientTalent).trim();
    const relatedLink = String(updates.relatedLink !== undefined ? updates.relatedLink : task.relatedLink).trim();
    const relatedContext = validateTaskRelatedContext_(relatedClientTalent, updates.relatedEntityType !== undefined ? updates.relatedEntityType : task.relatedEntityType, relatedLink);
    const relatedEntityType = relatedContext.type;
    const requestedAssigneeEmail = normalizeEmail(updates.assignedToEmail !== undefined ? updates.assignedToEmail : task.assignedToEmail);

    if (!title) throw new Error('Task title is required.');
    if (title.length > 140) throw new Error('Task title is too long. Please keep it under 140 characters.');
    if (!stripHtmlToText(instructions)) throw new Error('Task instructions are required.');
    requireSheetCellLength(instructions, 'Task instructions');
    if (updates.instructions !== undefined) rejectEmbeddedBase64Image(instructions, 'Task instructions');
    const assignee = requireTaskAssignee(requestedAssigneeEmail);

    const assigneeChanged = assignee.email !== task.assignedToEmail;
    const supportOwnerEmail = (assigneeChanged || !task.supportOwnerEmail)
      ? resolveTaskSupportOwnerEmail(assignee)
      : task.supportOwnerEmail;
    const supportOwnerChanged = supportOwnerEmail !== task.supportOwnerEmail;
    const changed = title !== task.title || instructions !== task.instructions || dueDate !== task.dueDate ||
      taskCategory !== task.category || relatedClientTalent !== task.relatedClientTalent || relatedLink !== task.relatedLink || relatedEntityType !== task.relatedEntityType || assigneeChanged || supportOwnerChanged;
    if (!changed) return { success: true, task: task, _skipDataVersion: true };

    const now = new Date();
    target.sheet.getRange(target.rowIndex, T_COL.TITLE).setValue(title);
    target.sheet.getRange(target.rowIndex, T_COL.INSTRUCTIONS).setValue(instructions);
    target.sheet.getRange(target.rowIndex, T_COL.ASSIGNED_TO).setValue(assignee.name);
    target.sheet.getRange(target.rowIndex, T_COL.ASSIGNED_TO_EMAIL).setValue(assignee.email);
    target.sheet.getRange(target.rowIndex, T_COL.DUE_DATE).setValue(dueDate);
    target.sheet.getRange(target.rowIndex, T_COL.PRIORITY).setValue(priority);
    target.sheet.getRange(target.rowIndex, T_COL.WORK_CATEGORY).setValue(taskCategory);
    target.sheet.getRange(target.rowIndex, T_COL.RELATED_EVENT).setValue(relatedClientTalent);
    target.sheet.getRange(target.rowIndex, T_COL.RELATED_CLIENT_TALENT).setValue(relatedClientTalent);
    target.sheet.getRange(target.rowIndex, T_COL.RELATED_LINK).setValue(relatedLink);
    target.sheet.getRange(target.rowIndex, T_COL.RELATED_ENTITY_TYPE).setValue(relatedEntityType);
    target.sheet.getRange(target.rowIndex, T_COL.SUPPORT_OWNER_EMAIL).setValue(supportOwnerEmail);

    let resultingStatus = task.status;
    if (assigneeChanged && task.status === TASK_STATUS_IN_PROGRESS) {
      resultingStatus = TASK_STATUS_PENDING;
      target.sheet.getRange(target.rowIndex, T_COL.STATUS).setValue(TASK_STATUS_PENDING);
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_AT).setValue('');
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETION_NOTE).setValue('');
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_BY_EMAIL).setValue('');
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_BY).setValue('');
    }

    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_AT).setValue(now);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_BY).setValue(actor.name);
    bumpTaskDataVersion();
    logAudit('TASK_EDIT', actor.email, actor.name, taskId, {
      entity: 'Task',
      titleChanged: title !== task.title,
      instructionsChanged: instructions !== task.instructions,
      dueDateFrom: task.dueDate,
      dueDateTo: dueDate,
      categoryFrom: task.category,
      categoryTo: taskCategory,
      relatedClientTalentFrom: task.relatedClientTalent,
      relatedClientTalentTo: relatedClientTalent,
      relatedLinkChanged: relatedLink !== task.relatedLink,
      relatedEntityTypeFrom: task.relatedEntityType,
      relatedEntityTypeTo: relatedEntityType,
      assigneeFrom: task.assignedToEmail,
      assigneeTo: assignee.email,
      supportOwnerFrom: task.supportOwnerEmail,
      supportOwnerTo: supportOwnerEmail,
      statusResetToPending: assigneeChanged && task.status === TASK_STATUS_IN_PROGRESS
    });

    const updatedRow = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const updatedTask = taskRowToObject(updatedRow);
    if (assigneeChanged) {
      safeCreateTaskNotification(
        updatedTask, assignee.email, assignee.name, 'reassigned', actor.name, actor.email,
        'reassigned:' + taskId + ':' + normalizeEmail(assignee.email) + ':' + now.toISOString()
      );
    }
    return {
      success: true,
      task: updatedTask,
      status: resultingStatus,
      taskEmailNotification: assigneeChanged && assignee.email !== actor.email ? {
        recipientEmail: assignee.email,
        recipientName: assignee.name,
        eventType: 'reassigned',
        actorName: actor.name,
        actorEmail: actor.email,
        task: updatedTask
      } : null
    };
  }, { bumpDataVersion: false });

  const notification = result && result.taskEmailNotification;
  if (notification) {
    sendTaskEmailIfEnabled(notification.recipientEmail, notification.recipientName, notification.eventType, notification.task, notification.actorName, notification.actorEmail);
    delete result.taskEmailNotification;
  }
  return result;
}


function updateTaskStatus(taskId, newStatus, completionNote, requestingEmail, expected) {
  const result = withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    const target = requireTaskRow(taskId);
    const row = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const task = taskRowToObject(row);
    const actorEmail = normalizeEmail(actor.email);
    const isAssignee = task.assignedToEmail === actorEmail;
    const isCreator = task.createdByEmail === actorEmail;
    const support = isSupportMember(actor);
    const admin = isAdminMember(actor);
    const requested = String(newStatus || '').trim();
    const allowedStatuses = [TASK_STATUS_PENDING, TASK_STATUS_IN_PROGRESS, TASK_STATUS_COMPLETED, TASK_STATUS_CANCELLED];
    if (!allowedStatuses.includes(requested)) throw new Error('That task status is not valid.');

    expected = expected || {};
    if (expected.status && String(expected.status) !== task.status) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest status.');
    }
    if (expected.updatedAt && task.updatedAt && String(expected.updatedAt) !== String(task.updatedAt)) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest details.');
    }

    if (task.status === TASK_STATUS_COMPLETED) {
      if (requested === TASK_STATUS_COMPLETED) return { success: true, task: task, _skipDataVersion: true };
      throw new Error('A completed task cannot be reopened. Create a follow-up task for additional work.');
    }

    if (!admin && task.status === TASK_STATUS_CANCELLED) {
      if (requested === task.status) return { success: true, task: task, _skipDataVersion: true };
      throw new Error('This task is already ' + task.status.toLowerCase() + ' and can no longer be changed.');
    }

    if (!admin) {
      if (requested === TASK_STATUS_CANCELLED) {
        if (!isCreator && !support) throw new Error('Only the task creator or Support can cancel this task.');
      } else {
        if (!isAssignee) throw new Error('Only the assigned team member can update the working status of this task.');
        if (task.status === TASK_STATUS_IN_PROGRESS && requested === TASK_STATUS_PENDING) {
          throw new Error('An in-progress task cannot be moved back to Pending.');
        }
      }
    }

    const now = new Date();
    const note = String(completionNote || '').trim();
    const becameCompleted = requested === TASK_STATUS_COMPLETED && task.status !== TASK_STATUS_COMPLETED;
    target.sheet.getRange(target.rowIndex, T_COL.STATUS).setValue(requested);
    if (requested === TASK_STATUS_COMPLETED) {
      if (becameCompleted) {
        target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_AT).setValue(now);
        target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_BY_EMAIL).setValue(actor.email);
        target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_BY).setValue(actor.name);
      }
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETION_NOTE).setValue(note);
    } else {
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_AT).setValue('');
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETION_NOTE).setValue('');
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_BY_EMAIL).setValue('');
      target.sheet.getRange(target.rowIndex, T_COL.COMPLETED_BY).setValue('');
    }
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_AT).setValue(now);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_BY).setValue(actor.name);
    bumpTaskDataVersion();
    logAudit('TASK_STATUS', actor.email, actor.name, taskId, {
      from: task.status, to: requested, hasCompletionNote: !!note,
      completedByEmail: becameCompleted ? actor.email : task.completedByEmail,
      entity: 'Task'
    });

    const updatedRow = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const updatedTask = taskRowToObject(updatedRow);
    const creatorEventType = becameCompleted ? 'completed' : (requested === TASK_STATUS_CANCELLED ? 'cancelled' : 'status');
    const creatorNotification = createTaskCreatorProgressNotification(
      updatedTask, creatorEventType, actor,
      creatorEventType + ':' + taskId + ':' + normalizeEmail(task.createdByEmail) + ':' + now.toISOString()
    );
    return {
      success: true,
      task: updatedTask,
      taskEmailNotification: creatorNotification
    };
  }, { bumpDataVersion: false });

  const notification = result && result.taskEmailNotification;
  if (notification) {
    sendTaskEmailIfEnabled(notification.recipientEmail, notification.recipientName, notification.eventType, notification.task, notification.actorName, notification.actorEmail);
    delete result.taskEmailNotification;
  }
  return result;
}

// Rich task replies are kept in the existing Audit Log so the task schema and
// historical rows remain unchanged. Links and Drive-backed image previews are
// short enough to stay safely within the Sheets cell limit.
function addTaskUpdate(taskId, updateHtml, requestingEmail) {
  const result = withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    const target = requireTaskRow(taskId);
    const taskRow = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const task = taskRowToObject(taskRow);
    const actorEmail = normalizeEmail(actor.email);
    const canView = isSupportMember(actor) || task.assignedToEmail === actorEmail || task.createdByEmail === actorEmail;
    if (!canView) throw new Error('Access denied: you cannot update this task.');
    if (task.status === TASK_STATUS_COMPLETED) {
      throw new Error('Completed tasks cannot receive new progress updates. Create a follow-up task for additional work.');
    }

    const content = String(updateHtml || '').trim();
    if (!stripHtmlToText(content)) throw new Error('Please write a task update before posting.');
    requireSheetCellLength(content, 'The task update');
    rejectEmbeddedBase64Image(content, 'The task update');

    const mentions = createMentionNotifications(taskId, task.title, content, actor, 'Task Update', []);
    const now = new Date();
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_AT).setValue(now);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_BY).setValue(actor.name);
    logAudit('TASK_UPDATE', actor.email, actor.name, taskId, {
      entity: 'Task', content: content, mentionedUsers: mentions.map(m => m.email)
    });
    bumpTaskDataVersion();

    const updatedTask = Object.assign({}, task, { updatedAt: now.toISOString(), updatedBy: actor.name });
    const creatorWasMentioned = mentions.some(function(member) {
      return emailsRepresentSameWorkspaceIdentity(member.email, task.createdByEmail);
    });
    const creatorNotification = creatorWasMentioned ? null : createTaskCreatorProgressNotification(
      updatedTask, 'updated', actor,
      'updated:' + taskId + ':' + normalizeEmail(task.createdByEmail) + ':' + now.toISOString()
    );

    return {
      success: true,
      updatedAt: now.toISOString(),
      taskEmailNotification: creatorNotification,
      mentionEmailNotifications: mentions.map(m => ({
        recipientEmail: m.email, recipientName: m.name,
        eventName: task.title, preview: content,
        actorName: actor.name, actorEmail: actor.email
      }))
    };
  }, { bumpDataVersion: false });

  if (result && Array.isArray(result.mentionEmailNotifications)) {
    result.mentionEmailNotifications.forEach(n => {
      sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail, 'task update');
    });
    delete result.mentionEmailNotifications;
  }
  const creatorNotification = result && result.taskEmailNotification;
  if (creatorNotification) {
    sendTaskEmailIfEnabled(creatorNotification.recipientEmail, creatorNotification.recipientName, creatorNotification.eventType, creatorNotification.task, creatorNotification.actorName, creatorNotification.actorEmail);
    delete result.taskEmailNotification;
  }
  return result;
}

// Saves a progress note and keeps active work open in one transaction. If the
// task was Pending it moves to In Progress; an already In Progress task stays
// there. This powers the safer completion decision in the task detail modal.
function addTaskUpdateAndKeepInProgress(taskId, updateHtml, requestingEmail, expected) {
  const result = withLock(() => {
    const actor = requireAuthenticatedMember(requestingEmail);
    const target = requireTaskRow(taskId);
    const taskRow = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    const task = taskRowToObject(taskRow);
    const actorEmail = normalizeEmail(actor.email);
    const canView = isSupportMember(actor) || task.assignedToEmail === actorEmail || task.createdByEmail === actorEmail;
    if (!canView) throw new Error('Access denied: you cannot update this task.');
    if (task.status === TASK_STATUS_COMPLETED || task.status === TASK_STATUS_CANCELLED) {
      throw new Error('Closed tasks cannot receive progress updates. Create a follow-up task for additional work.');
    }

    expected = expected || {};
    if (expected.status && String(expected.status) !== task.status) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest status.');
    }
    if (expected.updatedAt && task.updatedAt && String(expected.updatedAt) !== String(task.updatedAt)) {
      throw new Error('This task changed after you opened it. Please refresh and review the latest details.');
    }

    const content = String(updateHtml || '').trim();
    if (!stripHtmlToText(content)) throw new Error('Please add a progress note before keeping the task in progress.');
    requireSheetCellLength(content, 'The task update');
    rejectEmbeddedBase64Image(content, 'The task update');

    const mentions = createMentionNotifications(taskId, task.title, content, actor, 'Task Update', []);
    const now = new Date();
    const statusChanged = task.status === TASK_STATUS_PENDING;
    if (statusChanged) target.sheet.getRange(target.rowIndex, T_COL.STATUS).setValue(TASK_STATUS_IN_PROGRESS);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_AT).setValue(now);
    target.sheet.getRange(target.rowIndex, T_COL.UPDATED_BY).setValue(actor.name);
    logAudit('TASK_UPDATE', actor.email, actor.name, taskId, {
      entity: 'Task', content: content, mentionedUsers: mentions.map(m => m.email)
    });
    if (statusChanged) {
      logAudit('TASK_STATUS', actor.email, actor.name, taskId, {
        from: TASK_STATUS_PENDING, to: TASK_STATUS_IN_PROGRESS, hasCompletionNote: false, entity: 'Task'
      });
    }
    bumpTaskDataVersion();

    const updatedTask = taskRowToObject(target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0]);
    const creatorWasMentioned = mentions.some(function(member) {
      return emailsRepresentSameWorkspaceIdentity(member.email, task.createdByEmail);
    });
    const creatorNotification = creatorWasMentioned ? null : createTaskCreatorProgressNotification(
      updatedTask, 'updated', actor,
      'updated:' + taskId + ':' + normalizeEmail(task.createdByEmail) + ':' + now.toISOString()
    );
    return {
      success: true, task: updatedTask, updatedAt: now.toISOString(),
      taskEmailNotification: creatorNotification,
      mentionEmailNotifications: mentions.map(m => ({
        recipientEmail: m.email, recipientName: m.name,
        eventName: task.title, preview: content,
        actorName: actor.name, actorEmail: actor.email
      }))
    };
  }, { bumpDataVersion: false, operation: 'addTaskUpdateAndKeepInProgress' });

  if (result && Array.isArray(result.mentionEmailNotifications)) {
    result.mentionEmailNotifications.forEach(n => {
      sendMentionEmailIfEnabled(n.recipientEmail, n.recipientName, n.eventName, n.preview, n.actorName, n.actorEmail, 'task update');
    });
    delete result.mentionEmailNotifications;
  }
  const creatorNotification = result && result.taskEmailNotification;
  if (creatorNotification) {
    sendTaskEmailIfEnabled(creatorNotification.recipientEmail, creatorNotification.recipientName, creatorNotification.eventType, creatorNotification.task, creatorNotification.actorName, creatorNotification.actorEmail);
    delete result.taskEmailNotification;
  }
  return result;
}


// ==========================================
// PHASE 5 TASK ACTIVITY HISTORY
// ==========================================
// Reuses the existing Audit Log rather than introducing another history sheet.
// Only users who can already see the task (assignee, creator, Support/Admin)
// may read its activity timeline.
function getTaskActivityHistory(taskId, requestingEmail) {
  const member = requireAuthenticatedMember(requestingEmail);
  const target = requireTaskRow(taskId);
  const taskRow = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
  const task = taskRowToObject(taskRow);
  const email = normalizeEmail(member.email);
  const canView = isSupportMember(member) || task.assignedToEmail === email || task.createdByEmail === email;
  if (!canView) throw new Error('Access denied: you cannot view activity for this task.');

  const audit = ensureAuditSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (audit.getLastRow() < 2) return [];
  const rows = audit.getRange(2, 1, audit.getLastRow() - 1, 6).getValues();
  const wantedId = String(taskId || '').trim();
  const allowedActions = new Set(['TASK_CREATE', 'TASK_CREATE_FROM_ANSWER', 'TASK_CREATE_FROM_ANSWERED', 'TASK_CHILD_CREATE', 'TASK_DUE_DATE', 'TASK_EDIT', 'TASK_STATUS', 'TASK_UPDATE']);
  const out = [];

  for (let i = rows.length - 1; i >= 0 && out.length < 100; i--) {
    const action = String(rows[i][3] || '').trim();
    const recordId = String(rows[i][4] || '').trim();
    if (recordId !== wantedId || !allowedActions.has(action)) continue;

    let details = {};
    const rawDetails = rows[i][5];
    if (rawDetails && typeof rawDetails === 'object' && !(rawDetails instanceof Date)) {
      details = rawDetails;
    } else {
      try { details = rawDetails ? JSON.parse(String(rawDetails)) : {}; } catch (e) { details = {}; }
    }

    out.push({
      timestamp: rows[i][0] ? safeIsoDate(rows[i][0]) : '',
      actorEmail: normalizeEmail(rows[i][1]),
      actorName: stripNoraPrefix(String(rows[i][2] || '').trim()),
      action: action,
      details: details
    });
  }
  return out;
}

// ==========================================
// TEAM ROSTER MANAGEMENT
// ==========================================
function saveTeamMember(member, requestingEmail) {
  return withLock(() => {
    ensureSheetsExist();
    member = member || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TEAM);

    const newEmail = normalizeEmail(member.email);
    const newName = String(member.name || '').trim();
    if (!newName) throw new Error("Name is required.");
    if (!newEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(newEmail)) throw new Error("A valid email address is required.");

    const newTitle = String(member.title || '').trim();
    if (newTitle && getAllTitles().every(t => t.toLowerCase() !== newTitle.toLowerCase())) {
      throw new Error("That title doesn't exist. Add it first from the \"Manage Titles\" panel.");
    }

    const newCategory = String(member.category || '').trim();
    if (!['Coordinator', 'Support', 'Admin'].includes(newCategory)) {
      throw new Error("Category must be Coordinator, Support, or Admin - these are the only values that carry real permissions.");
    }

    const originalEmail = normalizeEmail(member.originalEmail);
    const originalProfileKey = String(member.originalProfileKey || '').trim();
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    if (originalProfileKey) {
      for (let i = 1; i < data.length; i++) {
        const rowMember = {
          name: stripNoraPrefix(String(data[i][0] || '').trim()),
          title: String(data[i][1] || '').trim(),
          status: String(data[i][2] || '').trim(),
          category: String(data[i][3] || '').trim(),
          email: normalizeEmail(data[i][4])
        };
        if (memberProfileKey(rowMember) === originalProfileKey) { targetRow = i + 1; break; }
      }
      if (targetRow === -1) throw new Error("This team profile no longer exists - someone may have just changed or removed it. Please refresh.");
    } else if (originalEmail) {
      // Legacy edit payload compatibility. Only use email if it identifies one row.
      const matches = [];
      for (let i = 1; i < data.length; i++) if (normalizeEmail(data[i][4]) === originalEmail) matches.push(i + 1);
      if (matches.length === 1) targetRow = matches[0];
      else if (matches.length > 1) throw new Error("This email has multiple profiles. Refresh Team Setup and edit the specific profile again.");
      else throw new Error("This team member no longer exists - someone may have just removed them. Please refresh.");
    }

    // The same Google Workspace mailbox may intentionally own more than one
    // Support Hub profile (for example Support + Admin). What must stay unique
    // is the email/category pair; two Support rows or two Admin rows using the
    // same mailbox would be ambiguous and are therefore blocked.
    for (let i = 1; i < data.length; i++) {
      if ((i + 1) === targetRow) continue;
      const rowEmail = normalizeEmail(data[i][4]);
      const rowCategory = String(data[i][3] || '').trim();
      if (rowEmail === newEmail && rowCategory === newCategory) {
        throw new Error("Another " + newCategory + " profile already uses that email address.");
      }
    }

    const newStatus = String(member.status || '').trim();
    const existingCategory = targetRow !== -1 ? String(sheet.getRange(targetRow, 4).getValue()).trim() : '';
    const touchesAdmin = newCategory === 'Admin' || existingCategory === 'Admin';

    if (touchesAdmin) {
      if (countWorkingAdmins() > 0) {
        requireAuthenticatedAdmin(requestingEmail);
      } else {
        requireAuthenticatedSupport(requestingEmail);
      }
    } else {
      requireAuthenticatedSupport(requestingEmail);
    }

    let passwordHash = '';
    let passwordSalt = '';
    let wasWorkingAdmin = false;
    if (targetRow !== -1) {
      const existing = sheet.getRange(targetRow, 1, 1, 7).getValues()[0];
      passwordHash = existing[5] || '';
      passwordSalt = existing[6] || '';
      wasWorkingAdmin = String(existing[3]).trim() === 'Admin'
        && String(existing[2]).trim().toLowerCase() === 'active'
        && !!String(existing[5] || '').trim();
    }

    const losesAdminAbility = wasWorkingAdmin && (newCategory !== 'Admin' || newStatus.toLowerCase() !== 'active');
    if (losesAdminAbility && countWorkingAdmins() <= 1) {
      throw new Error("This is the last working Admin account - it cannot be demoted or deactivated. Create another Admin first.");
    }

    if (newCategory === 'Admin') {
      if (member.password) {
        passwordSalt = Utilities.getUuid();
        passwordHash = hashPassword(member.password, passwordSalt);
      }
      if (targetRow === -1 && !passwordHash) {
        throw new Error("A password is required when creating a new Admin account.");
      }
    } else {
      passwordHash = '';
      passwordSalt = '';
    }

    const existingPrimary = targetRow !== -1 ? normalizeEmail(data[targetRow - 1][TEAM_COL.PRIMARY_SUPPORT_EMAIL - 1]) : '';
    const existingBackup = targetRow !== -1 ? normalizeEmail(data[targetRow - 1][TEAM_COL.BACKUP_SUPPORT_EMAIL - 1]) : '';
    const hasRoutingUpdate = Object.prototype.hasOwnProperty.call(member, 'primarySupportEmail') || Object.prototype.hasOwnProperty.call(member, 'backupSupportEmail');
    if (hasRoutingUpdate) requireAuthenticatedAdmin(requestingEmail);

    let primarySupportEmail = Object.prototype.hasOwnProperty.call(member, 'primarySupportEmail') ? normalizeEmail(member.primarySupportEmail) : existingPrimary;
    let backupSupportEmail = Object.prototype.hasOwnProperty.call(member, 'backupSupportEmail') ? normalizeEmail(member.backupSupportEmail) : existingBackup;
    if (newCategory !== 'Coordinator') {
      primarySupportEmail = '';
      backupSupportEmail = '';
    } else {
      if (primarySupportEmail) primarySupportEmail = requireQuestionAssignee(primarySupportEmail).email;
      if (backupSupportEmail) backupSupportEmail = requireQuestionAssignee(backupSupportEmail).email;
      if (primarySupportEmail && backupSupportEmail && primarySupportEmail === backupSupportEmail) {
        throw new Error('Primary Support and Backup Support must be different people.');
      }
    }

    const actor = requireAuthenticatedMember(requestingEmail);
    const row = [newName, newTitle, newStatus, newCategory, newEmail, passwordHash, passwordSalt, primarySupportEmail, backupSupportEmail];

    if (targetRow !== -1) {
      sheet.getRange(targetRow, 1, 1, TEAM_WIDTH).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    invalidateTeamCache();
    logAudit(targetRow !== -1 ? 'TEAM_EDIT' : 'TEAM_ADD', actor.email, actor.name, '', { member: newEmail, category: newCategory, status: newStatus, title: newTitle, primarySupportEmail: primarySupportEmail, backupSupportEmail: backupSupportEmail });
    return {
      success: true,
      member: {
        name: stripNoraPrefix(newName),
        title: newTitle,
        status: newStatus,
        category: newCategory,
        email: newEmail,
        primarySupportEmail: primarySupportEmail,
        backupSupportEmail: backupSupportEmail
      }
    };
  });
}

function deleteTeamMember(selector, requestingEmail) {
  return withLock(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TEAM);
    const profileKey = selector && typeof selector === 'object' ? String(selector.profileKey || '').trim() : '';
    const targetEmail = normalizeEmail(selector && typeof selector === 'object' ? selector.email : selector);

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    let targetCategory = '';
    let targetIsWorkingAdmin = false;
    for (let i = 1; i < data.length; i++) {
      const rowMember = {
        name: stripNoraPrefix(String(data[i][0] || '').trim()),
        title: String(data[i][1] || '').trim(),
        status: String(data[i][2] || '').trim(),
        category: String(data[i][3] || '').trim(),
        email: normalizeEmail(data[i][4])
      };
      const matches = profileKey ? memberProfileKey(rowMember) === profileKey : rowMember.email === targetEmail;
      if (!matches) continue;
      targetRow = i + 1;
      targetCategory = rowMember.category;
      targetIsWorkingAdmin = targetCategory === 'Admin'
        && String(data[i][2]).trim().toLowerCase() === 'active'
        && !!String(data[i][5] || '').trim();
      break;
    }
    if (targetRow === -1) throw new Error("This team profile no longer exists - someone may have already removed it.");

    if (targetCategory === 'Admin') {
      if (countWorkingAdmins() > 0) requireAuthenticatedAdmin(requestingEmail);
      else requireAuthenticatedSupport(requestingEmail);
    } else {
      requireAuthenticatedSupport(requestingEmail);
    }

    if (targetIsWorkingAdmin && countWorkingAdmins() <= 1) {
      throw new Error("This is the last working Admin account - it cannot be deleted. Create another Admin first.");
    }

    const actor = requireAuthenticatedMember(requestingEmail);
    const deletedName = String(data[targetRow - 1][0] || '').trim();
    const deletedEmail = normalizeEmail(data[targetRow - 1][4]);
    sheet.deleteRow(targetRow);
    invalidateTeamCache();
    logAudit('TEAM_DELETE', actor.email, actor.name, '', { member: deletedEmail, profile: deletedName, category: targetCategory });
    return { success: true };
  });
}

// ==========================================
// ADMIN-ONLY: DELETE TASK (SOFT ARCHIVE)
// ==========================================
function deleteTask(taskId, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const target = requireTaskRow(taskId);
    const rowValues = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
    archiveDeletedRow(SHEET_TASKS, rowValues, taskId, admin.email);
    target.sheet.deleteRow(target.rowIndex);
    bumpTaskDataVersion();
    logAudit('TASK_DELETE', admin.email, admin.name, taskId, { archived: true, entity: 'Task' });
    return { success: true };
  }, { bumpDataVersion: false });
}


function adminBulkDeleteTasks(taskIds, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const ids = Array.from(new Set((taskIds || []).map(id => String(id || '').trim()).filter(Boolean)));
    if (!ids.length) throw new Error('Select at least one task to delete.');
    if (ids.length > 500) throw new Error('Please delete 500 tasks or fewer at one time.');

    const targets = ids.map(id => {
      const target = requireTaskRow(id);
      const rowValues = target.sheet.getRange(target.rowIndex, 1, 1, T_WIDTH).getValues()[0];
      return { id: id, sheet: target.sheet, rowIndex: target.rowIndex, rowValues: rowValues };
    });

    targets.forEach(item => archiveDeletedRow(SHEET_TASKS, item.rowValues, item.id, admin.email));
    targets.slice().sort((a, b) => b.rowIndex - a.rowIndex).forEach(item => item.sheet.deleteRow(item.rowIndex));
    targets.forEach(item => logAudit('TASK_DELETE', admin.email, admin.name, item.id, { archived: true, entity: 'Task', bulk: true }));
    bumpTaskDataVersion();
    return { success: true, deletedCount: targets.length };
  }, { bumpDataVersion: false });
}

// ==========================================
// ADMIN-ONLY: EDIT / DELETE QUESTIONS & ANSWERS
// ==========================================
function updateQuestion(ticketId, updates, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    updates = updates || {};
    const target = requireQuestionRow(ticketId);
    const qSheet = target.sheet;
    const rowIndex = target.rowIndex;

    const current = qSheet.getRange(rowIndex, 1, 1, Q_WIDTH).getValues()[0];

    if (updates.expected && updates.expected.question !== undefined) {
      if (String(current[Q_COL.QUESTION - 1] || '').trim() !== String(updates.expected.question || '').trim()) {
        throw new Error("This record changed after you opened it (someone else edited or updated it). Please refresh and try again.");
      }
    }

    if (updates.eventName !== undefined && !isValidEventName(updates.eventName)) {
      throw new Error(EVENT_NAME_INVALID_MSG);
    }
    if (updates.caseLink !== undefined && String(updates.caseLink || '').trim() && !isValidCaseLink(updates.caseLink)) {
      throw new Error("Please provide a valid Case / Event URL (it must start with http:// or https://).");
    }

    const merged = current.slice();
    if (updates.question !== undefined) {
      requireSheetCellLength(updates.question, 'The question');
      merged[Q_COL.QUESTION - 1] = updates.question;
    }
    if (updates.eventName !== undefined) merged[Q_COL.EVENT - 1] = updates.eventName;
    if (updates.askedBy !== undefined) merged[Q_COL.ASKED_BY - 1] = updates.askedBy;
    if (updates.isUrgent !== undefined) {
      merged[Q_COL.PRIORITY - 1] = ""; // Legacy priority column intentionally unused.
      merged[Q_COL.EVENT_THIS_WEEKEND - 1] = !!updates.isUrgent;
    }
    if (updates.eventDate !== undefined) merged[Q_COL.EVENT_DATE - 1] = normalizeEventDate(updates.eventDate);
    if (updates.talentRole !== undefined) merged[Q_COL.TALENT_ROLE - 1] = normalizeEventRole(updates.talentRole);
    if (updates.category !== undefined) merged[Q_COL.WORK_CATEGORY - 1] = normalizeWorkCategory(updates.category);
    if (updates.caseLink !== undefined) merged[Q_COL.LINK - 1] = updates.caseLink;

    if (merged[Q_COL.EVENT_THIS_WEEKEND - 1] === true && !normalizeEventDate(merged[Q_COL.EVENT_DATE - 1])) {
      throw new Error('Please add the event date when selecting Event this weekend.');
    }

    qSheet.getRange(rowIndex, 1, 1, Q_WIDTH).setValues([merged]);
    logAudit('EDIT_QUESTION', admin.email, admin.name, ticketId, { fields: Object.keys(updates).filter(k => k !== 'expected') });
    return { success: true };
  });
}

function deleteQuestion(ticketId, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const target = requireQuestionRow(ticketId);
    const rowValues = target.sheet.getRange(target.rowIndex, 1, 1, Q_WIDTH).getValues()[0];
    archiveDeletedRow(SHEET_QUESTIONS, rowValues, ticketId, admin.email);
    target.sheet.deleteRow(target.rowIndex);
    logAudit('DELETE_QUESTION', admin.email, admin.name, ticketId, { archived: true });
    return { success: true };
  });
}

function updateAnswer(ticketId, updates, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    updates = updates || {};
    const target = requireAnsweredRow(ticketId);
    const aSheet = target.sheet;
    const rowIndex = target.rowIndex;

    const current = aSheet.getRange(rowIndex, 1, 1, A_WIDTH).getValues()[0];

    if (updates.expected && updates.expected.question !== undefined) {
      const questionChanged = String(current[A_COL.QUESTION - 1] || '').trim() !== String(updates.expected.question || '').trim();
      const answerChanged = updates.expected.answer !== undefined
        && String(current[A_COL.ANSWER - 1] || '').trim() !== String(updates.expected.answer || '').trim();
      if (questionChanged || answerChanged) {
        throw new Error("This record changed after you opened it (someone else edited or updated it). Please refresh and try again.");
      }
    }

    if (updates.eventName !== undefined && !isValidEventName(updates.eventName)) {
      throw new Error(EVENT_NAME_INVALID_MSG);
    }
    if (updates.caseLink !== undefined && String(updates.caseLink || '').trim() && !isValidCaseLink(updates.caseLink)) {
      throw new Error("Please provide a valid Case / Event URL (it must start with http:// or https://).");
    }

    const merged = current.slice();
    if (updates.question !== undefined) {
      requireSheetCellLength(updates.question, 'The question');
      merged[A_COL.QUESTION - 1] = updates.question;
    }
    if (updates.eventName !== undefined) merged[A_COL.EVENT - 1] = updates.eventName;
    if (updates.askedBy !== undefined) merged[A_COL.ASKED_BY - 1] = updates.askedBy;
    if (updates.caseLink !== undefined) merged[A_COL.LINK - 1] = updates.caseLink;
    if (updates.answer !== undefined) {
      requireSheetCellLength(updates.answer, 'The answer');
      merged[A_COL.ANSWER - 1] = updates.answer;
    }
    if (updates.answeredBy !== undefined) merged[A_COL.ANSWERED_BY - 1] = updates.answeredBy;
    if (updates.eventDate !== undefined) merged[A_COL.EVENT_DATE - 1] = normalizeEventDate(updates.eventDate);
    if (updates.talentRole !== undefined) merged[A_COL.TALENT_ROLE - 1] = normalizeEventRole(updates.talentRole);
    if (updates.category !== undefined) merged[A_COL.WORK_CATEGORY - 1] = normalizeWorkCategory(updates.category);

    aSheet.getRange(rowIndex, 1, 1, A_WIDTH).setValues([merged]);
    logAudit('EDIT_ANSWER', admin.email, admin.name, ticketId, { fields: Object.keys(updates).filter(k => k !== 'expected') });
    return { success: true };
  });
}

function deleteAnswer(ticketId, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const target = requireAnsweredRow(ticketId);
    const rowValues = target.sheet.getRange(target.rowIndex, 1, 1, A_WIDTH).getValues()[0];
    archiveDeletedRow(SHEET_ANSWERED, rowValues, ticketId, admin.email);
    target.sheet.deleteRow(target.rowIndex);
    logAudit('DELETE_ANSWER', admin.email, admin.name, ticketId, { archived: true });
    return { success: true };
  });
}


function adminBulkDeleteTickets(records, requestingEmail) {
  return withLock(() => {
    const admin = requireAuthenticatedAdmin(requestingEmail);
    const raw = Array.isArray(records) ? records : [];
    const seen = new Set();
    const normalized = raw.map(item => ({
      sheet: String(item && item.sheet || '').trim(),
      ticketId: String(item && item.ticketId || '').trim()
    })).filter(item => {
      if (!item.ticketId || !['Questions Tracker', 'Answered'].includes(item.sheet)) return false;
      const key = item.sheet + '::' + item.ticketId;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    if (!normalized.length) throw new Error('Select at least one ticket to delete.');
    if (normalized.length > 500) throw new Error('Please delete 500 tickets or fewer at one time.');

    const targets = normalized.map(item => {
      const target = item.sheet === 'Answered' ? requireAnsweredRow(item.ticketId) : requireQuestionRow(item.ticketId);
      const width = item.sheet === 'Answered' ? A_WIDTH : Q_WIDTH;
      const rowValues = target.sheet.getRange(target.rowIndex, 1, 1, width).getValues()[0];
      return {
        sheetName: item.sheet,
        ticketId: item.ticketId,
        sheet: target.sheet,
        rowIndex: target.rowIndex,
        rowValues: rowValues
      };
    });

    targets.forEach(item => archiveDeletedRow(item.sheetName === 'Answered' ? SHEET_ANSWERED : SHEET_QUESTIONS, item.rowValues, item.ticketId, admin.email));

    const grouped = {};
    targets.forEach(item => {
      const name = item.sheetName;
      if (!grouped[name]) grouped[name] = [];
      grouped[name].push(item);
    });
    Object.keys(grouped).forEach(name => {
      grouped[name].sort((a, b) => b.rowIndex - a.rowIndex).forEach(item => item.sheet.deleteRow(item.rowIndex));
    });

    targets.forEach(item => logAudit(item.sheetName === 'Answered' ? 'DELETE_ANSWER' : 'DELETE_QUESTION', admin.email, admin.name, item.ticketId, {
      archived: true,
      bulk: true
    }));
    return { success: true, deletedCount: targets.length };
  });
}


// ==========================================
// PHASE 9.12 — PEER QUESTIONS
// Lightweight Coordinator-to-Coordinator help. Kept separate from Supervisor
// Desk so peer help does not affect Support SLA/queue reporting until escalated.
// ==========================================
const P_COL = { ID:1, QUESTION:2, EVENT:3, ASKED_BY:4, ASKED_EMAIL:5, ASSIGNED_TO:6, ASSIGNED_EMAIL:7, CREATED:8, UPDATED:9, STATUS:10, THREAD:11, CASE_LINK:12, CLOSED_AT:13, CLOSED_BY:14 };
const P_WIDTH = 14;
const PEER_STATUS_OPEN = 'Open';
const PEER_STATUS_CLOSED = 'Closed';
const CONFIG_KEY_PEER_VERSION = 'PEER_DATA_VERSION_COUNTER';

function ensurePeerQuestionsSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_PEER_QUESTIONS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PEER_QUESTIONS);
    sh.appendRow(['Peer Ticket ID','Question','Event / Client / Talent Name','Asked By','Asked By Email','Assigned To','Assigned To Email','Created At','Updated At','Status','Conversation JSON','Case / Event Link','Closed At','Closed By']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function getPeerDataVersion(){ return Number(PropertiesService.getScriptProperties().getProperty(CONFIG_KEY_PEER_VERSION)||'0'); }
function bumpPeerDataVersion(){ const n=getPeerDataVersion()+1; PropertiesService.getScriptProperties().setProperty(CONFIG_KEY_PEER_VERSION,String(n)); return n; }
function isPeerEligibleMember(m){ return !!m && String(m.status||'').trim().toLowerCase()==='active' && String(m.category||'').trim().toLowerCase()!=='admin'; }

// Peer Questions intentionally require a real web URL. If there is no specific
// Salesforce record URL, the UI can suggest http://login.salesforce.com.
function isValidPeerRelatedLink(value) {
  const link = String(value || '').trim();
  return /^https?:\/\/[^\s]+$/i.test(link);
}

function requirePeerRelatedLink(value) {
  const link = String(value || '').trim();
  if (!link) {
    throw new Error('Related Link is required. If there is no specific URL, use http://login.salesforce.com.');
  }
  if (!isValidPeerRelatedLink(link)) {
    throw new Error('Enter a valid Related Link starting with http:// or https://. If needed, use http://login.salesforce.com.');
  }
  return link;
}

function requirePeerEligible(email){
  const e=normalizeEmail(email);
  const matches=_getTeamMembersInternal().filter(m=>normalizeEmail(m.email)===e && isPeerEligibleMember(m));
  if(matches.length===1) return matches[0];
  if(!matches.length) throw new Error('Peer Questions can only be assigned to an active non-Admin team member.');
  throw new Error('More than one active Coordinator profile uses this email. Please correct Team Setup.');
}
function parsePeerThread(raw){ try{ const a=JSON.parse(String(raw||'[]')); return Array.isArray(a)?a:[]; }catch(e){ return []; } }
function peerRowToObject(row){
  return { id:String(row[0]||''), question:String(row[1]||''), eventName:String(row[2]||''), askedBy:String(row[3]||''), askedByEmail:normalizeEmail(row[4]), assignedTo:String(row[5]||''), assignedToEmail:normalizeEmail(row[6]), created:safeIsoDate(row[7]), updated:safeIsoDate(row[8]), status:String(row[9]||PEER_STATUS_OPEN), thread:parsePeerThread(row[10]), caseLink:String(row[11]||''), closedAt:row[12]?safeIsoDate(row[12]):'', closedBy:String(row[13]||'') };
}
function getPeerQuestionsData(requestingEmail){
  const actor=requireAuthenticatedMember(requestingEmail);
  const actorEmail=normalizeEmail(actor.email);
  const sh=ensurePeerQuestionsSheet();
  if(sh.getLastRow()<2) return {version:getPeerDataVersion(),data:[]};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,P_WIDTH).getValues();
  const canSeeAll=isSupportMember(actor)||isAdminMember(actor);
  const data=rows.map(peerRowToObject).filter(p=>canSeeAll || p.askedByEmail===actorEmail || p.assignedToEmail===actorEmail);
  return {version:getPeerDataVersion(),data:data};
}
function getPeerQuestionsDataIfChanged(requestingEmail, clientVersion){
  const actor=requireAuthenticatedMember(requestingEmail);
  const v=getPeerDataVersion();
  if(Number(clientVersion)===v) return {unchanged:true,version:v};
  return getPeerQuestionsData(actor.email);
}
function findPeerRow(id){
  const sh=ensurePeerQuestionsSheet(); if(sh.getLastRow()<2) throw new Error('Peer question not found.');
  const vals=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(let i=0;i<vals.length;i++) if(String(vals[i][0])===String(id)) return {sheet:sh,rowIndex:i+2,row:sh.getRange(i+2,1,1,P_WIDTH).getValues()[0]};
  throw new Error('Peer question not found.');
}
// One server call supplies the two dropdowns used by the Peer Question action
// dialogs. This keeps the browser from asking users to type an email address.
function getPeerQuestionActionOptions(id, requestingEmail) {
  const actor = requireAuthenticatedMember(requestingEmail);
  const target = findPeerRow(id);
  const peer = peerRowToObject(target.row);
  const actorEmail = normalizeEmail(actor.email);
  const isParticipant = actorEmail===peer.askedByEmail || actorEmail===peer.assignedToEmail;
  if(!isParticipant && !isSupportMember(actor) && !isAdminMember(actor)) {
    throw new Error('Access denied: you cannot manage this Peer Question.');
  }

  const roster = _getTeamMembersInternal().filter(m => String(m.status||'').trim().toLowerCase()==='active');
  const peers = roster
    .filter(isPeerEligibleMember)
    .filter(m => normalizeEmail(m.email)!==peer.askedByEmail && normalizeEmail(m.email)!==peer.assignedToEmail)
    .map(m => ({name:m.name, email:normalizeEmail(m.email), title:String(m.title||'')}))
    .sort((a,b)=>String(a.name||'').localeCompare(String(b.name||'')));

  const support = roster
    .filter(m => String(m.category||'').trim().toLowerCase()==='support')
    .map(m => ({name:m.name, email:normalizeEmail(m.email), title:String(m.title||'')}))
    .sort((a,b)=>String(a.name||'').localeCompare(String(b.name||'')));

  return {
    success:true,
    peerId:peer.id,
    reassignOptions:peers,
    supportOptions:support,
    canCloseWithoutReply:actorEmail===peer.askedByEmail,
    replyCount:(peer.thread||[]).filter(item => String(item.type||'').toLowerCase()==='reply').length
  };
}

function createPeerQuestion(payload){
  return withLock(()=>{
    payload=payload||{};
    const actor=requireAuthenticatedMember(payload.askedByEmail);
    if(!isPeerEligibleMember(actor)) throw new Error('Peer Questions are available to active non-Admin team members.');
    const assignee=requirePeerEligible(payload.assignedToEmail);
    if(normalizeEmail(assignee.email)===normalizeEmail(actor.email)) throw new Error('Choose another team member for a Peer Question.');
    const q=String(payload.question||'').trim(), ev=String(payload.eventName||'').trim();
    const link=requirePeerRelatedLink(payload.caseLink);
    if(!q) throw new Error('Please enter your question.');
    requireSheetCellLength(q, 'The Peer Question');
    if(!isValidEventName(ev)) throw new Error('Enter the Event / Client / Talent Name as text.');
    const now=new Date(), id=Utilities.getUuid();
    const thread=[{type:'question',text:q,by:actor.name,email:actor.email,at:now.toISOString()}];
    const serializedThread=JSON.stringify(thread); requireSheetCellLength(serializedThread,'The Peer Question conversation');
    ensurePeerQuestionsSheet().appendRow([id,q,ev,actor.name,actor.email,assignee.name,assignee.email,now,now,PEER_STATUS_OPEN,serializedThread,link,'','']);
    bumpPeerDataVersion();
    logAudit('PEER_QUESTION_CREATE',actor.email,actor.name,id,{assignedTo:assignee.email,event:ev});
    return {success:true,id:id};
  },{bumpDataVersion:false});
}
function replyPeerQuestion(id,text,requestingEmail){
  return withLock(()=>{
    const actor=requireAuthenticatedMember(requestingEmail), t=findPeerRow(id), p=peerRowToObject(t.row), ae=normalizeEmail(actor.email);
    if(p.status!==PEER_STATUS_OPEN) throw new Error('This Peer Question is closed.');
    if(!(ae===p.askedByEmail || ae===p.assignedToEmail || isSupportMember(actor))) throw new Error('Only the asker, assigned peer, or Support can reply.');
    const msg=String(text||'').trim(); if(!msg) throw new Error('Write a reply first.');
    const thread=p.thread;
    const last=thread.length?thread[thread.length-1]:null;
    if(last && normalizeEmail(last.email)===ae && String(last.text||'').trim()===msg){
      const lastAt=new Date(last.at||0); if(!isNaN(lastAt.getTime()) && new Date()-lastAt<TASK_DUPLICATE_WINDOW_MS) return {success:true,duplicate:true,_skipDataVersion:true};
    }
    thread.push({type:'reply',text:msg,by:actor.name,email:actor.email,at:new Date().toISOString()});
    const serialized=JSON.stringify(thread); requireSheetCellLength(serialized,'The Peer Question conversation');
    t.sheet.getRange(t.rowIndex,P_COL.THREAD).setValue(serialized); t.sheet.getRange(t.rowIndex,P_COL.UPDATED).setValue(new Date());
    bumpPeerDataVersion(); logAudit('PEER_QUESTION_REPLY',actor.email,actor.name,id,{recipient:ae===p.askedByEmail?p.assignedToEmail:p.askedByEmail});
    return {success:true};
  },{bumpDataVersion:false});
}
function reassignPeerQuestion(id,newEmail,requestingEmail){
  return withLock(()=>{
    const actor=requireAuthenticatedMember(requestingEmail), t=findPeerRow(id), p=peerRowToObject(t.row), ae=normalizeEmail(actor.email);
    if(p.status!==PEER_STATUS_OPEN) throw new Error('This Peer Question is closed.');
    if(!(ae===p.askedByEmail || ae===p.assignedToEmail || isSupportMember(actor))) throw new Error('Only participants or Support can reassign this Peer Question.');
    const m=requirePeerEligible(newEmail);
    if(normalizeEmail(m.email)===p.askedByEmail) throw new Error('Choose a teammate other than the original asker.');
    if(normalizeEmail(m.email)===p.assignedToEmail) return {success:true,_skipDataVersion:true};
    t.sheet.getRange(t.rowIndex,P_COL.ASSIGNED_TO,1,2).setValues([[m.name,m.email]]); t.sheet.getRange(t.rowIndex,P_COL.UPDATED).setValue(new Date());
    bumpPeerDataVersion(); logAudit('PEER_QUESTION_REASSIGN',actor.email,actor.name,id,{assignedTo:m.email}); return {success:true};
  },{bumpDataVersion:false});
}
function closePeerQuestion(id,requestingEmail){
  return withLock(()=>{
    const actor=requireAuthenticatedMember(requestingEmail), t=findPeerRow(id), p=peerRowToObject(t.row), ae=normalizeEmail(actor.email);
    if(!(ae===p.askedByEmail || ae===p.assignedToEmail || isSupportMember(actor))) throw new Error('Only participants or Support can close this Peer Question.');
    if(p.status===PEER_STATUS_CLOSED) return {success:true,_skipDataVersion:true};

    // The original asker may close their own Peer Question whenever it is no
    // longer needed. The assigned peer must provide at least one reply before
    // closing so a Peer Question cannot be closed without an answer.
    const isAsker = ae===p.askedByEmail;
    const replyCount = (p.thread||[]).filter(item => String(item.type||'').toLowerCase()==='reply').length;
    if(!isAsker && replyCount<1){
      throw new Error('Add at least one reply before closing this Peer Question. The original asker can close it without a reply if help is no longer needed.');
    }

    const now=new Date(); t.sheet.getRange(t.rowIndex,P_COL.STATUS).setValue(PEER_STATUS_CLOSED); t.sheet.getRange(t.rowIndex,P_COL.UPDATED).setValue(now); t.sheet.getRange(t.rowIndex,P_COL.CLOSED_AT,1,2).setValues([[now,actor.name]]);
    bumpPeerDataVersion(); logAudit('PEER_QUESTION_CLOSE',actor.email,actor.name,id,{replyCount:replyCount}); return {success:true};
  },{bumpDataVersion:false});
}
function escalatePeerQuestion(id,supportEmail,requestingEmail){
  return withLock(()=>{
    const actor=requireAuthenticatedMember(requestingEmail), t=findPeerRow(id), p=peerRowToObject(t.row), ae=normalizeEmail(actor.email);
    if(!(ae===p.askedByEmail || ae===p.assignedToEmail)) throw new Error('Only participants can escalate this Peer Question.');
    if(p.status!==PEER_STATUS_OPEN) throw new Error('This Peer Question is already closed.');
    const support=requireQuestionAssignee(supportEmail);
    const threadText=p.thread.map(x=>'['+(x.by||'User')+'] '+String(x.text||'')).join('\n\n');
    requireSheetCellLength('Peer escalation:\n'+threadText,'The escalated Peer Question history');
    const qSheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS), now=new Date(), ticketId=Utilities.getUuid();
    qSheet.appendRow(['Peer escalation:\n'+threadText,p.eventName,p.askedBy,now,'','', '',p.caseLink,'',STATUS_OPEN,support.email,ticketId,'','',0,p.askedByEmail,true,'Escalated from Peer Question '+p.id+' by '+actor.name,now,actor.name,inferWorkCategory(threadText,p.eventName),false,'','',support.email,false,'','','','','','','','',false,'','','','','']);
    t.sheet.getRange(t.rowIndex,P_COL.STATUS).setValue(PEER_STATUS_CLOSED); t.sheet.getRange(t.rowIndex,P_COL.UPDATED).setValue(now); t.sheet.getRange(t.rowIndex,P_COL.CLOSED_AT,1,2).setValues([[now,actor.name]]);
    bumpPeerDataVersion(); logAudit('PEER_QUESTION_ESCALATE',actor.email,actor.name,id,{support:support.email,ticketId:ticketId});
    return {success:true,ticketId:ticketId};
  });
}
