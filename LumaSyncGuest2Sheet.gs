/**
 * LumaSyncGuest2Sheet.gs
 * ---------------------------------------------------------------------------
 * Google Apps Script that pulls an event's guest list from the Luma public API
 * and writes it into a Google Sheet, then stamps a "last updated" timestamp on
 * a dashboard sheet.
 *
 * Highlights:
 *   - Fetches every approval status you list (Luma's get-guests filters by status,
 *     so omitting a status drops those guests — e.g. all pending_approval ones).
 *   - Adopts the column order already present in the destination sheet's header
 *     row; custom registration questions are matched by label, ignoring case and
 *     stray whitespace. Questions with no column yet are appended automatically
 *     (see AUTO_ADD_QUESTION_COLUMNS), so a new event's questions cannot go missing.
 *   - Normalises Luma's per-question-type answer shapes (LinkedIn/GitHub paths,
 *     multi-select arrays, the bundled company+job-title answer, terms booleans).
 *   - Optionally posts a Slack update (via an Incoming Webhook) when new
 *     registrations appear, including the new + total counts and event name/date.
 *
 * SETUP
 *   1. Requires a Luma Plus calendar (the API is gated behind it).
 *   2. In Apps Script: Project Settings -> Script Properties, add:
 *        LUMA_API_KEY        = <your Luma API key>
 *        LUMA_EVENT_ID       = <your event id, looks like "evt-XXXXXXXX">
 *        SLACK_WEBHOOK_URL   = <optional: Slack Incoming Webhook URL for updates>
 *        TOA_EVENT_SHEET_URL = <optional: when set, the Slack message links to this
 *                               URL as "View TOA Event Sheet" instead of to Luma>
 *        TOA_EVENT_NAME      = <optional: event name to show in the Slack message
 *                               instead of the name fetched from Luma>
 *   3. Edit the Config constants below to match your spreadsheet's sheet names
 *      and your event's question labels.
 *   4. Run syncLumaGuests once to authorise, then add a time-driven trigger
 *      (Triggers -> Add Trigger -> syncLumaGuests -> Time-driven) for auto-sync.
 *
 * No secrets are stored in this file.
 */

// ---- Config (edit to match your spreadsheet) ----
const SHEET_NAME  = 'Guests';       // destination sheet that receives the guest rows
const STAMP_SHEET = 'Dashboard';    // sheet to write the "last updated" timestamp on
const STAMP_CELL  = 'B1';           // cell to write the timestamp into
const STATUSES    = ['approved', 'pending_approval', 'waitlist', 'declined']; // add 'invited','session' if needed
const PAGE_LIMIT  = 100;

// When true, any registration question present in the Luma data that has no column
// in row 1 is appended as a new column at the end. Existing columns are never
// reordered or removed. This keeps the sheet complete when an event's questions
// change; set to false to freeze the layout to exactly what row 1 says.
const AUTO_ADD_QUESTION_COLUMNS = true;

// Luma's "company" question type returns BOTH the company and the job title in a
// single answer object — the job title has no answer entry of its own. The company
// question is detected automatically by its type, so no label needs configuring.
// This names the column that should receive that bundled job title when the event
// has no standalone job-title question. Matching ignores case and extra spaces, and
// any column whose name mentions "job title" or "role" is also accepted.
const JOB_TITLE_COLUMN = 'What is your job title?';

// Fallback column order — used ONLY if the destination sheet has no header row yet.
// Normally the script keeps whatever order already exists in row 1, so add your
// custom-question columns there (named as the Luma question label).
const COLUMNS = [
  'guest_id','name','first_name','last_name','email','phone_number','created_at',
  'approval_status','checked_in_at','utm_source','referrer','referred_by',
  'qr_code_url','amount','amount_tax',
  'amount_discount','currency','coupon_code','eth_address','solana_address',
  'survey_response_rating','survey_response_feedback','ticket_type_id','ticket_name'
];

const LUMA_API_BASE = 'https://public-api.luma.com/v1';

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('LUMA_API_KEY');
  const eventId = props.getProperty('LUMA_EVENT_ID');
  if (!apiKey)  throw new Error('Set LUMA_API_KEY in Project Settings -> Script Properties');
  if (!eventId) throw new Error('Set LUMA_EVENT_ID in Project Settings -> Script Properties');
  return { apiKey: apiKey, eventId: eventId };
}

function syncLumaGuests() {
  const cfg = getConfig();
  const slackWebhook = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');

  const all = [];
  STATUSES.forEach(function (s) {
    all.push.apply(all, fetchGuestsByStatus(cfg.apiKey, cfg.eventId, s));
  });

  // De-dupe guests that appear more than once
  const seen = {}, guests = [];
  all.forEach(function (g) {
    const id = g.api_id || g.user_email;
    if (id && !seen[id]) { seen[id] = true; guests.push(g); }
  });

  // Newest registration first
  guests.sort(function (a, b) {
    return String(b.created_at || b.registered_at || '')
      .localeCompare(String(a.created_at || a.registered_at || ''));
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  // Read guests already in the sheet BEFORE clearing, to detect newly-added ones
  const prev = existingState(sheet);
  let newCount = 0;
  guests.forEach(function (g) {
    const key = String(g[prev.keyField] || g.api_id || g.user_email || '');
    if (key && !prev.keys[key]) newCount++;
  });
  const total = guests.length;

  // Adopt the column order already in row 1; fall back to COLUMNS
  const lastCol = sheet.getLastColumn();
  let columns = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String) : [];
  if (!columns.length) columns = COLUMNS;

  columns = withMissingQuestionColumns(columns, guests);

  const rows = guests.map(function (g) {
    return columns.map(function (col) { return getValue(g, col); });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  if (rows.length) sheet.getRange(2, 1, rows.length, columns.length).setValues(rows);

  // Stamp the last-update time on the dashboard sheet
  const tz = ss.getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss z');
  const stampSheet = ss.getSheetByName(STAMP_SHEET) || ss.insertSheet(STAMP_SHEET);
  stampSheet.getRange(STAMP_CELL).setValue('Last updated: ' + stamp);

  Logger.log('Synced %s guests (%s new) in %s columns; stamped %s!%s',
    total, newCount, columns.length, STAMP_SHEET, STAMP_CELL);

  // Notify Slack only when new registrations appeared
  if (slackWebhook && newCount > 0) {
    postToSlack(slackWebhook, fetchEvent(cfg.apiKey, cfg.eventId), newCount, total);
  }
}

function getValue(g, col) {
  const t = (g.event_tickets && g.event_tickets[0]) || {};

  switch (col) {
    case 'guest_id':        return g.api_id || '';
    case 'name':            return g.user_name || g.name || '';
    case 'first_name':      return g.user_first_name || '';
    case 'last_name':       return g.user_last_name || '';
    case 'email':           return g.user_email || g.email || '';
    case 'phone_number':    return g.phone_number || '';
    case 'created_at':      return g.created_at || g.registered_at || '';
    case 'approval_status': return g.approval_status || '';
    case 'checked_in_at':   return g.checked_in_at || '';
    case 'utm_source':      return g.utm_source || '';
    case 'referrer':        return referrerOf(g);
    case 'referred_by':     return personName(g.referred_by);
    case 'qr_code_url':     return g.check_in_qr_code || '';
    case 'amount':          return money(t.amount);
    case 'amount_tax':      return money(t.amount_tax);
    case 'amount_discount': return money(t.amount_discount);
    case 'currency':        return t.currency || '';
    case 'coupon_code':     return g.coupon_code || '';                 // not returned by get-guests
    case 'eth_address':     return g.eth_address || '';
    case 'solana_address':  return g.solana_address || '';
    case 'survey_response_rating':   return g.survey_response_rating || '';   // not returned by get-guests
    case 'survey_response_feedback': return g.survey_response_feedback || ''; // not returned by get-guests
    case 'ticket_type_id':  return t.event_ticket_type_id || t.event_ticket_type_api_id || '';
    case 'ticket_name':     return t.name || '';
  }

  // Not a standard field, so it is a registration question, matched by label.
  const a = answerByLabel(g, col);
  if (a) {
    if (a.question_type === 'company') {
      if (a.answer_company != null) return a.answer_company;
      if (a.value && a.value.company != null) return a.value.company;
      return a.answer || '';
    }
    return formatAnswer(a);
  }

  // No question carries this label. It may be the job-title half of Luma's bundled
  // "company" question, which is not returned as an answer of its own.
  const c = companyAnswer(g);
  if (c && looksLikeJobTitle(col)) {
    if (c.answer_job_title != null) return c.answer_job_title;
    return (c.value && c.value.job_title) || '';
  }

  return '';
}

// Labels are compared loosely: case, repeated spaces and spaces before punctuation
// are ignored, so a Luma label like "Which  area will you focus on ?" still matches
// "Which area will you focus on?" typed in the sheet. Luma preserves whatever the
// host typed when creating the question, stray spaces included.
function normLabel(s) {
  return String(s == null ? '' : s)
    .replace(/\s+/g, ' ')
    .replace(/\s+([?!.:,])/g, '$1')
    .trim()
    .toLowerCase();
}

function answerByLabel(g, label) {
  const want = normLabel(label);
  if (!want) return undefined;
  const answers = g.registration_answers || [];
  for (let i = 0; i < answers.length; i++) {
    if (normLabel(answers[i].label) === want) return answers[i];
  }
  return undefined;
}

// Luma's bundled company question, found by type rather than by a configured label.
function companyAnswer(g) {
  return (g.registration_answers || []).filter(function (a) {
    return a.question_type === 'company';
  })[0];
}

function looksLikeJobTitle(col) {
  const n = normLabel(col);
  if (!n) return false;
  if (JOB_TITLE_COLUMN && n === normLabel(JOB_TITLE_COLUMN)) return true;
  return n.indexOf('job title') >= 0 || n.indexOf('role') >= 0;
}

// referred_by may be a plain string or a user object depending on the event.
function personName(v) {
  if (v == null) return '';
  if (typeof v === 'object') return v.name || v.email || v.api_id || '';
  return v;
}

// Luma's CSV export has a "referrer" column, but the guest object returned by the
// API has not always used that name — older payloads carried "custom_source"
// instead. Try each known candidate in order. If the column still comes out empty,
// run debugGuestFields() to see exactly which field holds the value.
const REFERRER_FIELDS = ['referrer', 'custom_source'];

function referrerOf(g) {
  for (let i = 0; i < REFERRER_FIELDS.length; i++) {
    const v = g[REFERRER_FIELDS[i]];
    if (v != null && v !== '') return (typeof v === 'object') ? personName(v) : v;
  }
  return '';
}

// Returns the header plus any registration question that has no column yet, so a
// new event's questions cannot silently go missing.
function withMissingQuestionColumns(columns, guests) {
  const known = {};
  columns.forEach(function (c) { known[normLabel(c)] = true; });

  const missing = [], seen = {};
  let hasCompanyQuestion = false;

  guests.forEach(function (g) {
    if (companyAnswer(g)) hasCompanyQuestion = true;
    (g.registration_answers || []).forEach(function (a) {
      const n = normLabel(a.label);
      if (!n || known[n] || seen[n]) return;
      seen[n] = true;
      missing.push(a.label);
    });
  });

  // The bundled job title has no answer entry, so check for its column separately.
  const needsJobTitleColumn = hasCompanyQuestion && JOB_TITLE_COLUMN &&
    !columns.some(looksLikeJobTitle) && !missing.some(looksLikeJobTitle);
  if (needsJobTitleColumn) missing.push(JOB_TITLE_COLUMN);

  if (!missing.length) return columns;

  Logger.log('Questions with no column in row 1: %s', missing.join(' | '));
  if (!AUTO_ADD_QUESTION_COLUMNS) {
    Logger.log('AUTO_ADD_QUESTION_COLUMNS is off — these will not be written.');
    return columns;
  }
  Logger.log('Appending %s new column(s) to the header.', missing.length);
  return columns.concat(missing);
}

function formatAnswer(a) {
  if (!a) return '';
  const v = (a.answer != null) ? a.answer : a.value;
  switch (a.question_type) {
    case 'linkedin':
      if (!v) return '';
      return /^https?:\/\//i.test(v) ? v : 'https://linkedin.com' + (v.charAt(0) === '/' ? v : '/' + v);
    case 'github':
      if (!v) return '';
      return /^https?:\/\//i.test(v) ? v : 'https://github.com/' + String(v).replace(/^\//, '');
    case 'multi-select': return Array.isArray(v) ? v.join(', ') : (v || '');
    case 'terms':        return v === true ? 'Agreed' : (v ? String(v) : '');
    default:
      if (Array.isArray(v)) return v.join(', ');
      if (v && typeof v === 'object') return JSON.stringify(v);
      return (v == null) ? '' : v;
  }
}

function money(cents) {
  if (cents == null) return '';
  return '$' + (Number(cents) / 100).toFixed(2); // Luma returns amounts in cents
}

function fetchGuestsByStatus(apiKey, eventId, status) {
  const out = [];
  let cursor = null, guard = 0;
  do {
    const url = LUMA_API_BASE + '/event/get-guests'
      + '?event_id=' + encodeURIComponent(eventId)
      + '&approval_status=' + encodeURIComponent(status)
      + '&pagination_limit=' + PAGE_LIMIT
      + (cursor ? '&pagination_cursor=' + encodeURIComponent(cursor) : '');

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'x-luma-api-key': apiKey, 'accept': 'application/json' },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log('Error %s for status %s: %s', res.getResponseCode(), status, res.getContentText());
      break;
    }

    const data = JSON.parse(res.getContentText());
    (data.entries || []).forEach(function (e) { out.push(e.guest || e); });
    cursor = data.has_more ? data.next_cursor : null;
  } while (cursor && ++guard < 100);

  return out;
}

// Reads the keys (guest_id, or email as fallback) already present in the sheet,
// so we can count how many of the freshly fetched guests are new since last sync.
function existingState(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { keyField: 'api_id', keys: {} };
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = header.indexOf('guest_id');
  let keyField = 'api_id';
  if (idx < 0) { idx = header.indexOf('email'); keyField = 'user_email'; }
  const keys = {};
  if (idx >= 0 && lastRow >= 2) {
    sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues().forEach(function (r) {
      if (r[0]) keys[String(r[0]).trim()] = true;
    });
  }
  return { keyField: keyField, keys: keys };
}

// Fetches the event so the Slack message can include its name and date.
function fetchEvent(apiKey, eventId) {
  const url = LUMA_API_BASE + '/event/get?id=' + encodeURIComponent(eventId);
  const res = UrlFetchApp.fetch(url, {
    headers: { 'x-luma-api-key': apiKey, 'accept': 'application/json' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('Event fetch error %s: %s', res.getResponseCode(), res.getContentText());
    return null;
  }
  const data = JSON.parse(res.getContentText());
  return data.event || data;
}

function formatEventDate(ev) {
  if (!ev || !ev.start_at) return '';
  const tz = ev.timezone || Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(ev.start_at), tz, "EEE, MMM d, yyyy 'at' h:mm a z");
}

// Posts a registration update to Slack via an Incoming Webhook.
// Returns the HTTP status code (200 = success).
function postToSlack(webhook, ev, newCount, total, isTest) {
  const props = PropertiesService.getScriptProperties();
  const name = props.getProperty('TOA_EVENT_NAME') || (ev && ev.name) || 'your event';
  const when = formatEventDate(ev);
  const sheetUrl = props.getProperty('TOA_EVENT_SHEET_URL');
  const lumaUrl = (ev && ev.url) || '';

  const lines = [];
  if (isTest) lines.push(':test_tube: _Test message from LumaSync2Gsheet — please ignore._');
  lines.push(':tada: *' + newCount + ' new registration' + (newCount === 1 ? '' : 's') + '* for *' + name + '*');
  if (when) lines.push(':calendar: ' + when);
  lines.push(':busts_in_silhouette: *' + total + '* total registered');
  if (sheetUrl) {
    lines.push('<' + sheetUrl + '|View TOA Event Sheet>');
  } else if (lumaUrl) {
    lines.push('<' + lumaUrl + '|View on Luma>');
  }

  const res = UrlFetchApp.fetch(webhook, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: lines.join('\n') }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code !== 200) Logger.log('Slack post error %s: %s', code, res.getContentText());
  return code;
}

// TEST: run this manually to verify the Slack webhook. It posts a clearly-labelled
// test message (real event name/date when Luma creds are set, with sample counts)
// and does NOT modify the spreadsheet.
function testSlackMessage() {
  const webhook = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhook) {
    throw new Error('SLACK_WEBHOOK_URL is not set — add it in Project Settings -> Script Properties.');
  }

  let ev = null;
  try {
    const cfg = getConfig();
    ev = fetchEvent(cfg.apiKey, cfg.eventId); // real name/date if Luma creds are configured
  } catch (e) {
    Logger.log('Test: skipping event lookup (%s)', e.message);
  }

  const code = postToSlack(webhook, ev, 3, 128, true); // sample counts
  if (code === 200) {
    Logger.log('Test message posted to Slack successfully (HTTP 200).');
  } else {
    throw new Error('Slack returned HTTP ' + code + ' — check that the webhook URL is correct and active.');
  }
}

// Run once to inspect the raw API response shape (helpful when adding columns).
// Tries each configured status so it still finds a sample when, for example, every
// guest is already approved.
function debugLumaResponse() {
  const cfg = getConfig();
  for (let i = 0; i < STATUSES.length; i++) {
    const url = LUMA_API_BASE + '/event/get-guests?event_id=' + encodeURIComponent(cfg.eventId)
      + '&approval_status=' + encodeURIComponent(STATUSES[i]) + '&pagination_limit=1';
    const res = UrlFetchApp.fetch(url, {
      headers: { 'x-luma-api-key': cfg.apiKey, 'accept': 'application/json' },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      Logger.log('Error %s for status %s: %s', res.getResponseCode(), STATUSES[i], res.getContentText());
      continue;
    }
    const data = JSON.parse(res.getContentText());
    if (data.entries && data.entries.length) {
      Logger.log('Sample guest (status: %s)\n%s', STATUSES[i], res.getContentText());
      return;
    }
  }
  Logger.log('No guests found in any of: %s', STATUSES.join(', '));
}

// Lists every top-level field on the guest objects, with how many guests have it
// populated and a sample value. Use this when a non-question column stays blank:
// it shows the exact API field name to map, and because it counts across all
// guests, a field that is only sometimes set cannot hide behind an empty sample.
function debugGuestFields() {
  const cfg = getConfig();
  const guests = [];
  STATUSES.forEach(function (s) {
    guests.push.apply(guests, fetchGuestsByStatus(cfg.apiKey, cfg.eventId, s));
  });
  if (!guests.length) { Logger.log('No guests found in any of: %s', STATUSES.join(', ')); return; }

  const skip = { registration_answers: 1, guest: 1, event_tickets: 1, event_ticket: 1 };
  const stats = {};
  guests.forEach(function (g) {
    Object.keys(g).forEach(function (k) {
      if (skip[k]) return;
      if (!stats[k]) stats[k] = { filled: 0, sample: '' };
      const v = g[k];
      if (v == null || v === '') return;
      stats[k].filled++;
      if (!stats[k].sample) {
        const s = (typeof v === 'object') ? JSON.stringify(v) : String(v);
        stats[k].sample = s.length > 80 ? s.slice(0, 80) + '...' : s;
      }
    });
  });

  Logger.log('Top-level guest fields across %s guests  —  field: filled/total  sample', guests.length);
  Object.keys(stats).sort().forEach(function (k) {
    Logger.log('  %s: %s/%s  %s', k, stats[k].filled, guests.length, stats[k].sample);
  });
  Logger.log('Referrer candidates tried, in order: %s', REFERRER_FIELDS.join(', '));
}

// Lists every registration question the event actually returns, with its type and
// the column the script would map it to. Run this whenever a column stays blank —
// it shows the exact label text to put in row 1.
function debugQuestionLabels() {
  const cfg = getConfig();
  const labels = {};
  STATUSES.forEach(function (s) {
    fetchGuestsByStatus(cfg.apiKey, cfg.eventId, s).forEach(function (g) {
      (g.registration_answers || []).forEach(function (a) {
        if (a.label && !labels[a.label]) labels[a.label] = a.question_type || '(unknown)';
      });
    });
  });

  const found = Object.keys(labels);
  if (!found.length) { Logger.log('No registration answers found.'); return; }

  Logger.log('Registration questions returned by Luma (copy these into row 1):');
  found.forEach(function (l) {
    Logger.log('  [%s]  %s%s', labels[l], l,
      labels[l] === 'company' ? '   <- also carries the job title' : '');
  });
  Logger.log('Job title column in use: %s', JOB_TITLE_COLUMN || '(none)');
}
