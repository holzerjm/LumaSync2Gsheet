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
 *     row; custom registration questions are matched by their exact label.
 *   - Normalises Luma's per-question-type answer shapes (LinkedIn/GitHub paths,
 *     multi-select arrays, the bundled company+job-title answer, terms booleans).
 *
 * SETUP
 *   1. Requires a Luma Plus calendar (the API is gated behind it).
 *   2. In Apps Script: Project Settings -> Script Properties, add:
 *        LUMA_API_KEY  = <your Luma API key>
 *        LUMA_EVENT_ID = <your event id, looks like "evt-XXXXXXXX">
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

// Luma's "company" question returns BOTH the company and the job title in a single
// answer object. Set these to your event's exact labels. Leave JOB_TITLE_COLUMN as
// '' if you don't keep job title in its own column.
const COMPANY_QUESTION_LABEL = 'What company do you work for?';
const JOB_TITLE_COLUMN       = 'What is your job title?';

// Fallback column order — used ONLY if the destination sheet has no header row yet.
// Normally the script keeps whatever order already exists in row 1, so add your
// custom-question columns there (named exactly as the Luma question label).
const COLUMNS = [
  'guest_id','name','first_name','last_name','email','phone_number','created_at',
  'approval_status','checked_in_at','utm_source','qr_code_url','amount','amount_tax',
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

  // Adopt the column order already in row 1; fall back to COLUMNS
  const lastCol = sheet.getLastColumn();
  let columns = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String) : [];
  if (!columns.length) columns = COLUMNS;

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

  Logger.log('Synced %s guests in %s columns; stamped %s!%s',
    rows.length, columns.length, STAMP_SHEET, STAMP_CELL);
}

function getValue(g, col) {
  const t = (g.event_tickets && g.event_tickets[0]) || {};

  // Luma bundles company + job title into one "company" answer
  if (JOB_TITLE_COLUMN && col === JOB_TITLE_COLUMN) {
    const c = answerByLabel(g, COMPANY_QUESTION_LABEL);
    return c ? (c.answer_job_title != null ? c.answer_job_title : ((c.value && c.value.job_title) || '')) : '';
  }
  if (col === COMPANY_QUESTION_LABEL) {
    const c = answerByLabel(g, COMPANY_QUESTION_LABEL);
    return c ? (c.answer_company != null ? c.answer_company : ((c.value && c.value.company) || c.answer || '')) : '';
  }

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
    case 'ticket_type_id':  return t.event_ticket_type_id || '';
    case 'ticket_name':     return t.name || '';
    default:
      return formatAnswer(answerByLabel(g, col)); // any custom registration question, matched by label
  }
}

function answerByLabel(g, label) {
  return (g.registration_answers || []).filter(function (a) { return a.label === label; })[0];
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

// Run once to inspect the raw API response shape (helpful when adding columns)
function debugLumaResponse() {
  const cfg = getConfig();
  const url = LUMA_API_BASE + '/event/get-guests?event_id=' + encodeURIComponent(cfg.eventId)
    + '&approval_status=pending_approval&pagination_limit=1';
  const res = UrlFetchApp.fetch(url, {
    headers: { 'x-luma-api-key': cfg.apiKey, 'accept': 'application/json' },
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}
