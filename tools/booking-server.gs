/**
 * Pulse: Scheduling (server side, Google Apps Script).
 *
 * VGCG's own booking pages (a Calendly replacement). Visitors pick a time on
 * the public page (book.html); the booking goes onto this script owner's
 * Google Calendar with a Zoom link and a VGCG-branded confirmation email.
 * Meeting types and availability are managed from the Scheduling page in Pulse.
 *
 * INSTALL (same project as the rest of the Pulse server):
 *   1. Extensions > Apps Script in the VGCG Client Hub sheet. Add a new script
 *      file named "booking-server" and paste this whole file in.
 *   2. In doPost, make this the FIRST line inside the function:
 *
 *          var bk = bookingRoute_(e); if (bk) return bk;
 *
 *      That answers the public booking calls (names starting "bkPub"), which
 *      need no sign-in, and leaves every other call to the existing code.
 *   3. Wherever the web app lists the functions it accepts (the same place
 *      getIntel was added), add: bkAdminGet, bkAdminSave, bkAdminCancel,
 *      bkAdminSetClient. They get the usual sign-in check.
 *   4. Deploy > Manage deployments > edit the current deployment > New version.
 *      Keep the same deployment so the URL doesn't change.
 *   5. The first run asks for Gmail (send mail) permission if the project
 *      doesn't have it yet; approve it.
 *
 * TIME ZONES: every instant is stored and passed as a UTC ISO string. Working
 * hours are wall-clock times in the host's IANA zone (default America/New_York)
 * and are converted to instants per date, so daylight-saving changes are
 * handled on the right day. The public page converts to the visitor's zone.
 *
 * DATA: settings and meeting types live in Script Properties (JSON). Bookings
 * are rows on a "Bookings" tab (created on first use).
 */

var BK_PROP = 'pulse.booking.v1';
var BK_SHEET = 'Bookings';
var BK_COLS = ['ID', 'Created', 'Type', 'Start', 'End', 'Name', 'Email', 'Organization', 'Notes',
  'Visitor time zone', 'Status', 'Event ID', 'Token', 'Client', 'Start UTC', 'End UTC', 'Updated'];

function bkDefaults_() {
  var wk = [['09:00', '17:00']];
  return {
    settings: {
      tz: 'America/New_York',
      hours: { '1': wk, '2': wk, '3': wk, '4': wk, '5': wk, '6': [], '7': [] }, // 1 = Monday … 7 = Sunday
      bufferMin: 15,        // gap kept clear before and after every meeting
      noticeHours: 24,      // earliest bookable time from now
      daysAhead: 45,        // how far out people can book
      dailyMax: 4,          // most booked meetings per day (0 = no cap)
      increment: 30,        // start times every N minutes
      blackout: [],         // ["2026-12-24", …] days off
      extraCalendars: [],   // other calendar IDs to treat as busy
      publicUrl: 'https://pulsedotdash.netlify.app/book',
      orgName: 'Victoria Gillespie Consulting Group',
      notifyEmail: '',      // blank = the script owner
      zoom: { phil: 'https://us06web.zoom.us/j/6581398657', victoria: 'https://us06web.zoom.us/j/9878011120' }
    },
    types: [
      { slug: 'discovery-call', name: 'Discovery Call', minutes: 30, zoom: 'phil', active: true,
        description: 'A short conversation about your organization and where we might help.',
        question: 'What would you like to talk about?' },
      { slug: 'strategy-session', name: 'Strategy Session', minutes: 60, zoom: 'phil', active: true,
        description: 'A working session on a specific strategy, search or communications need.',
        question: 'What should we prepare before we meet?' }
    ]
  };
}

/* ---------- routing ---------- */

function bookingRoute_(e) {
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return null; }
  var fn = String(body.fn || '');
  if (fn.indexOf('bkPub') !== 0) return null;
  var pub = { bkPubTypes: bkPubTypes, bkPubSlots: bkPubSlots, bkPubBook: bkPubBook, bkPubGet: bkPubGet,
    bkPubCancel: bkPubCancel, bkPubReschedule: bkPubReschedule };
  var out;
  try {
    if (!pub[fn]) throw new Error('Unknown request.');
    out = { ok: true, result: pub[fn].apply(null, body.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- config ---------- */

function bkConfig_() {
  var d = bkDefaults_(), raw = PropertiesService.getScriptProperties().getProperty(BK_PROP), c = {};
  try { c = raw ? JSON.parse(raw) : {}; } catch (x) { c = {}; }
  var s = c.settings || {}, out = {};
  Object.keys(d.settings).forEach(function (k) { out[k] = s[k] !== undefined ? s[k] : d.settings[k]; });
  return { settings: out, types: c.types && c.types.length ? c.types : d.types };
}

function bkType_(cfg, slug) {
  var t = cfg.types.filter(function (x) { return x.slug === slug && x.active; })[0];
  if (!t) throw new Error('That meeting type isn\'t available.');
  return t;
}

/* ---------- time zone helpers ---------- */

// Minutes east of UTC for zone tz at instant ms ("-0500" -> -300).
function bkOffset_(ms, tz) {
  var z = Utilities.formatDate(new Date(ms), tz, 'Z');
  var sign = z.charAt(0) === '-' ? -1 : 1;
  return sign * (Number(z.substr(1, 2)) * 60 + Number(z.substr(3, 2)));
}

// Wall-clock date ("YYYY-MM-DD") and time ("HH:MM") in tz -> UTC milliseconds.
function bkWallToUtc_(ymd, hm, tz) {
  var p = ymd.split('-'), t = hm.split(':');
  var guess = Date.UTC(+p[0], +p[1] - 1, +p[2], +t[0], +t[1]);
  var o1 = bkOffset_(guess, tz), ms = guess - o1 * 60000, o2 = bkOffset_(ms, tz);
  if (o2 !== o1) ms = guess - o2 * 60000;
  return ms;
}

function bkYmd_(ms, tz) { return Utilities.formatDate(new Date(ms), tz, 'yyyy-MM-dd'); }
function bkAddDays_(ymd, n) {
  var p = ymd.split('-'), d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.toISOString().slice(0, 10);
}
function bkIsoDow_(ymd) { var p = ymd.split('-'); var w = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay(); return w === 0 ? 7 : w; }
function bkMin_(hm) { var t = String(hm).split(':'); return +t[0] * 60 + +t[1]; }
function bkHm_(m) { return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }

/* ---------- availability ---------- */

// Busy intervals [startMs, endMs] from the calendar(s), all-day and declined events skipped.
function bkBusy_(s, fromMs, toMs, skipEventId) {
  var cals = [CalendarApp.getDefaultCalendar()];
  (s.extraCalendars || []).forEach(function (id) { try { var c = CalendarApp.getCalendarById(id); if (c) cals.push(c); } catch (x) {} });
  var out = [];
  cals.forEach(function (cal) {
    cal.getEvents(new Date(fromMs), new Date(toMs)).forEach(function (ev) {
      if (ev.isAllDayEvent()) return;
      if (skipEventId && ev.getId() === skipEventId) return;
      try { if (ev.getMyStatus() === CalendarApp.GuestStatus.NO) return; } catch (x) {}
      out.push([ev.getStartTime().getTime(), ev.getEndTime().getTime()]);
    });
  });
  return out;
}

// Core slot computation, kept free of Google services so it can be tested.
// busy: [[startMs,endMs]], perDay: {"YYYY-MM-DD": count of confirmed bookings (host zone)}
function bkComputeSlots_(s, minutes, fromYmd, toYmd, nowMs, busy, perDay) {
  var tz = s.tz, inc = Math.max(5, +s.increment || 30), buf = (+s.bufferMin || 0) * 60000, dur = minutes * 60000;
  var earliest = nowMs + (+s.noticeHours || 0) * 3600000;
  var latest = nowMs + (+s.daysAhead || 60) * 86400000;
  var black = {}; (s.blackout || []).forEach(function (d) { black[d] = 1; });
  var slots = [];
  for (var ymd = fromYmd, guard = 0; ymd <= toYmd && guard < 70; ymd = bkAddDays_(ymd, 1), guard++) {
    if (black[ymd]) continue;
    if (+s.dailyMax && (perDay[ymd] || 0) >= +s.dailyMax) continue;
    var wins = (s.hours || {})[String(bkIsoDow_(ymd))] || [];
    wins.forEach(function (w) {
      var a = bkMin_(w[0]), b = bkMin_(w[1]);
      for (var m = a; m + minutes <= b; m += inc) {
        var st = bkWallToUtc_(ymd, bkHm_(m), tz), en = st + dur;
        if (st < earliest || st > latest) continue;
        // a nonexistent wall time (spring-forward gap) maps to another hour; skip it
        if (Utilities.formatDate(new Date(st), tz, 'HH:mm') !== bkHm_(m)) continue;
        var clash = busy.some(function (x) { return st - buf < x[1] && en + buf > x[0]; });
        if (!clash) slots.push(new Date(st).toISOString());
      }
    });
  }
  return slots;
}

function bkSlotsFor_(cfg, type, fromYmd, toYmd, skip) {
  var s = cfg.settings, tz = s.tz, now = Date.now();
  var minY = bkYmd_(now, tz), maxY = bkYmd_(now + (+s.daysAhead || 60) * 86400000, tz);
  if (fromYmd < minY) fromYmd = minY;
  if (toYmd > maxY) toYmd = maxY;
  if (toYmd < fromYmd) return [];
  var fromMs = bkWallToUtc_(fromYmd, '00:00', tz) - 86400000, toMs = bkWallToUtc_(bkAddDays_(toYmd, 1), '00:00', tz) + 86400000;
  var busy = bkBusy_(s, fromMs, toMs, skip && skip.eventId);
  var perDay = {};
  bkRows_().forEach(function (r) {
    if (r.status !== 'confirmed' || (skip && r.id === skip.id)) return;
    var d = bkYmd_(Date.parse(r.startUtc), tz); perDay[d] = (perDay[d] || 0) + 1;
  });
  return bkComputeSlots_(s, +type.minutes, fromYmd, toYmd, now, busy, perDay);
}

/* ---------- bookings sheet ---------- */

function bkSheet_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(BK_SHEET);
  if (!sh) { sh = ss.insertSheet(BK_SHEET); sh.appendRow(BK_COLS); sh.setFrozenRows(1); }
  return sh;
}
function bkRows_() {
  var sh = bkSheet_(); if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, BK_COLS.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0]), created: r[1], type: String(r[2]), name: String(r[5]), email: String(r[6]),
      org: String(r[7]), notes: String(r[8]), visitorTz: String(r[9]), status: String(r[10]), eventId: String(r[11]),
      token: String(r[12]), client: String(r[13]), startUtc: String(r[14]), endUtc: String(r[15]) };
  });
}
function bkByToken_(token) {
  token = String(token || '');
  if (token.length < 20) throw new Error('That booking link isn\'t valid.');
  var r = bkRows_().filter(function (x) { return x.token === token; })[0];
  if (!r) throw new Error('That booking link isn\'t valid.');
  return r;
}
function bkWriteTimes_(r, cfg, startMs, endMs, status) {
  var tz = cfg.settings.tz, f = 'yyyy-MM-dd HH:mm';
  var sh = bkSheet_();
  sh.getRange(r.row, 4, 1, 2).setValues([[Utilities.formatDate(new Date(startMs), tz, f), Utilities.formatDate(new Date(endMs), tz, f)]]);
  sh.getRange(r.row, 11).setValue(status);
  sh.getRange(r.row, 15, 1, 3).setValues([[new Date(startMs).toISOString(), new Date(endMs).toISOString(), new Date()]]);
}

/* ---------- public calls ---------- */

function bkPubTypes() {
  var cfg = bkConfig_();
  return { org: cfg.settings.orgName, tz: cfg.settings.tz,
    types: cfg.types.filter(function (t) { return t.active; }).map(function (t) {
      return { slug: t.slug, name: t.name, minutes: +t.minutes, description: t.description || '', question: t.question || '' };
    }) };
}

function bkPubSlots(slug, fromYmd, toYmd, token) {
  var cfg = bkConfig_(), type = bkType_(cfg, slug), day = /^\d{4}-\d{2}-\d{2}$/;
  if (!day.test(fromYmd) || !day.test(toYmd)) throw new Error('Bad date range.');
  if (bkAddDays_(fromYmd, 45) < toYmd) toYmd = bkAddDays_(fromYmd, 45);
  var skip = null;
  if (token) { var r = bkByToken_(token); skip = { id: r.id, eventId: r.eventId }; }
  return { minutes: +type.minutes, slots: bkSlotsFor_(cfg, type, fromYmd, toYmd, skip) };
}

function bkPubBook(form) {
  form = form || {};
  if (form.website) throw new Error('Something went wrong. Please try again.'); // honeypot
  var name = String(form.name || '').trim().slice(0, 120), email = String(form.email || '').trim().toLowerCase().slice(0, 160);
  var org = String(form.org || '').trim().slice(0, 160), notes = String(form.notes || '').trim().slice(0, 2000);
  var vtz = String(form.tz || '').slice(0, 60);
  if (!name) throw new Error('Please enter your name.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('Please enter a valid email address.');
  var cache = CacheService.getScriptCache(), ck = 'bk.rate.' + email, n = +(cache.get(ck) || 0);
  if (n >= 3) throw new Error('You\'ve booked several times in the last hour. Please reply to your confirmation email instead.');

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var cfg = bkConfig_(), s = cfg.settings, type = bkType_(cfg, form.slug);
    var startMs = Date.parse(form.start);
    if (isNaN(startMs)) throw new Error('Please pick a time.');
    var ymd = bkYmd_(startMs, s.tz), iso = new Date(startMs).toISOString();
    if (bkSlotsFor_(cfg, type, ymd, ymd).indexOf(iso) < 0) throw new Error('Sorry, that time was just taken. Please pick another.');
    var endMs = startMs + type.minutes * 60000;
    var id = 'B' + Utilities.formatDate(new Date(), 'UTC', 'yyMMdd') + Math.random().toString(36).slice(2, 6).toUpperCase();
    var token = Utilities.getUuid().replace(/-/g, '') + Math.random().toString(36).slice(2, 10);
    var manage = s.publicUrl.replace(/\/$/, '') + '?b=' + token;
    var zoom = (s.zoom || {})[type.zoom] || '';
    var title = type.name + ': ' + name + (org ? ' (' + org + ')' : '');
    var desc = type.name + ' booked through VGCG scheduling.\n\nName: ' + name + '\nEmail: ' + email + (org ? '\nOrganization: ' + org : '') +
      (notes ? '\n\n' + (type.question || 'Notes') + '\n' + notes : '') + (zoom ? '\n\nJoin Zoom Meeting\n' + zoom : '') +
      '\n\nNeed to reschedule or cancel? ' + manage;
    var ev = CalendarApp.getDefaultCalendar().createEvent(title, new Date(startMs), new Date(endMs),
      { description: desc, location: zoom, guests: email, sendInvites: true });
    try { ev.setTag('pulseBooking', id); } catch (x) {}
    bkSheet_().appendRow([id, new Date(), type.slug, '', '', name, email, org, notes, vtz, 'confirmed', ev.getId(), token, '', '', '', '']);
    var r = bkRows_().filter(function (x) { return x.id === id; })[0];
    bkWriteTimes_(r, cfg, startMs, endMs, 'confirmed');
    cache.put(ck, String(n + 1), 3600);
    var b = { id: id, type: type, name: name, email: email, org: org, notes: notes, vtz: vtz, startMs: startMs, endMs: endMs, zoom: zoom, manage: manage };
    bkMail_(cfg, 'confirmed', b);
    return { id: id, token: token, start: iso, end: new Date(endMs).toISOString(), type: type.name, minutes: +type.minutes, zoom: zoom };
  } finally { lock.releaseLock(); }
}

function bkPublicView_(cfg, r) {
  var t = cfg.types.filter(function (x) { return x.slug === r.type; })[0] || { name: r.type, minutes: 0 };
  return { id: r.id, status: r.status, type: t.name, slug: r.type, minutes: +t.minutes, start: r.startUtc, end: r.endUtc,
    name: r.name, zoom: r.status === 'confirmed' ? ((cfg.settings.zoom || {})[t.zoom] || '') : '',
    past: Date.parse(r.startUtc) < Date.now() };
}

function bkPubGet(token) { var cfg = bkConfig_(); return bkPublicView_(cfg, bkByToken_(token)); }

function bkPubCancel(token, reason) {
  var cfg = bkConfig_(), r = bkByToken_(token);
  if (r.status !== 'confirmed') throw new Error('This meeting is already cancelled.');
  if (Date.parse(r.startUtc) < Date.now()) throw new Error('This meeting has already happened.');
  bkCancel_(cfg, r, String(reason || '').slice(0, 500), 'guest');
  return bkPublicView_(cfg, bkByToken_(token));
}

function bkCancel_(cfg, r, reason, by) {
  try { var ev = CalendarApp.getDefaultCalendar().getEventById(r.eventId); if (ev) ev.deleteEvent(); } catch (x) {}
  bkSheet_().getRange(r.row, 11).setValue('cancelled');
  bkSheet_().getRange(r.row, 17).setValue(new Date());
  var t = cfg.types.filter(function (x) { return x.slug === r.type; })[0] || { name: r.type, minutes: 0 };
  bkMail_(cfg, 'cancelled', { id: r.id, type: t, name: r.name, email: r.email, org: r.org, vtz: r.visitorTz,
    startMs: Date.parse(r.startUtc), endMs: Date.parse(r.endUtc), reason: reason, by: by,
    manage: cfg.settings.publicUrl.replace(/\/$/, '') + '/' + r.type });
}

function bkPubReschedule(token, newStart) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var cfg = bkConfig_(), s = cfg.settings, r = bkByToken_(token);
    if (r.status !== 'confirmed') throw new Error('This meeting was cancelled. Please book a new time.');
    var type = bkType_(cfg, r.type), startMs = Date.parse(newStart);
    if (isNaN(startMs)) throw new Error('Please pick a time.');
    var ymd = bkYmd_(startMs, s.tz), iso = new Date(startMs).toISOString();
    if (bkSlotsFor_(cfg, type, ymd, ymd, { id: r.id, eventId: r.eventId }).indexOf(iso) < 0) throw new Error('Sorry, that time was just taken. Please pick another.');
    var endMs = startMs + type.minutes * 60000, oldStart = Date.parse(r.startUtc);
    var ev = CalendarApp.getDefaultCalendar().getEventById(r.eventId);
    if (ev) ev.setTime(new Date(startMs), new Date(endMs));
    bkWriteTimes_(r, cfg, startMs, endMs, 'confirmed');
    var manage = s.publicUrl.replace(/\/$/, '') + '?b=' + token;
    bkMail_(cfg, 'rescheduled', { id: r.id, type: type, name: r.name, email: r.email, org: r.org, notes: r.notes, vtz: r.visitorTz,
      startMs: startMs, endMs: endMs, oldStartMs: oldStart, zoom: (s.zoom || {})[type.zoom] || '', manage: manage });
    return bkPublicView_(cfg, bkByToken_(token));
  } finally { lock.releaseLock(); }
}

/* ---------- admin calls (signed in through Pulse) ---------- */

function bkAdminGet() {
  var cfg = bkConfig_(), cutoff = Date.now() - 120 * 86400000;
  return { settings: cfg.settings, types: cfg.types, defaults: bkDefaults_(),
    bookings: bkRows_().filter(function (r) { return Date.parse(r.startUtc) >= cutoff; }).map(function (r) {
      return { id: r.id, type: r.type, name: r.name, email: r.email, org: r.org, notes: r.notes, visitorTz: r.visitorTz,
        status: r.status, client: r.client, start: r.startUtc, end: r.endUtc,
        created: r.created instanceof Date ? r.created.toISOString() : String(r.created) };
    }) };
}

function bkAdminSave(settings, types) {
  var d = bkDefaults_().settings, s = {}, hm = /^([01]\d|2[0-3]):[0-5]\d$/;
  settings = settings || {};
  Object.keys(d).forEach(function (k) { s[k] = settings[k] !== undefined ? settings[k] : d[k]; });
  try { Utilities.formatDate(new Date(), s.tz, 'Z'); } catch (x) { throw new Error('Unknown time zone.'); }
  Object.keys(s.hours || {}).forEach(function (k) {
    (s.hours[k] || []).forEach(function (w) { if (!hm.test(w[0]) || !hm.test(w[1]) || w[1] <= w[0]) throw new Error('Check the working hours: each block needs a start before its end.'); });
  });
  ['bufferMin', 'noticeHours', 'daysAhead', 'dailyMax', 'increment'].forEach(function (k) { s[k] = Math.max(0, Math.round(+s[k] || 0)); });
  if (s.increment < 5) s.increment = 5;
  if (s.daysAhead < 1) s.daysAhead = 1;
  var seen = {};
  types = (types || []).map(function (t) {
    var slug = String(t.slug || t.name || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
    if (!slug || !String(t.name || '').trim()) throw new Error('Every meeting type needs a name.');
    if (seen[slug]) throw new Error('Two meeting types share the link "' + slug + '".');
    seen[slug] = 1;
    var min = Math.round(+t.minutes);
    if (!(min >= 5 && min <= 480)) throw new Error('"' + t.name + '" needs a length between 5 and 480 minutes.');
    return { slug: slug, name: String(t.name).trim(), minutes: min, zoom: ['phil', 'victoria', 'none'].indexOf(t.zoom) >= 0 ? t.zoom : 'phil',
      active: !!t.active, description: String(t.description || '').slice(0, 600), question: String(t.question || '').slice(0, 200) };
  });
  if (!types.length) throw new Error('Keep at least one meeting type.');
  PropertiesService.getScriptProperties().setProperty(BK_PROP, JSON.stringify({ settings: s, types: types }));
  return bkAdminGet();
}

function bkAdminCancel(id, reason) {
  var cfg = bkConfig_(), r = bkRows_().filter(function (x) { return x.id === String(id); })[0];
  if (!r) throw new Error('Booking not found.');
  if (r.status !== 'confirmed') throw new Error('Already cancelled.');
  bkCancel_(cfg, r, String(reason || '').slice(0, 500), 'host');
  return bkAdminGet();
}

function bkAdminSetClient(id, clientId) {
  var r = bkRows_().filter(function (x) { return x.id === String(id); })[0];
  if (!r) throw new Error('Booking not found.');
  bkSheet_().getRange(r.row, 14).setValue(String(clientId || ''));
  return true;
}

/* ---------- email ---------- */

function bkFmt_(ms, tz) {
  return Utilities.formatDate(new Date(ms), tz, "EEEE, MMMM d, yyyy 'at' h:mm a") + ' ' + bkTzLabel_(ms, tz);
}
function bkTzLabel_(ms, tz) { var z = Utilities.formatDate(new Date(ms), tz, 'z'); return /^GMT|^UTC/.test(z) ? tz.replace(/_/g, ' ') : z; }
function bkEsc_(t) { return String(t || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

function bkMail_(cfg, kind, b) {
  var s = cfg.settings, gtz = b.vtz || s.tz, logo = s.publicUrl.replace(/\/book\/?$/, '') + '/vgcg-mark.png';
  try { Utilities.formatDate(new Date(), gtz, 'Z'); } catch (x) { gtz = s.tz; }
  var when = bkFmt_(b.startMs, gtz), whenHost = bkFmt_(b.startMs, s.tz);
  var head = { confirmed: 'You\'re booked', rescheduled: 'Your meeting has moved', cancelled: 'Your meeting is cancelled' }[kind];
  var lines = [];
  lines.push('<p style="margin:0 0 6px;font-size:13px;color:#67626e;text-transform:uppercase;letter-spacing:.08em">' + bkEsc_(b.type.name) + ' · ' + b.type.minutes + ' minutes</p>');
  lines.push('<p style="margin:0 0 18px;font-size:20px;font-weight:600;color:#2b2730">' + bkEsc_(when) + '</p>');
  if (kind === 'rescheduled' && b.oldStartMs) lines.push('<p style="margin:0 0 16px;color:#67626e">Previously ' + bkEsc_(bkFmt_(b.oldStartMs, gtz)) + '</p>');
  if (kind !== 'cancelled' && b.zoom) lines.push('<p style="margin:0 0 18px"><a href="' + bkEsc_(b.zoom) + '" style="display:inline-block;background:#0077a6;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:600">Join on Zoom</a></p>');
  if (kind === 'cancelled' && b.reason) lines.push('<p style="margin:0 0 16px;color:#2b2730">Note: ' + bkEsc_(b.reason) + '</p>');
  if (kind === 'cancelled') lines.push('<p style="margin:0 0 16px"><a href="' + bkEsc_(b.manage) + '" style="color:#0077a6">Book a new time</a></p>');
  else lines.push('<p style="margin:0 0 16px;color:#67626e;font-size:14px">A calendar invitation is on its way separately. Need a different time? <a href="' + bkEsc_(b.manage) + '" style="color:#0077a6">Reschedule or cancel</a>.</p>');
  var html = '<div style="background:#f3f7f8;padding:28px 12px;font-family:Helvetica,Arial,sans-serif">' +
    '<div style="max-width:520px;margin:0 auto;background:#fff;border:1px solid #dbe5e8;border-radius:14px;padding:28px">' +
    '<img src="' + bkEsc_(logo) + '" width="44" height="42" alt="VGCG" style="display:block;margin-bottom:18px">' +
    '<h1 style="margin:0 0 18px;font-size:22px;color:#2b2730">' + head + (b.name ? ', ' + bkEsc_(b.name.split(' ')[0]) : '') + '</h1>' +
    lines.join('') +
    '<p style="margin:22px 0 0;padding-top:16px;border-top:1px solid #dbe5e8;font-size:12px;color:#67626e">' + bkEsc_(s.orgName) + '</p>' +
    '</div></div>';
  var subj = { confirmed: 'Confirmed: ', rescheduled: 'Updated: ', cancelled: 'Cancelled: ' }[kind] + b.type.name + ' on ' +
    Utilities.formatDate(new Date(b.startMs), gtz, 'MMM d') ;
  var host = s.notifyEmail || Session.getEffectiveUser().getEmail();
  try {
    MailApp.sendEmail({ to: b.email, subject: subj, htmlBody: html, name: s.orgName, replyTo: host,
      body: head + '\n' + b.type.name + '\n' + when + (b.zoom && kind !== 'cancelled' ? '\nZoom: ' + b.zoom : '') + '\n' + (b.manage || '') });
  } catch (x) {}
  if (b.by === 'host') return; // host cancelled it themselves
  try {
    MailApp.sendEmail({ to: host, subject: '[Booking] ' + subj + ' with ' + b.name,
      body: kind.toUpperCase() + ': ' + b.type.name + '\n' + whenHost + '\n\n' + b.name + ' <' + b.email + '>' + (b.org ? '\n' + b.org : '') +
        (b.vtz ? '\nTheir time zone: ' + b.vtz : '') + (b.notes ? '\n\n' + b.notes : '') + (b.reason ? '\n\nReason: ' + b.reason : '') +
        '\n\nBooking ' + b.id + ' · manage it on the Scheduling page in Pulse.' });
  } catch (x) {}
}
