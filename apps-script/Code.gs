/**
 * KovsieX Experience — staff portal (Apps Script web app, backed by this Google Sheet)
 *
 * Roles:  Staff       → own dashboard, log hours, request leave / submit excuses
 *         Supervisor  → + management view, first-stage approval of requests
 *         Manager     → + final approval, staff accounts, settings (the "admin")
 *
 * Setup: open the sheet → Extensions → Apps Script → paste this file and Index.html,
 * reload the sheet, then run  KovsieX → 1. Set up sheets & first Manager.
 */

const SS_ID = '';            // leave blank when the script is bound to the sheet
const SESSION_SECONDS = 21600; // 6 hours (Apps Script cache maximum)
const ROLES = ['Staff', 'Supervisor', 'Manager'];
// Radio only for now — add 'TV/Video', 'Online & Social' back here when those areas join.
const AREAS = ['Radio', 'All areas'];
const MAX_SHOW_HOURS = 6;
const REQUEST_TYPES = ['Leave', 'Excuse'];
const MAX_FAILED_LOGINS = 5;

const SHEETS = {
  Users: ['Email', 'Name', 'Role', 'Area', 'Salt', 'PasswordHash', 'Active', 'WeeklyTarget',
          'MustChangePassword', 'CreatedAt', 'LastLogin',
          'Status', 'Phone', 'StudentNumber', 'StartDate', 'SuspendedUntil', 'LeftOn'],
  Hours: ['ID', 'Email', 'Date', 'Hours', 'Area', 'Activity', 'LoggedAt', 'Status'],
  Requests: ['ID', 'Email', 'Name', 'Type', 'StartDate', 'EndDate', 'Reason', 'Status',
             'SupervisorBy', 'SupervisorAt', 'ManagerBy', 'ManagerAt', 'Comment', 'CreatedAt'],
  Bookings: ['ID', 'Email', 'Name', 'Date', 'Start', 'End', 'Show', 'With', 'WithEmails', 'Permanent',
             'Status', 'SkipDates', 'NoShowDates', 'EndDate', 'CreatedAt', 'Image'],
  HR: ['ID', 'Email', 'Name', 'Type', 'Details', 'Date', 'By', 'CreatedAt'],
  Notifications: ['ID', 'Email', 'Type', 'Title', 'Body', 'Page', 'Read', 'CreatedAt'],
  Reviews: ['ID', 'BookingID', 'Date', 'Show', 'HostEmail', 'Emails', 'Music', 'Mic', 'Engagement', 'Imaging',
            'Comments', 'By', 'CreatedAt'],
  Listener: ['ID', 'Name', 'Type', 'Message', 'Source', 'Email', 'Verified', 'Status', 'Pinned', 'CreatedAt'],
  Played: ['ID', 'Title', 'Artist', 'Art', 'PlayedAt', 'By'],
  Settings: ['Key', 'Value']
};
const LISTENER_TYPES = ['Request', 'Shoutout', 'Traffic', 'Joke', 'Competition', 'Other'];
const DEFAULT_LINKS = 'Welcome package | \nHow-to guide | \nTraffic (SANRAL i-traffic) | https://www.i-traffic.co.za\nWeather (SA Weather Service) | https://www.weathersa.co.za';
// Columns kept as plain text so Sheets never reformats them
const TEXT_COLS = {
  Users: ['Email', 'Salt', 'PasswordHash', 'CreatedAt', 'LastLogin', 'Phone', 'StudentNumber',
          'StartDate', 'SuspendedUntil', 'LeftOn'],
  HR: ['ID', 'Date', 'CreatedAt'],
  Notifications: ['ID', 'Email', 'Page', 'CreatedAt'],
  Reviews: ['ID', 'BookingID', 'Date', 'HostEmail', 'Emails', 'CreatedAt'],
  Listener: ['ID', 'Email', 'CreatedAt'],
  Played: ['ID', 'PlayedAt'],
  Hours: ['ID', 'Date', 'LoggedAt'],
  Requests: ['ID', 'StartDate', 'EndDate', 'SupervisorAt', 'ManagerAt', 'CreatedAt'],
  Bookings: ['ID', 'Date', 'Start', 'End', 'WithEmails', 'SkipDates', 'NoShowDates', 'EndDate', 'CreatedAt', 'Image'],
  Settings: ['Key', 'Value']
};

/* ───────────────────────── Web app & sheet menu ───────────────────────── */

function doGet(e) {
  const p = e && e.parameter && e.parameter.p;
  if (p === 'listen') {
    return HtmlService.createHtmlOutputFromFile('Public')
      .setTitle('KovsieX Radio')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('KovsieX Experience')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('KovsieX')
    .addItem('1. Set up sheets & first Manager', 'setup')
    .addSeparator()
    .addItem('Add staff member', 'menuAddStaff')
    .addItem('Reset a password', 'menuResetPassword')
    .addToUi();
}

function setup() {
  const ss = ss_();
  const ui = SpreadsheetApp.getUi();
  Object.keys(SHEETS).forEach(function (name) { ensureSheet_(ss, name); });
  const s = settings_();
  if (!s.termStart) setSetting_('TermStart', monday_(today_()));
  if (!s.defaultTarget) setSetting_('DefaultWeeklyTarget', '4');

  const hasManager = rows_('Users').some(function (u) { return u.Role === 'Manager'; });
  if (hasManager) { ui.alert('Sheets are ready. A Manager account already exists.'); return; }

  const e = ui.prompt('First Manager account', 'Manager email address:', ui.ButtonSet.OK_CANCEL);
  if (e.getSelectedButton() !== ui.Button.OK) return;
  const n = ui.prompt('First Manager account', 'Manager full name:', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  const pw = createUser_({ email: e.getResponseText(), name: n.getResponseText(), role: 'Manager', area: 'All areas' });
  ui.alert('Manager created.\n\nEmail: ' + e.getResponseText().trim().toLowerCase() +
           '\nTemporary password: ' + pw + '\n\nCopy it now — it is not stored anywhere readable. ' +
           'They will be asked to choose a new password at first login.');
}

function menuAddStaff() {
  const ui = SpreadsheetApp.getUi();
  const ask = function (q) {
    const r = ui.prompt('Add staff member', q, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) throw new Error('Cancelled');
    return r.getResponseText();
  };
  try {
    const email = ask('Email address:');
    const name = ask('Full name:');
    const role = ask('Role (' + ROLES.join(' / ') + '):') || 'Staff';
    const area = ask('Area (' + AREAS.join(' / ') + '):');
    const pw = createUser_({ email: email, name: name, role: role, area: area });
    ui.alert('Account created.\n\nEmail: ' + email.trim().toLowerCase() + '\nTemporary password: ' + pw);
  } catch (err) { if (err.message !== 'Cancelled') ui.alert('Could not add: ' + err.message); }
}

function menuResetPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Reset password', 'Email address of the account:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try { ui.alert('New temporary password: ' + resetPassword_(r.getResponseText())); }
  catch (err) { ui.alert(err.message); }
}

/* ───────────────────────── Auth ───────────────────────── */

function login(email, password) {
  email = String(email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + email;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILED_LOGINS) throw new Error('Too many attempts. Try again in 15 minutes or ask a Manager to reset your password.');

  const u = findUser_(email);
  if (!u || hash_(password, u.Salt) !== u.PasswordHash) {
    cache.put(failKey, String(fails + 1), 900);
    throw new Error('Email or password is incorrect.');
  }
  liftIfDue_(u);
  if (u.Status === 'Terminated') throw new Error('This account is no longer active. Please speak to a Manager.');
  if (u.Status !== 'Suspended' && !isTrue_(u.Active)) throw new Error('This account is inactive. Please speak to a Manager.');
  cache.remove(failKey);

  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('s_' + token, email, SESSION_SECONDS);
  setCells_('Users', u._row, { LastLogin: now_() });
  return { token: token, user: publicUser_(u), mustChange: isTrue_(u.MustChangePassword) };
}

function logout(token) { CacheService.getScriptCache().remove('s_' + token); return true; }

function changePassword(token, current, next) {
  const u = auth_(token, null, true);
  if (hash_(current, u.Salt) !== u.PasswordHash) throw new Error('Current password is incorrect.');
  next = String(next || '');
  if (next.length < 8) throw new Error('New password must be at least 8 characters.');
  if (next === current) throw new Error('Choose a password different from the current one.');
  const salt = Utilities.getUuid();
  setCells_('Users', u._row, { Salt: salt, PasswordHash: hash_(next, salt), MustChangePassword: false });
  return true;
}

/* ───────────────────────── Staff dashboard ───────────────────────── */

function getDashboard(token) {
  const u = auth_(token, null, true);
  const s = settings_();
  const users = rows_('Users');
  const hours = rows_('Hours');
  const reqs = rows_('Requests');
  const books = rows_('Bookings');
  const today = today_(), now = now_();
  const stats = stats_(u, hours, reqs, books, s);

  const mine = function (o) { return o.emails.indexOf(u.Email) >= 0; };
  const sched = schedule_(books, addDays_(today, -28), addDays_(today, s.bookingWindow));
  const upcoming = sched.filter(function (o) { return mine(o) && o.endAt > now; });
  const recent = sched.filter(function (o) { return mine(o) && o.endAt <= now; }).reverse().slice(0, 15);
  const extra = hours.filter(function (h) { return h.Email === u.Email && h.Status !== 'Removed'; })
    .sort(byDesc_('Date')).slice(0, 10)
    .map(function (h) { return { date: h.Date, hours: Number(h.Hours), activity: h.Activity }; });
  const studio = sched.filter(function (o) { return o.date >= today; })
    .map(function (o) { return { id: o.id, date: o.date, start: o.start, end: o.end, show: o.show, name: o.name,
                                 with: o.with, permanent: o.permanent }; });
  const staffList = users.filter(function (x) { return isTrue_(x.Active) && x.Email !== u.Email; })
    .map(function (x) { return { email: x.Email, name: x.Name }; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
  const myReqs = reqs.filter(function (r) { return r.Email === u.Email; })
    .sort(byDesc_('CreatedAt')).slice(0, 25).map(publicRequest_);

  const warnings = rows_('HR').filter(function (h) { return h.Email === u.Email && h.Type === 'Warning'; })
    .sort(byDesc_('CreatedAt')).map(function (h) { return { date: h.Date, details: h.Details, by: h.By }; });
  const reviews = rows_('Reviews').filter(function (r) { return list_(r.Emails).indexOf(u.Email) >= 0; })
    .sort(byDesc_('CreatedAt')).slice(0, 30).map(publicReview_);

  return { user: publicUser_(u), stats: stats, live: live_(books), upcoming: upcoming.map(pubOcc_(u)),
           recent: recent.map(pubOcc_(u)), extraHours: extra, studio: studio, staffList: staffList,
           requests: myReqs, today: today, maxDate: addDays_(today, s.bookingWindow),
           bookingWindow: s.bookingWindow, warnings: warnings, maxWarnings: s.maxWarnings,
           reviews: reviews, links: s.links, unread: unreadCount_(u.Email), publicUrl: publicUrl_() };
}

/* ───────────────────────── Booking studio time ───────────────────────── */

function bookSlot(token, o) {
  const u = auth_(token);
  const s = settings_();
  const today = today_();
  const date = cleanDate_(o.date);
  const start = cleanTime_(o.start), end = cleanTime_(o.end);
  if (date < today) throw new Error('You cannot book a date in the past.');
  if (date > addDays_(today, s.bookingWindow))
    throw new Error('You can only book up to ' + s.bookingWindow + ' days ahead.');
  if (end <= start) throw new Error('End time must be after start time.');
  const h = hoursBetween_(start, end);
  if (h < 0.5 || h > MAX_SHOW_HOURS) throw new Error('A booking must be between 30 minutes and ' + MAX_SHOW_HOURS + ' hours.');
  if (date + ' ' + start <= now_()) throw new Error('That start time has already passed.');
  const show = String(o.show || '').trim();
  if (!show) throw new Error('Enter the show name.');

  const users = rows_('Users');
  const withEmails = (o.withEmails || []).map(function (e) { return String(e).trim().toLowerCase(); })
    .filter(function (e, i, a) { return e && e !== u.Email && a.indexOf(e) === i; });
  const names = withEmails.map(function (e) {
    const x = users.filter(function (v) { return v.Email === e && isTrue_(v.Active); })[0];
    if (!x) throw new Error('One of the co-hosts is not an active staff member.');
    return x.Name;
  });
  const guests = String(o.withOther || '').trim();
  if (guests) names.push(guests);
  const permanent = !!o.permanent;

  return withLock_(function () {
    const clash = findClash_(rows_('Bookings'), date, start, end, permanent, null);
    if (clash) throw new Error('The studio is already booked ' + fmtClash_(clash) + '. Pick another time.');
    const image = o.image ? saveImage_(o.image, show) : '';
    append_('Bookings', { ID: id_(), Email: u.Email, Name: u.Name, Date: date, Start: start, End: end, Image: image,
      Show: show.slice(0, 120), With: names.join(', ').slice(0, 300), WithEmails: withEmails.join(','),
      Permanent: permanent, Status: 'Booked', SkipDates: '', NoShowDates: '', EndDate: '', CreatedAt: now_() });
    withEmails.forEach(function (e) {
      notify_(e, 'Booking', u.Name + ' added you to a show',
        show + ' · ' + date + ' ' + start + '–' + end + (permanent ? ' (every week)' : ''), 'shows');
    });
    return getDashboard(token);
  });
}

/** scope: 'one' = just this date, 'series' = this date and every week after (permanent shows) */
function cancelBooking(token, id, date, scope) {
  const u = auth_(token);
  const isMgmt = u.Role !== 'Staff';
  withLock_(function () {
    const b = rows_('Bookings').filter(function (x) { return x.ID === id; })[0];
    if (!b) throw new Error('Booking not found.');
    if (b.Email !== u.Email && !isMgmt) throw new Error('Only the person who booked the show can cancel it.');
    date = cleanDate_(date);
    if (!isMgmt && date + ' ' + b.Start <= now_()) throw new Error('This show has already started.');
    if (!isTrue_(b.Permanent) || (scope === 'series' && date <= b.Date)) {
      setCells_('Bookings', b._row, { Status: 'Cancelled' });
    } else if (scope === 'series') {
      setCells_('Bookings', b._row, { EndDate: addDays_(date, -1) });
    } else {
      setCells_('Bookings', b._row, { SkipDates: list_(b.SkipDates).concat([date]).join(',') });
    }
  });
  return getDashboard(token);
}

/** Change the show avatar on a booking (host, or management) */
function setBookingImage(token, id, dataUrl) {
  const u = auth_(token);
  const b = rows_('Bookings').filter(function (x) { return x.ID === id; })[0];
  if (!b) throw new Error('Booking not found.');
  if (b.Email !== u.Email && u.Role === 'Staff') throw new Error('Only the host can change the show avatar.');
  setCells_('Bookings', b._row, { Image: saveImage_(dataUrl, b.Show) });
  return getDashboard(token);
}

function managerCancelBooking(token, id, date, scope) {
  const me = auth_(token, ['Supervisor', 'Manager']);
  const b = rows_('Bookings').filter(function (x) { return x.ID === id; })[0];
  cancelBooking(token, id, date, scope);
  if (b && b.Email !== me.Email) notify_(b.Email, 'Booking', 'Your show was cancelled by ' + me.Name,
    b.Show + ' · ' + date + (scope === 'series' ? ' and every week after' : ''), 'shows');
  return getManagement(token);
}

function setNoShow(token, id, date, noShow) {
  auth_(token, ['Supervisor', 'Manager']);
  withLock_(function () {
    const b = rows_('Bookings').filter(function (x) { return x.ID === id; })[0];
    if (!b) throw new Error('Booking not found.');
    let l = list_(b.NoShowDates).filter(function (d) { return d !== date; });
    if (noShow) l.push(cleanDate_(date));
    setCells_('Bookings', b._row, { NoShowDates: l.join(',') });
    if (noShow) [b.Email].concat(list_(b.WithEmails)).forEach(function (e) {
      notify_(e, 'Booking', 'Show marked as a no-show', b.Show + ' on ' + date + ' — its hours were removed.', 'shows');
    });
  });
  return getManagement(token);
}

function deleteHours(token, id) {
  const u = auth_(token);
  const h = rows_('Hours').filter(function (r) { return r.ID === id; })[0];
  if (!h) throw new Error('Entry not found.');
  const isMgmt = u.Role !== 'Staff';
  if (!isMgmt && (h.Email !== u.Email || h.Date < addDays_(today_(), -7)))
    throw new Error('You can only remove your own entries from the last 7 days.');
  setCells_('Hours', h._row, { Status: 'Removed' });
  return true;
}

function submitRequest(token, req) {
  const u = auth_(token);
  if (REQUEST_TYPES.indexOf(req.type) < 0) throw new Error('Choose Leave or Excuse.');
  const start = cleanDate_(req.start);
  const end = cleanDate_(req.end || req.start);
  if (end < start) throw new Error('End date is before start date.');
  const reason = String(req.reason || '').trim();
  if (reason.length < 5) throw new Error('Please give a short reason.');
  withLock_(function () {
    append_('Requests', { ID: id_(), Email: u.Email, Name: u.Name, Type: req.type, StartDate: start,
      EndDate: end, Reason: reason.slice(0, 500),
      Status: u.Role === 'Staff' ? 'Pending Supervisor' : 'Pending Manager', CreatedAt: now_() });
  });
  notifyRole_(u.Role === 'Staff' ? ['Supervisor', 'Manager'] : ['Manager'], u.Email, 'Leave',
    'New ' + req.type.toLowerCase() + ' request from ' + u.Name, start + (end !== start ? ' to ' + end : '') + ' · ' + reason.slice(0, 120), 'm:requests');
  return getDashboard(token);
}

function cancelRequest(token, id) {
  const u = auth_(token);
  const r = rows_('Requests').filter(function (x) { return x.ID === id && x.Email === u.Email; })[0];
  if (!r) throw new Error('Request not found.');
  if (r.Status.indexOf('Pending') !== 0) throw new Error('Only pending requests can be cancelled.');
  setCells_('Requests', r._row, { Status: 'Cancelled' });
  return getDashboard(token);
}

/* ───────────────────────── Management view ───────────────────────── */

function getManagement(token) {
  const me = auth_(token, ['Supervisor', 'Manager']);
  const s = settings_();
  const users = rows_('Users');
  const hours = rows_('Hours');
  const reqs = rows_('Requests');
  const books = rows_('Bookings');
  const hr = rows_('HR');
  const names = {}, warnings = {};
  users.forEach(function (u) { names[u.Email] = u.Name; });
  hr.forEach(function (h) { if (h.Type === 'Warning') warnings[h.Email] = (warnings[h.Email] || 0) + 1; });

  const staff = users.map(function (u) {
    const st = stats_(u, hours, reqs, books, s);
    return { email: u.Email, name: u.Name, role: u.Role, area: u.Area, active: isTrue_(u.Active),
      status: statusOf_(u), phone: u.Phone, studentNumber: u.StudentNumber, startDate: u.StartDate,
      suspendedUntil: u.SuspendedUntil, leftOn: u.LeftOn, warnings: warnings[u.Email] || 0,
      target: Number(u.WeeklyTarget) || s.defaultTarget, lastLogin: u.LastLogin,
      week: st.thisWeek, pastShort: st.pastShort, termWorked: st.termWorked,
      weeksMet: st.weeksMet, pastWeeks: st.pastWeeks };
  }).sort(function (a, b) { return a.name.localeCompare(b.name); });

  const today = today_();
  const reviewed = {};
  rows_('Reviews').forEach(function (r) { reviewed[r.BookingID + '|' + r.Date] = true; });
  const schedule = schedule_(books, addDays_(today, -14), addDays_(today, s.bookingWindow)).map(pubOcc_(me))
    .map(function (o) { o.reviewed = !!reviewed[o.id + '|' + o.date]; return o; });

  const requests = reqs.sort(byDesc_('CreatedAt')).map(function (r) {
    const p = publicRequest_(r);
    p.canAct = canDecide_(me, r);
    return p;
  });
  const recentHours = hours.filter(function (h) { return h.Status !== 'Removed'; })
    .sort(byDesc_('Date')).slice(0, 150).map(function (h) {
      return { id: h.ID, email: h.Email, name: names[h.Email] || h.Email, date: h.Date,
               hours: Number(h.Hours), area: h.Area, activity: h.Activity };
    });
  const hrLog = me.Role === 'Manager' ? hr.sort(byDesc_('CreatedAt')).map(function (h) {
    return { email: h.Email, name: h.Name, type: h.Type, details: h.Details, date: h.Date, by: h.By, created: h.CreatedAt };
  }) : [];
  return { me: publicUser_(me), staff: staff, requests: requests, hours: recentHours, schedule: schedule,
           live: live_(books), hr: hrLog, hrTypes: ['Warning', 'Note'],
           settings: s, roles: ROLES, areas: AREAS, isManager: me.Role === 'Manager', today: today };
}

function decideRequest(token, id, decision, comment) {
  const me = auth_(token, ['Supervisor', 'Manager']);
  return withLock_(function () {
    const r = rows_('Requests').filter(function (x) { return x.ID === id; })[0];
    if (!r) throw new Error('Request not found.');
    if (!canDecide_(me, r)) throw new Error('You cannot act on this request at its current stage.');
    const upd = { Comment: String(comment || '').slice(0, 300) || r.Comment };
    if (decision === 'Decline') {
      upd.Status = 'Declined';
      if (r.Status === 'Pending Supervisor') { upd.SupervisorBy = me.Name; upd.SupervisorAt = now_(); }
      else { upd.ManagerBy = me.Name; upd.ManagerAt = now_(); }
    } else if (decision === 'Approve') {
      if (r.Status === 'Pending Supervisor' && me.Role === 'Supervisor') {
        upd.Status = 'Pending Manager'; upd.SupervisorBy = me.Name; upd.SupervisorAt = now_();
      } else {
        upd.Status = 'Approved'; upd.ManagerBy = me.Name; upd.ManagerAt = now_();
        if (r.Status === 'Pending Supervisor') upd.SupervisorBy = '(skipped by Manager)';
      }
    } else throw new Error('Unknown decision.');
    setCells_('Requests', r._row, upd);
    const when = r.StartDate + (r.EndDate && r.EndDate !== r.StartDate ? ' to ' + r.EndDate : '');
    if (upd.Status === 'Approved' || upd.Status === 'Declined') {
      notify_(r.Email, 'Leave', 'Your ' + String(r.Type).toLowerCase() + ' was ' + upd.Status.toLowerCase(),
        when + ' · by ' + me.Name + (upd.Comment ? ' — ' + upd.Comment : ''), 'leave');
    } else if (upd.Status === 'Pending Manager') {
      notify_(r.Email, 'Leave', 'Your ' + String(r.Type).toLowerCase() + ' was approved by your Supervisor',
        when + ' · now waiting for a Manager', 'leave');
      notifyRole_(['Manager'], me.Email, 'Leave', r.Name + '’s ' + String(r.Type).toLowerCase() + ' needs your approval',
        when + ' · approved by ' + me.Name, 'm:requests');
    }
    return getManagement(token);
  });
}

function managerAddHours(token, entry) {
  auth_(token, ['Manager']);
  const u = findUser_(entry.email);
  if (!u) throw new Error('No staff member with that email.');
  const h = Number(entry.hours);
  if (!(h > 0 && h <= 24)) throw new Error('Hours must be between 0 and 24.');
  append_('Hours', { ID: id_(), Email: u.Email, Date: cleanDate_(entry.date), Hours: h,
    Area: entry.area || u.Area, Activity: String(entry.activity || 'Added by Manager').slice(0, 300),
    LoggedAt: now_(), Status: 'Logged' });
  return getManagement(token);
}

function removeHours(token, id) {
  auth_(token, ['Supervisor', 'Manager']);
  deleteHours(token, id);
  return getManagement(token);
}

/* Accounts (Manager only) */

function addUser(token, o) {
  const me = auth_(token, ['Manager']);
  const pw = withLock_(function () { return createUser_(o); });
  hrLog_(findUser_(o.email), 'Hired', 'Added as ' + (o.role || 'Staff') + ' (' + (o.area || AREAS[0]) + ')',
         o.startDate || today_(), me.Name);
  return { password: pw, data: getManagement(token) };
}

function resetUserPassword(token, email) {
  auth_(token, ['Manager']);
  return { password: resetPassword_(email) };
}

function updateUser(token, o) {
  const me = auth_(token, ['Manager']);
  const u = findUser_(o.email);
  if (!u) throw new Error('User not found.');
  if (ROLES.indexOf(o.role) < 0) throw new Error('Invalid role.');
  if (AREAS.indexOf(o.area) < 0) throw new Error('Invalid area.');
  if (u.Email === me.Email && o.role !== 'Manager') throw new Error('You cannot demote your own account.');
  const t = Number(o.target);
  const start = o.startDate ? cleanDate_(o.startDate) : '';
  setCells_('Users', u._row, { Role: o.role, Area: o.area, WeeklyTarget: t > 0 ? t : '',
    Name: String(o.name || u.Name).trim(), Phone: String(o.phone || '').trim().slice(0, 30),
    StudentNumber: String(o.studentNumber || '').trim().slice(0, 30), StartDate: start });
  if (o.role !== u.Role) hrLog_(u, 'Role change', u.Role + ' → ' + o.role, today_(), me.Name);
  return getManagement(token);
}

/**
 * HR actions (Manager only):
 *  Warning / Note  → recorded on the person's HR file
 *  Suspend         → can't sign in until reinstated (or until the 'until' date passes)
 *  Reinstate       → active again
 *  Terminate       → removed from the team: sign-in blocked, future bookings and pending requests cancelled.
 *                    Their past hours and HR file are kept.
 */
function hrAction(token, o) {
  const me = auth_(token, ['Manager']);
  const action = String(o.action || '');
  const details = String(o.details || '').trim().slice(0, 1000);
  withLock_(function () {
    const u = findUser_(o.email);
    if (!u) throw new Error('Person not found.');
    const date = o.date ? cleanDate_(o.date) : today_();
    if (u.Email === me.Email && ['Suspend', 'Terminate'].indexOf(action) >= 0)
      throw new Error('You cannot suspend or remove your own account.');
    if (action === 'Warning' || action === 'Note') {
      if (!details) throw new Error('Write the details for the ' + action.toLowerCase() + '.');
      hrLog_(u, action, details, date, me.Name);
      if (action === 'Warning') {
        const max = settings_().maxWarnings;
        const count = rows_('HR').filter(function (h) { return h.Email === u.Email && h.Type === 'Warning'; }).length;
        notify_(u.Email, 'Warning', 'You have received a written warning (' + count + ' of ' + max + ')',
          details.slice(0, 200) + (count >= max ? ' — you have reached the warning limit.' :
          ' — ' + (max - count) + ' more and you will be removed from the team.'), 'shows');
      }
    } else if (action === 'Suspend') {
      if (!details) throw new Error('Give a reason for the suspension.');
      const until = o.until ? cleanDate_(o.until) : '';
      if (until && until < date) throw new Error('The end date is before the start date.');
      setCells_('Users', u._row, { Active: false, Status: 'Suspended', SuspendedUntil: until });
      const n = cancelRange_(u.Email, date, until);
      hrLog_(u, 'Suspended', details + (until ? ' (until ' + until + ')' : '') + (n ? ' · ' + n + ' show(s) unbooked' : ''), date, me.Name);
      notify_(u.Email, 'Suspension', 'Your account has been suspended' + (until ? ' until ' + until : ''),
        details.slice(0, 200) + ' — you can still view your dashboard but can’t make changes.' +
        (n ? ' ' + n + ' booked show(s) in that time were unbooked.' : ''), 'dash');
    } else if (action === 'Reinstate') {
      setCells_('Users', u._row, { Active: true, Status: 'Active', SuspendedUntil: '', LeftOn: '' });
      hrLog_(u, 'Reinstated', details || 'Returned to the team', date, me.Name);
      notify_(u.Email, 'Suspension', 'Your account is active again', details || 'Welcome back!', 'dash');
    } else if (action === 'Terminate') {
      if (!details) throw new Error('Give a reason for removing this person.');
      setCells_('Users', u._row, { Active: false, Status: 'Terminated', LeftOn: date, SuspendedUntil: '' });
      const cancelled = cancelFutureFor_(u.Email, date);
      hrLog_(u, 'Terminated', details + (cancelled ? ' · ' + cancelled + ' future booking(s) cancelled' : ''), date, me.Name);
      endSessions_(u.Email);
    } else throw new Error('Unknown action.');
  });
  return getManagement(token);
}

function getLive(token) { auth_(token, null, true); return live_(rows_('Bookings')); }
function getPulse(token) {
  const u = auth_(token, null, true);
  return { live: live_(rows_('Bookings')), unread: unreadCount_(u.Email) };
}

/* ───────────────────────── Notifications ───────────────────────── */

function getNotifications(token) {
  const u = auth_(token, null, true);
  const list = rows_('Notifications').filter(function (n) { return n.Email === u.Email; })
    .sort(byDesc_('CreatedAt')).slice(0, 40).map(function (n) {
      return { id: n.ID, type: n.Type, title: n.Title, body: n.Body, page: n.Page, read: isTrue_(n.Read), created: n.CreatedAt };
    });
  return { list: list, unread: list.filter(function (n) { return !n.read; }).length };
}
function markNotificationsRead(token) {
  const u = auth_(token, null, true);
  withLock_(function () {
    rows_('Notifications').forEach(function (n) {
      if (n.Email === u.Email && !isTrue_(n.Read)) setCells_('Notifications', n._row, { Read: true });
    });
  });
  return true;
}
function notify_(email, type, title, body, page) {
  if (!email) return;
  append_('Notifications', { ID: id_(), Email: String(email).toLowerCase(), Type: type, Title: String(title).slice(0, 160),
    Body: String(body || '').slice(0, 400), Page: page || '', Read: false, CreatedAt: now_() });
}
function notifyRole_(roles, exceptEmail, type, title, body, page) {
  rows_('Users').forEach(function (x) {
    if (roles.indexOf(x.Role) >= 0 && isTrue_(x.Active) && x.Email !== exceptEmail) notify_(x.Email, type, title, body, page);
  });
}
function unreadCount_(email) {
  return rows_('Notifications').filter(function (n) { return n.Email === email && !isTrue_(n.Read); }).length;
}

/* ───────────────────────── Show reviews (audits) ───────────────────────── */

function reviewShow(token, o) {
  const me = auth_(token, ['Supervisor', 'Manager']);
  const b = rows_('Bookings').filter(function (x) { return x.ID === o.id; })[0];
  if (!b) throw new Error('Show not found.');
  const date = cleanDate_(o.date);
  const score = function (v) { const n = Number(v); if (!(n >= 0 && n <= 5)) throw new Error('Scores must be 0 to 5.'); return Math.round(n); };
  const emails = [b.Email].concat(list_(b.WithEmails));
  withLock_(function () {
    append_('Reviews', { ID: id_(), BookingID: b.ID, Date: date, Show: b.Show, HostEmail: b.Email, Emails: emails.join(','),
      Music: score(o.music), Mic: score(o.mic), Engagement: score(o.engagement), Imaging: score(o.imaging),
      Comments: String(o.comments || '').trim().slice(0, 3000), By: me.Name, CreatedAt: now_() });
  });
  emails.forEach(function (e) {
    notify_(e, 'Review', 'Your show was reviewed by ' + me.Name, b.Show + ' · ' + date + ' ' + b.Start + '–' + b.End, 'shows');
  });
  return getManagement(token);
}
function publicReview_(r) {
  return { id: r.ID, bookingId: r.BookingID, date: r.Date, show: r.Show, music: Number(r.Music), mic: Number(r.Mic),
    engagement: Number(r.Engagement), imaging: Number(r.Imaging), comments: r.Comments, by: r.By, created: r.CreatedAt };
}

function saveSettings(token, o) {
  auth_(token, ['Manager']);
  setSetting_('TermStart', monday_(cleanDate_(o.termStart)));
  const t = Number(o.defaultTarget);
  if (!(t > 0 && t <= 40)) throw new Error('Weekly target must be between 1 and 40.');
  setSetting_('DefaultWeeklyTarget', String(t));
  const w = Number(o.bookingWindow);
  if (!(w >= 1 && w <= 90)) throw new Error('Booking window must be between 1 and 90 days.');
  setSetting_('BookingWindowDays', String(Math.round(w)));
  if (o.maxWarnings !== undefined) {
    const mw = Number(o.maxWarnings);
    if (!(mw >= 1 && mw <= 10)) throw new Error('Warning limit must be between 1 and 10.');
    setSetting_('MaxWarnings', String(Math.round(mw)));
  }
  if (o.streamUrl !== undefined) {
    const su = String(o.streamUrl || '').trim();
    if (su && su.indexOf('https:') !== 0) throw new Error('The stream link must start with https:');
    setSetting_('StreamUrl', su);
  }
  if (o.links !== undefined) { parseLinks_(o.links); setSetting_('Links', String(o.links).slice(0, 3000)); }
  if (o.ionoStream !== undefined) { setSetting_('IonoStream', ionoId_(o.ionoStream)); CacheService.getScriptCache().remove('iono'); }
  return getManagement(token);
}

/* ───────────────────────── Hours maths ───────────────────────── */

/**
 * Worked hours = booked shows that have finished airing (minus no-shows, for the host and co-hosts)
 *              + any hours a Manager added by hand (Hours sheet).
 * Booked = shows still to come this week.
 */
function stats_(user, hoursAll, reqAll, booksAll, s) {
  const target = Number(user.WeeklyTarget) || s.defaultTarget;
  const thisMon = monday_(today_());
  const now = now_();
  let startMon = monday_(s.termStart || thisMon);
  if (startMon > thisMon) startMon = thisMon;

  const worked = {}, booked = {};
  hoursAll.forEach(function (h) {
    if (h.Email !== user.Email || h.Status === 'Removed' || !h.Date) return;
    const w = monday_(h.Date);
    worked[w] = (worked[w] || 0) + (Number(h.Hours) || 0);
  });
  schedule_(booksAll, startMon, addDays_(thisMon, 6)).forEach(function (o) {
    if (o.emails.indexOf(user.Email) < 0 || o.noShow) return;
    const w = monday_(o.date);
    if (o.endAt <= now) worked[w] = (worked[w] || 0) + o.hours;
    else booked[w] = (booked[w] || 0) + o.hours;
  });
  const leave = reqAll.filter(function (r) { return r.Email === user.Email && r.Status === 'Approved'; });

  const weeks = [];
  for (let w = startMon; w <= thisMon && weeks.length < 60; w = addDays_(w, 7)) {
    let off = 0;
    for (let i = 0; i < 5; i++) {
      const d = addDays_(w, i);
      if (leave.some(function (r) { return r.StartDate <= d && d <= (r.EndDate || r.StartDate); })) off++;
    }
    const t = Math.round(target * (5 - off) / 5 * 2) / 2;
    const wk = round_(worked[w] || 0), bk = round_(booked[w] || 0);
    weeks.push({ week: w, worked: wk, booked: bk, target: t, leaveDays: off,
                 short: Math.max(0, round_(t - wk)), unbooked: Math.max(0, round_(t - wk - bk)) });
  }
  const cur = weeks[weeks.length - 1];
  const past = weeks.slice(0, -1);
  return {
    target: target,
    thisWeek: cur,
    termStart: startMon,
    termWorked: round_(weeks.reduce(function (a, w) { return a + w.worked; }, 0)),
    termExpected: round_(weeks.reduce(function (a, w) { return a + w.target; }, 0)),
    pastShort: round_(past.reduce(function (a, w) { return a + w.short; }, 0)),
    weeksMet: past.filter(function (w) { return w.short === 0; }).length,
    pastWeeks: past.length,
    recentWeeks: weeks.slice(-8)
  };
}

/* Every show occurrence between two dates (permanent shows expanded week by week) */
function schedule_(books, from, to) {
  const out = [];
  books.forEach(function (b) {
    if (b.Status === 'Cancelled' || !b.Date) return;
    const noShows = list_(b.NoShowDates);
    const emails = [b.Email].concat(list_(b.WithEmails));
    const h = hoursBetween_(b.Start, b.End);
    occurrences_(b, from, to).forEach(function (d) {
      out.push({ id: b.ID, date: d, start: b.Start, end: b.End, endAt: d + ' ' + b.End,
        show: b.Show, name: b.Name, owner: b.Email, with: b.With, emails: emails,
        permanent: isTrue_(b.Permanent), hours: h, noShow: noShows.indexOf(d) >= 0, image: b.Image || '' });
    });
  });
  return out.sort(function (a, b) { return (a.date + a.start).localeCompare(b.date + b.start); });
}
function occurrences_(b, from, to) {
  const skip = list_(b.SkipDates);
  if (!isTrue_(b.Permanent)) return (b.Date >= from && b.Date <= to && skip.indexOf(b.Date) < 0) ? [b.Date] : [];
  let d = b.Date;
  if (d < from) d = addDays_(d, Math.ceil((d_(from) - d_(d)) / 864e5 / 7) * 7);
  const stop = (b.EndDate && b.EndDate < to) ? b.EndDate : to;
  const out = [];
  for (; d <= stop; d = addDays_(d, 7)) if (skip.indexOf(d) < 0) out.push(d);
  return out;
}
function findClash_(books, date, start, end, permanent, ignoreId) {
  const dates = [date];
  if (permanent) for (let k = 1; k < 26; k++) dates.push(addDays_(date, 7 * k));
  for (let i = 0; i < dates.length; i++) {
    const hit = schedule_(books.filter(function (b) { return b.ID !== ignoreId; }), dates[i], dates[i])
      .filter(function (o) { return o.start < end && start < o.end; })[0];
    if (hit) return hit;
  }
  return null;
}
function fmtClash_(o) {
  return 'on ' + o.date + ' from ' + o.start + ' to ' + o.end + ' (' + o.show + ' with ' + o.name +
         (o.permanent ? ', a permanent show' : '') + ')';
}
function pubOcc_(viewer) {
  return function (o) {
    return { id: o.id, date: o.date, start: o.start, end: o.end, show: o.show, name: o.name, with: o.with,
      permanent: o.permanent, hours: o.hours, noShow: o.noShow, past: o.endAt <= now_(), hasImage: !!o.image,
      isOwner: o.owner === viewer.Email };
  };
}

/* ───────────────────────── Helpers ───────────────────────── */

function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
const _checked = {};
function sheet_(name) {
  const ss = ss_();
  if (!SHEETS[name]) throw new Error('Sheet "' + name + '" is missing.');
  if (_checked[name]) return _checked[name];
  // Creates missing tabs (e.g. Bookings, HR) and adds any new columns to existing ones
  return (_checked[name] = ensureSheet_(ss, name));
}
function ensureSheet_(ss, name) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
  } else {
    const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    const missing = SHEETS[name].filter(function (c) { return head.indexOf(c) < 0; });
    if (missing.length) sh.getRange(1, head.length + 1, 1, missing.length).setValues([missing]);
  }
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8eaf6');
  (TEXT_COLS[name] || []).forEach(function (col) {
    const i = head.indexOf(col) + 1;
    if (i > 0) sh.getRange(1, i, sh.getMaxRows(), 1).setNumberFormat('@');
  });
  return sh;
}

/* Who is on air right now, and what's next */
function live_(books) {
  const today = today_(), now = now_(), tz = tz_();
  const ms = function (d, t) { return Utilities.parseDate(d + ' ' + t, tz, 'yyyy-MM-dd HH:mm').getTime(); };
  const occ = schedule_(books, today, addDays_(today, 7)).filter(function (o) { return !o.noShow; });
  const pub = function (o) {
    return { show: o.show, name: o.name, with: o.with, date: o.date, start: o.start, end: o.end,
             startMs: ms(o.date, o.start), endMs: ms(o.date, o.end), permanent: o.permanent,
             image: imageData_(o.image) };
  };
  const onAir = occ.filter(function (o) { return o.date + ' ' + o.start <= now && now < o.endAt; }).map(pub);
  const next = occ.filter(function (o) { return o.date + ' ' + o.start > now; })[0];
  return { onAir: onAir, next: next ? pub(next) : null, serverNow: Date.now() };
}

/* Show avatars: saved privately in a Drive folder the script owns, served to the page as data, so no public links */
function saveImage_(dataUrl, label) {
  const m = /^data:(image\/(jpeg|png|webp));base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('The avatar must be a JPG, PNG or WebP image.');
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 1500000) throw new Error('That image is too large. Please pick a smaller one.');
  const blob = Utilities.newBlob(bytes, m[1], 'show-' + String(label || 'avatar').replace(/[^\w-]+/g, '_').slice(0, 40) + '-' + id_() + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2]));
  return imageFolder_().createFile(blob).getId();
}
function imageFolder_() {
  const id = settingsRaw_('ImageFolderId');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('KovsieX Experience – Show Avatars');
  setSetting_('ImageFolderId', f.getId());
  return f;
}
function imageData_(id) {
  if (!id) return '';
  const cache = CacheService.getScriptCache();
  const hit = cache.get('img_' + id);
  if (hit) return hit;
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    const data = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    if (data.length < 95000) cache.put('img_' + id, data, 21600);
    return data;
  } catch (e) { return ''; }
}
function settingsRaw_(key) {
  const r = rows_('Settings').filter(function (x) { return x.Key === key; })[0];
  return r ? String(r.Value) : '';
}

function statusOf_(u) { return u.Status || (isTrue_(u.Active) ? 'Active' : 'Suspended'); }
function hrLog_(u, type, details, date, by) {
  if (!u) return;
  append_('HR', { ID: id_(), Email: u.Email, Name: u.Name, Type: type, Details: details,
                  Date: date, By: by, CreatedAt: now_() });
}
function endSessions_(email) { /* sessions expire on their own; auth_ already blocks inactive users */ }
function cancelFutureFor_(email, fromDate) {
  let n = 0;
  rows_('Bookings').forEach(function (b) {
    if (b.Email !== email || b.Status === 'Cancelled') return;
    if (isTrue_(b.Permanent)) {
      if (b.Date >= fromDate) { setCells_('Bookings', b._row, { Status: 'Cancelled' }); n++; }
      else if (!b.EndDate || b.EndDate >= fromDate) { setCells_('Bookings', b._row, { EndDate: addDays_(fromDate, -1) }); n++; }
    } else if (b.Date >= fromDate) { setCells_('Bookings', b._row, { Status: 'Cancelled' }); n++; }
  });
  rows_('Requests').forEach(function (r) {
    if (r.Email === email && String(r.Status).indexOf('Pending') === 0) setCells_('Requests', r._row, { Status: 'Cancelled' });
  });
  return n;
}
function list_(v) { return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(String); }
function cleanTime_(v) {
  const s = String(v || '').trim();
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(s)) throw new Error('Please enter a valid time.');
  return s;
}
function hoursBetween_(a, b) {
  const m = function (t) { const p = String(t).split(':').map(Number); return p[0] * 60 + p[1]; };
  return round_((m(b) - m(a)) / 60);
}
function tz_() { return ss_().getSpreadsheetTimeZone() || 'Africa/Johannesburg'; }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
function id_() { return Utilities.getUuid().slice(0, 8); }
function round_(n) { return Math.round(n * 100) / 100; }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function byDesc_(k) { return function (a, b) { return String(b[k]).localeCompare(String(a[k])); }; }

function d_(ymd) { const p = ymd.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2])); }
function ymd_(d) { return d.toISOString().slice(0, 10); }
function addDays_(ymd, n) { const d = d_(ymd); d.setUTCDate(d.getUTCDate() + n); return ymd_(d); }
function monday_(ymd) { const d = d_(ymd); d.setUTCDate(d.getUTCDate() - (d.getUTCDay() + 6) % 7); return ymd_(d); }
function cleanDate_(v) {
  const s = String(v || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || isNaN(d_(s).getTime())) throw new Error('Please enter a valid date.');
  return s;
}

function rows_(name) {
  const sh = sheet_(name);
  const v = sh.getDataRange().getValues();
  const head = v.shift() || [];
  const tz = tz_();
  return v.filter(function (r) { return r.join('') !== ''; }).map(function (r, i) {
    const o = { _row: i + 2 };
    head.forEach(function (k, j) {
      let x = r[j];
      if (x instanceof Date) {
        x = (x.getHours() || x.getMinutes()) ? Utilities.formatDate(x, tz, 'yyyy-MM-dd HH:mm')
                                             : Utilities.formatDate(x, tz, 'yyyy-MM-dd');
      }
      o[k] = (k === 'Email') ? String(x).trim().toLowerCase() : x;
    });
    return o;
  });
}
function header_(name) { const sh = sheet_(name); return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }
function append_(name, obj) {
  sheet_(name).appendRow(header_(name).map(function (k) { return obj[k] === undefined ? '' : obj[k]; }));
}
function setCells_(name, row, obj) {
  const sh = sheet_(name), head = header_(name);
  Object.keys(obj).forEach(function (k) {
    const c = head.indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(obj[k]);
  });
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function settings_() {
  const m = {};
  try { rows_('Settings').forEach(function (r) { m[r.Key] = r.Value; }); } catch (e) {}
  const linksRaw = m.Links !== undefined && m.Links !== '' ? String(m.Links) : DEFAULT_LINKS;
  return { termStart: m.TermStart ? String(m.TermStart) : '', defaultTarget: Number(m.DefaultWeeklyTarget) || 4,
           bookingWindow: Number(m.BookingWindowDays) || 21, maxWarnings: Number(m.MaxWarnings) || 3,
           streamUrl: m.StreamUrl ? String(m.StreamUrl) : '', linksRaw: linksRaw, links: parseLinks_(linksRaw),
           ionoStream: m.IonoStream ? String(m.IonoStream) : '' };
}
function parseLinks_(raw) {
  return String(raw || '').split('\n').map(function (line) {
    const i = line.indexOf('|');
    const label = (i >= 0 ? line.slice(0, i) : line).trim();
    const url = (i >= 0 ? line.slice(i + 1) : '').trim();
    if (url && url.indexOf('https:') !== 0 && url.indexOf('http:') !== 0) throw new Error('Link for "' + label + '" must start with https:');
    return { label: label, url: url };
  }).filter(function (l) { return l.label; }).slice(0, 12);
}
function setSetting_(key, value) {
  const r = rows_('Settings').filter(function (x) { return x.Key === key; })[0];
  if (r) setCells_('Settings', r._row, { Value: value }); else append_('Settings', { Key: key, Value: value });
}

function findUser_(email) {
  email = String(email || '').trim().toLowerCase();
  return rows_('Users').filter(function (u) { return u.Email === email; })[0];
}
/** readOnly: suspended people may view (dashboard, notifications) but nothing else */
function auth_(token, roles, readOnly) {
  const email = token && CacheService.getScriptCache().get('s_' + token);
  if (!email) throw new Error('SESSION_EXPIRED');
  const u = findUser_(email);
  if (!u) throw new Error('SESSION_EXPIRED');
  liftIfDue_(u);
  if (u.Status === 'Terminated') throw new Error('SESSION_EXPIRED');
  if (u.Status === 'Suspended') {
    if (!readOnly || roles) throw new Error('Your account is suspended' + (u.SuspendedUntil ? ' until ' + u.SuspendedUntil : '') +
      '. You can view your dashboard but can’t make changes.');
  } else if (!isTrue_(u.Active)) throw new Error('SESSION_EXPIRED');
  if (roles && roles.indexOf(u.Role) < 0) throw new Error('You do not have access to this.');
  return u;
}
function liftIfDue_(u) {
  if (u.Status === 'Suspended' && u.SuspendedUntil && today_() > u.SuspendedUntil) {
    setCells_('Users', u._row, { Active: true, Status: 'Active', SuspendedUntil: '' });
    hrLog_(u, 'Reinstated', 'Suspension ended automatically', today_(), 'System');
    notify_(u.Email, 'Suspension', 'Your suspension has ended', 'Your account is fully active again.', 'dash');
    u.Active = true; u.Status = 'Active'; u.SuspendedUntil = '';
  }
  return u;
}
/** Unbook someone's own shows between two dates (until blank = from date onwards) */
function cancelRange_(email, from, until) {
  if (!until) return cancelFutureFor_(email, from);
  let n = 0;
  rows_('Bookings').forEach(function (b) {
    if (b.Email !== email || b.Status === 'Cancelled') return;
    if (isTrue_(b.Permanent)) {
      const hits = occurrences_(b, from, until);
      if (hits.length) { setCells_('Bookings', b._row, { SkipDates: list_(b.SkipDates).concat(hits).join(',') }); n += hits.length; }
    } else if (b.Date >= from && b.Date <= until) { setCells_('Bookings', b._row, { Status: 'Cancelled' }); n++; }
  });
  return n;
}
function canDecide_(me, r) {
  if (r.Email === me.Email) return false;
  if (r.Status === 'Pending Supervisor') return me.Role === 'Supervisor' || me.Role === 'Manager';
  if (r.Status === 'Pending Manager') return me.Role === 'Manager';
  return false;
}
function publicUser_(u) {
  return { email: u.Email, name: u.Name, role: u.Role, area: u.Area, startDate: u.StartDate || String(u.CreatedAt || '').slice(0, 10),
           status: statusOf_(u), suspendedUntil: u.SuspendedUntil || '',
           target: Number(u.WeeklyTarget) || settings_().defaultTarget };
}
function publicRequest_(r) {
  return { id: r.ID, email: r.Email, name: r.Name, type: r.Type, start: r.StartDate, end: r.EndDate,
    reason: r.Reason, status: r.Status, supervisorBy: r.SupervisorBy, managerBy: r.ManagerBy,
    comment: r.Comment, created: r.CreatedAt };
}

function hash_(pw, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(salt) + '|' + String(pw), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function genPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const seed = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Math.random());
  let out = '';
  for (let i = 0; i < 10; i++) out += chars[((seed[i] + 256) % 256) % chars.length];
  return 'Kx-' + out;
}
function createUser_(o) {
  const email = String(o.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address.');
  if (findUser_(email)) throw new Error('An account with that email already exists.');
  const name = String(o.name || '').trim();
  if (!name) throw new Error('Enter a name.');
  const role = ROLES.indexOf(o.role) >= 0 ? o.role : 'Staff';
  const area = AREAS.indexOf(o.area) >= 0 ? o.area : AREAS[0];
  const pw = genPassword_();
  const salt = Utilities.getUuid();
  append_('Users', { Email: email, Name: name, Role: role, Area: area, Salt: salt,
    PasswordHash: hash_(pw, salt), Active: true, WeeklyTarget: Number(o.target) > 0 ? Number(o.target) : '',
    MustChangePassword: true, CreatedAt: now_(), LastLogin: '', Status: 'Active',
    Phone: String(o.phone || '').trim().slice(0, 30), StudentNumber: String(o.studentNumber || '').trim().slice(0, 30),
    StartDate: o.startDate ? cleanDate_(o.startDate) : today_() });
  return pw;
}
function resetPassword_(email) {
  const u = findUser_(email);
  if (!u) throw new Error('No account with that email.');
  const pw = genPassword_();
  const salt = Utilities.getUuid();
  setCells_('Users', u._row, { Salt: salt, PasswordHash: hash_(pw, salt), MustChangePassword: true });
  CacheService.getScriptCache().remove('fail_' + u.Email);
  return pw;
}


/* ───────────────────────── Listener requests ───────────────────────── */

/** From the public site — anyone can send one. Light spam protection. */
function publicSendRequest(o) {
  o = o || {};
  if (o.website) return { ok: true };                      // hidden honeypot field filled → bot
  const name = String(o.name || '').trim().slice(0, 40);
  const message = String(o.message || '').trim().slice(0, 300);
  const type = LISTENER_TYPES.indexOf(o.type) >= 0 ? o.type : 'Request';
  if (!name) throw new Error('Please enter your name.');
  if (message.length < 2) throw new Error('Please write a message.');
  const cache = CacheService.getScriptCache();
  const perName = 'lr_' + name.toLowerCase();
  if (cache.get(perName)) throw new Error('Thanks! Please wait a minute before sending another one.');
  const minute = 'lr_rate_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMddHHmm');
  const n = Number(cache.get(minute) || 0);
  if (n > 40) throw new Error('We’re getting a lot of messages right now — please try again in a minute.');
  cache.put(minute, String(n + 1), 120);
  cache.put(perName, '1', 60);
  withLock_(function () {
    append_('Listener', { ID: id_(), Name: name, Type: type, Message: message, Source: 'Website', Email: '',
      Verified: false, Status: 'New', Pinned: false, CreatedAt: nowSec_() });
  });
  return { ok: true };
}

/** From inside the staff portal — shows as verified (red) */
function staffSendRequest(token, o) {
  const u = auth_(token);
  const message = String(o.message || '').trim().slice(0, 300);
  if (message.length < 2) throw new Error('Please write a message.');
  const type = LISTENER_TYPES.indexOf(o.type) >= 0 ? o.type : 'Request';
  withLock_(function () {
    append_('Listener', { ID: id_(), Name: u.Name, Type: type, Message: message, Source: 'Staff portal', Email: u.Email,
      Verified: true, Status: 'New', Pinned: false, CreatedAt: nowSec_() });
  });
  return getListenerRequests(token);
}

function getListenerRequests(token) {
  const u = auth_(token, null, true);
  const since = addDays_(today_(), -3);
  const list = rows_('Listener').filter(function (r) { return String(r.CreatedAt) >= since || isTrue_(r.Pinned); })
    .sort(byDesc_('CreatedAt')).slice(0, 400).map(function (r) {
      return { id: r.ID, name: r.Name, type: r.Type, message: r.Message, source: r.Source, verified: isTrue_(r.Verified),
               status: r.Status, pinned: isTrue_(r.Pinned), created: r.CreatedAt };
    });
  syncIono_();
  return { list: list, types: LISTENER_TYPES, nowPlaying: nowPlaying_(), recent: played_(10), autoNowPlaying: !!settings_().ionoStream,
           publicUrl: publicUrl_(), readOnly: u.Status === 'Suspended' };
}

/** action: pin | unpin | done | new | delete | restore */
function updateListenerRequest(token, id, action) {
  auth_(token);
  withLock_(function () {
    const r = rows_('Listener').filter(function (x) { return x.ID === id; })[0];
    if (!r) throw new Error('Message not found.');
    const upd = { pin: { Pinned: true }, unpin: { Pinned: false }, done: { Status: 'Done' }, 'new': { Status: 'New' },
                  'delete': { Status: 'Deleted', Pinned: false }, restore: { Status: 'New' } }[action];
    if (!upd) throw new Error('Unknown action.');
    setCells_('Listener', r._row, upd);
  });
  return getListenerRequests(token);
}

/* ───────────────────────── Now playing ───────────────────────── */

function setNowPlaying(token, o) {
  const u = auth_(token);
  const title = String(o.title || '').trim().slice(0, 120);
  const artist = String(o.artist || '').trim().slice(0, 120);
  if (!title) throw new Error('Enter the song title.');
  const art = lookupArt_(title, artist);
  withLock_(function () {
    append_('Played', { ID: id_(), Title: title, Artist: artist, Art: art, PlayedAt: nowSec_(), By: u.Name });
  });
  return getListenerRequests(token);
}
function lookupArt_(title, artist) {
  try {
    const url = 'https://itunes.apple.com/search?entity=song&limit=1&country=ZA&term=' +
                encodeURIComponent((artist + ' ' + title).trim());
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return '';
    const j = JSON.parse(res.getContentText());
    const a = j.results && j.results[0] && j.results[0].artworkUrl100;
    return a ? a.replace('100x100bb', '600x600bb') : '';
  } catch (e) { return ''; }
}
function played_(n) {
  return rows_('Played').sort(byDesc_('PlayedAt')).slice(0, n).map(function (p) {
    return { title: p.Title, artist: p.Artist, art: p.Art, at: String(p.PlayedAt).slice(11, 16), date: String(p.PlayedAt).slice(0, 10) };
  });
}
function nowPlaying_() { const p = played_(1)[0]; return p && p.date === today_() ? p : null; }

/* ───────────────────────── Public listener site (?p=listen) ───────────────────────── */

function publicGetState() {
  const s = settings_();
  syncIono_();
  const live = live_(rows_('Bookings'));
  const strip = function (o) { return o ? { show: o.show, name: o.name, with: o.with, start: o.start, end: o.end,
    date: o.date, image: o.image, startMs: o.startMs, endMs: o.endMs } : null; };
  return { onAir: live.onAir.map(strip), next: strip(live.next), nowPlaying: nowPlaying_(), recent: played_(15),
           streamUrl: s.streamUrl, types: LISTENER_TYPES, serverNow: Date.now() };
}

function publicUrl_() {
  try { return ScriptApp.getService().getUrl() + '?p=listen'; } catch (e) { return ''; }
}
function nowSec_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }


/* ───────────────────────── Iono.fm now playing (RadioDJ → Iono.fm → here) ─────────────────────────
 * Iono.fm publishes the stream's current song at https://edge.iono.fm/xmetadata/<stream_id>.
 * We check it at most every 30 seconds (Iono asks for no more than every 15s), and only when
 * someone has the listener site or the Requests page open. New songs are added to "Played". */
function ionoId_(v) {
  let x = String(v || '').trim();
  const i = x.indexOf('xmetadata/');
  if (i >= 0) x = x.slice(i + 10);
  x = x.split('?')[0].split('/')[0].trim();
  if (x && !/^[A-Za-z0-9_-]+$/.test(x)) throw new Error('That doesn’t look like an Iono.fm stream ID.');
  return x;
}
function fetchIono_(id) {
  const res = UrlFetchApp.fetch('https://edge.iono.fm/xmetadata/' + encodeURIComponent(id), { muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code === 404) return { off: true };
  if (code !== 200) return { error: 'Iono.fm answered ' + code };
  return JSON.parse(res.getContentText());
}
function syncIono_() {
  const id = settings_().ionoStream;
  if (!id) return;
  const cache = CacheService.getScriptCache();
  if (cache.get('iono')) return;
  cache.put('iono', '1', 30);
  try {
    const j = fetchIono_(id);
    if (!j || j.off || j.error) return;
    const event = String(j.event || 'song').toLowerCase();
    const title = String(j.title || '').trim(), artist = String(j.artist || '').trim();
    if (event !== 'song' || !title || !artist) return;                          // skip ads, news, jingles without an artist
    if (/kovsie/i.test(title + ' ' + artist) || /(^|\W)ad(\W|$)/i.test(title)) return; // station imaging / ads
    const last = rows_('Played').sort(byDesc_('PlayedAt'))[0];
    if (last && String(last.Title).toLowerCase() === title.toLowerCase() &&
        String(last.Artist).toLowerCase() === artist.toLowerCase()) return;   // same song still playing
    const art = (j.artwork && String(j.artwork).indexOf('http') === 0) ? String(j.artwork) : lookupArt_(title, artist);
    withLock_(function () {
      append_('Played', { ID: id_(), Title: title.slice(0, 120), Artist: artist.slice(0, 120), Art: art, PlayedAt: nowSec_(), By: 'Iono.fm' });
    });
  } catch (e) { /* network hiccup — try again on the next check */ }
}
/** Settings → "Test" button */
function testIono(token, v) {
  auth_(token, ['Manager']);
  const id = ionoId_(v);
  if (!id) throw new Error('Enter your Iono.fm stream ID first.');
  const j = fetchIono_(id);
  if (j.off) return 'Connected, but Iono.fm has no song info right now (stream offline, or RadioDJ isn’t sending titles).';
  if (j.error) throw new Error(j.error + ' — check the stream ID.');
  return 'Connected! Iono.fm says: ' + (j.event || 'song') + ' — ' + (j.artist || '?') + ' – ' + (j.title || '?');
}

/* ───────────────────────── JSON API (used by the GitHub Pages site) ─────────────────────────
 * The pages on GitHub send:  POST <this web app URL>   body: {"fn": "getDashboard", "args": [token, ...]}
 * Only the functions listed here can be called from outside. Every one still checks the sign-in token itself. */
const API_FUNCTIONS = {
  login: login, logout: logout, changePassword: changePassword,
  getDashboard: getDashboard, bookSlot: bookSlot, cancelBooking: cancelBooking, setBookingImage: setBookingImage,
  submitRequest: submitRequest, cancelRequest: cancelRequest, deleteHours: deleteHours,
  getManagement: getManagement, decideRequest: decideRequest, managerAddHours: managerAddHours, removeHours: removeHours,
  managerCancelBooking: managerCancelBooking, setNoShow: setNoShow, reviewShow: reviewShow,
  addUser: addUser, resetUserPassword: resetUserPassword, updateUser: updateUser, hrAction: hrAction,
  saveSettings: saveSettings, testIono: testIono,
  getLive: getLive, getPulse: getPulse, getNotifications: getNotifications, markNotificationsRead: markNotificationsRead,
  staffSendRequest: staffSendRequest, getListenerRequests: getListenerRequests, updateListenerRequest: updateListenerRequest,
  setNowPlaying: setNowPlaying, publicGetState: publicGetState, publicSendRequest: publicSendRequest
};
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API_FUNCTIONS.hasOwnProperty(body.fn) ? API_FUNCTIONS[body.fn] : null;
    if (!fn) throw new Error('Unknown request.');
    const args = Array.isArray(body.args) ? body.args.slice(0, 6) : [];
    out = { ok: true, result: fn.apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
