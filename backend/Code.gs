/**
 * Let-Out Board backend (Google Apps Script, bound to a Google Sheet).
 *
 * Setup (one time):
 *   1. Paste this file into Extensions > Apps Script of a new Google Sheet.
 *   2. Put the family passcode in PASSCODE below (only in your Sheet's copy,
 *      never in the public GitHub repo).
 *   3. Run setup() once and approve the permissions.
 *   4. Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 *
 * Tabs it keeps:
 *   People   who exists and their usual week (hours worked per weekday)
 *   Dogs     each dog's let-out rule
 *   Hours    one-off hours for a specific date (0 = off)
 *   Claims   who claimed which let-out
 *   Counts   let-out counts set by hand for a specific date
 *   Daycare  day care drop-offs (which let-out drops the dog off) and who picks up
 *   Coverage this week and next, one row per let-out (rebuilt on every change and nightly)
 */

const PASSCODE = '';            // set this in your Sheet's copy only
const SHEET_ID = '';            // blank = the Sheet this script is attached to
const TZ = 'America/Chicago';
function book() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive(); }

const HEADERS = {
  People:   ['id', 'name', 'schedule', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun', 'order'],
  Dogs:     ['id', 'name', 'house', 'owners', 'letouts', 'long_day_hours', 'long_day_letouts', 'second_letout_helpers', 'order'],
  Hours:    ['date', 'person', 'hours', 'updated'],
  Claims:   ['date', 'dog', 'slot', 'person', 'updated'],
  Counts:   ['date', 'dog', 'letouts', 'updated'],
  Daycare:  ['date', 'dog', 'slot', 'pickup', 'updated'],
  Coverage: ['Date', 'Day', 'Dog', 'House', 'Let-out', 'Status', 'Claimed by', 'Note']
};
const DAY_NAMES = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const WEEK_COLS = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];

// ---------------------------------------------------------------- setup

function setup() {
  const ss = book();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.getRange('A:A').setNumberFormat('@');   // dates stay plain text like 2026-10-19
  });
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);

  const people = ss.getSheetByName('People');
  if (people.getLastRow() < 2) {
    people.getRange(2, 1, 5, 11).setValues([
      ['bryce',   'Bryce',   'fixed',  9.25, 9.25, 9.25, 9.25, 9.25, 0, 0, 1],
      ['mckayla', 'McKayla', 'fixed',  9.25, 9.25, 9.25, 9.25, 9.25, 0, 0, 2],
      ['brynn',   'Brynn',   'weekly', '', '', '', '', '', '', '', 3],
      ['mom',     'Mom',     'none',   '', '', '', '', '', '', '', 4],
      ['grandma', 'Grandma', 'none',   '', '', '', '', '', '', '', 5]
    ]);
  }
  const dogs = ss.getSheetByName('Dogs');
  if (dogs.getLastRow() < 2) {
    dogs.getRange(2, 1, 2, 9).setValues([
      ['bubba',  'Bubba',  'Bryce & McKayla\'s house', 'bryce,mckayla', 1, '', '', '', 1],
      ['harvey', 'Harvey', 'Brynn\'s house',           'brynn',         1, 12, 2, 'bryce,mckayla', 2]
    ]);
  }

  const props = PropertiesService.getScriptProperties();
  if (PASSCODE) props.setProperty('PASSCODE', PASSCODE);
  if (!props.getProperty('PASSCODE')) throw new Error('Put the family passcode in PASSCODE at the top of this file, then run setup() again.');

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'nightly') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightly').timeBased().everyDays(1).atHour(1).inTimezone(TZ).create();

  rebuildCoverage();
}

function nightly() { rebuildCoverage(); }

// ---------------------------------------------------------------- web app

function doGet() {
  return json({ok: true, message: 'Let-Out Board backend is running.'});
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json({ok: false, error: 'bad_request'}); }

  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE') || '';
  if (!pass || String(req.passcode || '').trim().toLowerCase() !== pass.trim().toLowerCase()) {
    return json({ok: false, error: 'passcode'});
  }
  try {
    if (req.action === 'load') return json({ok: true, data: snapshot()});
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (req.action === 'setHours') setHours(req);
      else if (req.action === 'claim') setClaim(req);
      else if (req.action === 'setCount') setCount(req);
      else if (req.action === 'daycare') setDaycare(req);
      else if (req.action === 'pickup') setPickup(req);
      else return json({ok: false, error: 'unknown_action'});
      SpreadsheetApp.flush();
      rebuildCoverage();
    } finally {
      lock.releaseLock();
    }
    return json({ok: true, data: snapshot()});
  } catch (err) {
    return json({ok: false, error: 'server', message: String(err && err.message || err)});
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- writes

function checkDate(d) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d || ''))) throw new Error('Bad date');
  return String(d);
}
function stamp() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function setHours(req) {
  const date = checkDate(req.date), person = String(req.person || '');
  if (!readPeople().some(function (p) { return p.id === person; })) throw new Error('Unknown person');
  const keys = {date: date, person: person};
  if (req.hours === null || req.hours === '' || req.hours === undefined) return removeRows('Hours', keys);
  const h = Number(req.hours);
  if (!(h >= 0 && h <= 24)) throw new Error('Hours must be 0 to 24');
  upsert('Hours', keys, {hours: h, updated: stamp()});
}

function setClaim(req) {
  const date = checkDate(req.date), dog = String(req.dog || ''), slot = Number(req.slot);
  if (!readDogs().some(function (d) { return d.id === dog; })) throw new Error('Unknown dog');
  if (!(slot >= 1 && slot <= 6)) throw new Error('Bad let-out number');
  const keys = {date: date, dog: dog, slot: slot};
  if (!req.person) {
    removeRows('Claims', keys);
    // If this let-out was the day care drop-off, the plan is off too, so the cancelled let-outs come back.
    removeRows('Daycare', keys);
    return;
  }
  if (!readPeople().some(function (p) { return p.id === req.person; })) throw new Error('Unknown person');
  upsert('Claims', keys, {person: String(req.person), updated: stamp()});
}

function setCount(req) {
  const date = checkDate(req.date), dog = String(req.dog || '');
  if (!readDogs().some(function (d) { return d.id === dog; })) throw new Error('Unknown dog');
  const keys = {date: date, dog: dog};
  if (req.letouts === null || req.letouts === '' || req.letouts === undefined) return removeRows('Counts', keys);
  const n = Math.round(Number(req.letouts));
  if (!(n >= 0 && n <= 6)) throw new Error('Let-outs must be 0 to 6');
  upsert('Counts', keys, {letouts: n, updated: stamp()});
}

function daycareRow(date, dog) {
  return rows('Daycare').filter(function (r) { return cellStr(r.date) === date && cellStr(r.dog) === dog; })[0] || null;
}

function setDaycare(req) {
  const date = checkDate(req.date), dog = String(req.dog || '');
  if (!readDogs().some(function (d) { return d.id === dog; })) throw new Error('Unknown dog');
  const keys = {date: date, dog: dog};
  if (!req.on) return removeRows('Daycare', keys);
  const slot = Number(req.slot);
  if (!(slot >= 1 && slot <= 6)) throw new Error('Bad let-out number');
  const existing = daycareRow(date, dog);
  upsert('Daycare', keys, {slot: slot, pickup: existing ? cellStr(existing.pickup) : '', updated: stamp()});
}

function setPickup(req) {
  const date = checkDate(req.date), dog = String(req.dog || '');
  const existing = daycareRow(date, dog);
  if (!existing) throw new Error('No day care drop-off that day');
  if (req.person && !readPeople().some(function (p) { return p.id === req.person; })) throw new Error('Unknown person');
  upsert('Daycare', {date: date, dog: dog}, {slot: Number(existing.slot), pickup: req.person ? String(req.person) : '', updated: stamp()});
}

function sheet(name) {
  const ss = book();
  let sh = ss.getSheetByName(name);
  if (!sh && HEADERS[name]) {                                        // tabs added after setup() appear on first use
    const cov = ss.getSheetByName('Coverage');
    sh = ss.insertSheet(name, cov ? cov.getIndex() - 1 : ss.getSheets().length);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function cellStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v).trim();
}
function rowMatches(head, row, keys) {
  return Object.keys(keys).every(function (k) { return cellStr(row[head.indexOf(k)]) === String(keys[k]); });
}
function upsert(name, keys, vals) {
  const sh = sheet(name), values = sh.getDataRange().getValues(), head = values[0];
  const full = {};
  Object.keys(keys).forEach(function (k) { full[k] = keys[k]; });
  Object.keys(vals).forEach(function (k) { full[k] = vals[k]; });
  const row = head.map(function (k) { return full[k] === undefined ? '' : full[k]; });
  for (let i = 1; i < values.length; i++) {
    if (rowMatches(head, values[i], keys)) { sh.getRange(i + 1, 1, 1, head.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}
function removeRows(name, keys) {
  const sh = sheet(name), values = sh.getDataRange().getValues(), head = values[0];
  for (let i = values.length - 1; i >= 1; i--) if (rowMatches(head, values[i], keys)) sh.deleteRow(i + 1);
}

// ---------------------------------------------------------------- reads

function rows(name) {
  const values = sheet(name).getDataRange().getValues();
  const head = values.shift() || [];
  return values.filter(function (r) { return r.some(function (c) { return c !== ''; }); }).map(function (r) {
    const o = {}; head.forEach(function (k, j) { o[k] = r[j]; }); return o;
  });
}
function readPeople() {
  return rows('People').map(function (r) {
    return {
      id: cellStr(r.id), name: cellStr(r.name) || cellStr(r.id),
      schedule: (cellStr(r.schedule) || 'none').toLowerCase(),
      week: WEEK_COLS.map(function (c) { return r[c] === '' ? null : Number(r[c]); }),
      order: Number(r.order) || 0
    };
  }).filter(function (p) { return p.id; }).sort(function (a, b) { return a.order - b.order; });
}
function readDogs() {
  return rows('Dogs').map(function (r) {
    return {
      id: cellStr(r.id), name: cellStr(r.name) || cellStr(r.id), house: cellStr(r.house),
      owners: splitIds(r.owners),
      letouts: r.letouts === '' ? 1 : Number(r.letouts),
      longHours: r.long_day_hours === '' ? null : Number(r.long_day_hours),
      longLetouts: r.long_day_letouts === '' ? null : Number(r.long_day_letouts),
      secondHelpers: splitIds(r.second_letout_helpers),
      order: Number(r.order) || 0
    };
  }).filter(function (d) { return d.id; }).sort(function (a, b) { return a.order - b.order; });
}
function splitIds(v) { return cellStr(v).split(',').map(function (s) { return s.trim(); }).filter(String); }

// ---------------------------------------------------------------- the plan

/* Pure date helpers on yyyy-MM-dd strings (UTC math, so no time-zone drift). */
function parseD(s) { const p = s.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])); }
function fmtD(d) { return d.toISOString().slice(0, 10); }
function addDays(s, n) { const d = parseD(s); d.setUTCDate(d.getUTCDate() + n); return fmtD(d); }
function weekdayIdx(s) { return (parseD(s).getUTCDay() + 6) % 7; }      // Mon = 0
function mondayOf(s) { return addDays(s, -weekdayIdx(s)); }

function snapshot() {
  return buildPlan({
    today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    people: readPeople(),
    dogs: readDogs(),
    hours: rows('Hours').map(function (r) { return {date: cellStr(r.date), person: cellStr(r.person), hours: Number(r.hours)}; }),
    claims: rows('Claims').map(function (r) { return {date: cellStr(r.date), dog: cellStr(r.dog), slot: Number(r.slot), person: cellStr(r.person)}; }),
    counts: rows('Counts').map(function (r) { return {date: cellStr(r.date), dog: cellStr(r.dog), letouts: Number(r.letouts)}; }),
    daycare: rows('Daycare').map(function (r) { return {date: cellStr(r.date), dog: cellStr(r.dog), slot: Number(r.slot), pickup: cellStr(r.pickup)}; })
  });
}

/* Everything below is pure: it takes plain data and returns the plan. */
function buildPlan(src) {
  const people = src.people, dogs = src.dogs;
  const byId = {}; people.forEach(function (p) { byId[p.id] = p; });
  const hoursMap = {}; src.hours.forEach(function (h) { hoursMap[h.date + '|' + h.person] = h.hours; });
  const claimMap = {}; src.claims.forEach(function (c) { claimMap[c.date + '|' + c.dog + '|' + c.slot] = c.person; });
  const countMap = {}; src.counts.forEach(function (c) { countMap[c.date + '|' + c.dog] = c.letouts; });
  const dcMap = {}; (src.daycare || []).forEach(function (c) { dcMap[c.date + '|' + c.dog] = c; });
  const nm = function (id) { return byId[id] ? byId[id].name : id; };
  const joinNames = function (ids) { const n = ids.map(nm); return n.length < 3 ? n.join(' and ') : n.slice(0, -1).join(', ') + ' and ' + n[n.length - 1]; };

  // Hours worked that day: number (0 = off), or null when the person enters hours weekly and hasn't yet.
  function hoursFor(id, date) {
    const k = date + '|' + id;
    if (hoursMap.hasOwnProperty(k)) return hoursMap[k];
    const p = byId[id]; if (!p) return null;
    if (p.schedule === 'fixed') { const v = p.week[weekdayIdx(date)]; return v === null || isNaN(v) ? 0 : v; }
    if (p.schedule === 'weekly') return null;
    return 0;                                                        // 'none': no job, always free
  }
  function usual(id, date) {
    const p = byId[id]; if (!p || p.schedule !== 'fixed') return null;
    const v = p.week[weekdayIdx(date)]; return v === null || isNaN(v) ? 0 : v;
  }

  function dogDay(d, date) {
    const owners = d.owners.filter(function (id) { return byId[id]; });
    const hrs = owners.map(function (id) { return {id: id, h: hoursFor(id, date)}; });
    const home = hrs.filter(function (x) { return x.h === 0; }).map(function (x) { return x.id; });
    const missing = hrs.filter(function (x) { return x.h === null; }).map(function (x) { return x.id; });
    let state, needed = 0, reason;
    if (!owners.length) { state = 'setup'; reason = 'Nobody set as living with ' + d.name; }
    else if (home.length) { state = 'home'; reason = joinNames(home) + (home.length > 1 ? ' are' : ' is') + ' off'; }
    else if (missing.length) { state = 'unknown'; needed = null; reason = joinNames(missing) + '’s hours not entered'; }
    else {
      const away = Math.min.apply(null, hrs.map(function (x) { return x.h; }));
      const long = d.longHours !== null && d.longLetouts !== null && away >= d.longHours;
      needed = long ? d.longLetouts : d.letouts;
      state = 'needs';
      reason = fmtHours(away) + ' workday';
    }
    const ck = date + '|' + d.id;
    const overridden = countMap.hasOwnProperty(ck);
    if (overridden) { needed = countMap[ck]; state = needed > 0 ? 'needs' : 'home'; reason = 'Set by hand'; }
    const slots = [];
    for (let k = 1; k <= (needed || 0); k++) {
      slots.push({n: k, claimedBy: claimMap[date + '|' + d.id + '|' + k] || null, free: freeFor(d, owners, date, k)});
    }
    // Day care: the drop-off let-out stays, later let-outs that day are cancelled, and a pickup task is added.
    let daycare = null;
    const dc = dcMap[date + '|' + d.id];
    if (dc && dc.slot >= 1 && dc.slot <= slots.length && slots[dc.slot - 1].claimedBy) {
      const cancelled = slots.length - dc.slot;
      slots.length = dc.slot;
      slots[dc.slot - 1].daycare = true;
      daycare = {slot: dc.slot, cancelled: cancelled, pickup: {claimedBy: dc.pickup || null, free: freeFor(d, owners, date, 2)}};
    }
    return {dog: d.id, state: state, needed: needed, reason: reason, overridden: overridden, slots: slots, daycare: daycare};
  }

  function freeFor(d, owners, date, k) {
    const out = [];
    people.forEach(function (p) {
      if (owners.indexOf(p.id) >= 0) return;
      if (hoursFor(p.id, date) === 0) out.push(p.id);
    });
    if (k >= 2) d.secondHelpers.forEach(function (id) { if (byId[id] && out.indexOf(id) < 0) out.push(id); });
    return out;
  }

  const mon = mondayOf(src.today);
  const weeks = [0, 7].map(function (off) {
    const m = addDays(mon, off);
    const days = [];
    for (let i = 0; i < 7; i++) {
      const date = addDays(m, i);
      const hours = {}, usualH = {};
      people.forEach(function (p) { if (p.schedule !== 'none'){ hours[p.id] = hoursFor(p.id, date); usualH[p.id] = usual(p.id, date); } });
      days.push({date: date, day: DAY_NAMES[i], hours: hours, usual: usualH, dogs: dogs.map(function (d) { return dogDay(d, date); })});
    }
    return {monday: m, days: days};
  });

  return {
    today: src.today,
    people: people.map(function (p) { return {id: p.id, name: p.name, schedule: p.schedule, week: p.week}; }),
    dogs: dogs.map(function (d) { return {id: d.id, name: d.name, house: d.house, owners: d.owners}; }),
    weeks: weeks
  };
}

function fmtHours(h) {
  const whole = Math.floor(h), m = Math.round((h - whole) * 60);
  return m ? whole + 'h ' + m + 'm' : whole + '-hour';
}

// ---------------------------------------------------------------- coverage tab (for Muse)

function coverageRows(plan) {
  const names = {}; plan.people.forEach(function (p) { names[p.id] = p.name; });
  const dogs = {}; plan.dogs.forEach(function (d) { dogs[d.id] = d; });
  const out = [];
  let unclaimed = 0, unknown = 0;
  plan.weeks.forEach(function (w) {
    w.days.forEach(function (day) {
      if (day.date < plan.today) return;                             // past days don't need covering
      day.dogs.forEach(function (dd) {
        const dog = dogs[dd.dog];
        if (dd.state === 'unknown') {
          unknown++;
          out.push([day.date, day.day, dog.name, dog.house, '', 'HOURS NOT ENTERED', '', dd.reason]);
        }
        const freeNote = function (free) { return free.length ? 'Free: ' + free.map(function (id) { return names[id] || id; }).join(', ') : 'Nobody listed as free'; };
        dd.slots.forEach(function (s) {
          if (!s.claimedBy) unclaimed++;
          const dropNote = s.daycare ? (dd.daycare.cancelled ? 'Day care drop-off; rest of day at day care' : 'Day care drop-off') : '';
          out.push([day.date, day.day, dog.name, dog.house, s.n + ' of ' + dd.needed + (s.daycare ? ' + day care drop-off' : ''),
            s.claimedBy ? 'Claimed' : 'UNCLAIMED', s.claimedBy ? (names[s.claimedBy] || s.claimedBy) : '',
            s.claimedBy ? dropNote : freeNote(s.free)]);
        });
        if (dd.daycare) {
          const pu = dd.daycare.pickup;
          if (!pu.claimedBy) unclaimed++;
          out.push([day.date, day.day, dog.name, dog.house, 'Day care pickup',
            pu.claimedBy ? 'Claimed' : 'UNCLAIMED', pu.claimedBy ? (names[pu.claimedBy] || pu.claimedBy) : '',
            pu.claimedBy ? '' : freeNote(pu.free)]);
        }
      });
    });
  });
  return {rows: out, unclaimed: unclaimed, unknown: unknown};
}

function rebuildCoverage() {
  const plan = snapshot();
  const c = coverageRows(plan);
  const sh = sheet('Coverage');
  sh.clearContents();
  const summary = [
    ['Let-out coverage from ' + plan.today + ' through ' + addDays(plan.weeks[1].monday, 6), '', '', '', '', '', '', ''],
    ['Updated ' + stamp() + ' · Unclaimed let-outs: ' + c.unclaimed + ' · Days with hours not entered: ' + c.unknown, '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    HEADERS.Coverage
  ];
  const all = summary.concat(c.rows);
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange(1, 1, all.length, 8).setValues(all);
  sh.getRange(4, 1, 1, 8).setFontWeight('bold');
  sh.getRange(1, 1).setFontWeight('bold');
}
