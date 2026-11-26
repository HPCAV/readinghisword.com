/**
 * ReadingHisWord.com — backend (Google Apps Script)
 *
 * Runs in Dick's Google account. It:
 *   1. Indexes the Sunday School Lessons folder in Drive (hourly).
 *   2. Serves the library and rendered lesson outlines to the website as JSON.
 *      Only documents inside the lessons folder can ever be served.
 *   3. Handles email sign-up (double opt-in), confirmation and unsubscribe.
 *   4. Sends each subscriber the next outline in their chosen series, daily or weekly, and invites
 *      them to pick a new series when they finish one.
 *
 * First-time setup: run setup() once from the editor, then deploy as a web app
 * (Execute as: Me, Who has access: Anyone).
 */

const CONFIG = {
  LIBRARY_FOLDER_ID: '1Bhw76UfoRGHOOc5nsQ08m7_A8zslbH_m', // Southview Baptist Church › Sunday School Lessons
  SITE_URL: 'https://readinghisword.com',
  SITE_NAME: 'Reading His Word',
  TIMEZONE: 'America/Chicago',      // Southview Baptist Church, Lincoln NE
  SEND_HOUR: 6,                     // emails go out at or after 6 a.m. local time
  MAX_CONFIRMATIONS_PER_DAY: 40,    // protects the 100-recipient/day Gmail quota from sign-up abuse
  INDEX_TIME_BUDGET_MS: 4.5 * 60 * 1000,
};

/* ------------------------------------------------------------------ */
/* Setup & triggers                                                    */
/* ------------------------------------------------------------------ */

function setup() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, 'Subscribers', SUB_HEADERS_);
  ensureSheet_(ss, 'Index', ['id', 'seriesOrder', 'series', 'date', 'passage', 'heading', 'name', 'modified']);
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('hourly').timeBased().everyHours(1).create();

  rebuildIndex();
  Logger.log('Setup complete. Data sheet: ' + ss.getUrl());
}

function hourly() {
  rebuildIndex();
  sendScheduled_();
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('ReadingHisWord.com – Data');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ------------------------------------------------------------------ */
/* Library index                                                       */
/* ------------------------------------------------------------------ */

function rebuildIndex() {
  const started = Date.now();
  const sh = getSpreadsheet_().getSheetByName('Index');
  const old = {};
  readRows_(sh).forEach(function (r) { old[r.id] = r; });

  const root = DriveApp.getFolderById(CONFIG.LIBRARY_FOLDER_ID);
  const rows = [];

  const seriesFolders = root.getFolders();
  while (seriesFolders.hasNext()) {
    const f = seriesFolders.next();
    if (isHiddenFolder_(f.getName())) continue;
    const m = f.getName().match(/^\s*(\d+)\s*[.)-]?\s*(.+)$/);
    const order = m ? Number(m[1]) : 999;
    const title = m ? m[2].trim() : f.getName();
    collectDocs_(f, order, title, rows);
  }
  collectDocs_(root, 1000, 'Other Lessons', rows, true);

  // Drop "Copy of X" when X also exists.
  const names = {};
  rows.forEach(function (r) { names[r.name] = true; });
  const filtered = rows.filter(function (r) {
    return !(/^Copy of /i.test(r.name) && names[r.name.replace(/^Copy of /i, '')]);
  });

  // Fill in headings (first line of each doc), reusing what we already know.
  filtered.forEach(function (r) {
    const prev = old[r.id];
    if (prev && String(prev.modified) === r.modified && prev.heading) {
      r.heading = prev.heading;
    } else if (Date.now() - started < CONFIG.INDEX_TIME_BUDGET_MS) {
      try { r.heading = firstLine_(r.id); } catch (e) { r.heading = ''; }
    } else {
      r.heading = prev ? prev.heading : '';
      if (!prev || !prev.heading) r.modified = ''; // retry next run
    }
  });

  filtered.sort(function (a, b) {
    return a.seriesOrder - b.seriesOrder || (a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
  });

  const headers = ['id', 'seriesOrder', 'series', 'date', 'passage', 'heading', 'name', 'modified'];
  const values = filtered.map(function (r) { return headers.map(function (h) { return r[h]; }); });
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), headers.length).clearContent();
  if (values.length) {
    sh.getRange(2, 1, values.length, headers.length).setNumberFormat('@').setValues(values);
  }
  CacheService.getScriptCache().remove('library_meta');
  return values.length;
}

function collectDocs_(folder, order, seriesTitle, rows, shallow) {
  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const parsed = parseName_(name);
    rows.push({
      id: file.getId(),
      seriesOrder: order,
      series: seriesTitle,
      date: parsed.date,
      passage: parsed.passage,
      heading: '',
      name: name,
      modified: String(file.getLastUpdated().getTime()),
    });
  }
  if (shallow) return;
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    if (isHiddenFolder_(sub.getName())) continue; // e.g. "Drafts": kept in Drive, never shown on the site
    collectDocs_(sub, order, seriesTitle, rows);
  }
}

/** Folders named "Drafts", "Archive", "Hidden", or starting with "_" are skipped. */
function isHiddenFolder_(name) {
  return /^\s*(_|drafts?\b|archive\b|hidden\b|not for site\b)/i.test(name);
}

function parseName_(name) {
  const d = name.match(/(\d{4})(\d{2})(\d{2})/);
  const date = d ? d[1] + '-' + d[2] + '-' + d[3] : '';
  let t = name.replace(/^(Copy of\s+)?Sunday School Lesson\s*/i, '');
  // Drop the date wherever it sits (start, middle or end), plus placeholders like "2022????".
  t = t.replace(/\s*[-–]?\s*(\d{8}|\d{4}\?+)\s*[-–]?\s*/g, ' - ');
  t = t.replace(/^[\s\-–]+|[\s\-–]+$/g, '').replace(/\s+-\s+-\s+/g, ' - ');
  t = t.replace(/_/g, ':').trim();
  return { date: date, passage: t || name };
}

function firstLine_(id) {
  const paras = DocumentApp.openById(id).getBody().getParagraphs();
  for (let i = 0; i < paras.length && i < 10; i++) {
    const t = paras[i].getText().trim();
    if (t) return t.length > 140 ? t.slice(0, 137) + '…' : t;
  }
  return '';
}

function getIndex_() {
  const cache = CacheService.getScriptCache();
  const hit = cacheGetBig_(cache, 'library_meta');
  if (hit) return JSON.parse(hit);
  const rows = readRows_(getSpreadsheet_().getSheetByName('Index'));
  cachePutBig_(cache, 'library_meta', JSON.stringify(rows), 3600);
  return rows;
}

function libraryPayload_() {
  const rows = getIndex_();
  const bySeries = {};
  const list = [];
  rows.forEach(function (r) {
    const key = r.seriesOrder + '|' + r.series;
    if (!bySeries[key]) {
      bySeries[key] = { order: Number(r.seriesOrder), title: r.series, slug: slug_(r.series), lessons: [] };
      list.push(bySeries[key]);
    }
    bySeries[key].lessons.push([r.id, r.date, r.passage, r.heading]);
  });
  list.sort(function (a, b) { return a.order - b.order; });
  return { ok: true, updated: new Date().toISOString(), series: list };
}

/* ------------------------------------------------------------------ */
/* Rendering a Google Doc outline to safe HTML                         */
/* ------------------------------------------------------------------ */

const GLYPH_CSS_ = {
  NUMBER: 'decimal', LATIN_UPPER: 'upper-alpha', LATIN_LOWER: 'lower-alpha',
  ROMAN_UPPER: 'upper-roman', ROMAN_LOWER: 'lower-roman',
  BULLET: 'disc', HOLLOW_BULLET: 'circle', SQUARE_BULLET: 'square',
};
const DEFAULT_LEVEL_CSS_ = ['decimal', 'lower-alpha', 'lower-roman', 'decimal', 'lower-alpha', 'lower-roman', 'decimal', 'lower-alpha', 'lower-roman'];

function getLesson_(id) {
  const rows = getIndex_();
  let idx = -1;
  for (let i = 0; i < rows.length; i++) if (rows[i].id === id) { idx = i; break; }
  if (idx < 0) return null; // not in the lessons folder: never served

  const r = rows[idx];
  const cache = CacheService.getScriptCache();
  const key = 'lesson_' + id + '_' + r.modified;
  const hit = cacheGetBig_(cache, key);
  let rendered;
  if (hit) rendered = JSON.parse(hit);
  else {
    rendered = renderDoc_(id);
    cachePutBig_(cache, key, JSON.stringify(rendered), 21600);
  }

  const sameSeries = rows.filter(function (x) { return x.series === r.series; });
  const pos = sameSeries.findIndex(function (x) { return x.id === id; });
  const link = function (x) { return x ? { id: x.id, date: x.date, passage: x.passage, heading: x.heading } : null; };

  return {
    id: id, series: r.series, seriesSlug: slug_(r.series), date: r.date, passage: r.passage,
    heading: r.heading, html: rendered.full, sections: rendered.sectionCount,
    prev: link(sameSeries[pos - 1]), next: link(sameSeries[pos + 1]),
  };
}

/** Returns {full, preface, sections[], sectionCount}. Each section is self-contained HTML. */
function renderDoc_(id) {
  const body = DocumentApp.openById(id).getBody();
  const n = body.getNumChildren();

  // A section starts at each top-level list item whose first character is bold.
  const starts = [];
  for (let i = 0; i < n; i++) {
    const c = body.getChild(i);
    if (c.getType() === DocumentApp.ElementType.LIST_ITEM && c.asListItem().getNestingLevel() === 0) {
      const t = c.asListItem().editAsText();
      if (t.getText().trim() && t.isBold(firstNonSpace_(t.getText()))) starts.push(i);
    }
  }

  const full = renderRange_(body, 0, n);
  const preface = starts.length ? renderRange_(body, 0, starts[0]) : '';
  const sections = starts.map(function (s, k) { return renderRange_(body, s, k + 1 < starts.length ? starts[k + 1] : n); });
  return { full: full, preface: preface, sections: sections, sectionCount: sections.length };
}

function firstNonSpace_(s) {
  const m = s.search(/\S/);
  return m < 0 ? 0 : m;
}

/** Renders children [from, to). List numbering is counted from the top of the doc so it stays correct. */
function renderRange_(body, from, to) {
  const out = [];
  const counters = {};
  let depth = 0;
  const liOpen = [];
  const emit = function (s, i) { if (i >= from) out.push(s); };

  const closeTo = function (level, i) {
    while (depth > level + 1) {
      if (liOpen[depth - 1]) emit('</li>', i);
      emit('</ol>', i);
      depth--;
    }
  };
  const openTo = function (level, i) {
    while (depth < level + 1) {
      if (depth > 0 && !liOpen[depth - 1]) { emit('<li class="nomark">', i); liOpen[depth - 1] = true; }
      emit('<ol>', i);
      liOpen[depth] = false;
      depth++;
    }
  };

  if (from >= to) return '';
  for (let i = 0; i < to; i++) {
    if (i === from) { depth = 0; liOpen.length = 0; } // nothing before the range was emitted
    const child = body.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.LIST_ITEM) {
      const li = child.asListItem();
      const level = li.getNestingLevel();
      const listId = li.getListId();
      const arr = counters[listId] || (counters[listId] = []);
      arr[level] = (arr[level] || 0) + 1;
      arr.length = level + 1;

      closeTo(level, i);
      openTo(level, i);
      if (liOpen[level]) emit('</li>', i);
      let glyph = null;
      try { glyph = li.getGlyphType(); } catch (e) { glyph = null; }
      const css = (glyph && GLYPH_CSS_[String(glyph)]) || DEFAULT_LEVEL_CSS_[level] || 'decimal';
      emit('<li value="' + arr[level] + '" style="list-style-type:' + css + '">' + inlineHtml_(li), i);
      liOpen[level] = true;
      continue;
    }

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const text = p.getText();
      if (!text.trim() && p.getNumChildren() <= 1) continue; // ignore blank lines; keep lists intact
      closeTo(-1, i);
      if (i < from) continue;
      const h = String(p.getHeading());
      const tag = h === 'TITLE' || h === 'HEADING1' ? 'h2' : h === 'HEADING2' || h === 'SUBTITLE' ? 'h3' : /^HEADING/.test(h) ? 'h4' : 'p';
      const align = String(p.getAlignment()) === 'CENTER' ? ' class="center"' : '';
      out.push('<' + tag + align + '>' + inlineHtml_(p) + '</' + tag + '>');
      continue;
    }

    if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      closeTo(-1, i);
      emit('<hr>', i);
      continue;
    }

    if (type === DocumentApp.ElementType.TABLE) {
      closeTo(-1, i);
      if (i < from) continue;
      const t = child.asTable();
      const rowsHtml = [];
      for (let r = 0; r < t.getNumRows(); r++) {
        const row = t.getRow(r);
        const cells = [];
        for (let c = 0; c < row.getNumCells(); c++) cells.push('<td>' + escapeHtml_(row.getCell(c).getText()).replace(/\n/g, '<br>') + '</td>');
        rowsHtml.push('<tr>' + cells.join('') + '</tr>');
      }
      out.push('<table>' + rowsHtml.join('') + '</table>');
    }
  }
  // close everything (always emitted)
  while (depth > 0) {
    if (liOpen[depth - 1]) out.push('</li>');
    out.push('</ol>');
    depth--;
  }
  return out.join('');
}

function inlineHtml_(container) {
  let html = '';
  for (let k = 0; k < container.getNumChildren(); k++) {
    const el = container.getChild(k);
    const t = el.getType();
    if (t === DocumentApp.ElementType.TEXT) html += textHtml_(el.asText());
    else if (t === DocumentApp.ElementType.HORIZONTAL_RULE) html += '<hr>';
  }
  return html;
}

function textHtml_(t) {
  const s = t.getText();
  if (!s) return '';
  const idx = t.getTextAttributeIndices();
  let html = '';
  for (let i = 0; i < idx.length; i++) {
    const start = idx[i];
    const end = i + 1 < idx.length ? idx[i + 1] : s.length;
    const seg = s.substring(start, end);
    if (!seg) continue;
    let h = escapeHtml_(seg).replace(/[\r\n\u000b]/g, '<br>');
    if (t.isBold(start)) h = '<strong>' + h + '</strong>';
    if (t.isItalic(start)) h = '<em>' + h + '</em>';
    const url = t.getLinkUrl(start);
    if (url && /^(https?:|mailto:)/i.test(url)) h = '<a href="' + escapeHtml_(url) + '" target="_blank" rel="noopener">' + h + '</a>';
    else if (t.isUnderline(start)) h = '<u>' + h + '</u>';
    html += h;
  }
  return html;
}

/* ------------------------------------------------------------------ */
/* Web app                                                             */
/* ------------------------------------------------------------------ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  let res;
  try {
    switch (p.action) {
      case 'library': res = libraryPayload_(); break;
      case 'lesson': res = getLesson_(String(p.id || '')) || { ok: false, error: 'Lesson not found.' }; if (res.id) res.ok = true; break;
      case 'subscribe': res = subscribe_(p); break;
      case 'confirm': res = confirm_(String(p.token || '')); break;
      case 'unsubscribe': res = unsubscribe_(String(p.token || '')); break;
      default: res = { ok: true, name: CONFIG.SITE_NAME };
    }
  } catch (err) {
    console.error(err);
    res = { ok: false, error: 'Something went wrong. Please try again later.' };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Subscriptions                                                       */
/* A reading plan = one series, from a starting lesson to the end of   */
/* that series, one outline per email, daily or weekly. When it ends,   */
/* the reader is invited to choose a new plan.                          */
/* ------------------------------------------------------------------ */

const SUB_HEADERS_ = ['email', 'frequency', 'status', 'token', 'created', 'confirmed', 'pendingFrequency', 'pendingStart', 'lastConfirmSent', 'startLessonId', 'lastLessonId', 'nextDue', 'sentCount'];

function subscribe_(p) {
  if (p.website) return { ok: true, message: 'Thanks! Check your inbox to confirm.' }; // honeypot
  const email = String(p.email || '').trim().toLowerCase();
  const freq = p.frequency === 'daily' ? 'daily' : 'weekly';
  if (!/^[^\s@<>,;]+@[^\s@<>,;]+\.[a-z]{2,}$/i.test(email) || email.length > 254) {
    return { ok: false, error: 'Please enter a valid email address.' };
  }
  const start = resolveStart_(String(p.start || ''));
  if (!start) return { ok: false, error: 'The library isn’t available right now. Please try again later.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = subscribersSheet_();
    const rows = readRows_(sh);
    const now = new Date();
    let row = rows.find(function (r) { return r.email === email; });

    if (row && row.lastConfirmSent && now - new Date(row.lastConfirmSent) < 10 * 60 * 1000) {
      return { ok: true, message: 'We just sent you a confirmation email. Please check your inbox (and spam folder).' };
    }
    if (!takeConfirmationSlot_()) {
      return { ok: false, error: 'Sign-ups are temporarily paused. Please try again tomorrow.' };
    }

    if (!row) {
      row = { email: email, frequency: freq, status: 'pending', token: Utilities.getUuid().replace(/-/g, ''), created: now.toISOString(), sentCount: '0' };
      appendRow_(sh, row);
    } else if (row.status === 'unsubscribed') {
      row.status = 'pending';
    }
    row.pendingFrequency = freq;
    row.pendingStart = start.id;
    row.lastConfirmSent = now.toISOString();
    writeRow_(sh, row);

    const link = CONFIG.SITE_URL + '/#/confirm/' + row.token;
    const what = (freq === 'daily' ? 'one outline a day' : 'one outline a week') + ' through <strong>' + escapeHtml_(start.series) + '</strong>, starting with ' + escapeHtml_(start.passage);
    const html = emailShell_(
      '<p style="margin:0 0 16px">Please confirm your ' + CONFIG.SITE_NAME + ' reading plan: ' + what + '.</p>' +
      '<p style="margin:0 0 24px"><a href="' + link + '" style="background:#1f3a5f;color:#fff;padding:12px 20px;border-radius:6px;text-decoration:none;display:inline-block">Confirm and send my first lesson</a></p>' +
      '<p style="margin:0;color:#6b6358;font-size:13px">If you didn’t ask for this, just ignore this email and nothing will be sent.</p>', null);
    MailApp.sendEmail({ to: email, subject: 'Confirm your ' + CONFIG.SITE_NAME + ' reading plan', htmlBody: html, name: CONFIG.SITE_NAME, body: 'Confirm your reading plan: ' + link });
    return { ok: true, message: 'Almost done — check your inbox for a confirmation link.' };
  } finally {
    lock.releaseLock();
  }
}

/** start may be a lesson id, a series slug, or empty (= first lesson in the library). */
function resolveStart_(start) {
  const rows = getIndex_();
  if (!rows.length) return null;
  let hit = rows.find(function (r) { return r.id === start; });
  if (!hit && start) hit = rows.find(function (r) { return slug_(r.series) === start; });
  return hit || rows[0];
}

function confirm_(token) {
  if (!/^[a-f0-9]{32}$/.test(token)) return { ok: false, error: 'That confirmation link isn’t valid.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = subscribersSheet_();
    const row = readRows_(sh).find(function (r) { return r.token === token; });
    if (!row) return { ok: false, error: 'That confirmation link isn’t valid.' };
    if (!row.pendingStart && row.status !== 'pending') {
      return { ok: true, message: 'Your reading plan is already active.' };
    }
    row.status = 'active';
    if (row.pendingFrequency) row.frequency = row.pendingFrequency;
    if (row.pendingStart) { row.startLessonId = row.pendingStart; row.lastLessonId = ''; }
    row.pendingFrequency = '';
    row.pendingStart = '';
    row.confirmed = new Date().toISOString();
    row.nextDue = today_();
    writeRow_(sh, row);

    // Send the first lesson right away.
    let sentNow = false;
    try { sentNow = sendNext_(sh, row, {}); } catch (e) { console.error(e); }
    const start = getIndex_().find(function (r) { return r.id === row.startLessonId; });
    return {
      ok: true, frequency: row.frequency,
      message: 'You’re subscribed: one outline ' + (row.frequency === 'daily' ? 'a day' : 'a week') +
        (start ? ' through ' + start.series + ', starting with ' + start.passage : '') + '. ' +
        (sentNow ? 'Your first lesson is on its way now.' : 'Your first lesson will arrive tomorrow morning.'),
    };
  } finally {
    lock.releaseLock();
  }
}

function unsubscribe_(token) {
  if (!/^[a-f0-9]{32}$/.test(token)) return { ok: false, error: 'That link isn’t valid.' };
  const sh = subscribersSheet_();
  const row = readRows_(sh).find(function (r) { return r.token === token; });
  if (!row) return { ok: false, error: 'That link isn’t valid.' };
  row.status = 'unsubscribed';
  writeRow_(sh, row);
  return { ok: true, message: 'You’ve been unsubscribed. You won’t receive any more emails.' };
}

function takeConfirmationSlot_() {
  const props = PropertiesService.getScriptProperties();
  const key = 'confirms_' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
  const n = Number(props.getProperty(key) || 0);
  if (n >= CONFIG.MAX_CONFIRMATIONS_PER_DAY) return false;
  props.setProperty(key, String(n + 1));
  return true;
}

function subscribersSheet_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('Subscribers');
  // Add any missing columns (keeps older sheets working).
  const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].filter(String);
  SUB_HEADERS_.forEach(function (h) {
    if (have.indexOf(h) < 0) { sh.getRange(1, have.length + 1).setValue(h).setFontWeight('bold'); have.push(h); }
  });
  return sh;
}

/* ------------------------------------------------------------------ */
/* Scheduled emails                                                    */
/* ------------------------------------------------------------------ */

function sendScheduled_() {
  const hour = Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'H'));
  if (hour < CONFIG.SEND_HOUR) return;
  const today = today_();
  const sh = subscribersSheet_();
  const rendered = {};
  readRows_(sh).forEach(function (row) {
    if (row.status !== 'active' || !row.nextDue || row.nextDue > today) return;
    if (MailApp.getRemainingDailyQuota() < 1) return; // the rest go out next hour / tomorrow
    sendNext_(sh, row, rendered);
  });
}

/** The series a lesson belongs to, and whether lessons are still being added to it (it holds the newest dated lesson). */
function seriesInfo_(rows, seriesName) {
  let newest = null;
  rows.forEach(function (r) { if (r.date && (!newest || r.date > newest.date)) newest = r; });
  return { name: seriesName, ongoing: !!newest && newest.series === seriesName };
}

/**
 * Sends the subscriber's next lesson in their plan's series. When the series is finished,
 * sends a "choose your next plan" email instead and marks the plan completed
 * (unless the series is still being written, in which case they wait for new lessons).
 * Returns true if an email went out.
 */
function sendNext_(sh, row, renderedCache) {
  const rows = getIndex_();
  const start = rows.find(function (r) { return r.id === row.startLessonId; });
  if (!start) return false;
  const series = seriesInfo_(rows, start.series);

  let idx;
  if (row.lastLessonId) {
    const last = rows.findIndex(function (r) { return r.id === row.lastLessonId; });
    idx = last < 0 ? -1 : last + 1;
  } else {
    idx = rows.findIndex(function (r) { return r.id === row.startLessonId; });
  }
  if (idx < 0) return false;
  if (MailApp.getRemainingDailyQuota() < 1) return false;

  const lesson = rows[idx];
  if (!lesson || lesson.series !== series.name) {
    if (series.ongoing) return false; // caught up; new lessons will come as they're added
    sendCompletion_(row, series.name, rows);
    row.status = 'completed';
    row.nextDue = '';
    writeRow_(sh, row);
    return true;
  }

  const r = renderedCache[lesson.id] || (renderedCache[lesson.id] = renderDoc_(lesson.id));
  const n = Number(row.sentCount || 0) + 1;
  const inSeries = rows.filter(function (x) { return x.series === series.name; });
  const pos = inSeries.findIndex(function (x) { return x.id === lesson.id; }) + 1;
  const subject = lesson.passage + (lesson.heading ? ' — ' + lesson.heading : '');
  const nextUp = rows[idx + 1] && rows[idx + 1].series === series.name ? rows[idx + 1] : null;
  const when = row.frequency === 'daily' ? 'tomorrow' : 'next week';

  const intro = '<p style="margin:0 0 6px;color:#8a6d3b;font-size:13px;letter-spacing:.08em;text-transform:uppercase">' + escapeHtml_(series.name) + ' · Lesson ' + pos + ' of ' + inSeries.length + '</p>' +
    '<p style="margin:0 0 20px"><a href="' + lessonUrl_(lesson.id) + '" style="color:#1f3a5f">Read this lesson on ' + CONFIG.SITE_NAME + '</a></p>';
  let outro;
  if (nextUp) outro = 'Next ' + when + ': ' + escapeHtml_(nextUp.passage);
  else if (series.ongoing) outro = 'You’re caught up with the newest lesson in ' + escapeHtml_(series.name) + '. New outlines will come to you as they’re written.';
  else outro = 'That’s the last lesson in ' + escapeHtml_(series.name) + '. ' + (row.frequency === 'daily' ? 'Tomorrow' : 'Next week') + ' you’ll get a note inviting you to choose your next reading plan.';
  outro = '<p style="margin:24px 0 0;color:#6b6358;font-size:14px">' + outro + '</p>';

  MailApp.sendEmail({ to: row.email, subject: subject, name: CONFIG.SITE_NAME,
    htmlBody: emailShell_(intro + '<div style="font-family:Georgia,serif;font-size:16px;line-height:1.55">' + r.full + '</div>' + outro, planFooter_(row)) });

  row.lastLessonId = lesson.id;
  row.sentCount = String(n);
  row.nextDue = addDays_(today_(), row.frequency === 'daily' ? 1 : 7);
  writeRow_(sh, row);
  return true;
}

function sendCompletion_(row, finishedSeries, rows) {
  const seen = {};
  const series = [];
  rows.forEach(function (r) {
    if (seen[r.series]) { seen[r.series].count++; return; }
    seen[r.series] = { title: r.series, order: Number(r.seriesOrder), count: 1 };
    series.push(seen[r.series]);
  });
  const done = seen[finishedSeries];
  const others = series.filter(function (x) { return x.title !== finishedSeries && x.order < 1000; });
  // Suggest the series that comes next in the library, wrapping around.
  const suggestion = others.find(function (x) { return done && x.order > done.order; }) || others[0];

  const btn = function (href, label) { return '<a href="' + href + '" style="background:#1f3a5f;color:#fff;padding:12px 20px;border-radius:6px;text-decoration:none;display:inline-block">' + label + '</a>'; };
  let html = '<p style="margin:0 0 6px;color:#8a6d3b;font-size:13px;letter-spacing:.08em;text-transform:uppercase">Reading plan complete</p>' +
    '<p style="margin:0 0 14px;font-size:22px;color:#1f3a5f">You’ve finished ' + escapeHtml_(finishedSeries) + '.</p>' +
    '<p style="margin:0 0 20px">Well done. Every lesson in the series has come your way' + (done ? ', all ' + done.count : '') + '. Would you like to keep going with another series?</p>';
  if (suggestion) {
    html += '<p style="margin:0 0 24px">' + btn(CONFIG.SITE_URL + '/#/subscribe/' + slug_(suggestion.title), 'Start ' + escapeHtml_(suggestion.title)) + '</p>';
  }
  html += '<p style="margin:0 0 8px">Or choose any series:</p><ul style="margin:0 0 20px;padding-left:20px">' +
    others.map(function (x) {
      return '<li style="margin:4px 0"><a href="' + CONFIG.SITE_URL + '/#/subscribe/' + slug_(x.title) + '" style="color:#1f3a5f">' + escapeHtml_(x.title) + '</a> <span style="color:#6b6358">(' + x.count + ' lessons)</span></li>';
    }).join('') + '</ul>' +
    '<p style="margin:0;color:#6b6358;font-size:14px">If you’d rather stop here, there’s nothing to do. No more emails will come unless you start a new plan.</p>';

  MailApp.sendEmail({ to: row.email, subject: 'You finished ' + finishedSeries + '. Choose your next reading plan', name: CONFIG.SITE_NAME,
    htmlBody: emailShell_(html, 'You received this because you finished a ' + CONFIG.SITE_NAME + ' reading plan.') });
}

function planFooter_(row) {
  return 'You’re getting one outline ' + (row.frequency === 'daily' ? 'a day' : 'a week') + '. ' +
    '<a href="' + CONFIG.SITE_URL + '/#/subscribe" style="color:#6b6358">Change your plan</a> · ' +
    '<a href="' + CONFIG.SITE_URL + '/#/unsubscribe/' + row.token + '" style="color:#6b6358">Unsubscribe</a>';
}

function emailShell_(inner, footer) {
  return '<div style="background:#f6f1e7;padding:24px 12px">' +
    '<div style="max-width:640px;margin:0 auto;background:#fffdf8;border:1px solid #e6dccb;border-radius:10px;padding:28px 28px 20px;color:#2b2620;font-family:Georgia,serif">' +
    '<p style="margin:0 0 20px;font-size:20px;color:#1f3a5f"><a href="' + CONFIG.SITE_URL + '" style="color:#1f3a5f;text-decoration:none">' + CONFIG.SITE_NAME + '</a></p>' +
    inner +
    (footer ? '<hr style="border:none;border-top:1px solid #e6dccb;margin:28px 0 12px"><p style="margin:0;font-size:12px;color:#6b6358;font-family:Arial,sans-serif">' + footer + '</p>' : '') +
    '</div></div>';
}

function today_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function readRows_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    const o = { _row: i + 1 };
    headers.forEach(function (h, j) { o[h] = values[i][j]; });
    rows.push(o);
  }
  return rows;
}

function appendRow_(sh, obj) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(function (h) { return obj[h] || ''; }));
  obj._row = sh.getLastRow();
}

function writeRow_(sh, obj) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.getRange(obj._row, 1, 1, headers.length).setNumberFormat('@').setValues([headers.map(function (h) { return obj[h] == null ? '' : String(obj[h]); })]);
}

function cachePutBig_(cache, key, str, ttl) {
  const size = 90000;
  const parts = Math.ceil(str.length / size) || 1;
  const map = {};
  for (let i = 0; i < parts; i++) map[key + '_' + i] = str.substr(i * size, size);
  map[key] = String(parts);
  try { cache.putAll(map, ttl); } catch (e) { /* too big to cache; fine */ }
}

function cacheGetBig_(cache, key) {
  const n = Number(cache.get(key) || 0);
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  const got = cache.getAll(keys);
  let s = '';
  for (let i = 0; i < n; i++) { if (got[keys[i]] == null) return null; s += got[keys[i]]; }
  return s;
}

function addDays_(ymd, n) {
  const d = new Date(ymd + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function prettyDate_(ymd) {
  if (!ymd) return '';
  return Utilities.formatDate(new Date(ymd + 'T12:00:00Z'), 'UTC', 'MMMM d, yyyy');
}

function lessonUrl_(id) { return CONFIG.SITE_URL + '/#/lesson/' + id; }
function slug_(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, ''); }
function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/* Manual test helpers (run from the editor) */
function testWeeklyEmailToMe() {
  const lesson = getIndex_().filter(function (r) { return r.date; }).sort(function (a, b) { return a.date < b.date ? -1 : 1; }).pop();
  const r = renderDoc_(lesson.id);
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: '[Test] This Sunday: ' + lesson.passage, htmlBody: emailShell_(r.full, 'Test email'), name: CONFIG.SITE_NAME });
}
