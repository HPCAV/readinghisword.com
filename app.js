(function () {
  'use strict';

  var API = window.RHW_API;
  var app = document.getElementById('app');
  var library = null;
  var libraryPromise = null;
  var lessonCache = {};
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];


  /* ---------- helpers ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function dash(s) { return String(s || '').replace(/(\d)\s*-\s*(\d)/g, '$1–$2'); }
  function prettyDate(ymd, short) {
    if (!ymd) return '';
    var p = ymd.split('-');
    var m = MONTHS[Number(p[1]) - 1];
    return (short ? m.slice(0, 3) : m) + ' ' + Number(p[2]) + ', ' + p[0];
  }
  function api(params) {
    var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); }).join('&');
    return fetch(API + '?' + qs, { redirect: 'follow' }).then(function (r) {
      if (!r.ok) throw new Error('Network error');
      return r.json();
    });
  }
  function store(key, val) { try { sessionStorage.setItem(key, JSON.stringify(val)); } catch (e) {} }
  function recall(key) { try { return JSON.parse(sessionStorage.getItem(key)); } catch (e) { return null; } }
  function loading() { app.innerHTML = '<div class="loading"><span></span><span></span><span></span></div>'; }
  function errorView(msg) {
    app.innerHTML = '<div class="card notice"><h1>Sorry</h1><p>' + esc(msg || 'The library couldn’t be loaded right now. Please try again in a moment.') + '</p><p><a class="btn ghost" href="#/">Back to the library</a></p></div>';
  }
  function setTitle(t) { document.title = t ? t + ' — Reading His Word' : 'Reading His Word — Sunday School Lesson Outlines'; }
  function setNav(name) {
    Array.prototype.forEach.call(document.querySelectorAll('.nav a'), function (a) {
      a.classList.toggle('active', a.getAttribute('href') === '#/' + (name || ''));
    });
  }

  function getLibrary() {
    if (library) return Promise.resolve(library);
    var cached = recall('rhw_library');
    if (cached && cached.series && Date.now() - cached._t < 30 * 60 * 1000) { library = cached; return Promise.resolve(library); }
    if (!libraryPromise) {
      libraryPromise = api({ action: 'library' }).then(function (d) {
        if (!d.ok) throw new Error(d.error);
        d._t = Date.now();
        d.series.forEach(function (s) {
          s.lessons = s.lessons.map(function (l) { return { id: l[0], date: l[1], passage: dash(l[2]), heading: l[3], series: s.title, slug: s.slug }; });
        });
        library = d;
        store('rhw_library', d);
        return d;
      }).catch(function (e) { libraryPromise = null; throw e; });
    }
    return libraryPromise;
  }

  function allLessons(lib) {
    var out = [];
    lib.series.forEach(function (s) { out = out.concat(s.lessons); });
    return out;
  }

  function lessonRow(l, showSeries) {
    return '<li><a class="lesson-row" href="#/lesson/' + esc(l.id) + '">' +
      '<span class="date">' + esc(prettyDate(l.date, true)) + '</span>' +
      '<span class="title">' + (showSeries ? '<span class="series-name">' + esc(l.series) + '</span>' : '') +
      '<strong>' + esc(l.passage) + '</strong>' + (l.heading && l.heading !== l.passage ? ' <span>— ' + esc(l.heading) + '</span>' : '') + '</span>' +
      '<span class="arrow" aria-hidden="true">→</span></a></li>';
  }

  function subscribeBlock(lib, startLesson, startSlug) {
    var opts = '';
    if (startLesson) opts += '<option value="' + esc(startLesson.id) + '" selected>From this lesson: ' + esc(startLesson.passage) + ' (' + esc(startLesson.series) + ')</option>';
    if (lib) opts += lib.series.filter(function (s) { return s.order < 1000; }).map(function (s, i) {
      var sel = !startLesson && (startSlug ? s.slug === startSlug : i === 0);
      return '<option value="' + esc(s.slug) + '"' + (sel ? ' selected' : '') + '>' + (s.order < 900 ? esc(s.order) + '. ' : '') + esc(s.title) + ' (' + s.lessons.length + ' lessons)</option>';
    }).join('');
    return '<section class="card subscribe" id="subscribe">' +
      '<div><span class="eyebrow">Free reading plan by email</span><h2>Work through the outlines at your own pace</h2>' +
      '<p>Choose a series and how often. You’ll get its outlines one at a time, in order, one a day or one a week. Your first lesson arrives as soon as you confirm. When you finish a series, we’ll invite you to pick your next one.</p></div>' +
      '<form class="sub-form" novalidate>' +
      '<input type="email" name="email" placeholder="you@example.com" autocomplete="email" required aria-label="Email address">' +
      '<label class="field"><span>Reading plan</span><select name="start">' + opts + '</select></label>' +
      '<div class="freq">' +
      '<label><input type="radio" name="frequency" value="weekly" checked>Weekly<small>One outline each week</small></label>' +
      '<label><input type="radio" name="frequency" value="daily">Daily<small>One outline each morning</small></label>' +
      '</div>' +
      '<input class="hp" type="text" name="website" tabindex="-1" autocomplete="off" aria-hidden="true">' +
      '<button class="btn" type="submit">Start my reading plan</button>' +
      '<div class="form-msg" role="status"></div>' +
      '</form></section>';
  }

  function wireSubscribe() {
    var form = app.querySelector('.sub-form');
    if (!form) return;
    form.addEventListener('submit', function (ev) {
      ev.preventDefault();
      var msg = form.querySelector('.form-msg');
      var btn = form.querySelector('button');
      var email = form.email.value.trim();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) { msg.className = 'form-msg err'; msg.textContent = 'Please enter a valid email address.'; return; }
      btn.disabled = true; btn.textContent = 'Sending…'; msg.textContent = '';
      api({ action: 'subscribe', email: email, frequency: form.querySelector('input[name=frequency]:checked').value, start: form.start ? form.start.value : '', website: form.website.value })
        .then(function (d) {
          msg.className = 'form-msg ' + (d.ok ? 'ok' : 'err');
          msg.textContent = d.ok ? d.message : d.error;
          if (d.ok) form.email.value = '';
        })
        .catch(function () { msg.className = 'form-msg err'; msg.textContent = 'Couldn’t reach the server. Please try again.'; })
        .then(function () { btn.disabled = false; btn.textContent = 'Start my reading plan'; });
    });
  }

  /* ---------- views ---------- */
  function homeView() {
    setTitle(''); setNav('');
    loading();
    getLibrary().then(function (lib) {
      var lessons = allLessons(lib).filter(function (l) { return l.date; });
      lessons.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
      var latest = lessons[0];
      var current = lib.series[0];

      var html = '<section class="hero">' +
        '<div class="hero-text"><span class="eyebrow">Free to read · Free to share · Public domain</span>' +
        '<h1>Reading His Word,<br>verse by verse</h1>' +
        '<p>A growing library of Bible study outlines that walk through Scripture one verse at a time, with the original-language notes and cross-references used in class.</p>' +
        '<a class="btn" href="#/subscribe">Get the outlines by email</a></div>';
      if (latest) {
        html += '<a class="card feature" href="#/lesson/' + esc(latest.id) + '">' +
          '<div><span class="badge">Latest lesson</span>' +
          '<h2>' + esc(latest.heading || latest.passage) + '</h2>' +
          '<div class="passage">' + esc(latest.passage) + '</div>' +
          '<div class="muted" style="margin-top:6px;font-size:14px">' + esc(latest.series) + ' · ' + esc(prettyDate(latest.date)) + '</div></div>' +
          '<div class="go">Read the outline →</div></a>';
      }
      html += '</section>';

      html += '<div class="toolbar"><h2>The Library</h2>' +
        '<label class="search"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>' +
        '<input type="search" id="q" placeholder="Search lessons, books, passages…" aria-label="Search lessons"></label></div>' +
        '<div id="results"></div>';

      html += '<div class="grid" id="series-grid">' + lib.series.map(function (s) {
        var dated = s.lessons.filter(function (l) { return l.date; });
        var range = dated.length ? dated[0].date.slice(0, 4) + (dated[dated.length - 1].date.slice(0, 4) !== dated[0].date.slice(0, 4) ? '–' + dated[dated.length - 1].date.slice(0, 4) : '') : '';
        return '<a class="card series-card" href="#/series/' + esc(s.slug) + '">' +
          (s === current ? '<span class="badge">Current</span>' : '') +
          '<div class="num">' + (s.order < 900 ? esc(s.order) : '•') + '</div>' +
          '<h3>' + esc(s.title) + '</h3>' +
          '<div class="meta">' + s.lessons.length + ' lesson' + (s.lessons.length === 1 ? '' : 's') + (range ? ' · ' + range : '') + '</div></a>';
      }).join('') + '</div>';

      html += subscribeBlock(lib);
      app.innerHTML = html;
      wireSubscribe();

      var q = document.getElementById('q');
      var results = document.getElementById('results');
      var grid = document.getElementById('series-grid');
      var all = allLessons(lib);
      q.addEventListener('input', function () {
        var term = q.value.trim().toLowerCase();
        if (!term) { results.innerHTML = ''; grid.style.display = ''; return; }
        var words = term.split(/\s+/);
        var hits = all.filter(function (l) {
          var hay = (l.passage + ' ' + l.heading + ' ' + l.series + ' ' + prettyDate(l.date)).toLowerCase();
          return words.every(function (w) { return hay.indexOf(w) >= 0; });
        }).slice(0, 80);
        grid.style.display = 'none';
        results.innerHTML = hits.length
          ? '<ul class="card lesson-list" style="margin-bottom:10px">' + hits.map(function (l) { return lessonRow(l, true); }).join('') + '</ul>'
          : '<div class="card empty">No lessons match “' + esc(q.value) + '”.</div>';
      });
    }).catch(function () { errorView(); });
  }

  function seriesView(slug) {
    setNav('');
    loading();
    getLibrary().then(function (lib) {
      var s = lib.series.filter(function (x) { return x.slug === slug; })[0];
      if (!s) return errorView('That series wasn’t found.');
      setTitle(s.title);
      var lessons = s.lessons.slice();
      app.innerHTML = '<div class="page-head"><div class="crumbs"><a href="#/">Library</a> › Series ' + esc(s.order < 900 ? s.order : '') + '</div>' +
        '<h1>' + esc(s.title) + '</h1><p>' + lessons.length + ' lesson' + (lessons.length === 1 ? '' : 's') + '</p></div>' +
        (lessons.length ? '<ul class="card lesson-list">' + lessons.map(function (l) { return lessonRow(l, false); }).join('') + '</ul>' : '<div class="card empty">No lessons yet.</div>');
      window.scrollTo(0, 0);
    }).catch(function () { errorView(); });
  }

  function lessonView(id) {
    setNav('');
    loading();
    window.scrollTo(0, 0);
    var p = lessonCache[id] ? Promise.resolve(lessonCache[id]) : api({ action: 'lesson', id: id });
    p.then(function (d) {
      if (!d.ok) return errorView(d.error || 'That lesson wasn’t found.');
      lessonCache[id] = d;
      d.passage = dash(d.passage);
      if (d.prev) d.prev.passage = dash(d.prev.passage);
      if (d.next) d.next.passage = dash(d.next.passage);
      setTitle(d.passage + (d.heading ? ' — ' + d.heading : ''));
      var pn = '';
      if (d.prev) pn += '<a href="#/lesson/' + esc(d.prev.id) + '"><small>← Previous</small>' + esc(d.prev.passage) + '</a>';
      if (d.next) pn += '<a href="#/lesson/' + esc(d.next.id) + '"><small>Next →</small>' + esc(d.next.passage) + '</a>';
      app.innerHTML = '<div class="page-head"><div class="crumbs"><a href="#/">Library</a> › <a href="#/series/' + esc(d.seriesSlug) + '">' + esc(d.series) + '</a></div>' +
        '<h1>' + esc(d.passage) + '</h1><p>' + esc(prettyDate(d.date)) + '</p></div>' +
        '<div class="reader"><article class="card outline">' + d.html + '</article>' +
        '<aside class="side">' +
        (pn ? '<div class="card pn"><h4>In this series</h4>' + pn + '</div>' : '') +
        '<div class="card"><h4>Use this lesson</h4><button class="btn ghost block" id="print">Print outline</button></div>' +
        '<div class="card"><h4>Reading plan</h4><p style="margin:0 0 12px;font-size:14px" class="muted">Read the rest of this series by email, starting here, one a day or one a week.</p><a class="btn block" href="#/subscribe/' + esc(d.id) + '">Start from this lesson</a></div>' +
        '</aside></div>';
      document.getElementById('print').addEventListener('click', function () { window.print(); });
      // Warm the cache for the next lesson.
      if (d.next && !lessonCache[d.next.id]) api({ action: 'lesson', id: d.next.id }).then(function (n) { if (n.ok) lessonCache[n.id] = n; }).catch(function () {});
    }).catch(function () { errorView('That lesson couldn’t be loaded right now. Please try again in a moment.'); });
  }

  function subscribeView(startId) {
    setTitle('Reading plan'); setNav('subscribe');
    loading();
    getLibrary().then(function (lib) {
      var startLesson = startId ? allLessons(lib).filter(function (l) { return l.id === startId; })[0] : null;
      var startSlug = !startLesson && startId ? startId : '';
      app.innerHTML = '<div class="page-head"><div class="crumbs"><a href="#/">Library</a> › Reading plan</div><h1>Outlines by email</h1>' +
        '<p>Free, and you can unsubscribe with one click from any email.</p></div>' + subscribeBlock(lib, startLesson, startSlug);
      app.querySelector('.subscribe').style.marginTop = '0';
      wireSubscribe();
      var input = app.querySelector('input[type=email]');
      if (input) input.focus();
    }).catch(function () { errorView(); });
  }

  function tokenView(action, token) {
    setTitle(action === 'confirm' ? 'Confirm subscription' : 'Unsubscribe'); setNav('');
    loading();
    api({ action: action, token: token }).then(function (d) {
      var title = d.ok ? (action === 'confirm' ? 'You’re subscribed' : 'Unsubscribed') : 'Hmm, that didn’t work';
      app.innerHTML = '<div class="card notice"><h1>' + esc(title) + '</h1><p>' + esc(d.ok ? d.message : d.error) + '</p>' +
        '<p style="margin-top:22px"><a class="btn" href="#/">Browse the library</a></p></div>';
    }).catch(function () { errorView('Couldn’t reach the server. Please try the link again in a moment.'); });
  }

  function aboutView() {
    setTitle('About'); setNav('about');
    app.innerHTML = '<article class="card prose"><h1>About</h1>' +
      '<p>Reading His Word is a library of Sunday School lesson outlines written for an adult class at Southview Baptist Church in Lincoln, Nebraska. Each outline works through a passage verse by verse, noting the verse references along the way, with occasional Hebrew and Greek word studies and cross-references.</p>' +
      '<p>The series run through Genesis, the history of Israel, the Minor Prophets, the Gospels, and the Epistles. New lessons appear here automatically as they’re written.</p>' +
      '<p>To study along, <a href="#/subscribe">start a reading plan</a>: choose a series, and its outlines come to your inbox one at a time, daily or weekly. When you finish, you’ll be invited to choose another.</p>' +
      '<h2 class="about-h">About the author</h2>' +
      '<div class="author"><img src="dick-clark.jpg" alt="Dick Clark" width="220" height="220">' +
      '<div><p><strong>Dick Clark</strong> trusted Christ as his Savior on January 27, 1986, shortly before his fifth birthday. He grew up in Slidell, Louisiana, the oldest of five sons of two New Orleans Baptist Theological Seminary graduates who raised their children to be active in the life of the church.</p>' +
      '<p>His first ministry was playing violin for special music at a small country church, and music has remained part of his service ever since. For years he and his family led monthly worship services for the residents of a Lincoln nursing and rehabilitation center.</p>' +
      '<p>At Southview Baptist Church he served as a deacon, led the church’s safety and security ministry, and taught an adult Sunday School class. The outlines in this library come from that class, and he continues to write them each week. He is a co-author of <em>Faith Seeking Freedom: Updated &amp; Expanded</em> (Libertarian Christian Institute, 2026), available on <a href="https://www.amazon.com/dp/B0GZV4J2FS" target="_blank" rel="noopener">Amazon</a>.</p>' +
      '<p>Dick and his wife have been married nearly nineteen years and have three children.</p></div></div>' +
      '<h2 class="about-h">Public domain</h2>' +
      '<p>Everything on this site, including every lesson outline and the site itself, is released into the public domain. You are free to use, copy, print, adapt, translate and distribute it, for any purpose, without asking permission or giving credit. Where a public-domain dedication isn’t legally recognized, the <a href="https://creativecommons.org/publicdomain/zero/1.0/" target="_blank" rel="noopener">CC0 1.0 Universal</a> waiver applies.</p>' +
      '</article>';
  }

  /* ---------- router ---------- */
  function route() {
    var h = location.hash.replace(/^#\/?/, '');
    var parts = h.split('/');
    if (!API || API.indexOf('PASTE_') === 0) return errorView('The site is being set up. Please check back soon.');
    switch (parts[0]) {
      case 'series': return seriesView(parts[1]);
      case 'lesson': return lessonView(parts[1]);
      case 'subscribe': return subscribeView(parts[1]);
      case 'about': return aboutView();
      case 'confirm': return tokenView('confirm', parts[1] || '');
      case 'unsubscribe': return tokenView('unsubscribe', parts[1] || '');
      default: return homeView();
    }
  }
  window.addEventListener('hashchange', route);
  route();
})();
