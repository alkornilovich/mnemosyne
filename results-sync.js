/**
 * results-sync.js: отправка результатов студента в Google Таблицу преподавателя.
 * Подключается в каждую игру одним <script>. Работает на любом статическом хостинге.
 *
 *   Sync.connect(url, name)  -> Promise<{ok:true}>  проверяет ссылку и сохраняет вход
 *   Sync.skip()                                     свободная игра, ничего не отправляется
 *   Sync.report({game, block, task, correct, attempts, seconds, score, extra})
 *   Sync.status()            -> {mode:'class'|'free'|'none', name, pending}
 *   Sync.logout()
 */
(function (root) {
  var KEY = 'mn_sync_v1';
  var URL_RE = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/;
  var memory = {};
  var flushing = false;

  function load() {
    try { var s = root.localStorage.getItem(KEY); if (s) return JSON.parse(s); } catch (e) {}
    return memory.state || { mode: 'none', url: '', name: '', queue: [] };
  }
  function save(st) {
    memory.state = st;
    try { root.localStorage.setItem(KEY, JSON.stringify(st)); } catch (e) {}
  }
  function uid() {
    return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }
  function post(url, payload) {
    // text/plain: простой запрос без preflight, Apps Script его принимает
    return root.fetch(url, {
      method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    }).then(function (r) { return r.json(); });
  }

  var Sync = {
    parseUrl: function (raw) {
      var u = String(raw || '').trim();
      return URL_RE.test(u) ? u : null;
    },
    connect: function (rawUrl, name) {
      var url = Sync.parseUrl(rawUrl);
      var nm = String(name || '').trim();
      if (!url) return Promise.reject(new Error('bad_url'));
      if (nm.length < 2) return Promise.reject(new Error('bad_name'));
      return root.fetch(url + '?ping=1', { redirect: 'follow' })
        .then(function (r) { return r.json(); })
        .then(function (j) {
          if (!j || j.ok !== true) throw new Error('not_results_link');
          var st = load();
          save({ mode: 'class', url: url, name: nm, queue: st.queue || [] });
          return { ok: true };
        }, function () { throw new Error('network'); });
    },
    skip: function () { var st = load(); st.mode = 'free'; save(st); },
    logout: function () { save({ mode: 'none', url: '', name: '', queue: [] }); },
    status: function () { var st = load(); return { mode: st.mode, name: st.name, pending: (st.queue || []).length }; },
    report: function (ev) {
      var st = load();
      if (st.mode !== 'class') return Promise.resolve({ sent: false, reason: st.mode });
      var e = ev || {};
      st.queue = (st.queue || []).concat([{
        id: uid(), student: st.name, game: e.game, block: e.block, task: e.task,
        correct: !!e.correct, attempts: e.attempts, seconds: e.seconds, score: e.score, extra: e.extra,
        topic: e.topic, mistakes: e.mistakes
      }]).slice(-500);
      save(st);
      return Sync.flush();
    },
    flush: function () {
      var st = load();
      if (st.mode !== 'class' || !(st.queue || []).length) return Promise.resolve({ sent: true, left: 0 });
      if (flushing) return Promise.resolve({ sent: false, reason: 'busy' });
      flushing = true;
      var batch = st.queue.slice(0, 50);
      return post(st.url, { events: batch }).then(function (j) {
        flushing = false;
        if (!j || j.ok !== true) return { sent: false, reason: (j && j.error) || 'server' };
        var cur = load();
        var done = {}; batch.forEach(function (b) { done[b.id] = 1; });
        cur.queue = (cur.queue || []).filter(function (q) { return !done[q.id]; });
        save(cur);
        return cur.queue.length ? Sync.flush() : { sent: true, left: 0 };
      }, function () {
        flushing = false;
        return { sent: false, reason: 'network' };
      });
    }
  };
  if (root.addEventListener) root.addEventListener('online', function () { Sync.flush(); });
  root.Sync = Sync;
  if (typeof module !== 'undefined') module.exports = Sync;
})(typeof window !== 'undefined' ? window : globalThis);
