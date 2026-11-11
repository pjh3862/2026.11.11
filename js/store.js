/* 브라우저 저장소 — localStorage 를 쓰되, 막혀 있거나 용량이 차면 메모리로만 동작합니다 (data09-25 에서 옮겨 옴) */
(function (root) {
  'use strict';
  var KEY_DB = 'data10-05.db';
  var KEY_UI = 'data10-05.ui';
  var KEY_OLD = 'data10-05.db.v1'; // 3차 개편 전(1·2차) 자료를 지우지 않고 따로 둠
  var memory = {};
  var ok = true;
  var lastError = '';
  function get(k) {
    try { var v = root.localStorage.getItem(k); return v == null && memory[k] != null ? memory[k] : v; }
    catch (e) { ok = false; return memory[k] == null ? null : memory[k]; }
  }
  function set(k, v) {
    memory[k] = v;
    try { root.localStorage.setItem(k, v); lastError = ''; return true; }
    catch (e) {
      ok = false;
      lastError = (e && (e.name === 'QuotaExceededError' || e.code === 22)) ? 'quota' : 'blocked';
      return false;
    }
  }
  function del(k) {
    delete memory[k];
    try { root.localStorage.removeItem(k); } catch (e) { ok = false; }
  }
  root.PPStore = {
    loadDb: function () {
      var raw = get(KEY_DB);
      if (!raw) return root.PPLogic.emptyDb();
      var o;
      try { o = JSON.parse(raw); } catch (e) { var d0 = root.PPLogic.emptyDb(); d0.notice = { from: 0, had: 0, items: 0, broken: true }; if (!get(KEY_OLD)) set(KEY_OLD, raw); return d0; }
      /* 옛 판(1·2차) 자료면 원래 글을 그대로 따로 보관하고(설정에서 JSON 으로 받을 수 있음), 새 형식으로 옮긴 것을 씁니다 */
      if (o && o.schemaVersion !== root.PPLogic.SCHEMA_VERSION && !get(KEY_OLD)) set(KEY_OLD, raw);
      try { return root.PPLogic.restoreDb(o); } catch (e) { return root.PPLogic.emptyDb(); }
    },
    /* true = 브라우저에 저장됨, false = 이번 창의 메모리에만 있음 */
    saveDb: function (db) { return set(KEY_DB, JSON.stringify(db)); },
    clearDb: function () { del(KEY_DB); },
    /* 3차 개편 전 자료(원래 JSON 글) — 없으면 null */
    loadOld: function () { return get(KEY_OLD); },
    clearOld: function () { del(KEY_OLD); },
    /* 입력 화면 편의값(마지막 공정·작업자수 등) — 없어도 동작 */
    loadUi: function () { try { return JSON.parse(get(KEY_UI) || '{}') || {}; } catch (e) { return {}; } },
    saveUi: function (o) { set(KEY_UI, JSON.stringify(o || {})); },
    lastError: function () { return lastError; },
    sizeKb: function () { var v = get(KEY_DB); return v ? Math.round(v.length * 2 / 1024) : 0; },
    available: function () { get(KEY_DB); return ok; }
  };
})(window);
