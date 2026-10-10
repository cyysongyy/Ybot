// ==UserScript==
// @name         Ybot：WebITR 差假同步（只在這台電腦）
// @namespace    https://cyysongyy.github.io/Ybot/
// @version      1.0.0
// @description  打開 WebITR 批核案件頁時，把清單交給同一台電腦上的 Ybot。只讀取表格：不登入、不按任何按鈕、資料不上傳。
// @match        https://webitr.yunlin.gov.tw/*
// @match        https://cyysongyy.github.io/Ybot/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addValueChangeListener
// @run-at       document-idle
// @updateURL    https://cyysongyy.github.io/Ybot/webitr-ybot.user.js
// @downloadURL  https://cyysongyy.github.io/Ybot/webitr-ybot.user.js
// ==/UserScript==

// 做的事只有兩件：
// 1. 在 WebITR 頁面：找到「申請人」那張表，把文字存進 Tampermonkey 自己的儲存空間。
// 2. 在 Ybot 頁面：把存好的清單交給 Ybot 整理。
// 兩邊都在同一台電腦、同一個 Chrome 裡，中間不經過任何伺服器。
// 刻意不做：自動登入、自動按「同意／不同意」、把資料送到網路上。
(function () {
  'use strict';
  const KEY = 'snapshots';
  const KEEP_MS = 21 * 864e5;   // 留三週：核准後案件會從批核清單消失，靠累積的快照才知道之後誰請假
  const MAX_SNAPSHOTS = 60;
  const MAX_TEXT = 20000;
  const HEAD_RE = /申請人|姓名|請假人|差假人/;
  const EVT = 'ybot-webitr-leave';
  const REQ = 'ybot-webitr-leave-request';

  function load() {
    try {
      const a = JSON.parse(GM_getValue(KEY, '[]'));
      return Array.isArray(a) ? a : [];
    } catch (e) { return []; }
  }
  function prune(list) {
    const cut = Date.now() - KEEP_MS;
    return list.filter(s => s && typeof s.text === 'string' && s.text && Date.parse(s.at) >= cut).slice(-MAX_SNAPSHOTS);
  }

  // ── Ybot 這一端 ──────────────────────────────
  if (location.hostname === 'cyysongyy.github.io') {
    const give = () => document.dispatchEvent(new CustomEvent(EVT, {
      detail: JSON.stringify({ v: 1, snapshots: prune(load()) })
    }));
    // Ybot 可能比外掛早載入，也可能比較晚：兩邊都會喊一聲，誰先誰後都接得上
    document.addEventListener(REQ, give);
    give();
    // WebITR 在另一個分頁更新時，開著的 Ybot 直接跟著變
    if (typeof GM_addValueChangeListener === 'function') GM_addValueChangeListener(KEY, give);
    return;
  }

  // ── WebITR 這一端 ────────────────────────────
  const cellText = c => (c.innerText || c.textContent || '').replace(/\s+/g, ' ').trim();
  const rowsOf = t => Array.from(t.rows || []).map(r => Array.from(r.cells).map(cellText));
  const visible = el => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
  // 表頭格要短（「申請人」），排版用的大外框表格整格文字很長，不會被誤認
  const isHeadRow = r => r.some(c => c.length <= 10 && HEAD_RE.test(c));

  function grab() {
    const tables = Array.from(document.querySelectorAll('table')).filter(visible);
    for (let i = 0; i < tables.length; i++) {
      const rows = rowsOf(tables[i]);
      const hi = rows.findIndex(isHeadRow);
      if (hi < 0) continue;
      let out = rows.slice(hi);
      // 有些系統把表頭和內容拆成兩個 table（表頭固定、內容捲動）。
      // 表頭這張沒有資料列時，接上後面欄數相近的那一張。
      if (out.length === 1) {
        const n = out[0].length;
        for (let j = i + 1; j < Math.min(tables.length, i + 4); j++) {
          if (tables[j].contains(tables[i]) || tables[i].contains(tables[j])) continue;
          const more = rowsOf(tables[j]).filter(r => r.length && Math.abs(r.length - n) <= 1);
          if (more.length) { out = out.concat(more); break; }
        }
      }
      return { text: out.map(r => r.join('\t')).join('\n').slice(0, MAX_TEXT), count: out.length - 1 };
    }
    return null;
  }

  let toastEl = null, toastTimer = null;
  function toast(msg) {
    if (!toastEl) {
      toastEl = document.createElement('div');
      // 不吃滑鼠點擊，絕對不會擋到 WebITR 的按鈕
      toastEl.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;pointer-events:none;'
        + 'background:#4c1d95;color:#fff;font:14px/1.4 sans-serif;padding:8px 14px;border-radius:10px;'
        + 'box-shadow:0 4px 14px rgba(0,0,0,.25);opacity:0;transition:opacity .3s';
      document.body.appendChild(toastEl);
    }
    toastEl.textContent = msg;
    toastEl.style.opacity = '1';
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.style.opacity = '0'; }, 3000);
  }

  let lastText = null, timer = null;
  function save() {
    const got = grab();
    if (!got || got.text === lastText) return;
    lastText = got.text;
    const list = prune(load());
    const now = new Date().toISOString();
    const last = list[list.length - 1];
    // 同一份清單只更新時間，不重複存
    if (last && last.text === got.text) last.at = now;
    else list.push({ at: now, text: got.text });
    GM_setValue(KEY, JSON.stringify(prune(list)));
    toast(got.count > 0 ? `Ybot 已讀取批核清單（${got.count} 筆）` : 'Ybot 已讀取批核清單（目前沒有案件）');
  }
  // WebITR 是動態載入的頁面，表格常在頁面打開後才出現、切分頁也不會重新整理：
  // 盯著畫面變動，停下來一秒多再讀一次。
  const kick = () => { clearTimeout(timer); timer = setTimeout(save, 1200); };
  new MutationObserver(kick).observe(document.documentElement, { childList: true, subtree: true, characterData: true });
  kick();
})();
