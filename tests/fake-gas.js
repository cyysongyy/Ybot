// 假的 Apps Script 環境：試算表、快取、Telegram、AI 全部是記憶體裡的假貨，不會連任何網路。
// 只放虛構資料。這個 repo 是公開的。
const vm = require('vm');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const BACKEND = path.join(__dirname, '..', 'ybot-backend.gs');

function fakeSheet() {
  const d = [];
  return {
    d,
    getDataRange: () => ({ getValues: () => d.map(r => r.slice()) }),
    getRange: (r, c) => ({
      setValues: v => v.forEach((row, i) => {
        d[r - 1 + i] = d[r - 1 + i] || [];
        row.forEach((x, j) => { d[r - 1 + i][c - 1 + j] = x; });
      }),
      setValue: x => { d[r - 1][c - 1] = x; },
      getValue: () => d[r - 1][c - 1],
      setFontWeight() {}
    }),
    appendRow: r => d.push(r.slice()),
    deleteRow: n => d.splice(n - 1, 1),
    getLastRow: () => d.length,
    getLastColumn: () => (d[0] ? d[0].length : 0),
    setFrozenRows() {}
  };
}

// 回傳一個全新、互不影響的後台。
function loadBackend() {
  const sheets = {};
  const cache = {};
  const props = { TELEGRAM_BOT_TOKEN: 'fake', TELEGRAM_CHAT_ID: '1', AI_KEY_GEMINI: 'fake' };
  const telegram = [];   // 送給 Telegram 的內容
  const aiScript = [];   // 假 AI 依序回覆：字串＝文字回答；[[工具名, 參數], …]＝呼叫工具
  const aiRequests = []; // 送給 AI 的內容
  const mail = [];
  const ok = o => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify(o) });
  const ss = { getSheetByName: n => sheets[n], insertSheet: n => (sheets[n] = fakeSheet()), getId: () => 'x' };

  const g = {
    console,
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => props[k] || null,
        setProperty: (k, v) => { props[k] = v; },
        deleteProperty: k => { delete props[k]; }
      })
    },
    CacheService: {
      getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } })
    },
    Utilities: { getUuid: () => crypto.randomUUID(), formatDate: d => d.toISOString() },
    ContentService: { createTextOutput: t => ({ setMimeType: () => JSON.parse(t) }), MimeType: { JSON: 1 } },
    GmailApp: { search: () => [], getThreadById: () => { throw new Error('fake: no such thread'); } },
    CalendarApp: { getAllCalendars: () => [] },
    MailApp: { sendEmail: (to, subject, text) => mail.push({ to, subject, text }) },
    UrlFetchApp: {
      fetch: (url, o) => {
        if (url.includes('api.telegram.org')) {
          telegram.push({ method: url.split('/').pop(), body: JSON.parse(o.payload) });
          return ok({ ok: true, result: { message_id: telegram.length } });
        }
        if (url.includes('generativelanguage')) {
          aiRequests.push(JSON.parse(o.payload));
          if (!aiScript.length) throw new Error('fake AI: no reply scripted');
          const s = aiScript.shift();
          const parts = typeof s === 'string' ? [{ text: s }] : s.map(([name, args]) => ({ functionCall: { name, args } }));
          return ok({ candidates: [{ content: { parts } }] });
        }
        throw new Error('fake: unexpected network call');
      }
    }
  };
  vm.createContext(g);
  vm.runInContext(fs.readFileSync(BACKEND, 'utf8'), g, { filename: 'ybot-backend.gs' });

  const run = code => vm.runInContext(code, g);
  const token = run('getApiToken()');
  const post = body => g.doPost({ postData: { contents: JSON.stringify(Object.assign({ t: token }, body)) } });
  const rows = name => {
    if (!sheets[name]) return [];
    const [head, ...rest] = sheets[name].d;
    return rest.map(r => Object.fromEntries(head.map((k, i) => [k, r[i]])));
  };
  return {
    g, run, post, rows, sheets, props, cache, telegram, aiScript, aiRequests, mail,
    memories: () => rows(run('QA_SHEET')).filter(r => r.kind === 'profile').map(r => r.a),
    runs: () => rows(run('RUN_SHEET')),
    agent: (script, history = [], question = 'test question') => {
      aiScript.length = 0;
      aiScript.push(...script);
      return post({ action: 'agent', question, history });
    }
  };
}

module.exports = { loadBackend };
