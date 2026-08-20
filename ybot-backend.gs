/**
 * Ybot 個人助理 — Google Apps Script 後台
 * 待辦／筆記／提醒的雲端儲存 + Gmail／日曆彙整 + 每日主動簡報（含🔴🟠🟢優先分級）
 * + 到點提醒信 + 晚間彙整（只在有未處理事項時才寄，平時不打擾）+ 起立提醒 + 本週回顧
 *
 * 使用方式：
 * 1. 開啟一份 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上此程式碼（取代所有內容）
 * 3. 執行一次 setupDailyBrief()（授權後排程「每日簡報」）
 *    再執行一次 setupReminderWatch()（排程「到點提醒」，每 30 分鐘檢查一次）
 *    再執行一次 setupEveningDigest()（排程「晚間彙整」，只在有未處理事項時才寄信）
 *    再執行一次 setupStandupWatch()（排程「起立提醒」，實際頻率依 ybot.html 設定調整）
 *    再執行一次 setupWeeklyReview()（排程「本週回顧」，每週一早上 7 點）
 * 4. 部署 → 新增部署作業 → 網頁應用程式
 *    - 以下列身分執行：我（Me）
 *    - 誰可以存取：所有人（Anyone）
 * 5. 複製部署網址，貼入 ybot.html「⚙️ 設定 → 雲端後台」→ 測試連線
 *
 * 授權提醒：Gmail／日曆彙整需要額外授權 Gmail 與 Calendar 權限，
 * 重新部署後首次執行會跳出授權視窗，請同意（僅你本人帳號讀取，不外傳）。
 *
 * 隱私：所有資料僅存於「你自己的」Google 試算表，由你本人的帳號執行；
 * Gmail／日曆內容只在你自己部署的後台內彙整、回傳給你自己的前端。
 */

// ── 試算表設定 ──────────────────────────────────
const SS = SpreadsheetApp.getActiveSpreadsheet();
const NOTE_SHEET = '待辦與筆記';
const KV_SHEET = '個人設定';
const BRIEF_SHEET = '每日簡報存檔';

const NOTE_COLS = ['id', 'type', 'content', 'dueAt', 'done', 'createdAt', 'notifiedAt'];
// type: note（瑣事筆記）／ todo（待辦）／ reminder（有時間點的提醒，到點會主動寄信）

// 通知信箱（留空則自動用試算表擁有者信箱）
const NOTIFY_EMAIL = '';

// ── AI 設定（選填，供每日簡報生成友善摘要用）────
// 執行一次 setAiConfig('gemini','你的KEY') 或 setAiConfig('openai','sk-...')
function setAiConfig(provider, key) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AI_PROVIDER', provider || 'gemini');
  props.setProperty('AI_KEY', key || '');
  return '✅ AI 設定完成：' + (provider || 'gemini');
}
function getAiConfig() {
  const props = PropertiesService.getScriptProperties();
  return { provider: props.getProperty('AI_PROVIDER') || 'gemini', key: props.getProperty('AI_KEY') || '' };
}
function callAI(prompt) {
  const { provider, key } = getAiConfig();
  if (!key) return '';
  try {
    if (provider === 'openai') {
      const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify({ model: 'gpt-4o-mini', messages: [{ role: 'user', content: prompt }], temperature: 0.6 })
      });
      const d = JSON.parse(res.getContentText());
      return d.choices && d.choices[0] ? d.choices[0].message.content : '';
    } else {
      const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + key;
      const res = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
      });
      const d = JSON.parse(res.getContentText());
      return d.candidates && d.candidates[0] ? d.candidates[0].content.parts[0].text : '';
    }
  } catch (err) { return ''; }
}

// ── 初始化試算表 ────────────────────────────────
function setupSheets() {
  const note = ensureSheet(NOTE_SHEET, NOTE_COLS);
  const kv = ensureSheet(KV_SHEET, ['key', 'value']);
  const brief = ensureSheet(BRIEF_SHEET, ['generatedAt', 'summary']);
  return { note, kv, brief };
}
function ensureSheet(name, cols) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── GET 處理 ────────────────────────────────────
function doGet(e) {
  const action = e?.parameter?.action || 'context';

  if (action === 'ping') {
    return jsonResp({ ok: true, message: 'Ybot 後台連線正常', time: new Date().toISOString() });
  }
  if (action === 'notes') {
    const { note } = setupSheets();
    return jsonResp({ ok: true, notes: sheetToObjects(note, NOTE_COLS) });
  }
  if (action === 'context') {
    return jsonResp({ ok: true, context: buildContext() });
  }
  return jsonResp({ ok: false, error: 'Unknown action: ' + action });
}

// ── POST 處理 ───────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonResp({ ok: false, error: 'Invalid JSON' }); }
  const action = body.action;

  if (action === 'addNote') {
    const { note } = setupSheets();
    const n = body.note || {};
    const id = n.id || Utilities.getUuid();
    appendObj(note, NOTE_COLS, {
      id, type: n.type || 'note', content: n.content || '', dueAt: n.dueAt || '',
      done: n.done ? 'true' : '', createdAt: n.createdAt || new Date().toISOString(), notifiedAt: ''
    });
    return jsonResp({ ok: true, message: '已新增', id });
  }
  if (action === 'updateNote') {
    const { note } = setupSheets();
    updatePartial(note, NOTE_COLS, body.id, body.patch || {});
    return jsonResp({ ok: true, message: '已更新' });
  }
  if (action === 'deleteNote') {
    const { note } = setupSheets();
    deleteRow(note, body.id);
    return jsonResp({ ok: true, message: '已刪除' });
  }
  if (action === 'syncAll') {
    const { note } = setupSheets();
    (body.notes || []).forEach(n => {
      if (!n.id) return;
      upsertRow(note, NOTE_COLS, n.id, {
        id: n.id, type: n.type || 'note', content: n.content || '', dueAt: n.dueAt || '',
        done: n.done ? 'true' : '', createdAt: n.createdAt || new Date().toISOString(), notifiedAt: n.notifiedAt || ''
      });
    });
    return jsonResp({ ok: true, message: `同步完成：${(body.notes || []).length} 筆` });
  }
  if (action === 'saveLinkedBackends') {
    writeKv({ remediationBackendUrl: body.remediationBackendUrl || '', healthBackendUrl: body.healthBackendUrl || '' });
    return jsonResp({ ok: true, message: '已儲存整合設定' });
  }
  if (action === 'saveOtherLinks') {
    writeKv({ otherLinksJson: JSON.stringify(body.links || []) });
    return jsonResp({ ok: true, message: '已儲存其他 App 設定' });
  }
  if (action === 'saveAiConfig') {
    setAiConfig(body.provider, body.key);
    return jsonResp({ ok: true, message: '已儲存 AI 設定' });
  }
  if (action === 'saveWeatherLocation') {
    writeKv({ weatherCity: body.city || '', weatherLat: body.lat || '', weatherLon: body.lon || '' });
    CacheService.getScriptCache().remove('weatherDigest'); // 換地點後清掉舊快取，下次立刻抓新地點
    return jsonResp({ ok: true, message: '已儲存天氣地點' });
  }
  if (action === 'addCalendarEvent') {
    try {
      const title = body.title || '(未命名事件)';
      const start = new Date(body.start);
      if (isNaN(start.getTime())) throw new Error('無效的開始時間');
      const end = (body.end && !isNaN(new Date(body.end).getTime())) ? new Date(body.end) : new Date(start.getTime() + 3600000);
      const ev = CalendarApp.getDefaultCalendar().createEvent(title, start, end);
      ev.addPopupReminder(30);
      return jsonResp({ ok: true, message: '已加入 Google 日曆', eventId: ev.getId() });
    } catch (err) {
      return jsonResp({ ok: false, error: '建立日曆事件失敗：' + err.message });
    }
  }
  if (action === 'saveStandupSettings') {
    writeKv({
      standupEnabled: body.enabled ? 'true' : '',
      standupIntervalMin: String(body.intervalMin || 60),
      standupStartHour: String(body.startHour ?? 8),
      standupEndHour: String(body.endHour ?? 18)
    });
    return jsonResp({ ok: true, message: '已儲存起立提醒設定' });
  }
  return jsonResp({ ok: false, error: 'Unknown action: ' + action });
}

// ── 彙整上下文：Gmail + 日曆 + 新聞 + 天氣 + 起立提醒設定 + 待辦提醒 + 其他系統摘要 ──
function buildContext() {
  return {
    generatedAt: new Date().toISOString(),
    notes: sheetToObjects(setupSheets().note, NOTE_COLS),
    gmail: getGmailDigest(),
    calendar: getCalendarDigest(),
    news: getNewsDigest(),
    weather: getWeatherDigest(),
    standup: getStandupSettings(),
    linked: getLinkedSummaries()
  };
}
function getStandupSettings() {
  const kv = readKv();
  return {
    enabled: kv.standupEnabled === 'true',
    intervalMin: Number(kv.standupIntervalMin || 60),
    startHour: Number(kv.standupStartHour ?? 8),
    endHour: Number(kv.standupEndHour ?? 18)
  };
}

// Gmail 彙整：近 2 天未讀信件，僅取主旨／寄件者／時間／摘要片段（不含全文）。
// 排除 Ybot 自己寄的信（避免自己的每日簡報/提醒被當成「未讀重要信件」形成迴圈）
// 與 Google 安全性快訊；含信用卡/繳費關鍵字的標記為重要並優先顯示；最多列 2 封。
const GMAIL_IMPORTANT_KEYWORDS = ['信用卡', '繳費', '帳單', '付款', '逾期', 'invoice', 'payment', 'credit card'];
function getGmailDigest() {
  try {
    const threads = GmailApp.search('is:unread newer_than:2d -from:me -from:no-reply@accounts.google.com', 0, 20);
    const out = threads.map(t => {
      const msgs = t.getMessages();
      const last = msgs[msgs.length - 1];
      const subject = t.getFirstMessageSubject() || '';
      const body = (last.getPlainBody() || '').replace(/\s+/g, ' ');
      const hay = (subject + ' ' + body).toLowerCase();
      const important = GMAIL_IMPORTANT_KEYWORDS.some(k => hay.includes(k.toLowerCase()));
      return {
        subject, from: last.getFrom(), date: last.getDate().toISOString(),
        snippet: body.slice(0, 120), important
      };
    });
    out.sort((a, b) => (b.important - a.important) || (new Date(b.date) - new Date(a.date)));
    return out.slice(0, 2);
  } catch (err) { return []; }
}

// 日曆彙整：未來 7 天的行程
// 讀取「所有」看得到的日曆（含訂閱/分享加入清單的），不只主要日曆，
// 避免副曆的行程被漏掉；同一事件若出現在多個日曆中則去重。
function getCalendarDigest() {
  try {
    const now = new Date();
    const until = new Date(now.getTime() + 7 * 86400000);
    const calendars = CalendarApp.getAllCalendars();
    const seen = new Set();
    const events = [];
    calendars.forEach(cal => {
      try {
        cal.getEvents(now, until).forEach(ev => {
          const key = ev.getId() + '_' + ev.getStartTime().getTime();
          if (seen.has(key)) return;
          seen.add(key);
          events.push({
            title: ev.getTitle(),
            start: ev.getStartTime().toISOString(),
            end: ev.getEndTime().toISOString(),
            allDay: ev.isAllDayEvent(),
            calendar: cal.getName()
          });
        });
      } catch (err) { /* 單一日曆讀取失敗就跳過，不影響其他日曆 */ }
    });
    events.sort((a, b) => new Date(a.start) - new Date(b.start));
    return events.slice(0, 50);
  } catch (err) { return []; }
}

// 手動清除新聞快取用：改了新聞來源設定後，在 Apps Script 編輯器執行一次這個函式，
// 不用等 30 分鐘快取自然過期就能立刻看到新設定的結果。
function clearNewsCache() {
  CacheService.getScriptCache().remove('newsDigest');
  return '✅ 已清除新聞快取，下次抓取會是最新設定的結果';
}
// 新聞彙整：國內／國際／教育各 2 則，來源 Google 新聞 RSS（免金鑰）。
// 快取 15 分鐘（原本 30 分鐘），提高新聞輪替更新的頻率。
function getNewsDigest() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('newsDigest');
  if (cached) { try { return JSON.parse(cached); } catch (err) { /* 快取壞掉就重抓 */ } }
  // 國際新聞改用英文（美國）語系，來源會是 CNN／Reuters／AP／BBC 等主流英語媒體，
  // 跟國內/教育新聞（中文語系）分開設定。
  const out = {
    domestic: fetchNewsRss('https://news.google.com/rss?hl=zh-TW&gl=TW&ceid=TW:zh-Tw', 2),
    world: fetchNewsRss('https://news.google.com/rss/headlines/section/topic/WORLD?hl=en-US&gl=US&ceid=US:en', 2),
    education: fetchNewsRss('https://news.google.com/rss/search?q=%E6%95%99%E8%82%B2&hl=zh-TW&gl=TW&ceid=TW:zh-Tw', 2)
  };
  try { cache.put('newsDigest', JSON.stringify(out), 900); } catch (err) { /* 超過 CacheService 容量就不快取，下次仍會重抓 */ }
  return out;
}
function fetchNewsRss(url, limit) {
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    const doc = XmlService.parse(res.getContentText());
    const items = doc.getRootElement().getChild('channel').getChildren('item');
    return items.slice(0, limit).map(it => ({
      title: it.getChildText('title') || '',
      link: it.getChildText('link') || ''
    }));
  } catch (err) { return []; }
}

// 天氣彙整：來源 Open-Meteo（免金鑰）。地點預設「台北」，可在設定 → 雲端後台改地點名稱
// （後台會用地名查經緯度），或前端用 GPS 直接送經緯度過來（跳過查詢，最準確）。
// 快取 30 分鐘；地名轉經緯度另外快取 24 小時，很少變動不用常查。
function getWeatherDigest() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('weatherDigest');
  if (cached) { try { return JSON.parse(cached); } catch (err) { /* 快取壞掉就重抓 */ } }

  const kv = readKv();
  let coords;
  if (kv.weatherLat && kv.weatherLon) {
    coords = { lat: kv.weatherLat, lon: kv.weatherLon, name: '目前位置' };
  } else {
    const cityLabel = kv.weatherCity || '台北';
    // 縣市名稱（如「雲林」）在地理編碼 API 常會誤配到縣內同名小地方而非縣治所在，
    // 有手動校正過座標的縣市一律優先用 DEFAULT_WEATHER_COORDS，不查即時地理編碼；
    // 中文、英文拼音（如「雲林」／「Yunlin」）都要能對到同一組固定座標。
    const enLabel = CITY_EN[cityLabel] || cityLabel;
    coords = DEFAULT_WEATHER_COORDS[cityLabel] || DEFAULT_WEATHER_COORDS[enLabel] || geocodeCity(cityLabel) || null;
  }
  if (!coords) return null;

  let out = null;
  try {
    const url = 'https://api.open-meteo.com/v1/forecast?latitude=' + coords.lat + '&longitude=' + coords.lon +
      '&current=temperature_2m,weather_code,relative_humidity_2m,precipitation_probability' +
      '&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max' +
      '&timezone=Asia%2FTaipei&forecast_days=1';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const d = JSON.parse(res.getContentText());
      const cur = d.current || {};
      const daily = d.daily || {};
      const info = weatherCodeInfo(cur.weather_code);
      out = {
        city: coords.name,
        icon: info[0], desc: info[1],
        temp: cur.temperature_2m,
        humidity: cur.relative_humidity_2m,
        tMax: (daily.temperature_2m_max || [])[0] ?? null,
        tMin: (daily.temperature_2m_min || [])[0] ?? null,
        rainChance: (daily.precipitation_probability_max || [])[0] ?? cur.precipitation_probability ?? null
      };
    }
  } catch (err) { out = null; }

  if (out) { try { cache.put('weatherDigest', JSON.stringify(out), 1800); } catch (err) { /* 超過容量就不快取 */ } }
  return out;
}
// Open-Meteo 地理編碼對中文城市名常常查不到（索引主要是英文/拼音），
// 常見台灣城市先轉英文名再查，大幅提高命中率；查不到再用下面的固定座標當保底。
const CITY_EN = {
  '台北': 'Taipei', '臺北': 'Taipei', '新北': 'New Taipei', '桃園': 'Taoyuan',
  '台中': 'Taichung', '臺中': 'Taichung', '台南': 'Tainan', '臺南': 'Tainan',
  '高雄': 'Kaohsiung', '基隆': 'Keelung', '新竹': 'Hsinchu', '嘉義': 'Chiayi',
  '宜蘭': 'Yilan', '花蓮': 'Hualien', '台東': 'Taitung', '臺東': 'Taitung',
  '南投': 'Nantou', '雲林': 'Yunlin', '彰化': 'Changhua', '苗栗': 'Miaoli',
  '屏東': 'Pingtung', '澎湖': 'Penghu', '金門': 'Kinmen'
};
const DEFAULT_WEATHER_COORDS = {
  '台北': { lat: 25.033, lon: 121.5654, name: '台北' },
  '臺北': { lat: 25.033, lon: 121.5654, name: '台北' },
  // 雲林是縣名不是單一地點，地理編碼查「Yunlin」容易誤配到縣內同名小地方（如「Yongling」），
  // 改用縣治斗六市的座標，結果較準確、穩定。
  '雲林': { lat: 23.7092, lon: 120.5414, name: '雲林' },
  'Yunlin': { lat: 23.7092, lon: 120.5414, name: '雲林' }
};
function geocodeCity(name) {
  const cache = CacheService.getScriptCache();
  const key = 'geo_' + name;
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch (err) { /* 快取壞掉就重查 */ } }
  const query = CITY_EN[name] || name;
  try {
    const url = 'https://geocoding-api.open-meteo.com/v1/search?count=1&language=zh&name=' + encodeURIComponent(query);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const d = JSON.parse(res.getContentText());
    const r = d.results && d.results[0];
    if (!r) return null;
    const out = { lat: r.latitude, lon: r.longitude, name: r.name };
    try { cache.put(key, JSON.stringify(out), 86400); } catch (err) { /* 超過容量就不快取，下次仍會重查 */ }
    return out;
  } catch (err) { return null; }
}
function weatherCodeInfo(code) {
  const map = {
    0: ['☀️', '晴朗'], 1: ['🌤️', '大致晴朗'], 2: ['⛅', '局部多雲'], 3: ['☁️', '陰天'],
    45: ['🌫️', '有霧'], 48: ['🌫️', '霧淞'],
    51: ['🌦️', '毛毛雨'], 53: ['🌦️', '毛毛雨'], 55: ['🌦️', '毛毛雨'],
    56: ['🌧️', '凍雨'], 57: ['🌧️', '凍雨'],
    61: ['🌧️', '小雨'], 63: ['🌧️', '中雨'], 65: ['🌧️', '大雨'],
    66: ['🌧️', '凍雨'], 67: ['🌧️', '凍雨'],
    71: ['🌨️', '小雪'], 73: ['🌨️', '中雪'], 75: ['🌨️', '大雪'], 77: ['🌨️', '雪粒'],
    80: ['🌦️', '短暫陣雨'], 81: ['🌦️', '短暫陣雨'], 82: ['⛈️', '強陣雨'],
    85: ['🌨️', '短暫陣雪'], 86: ['🌨️', '短暫陣雪'],
    95: ['⛈️', '雷雨'], 96: ['⛈️', '雷雨挾冰雹'], 99: ['⛈️', '強雷雨挾冰雹']
  };
  return map[code] || ['🌡️', '天氣'];
}

// 串接考卷批改／健康預測兩套系統的後台（若已在設定中填入網址）
function getLinkedSummaries() {
  const kv = readKv();
  const out = {};
  if (kv.remediationBackendUrl) {
    try {
      const res = UrlFetchApp.fetch(kv.remediationBackendUrl + '?action=all', { muteHttpExceptions: true });
      const d = JSON.parse(res.getContentText());
      if (d.ok) {
        const subs = d.submissions || [];
        const recent = subs.slice(-10);
        const avg = recent.length ? Math.round(recent.reduce((s, r) => s + (Number(r.percentage) || 0), 0) / recent.length) : null;
        const misconceptions = {};
        recent.forEach(r => (r.misconceptions || []).forEach(m => { misconceptions[m] = (misconceptions[m] || 0) + 1; }));
        out.remediation = {
          totalSubmissions: subs.length,
          recentAvgPercentage: avg,
          topMisconceptions: Object.entries(misconceptions).sort((a, b) => b[1] - a[1]).slice(0, 5).map(x => x[0])
        };
      }
    } catch (err) { /* 忽略連線失敗 */ }
  }
  if (kv.healthBackendUrl) {
    try {
      const res = UrlFetchApp.fetch(kv.healthBackendUrl + '?action=latest', { muteHttpExceptions: true });
      const d = JSON.parse(res.getContentText());
      if (d.ok && d.snapshot) {
        out.health = {
          date: d.snapshot.date, score: d.snapshot.score, bio: d.snapshot.bio,
          cvdLevel: d.snapshot.cvdLevel, dmLevel: d.snapshot.dmLevel
        };
      }
    } catch (err) { /* 忽略連線失敗 */ }
  }
  // 使用者在「我的其他 App」自行填入的任意串聯 App（若有填後台 API 網址則嘗試best-effort讀取）
  try {
    const others = JSON.parse(kv.otherLinksJson || '[]');
    const otherOut = {};
    others.forEach(link => {
      if (!link || !link.apiUrl) return;
      try {
        const res = UrlFetchApp.fetch(link.apiUrl + '?action=all', { muteHttpExceptions: true });
        let text = res.getContentText() || '';
        if (text.length > 1500) text = text.slice(0, 1500) + '…';
        otherOut[link.name || link.apiUrl] = text;
      } catch (err) { /* 忽略單一 App 連線失敗 */ }
    });
    if (Object.keys(otherOut).length) out.other = otherOut;
  } catch (err) { /* 忽略解析失敗 */ }
  return out;
}

// ── 自動化：每日主動簡報 ────────────────────────
function setupDailyBrief() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyBrief') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBrief').timeBased().everyDays(1).atHour(7).create();
  return '✅ 已排程每日早上 7 點自動產生 Ybot 簡報';
}

function dailyBrief() {
  const ctx = buildContext();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const { items: priorities, narrative } = buildPriorities(ctx);
  const groups = { red: [], orange: [], green: [] };
  priorities.forEach(p => groups[p.level].push(p));

  let lines = [`【Ybot 每日簡報】${new Date().toLocaleDateString('zh-TW')}`, ''];

  if (ctx.weather) {
    const w = ctx.weather;
    lines.push(`${w.icon} ${w.city}天氣：${w.desc}，現在 ${w.temp}°C（今日 ${w.tMin}~${w.tMax}°C，降雨機率 ${w.rainChance ?? '—'}%）`);
    lines.push('');
  }
  if (ctx.calendar.length) {
    lines.push('📅 未來行程：');
    ctx.calendar.forEach(ev => {
      const t = ev.allDay ? '全天' : new Date(ev.start).toLocaleString('zh-TW', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' });
      lines.push(`　${t}　${ev.title}`);
    });
    lines.push('');
  }
  if (priorities.length) {
    lines.push(`🎯 今天真正需要處理（共 ${priorities.length} 項，含今明行程）：`);
    if (groups.red.length) {
      lines.push('🔴 今天必須：');
      groups.red.forEach(p => lines.push(fmtItemLine(p)));
    }
    if (groups.orange.length) {
      lines.push('🟠 建議今天：');
      groups.orange.forEach(p => lines.push(fmtItemLine(p)));
    }
    if (groups.green.length) {
      lines.push(`🟢 可以延後（共 ${groups.green.length} 項，僅列前 5）：`);
      groups.green.slice(0, 5).forEach(p => lines.push(fmtItemLine(p)));
    }
    lines.push('');
  }
  if (ctx.gmail.length) {
    lines.push(`📬 未讀重要信件（${ctx.gmail.length} 封）：`);
    ctx.gmail.forEach(m => lines.push(`　${m.important ? '💳' : '・'} ${m.subject}（${m.from}）`));
    lines.push('');
  }
  const news = ctx.news || {};
  if ((news.domestic || []).length || (news.world || []).length || (news.education || []).length) {
    lines.push('📰 今日新聞：');
    if ((news.domestic || []).length) { lines.push('　國內：'); news.domestic.forEach(n => lines.push(`　　・${n.title}`)); }
    if ((news.world || []).length) { lines.push('　國際：'); news.world.forEach(n => lines.push(`　　・${n.title}`)); }
    if ((news.education || []).length) { lines.push('　教育：'); news.education.forEach(n => lines.push(`　　・${n.title}`)); }
    lines.push('');
  }
  if (ctx.linked.remediation) {
    const r = ctx.linked.remediation;
    lines.push(`🎯 考卷批改：累計 ${r.totalSubmissions} 筆，近期平均 ${r.recentAvgPercentage ?? '—'}%，常見迷思：${(r.topMisconceptions || []).join('、') || '無'}`);
  }
  if (ctx.linked.health) {
    const h = ctx.linked.health;
    lines.push(`🫀 健康：最新健康分 ${h.score}，生理年齡 ${h.bio}（${h.date}）`);
  }
  if (ctx.linked.other && Object.keys(ctx.linked.other).length) {
    lines.push('🧩 其他串聯 App：');
    Object.entries(ctx.linked.other).forEach(([name, text]) => {
      lines.push(`　・${name}：${String(text).slice(0, 200)}`);
    });
  }
  lines.push('');

  if (narrative) lines.push('💬 Ybot 想跟你說：\n' + narrative + '\n');

  lines.push('（本簡報由 Ybot 後台每日自動產生。）');
  const summary = lines.join('\n');

  appendObj(setupSheets().brief, ['generatedAt', 'summary'], { generatedAt: new Date().toISOString(), summary });
  try { MailApp.sendEmail(email, '🤖 Ybot 每日簡報', summary); } catch (err) { }
}

// 把「待辦／提醒」與「今明的行程」合併分成 🔴今天必須／🟠建議今天／🟢可延後，
// 規則先判斷（有到期日／行程時間就照時間分級），AI 只負責「升級」沒有到期日、但內容看起來緊急的
// 待辦事項（不會動到行程），不確定就不動——避免 AI 亂猜出不存在的急迫性。
// 較遠期（3天以上）的行程不併入這裡，避免跟下面「📅 未來行程」整週清單重複列。
function buildPriorities(ctx) {
  const now = new Date();
  const actionable = ctx.notes.filter(n => n.type !== 'note' && n.done !== 'true');
  const todoItems = actionable.map(n => ({ id: n.id, type: n.type, content: n.content, dueAt: n.dueAt, level: ruleLevel(n, now) }));
  const ai = aiPrioritize(todoItems, ctx);
  (ai.upgrades || []).forEach(u => {
    const it = todoItems.find(x => x.id === u.id);
    if (it && (u.level === 'red' || u.level === 'orange') && rank(u.level) > rank(it.level)) it.level = u.level;
  });

  const eventItems = (ctx.calendar || [])
    .map((ev, i) => ({ id: 'evt' + i, type: 'event', content: ev.title, dueAt: ev.start, allDay: ev.allDay, level: eventLevel(ev, now) }))
    .filter(e => e.level !== 'green');

  const items = todoItems.concat(eventItems);
  items.sort((a, b) => rank(b.level) - rank(a.level));
  return { items, narrative: ai.narrative || '' };
}
function ruleLevel(n, now) {
  if (!n.dueAt) return 'green';
  const diffDays = (new Date(n.dueAt) - now) / 86400000;
  if (diffDays <= 1) return 'red';
  if (diffDays <= 3) return 'orange';
  return 'green';
}
function eventLevel(ev, now) {
  const diffDays = (new Date(ev.start) - now) / 86400000;
  if (diffDays < 0) return 'green'; // 已開始/已過，不再算今天要處理
  if (diffDays <= 1) return 'red'; // 今天
  if (diffDays <= 2) return 'orange'; // 明天，先提醒你準備
  return 'green';
}
function rank(level) { return { green: 0, orange: 1, red: 2 }[level] || 0; }
function fmtDue(dueAt) {
  try { return new Date(dueAt).toLocaleString('zh-TW', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }); }
  catch (err) { return dueAt; }
}
function fmtItemLine(p) {
  if (p.type === 'event') {
    const t = p.allDay ? '全天' : fmtDue(p.dueAt);
    return `　📅 ${t}　${p.content}`;
  }
  return `　・${p.content}${p.dueAt ? '（' + fmtDue(p.dueAt) + '）' : ''}`;
}

// AI 輔助分級 + 今日提醒文字（選填，需先設定 AI Key）。單一次呼叫回傳 JSON，
// 只允許「升級」清單內既有項目，不允許新增清單外的項目，避免幻覺。
function aiPrioritize(items, ctx) {
  const { key } = getAiConfig();
  if (!key || !items.length) return {};
  const prompt = `你是 Young 的個人行政幕僚 Ybot，個性溫暖直接。以下是他目前的待辦／提醒清單（JSON），每項已有初步等級（red=今天必須, orange=建議今天, green=可延後）：\n` +
    JSON.stringify(items.map(i => ({ id: i.id, content: i.content, dueAt: i.dueAt, level: i.level }))) + '\n\n' +
    `背景資訊（僅供你判斷是否要升級等級，不要新增清單外的項目、不要臆測不存在的細節）：\n` +
    `未讀信件主旨：${JSON.stringify(ctx.gmail.slice(0, 8).map(m => m.subject))}\n` +
    `未來行程：${JSON.stringify(ctx.calendar.map(e => e.title))}\n` +
    `其他系統摘要：${JSON.stringify(ctx.linked)}\n\n` +
    `請完成兩件事，只回傳一個 JSON 物件，不要有任何其他文字或說明：\n` +
    `1. upgrades：陣列，只列出你「有把握」該升級等級的項目（例如沒有到期日但內容明顯緊急），格式 [{"id":"...","level":"red"}]，沒有就給空陣列，不確定的不要動。\n` +
    `2. narrative：100 字內的繁體中文「今日提醒」，語氣像朋友提醒，聚焦在今天最該優先做的 1-2 件事並給一句鼓勵，不要條列複誦清單內容。\n\n` +
    `輸出格式：{"upgrades":[...],"narrative":"..."}`;
  const raw = callAI(prompt);
  if (!raw) return {};
  try {
    const match = raw.match(/\{[\s\S]*\}/);
    const obj = JSON.parse(match ? match[0] : raw);
    return {
      upgrades: Array.isArray(obj.upgrades) ? obj.upgrades : [],
      narrative: typeof obj.narrative === 'string' ? obj.narrative : ''
    };
  } catch (err) { return {}; }
}

// ── 自動化：到點提醒（每 30 分鐘檢查一次）──────
function setupReminderWatch() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'reminderWatch') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reminderWatch').timeBased().everyMinutes(30).create();
  return '✅ 已排程每 30 分鐘檢查一次到點提醒';
}

function reminderWatch() {
  const { note } = setupSheets();
  const all = sheetToObjects(note, NOTE_COLS);
  const now = new Date();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const due = all.filter(n => n.type === 'reminder' && n.done !== 'true' && !n.notifiedAt && n.dueAt && new Date(n.dueAt) <= now);
  due.forEach(n => {
    try {
      MailApp.sendEmail(email, '⏰ Ybot 提醒：' + n.content, `到了你設定的提醒時間：\n\n${n.content}\n\n（設定時間：${n.dueAt}）`);
      updatePartial(note, NOTE_COLS, n.id, { notifiedAt: new Date().toISOString() });
    } catch (err) { /* 忽略單筆寄送失敗 */ }
  });
}

// ── 自動化：本週回顧（每週一早上 7 點）──
// 注意：這是「本週活動摘要＋目前狀態快照」，不是「跟上週比較的趨勢」——
// 因為目前沒有存歷史快照，做不出真正的週對週比較，誠實只呈現看得到的資料。
function setupWeeklyReview() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'weeklyReview') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyReview').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  return '✅ 已排程每週一早上 7 點產生本週回顧';
}

function weeklyReview() {
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;
  const ctx = buildContext();
  const weekAgo = new Date(Date.now() - 7 * 86400000);
  const allNotes = sheetToObjects(setupSheets().note, NOTE_COLS);
  const addedThisWeek = allNotes.filter(n => n.createdAt && new Date(n.createdAt) >= weekAgo);
  const decisionNotesThisWeek = allNotes.filter(n => n.type === 'note' && n.createdAt && new Date(n.createdAt) >= weekAgo);

  let lines = [`【Ybot 本週回顧】${new Date().toLocaleDateString('zh-TW')}`, ''];
  lines.push(`📝 本週新增 ${addedThisWeek.length} 筆（待辦／提醒／瑣事）`);
  if (decisionNotesThisWeek.length) {
    lines.push('📌 本週瑣事筆記：');
    decisionNotesThisWeek.slice(0, 15).forEach(n => lines.push(`　・${n.content}`));
  }
  lines.push('');
  if (ctx.linked.remediation) {
    const r = ctx.linked.remediation;
    lines.push(`🎯 考卷批改目前狀態：累計 ${r.totalSubmissions} 筆，近期平均 ${r.recentAvgPercentage ?? '—'}%，常見迷思：${(r.topMisconceptions || []).join('、') || '無'}`);
  }
  if (ctx.linked.health) {
    const h = ctx.linked.health;
    lines.push(`🫀 健康目前狀態：健康分 ${h.score}，生理年齡 ${h.bio}（${h.date}）`);
  }
  if (ctx.linked.other && Object.keys(ctx.linked.other).length) {
    lines.push('🧩 其他串聯 App：');
    Object.entries(ctx.linked.other).forEach(([name, text]) => lines.push(`　・${name}：${String(text).slice(0, 200)}`));
  }
  lines.push('');

  const narrative = aiWeeklyNarrative(addedThisWeek, decisionNotesThisWeek, ctx);
  if (narrative) lines.push('💬 Ybot 本週小結：\n' + narrative + '\n');

  lines.push('（本回顧由 Ybot 後台每週一自動產生，內容為本週活動摘要與目前狀態，非跨週趨勢比較。）');
  try { MailApp.sendEmail(email, '🗓️ Ybot 本週回顧', lines.join('\n')); } catch (err) { }
}

function aiWeeklyNarrative(addedThisWeek, decisionNotesThisWeek, ctx) {
  const { key } = getAiConfig();
  if (!key) return '';
  const prompt = `你是 Young 的個人行政幕僚 Ybot，個性溫暖直接。以下是本週資料：\n` +
    `本週新增待辦/提醒/瑣事共 ${addedThisWeek.length} 筆\n` +
    `本週瑣事筆記：${JSON.stringify(decisionNotesThisWeek.map(n => n.content))}\n` +
    `其他系統目前狀態：${JSON.stringify(ctx.linked)}\n\n` +
    `請用繁體中文寫一段 120 字內的「本週小結」，語氣溫暖直接，重點放在這週值得注意的事跟一句給下週的建議，不要條列複誦上面內容，不要臆測不存在的細節。`;
  return callAI(prompt);
}

// ── 自動化：晚間彙整（只在有事沒處理時才寄，不打擾）──
// 早上簡報看過的事，如果晚上還沒完成，累積提醒一次；沒有待處理事項就完全不寄信。
function setupEveningDigest() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'eveningDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('eveningDigest').timeBased().everyDays(1).atHour(20).create();
  return '✅ 已排程每天晚上 8 點檢查（有未處理的🔴/🟠事項才會寄信，沒有就不打擾）';
}

function eveningDigest() {
  const ctx = buildContext();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const { items } = buildPriorities(ctx);
  const pending = items.filter(p => p.level === 'red' || p.level === 'orange');
  if (!pending.length) return; // 沒有需要留意的事，不寄信

  let lines = [`【Ybot 晚間提醒】${new Date().toLocaleDateString('zh-TW')}`, '', `還有 ${pending.length} 項需要留意：`];
  pending.slice(0, 8).forEach(p => {
    const flag = p.level === 'red' ? '🔴' : '🟠';
    const label = p.type === 'event' ? `${flag} 📅 ${p.allDay ? '全天' : fmtDue(p.dueAt)}　${p.content}` : `${flag} ${p.content}${p.dueAt ? '（' + fmtDue(p.dueAt) + '）' : ''}`;
    lines.push('　' + label);
  });
  lines.push('', '（僅在有待留意事項時才會寄這封信。）');
  try { MailApp.sendEmail(email, '🌙 Ybot 晚間提醒', lines.join('\n')); } catch (err) { }
}

// ── 自動化：起立提醒（每 15 分鐘檢查一次，依設定的間隔與時段決定是否真的寄信）──
// 間隔／時段可在 ybot.html「設定」調整，不用重新排程觸發器；這裡只是頻繁檢查、
// 依 KV 存的設定與上次寄送時間判斷「這次要不要真的寄」。
function setupStandupWatch() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'standupWatch') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('standupWatch').timeBased().everyMinutes(15).create();
  return '✅ 已排程每 15 分鐘檢查一次起立提醒（實際頻率依「設定」中的間隔與時段）';
}

function standupWatch() {
  const s = getStandupSettings();
  if (!s.enabled) return;
  const now = new Date();
  const hour = now.getHours();
  if (s.startHour <= s.endHour) {
    if (hour < s.startHour || hour >= s.endHour) return; // 不在活躍時段
  } else {
    if (hour < s.startHour && hour >= s.endHour) return; // 跨午夜的時段（例如 22~6）
  }
  const kv = readKv();
  const last = kv.standupLastSent ? new Date(kv.standupLastSent) : null;
  if (last && (now - last) / 60000 < s.intervalMin) return; // 還沒到下次提醒時間

  // 建立一個「現在」的極短行事曆事件並設定立即彈出提醒——手機上的日曆 App（Google 日曆／
  // 有訂閱進 Apple 日曆的話也算）會用系統原生通知鈴聲跳出提醒，不需要打開 Ybot 也會響。
  try {
    const end = new Date(now.getTime() + 5 * 60000);
    const ev = CalendarApp.getDefaultCalendar().createEvent('🧍 起立走動一下', now, end,
      { description: `每 ${s.intervalMin} 分鐘提醒一次，時段 ${s.startHour}:00~${s.endHour}:00，可在 Ybot「設定」調整或關閉。` });
    ev.addPopupReminder(0);
  } catch (err) { /* 忽略建立失敗 */ }

  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (email) {
    try {
      MailApp.sendEmail(email, '🧍 起立提醒',
        `久坐了，站起來動一動、喝口水吧！\n\n（每 ${s.intervalMin} 分鐘提醒一次，時段 ${s.startHour}:00~${s.endHour}:00，可在 Ybot「設定」調整或關閉。）`);
    } catch (err) { /* 忽略單次寄送失敗 */ }
  }
  writeKv({ standupLastSent: now.toISOString() });
}

function getOwnerEmail() {
  try { return Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; }
}

// ── KV 設定 ─────────────────────────────────────
function readKv() {
  const { kv } = setupSheets();
  const data = kv.getDataRange().getValues();
  const obj = {};
  for (let i = 1; i < data.length; i++) { if (data[i][0]) obj[data[i][0]] = data[i][1]; }
  return obj;
}
function writeKv(map) {
  const { kv } = setupSheets();
  Object.keys(map).forEach(k => upsertRow(kv, ['key', 'value'], k, { key: k, value: map[k] }));
}

// ── 通用工具 ────────────────────────────────────
function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheetToObjects(sheet, cols) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => {
    const o = {}; cols.forEach((c, i) => o[c] = row[i] ?? ''); return o;
  }).filter(o => o.id);
}
function appendObj(sheet, cols, obj) {
  sheet.appendRow(cols.map(c => obj[c] !== undefined ? obj[c] : ''));
}
function upsertRow(sheet, cols, id, rowData) {
  const data = sheet.getDataRange().getValues();
  let target = -1;
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { target = i + 1; break; } }
  const row = cols.map(c => rowData[c] !== undefined ? rowData[c] : '');
  if (target > 0) sheet.getRange(target, 1, 1, cols.length).setValues([row]);
  else sheet.appendRow(row);
}
function updatePartial(sheet, cols, id, patch) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const rowData = {};
      cols.forEach((c, j) => rowData[c] = data[i][j]);
      Object.assign(rowData, patch);
      const row = cols.map(c => rowData[c] !== undefined ? rowData[c] : '');
      sheet.getRange(i + 1, 1, 1, cols.length).setValues([row]);
      return;
    }
  }
}
function deleteRow(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) { if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); break; } }
}
