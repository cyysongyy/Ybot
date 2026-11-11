/**
 * Ybot 個人助理 — Google Apps Script 後台
 * 待辦／筆記／提醒的雲端儲存 + Gmail／日曆彙整 + 每日主動簡報（含🔴🟠🟢優先分級）
 * + 到點提醒信 + 晚間彙整（只在有未處理事項時才寄，平時不打擾）+ 起立提醒 + 本週回顧
 *
 * 使用方式：
 * 1. 開啟一份 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上此程式碼（取代所有內容）
 * 3. 執行一次 generateApiToken()，把回傳的那串通關密語貼進
 *    Ybot「設定 → 雲端後台 → 通關密語」（沒設的話，任何知道網址的人都讀得到你的資料）
 * 3-1. 執行一次 setupDailyBrief()（授權後排程「每日簡報」）
 *    再執行一次 setupReminderWatch()（排程「到點提醒」，每 30 分鐘檢查一次）
 *    再執行一次 setupEveningDigest()（排程「晚間彙整」，只在有未處理事項時才寄信）
 *    再執行一次 setupStandupWatch()（排程「起立提醒」，實際頻率依 ybot.html 設定調整）
 *    再執行一次 setupWeeklyReview()（排程「本週回顧」，每週一早上 7 點）
 * 3-2.（選填）Telegram 通知：手機鎖著也收得到新的【差勤批核】筆記。
 *    a. 在 Telegram 找 @BotFather 建立機器人（或沿用股票提醒那一個），拿到 token
 *    b. 專案設定（齒輪）→ 指令碼屬性 → 新增屬性 TELEGRAM_BOT_TOKEN，值貼 token
 *       （token 只放這裡，不要寫進程式碼）
 *    c. 在 Telegram 對你的機器人隨便傳一句話，再執行一次 findTelegramChatId()
 *    d. 執行一次 setupTelegramWatch()（每 10 分鐘檢查一次，並先傳一則測試訊息）
 *    會推：新的【差勤批核】筆記、行程開始前約 20～30 分鐘提醒
 * 4. 部署 → 新增部署作業 → 網頁應用程式
 *    - 以下列身分執行：我（Me）
 *    - 誰可以存取：所有人（Anyone）
 * 5. 複製部署網址，貼入 ybot.html「⚙️ 設定 → 雲端後台」→ 測試連線
 *
 * AI 摘要（選填，供每日簡報/晚間提醒/本週回顧生成友善文字用）：
 * 支援 Gemini／OpenAI／NVIDIA，可只設定一組，也可都設定（自動依優先順序備援）。
 * 每個 provider 各自執行一次：
 *   setAiConfig('gemini','你的KEY') / setAiConfig('openai','sk-...') / setAiConfig('nvidia','nvapi-...')
 * 選填：想指定優先用哪個，執行一次 setAiPriority('nvidia')（預設 'auto'：依序 gemini→openai→nvidia）
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
// 問答資料庫：前端每問一次就存一筆，這裡是雲端備份。
// localStorage 會被「清除瀏覽器資料」清掉，也可能被 iOS 在空間吃緊時回收，
// 累積幾個月的問答只放本機太脆弱。
const QA_SHEET = '問答資料庫';

// ── 時間一律台北 ────────────────────────────────
// Apps Script 的 new Date().toLocaleString() 與 getHours() 都跟著「專案時區」走，
// 而專案時區在部署時可能不是台北（新建專案常常是 America/New_York）。
// 起立提醒的時段判斷、每日簡報的日期都靠這些，抓錯就整個歪掉，所以全部寫死台北。
const TZ = 'Asia/Taipei';
function tpeFormat(d, pattern) {
  return Utilities.formatDate(d || new Date(), TZ, pattern);
}
// 台北的日期字串，例如 2026/9/5
function tpeDateStr(d) { return tpeFormat(d, 'yyyy/M/d'); }
// 台北的月日時分，例如 9/5 14:30
function tpeDateTimeStr(d) { return tpeFormat(d, 'M/d HH:mm'); }
// 台北的「幾點」，用來判斷活躍時段
function tpeHour(d) { return Number(tpeFormat(d, 'H')); }
// 定時觸發器（每日簡報 6 點、晚間提醒 8 點…）是照「專案時區」發動的，
// 那個時區只能在 Apps Script 的專案設定裡改，程式改不了。所以這裡負責提醒：
// 排程函式的回傳訊息會直接告訴你目前時區對不對。
function scriptTz() {
  try { return Session.getScriptTimeZone(); } catch (e) { return ''; }
}
function tzWarn() {
  const tz = scriptTz();
  if (tz === TZ) return '';
  return '\n\n⚠️ 注意：這個 Apps Script 專案的時區目前是「' + (tz || '未知') + '」，不是台北。'
    + '定時排程會照專案時區發動，所以簡報／提醒會在錯的時間寄出。'
    + '請到「專案設定（齒輪）→ 時區」改成 (GMT+08:00) Taipei，再重新執行一次這個排程函式。';
}
// 想單獨確認時區的話執行這一個
function checkTimezone() {
  const tz = scriptTz();
  return (tz === TZ)
    ? '✅ 專案時區正確：' + tz
    : '⚠️ 專案時區是「' + (tz || '未知') + '」，請改成 (GMT+08:00) Taipei。顯示用的時間已經一律是台北，但定時排程會照專案時區發動。';
}

// owner：交辦給誰。空白＝自己要做的。
// 加在最後面，舊資料的那七欄位置不動，既有資料不用搬。
const NOTE_COLS = ['id', 'type', 'content', 'dueAt', 'done', 'createdAt', 'notifiedAt', 'owner'];
// kind: chat（對話問答）／ relax（鬆弛練習）
const QA_COLS = ['id', 'q', 'a', 'kind', 'at'];
// type: note（瑣事筆記）／ todo（待辦）／ reminder（有時間點的提醒，到點會主動寄信）

// 通知信箱（留空則自動用試算表擁有者信箱）
const NOTIFY_EMAIL = '';

// ── AI 設定（選填，供每日簡報生成友善摘要用）────
// 支援 Gemini／OpenAI／NVIDIA 三選一或都設定，設定多組時會依優先順序自動備援。
// 每個 provider 各自執行一次 setAiConfig('gemini','你的KEY') / setAiConfig('openai','sk-...') / setAiConfig('nvidia','nvapi-...')
// 選填：想指定優先用哪個 provider，執行一次 setAiPriority('nvidia')（不設定則為 'auto'：依序 gemini→openai→nvidia，跳過沒設 Key 的）
const AI_PROVIDERS = ['gemini', 'openai', 'nvidia'];
function setAiConfig(provider, key) {
  if (AI_PROVIDERS.indexOf(provider) === -1) throw new Error('provider 必須是 gemini／openai／nvidia 其中之一');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AI_KEY_' + provider.toUpperCase(), key || '');
  return '✅ 已儲存 ' + provider + ' 的 AI Key';
}
function setAiPriority(provider) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AI_PRIORITY', (provider && AI_PROVIDERS.indexOf(provider) !== -1) ? provider : 'auto');
  return '✅ AI 優先順序設定為：' + (props.getProperty('AI_PRIORITY'));
}
function getAiConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    priority: props.getProperty('AI_PRIORITY') || 'auto',
    gemini: props.getProperty('AI_KEY_GEMINI') || '',
    openai: props.getProperty('AI_KEY_OPENAI') || '',
    nvidia: props.getProperty('AI_KEY_NVIDIA') || ''
  };
}
function hasAnyAiKey(cfg) { return !!(cfg.gemini || cfg.openai || cfg.nvidia); }
// 依優先順序排出要嘗試的 provider 清單，只列出有設 Key 的
function aiOrder(cfg) {
  const available = AI_PROVIDERS.filter(id => cfg[id]);
  if (cfg.priority === 'auto' || !cfg.priority) return available;
  return [cfg.priority].concat(available.filter(id => id !== cfg.priority)).filter(id => cfg[id]);
}

// NVIDIA 模型目錄常會輪替／棄用，若又失效請至 build.nvidia.com/models 換一個目前可用的模型 id
// （跟前端 index.html 的 NV_DEFAULT_MODEL 保持同一個，方便一起更新）
const NV_ENDPOINT = 'https://integrate.api.nvidia.com/v1/chat/completions';
const NV_DEFAULT_MODEL = 'nvidia/llama-3.1-nemotron-70b-instruct';

// 依序嘗試設定好的 provider，第一個成功回傳文字的就用它；全部失敗則回傳空字串
// （呼叫端本來就會把空字串當「沒有 AI 摘要」處理，不影響其他功能正常運作）
function callAI(prompt) {
  const cfg = getAiConfig();
  const order = aiOrder(cfg);
  for (let i = 0; i < order.length; i++) {
    try {
      const text = callAiProvider(order[i], cfg[order[i]], prompt);
      if (text) return text;
    } catch (err) { /* 這個 provider 失敗就換下一個，不中斷 */ }
  }
  return '';
}
function callAiProvider(provider, key, prompt) {
  if (provider === 'gemini') {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + key;
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.6, maxOutputTokens: 1024 } })
    });
    const d = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) throw new Error((d.error && d.error.message) || ('Gemini ' + res.getResponseCode()));
    return (d.candidates && d.candidates[0] && d.candidates[0].content.parts[0].text) || '';
  }
  // OpenAI／NVIDIA 都是 OpenAI 相容的 chat/completions 介面
  const isNv = provider === 'nvidia';
  const endpoint = isNv ? NV_ENDPOINT : 'https://api.openai.com/v1/chat/completions';
  const model = isNv ? NV_DEFAULT_MODEL : 'gpt-4o-mini';
  const res = UrlFetchApp.fetch(endpoint, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({ model: model, messages: [{ role: 'user', content: prompt }], temperature: 0.6, max_tokens: 1024 })
  });
  const d = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error((d.error && d.error.message) || d.detail || (provider + ' ' + res.getResponseCode()));
  return (d.choices && d.choices[0] && d.choices[0].message.content) || '';
}

// ── 通關密語（保護後台不被「知道網址的人」隨意讀寫）──────────
// 這個網頁應用程式必須部署成「所有人皆可存取」，瀏覽器才連得上，
// 也就是說：任何知道網址的人都打得到，後台不會因為你換手機就認得你。
// 網址是一長串亂碼、猜不出來，但它會出現在瀏覽器歷史紀錄、截圖、
// 轉貼的訊息裡——一旦流出去，對方就能讀到 Gmail 摘要、行事曆與全部
// 筆記問答，而且你不會收到通知，也查不出誰讀過。
// 設了密語之後，光有網址沒有用；密語萬一外洩，重新產生一組就好，
// 網址不用換、Ybot 那邊也只要改一個欄位。
//
// 啟用：在編輯器執行一次 generateApiToken()，把回傳的那串貼進
// Ybot「設定 → 雲端後台 → 通關密語」。
const TOKEN_PROP = 'API_TOKEN';
function generateApiToken() {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty(TOKEN_PROP, token);
  return '✅ 新的通關密語（整串複製，貼進 Ybot 設定 → 雲端後台 → 通關密語）：\n\n'
    + token + '\n\n舊的密語即刻失效，網址不用換。';
}
function getApiToken() {
  return PropertiesService.getScriptProperties().getProperty(TOKEN_PROP) || '';
}
// 想確認目前有沒有保護，執行這一個（不會把密語印出來）
function checkApiToken() {
  return getApiToken()
    ? '✅ 已啟用通關密語，只有帶對密語的請求讀得到資料。'
    : '⚠️ 尚未設定通關密語，任何知道網址的人都能讀寫這個後台。執行 generateApiToken() 產生一組。';
}
function clearApiToken() {
  PropertiesService.getScriptProperties().deleteProperty(TOKEN_PROP);
  return '⚠️ 已關閉通關密語檢查，現在任何知道網址的人都能讀寫這個後台。';
}
// 還沒設密語時一律放行。不然貼上新程式、還來不及執行 generateApiToken()，
// App 就整個連不上，反而會被誤判成「後台又失效了」。放行時回應會帶
// tokenRequired:false，前端據此顯示「尚未保護」的警告。
function authOk(supplied) {
  const token = getApiToken();
  if (!token) return true;
  return String(supplied || '') === token;
}
function authFail() {
  return jsonResp({ ok: false, error: '通關密語不正確或尚未填寫', authRequired: true });
}

// ── 初始化試算表 ────────────────────────────────
function setupSheets() {
  const note = ensureSheet(NOTE_SHEET, NOTE_COLS);
  const kv = ensureSheet(KV_SHEET, ['key', 'value']);
  const brief = ensureSheet(BRIEF_SHEET, ['generatedAt', 'summary']);
  const qa = ensureSheet(QA_SHEET, QA_COLS);
  return { note, kv, brief, qa };
}
function ensureSheet(name, cols) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // 之後加欄位時（例如 owner），既有的工作表不會自己長出標題列。
  // 資料本身照 index 讀寫沒問題，但表頭缺一格，人去看試算表會一頭霧水。
  try {
    if (sh.getLastColumn() < cols.length) {
      sh.getRange(1, 1, 1, cols.length).setValues([cols]);
      sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    }
  } catch (err) { /* 補標題失敗不影響讀寫，略過 */ }
  return sh;
}

// ── GET 處理 ────────────────────────────────────
function doGet(e) {
  const action = e?.parameter?.action || 'context';
  if (!authOk(e?.parameter?.t)) return authFail();

  if (action === 'ping') {
    // tokenRequired 讓前端知道這個後台到底有沒有上鎖，沒上鎖就在設定頁提醒。
    return jsonResp({ ok: true, message: 'Ybot 後台連線正常', time: new Date().toISOString(),
      tokenRequired: !!getApiToken() });
  }
  if (action === 'notes') {
    const { note } = setupSheets();
    return jsonResp({ ok: true, notes: sheetToObjects(note, NOTE_COLS) });
  }
  if (action === 'context') {
    return jsonResp({ ok: true, context: buildContext() });
  }
  if (action === 'qa') {
    const { qa } = setupSheets();
    return jsonResp({ ok: true, qa: sheetToObjects(qa, QA_COLS) });
  }
  return jsonResp({ ok: false, error: 'Unknown action: ' + action });
}

// ── POST 處理 ───────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonResp({ ok: false, error: 'Invalid JSON' }); }
  if (!authOk(body.t)) return authFail();
  const action = body.action;

  if (action === 'addNote') {
    const { note } = setupSheets();
    const n = body.note || {};
    const id = n.id || Utilities.getUuid();
    appendObj(note, NOTE_COLS, {
      id, type: n.type || 'note', content: n.content || '', dueAt: n.dueAt || '',
      done: n.done ? 'true' : '', createdAt: n.createdAt || new Date().toISOString(), notifiedAt: '',
      owner: n.owner || ''
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
        done: n.done ? 'true' : '', createdAt: n.createdAt || new Date().toISOString(),
        notifiedAt: n.notifiedAt || '', owner: n.owner || ''
      });
    });
    return jsonResp({ ok: true, message: '同步完成：' + ((body.notes || []).length) + ' 筆' });
  }
  if (action === 'addQa') {
    const { qa } = setupSheets();
    const r = body.qa || {};
    if (!r.id) return jsonResp({ ok: false, error: '缺少 id' });
    // 用 upsert 而不是 append：離線補送或重試時同一筆才不會變成兩列。
    upsertRow(qa, QA_COLS, r.id, {
      id: r.id, q: r.q || '', a: r.a || '', kind: r.kind || 'chat',
      at: r.at || new Date().toISOString()
    });
    return jsonResp({ ok: true, message: '已存入問答庫', id: r.id });
  }
  if (action === 'syncQa') {
    const { qa } = setupSheets();
    const list = body.qa || [];
    list.forEach(r => {
      if (!r.id) return;
      upsertRow(qa, QA_COLS, r.id, {
        id: r.id, q: r.q || '', a: r.a || '', kind: r.kind || 'chat',
        at: r.at || new Date().toISOString()
      });
    });
    return jsonResp({ ok: true, message: '問答庫同步完成：' + list.length + ' 筆' });
  }
  if (action === 'deleteQa') {
    const { qa } = setupSheets();
    deleteRow(qa, body.id);
    return jsonResp({ ok: true, message: '已刪除' });
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
  if (action === 'notify') {
    // 立刻寄一封通知信給自己（外部腳本用，例如差勤批核結果），不用等 reminderWatch 的 30 分鐘排程。
    const email = NOTIFY_EMAIL || getOwnerEmail();
    if (!email) return jsonResp({ ok: false, error: '找不到通知信箱' });
    try {
      const subject = String(body.subject || '通知').slice(0, 120);
      MailApp.sendEmail(email, '🔔 Ybot：' + subject, String(body.message || subject));
      return jsonResp({ ok: true, message: '已寄出通知' });
    } catch (err) {
      return jsonResp({ ok: false, error: '寄信失敗：' + err.message });
    }
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
const GMAIL_IMPORTANT_KEYWORDS = ['聯邦', '信用卡', '繳費', '帳單', '付款', '逾期', 'invoice', 'payment', 'credit card'];
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
        snippet: body.slice(0, 120), important, threadId: t.getId()
      };
    });
    out.sort((a, b) => (b.important - a.important) || (new Date(b.date) - new Date(a.date)));
    return out.slice(0, 2);
  } catch (err) { return []; }
}

// 日曆彙整：未來 14 天的行程
// 原本是 7 天。問題是：禮拜一問「下禮拜的行程」，下禮拜五已經超過 7 天，
// 答案會缺一截，而且前端看不出來是缺的還是真的沒有。14 天涵蓋得住
// 「這禮拜＋下禮拜」的任何問法。
// 讀取「所有」看得到的日曆（含訂閱/分享加入清單的），不只主要日曆，
// 避免副曆的行程被漏掉；同一事件若出現在多個日曆中則去重。
function getCalendarDigest() {
  try {
    const now = new Date();
    const until = new Date(now.getTime() + 14 * 86400000);
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
    return events.slice(0, 80);   // 兩週的量比一週多，上限跟著放寬
  } catch (err) { return []; }
}

// 抓行程撞期：只比對有明確時間的行程，全天事件排除在外——一整天的活動
// 跟某場會議「撞」在同一天不是真正的衝突。events 先照開始時間排序，一旦
// 後面那筆的開始時間晚於前面那筆的結束時間，後面更晚的也不會跟前面重疊，
// 可以直接跳出內層迴圈，不用兩兩全比對。
function detectCalendarConflicts(events) {
  const timed = (events || []).filter(function (ev) { return !ev.allDay; })
    .map(function (ev) { return { title: ev.title, start: new Date(ev.start), end: new Date(ev.end) }; })
    .filter(function (ev) { return !isNaN(ev.start) && !isNaN(ev.end); })
    .sort(function (a, b) { return a.start - b.start; });
  const conflicts = [];
  for (let i = 0; i < timed.length; i++) {
    for (let j = i + 1; j < timed.length; j++) {
      if (timed[j].start >= timed[i].end) break;
      conflicts.push([timed[i], timed[j]]);
    }
  }
  return conflicts;
}
function fmtConflictEvent(ev) {
  return tpeDateTimeStr(ev.start) + '~' + tpeFormat(ev.end, 'HH:mm') + ' ' + ev.title;
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
  ScriptApp.newTrigger('dailyBrief').timeBased().everyDays(1).atHour(6).nearMinute(0).create();
  return '✅ 已排程每日早上 6 點自動產生 Ybot 簡報與 Telegram 晨報（台北時間，實際會在 6:00～6:15 之間送出）' + tzWarn();
}

function dailyBriefImpl() {
  const ctx = buildContext();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const { items: priorities, narrative } = buildPriorities(ctx);
  const groups = { red: [], orange: [], green: [] };
  priorities.forEach(p => groups[p.level].push(p));

  let lines = ['【Ybot 每日簡報】' + tpeDateStr() + '（台北時間）', ''];

  if (ctx.weather) {
    const w = ctx.weather;
    lines.push(w.icon + ' ' + w.city + '天氣：' + w.desc + '，現在 ' + w.temp + '°C（今日 ' + w.tMin + '~' + w.tMax + '°C，降雨機率 ' + (w.rainChance ?? '—') + '%）');
    lines.push('');
  }
  // 撞期用完整 14 天的行事曆資料，不像下面「未來一週行程」只列 7 天——
  // 衝突通常是提前排的，越早發現越有時間喬開，不用等到剩不到一週才知道。
  const conflicts = detectCalendarConflicts(ctx.calendar);
  if (conflicts.length) {
    lines.push('⚠️ 行程撞期（' + conflicts.length + ' 組，請確認）：');
    conflicts.forEach(function (pair) {
      lines.push('　' + fmtConflictEvent(pair[0]) + '　撞　' + fmtConflictEvent(pair[1]));
    });
    lines.push('');
  }
  // context 給到 14 天是為了讓對話問得到「下禮拜」，但早上這封信不需要跟著
  // 變成兩倍長——信是拿來掃一眼的，只列一週。
  const briefWindow = new Date(Date.now() + 7 * 86400000);
  const briefEvents = (ctx.calendar || []).filter(function (ev) {
    const d = new Date(ev.start);
    return !isNaN(d) && d <= briefWindow;
  });
  if (briefEvents.length) {
    lines.push('📅 未來一週行程：');
    briefEvents.forEach(ev => {
      const t = ev.allDay ? '全天' : tpeDateTimeStr(new Date(ev.start));
      lines.push('　' + t + '　' + ev.title);
    });
    lines.push('');
  }
  if (priorities.length) {
    lines.push('🎯 今天真正需要處理（共 ' + priorities.length + ' 項，含今明行程）：');
    if (groups.red.length) {
      lines.push('🔴 今天必須：');
      groups.red.forEach(p => lines.push(fmtItemLine(p)));
    }
    if (groups.orange.length) {
      lines.push('🟠 建議今天：');
      groups.orange.forEach(p => lines.push(fmtItemLine(p)));
    }
    if (groups.green.length) {
      lines.push('🟢 可以延後（共 ' + groups.green.length + ' 項，僅列前 5）：');
      groups.green.slice(0, 5).forEach(p => lines.push(fmtItemLine(p)));
    }
    lines.push('');
  }
  if ((ctx.gmail || []).length) {
    lines.push('📬 未讀重要信件（' + ctx.gmail.length + ' 封）：');
    ctx.gmail.forEach(m => lines.push('　' + (m.important ? '💳' : '・') + ' ' + m.subject + '（' + m.from + '）'));
    lines.push('');
  }
  const news = ctx.news || {};
  if ((news.domestic || []).length || (news.world || []).length || (news.education || []).length) {
    lines.push('📰 今日新聞：');
    if ((news.domestic || []).length) { lines.push('　國內：'); news.domestic.forEach(n => lines.push('　　・' + n.title)); }
    if ((news.world || []).length) { lines.push('　國際：'); news.world.forEach(n => lines.push('　　・' + n.title)); }
    if ((news.education || []).length) { lines.push('　教育：'); news.education.forEach(n => lines.push('　　・' + n.title)); }
    lines.push('');
  }
  if (ctx.linked.remediation) {
    const r = ctx.linked.remediation;
    lines.push('🎯 考卷批改：累計 ' + r.totalSubmissions + ' 筆，近期平均 ' + (r.recentAvgPercentage ?? '—') + '%，常見迷思：' + ((r.topMisconceptions || []).join('、') || '無'));
  }
  const stockLine = getStockSummaryLine();
  if (stockLine) lines.push(stockLine);
  if (ctx.linked.health) {
    const h = ctx.linked.health;
    lines.push('🫀 健康：最新健康分 ' + h.score + '，生理年齡 ' + h.bio + '（' + h.date + '）');
  }
  if (ctx.linked.other && Object.keys(ctx.linked.other).length) {
    lines.push('🧩 其他串聯 App：');
    Object.entries(ctx.linked.other).forEach(([name, text]) => {
      lines.push('　・' + name + '：' + (String(text).slice(0, 200)));
    });
  }
  lines.push('');

  if (narrative) lines.push('💬 Ybot 想跟你說：\n' + narrative + '\n');

  lines.push('（本簡報由 Ybot 後台每日自動產生。）');
  const summary = lines.join('\n');

  appendObj(setupSheets().brief, ['generatedAt', 'summary'], { generatedAt: new Date().toISOString(), summary });
  try { MailApp.sendEmail(email, '🤖 Ybot 每日簡報', summary); } catch (err) { }
  // 同一份簡報的精簡版也傳 Telegram（有設定 token 與 chat id 才傳；失敗不影響上面的 Email）
  try {
    const tg = tgConfig();
    if (tg.token && tg.chatId) {
      telegramSend(buildTelegramBrief(ctx, groups, narrative, conflicts, briefEvents, stockLine));
    }
  } catch (err) { reportError('Telegram 晨報', err, 24); /* Email 已經寄出 */ }
}

// Telegram 版晨報：手機上掃一眼用，所以只放天氣、撞期、今明行程、🔴🟠待辦、Ybot 的話。
// 不放新聞與信件寄件者（太長，也不必把信件資訊送到 Telegram）。
function buildTelegramBrief(ctx, groups, narrative, conflicts, briefEvents, stockLine) {
  const L = ['🌅 Ybot 晨報　' + tpeDateStr()];
  if (ctx.weather) {
    const w = ctx.weather;
    L.push(w.icon + ' ' + w.city + ' ' + w.desc + ' ' + w.tMin + '~' + w.tMax + '°C，降雨 ' + (w.rainChance ?? '—') + '%');
  }
  if (conflicts && conflicts.length) {
    L.push('', '⚠️ 行程撞期 ' + conflicts.length + ' 組：');
    conflicts.forEach(pair => L.push('　' + fmtConflictEvent(pair[0]) + ' 撞 ' + fmtConflictEvent(pair[1])));
  }
  // 只列今天與明天的行程；整週清單留在 Email
  const todayStr = tpeFormat(new Date(), 'yyyy-MM-dd');
  const tomorrowStr = tpeFormat(new Date(Date.now() + 86400000), 'yyyy-MM-dd');
  const near = (briefEvents || []).filter(ev => {
    const d = tpeFormat(new Date(ev.start), 'yyyy-MM-dd');
    return d === todayStr || d === tomorrowStr;
  });
  if (near.length) {
    L.push('', '📅 今明行程：');
    near.slice(0, 8).forEach(ev => {
      const t = ev.allDay ? '全天' : tpeDateTimeStr(new Date(ev.start));
      L.push('　' + t + ' ' + ev.title);
    });
  }
  if (groups.red.length || groups.orange.length) {
    L.push('', '🎯 今天要處理：');
    groups.red.forEach(p => L.push('🔴 ' + String(p.content || p.title || '').slice(0, 60)));
    groups.orange.slice(0, 5).forEach(p => L.push('🟠 ' + String(p.content || p.title || '').slice(0, 60)));
  } else {
    L.push('', '🎯 今天沒有必須處理的待辦');
  }
  if (stockLine) L.push('', stockLine);
  if (narrative) L.push('', '💬 ' + String(narrative).slice(0, 300));
  return L.join('\n');
}

// 把「待辦／提醒」與「今明的行程」合併分成 🔴今天必須／🟠建議今天／🟢可延後，
// 規則先判斷（有到期日／行程時間就照時間分級），AI 只負責「升級」沒有到期日、但內容看起來緊急的
// 待辦事項（不會動到行程），不確定就不動——避免 AI 亂猜出不存在的急迫性。
// 較遠期（3天以上）的行程不併入這裡，避免跟下面「📅 未來行程」整週清單重複列。
function buildPriorities(ctx) {
  const now = new Date();
  const actionable = ctx.notes.filter(n => n.type !== 'note' && n.type !== 'idea' && n.done !== 'true');
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
  try { return tpeDateTimeStr(new Date(dueAt)); }
  catch (err) { return dueAt; }
}
function fmtItemLine(p) {
  if (p.type === 'event') {
    const t = p.allDay ? '全天' : fmtDue(p.dueAt);
    return '　📅 ' + t + '　' + p.content;
  }
  return '　・' + p.content + (p.dueAt ? '（' + fmtDue(p.dueAt) + '）' : '');
}

// AI 輔助分級 + 今日提醒文字（選填，需先設定 AI Key）。單一次呼叫回傳 JSON，
// 只允許「升級」清單內既有項目，不允許新增清單外的項目，避免幻覺。
function aiPrioritize(items, ctx) {
  const cfg = getAiConfig();
  if (!hasAnyAiKey(cfg) || !items.length) return {};
  const prompt = '你是 Young 的個人行政幕僚 Ybot，個性溫暖直接。以下是他目前的待辦／提醒清單（JSON），每項已有初步等級（red=今天必須, orange=建議今天, green=可延後）：\n' +
    JSON.stringify(items.map(i => ({ id: i.id, content: i.content, dueAt: i.dueAt, level: i.level }))) + '\n\n' +
    '背景資訊（僅供你判斷是否要升級等級，不要新增清單外的項目、不要臆測不存在的細節）：\n' +
    '未讀信件主旨：' + (JSON.stringify(ctx.gmail.slice(0, 8).map(m => m.subject))) + '\n' +
    '未來行程：' + (JSON.stringify(ctx.calendar.map(e => e.title))) + '\n' +
    '其他系統摘要：' + (JSON.stringify(ctx.linked)) + '\n\n' +
    '請完成兩件事，只回傳一個 JSON 物件，不要有任何其他文字或說明：\n' +
    '1. upgrades：陣列，只列出你「有把握」該升級等級的項目（例如沒有到期日但內容明顯緊急），格式 [{"id":"...","level":"red"}]，沒有就給空陣列，不確定的不要動。\n' +
    '2. narrative：100 字內的繁體中文「今日提醒」，語氣像朋友提醒，聚焦在今天最該優先做的 1-2 件事並給一句鼓勵，不要條列複誦清單內容。\n\n' +
    '輸出格式：{"upgrades":[...],"narrative":"..."}';
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
  ScriptApp.newTrigger('reminderWatch').timeBased().everyMinutes(10).create();
  return '✅ 已排程每 30 分鐘檢查一次到點提醒' + tzWarn();
}

function reminderWatchImpl() {
  const { note } = setupSheets();
  const all = sheetToObjects(note, NOTE_COLS);
  const now = new Date();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  const tgReady = !!(tgConfig().token && tgConfig().chatId);
  if (!email && !tgReady) return;

  const due = all.filter(n => n.type === 'reminder' && n.done !== 'true' && !n.notifiedAt && n.dueAt && new Date(n.dueAt) <= now);
  due.forEach(n => {
    try {
      // 先傳 Telegram（附「完成／延後 1 小時」按鈕）；沒設定或傳失敗才退回寄 Email
      let sent = tgReady && sendReminderTelegram(n);
      if (!sent && email) {
        MailApp.sendEmail(email, '⏰ Ybot 提醒：' + n.content, '到了你設定的提醒時間：\n\n' + n.content + '\n\n（設定時間：' + n.dueAt + '）');
        sent = true;
      }
      if (sent) updatePartial(note, NOTE_COLS, n.id, { notifiedAt: new Date().toISOString() });
    } catch (err) { /* 忽略單筆寄送失敗 */ }
  });

  if (email) autoEscalateStaleReminders(note, all, now, email);
}

// 自主動作：提醒已經到點寄過信，卻超過 AUTO_ESCALATE_HOURS 小時還沒被標記完成，
// 代表大概率已經錯過那個時間點、或這件事其實還沒做完。提醒一旦寄過信
// （notifiedAt 已設）reminderWatch 就不會再理它，等於悄悄消失在提醒清單裡；
// 與其這樣，不如自動轉成「待辦」，讓它繼續出現在待辦清單／每日簡報。
// 這是真的會改資料的自主動作，所以每次都寄信讓使用者知道 Ybot 做了什麼，
// 且完全可逆（把類型改回「提醒」即可）。
const AUTO_ESCALATE_HOURS = 24;
function autoEscalateStaleReminders(note, all, now, email) {
  const stale = all.filter(n => n.type === 'reminder' && n.done !== 'true' && n.notifiedAt &&
    (now - new Date(n.notifiedAt)) >= AUTO_ESCALATE_HOURS * 3600000);
  if (!stale.length) return;
  stale.forEach(n => updatePartial(note, NOTE_COLS, n.id, { type: 'todo' }));
  try {
    const lines = stale.map(n => '・' + n.content + '（原提醒時間：' + n.dueAt + '）');
    MailApp.sendEmail(email, '🤖 Ybot 自動處理：' + stale.length + ' 筆逾期提醒已轉為待辦',
      '這些提醒已經過了 ' + AUTO_ESCALATE_HOURS + ' 小時還沒標記完成，Ybot 幫你自動轉成「待辦」，之後會繼續出現在待辦清單與每日簡報裡（不會像提醒一樣就此消失）：\n\n' +
      lines.join('\n') + '\n\n如果不想改，打開 App 把類型改回「提醒」即可。');
  } catch (err) { /* 忽略寄送失敗，資料已經改完 */ }
}

// ── 自動化：本週回顧（每週一早上 7 點）──
// 注意：這是「本週活動摘要＋目前狀態快照」，不是「跟上週比較的趨勢」——
// 因為目前沒有存歷史快照，做不出真正的週對週比較，誠實只呈現看得到的資料。
function setupWeeklyReview() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'weeklyReview') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyReview').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  return '✅ 已排程每週一早上 7 點產生本週回顧（台北時間）' + tzWarn();
}

function weeklyReviewImpl() {
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;
  const ctx = buildContext();
  const weekAgo = new Date(Date.now() - 7 * 86400000);
  const allNotes = sheetToObjects(setupSheets().note, NOTE_COLS);
  const addedThisWeek = allNotes.filter(n => n.createdAt && new Date(n.createdAt) >= weekAgo);
  const decisionNotesThisWeek = allNotes.filter(n => n.type === 'note' && n.createdAt && new Date(n.createdAt) >= weekAgo);

  let lines = ['【Ybot 本週回顧】' + tpeDateStr() + '（台北時間）', ''];
  lines.push('📝 本週新增 ' + addedThisWeek.length + ' 筆（待辦／提醒／瑣事）');
  if (decisionNotesThisWeek.length) {
    lines.push('📌 本週瑣事筆記：');
    decisionNotesThisWeek.slice(0, 15).forEach(n => lines.push('　・' + n.content));
  }
  lines.push('');
  if (ctx.linked.remediation) {
    const r = ctx.linked.remediation;
    lines.push('🎯 考卷批改目前狀態：累計 ' + r.totalSubmissions + ' 筆，近期平均 ' + (r.recentAvgPercentage ?? '—') + '%，常見迷思：' + ((r.topMisconceptions || []).join('、') || '無'));
  }
  if (ctx.linked.health) {
    const h = ctx.linked.health;
    lines.push('🫀 健康目前狀態：健康分 ' + h.score + '，生理年齡 ' + h.bio + '（' + h.date + '）');
  }
  if (ctx.linked.other && Object.keys(ctx.linked.other).length) {
    lines.push('🧩 其他串聯 App：');
    Object.entries(ctx.linked.other).forEach(([name, text]) => lines.push('　・' + name + '：' + (String(text).slice(0, 200))));
  }
  lines.push('');

  const narrative = aiWeeklyNarrative(addedThisWeek, decisionNotesThisWeek, ctx);
  if (narrative) lines.push('💬 Ybot 本週小結：\n' + narrative + '\n');

  lines.push('（本回顧由 Ybot 後台每週一自動產生，內容為本週活動摘要與目前狀態，非跨週趨勢比較。）');
  try { MailApp.sendEmail(email, '🗓️ Ybot 本週回顧', lines.join('\n')); } catch (err) { }
}

function aiWeeklyNarrative(addedThisWeek, decisionNotesThisWeek, ctx) {
  if (!hasAnyAiKey(getAiConfig())) return '';
  const prompt = '你是 Young 的個人行政幕僚 Ybot，個性溫暖直接。以下是本週資料：\n' +
    '本週新增待辦/提醒/瑣事共 ' + addedThisWeek.length + ' 筆\n' +
    '本週瑣事筆記：' + (JSON.stringify(decisionNotesThisWeek.map(n => n.content))) + '\n' +
    '其他系統目前狀態：' + (JSON.stringify(ctx.linked)) + '\n\n' +
    '請用繁體中文寫一段 120 字內的「本週小結」，語氣溫暖直接，重點放在這週值得注意的事跟一句給下週的建議，不要條列複誦上面內容，不要臆測不存在的細節。';
  return callAI(prompt);
}

// ── 自動化：晚間彙整（只在有事沒處理時才寄，不打擾）──
// 早上簡報看過的事，如果晚上還沒完成，累積提醒一次；沒有待處理事項就完全不寄信。
function setupEveningDigest() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'eveningDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('eveningDigest').timeBased().everyDays(1).atHour(20).create();
  return '✅ 已排程每天晚上 8 點檢查（台北時間；有未處理的🔴/🟠事項才會寄信，沒有就不打擾）' + tzWarn();
}

function eveningDigestImpl() {
  const ctx = buildContext();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const { items } = buildPriorities(ctx);
  const pending = items.filter(p => p.level === 'red' || p.level === 'orange');
  if (!pending.length) return; // 沒有需要留意的事，不寄信

  let lines = ['【Ybot 晚間提醒】' + tpeDateStr() + '（台北時間）', '', '還有 ' + pending.length + ' 項需要留意：'];
  pending.slice(0, 8).forEach(p => {
    const flag = p.level === 'red' ? '🔴' : '🟠';
    const label = p.type === 'event' ? flag + ' 📅 ' + (p.allDay ? '全天' : fmtDue(p.dueAt)) + '　' + p.content : flag + ' ' + p.content + (p.dueAt ? '（' + fmtDue(p.dueAt) + '）' : '');
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
  return '✅ 已排程每 15 分鐘檢查一次起立提醒（時段判斷用台北時間）' + tzWarn();
}

function standupWatchImpl() {
  const s = getStandupSettings();
  if (!s.enabled) return;
  const now = new Date();
  const hour = tpeHour(now);   // 用台北的小時，不要用專案時區
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
      { description: '每 ' + s.intervalMin + ' 分鐘提醒一次，時段 ' + s.startHour + ':00~' + s.endHour + ':00，可在 Ybot「設定」調整或關閉。' });
    ev.addPopupReminder(0);
  } catch (err) { /* 忽略建立失敗 */ }

  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (email) {
    try {
      MailApp.sendEmail(email, '🧍 起立提醒',
        '久坐了，站起來動一動、喝口水吧！\n\n（每 ' + s.intervalMin + ' 分鐘提醒一次，時段 ' + s.startHour + ':00~' + s.endHour + ':00，可在 Ybot「設定」調整或關閉。）');
    } catch (err) { /* 忽略單次寄送失敗 */ }
  }
  writeKv({ standupLastSent: now.toISOString() });
}

// ── Telegram 通知 ───────────────────────────────
// iPhone 上 Ybot 網頁要開著才會跳通知；改由後台每 10 分鐘檢查一次、直接傳 Telegram，
// 手機鎖著也會響。推：新的【差勤批核】筆記、行程前提醒。
// token 和 chat id 只存在「指令碼屬性」，不寫進程式碼。
const TG_TOKEN_PROP = 'TELEGRAM_BOT_TOKEN';
const TG_CHAT_PROP = 'TELEGRAM_CHAT_ID';
const TG_SENT_PROP = 'TELEGRAM_SENT_IDS';
const TG_APPROVAL_TAG = '【差勤批核】';
const TG_APPROVAL_FRESH_MS = 36 * 3600000;   // 跟前端首頁一樣：一天半內的才算新的
function tgProps() { return PropertiesService.getScriptProperties(); }
function tgConfig() {
  const p = tgProps();
  return { token: p.getProperty(TG_TOKEN_PROP) || '', chatId: p.getProperty(TG_CHAT_PROP) || '' };
}
// 也可以不用指令碼屬性畫面，改在編輯器執行 setTelegramConfig('token', 'chatId')，
// 執行完記得把括號裡的 token 刪掉，不要存檔留在程式碼裡。
function setTelegramConfig(token, chatId) {
  const p = tgProps();
  if (token) p.setProperty(TG_TOKEN_PROP, String(token).trim());
  if (chatId) p.setProperty(TG_CHAT_PROP, String(chatId).trim());
  return '✅ 已儲存 Telegram 設定';
}
function tgApi(method, payload) {
  const { token } = tgConfig();
  if (!token) throw new Error('還沒設定 ' + TG_TOKEN_PROP + '（專案設定 → 指令碼屬性）');
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload || {}), muteHttpExceptions: true
  });
  let d = {};
  try { d = JSON.parse(r.getContentText()); } catch (e) { /* 下面統一當失敗 */ }
  if (!d.ok) throw new Error('Telegram 回應錯誤：' + (d.description || ('HTTP ' + r.getResponseCode())));
  return d.result;
}
// 先在 Telegram 對機器人傳一句話，再執行這個：自動找到你的 chat id 存起來。
function findTelegramChatId() {
  const updates = tgApi('getUpdates', {});
  const chats = {};
  (updates || []).forEach(u => {
    const m = u.message || u.edited_message || u.channel_post;
    if (m && m.chat) chats[m.chat.id] = m.chat.first_name || m.chat.title || m.chat.username || '';
  });
  const ids = Object.keys(chats);
  // Apps Script 的執行記錄不會顯示函式的回傳值，所以結果一律 console.log 出來，
  // 失敗的情況直接丟錯誤（記錄裡是紅字），不然使用者只會看到「執行完畢」不知道成功沒。
  if (!ids.length) throw new Error('找不到對話：請先在 Telegram 打開你的機器人（不是 BotFather），按「開始」再傳一句 hi，然後再執行一次。');
  if (ids.length > 1) throw new Error('找到好幾個對話，請挑你自己的那個，執行 setTelegramConfig(\'\', \'chat id\')：' +
    ids.map(id => id + '（' + chats[id] + '）').join('、'));
  tgProps().setProperty(TG_CHAT_PROP, ids[0]);
  const msg = '✅ 已儲存 chat id（' + chats[ids[0]] + '）。接著執行 setupTelegramWatch()。';
  console.log(msg);
  return msg;
}
function telegramSend(text) {
  const { chatId } = tgConfig();
  if (!chatId) throw new Error('還沒設定 chat id：先對機器人傳一句話，再執行 findTelegramChatId()');
  tgApi('sendMessage', { chat_id: chatId, text: String(text).slice(0, 4000), disable_web_page_preview: true });
}
function testTelegram() {
  telegramSend('✅ Ybot Telegram 通知設定成功。之後有新的差勤批核、行程前提醒會傳到這裡。');
  console.log('✅ 已傳出測試訊息，請看 Telegram');
  return '✅ 已傳出測試訊息，請看 Telegram';
}
function setupTelegramWatch() {
  // 還沒有 chat id 就順手找一次，少一個步驟
  if (!tgConfig().chatId) findTelegramChatId();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'telegramWatch') ScriptApp.deleteTrigger(t);
  });
  testTelegram();
  ScriptApp.newTrigger('telegramWatch').timeBased().everyMinutes(10).create();
  const msg = '✅ 已排程每 10 分鐘檢查一次，並傳出一則測試訊息' + tzWarn();
  console.log(msg);
  return msg;
}
function telegramWatchImpl() {
  const cfg = tgConfig();
  if (!cfg.token || !cfg.chatId) return;
  const p = tgProps();
  let sent = [];
  try { sent = JSON.parse(p.getProperty(TG_SENT_PROP) || '[]'); } catch (e) { sent = []; }
  const now = Date.now();
  const outbox = [];
  sheetToObjects(setupSheets().note, NOTE_COLS).forEach(n => {
    const c = String(n.content || '').trim();
    if (!c.startsWith(TG_APPROVAL_TAG)) return;
    if (!(now - new Date(n.createdAt).getTime() <= TG_APPROVAL_FRESH_MS)) return;
    outbox.push({ key: 'ap:' + n.id, text: '📋 差勤批核\n' + c.slice(TG_APPROVAL_TAG.length).trim() });
  });
  // 行程前提醒：30 分鐘內要開始、有明確時間的行程（全天活動不提醒）
  getSoonEvents().forEach(ev => {
    const mins = Math.max(1, Math.round((ev.start.getTime() - now) / 60000));
    outbox.push({ key: 'ev:' + ev.key, text: '📅 ' + tpeFormat(ev.start, 'HH:mm') + ' ' + ev.title +
      (ev.location ? '（' + ev.location + '）' : '') + '\n' + mins + ' 分鐘後開始' });
  });
  // 撞期不推 Telegram（使用者交代只要行程前提醒）；撞期照樣顯示在 Ybot 首頁。
  let changed = false;
  outbox.forEach(o => {
    const keys = o.keys || [o.key];
    if (keys.every(k => sent.indexOf(k) !== -1)) return;
    try { telegramSend(o.text); keys.forEach(k => sent.push(k)); changed = true; }
    catch (err) { /* 這次失敗就下次再試，不記成已送 */ }
  });
  if (changed) p.setProperty(TG_SENT_PROP, JSON.stringify(sent.slice(-500)));
}
const TG_EVENT_LEAD_MS = 30 * 60000;
// 只抓接下來 30 分鐘內「還沒開始」的行程，連地點一起，給行程前提醒用。
// 每 10 分鐘跑一次，所以實際大約會在開始前 20～30 分鐘收到。
function getSoonEvents() {
  const now = new Date();
  const until = new Date(now.getTime() + TG_EVENT_LEAD_MS);
  const seen = {};
  const out = [];
  try {
    CalendarApp.getAllCalendars().forEach(cal => {
      try {
        cal.getEvents(now, until).forEach(ev => {
          if (ev.isAllDayEvent()) return;
          const st = ev.getStartTime();
          if (st.getTime() < now.getTime()) return;   // 已經開始的不提醒
          const key = ev.getId() + '_' + st.getTime();
          if (seen[key]) return;
          seen[key] = true;
          out.push({ key: key, title: ev.getTitle(), start: st, end: ev.getEndTime(), location: ev.getLocation() || '' });
        });
      } catch (err) { /* 單一日曆讀取失敗就跳過 */ }
    });
  } catch (err) { return []; }
  return out.sort((a, b) => a.start - b.start);
}

// ── 台股持股摘要（接「台股投資 Dashboard」的 Google Sheet 同步後台）────────────
// 只算彙總（檔數、以昨收計的總未實現損益%），不輸出每檔股數與成本。
// 網址與密碼只存在「指令碼屬性」：在編輯器執行 setStockConfig('網址', '密碼')，
// 執行完記得把括號裡的內容刪掉，不要存檔留在程式碼裡。
const STOCK_URL_PROP = 'STOCK_SYNC_URL';
const STOCK_SECRET_PROP = 'STOCK_SYNC_SECRET';
function setStockConfig(url, secret) {
  const p = PropertiesService.getScriptProperties();
  if (url) p.setProperty(STOCK_URL_PROP, String(url).trim());
  if (secret) p.setProperty(STOCK_SECRET_PROP, String(secret).trim());
  return '✅ 已儲存台股同步設定';
}
function stockToNum(v) {
  const n = Number(String(v == null ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : NaN;
}
// 台股同步後台可能用不同的參數名稱／動作名稱，依序試，第一個不是 Unauthorized／Unknown 的就用。
const STOCK_SECRET_PARAMS = ['secret', 'token', 'key', 'pass', 'password', 'k'];
const STOCK_ACTIONS = ['read', 'load', 'get', 'holdings', 'list'];
function stockFetchRaw(url, secret, param, action) {
  const r = UrlFetchApp.fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + 'action=' + action + '&' + param + '=' + encodeURIComponent(secret),
    { muteHttpExceptions: true, followRedirects: true });
  let d = null;
  try { d = JSON.parse(r.getContentText()); } catch (e) { /* 留 null */ }
  return { code: r.getResponseCode(), data: d };
}
function stockProbe(url, secret) {
  const tried = [];
  for (let i = 0; i < STOCK_SECRET_PARAMS.length; i++) {
    const param = STOCK_SECRET_PARAMS[i];
    const res = stockFetchRaw(url, secret, param, STOCK_ACTIONS[0]);
    const err = res.data && res.data.ok === false ? String(res.data.error || '') : '';
    tried.push(param + (res.data ? (err ? '→' + err : '→ok') : '→非JSON'));
    if (res.data && !/unauthor/i.test(err)) {
      // 通過驗證了；若動作不對再換動作名稱
      if (/unknown|invalid|action/i.test(err)) {
        for (let j = 1; j < STOCK_ACTIONS.length; j++) {
          const r2 = stockFetchRaw(url, secret, param, STOCK_ACTIONS[j]);
          const e2 = r2.data && r2.data.ok === false ? String(r2.data.error || '') : '';
          tried.push(param + '+' + STOCK_ACTIONS[j] + (r2.data ? (e2 ? '→' + e2 : '→ok') : '→非JSON'));
          if (r2.data && !/unauthor|unknown|invalid|action/i.test(e2)) return { param: param, action: STOCK_ACTIONS[j], data: r2.data, tried: tried };
        }
        continue;
      }
      return { param: param, action: STOCK_ACTIONS[0], data: res.data, tried: tried };
    }
  }
  return { param: '', action: '', data: null, tried: tried };
}
// 直接讀你自己的持股 Sheet（不經過同步後台、不需要金鑰）：
// 在編輯器執行一次 setStockSheet('試算表ID')；試算表ID 就是 Sheet 網址 /d/ 與 /edit 中間那串。
// 會自動找「第一列有 code、qty、cost 欄位」的分頁。只在 Ybot 後台讀取，不會寫入。
const STOCK_SHEET_PROP = 'STOCK_SHEET_ID';
function setStockSheet(id) {
  if (id) PropertiesService.getScriptProperties().setProperty(STOCK_SHEET_PROP, String(id).trim());
  return '✅ 已儲存持股 Sheet';
}
function readStockSheetDirect() {
  const id = PropertiesService.getScriptProperties().getProperty(STOCK_SHEET_PROP);
  if (!id) return null;
  const sheets = SpreadsheetApp.openById(id).getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const vals = sheets[i].getDataRange().getValues();
    if (!vals.length) continue;
    const head = vals[0].map(h => String(h).trim().toLowerCase());
    const ci = head.indexOf('code'), qi = head.indexOf('qty'), co = head.indexOf('cost');
    if (ci < 0 || qi < 0 || co < 0) continue;
    return vals.slice(1).map(r => ({ code: String(r[ci]).trim(), qty: r[qi], cost: r[co] }))
      .filter(h => h.code && stockToNum(h.qty) > 0 && stockToNum(h.cost) > 0);
  }
  return [];
}
// 讀台股後台的持股；格式不如預期就回空陣列（晨報照常，只是不顯示這一行）
function getStockHoldings() {
  const direct = readStockSheetDirect();
  if (direct) return direct;
  const p = PropertiesService.getScriptProperties();
  const url = p.getProperty(STOCK_URL_PROP), secret = p.getProperty(STOCK_SECRET_PROP);
  if (!url || !secret) return [];
  const d = stockProbe(url, secret).data;
  if (!d || d.ok === false) return [];
  let arr = d.holdings || d.data || d.rows || d;
  if (!Array.isArray(arr) && arr && typeof arr === 'object') {
    // 多使用者時可能是 { 使用者: [持股...] }，攤平成一個陣列
    arr = Object.keys(arr).reduce((acc, k) => acc.concat(Array.isArray(arr[k]) ? arr[k] : []), []);
  }
  return (Array.isArray(arr) ? arr : []).filter(h => h && h.code && stockToNum(h.qty) > 0 && stockToNum(h.cost) > 0);
}
// 證交所即時報價（上市 tse／上櫃 otc 都問一次）。盤前 z 是「-」，改用昨收 y。
function getStockPrices(codes) {
  const exch = [];
  codes.forEach(c => { exch.push('tse_' + c + '.tw'); exch.push('otc_' + c + '.tw'); });
  const out = {};
  for (let i = 0; i < exch.length; i += 40) {
    const r = UrlFetchApp.fetch('https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=' +
      encodeURIComponent(exch.slice(i, i + 40).join('|')) + '&json=1&delay=0', { muteHttpExceptions: true });
    const d = JSON.parse(r.getContentText());
    (d.msgArray || []).forEach(m => {
      const z = stockToNum(m.z), y = stockToNum(m.y);
      const price = z > 0 ? z : y;
      if (m.c && price > 0) out[m.c] = price;
    });
  }
  return out;
}
// 純計算：持股 + 報價 → { n, total, pct }。沒有任何一檔對得上報價時 n = 0。
function stockSummaryCompute(hs, prices) {
  let cost = 0, value = 0, n = 0;
  hs.forEach(h => {
    const px = prices[String(h.code).trim()];
    if (!px) return;
    const q = stockToNum(h.qty), c = stockToNum(h.cost);
    if (!(q > 0) || !(c > 0)) return;
    cost += q * c; value += q * px; n++;
  });
  return { n: n, total: hs.length, pct: cost > 0 ? (value - cost) / cost * 100 : NaN };
}
function stockSummaryText(r) {
  return '📈 台股持股 ' + r.n + ' 檔，以昨收計總未實現損益 ' + (r.pct >= 0 ? '+' : '') + r.pct.toFixed(1) + '%';
}
// 晨報用的一行彙總；任何一步失敗都回空字串，不影響晨報其他內容。
function getStockSummaryLine() {
  try {
    const hs = getStockHoldings();
    if (!hs.length) return '';
    const r = stockSummaryCompute(hs, getStockPrices(hs.map(h => String(h.code).trim())));
    return r.n && isFinite(r.pct) ? stockSummaryText(r) : '';
  } catch (err) { reportError('台股持股摘要', err, 24); return ''; }
}
// 手動測試：在編輯器執行，逐步回報卡在哪一步（只顯示欄位名稱與筆數，不顯示網址、密碼或數字）。
function testStockSummary() {
  const log = m => console.log(m);
  const p = PropertiesService.getScriptProperties();
  const url = p.getProperty(STOCK_URL_PROP), secret = p.getProperty(STOCK_SECRET_PROP);
  if (p.getProperty(STOCK_SHEET_PROP)) {
    log('① 設定：已指定持股 Sheet（直接讀取，不使用同步後台與金鑰）');
    let hsd;
    try { hsd = readStockSheetDirect() || []; } catch (e) { log('❌ 讀 Sheet 失敗：' + e.message + '（第一次可能要在跳出的視窗按「允許」授權）'); return '讀取失敗'; }
    log('② 讀到可用持股：' + hsd.length + ' 筆' + (hsd.length ? '' : '（0 筆：找不到第一列有 code、qty、cost 的分頁）'));
    if (!hsd.length) return '無持股';
    const px = getStockPrices(hsd.map(h => String(h.code).trim()));
    log('③ 抓到報價：' + Object.keys(px).length + ' / ' + hsd.length + ' 檔');
    try {
      const r = stockSummaryCompute(hsd, px);
      log('④ 對得上報價的持股：' + r.n + ' / ' + r.total + ' 檔' + (r.n < r.total ? '（沒對上的代號：' + hsd.filter(h => !px[String(h.code).trim()]).map(h => String(h.code).trim()).join('、') + '）' : ''));
      if (r.n && isFinite(r.pct)) { const ln = stockSummaryText(r); log(ln); return ln; }
      log('⑤ 算不出摘要：對得上的檔數為 0，或成本為 0。報價有的代號：' + Object.keys(px).join('、'));
      return '';
    } catch (e) { log('❌ 計算時出錯：' + e.message); return ''; }
  }
  log('① 設定：網址 ' + (url ? '有' : '❌沒有') + '、密碼 ' + (secret ? '有' : '❌沒有') + (url && secret ? '' : '（請先執行 setStockConfig）'));
  if (!url || !secret) return '未設定';
  let pr;
  try { pr = stockProbe(url, secret); } catch (e) { log('❌ 連線失敗：' + e.message); return '連線失敗'; }
  log('② 嘗試結果（參數名稱→回應）：' + pr.tried.join('；'));
  if (!pr.data) {
    log('❌ 所有參數名稱都被拒絕或回傳不是 JSON。→ 金鑰與後台不一致（或後台沒有重新部署新版本），請把後台的 SECRET_TOKEN 與這裡存的重新對一次');
    return '驗證失敗';
  }
  log('③ 通過驗證：參數「' + pr.param + '」、動作「' + pr.action + '」；回傳最上層欄位：' + (Array.isArray(pr.data) ? '（陣列，' + pr.data.length + ' 筆）' : Object.keys(pr.data).join(', ')));
  const hs = getStockHoldings();
  log('④ 讀到可用持股：' + hs.length + ' 筆' + (hs.length ? '；欄位：' + Object.keys(hs[0]).join(', ') : '（0 筆代表格式不符或 qty／cost 讀不到）'));
  if (!hs.length) return '無持股';
  const prices = getStockPrices(hs.map(h => String(h.code).trim()));
  log('⑤ 抓到報價：' + Object.keys(prices).length + ' / ' + hs.length + ' 檔');
  const line = getStockSummaryLine();
  log(line || '⑥ 仍然沒有摘要：請把上面 ③④ 的欄位名稱貼給我');
  return line;
}

// ── Telegram 提醒按鈕（✅ 完成／⏰ 延後 1 小時）＋ 每分鐘輪詢按鈕 ─────────────
// 不用 webhook（不必公開後台網址、也不會有 Apps Script 302 重送問題），
// 改由 setupTelegramPoll() 排程的觸發器每分鐘用 getUpdates 收按鈕，約 1 分鐘內生效。
// 只處理來自你自己 chat id 的按鈕，其他人按了無效。
const TG_OFFSET_PROP = 'TELEGRAM_UPDATE_OFFSET';
const SNOOZE_MS = 3600000;
function telegramSendButtons(text, keyboard) {
  const { chatId } = tgConfig();
  if (!chatId) throw new Error('還沒設定 chat id');
  return tgApi('sendMessage', { chat_id: chatId, text: String(text).slice(0, 4000),
    disable_web_page_preview: true, reply_markup: { inline_keyboard: keyboard } });
}
function sendReminderTelegram(n) {
  try {
    telegramSendButtons('⏰ 提醒\n' + n.content, [[
      { text: '✅ 完成', callback_data: 'done:' + n.id },
      { text: '⏰ 延後 1 小時', callback_data: 'snz:' + n.id }
    ]]);
    return true;
  } catch (err) { return false; }
}
function setupTelegramPoll() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'telegramPoll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('telegramPoll').timeBased().everyMinutes(1).create();
  const msg = '✅ 已排程每分鐘接收 Telegram 按鈕（提醒的「完成／延後」）' + tzWarn();
  console.log(msg);
  return msg;
}
function telegramPollImpl() {
  const cfg = tgConfig();
  if (!cfg.token || !cfg.chatId) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;          // 上一輪還沒跑完就跳過，避免同一個按鈕處理兩次
  try {
    const p = tgProps();
    const offset = Number(p.getProperty(TG_OFFSET_PROP) || 0);
    const updates = tgApi('getUpdates', { offset: offset || undefined, timeout: 0, allowed_updates: ['message', 'callback_query'] }) || [];
    if (!updates.length) return;
    let next = offset;
    updates.forEach(u => {
      next = Math.max(next, u.update_id + 1);
      if (u.callback_query) {
        try { handleReminderCallback(u.callback_query, cfg.chatId); } catch (err) { /* 單一按鈕失敗不擋其他 */ }
      }
    });
    p.setProperty(TG_OFFSET_PROP, String(next));
  } finally { lock.releaseLock(); }
}
function handleReminderCallback(cq, chatId) {
  const msg = cq.message;
  if (!msg || !msg.chat || String(msg.chat.id) !== String(chatId)) {
    try { tgApi('answerCallbackQuery', { callback_query_id: cq.id }); } catch (e) { /* 忽略 */ }
    return;
  }
  const m = String(cq.data || '').match(/^(done|snz):(.+)$/);
  let reply = '無法辨識的按鈕';
  if (m) {
    const { note } = setupSheets();
    const n = sheetToObjects(note, NOTE_COLS).find(x => String(x.id) === m[2]);
    if (!n) {
      reply = '找不到這筆提醒（可能已刪除）';
    } else if (m[1] === 'done') {
      updatePartial(note, NOTE_COLS, n.id, { done: 'true' });
      reply = '✅ 已完成';
    } else {
      const due = new Date(Date.now() + SNOOZE_MS);
      updatePartial(note, NOTE_COLS, n.id, { dueAt: due.toISOString(), notifiedAt: '' });
      reply = '⏰ 已延後，' + tpeFormat(due, 'HH:mm') + ' 再提醒';
    }
  }
  try { tgApi('answerCallbackQuery', { callback_query_id: cq.id, text: reply }); } catch (e) { /* 忽略 */ }
  // 把訊息改成結果並拿掉按鈕（不給 reply_markup 就會移除按鈕），避免重複按
  try { tgApi('editMessageText', { chat_id: chatId, message_id: msg.message_id, text: String(msg.text || '').slice(0, 3900) + '\n\n' + reply }); } catch (e) { /* 忽略 */ }
}

// ── 出錯自動通知（Telegram，沒設定就寄 Email）────────────────────────
// 同一項在 quietHours 小時內只通知一次，不會洗版。訊息只放項目名稱與錯誤摘要，
// 並把網址與 secret 參數遮掉，不會把任何內容資料送出去。
const ERR_PROP = 'ERROR_ALERT_SENT';
function scrubError(text) {
  return String(text || '').replace(/https?:\/\/\S+/g, '[網址]').replace(/(secret|token|key|pass|password)=\S+/gi, '$1=***').slice(0, 300);
}
function reportError(name, err, quietHours) {
  try {
    const q = (quietHours === undefined ? 6 : quietHours) * 3600000;
    const p = tgProps();
    let seen = {};
    try { seen = JSON.parse(p.getProperty(ERR_PROP) || '{}'); } catch (e) { seen = {}; }
    const now = Date.now();
    if (seen[name] && now - seen[name] < q) return;
    seen[name] = now;
    p.setProperty(ERR_PROP, JSON.stringify(seen));
    const text = '⚠️ Ybot 後台異常：' + name + '\n' + scrubError((err && err.message) || err) +
      '\n（' + tpeDateTimeStr(new Date()) + '；同一項 ' + (q / 3600000) + ' 小時內不重複通知）';
    const tg = tgConfig();
    if (tg.token && tg.chatId) { try { telegramSend(text); return; } catch (e) { /* 退回 Email */ } }
    const email = NOTIFY_EMAIL || getOwnerEmail();
    if (email) MailApp.sendEmail(email, '⚠️ Ybot 後台異常：' + name, text);
  } catch (e) { /* 通知本身失敗不能再讓程式出錯 */ }
}
function guard(name, fn) {
  try { return fn(); } catch (err) { reportError(name, err); throw err; }
}
function testErrorReport() {
  reportError('測試', new Error('這是一則測試訊息，代表出錯通知正常（可忽略）'), 0);
  console.log('已送出測試通知，請看 Telegram（沒設定 Telegram 就看 Email）');
}

// 每天早上 8:30 檢查：今天的每日簡報有沒有產生、核心排程有沒有被刪掉。
function setupHealthCheck() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'healthCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('healthCheck').timeBased().everyDays(1).atHour(8).nearMinute(30).create();
  const msg = '✅ 已排程每天早上 8:30 檢查晨報與排程是否正常（有問題才通知）' + tzWarn();
  console.log(msg);
  return msg;
}
function healthCheckImpl() {
  const problems = [];
  const sh = setupSheets().brief;
  const last = sh.getLastRow();
  let lastAt = null;
  if (last > 1) { const d = new Date(sh.getRange(last, 1).getValue()); if (!isNaN(d)) lastAt = d; }
  if (!lastAt || tpeDateStr(lastAt) !== tpeDateStr(new Date())) {
    problems.push('今天的每日簡報沒有產生（最後一次：' + (lastAt ? tpeDateTimeStr(lastAt) : '從未') + '）');
  }
  const have = {};
  ScriptApp.getProjectTriggers().forEach(t => { have[t.getHandlerFunction()] = true; });
  const need = ['dailyBrief', 'reminderWatch'];
  if (tgConfig().token) need.push('telegramWatch', 'telegramPoll');
  const missing = need.filter(h => !have[h]);
  if (missing.length) problems.push('缺少排程：' + missing.join('、') + '（請重新執行對應的 setup 函式）');
  if (problems.length) reportError('健康檢查', { message: problems.join('；') }, 20);
}
function testHealthCheck() {
  healthCheckImpl();
  console.log('已執行健康檢查；有問題才會通知，沒有通知就代表正常');
}

// ── 排程入口：出錯時自動通知 ──
function dailyBrief() { return guard('dailyBrief', dailyBriefImpl); }
function reminderWatch() { return guard('reminderWatch', reminderWatchImpl); }
function eveningDigest() { return guard('eveningDigest', eveningDigestImpl); }
function weeklyReview() { return guard('weeklyReview', weeklyReviewImpl); }
function standupWatch() { return guard('standupWatch', standupWatchImpl); }
function telegramWatch() { return guard('telegramWatch', telegramWatchImpl); }
function healthCheck() { return guard('healthCheck', healthCheckImpl); }
function telegramPoll() { return guard('telegramPoll', telegramPollImpl); }

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
