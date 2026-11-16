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
 * 3-1. 最簡單：執行一次 setupAll()，會把下面所有排程一次裝好（已存在的不動）。
 *    想分開裝也可以：執行一次 setupDailyBrief()（授權後排程「每日簡報」）
 *    再執行一次 setupReminderWatch()（排程「到點提醒」，每 10 分鐘檢查一次）
 *    再執行一次 setupEveningDigest()（排程「晚間彙整」，只在有未處理事項時才寄信）
 *    再執行一次 setupStandupWatch()（排程「起立提醒」，實際頻率依 ybot.html 設定調整）
 *    再執行一次 setupWeeklyReview()（排程「本週回顧」，每週一早上 7 點）
 * 3-2.（選填）Telegram 通知：手機鎖著也收得到新的【差勤批核】筆記和聯邦帳單。
 *    a. 在 Telegram 找 @BotFather 建立機器人（或沿用股票提醒那一個），拿到 token
 *    b. 專案設定（齒輪）→ 指令碼屬性 → 新增屬性 TELEGRAM_BOT_TOKEN，值貼 token
 *       （token 只放這裡，不要寫進程式碼）
 *    c. 在 Telegram 對你的機器人隨便傳一句話，再執行一次 findTelegramChatId()
 *    d. 執行一次 setupTelegramWatch()（每 10 分鐘檢查一次，並先傳一則測試訊息）
 *    會推：新的【差勤批核】筆記、聯邦帳單、行程開始前約 20～30 分鐘提醒
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
// 後台版本：每次改這個檔案就把日期往後改。前端設定頁會拿它跟 GitHub 上的最新版比，
// 不一樣就提醒「後台還沒更新」——不然貼新程式這一步漏了，手機上完全看不出來。
const BACKEND_VERSION = '2026-11-16-4';
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
// 定時觸發器（每日簡報 7 點、晚間提醒 8 點…）是照「專案時區」發動的，
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
// 中央氣象署開放資料的授權碼（免費，到 https://opendata.cwa.gov.tw 註冊後取得）。
// 只存在這個後台的 PropertiesService，不要貼進程式碼。在 Apps Script 編輯器
// 執行 setCwaKey('CWA-xxxx') 一次即可；沒設就維持原本的 Open-Meteo 溫度範圍。
function setCwaKey(key) {
  PropertiesService.getScriptProperties().setProperty('CWA_API_KEY', String(key || '').trim());
  CacheService.getScriptCache().remove('weatherDigest');
  return key ? '✅ 已儲存中央氣象署授權碼' : '✅ 已清除中央氣象署授權碼（改回 Open-Meteo）';
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
      tokenRequired: !!getApiToken(), version: BACKEND_VERSION });
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
  if (action === 'savePaDays') {
    return jsonResp({ ok: true, message: savePaDays(body.days) });
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
  if (action === 'agent') {
    // App 對話：前端把問題與最近幾輪對話送來，後台跑 agent 迴圈後回答。
    const q = String(body.question || '').trim();
    if (!q) return jsonResp({ ok: false, error: '問題是空的' });
    const r = runAgent(q, Array.isArray(body.history) ? body.history : [], { channel: 'app' });
    if (r.error && !r.text) return jsonResp({ ok: false, error: r.error, steps: r.steps || [] });
    return jsonResp({ ok: true, text: r.text, steps: (r.steps || []).map(st => ({ tool: st.tool, label: agentToolLabel(st.tool), ok: st.ok })), warning: r.error || '' });
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

// ── 聯邦帳單 ──────────────────────────────────
// 有「聯邦」兩個字的信就是聯邦銀行的帳單：不管已讀未讀、抓近 45 天最多 3 封，
// 從全文抓本期應繳金額和繳款截止日（抓不到就留空，前端只顯示主旨）。
function parseBillText(text) {
  const t = String(text || '').replace(/\s+/g, ' ');
  const a = t.match(/(?:本期)?應繳(?:總)?(?:金額|總額|款項)?\s*[:：]?\s*(?:NT\$|新[臺台]幣|TWD)?\s*(-?[\d,]+(?:\.\d+)?)\s*元?/);
  const d = t.match(/(?:繳款截止日|繳款期限|最後繳款日|繳費期限|繳款截止日期)\s*[:：]?\s*(\d{2,4}\s*[\/\-.年]\s*\d{1,2}\s*[\/\-.月]\s*\d{1,2})/);
  return { amount: a ? a[1] : '', due: d ? d[1].replace(/\s+/g, '').replace(/[年月]/g, '/') : '' };
}
function getBillDigest() {
  try {
    const threads = GmailApp.search('聯邦 newer_than:45d -from:me', 0, 10);
    const out = [];
    threads.forEach(t => {
      const msgs = t.getMessages();
      const last = msgs[msgs.length - 1];
      const subject = t.getFirstMessageSubject() || '';
      const body = (last.getPlainBody() || '').replace(/\s+/g, ' ');
      if (!/聯邦/.test(subject + ' ' + last.getFrom() + ' ' + body)) return;
      const p = parseBillText(body);
      out.push({
        subject, from: last.getFrom(), date: last.getDate().toISOString(),
        snippet: body.slice(0, 120), amount: p.amount, due: p.due,
        unread: t.isUnread(), threadId: t.getId()
      });
    });
    out.sort((a, b) => new Date(b.date) - new Date(a.date));
    return out.slice(0, 3);
  } catch (err) { return []; }
}

// ── 彙整上下文：Gmail + 日曆 + 新聞 + 天氣 + 起立提醒設定 + 待辦提醒 + 其他系統摘要 ──
function buildContext() {
  return {
    generatedAt: new Date().toISOString(),
    notes: sheetToObjects(setupSheets().note, NOTE_COLS),
    gmail: getGmailDigest(),
    bills: getBillDigest(),
    calendar: getCalendarDigest(),
    news: getNewsDigest(),
    weather: getWeatherDigest(),
    standup: getStandupSettings(),
    linked: getLinkedSummaries(),
    stock: getStockDetail(),
    pa: getPaDay()
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
        rainChance: (daily.precipitation_probability_max || [])[0] ?? cur.precipitation_probability ?? null,
        src: 'open-meteo'
      };
      // 有設中央氣象署授權碼就改用氣象署的當日最高／最低溫（和手機氣象 App 對得上）
      try {
        const cwa = getCwaTodayRange(kv.weatherLat && kv.weatherLon ? null : (kv.weatherCity || '台北'), coords);
        if (cwa) {
          // 傍晚 17 時以後氣象署只發「今晚」以後的預報，白天那一段已經不在了，
          // 直接用會把今天最高溫低估；這時和 Open-Meteo 的全天值取較極端的。
          const hi = cwa.hasDay ? cwa.tMax : Math.max(cwa.tMax, out.tMax == null ? -99 : out.tMax);
          const lo = cwa.hasDay ? cwa.tMin : Math.min(cwa.tMin, out.tMin == null ? 99 : out.tMin);
          out.tMax = hi; out.tMin = lo;
          if (cwa.rain != null) out.rainChance = cwa.rain;
          out.src = 'cwa';
        }
      } catch (err) { /* 氣象署連不上就維持 Open-Meteo 的值 */ }
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
// ── 中央氣象署（CWA）今明 36 小時預報：取今天的最高／最低溫 ──────────
// 資料集 F-C0032-001，以縣市為單位、每 12 小時一段，要授權碼（見 setCwaKey）。
// 縣市名稱要用「臺」「縣／市」的正式寫法；GPS 定位時挑離座標最近的縣市。
const CWA_COUNTIES = [
  ['臺北市', 25.033, 121.565], ['新北市', 25.012, 121.465], ['桃園市', 24.994, 121.301],
  ['臺中市', 24.148, 120.674], ['臺南市', 22.999, 120.227], ['高雄市', 22.627, 120.301],
  ['基隆市', 25.128, 121.739], ['新竹市', 24.804, 120.969], ['新竹縣', 24.839, 121.007],
  ['苗栗縣', 24.560, 120.821], ['彰化縣', 24.075, 120.542], ['南投縣', 23.961, 120.972],
  ['雲林縣', 23.709, 120.541], ['嘉義市', 23.480, 120.449], ['嘉義縣', 23.459, 120.293],
  ['屏東縣', 22.668, 120.488], ['宜蘭縣', 24.757, 121.753], ['花蓮縣', 23.991, 121.611],
  ['臺東縣', 22.756, 121.150], ['澎湖縣', 23.571, 119.579], ['金門縣', 24.449, 118.377],
  ['連江縣', 26.158, 119.950]
];
function cwaCountyName(cityLabel, coords) {
  if (cityLabel) {
    const t = String(cityLabel).replace(/台/g, '臺').trim();
    for (let i = 0; i < CWA_COUNTIES.length; i++) {
      if (CWA_COUNTIES[i][0] === t || CWA_COUNTIES[i][0].indexOf(t) === 0) return CWA_COUNTIES[i][0];
    }
  }
  if (!coords) return null;
  let best = null, bestD = Infinity;
  CWA_COUNTIES.forEach(c => {
    const dLat = c[1] - Number(coords.lat), dLon = (c[2] - Number(coords.lon)) * 0.9;
    const d = dLat * dLat + dLon * dLon;
    if (d < bestD) { bestD = d; best = c[0]; }
  });
  return best;
}
function getCwaTodayRange(cityLabel, coords) {
  const key = PropertiesService.getScriptProperties().getProperty('CWA_API_KEY');
  if (!key) return null;
  const county = cwaCountyName(cityLabel, coords);
  if (!county) return null;
  const url = 'https://opendata.cwa.gov.tw/api/v1/rest/datastore/F-C0032-001?Authorization=' + encodeURIComponent(key) +
    '&format=JSON&locationName=' + encodeURIComponent(county) + '&elementName=MinT,MaxT,PoP';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const d = JSON.parse(res.getContentText());
  const loc = d && d.records && d.records.location && d.records.location[0];
  if (!loc) return null;
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const pick = name => {
    const el = (loc.weatherElement || []).filter(e => e.elementName === name)[0];
    return el ? el.time || [] : [];
  };
  // 只取「開始時間在今天」的時段；時間字串是台北時間，如 2026-11-16 06:00:00
  const todayOnly = arr => arr.filter(t => String(t.startTime).slice(0, 10) === today);
  const num = t => Number(t.parameter && t.parameter.parameterName);
  const mins = todayOnly(pick('MinT')).map(num).filter(isFinite);
  const maxs = todayOnly(pick('MaxT')).map(num).filter(isFinite);
  if (!mins.length || !maxs.length) return null;
  const rains = todayOnly(pick('PoP')).map(num).filter(isFinite);
  // 今天有從早上（18 點以前）開始的時段，表示白天預報還在
  const hasDay = todayOnly(pick('MaxT')).some(t => Number(String(t.startTime).slice(11, 13)) < 18);
  return {
    tMin: Math.min.apply(null, mins), tMax: Math.max.apply(null, maxs),
    rain: rains.length ? Math.max.apply(null, rains) : null, hasDay: hasDay, county: county
  };
}
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
    // 檢驗數據庫（糖化血色素、血糖、血脂…）：每一項只取最新一筆，併進健康摘要，
    // 這樣在 Ybot 對話裡問「我的糖化血色素是多少」才答得出來。
    let labsLatest = null;
    try {
      const lr = UrlFetchApp.fetch(kv.healthBackendUrl + '?action=labs', { muteHttpExceptions: true });
      const ld = JSON.parse(lr.getContentText());
      if (ld.ok && Array.isArray(ld.labs)) labsLatest = latestLabsByMetric(ld.labs);
    } catch (err) { /* 檢驗資料讀不到就維持原本的摘要 */ }
    try {
      const res = UrlFetchApp.fetch(kv.healthBackendUrl + '?action=latest', { muteHttpExceptions: true });
      const d = JSON.parse(res.getContentText());
      if (d.ok && d.snapshot) {
        out.health = {
          date: d.snapshot.date, score: d.snapshot.score, bio: d.snapshot.bio,
          cvdLevel: d.snapshot.cvdLevel, dmLevel: d.snapshot.dmLevel
        };
        if (labsLatest && Object.keys(labsLatest).length) out.health.labs = labsLatest;
        try { PropertiesService.getScriptProperties().setProperty('HEALTH_LATEST', JSON.stringify(out.health)); } catch (e) { /* 只是快取 */ }
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
  return '✅ 已排程每日早上 7 點自動產生 Ybot 簡報（台北時間）' + tzWarn();
}

function dailyBriefImpl() {
  // 順手確認核心排程都還在（例如健康檢查被刪掉時由這裡補回），失敗不影響簡報
  try { ensureTriggers(false); } catch (err) { /* 略過 */ }
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
  if ((ctx.bills || []).length) {
    lines.push('💳 聯邦帳單：');
    ctx.bills.forEach(m => lines.push('　・' + m.subject +
      (m.amount ? '｜應繳 ' + m.amount + ' 元' : '') + (m.due ? '｜截止 ' + m.due : '')));
    lines.push('');
  }
  if (ctx.gmail.length) {
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
  const paLines = paDayLines();
  if (paLines.length) { lines.push('🏫 校長助理：'); paLines.forEach(function (t) { lines.push('　・' + t); }); lines.push(''); }
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

// Telegram 版晨報：手機上掃一眼用，所以只放天氣、撞期、今明行程、🔴🟠待辦、帳單、Ybot 的話。
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
  if ((ctx.bills || []).length) {
    L.push('', '💳 帳單：');
    ctx.bills.slice(0, 3).forEach(m => L.push('　・' + m.subject + (m.amount ? '｜' + m.amount + ' 元' : '') + (m.due ? '｜截止 ' + m.due : '')));
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
  return '✅ 已排程每 10 分鐘檢查一次到點提醒' + tzWarn();
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
// 手機鎖著也會響。推：新的【差勤批核】筆記、聯邦帳單、行程前提醒。
// token 和 chat id 只存在「指令碼屬性」，不寫進程式碼。
const TG_TOKEN_PROP = 'TELEGRAM_BOT_TOKEN';
const TG_CHAT_PROP = 'TELEGRAM_CHAT_ID';
const TG_SENT_PROP = 'TELEGRAM_SENT_IDS';
const TG_APPROVAL_TAG = '【差勤批核】';
const TG_APPROVAL_FRESH_MS = 36 * 3600000;   // 跟前端首頁一樣：一天半內的才算新的
const TG_BILL_FRESH_MS = 2 * 86400000;       // 帳單只推 2 天內寄來的，舊信不吵
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
  telegramSend('✅ Ybot Telegram 通知設定成功。之後有新的差勤批核、聯邦帳單、行程前提醒會傳到這裡。');
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
  getBillDigest().forEach(m => {
    if (!(now - new Date(m.date).getTime() <= TG_BILL_FRESH_MS)) return;
    const info = [m.amount ? '應繳 ' + m.amount + ' 元' : '', m.due ? '截止 ' + m.due : ''].filter(Boolean).join('｜');
    outbox.push({ key: 'bill:' + m.threadId, text: '💳 聯邦帳單\n' + m.subject + (info ? '\n' + info : '') });
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
      } else if (u.message) {
        try { handleTelegramText(u.message, cfg.chatId); } catch (err) { reportError('Telegram 指令', err); }
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
  // 待辦清單上的「✅」與剛新增待辦的「↩️ 取消」（見下方 Telegram 指令）
  const lm = String(cq.data || '').match(/^(ld|del):(.+)$/);
  if (lm) { handleListCallback(cq, chatId, lm[1], lm[2]); return; }
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
  // 自我修復：核心排程被刪掉就直接補回來，不必等使用者回來重跑 setup
  const fixed = [];
  try {
    const repaired = ensureTriggers(false);
    if (repaired.length) fixed.push('已自動補回排程：' + repaired.join('、'));
  } catch (err) {
    problems.push('排程檢查失敗：' + scrubError(err && err.message));
  }
  // 自我修復：今天的簡報沒產生就補發一次；補發也失敗才算問題
  if (problems.length && problems[0].indexOf('今天的每日簡報沒有產生') === 0) {
    try {
      dailyBriefImpl();
      fixed.push('今天的每日簡報原本沒有產生，已自動補發');
      problems.shift();
    } catch (err) {
      problems[0] += '；自動補發也失敗：' + scrubError(err && err.message);
    }
  }
  if (problems.length) reportError('健康檢查', { message: problems.concat(fixed).join('；') }, 20);
  else if (fixed.length) reportError('健康檢查（已自動修復）', { message: fixed.join('；') }, 20);
}
function testHealthCheck() {
  healthCheckImpl();
  console.log('已執行健康檢查；有問題才會通知，沒有通知就代表正常');
}

// ── 排程自我修復＋一鍵安裝 ─────────────────────────────────────────
// 每個排程原本要各自執行一次 setup 函式，漏掉一個、或排程被刪掉，那項功能就默默停擺。
// ensureTriggers 只補「缺少的」，已存在的不動（不會重設時間）。
//   core：一定要有；tg：有設定 Telegram 才需要；其餘是選用，只有 setupAll() 會裝。
const TRIGGER_SPECS = [
  { h: 'dailyBrief', kind: 'core', make: function () { ScriptApp.newTrigger('dailyBrief').timeBased().everyDays(1).atHour(7).create(); } },
  { h: 'reminderWatch', kind: 'core', make: function () { ScriptApp.newTrigger('reminderWatch').timeBased().everyMinutes(10).create(); } },
  { h: 'healthCheck', kind: 'core', make: function () { ScriptApp.newTrigger('healthCheck').timeBased().everyDays(1).atHour(8).nearMinute(30).create(); } },
  { h: 'telegramWatch', kind: 'tg', make: function () { ScriptApp.newTrigger('telegramWatch').timeBased().everyMinutes(10).create(); } },
  { h: 'telegramPoll', kind: 'tg', make: function () { ScriptApp.newTrigger('telegramPoll').timeBased().everyMinutes(1).create(); } },
  { h: 'autoWatch', kind: 'tg', make: function () { ScriptApp.newTrigger('autoWatch').timeBased().everyMinutes(15).create(); } },
  { h: 'eveningDigest', kind: 'optional', make: function () { ScriptApp.newTrigger('eveningDigest').timeBased().everyDays(1).atHour(20).create(); } },
  { h: 'weeklyReview', kind: 'optional', make: function () { ScriptApp.newTrigger('weeklyReview').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create(); } },
  { h: 'standupWatch', kind: 'optional', make: function () { ScriptApp.newTrigger('standupWatch').timeBased().everyMinutes(15).create(); } }
];
// 回傳這次補回來的排程名稱（沒有缺就回傳空陣列）
function ensureTriggers(includeOptional) {
  const have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  const tg = tgConfig();
  const tgReady = !!(tg.token && tg.chatId);
  const created = [];
  TRIGGER_SPECS.forEach(function (s) {
    if (have[s.h]) return;
    if (s.kind === 'tg' && !tgReady) return;
    if (s.kind === 'optional' && !includeOptional) return;
    s.make();
    created.push(s.h);
  });
  return created;
}
// 一鍵安裝：把所有排程一次裝好（已經有的不動）。新部署或懷疑漏裝時執行這一個就好。
function setupAll() {
  const created = ensureTriggers(true);
  const tg = tgConfig();
  const msg = (created.length ? '✅ 已補上排程：' + created.join('、') : '✅ 所有排程都已存在，沒有變動') +
    ((tg.token && tg.chatId) ? '' : '\n（還沒設定 Telegram，所以沒有裝 telegramWatch／telegramPoll）') + tzWarn();
  console.log(msg);
  return msg;
}

// ── Telegram 指令：在手機上直接對 Ybot 下指令 ──────────────────────────
// 由 telegramPoll（每分鐘）收訊息，所以約 1 分鐘內回覆。只理會自己 chat id 傳來的文字，
// 超過 10 分鐘的舊訊息不處理（避免剛裝好時把以前傳的「hi」變成待辦）。
//   直接打字            → 加成待辦（附「↩️ 取消」按鈕）
//   提醒 明天 9:00 開會  → 定時提醒（也接受 14:30、下午3點、30分鐘後、2小時後、11/20 9:00）
//   筆記 內容            → 瑣事筆記
//   待辦（或 /list）     → 列出未完成事項，每筆一個 ✅ 按鈕
//   今天（或 /now）      → 今明行程＋要處理的事
//   狀態（或 /status）   → 排程與最近一次簡報
//   問 問題（或 /ask）   → 用已設定的 AI 依待辦與行事曆回答（只送待辦與行程，不送信件）
//   說明（或 /help）
const TG_MSG_MAX_AGE_MS = 10 * 60000;
const TG_HELP = [
  '🤖 Ybot 指令',
  '直接打字 → 由 AI 判斷：查資料、做事、或加成待辦（agent 模式）',
  '待辦 內容 → 一定加成待辦',
  '重來 → 清除這段對話的記憶',
  'agent off / agent on → 關閉／開啟 agent 模式',
  '提醒 明天 9:00 開會 → 定時提醒',
  '　（也可以：14:30、下午3點、30分鐘後、2小時後、11/20 9:00）',
  '筆記 內容 → 存成筆記',
  '待辦 → 未完成清單（可按 ✅）',
  '今天 → 今明行程與要處理的事',
  '股票 → 持股今日漲跌與總損益%',
  '健康 → 最新健康分與提醒',
  '校長 → 校長助理今明記事與代課節數',
  '狀態 → 排程是否正常',
  '問 問題 → 請 AI 回答（可查持股、檢驗、行事曆、郵件）'
].join('\n');

function tgOpenItems() {
  return sheetToObjects(setupSheets().note, NOTE_COLS)
    .filter(function (n) { return (n.type === 'todo' || n.type === 'reminder') && n.done !== 'true'; });
}
function tgAddNote(type, content, dueAt) {
  const id = Utilities.getUuid();
  appendObj(setupSheets().note, NOTE_COLS, {
    id: id, type: type, content: content, dueAt: dueAt || '', done: '',
    createdAt: new Date().toISOString(), notifiedAt: '', owner: ''
  });
  return id;
}
// 把「提醒」後面的文字拆成時間與內容；拆不出來回傳 null。時間一律當台北時間。
function parseTgReminder(body, now) {
  const s = String(body || '').trim();
  now = now || new Date();
  let m = s.match(/^(\d{1,3})\s*(分鐘|分|小時|個小時)後\s*(.+)$/);
  if (m) {
    const ms = Number(m[1]) * (m[2].indexOf('分') === 0 ? 60000 : 3600000);
    return { dueAt: new Date(now.getTime() + ms), content: m[3].trim() };
  }
  const pad = function (n) { return (Number(n) < 10 ? '0' : '') + Number(n); };
  const at = function (y, mo, d, h, mi) {
    return new Date(y + '-' + pad(mo) + '-' + pad(d) + 'T' + pad(h) + ':' + pad(mi) + ':00+08:00');
  };
  const today = tpeFormat(now, 'yyyy-M-d').split('-').map(Number);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\s+(\d{1,2})[:：](\d{2})\s*(.+)$/);
  if (m) {
    if (Number(m[1]) < 1 || Number(m[1]) > 12 || Number(m[2]) < 1 || Number(m[2]) > 31 || Number(m[3]) > 23 || Number(m[4]) > 59) return null;
    let due = at(today[0], m[1], m[2], m[3], m[4]);
    if (due <= now) due = at(today[0] + 1, m[1], m[2], m[3], m[4]);   // 日期已過就當明年
    return isNaN(due) ? null : { dueAt: due, content: m[5].trim() };
  }
  m = s.match(/^(今天|明天|後天)?\s*(上午|早上|中午|下午|晚上)?\s*(\d{1,2})(?:[:：](\d{2})|點(?:(\d{1,2})分?|半)?)\s*(.+)$/);
  if (m) {
    let h = Number(m[3]);
    const mi = m[4] !== undefined ? Number(m[4]) : (m[5] !== undefined ? Number(m[5]) : (/點半/.test(s) ? 30 : 0));
    if ((m[2] === '下午' || m[2] === '晚上') && h < 12) h += 12;
    if (m[2] === '中午' && h < 11) h += 12;
    if (h > 23 || mi > 59) return null;
    const base = at(today[0], today[1], today[2], h, mi);
    if (isNaN(base)) return null;
    const add = m[1] === '明天' ? 1 : (m[1] === '後天' ? 2 : 0);
    let due = new Date(base.getTime() + add * 86400000);
    if (!m[1] && due <= now) due = new Date(due.getTime() + 86400000);   // 沒指定哪天、時間已過就當明天
    return { dueAt: due, content: m[6].trim() };
  }
  return null;
}
function tgListView() {
  const now = new Date();
  const items = tgOpenItems();
  if (!items.length) return { text: '🎉 沒有未完成的待辦', keyboard: [] };
  const order = { red: 0, orange: 1, green: 2 };
  const icon = { red: '🔴', orange: '🟠', green: '🟢' };
  const rows = items.map(function (n) { return { n: n, level: ruleLevel(n, now) }; })
    .sort(function (a, b) { return order[a.level] - order[b.level]; });
  const shown = rows.slice(0, 12);
  const L = ['📝 未完成 ' + items.length + ' 項' + (items.length > shown.length ? '（列前 ' + shown.length + ' 項）' : '') + '：'];
  shown.forEach(function (r, i) {
    L.push((i + 1) + '. ' + icon[r.level] + ' ' + String(r.n.content).slice(0, 60) + (r.n.dueAt ? '（' + fmtDue(r.n.dueAt) + '）' : ''));
  });
  const keyboard = shown.map(function (r, i) {
    return [{ text: '✅ ' + (i + 1) + '. ' + String(r.n.content).slice(0, 24), callback_data: 'ld:' + r.n.id }];
  }).filter(function (row) { return row[0].callback_data.length <= 64; });   // Telegram 上限 64 bytes
  return { text: L.join('\n'), keyboard: keyboard };
}
function tgNowText() {
  const now = new Date();
  const L = ['📍 ' + tpeDateTimeStr(now) + ' 現況'];
  const todayStr = tpeFormat(now, 'yyyy-MM-dd');
  const tomorrowStr = tpeFormat(new Date(now.getTime() + 86400000), 'yyyy-MM-dd');
  const near = (getCalendarDigest() || []).filter(function (ev) {
    const d = tpeFormat(new Date(ev.start), 'yyyy-MM-dd');
    return d === todayStr || d === tomorrowStr;
  });
  if (near.length) {
    L.push('', '📅 今明行程：');
    near.slice(0, 10).forEach(function (ev) {
      L.push('　' + (ev.allDay ? '全天' : tpeDateTimeStr(new Date(ev.start))) + ' ' + ev.title);
    });
  } else {
    L.push('', '📅 今明沒有行程');
  }
  const items = tgOpenItems();
  const red = items.filter(function (n) { return ruleLevel(n, now) === 'red'; });
  const orange = items.filter(function (n) { return ruleLevel(n, now) === 'orange'; });
  if (red.length || orange.length) {
    L.push('', '🎯 要處理：');
    red.forEach(function (n) { L.push('🔴 ' + String(n.content).slice(0, 60)); });
    orange.slice(0, 5).forEach(function (n) { L.push('🟠 ' + String(n.content).slice(0, 60)); });
  } else {
    L.push('', '🎯 沒有今明到期的待辦');
  }
  L.push('', '其他未完成：' + (items.length - red.length - orange.length) + ' 項（傳「待辦」看清單）');
  const pa = paDayLines();
  if (pa.length) { L.push('', '🏫 校長助理：'); pa.forEach(function (t) { L.push('　' + t); }); }
  const st = getStockDetail();
  if (st) L.push('', stockOneLine(st));
  const hl = healthLatest();
  if (hl) L.push('', healthOneLine(hl));
  return L.join('\n');
}
function tgStatusText() {
  const have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  const L = ['🩺 Ybot 後台狀態　' + tpeDateTimeStr(new Date())];
  TRIGGER_SPECS.forEach(function (s) {
    L.push((have[s.h] ? '✅ ' : (s.kind === 'optional' ? '➖ ' : '❌ ')) + s.h + (have[s.h] ? '' : (s.kind === 'optional' ? '（選用，未安裝）' : '（缺少，健康檢查會自動補回）')));
  });
  const sh = setupSheets().brief;
  const last = sh.getLastRow();
  let lastAt = null;
  if (last > 1) { const d = new Date(sh.getRange(last, 1).getValue()); if (!isNaN(d)) lastAt = d; }
  L.push('', '最近一次每日簡報：' + (lastAt ? tpeDateTimeStr(lastAt) : '從未'));
  L.push('AI：' + (hasAnyAiKey(getAiConfig()) ? '已設定' : '未設定'));
  return L.join('\n');
}
function tgAskAi(question) {
  if (!hasAnyAiKey(getAiConfig())) return '還沒設定 AI 金鑰（在 Apps Script 執行 setAiConfig），所以無法回答問題。';
  const now = new Date();
  const items = tgOpenItems().slice(0, 40).map(function (n) {
    return '- [' + n.type + '] ' + n.content + (n.dueAt ? '（到期 ' + fmtDue(n.dueAt) + '）' : '');
  });
  const events = (getCalendarDigest() || []).slice(0, 40).map(function (ev) {
    return '- ' + (ev.allDay ? tpeFormat(new Date(ev.start), 'M/d') + ' 全天' : tpeDateTimeStr(new Date(ev.start))) + ' ' + ev.title;
  });
  const prompt = '你是 Ybot，使用者的個人助理。現在是台北時間 ' + tpeFormat(now, 'yyyy/M/d (E) HH:mm') + '。\n' +
    '只根據下面的資料回答，資料裡沒有就說不知道，不要編造。用繁體中文、簡短（200 字內）、不要用 Markdown。\n\n' +
    '【未完成待辦與提醒】\n' + (items.join('\n') || '（無）') + '\n\n' +
    '【未來 14 天行事曆】\n' + (events.join('\n') || '（無）') + '\n\n' +
    '【校長助理今明記事】\n' + (paDayLines().join('\n') || '（無）') + '\n\n' +
    '【台股持股（只有百分比）】\n' + (function () { const s = getStockDetail(); return s ? stockReportText(s) : '（無）'; })() + '\n\n' +
    '【健康】\n' + (function () { const h = healthLatest(); return h ? healthOneLine(h) : '（無）'; })() + '\n\n' +
    '【問題】' + question;
  const ans = callAI(prompt);
  return ans ? String(ans).slice(0, 1500) : 'AI 沒有回應（可能是金鑰或模型失效），請稍後再試。';
}
function handleTelegramText(msg, chatId) {
  if (!msg || !msg.chat || String(msg.chat.id) !== String(chatId)) return;   // 只理自己
  if (!msg.text) return;
  if (msg.date && Date.now() - msg.date * 1000 > TG_MSG_MAX_AGE_MS) return;    // 舊訊息不處理
  const text = String(msg.text).trim();
  if (!text) return;
  const lower = text.toLowerCase();
  let m;
  if (/^(\/help|\/start|說明|幫助|指令)$/.test(lower)) { telegramSend(TG_HELP); return; }
  if (/^(\/list|待辦|清單)$/.test(lower)) {
    const v = tgListView();
    if (v.keyboard.length) telegramSendButtons(v.text, v.keyboard); else telegramSend(v.text);
    return;
  }
  if (/^(\/now|\/brief|今天|現況|簡報)$/.test(lower)) { telegramSend(tgNowText()); return; }
  if (/^(\/status|狀態)$/.test(lower)) { telegramSend(tgStatusText()); return; }
  if (/^(\/stock|股票|台股|持股)$/.test(lower)) { telegramSend(stockReportText(getStockDetail(true))); return; }
  if (/^(\/health|健康)$/.test(lower)) { telegramSend(healthReportText()); return; }
  if (/^(\/pa|校長|校長助理|代課)$/.test(lower)) { telegramSend(paReportText()); return; }
  if (/^(\/agent|agent)\s+(on|off|開|關)$/i.test(lower)) {
    const on = /on|開/.test(lower);
    writeKv({ tgAgent: on ? 'on' : 'off' });
    telegramSend(on ? '🤖 Agent 模式開啟：直接打字就會由 AI 判斷要查什麼、做什麼。' : '🤖 Agent 模式關閉：直接打字＝加待辦，「問 …」才會問 AI。');
    return;
  }
  if (/^(\/reset|重來|清除對話)$/.test(lower)) { agentHistClear(chatId); telegramSend('🧹 對話記憶已清除。'); return; }
  m = text.match(/^(?:\/todo|待辦)[\s:：]+([\s\S]+)$/i);
  if (m) {
    const tid = tgAddNote('todo', m[1].trim());
    telegramSendButtons('✅ 已加入待辦：' + m[1].trim().slice(0, 100), [[{ text: '↩️ 取消', callback_data: 'del:' + tid }]]);
    return;
  }
  m = text.match(/^(?:\/ask|問)[\s:：]+([\s\S]+)$/i);
  if (m) { if (agentEnabledForTelegram()) tgAgentReply(m[1].trim(), chatId); else telegramSend(tgAskAi(m[1].trim())); return; }
  m = text.match(/^(?:\/note|筆記)[\s:：]+([\s\S]+)$/i);
  if (m) { tgAddNote('note', m[1].trim()); telegramSend('📌 已存成筆記：' + m[1].trim().slice(0, 100)); return; }
  m = text.match(/^(\/remind|提醒)([\s:：]*)([\s\S]+)$/i);
  // 「提醒老師交成績」這種沒有空格、也拆不出時間的句子不是指令，往下當一般待辦
  if (m && !m[2] && m[1] === '提醒' && !parseTgReminder(m[3], new Date())) m = null;
  if (m) {
    m = [m[0], m[3]];
    const r = parseTgReminder(m[1], new Date());
    if (!r || !r.content) {
      telegramSend('看不懂提醒時間。請這樣寫：\n提醒 明天 9:00 開會\n提醒 14:30 打電話\n提醒 30分鐘後 收衣服\n提醒 11/20 9:00 繳費');
      return;
    }
    const rid = tgAddNote('reminder', r.content, r.dueAt.toISOString());
    telegramSendButtons('⏰ 已設定提醒：' + r.content.slice(0, 100) + '\n時間：' + tpeDateTimeStr(r.dueAt),
      [[{ text: '↩️ 取消', callback_data: 'del:' + rid }]]);
    return;
  }
  if (text.charAt(0) === '/') { telegramSend('不認得這個指令。\n\n' + TG_HELP); return; }
  // Agent 模式：其餘文字交給 AI 判斷（查東西、做事、或加成待辦）。失敗就退回原本的「加成待辦」。
  if (agentEnabledForTelegram()) {
    try { if (tgAgentReply(text, chatId)) return; } catch (err) { reportError('Telegram agent', err); }
  }
  const id = tgAddNote('todo', text);
  telegramSendButtons('✅ 已加入待辦：' + text.slice(0, 100), [[{ text: '↩️ 取消', callback_data: 'del:' + id }]]);
}
// ld：待辦清單上按 ✅ → 標記完成並把清單重畫；del：剛新增的那筆按「↩️ 取消」→ 刪除
function handleListCallback(cq, chatId, kind, id) {
  const msg = cq.message;
  const { note } = setupSheets();
  const n = sheetToObjects(note, NOTE_COLS).find(function (x) { return String(x.id) === String(id); });
  let reply;
  if (!n) reply = '找不到這筆（可能已刪除）';
  else if (kind === 'ld') { updatePartial(note, NOTE_COLS, n.id, { done: 'true' }); reply = '✅ 已完成'; }
  else { deleteRow(note, n.id); reply = '↩️ 已取消'; }
  try { tgApi('answerCallbackQuery', { callback_query_id: cq.id, text: reply }); } catch (e) { /* 忽略 */ }
  try {
    if (kind === 'ld') {
      const v = tgListView();
      const payload = { chat_id: chatId, message_id: msg.message_id, text: v.text };
      if (v.keyboard.length) payload.reply_markup = { inline_keyboard: v.keyboard };
      tgApi('editMessageText', payload);
    } else {
      tgApi('editMessageText', { chat_id: chatId, message_id: msg.message_id, text: String(msg.text || '').slice(0, 3900) + '\n\n' + reply });
    }
  } catch (e) { /* 訊息沒變或已過期就略過 */ }
}


// ══ 三合一：台股／健康／校長助理 的自動化 ═══════════════════════════
// 原則：能自動的就自動，但每一種通知一天最多一次，不洗版。
// 隱私：台股只輸出百分比與公開股價，不輸出股數、成本、金額；
//       校長助理只收「今明手動記事」與「代課節數」，不收老師姓名（那些留在手機裡）。

// ── 台股明細 ──
// 報價連名稱、昨收、成交日一起拿（getStockPrices 只拿價格，留著給舊的晨報用）。
function getStockQuotes(codes) {
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
      if (m.c && price > 0) out[m.c] = { price: price, prev: y > 0 ? y : price, name: String(m.n || ''), traded: z > 0, day: String(m.d || '') };
    });
  }
  return out;
}
// 純計算：持股 + 報價 → 彙總與每檔今日漲跌%。同一檔分多筆買的會合併。
function stockDetailCompute(hs, quotes, todayYmd) {
  let cost = 0, value = 0, prevValue = 0;
  const byCode = {};
  hs.forEach(h => {
    const code = String(h.code).trim(), x = quotes[code];
    if (!x) return;
    const q = stockToNum(h.qty), c = stockToNum(h.cost);
    if (!(q > 0) || !(c > 0)) return;
    cost += q * c; value += q * x.price; prevValue += q * x.prev;
    // 成交日不是今天（假日、盤前）就不算「今天有交易」，不發漲跌通知
    const live = x.traded && (!x.day || !todayYmd || x.day === todayYmd);
    byCode[code] = { code: code, name: x.name, price: x.price, dayPct: x.prev > 0 ? (x.price - x.prev) / x.prev * 100 : 0, live: live };
  });
  const items = Object.keys(byCode).map(k => byCode[k]).sort((a, b) => Math.abs(b.dayPct) - Math.abs(a.dayPct));
  if (!items.length) return null;
  return {
    n: items.length,
    plPct: cost > 0 ? (value - cost) / cost * 100 : 0,
    dayPct: prevValue > 0 ? (value - prevValue) / prevValue * 100 : 0,
    live: items.some(i => i.live),
    items: items
  };
}
function getStockDetail(noCache) {
  try {
    const cache = CacheService.getScriptCache();
    if (!noCache) { const hit = cache.get('stockDetail_v1'); if (hit) return JSON.parse(hit); }
    const hs = getStockHoldings();
    if (!hs.length) return null;
    const now = new Date();
    const out = stockDetailCompute(hs, getStockQuotes(hs.map(h => String(h.code).trim())), tpeFormat(now, 'yyyyMMdd'));
    if (!out) return null;
    out.at = now.toISOString();
    cache.put('stockDetail_v1', JSON.stringify(out), 300);
    return out;
  } catch (err) { reportError('台股明細', err, 24); return null; }
}
function pctStr(v) { return (v >= 0 ? '+' : '') + Number(v).toFixed(1) + '%'; }
function stockOneLine(s) {
  return '📈 台股 ' + s.n + ' 檔：' + (s.live ? '今日 ' + pctStr(s.dayPct) + '，' : '') + '總未實現損益 ' + pctStr(s.plPct);
}
function stockReportText(s) {
  if (!s) return '📈 還讀不到台股持股。請先在 Apps Script 執行 setStockSheet(\'試算表ID\')（或 setStockConfig），再執行 testStockSummary 檢查。';
  const L = [stockOneLine(s) + '　' + tpeDateTimeStr(new Date(s.at || Date.now()))];
  if (!s.live) L.push('（今天還沒有成交價，漲跌以昨收計）');
  const up = s.items.filter(i => i.dayPct > 0).sort((a, b) => b.dayPct - a.dayPct).slice(0, 5);
  const down = s.items.filter(i => i.dayPct < 0).sort((a, b) => a.dayPct - b.dayPct).slice(0, 5);
  const nm = i => (i.name || i.code) + ' ' + pctStr(i.dayPct);
  if (up.length) L.push('', '🔺 漲：' + up.map(nm).join('、'));
  if (down.length) L.push('🔻 跌：' + down.map(nm).join('、'));
  return L.join('\n');
}

// ── 健康 ──
// 健康系統的檢驗數據是「一列一個項目」，這裡每個項目只留最新一筆，格式 { hba1c:{v:6.1,unit:'%',date:'2026-07-20'} }
function latestLabsByMetric(rows) {
  const best = {};
  rows.forEach(function (r) {
    if (!r || !r.metric) return;
    const v = Number(r.value);
    if (r.value === '' || r.value === null || isNaN(v)) return;
    const date = String(r.measuredAt || r.importedAt || '').slice(0, 10);
    const cur = best[r.metric];
    if (!cur || date >= cur.date) best[r.metric] = { v: v, unit: r.unit || '', date: date };
  });
  return best;
}
const HEALTH_LAB_NAMES = {
  hba1c: '糖化血色素', glucose: '空腹血糖', ldl: 'LDL', hdl: 'HDL', chol: '總膽固醇', tg: '三酸甘油酯',
  sbp: '收縮壓', dbp: '舒張壓', weight: '體重', waist: '腰圍'
};
function healthLatest() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('HEALTH_LATEST');
    if (raw) return JSON.parse(raw);
  } catch (e) { /* 往下重新抓 */ }
  return getLinkedSummaries().health || null;
}
function healthAgeDays(h) {
  const t = new Date(h && h.date).getTime();
  return isNaN(t) ? null : Math.floor((Date.now() - t) / 86400000);
}
function healthOneLine(h) {
  const age = healthAgeDays(h);
  return '🫀 健康分 ' + h.score + '、生理年齡 ' + h.bio + (age === null ? '' : '（' + (age <= 0 ? '今天' : age + ' 天前') + '記錄）');
}
function healthReportText() {
  const h = getLinkedSummaries().health || healthLatest();
  if (!h) return '🫀 還讀不到健康資料。請在 Ybot「設定 → 整合其他系統」填入 LifeSpan 健康的後台網址。';
  const L = [healthOneLine(h)];
  if (h.cvdLevel !== undefined && h.cvdLevel !== '') L.push('心血管風險：' + h.cvdLevel);
  if (h.dmLevel !== undefined && h.dmLevel !== '') L.push('糖尿病風險：' + h.dmLevel);
  if (h.labs) {
    Object.keys(HEALTH_LAB_NAMES).forEach(function (k) {
      const x = h.labs[k];
      if (x) L.push(HEALTH_LAB_NAMES[k] + '：' + x.v + (x.unit ? ' ' + x.unit : '') + (x.date ? '（' + x.date + '）' : ''));
    });
  }
  const age = healthAgeDays(h);
  if (age !== null && age >= 7) L.push('', '⏰ 已經 ' + age + ' 天沒有更新健康數據了。');
  return L.join('\n');
}

// ── 校長助理 ──
// 校長助理的資料只存在手機瀏覽器裡，後台看不到。所以 Ybot App 每次打開（或資料有變）
// 會把「未來 14 天的手動記事」和「每天代課節數」推上來，晨報和 Telegram 才讀得到。
function savePaDays(days) {
  const clean = {};
  Object.keys(days || {}).slice(0, 40).forEach(function (k) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(k)) return;
    const d = days[k] || {};
    clean[k] = {
      items: (Array.isArray(d.items) ? d.items : []).slice(0, 30).map(function (s) { return String(s).slice(0, 80); }),
      subs: Number(d.subs) || 0
    };
  });
  writeKv({ paDaysJson: JSON.stringify(clean), paDaysAt: new Date().toISOString() });
  return '已同步校長助理 ' + Object.keys(clean).length + ' 天';
}
function getPaDay() {
  try {
    const kv = readKv();
    const days = JSON.parse(kv.paDaysJson || '{}');
    const now = new Date();
    const t = tpeFormat(now, 'yyyy-MM-dd'), m = tpeFormat(new Date(now.getTime() + 86400000), 'yyyy-MM-dd');
    return { today: days[t] || null, tomorrow: days[m] || null, syncedAt: kv.paDaysAt || '' };
  } catch (e) { return { today: null, tomorrow: null, syncedAt: '' }; }
}
function paDayLines() {
  const p = getPaDay(), L = [];
  [['今天', p.today], ['明天', p.tomorrow]].forEach(function (x) {
    const d = x[1];
    if (!d) return;
    (d.items || []).forEach(function (s) { L.push(x[0] + ' ' + s); });
    if (d.subs > 0) L.push(x[0] + '代課 ' + d.subs + ' 節');
  });
  return L;
}
function paReportText() {
  const L = paDayLines(), p = getPaDay();
  if (!p.syncedAt) return '🏫 還沒有校長助理的資料。打開 Ybot App 一次，它會自動把今明記事同步上來。';
  return '🏫 校長助理' + '\n' + (L.length ? L.join('\n') : '今明沒有手動記事，也沒有代課。') +
    '\n\n（資料同步於 ' + tpeDateTimeStr(new Date(p.syncedAt)) + '；代課的老師姓名只留在手機裡）';
}

// ── 每 15 分鐘的自動巡邏 ──
// 台股：盤中任一檔漲跌超過門檻（預設 5%，指令碼屬性 STOCK_ALERT_PCT 可改）通知一次；收盤後發一次收盤小結。
// 健康：一天看一次，健康分掉 5 分以上、風險等級改變、或超過 7 天沒更新就提醒（7 天最多提醒一次）。
// 校長助理：早上 7:30 後，今天有記事或代課就推一次「今天學校的事」。
function autoWatchImpl() {
  const cfg = tgConfig();
  if (!cfg.token || !cfg.chatId) return;
  const p = PropertiesService.getScriptProperties();
  const now = new Date();
  const ymd = tpeFormat(now, 'yyyy-MM-dd');
  let st = {};
  try { st = JSON.parse(p.getProperty('AUTO_WATCH_STATE') || '{}'); } catch (e) { st = {}; }
  if (st.date !== ymd) st = { date: ymd, stockCodes: [], stockClose: false, health: false, pa: false, healthNudge: st.healthNudge || '' };
  const hm = Number(tpeFormat(now, 'H')) * 60 + Number(tpeFormat(now, 'm'));
  // 台北的星期幾：0＝週日 … 6＝週六（用台北日期的正午換算，不受專案時區影響）
  const dow = new Date(ymd + 'T12:00:00Z').getUTCDay();
  const out = [];

  if (dow >= 1 && dow <= 5 && hm >= 9 * 60 + 5 && hm <= 14 * 60 + 30 && !st.stockClose) {
    const s = getStockDetail(true);
    if (s && s.live) {
      const th = Number(p.getProperty('STOCK_ALERT_PCT')) || 5;
      if (hm <= 13 * 60 + 30) {
        const hits = s.items.filter(function (i) { return i.live && Math.abs(i.dayPct) >= th && st.stockCodes.indexOf(i.code) < 0; });
        if (hits.length) {
          hits.forEach(function (i) { st.stockCodes.push(i.code); });
          out.push('📈 持股大幅波動（±' + th + '% 以上）\n' + hits.map(function (i) { return (i.dayPct > 0 ? '🔺 ' : '🔻 ') + (i.name || i.code) + ' ' + pctStr(i.dayPct); }).join('\n'));
        }
      } else if (hm >= 13 * 60 + 35) {
        st.stockClose = true;
        out.push('🔔 收盤小結\n' + stockReportText(s));
      }
    }
  }

  if (!st.health && hm >= 9 * 60) {
    st.health = true;
    const h = getLinkedSummaries().health;
    if (h) {
      let prev = null;
      try { prev = JSON.parse(p.getProperty('HEALTH_SEEN') || 'null'); } catch (e) { prev = null; }
      const msgs = [];
      if (prev && String(prev.date) !== String(h.date)) {
        const diff = Number(h.score) - Number(prev.score);
        if (isFinite(diff) && Math.abs(diff) >= 5) msgs.push('健康分 ' + prev.score + ' → ' + h.score + '（' + (diff > 0 ? '+' : '') + diff + '）');
        if (String(prev.cvdLevel) !== String(h.cvdLevel)) msgs.push('心血管風險：' + prev.cvdLevel + ' → ' + h.cvdLevel);
        if (String(prev.dmLevel) !== String(h.dmLevel)) msgs.push('糖尿病風險：' + prev.dmLevel + ' → ' + h.dmLevel);
      }
      p.setProperty('HEALTH_SEEN', JSON.stringify(h));
      const age = healthAgeDays(h);
      const nudgeAge = st.healthNudge ? Math.floor((now.getTime() - new Date(st.healthNudge).getTime()) / 86400000) : 99;
      if (age !== null && age >= 7 && nudgeAge >= 7) { msgs.push('已經 ' + age + ' 天沒有更新健康數據，找時間量一下。'); st.healthNudge = now.toISOString(); }
      if (msgs.length) out.push('🫀 健康提醒\n' + msgs.join('\n'));
    }
  }

  if (!st.pa && hm >= 7 * 60 + 30 && hm <= 12 * 60) {
    st.pa = true;
    const d = getPaDay().today;
    if (d && ((d.items || []).length || d.subs > 0)) {
      out.push('🏫 今天學校的事\n' + (d.items || []).map(function (s) { return '・' + s; }).join('\n') + (d.subs > 0 ? '\n・代課 ' + d.subs + ' 節' : ''));
    }
  }

  // 先存狀態再送：就算送到一半失敗，也不會下一輪又重送一次
  p.setProperty('AUTO_WATCH_STATE', JSON.stringify(st));
  out.forEach(function (t) { telegramSend(t); });
}
// 手動測試：在編輯器執行，會把三個系統目前讀到的內容各傳一則到 Telegram。
function testThreeAgents() {
  telegramSend(stockReportText(getStockDetail(true)));
  telegramSend(healthReportText());
  telegramSend(paReportText());
  console.log('✅ 已傳出台股／健康／校長助理三則訊息，請看 Telegram');
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
function autoWatch() { return guard('autoWatch', autoWatchImpl); }

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

// ═══════════════════════════════════════════════════════════════════
// 🤖 Agent（工具迴圈）
// ─────────────────────────────────────────────────────────────────
// 以前的「問」是一次問答：把待辦、行程、台股一行摘要塞進 prompt，AI 只能照著回。
// 這裡改成真正的 agent：AI 拿到一組工具（查台股、查檢驗、查行事曆、搜郵件、
// 加待辦……），自己決定要呼叫哪個、拿到結果再決定下一步，做完才回話。
// Telegram 的自由文字和 App 的對話都走這條路。
//
// 安全邊界（刻意的，不是漏做）：
// - Gmail 只能搜尋與讀取，不能寄信、不能刪信。
// - 行事曆可以查、加、改，不能刪。
// - 待辦可以查、加、改、標完成，不能刪（App 裡自己刪）。
// - 每次最多 AGENT_MAX_STEPS 步、AGENT_TIME_BUDGET_MS 毫秒，超過就停下來回報做到哪。
// ═══════════════════════════════════════════════════════════════════
const AGENT_MAX_STEPS = 8;
const AGENT_TIME_BUDGET_MS = 110000;   // Apps Script 單次上限 6 分鐘，Telegram 輪詢又每分鐘跑一次，留寬裕
const AGENT_HISTORY_MAX = 10;
const AGENT_HIST_PROP = 'agentHist_';   // CacheService key 前綴（Telegram 對話記憶）

const HEALTH_METRIC_NAMES = {
  hba1c: '糖化血色素', glucose: '空腹血糖', ldl: 'LDL 低密度膽固醇', hdl: 'HDL 高密度膽固醇',
  chol: '總膽固醇', tg: '三酸甘油酯', sbp: '收縮壓', dbp: '舒張壓', weight: '體重', waist: '腰圍',
  height: '身高', hr: '心率', bmi: 'BMI', egfr: 'eGFR 腎絲球過濾率', cr: '肌酸酐', ua: '尿酸',
  ast: 'AST(GOT)', alt: 'ALT(GPT)', ggt: 'γ-GT', hb: '血紅素', tsh: 'TSH 甲狀腺刺激素', k: '鉀', na: '鈉'
};

// ── 工具定義 ──
// 每個工具：name / description（給 AI 看的，要寫清楚什麼時候用）/ parameters（JSON Schema）/ run(args)。
// run 回傳的物件會原封不動 JSON 給 AI，所以只放 AI 需要的欄位，不放網址與密鑰。
const AGENT_TOOLS = [
  {
    name: 'get_today_overview',
    description: '取得今明行程、到期待辦、校長助理記事、台股一行摘要、健康一行摘要。使用者問「今天」「現在狀況」「有什麼要做」時先用這個。',
    parameters: { type: 'object', properties: {} },
    run: function () { return { overview: tgNowText() }; }
  },
  {
    name: 'list_notes',
    description: '列出待辦／提醒／筆記／心法。可用 status 篩未完成或已完成、type 篩類型、query 做關鍵字比對。回傳每筆的 id（要修改或標完成時用）。',
    parameters: {
      type: 'object',
      properties: {
        status: { type: 'string', enum: ['open', 'done', 'all'], description: '預設 open（未完成）' },
        type: { type: 'string', enum: ['todo', 'reminder', 'note', 'idea'], description: '不給就全部類型' },
        query: { type: 'string', description: '內容關鍵字（可省略）' }
      }
    },
    run: function (a) {
      a = a || {};
      let rows = sheetToObjects(setupSheets().note, NOTE_COLS);
      const st = a.status || 'open';
      if (st === 'open') rows = rows.filter(n => n.done !== 'true');
      if (st === 'done') rows = rows.filter(n => n.done === 'true');
      if (a.type) rows = rows.filter(n => n.type === a.type);
      if (a.query) { const q = String(a.query).toLowerCase(); rows = rows.filter(n => String(n.content).toLowerCase().indexOf(q) >= 0); }
      rows.sort((x, y) => String(y.createdAt).localeCompare(String(x.createdAt)));
      return {
        count: rows.length,
        items: rows.slice(0, 60).map(n => ({
          id: n.id, type: n.type, content: n.content, dueAt: n.dueAt ? tpeFormat(new Date(n.dueAt), 'yyyy/M/d HH:mm') : '',
          done: n.done === 'true', owner: n.owner || '', createdAt: tpeFormat(new Date(n.createdAt), 'M/d')
        }))
      };
    }
  },
  {
    name: 'add_note',
    description: '新增一筆。type：todo=待辦、reminder=特定時間要提醒（必須給 dueAt）、note=瑣事筆記、idea=想法與心法。owner 是「交辦給誰」，使用者自己要做的不要填。',
    parameters: {
      type: 'object',
      properties: {
        type: { type: 'string', enum: ['todo', 'reminder', 'note', 'idea'] },
        content: { type: 'string' },
        dueAt: { type: 'string', description: '到期或提醒時間，台北時間，格式 2026-11-20 09:00（可省略）' },
        owner: { type: 'string', description: '交辦對象（可省略）' }
      },
      required: ['type', 'content']
    },
    run: function (a) {
      const due = agentParseTime(a.dueAt);
      if (a.type === 'reminder' && !due) return { error: '提醒一定要有時間，請給 dueAt' };
      const id = Utilities.getUuid();
      appendObj(setupSheets().note, NOTE_COLS, {
        id: id, type: a.type || 'todo', content: String(a.content || '').trim(), dueAt: due ? due.toISOString() : '',
        done: '', createdAt: new Date().toISOString(), notifiedAt: '', owner: a.owner || ''
      });
      return { ok: true, id: id, dueAt: due ? tpeFormat(due, 'M/d HH:mm') : '' };
    }
  },
  {
    name: 'update_note',
    description: '修改既有的待辦／提醒／筆記。只放要改的欄位。改時間給 dueAt、改交辦對象給 owner（收回自己做給空字串）。',
    parameters: {
      type: 'object',
      properties: {
        id: { type: 'string' }, content: { type: 'string' }, dueAt: { type: 'string', description: '台北時間，格式 2026-11-20 09:00' },
        type: { type: 'string', enum: ['todo', 'reminder', 'note', 'idea'] }, owner: { type: 'string' }
      },
      required: ['id']
    },
    run: function (a) {
      const patch = {};
      if (a.content !== undefined) patch.content = String(a.content);
      if (a.type !== undefined) patch.type = a.type;
      if (a.owner !== undefined) patch.owner = a.owner;
      if (a.dueAt !== undefined) {
        const d = agentParseTime(a.dueAt);
        if (a.dueAt && !d) return { error: '看不懂時間：' + a.dueAt };
        patch.dueAt = d ? d.toISOString() : '';
        patch.notifiedAt = '';
      }
      if (!Object.keys(patch).length) return { error: '沒有要改的欄位' };
      const { note } = setupSheets();
      if (!sheetToObjects(note, NOTE_COLS).some(n => String(n.id) === String(a.id))) return { error: '找不到這筆 id' };
      updatePartial(note, NOTE_COLS, a.id, patch);
      return { ok: true };
    }
  },
  {
    name: 'complete_note',
    description: '把待辦或提醒標記為完成。',
    parameters: { type: 'object', properties: { id: { type: 'string' } }, required: ['id'] },
    run: function (a) {
      const { note } = setupSheets();
      const n = sheetToObjects(note, NOTE_COLS).find(x => String(x.id) === String(a.id));
      if (!n) return { error: '找不到這筆 id' };
      updatePartial(note, NOTE_COLS, a.id, { done: 'true' });
      return { ok: true, content: n.content };
    }
  },
  {
    name: 'list_calendar',
    description: '列出 Google 日曆行程（所有看得到的日曆）。預設今天起 14 天；可給 from / to 查更遠或更久以前。回傳 eventId（要修改時用）。',
    parameters: {
      type: 'object',
      properties: {
        from: { type: 'string', description: '起日，台北時間，格式 2026-11-20（可省略＝今天）' },
        to: { type: 'string', description: '迄日，格式 2026-12-04（可省略＝起日 +14 天）' }
      }
    },
    run: function (a) {
      a = a || {};
      const from = agentParseTime(a.from) || new Date();
      const to = agentParseTime(a.to) || new Date(from.getTime() + 14 * 86400000);
      if (a.to && !a.to.match(/\d{1,2}:\d{2}/)) to.setHours(23, 59, 59, 0);
      const seen = {}; const out = [];
      CalendarApp.getAllCalendars().forEach(cal => {
        try {
          cal.getEvents(from, to).forEach(ev => {
            const key = ev.getId() + '_' + ev.getStartTime().getTime();
            if (seen[key]) return; seen[key] = 1;
            out.push({
              eventId: ev.getId(), title: ev.getTitle(), calendar: cal.getName(), allDay: ev.isAllDayEvent(),
              start: tpeFormat(ev.getStartTime(), 'yyyy/M/d (E) HH:mm'), end: tpeFormat(ev.getEndTime(), 'yyyy/M/d HH:mm'),
              location: ev.getLocation() || ''
            });
          });
        } catch (e) { /* 單一日曆失敗跳過 */ }
      });
      out.sort((x, y) => x.start.localeCompare(y.start));
      return { range: tpeFormat(from, 'M/d') + '～' + tpeFormat(to, 'M/d'), count: out.length, events: out.slice(0, 80) };
    }
  },
  {
    name: 'add_calendar_event',
    description: '在主要 Google 日曆新增一筆行程。沒給 end 就是 1 小時。',
    parameters: {
      type: 'object',
      properties: {
        title: { type: 'string' }, start: { type: 'string', description: '台北時間，格式 2026-11-20 14:00' },
        end: { type: 'string', description: '格式 2026-11-20 15:00（可省略）' },
        location: { type: 'string' }, description: { type: 'string' }
      },
      required: ['title', 'start']
    },
    run: function (a) {
      const s = agentParseTime(a.start);
      if (!s) return { error: '看不懂開始時間：' + a.start };
      const e = agentParseTime(a.end) || new Date(s.getTime() + 3600000);
      const ev = CalendarApp.getDefaultCalendar().createEvent(String(a.title || '(未命名)'), s, e,
        { location: a.location || '', description: a.description || '' });
      ev.addPopupReminder(30);
      return { ok: true, eventId: ev.getId(), start: tpeFormat(s, 'M/d (E) HH:mm'), end: tpeFormat(e, 'HH:mm') };
    }
  },
  {
    name: 'update_calendar_event',
    description: '修改既有行程的標題、時間或地點。eventId 從 list_calendar 取得。只放要改的欄位。',
    parameters: {
      type: 'object',
      properties: {
        eventId: { type: 'string' }, title: { type: 'string' },
        start: { type: 'string', description: '台北時間，格式 2026-11-20 14:00' }, end: { type: 'string' }, location: { type: 'string' }
      },
      required: ['eventId']
    },
    run: function (a) {
      const ev = agentFindEvent(a.eventId);
      if (!ev) return { error: '找不到這筆行程' };
      if (a.title) ev.setTitle(String(a.title));
      if (a.location !== undefined) ev.setLocation(String(a.location));
      if (a.start || a.end) {
        const s = agentParseTime(a.start) || ev.getStartTime();
        const e = agentParseTime(a.end) || new Date(s.getTime() + (ev.getEndTime().getTime() - ev.getStartTime().getTime()));
        if (a.start && !agentParseTime(a.start)) return { error: '看不懂開始時間：' + a.start };
        ev.setTime(s, e);
      }
      return { ok: true, title: ev.getTitle(), start: tpeFormat(ev.getStartTime(), 'M/d (E) HH:mm'), end: tpeFormat(ev.getEndTime(), 'HH:mm') };
    }
  },
  {
    name: 'search_gmail',
    description: '用 Gmail 搜尋語法找信（例如 "from:某人 newer_than:7d"、"帳單 is:unread"、"subject:公文"）。只回主旨、寄件者、日期與摘要片段；要看內文再用 read_gmail_thread。只能讀，不能寄、不能刪。',
    parameters: {
      type: 'object',
      properties: { query: { type: 'string' }, limit: { type: 'integer', description: '最多幾封，預設 8，上限 15' } },
      required: ['query']
    },
    run: function (a) {
      const lim = Math.min(Math.max(Number(a.limit) || 8, 1), 15);
      const threads = GmailApp.search(String(a.query || ''), 0, lim);
      return {
        count: threads.length,
        threads: threads.map(t => {
          const msgs = t.getMessages(); const last = msgs[msgs.length - 1];
          return {
            threadId: t.getId(), subject: t.getFirstMessageSubject() || '(無主旨)', from: last.getFrom(),
            date: tpeFormat(last.getDate(), 'M/d HH:mm'), unread: t.isUnread(), messages: msgs.length,
            snippet: (last.getPlainBody() || '').replace(/\s+/g, ' ').slice(0, 200)
          };
        })
      };
    }
  },
  {
    name: 'read_gmail_thread',
    description: '讀一封信的內文（最後一則訊息，最多 3000 字）。threadId 從 search_gmail 取得。',
    parameters: { type: 'object', properties: { threadId: { type: 'string' } }, required: ['threadId'] },
    run: function (a) {
      const t = GmailApp.getThreadById(String(a.threadId || ''));
      if (!t) return { error: '找不到這封信' };
      const msgs = t.getMessages(); const last = msgs[msgs.length - 1];
      return {
        subject: t.getFirstMessageSubject(), from: last.getFrom(), to: last.getTo(), date: tpeFormat(last.getDate(), 'yyyy/M/d HH:mm'),
        body: (last.getPlainBody() || '').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').slice(0, 3000),
        attachments: last.getAttachments().map(x => x.getName())
      };
    }
  },
  {
    name: 'get_stock_holdings',
    description: '取得台股全部持股明細：每檔的股數、平均成本、現價、今日漲跌%、未實現損益金額與%，以及總成本、總市值、總損益。問「哪一支賺最多」「總共賺多少」「某支現在如何」都用這個。',
    parameters: { type: 'object', properties: {} },
    run: function () {
      const hs = getStockHoldings();
      if (!hs.length) return { error: '讀不到持股。請先在 Apps Script 執行 setStockSheet 或 setStockConfig。' };
      const quotes = getStockQuotes(hs.map(h => String(h.code).trim()));
      const merged = {};
      hs.forEach(h => {
        const code = String(h.code).trim(), q = stockToNum(h.qty), c = stockToNum(h.cost);
        if (!(q > 0) || !(c > 0)) return;
        if (!merged[code]) merged[code] = { code: code, qty: 0, cost: 0 };
        merged[code].qty += q; merged[code].cost += q * c;
      });
      let tCost = 0, tValue = 0, tPrev = 0;
      const items = Object.keys(merged).map(code => {
        const m = merged[code], x = quotes[code];
        const avg = m.cost / m.qty;
        if (!x) return { code: code, qty: m.qty, avgCost: round2(avg), note: '查不到報價' };
        const value = m.qty * x.price, pl = value - m.cost;
        tCost += m.cost; tValue += value; tPrev += m.qty * x.prev;
        return {
          code: code, name: x.name, qty: m.qty, avgCost: round2(avg), price: x.price,
          dayPct: round2(x.prev > 0 ? (x.price - x.prev) / x.prev * 100 : 0),
          value: Math.round(value), plAmount: Math.round(pl), plPct: round2(m.cost > 0 ? pl / m.cost * 100 : 0)
        };
      }).sort((a, b) => (b.plAmount || 0) - (a.plAmount || 0));
      return {
        asOf: tpeDateTimeStr(new Date()), count: items.length,
        totalCost: Math.round(tCost), totalValue: Math.round(tValue), totalPlAmount: Math.round(tValue - tCost),
        totalPlPct: round2(tCost > 0 ? (tValue - tCost) / tCost * 100 : 0),
        dayPct: round2(tPrev > 0 ? (tValue - tPrev) / tPrev * 100 : 0),
        items: items
      };
    }
  },
  {
    name: 'get_stock_quote',
    description: '查任一檔台股（不一定有持股）的即時報價：代號、名稱、現價、昨收、今日漲跌%。',
    parameters: { type: 'object', properties: { code: { type: 'string', description: '股票代號，例如 2330' } }, required: ['code'] },
    run: function (a) {
      const code = String(a.code || '').trim();
      const x = getStockQuotes([code])[code];
      if (!x) return { error: '查不到 ' + code + ' 的報價（可能代號錯或非交易時段）' };
      return { code: code, name: x.name, price: x.price, prev: x.prev, dayPct: round2(x.prev > 0 ? (x.price - x.prev) / x.prev * 100 : 0), traded: x.traded };
    }
  },
  {
    name: 'get_health_summary',
    description: '健康系統最新快照：健康分、生理年齡、心血管與糖尿病風險，以及各檢驗項目的最新一筆數值。',
    parameters: { type: 'object', properties: {} },
    run: function () {
      const h = getLinkedSummaries().health || healthLatest();
      if (!h) return { error: '讀不到健康資料。請在 Ybot 設定 → 整合其他系統填入健康後台網址。' };
      const labs = {};
      Object.keys(h.labs || {}).forEach(k => { labs[HEALTH_METRIC_NAMES[k] || k] = h.labs[k]; });
      return { date: h.date, score: h.score, bioAge: h.bio, cvdLevel: h.cvdLevel, dmLevel: h.dmLevel, latestLabs: labs };
    }
  },
  {
    name: 'get_health_labs',
    description: '健康系統「檢驗數據庫」的歷史紀錄（糖化血色素、血糖、血脂、血壓、體重…）。可用 metric 只看一項（用英文代碼如 hba1c、glucose、ldl、hdl、tg、chol、sbp、dbp、weight、waist、egfr、ua）。問趨勢、上次數值、是否改善時用這個。',
    parameters: {
      type: 'object',
      properties: { metric: { type: 'string', description: '英文代碼，省略＝全部' }, limit: { type: 'integer', description: '每項最多幾筆，預設 12' } }
    },
    run: function (a) {
      a = a || {};
      const kv = readKv();
      if (!kv.healthBackendUrl) return { error: '還沒設定健康後台網址' };
      const r = UrlFetchApp.fetch(kv.healthBackendUrl + '?action=labs', { muteHttpExceptions: true });
      const d = JSON.parse(r.getContentText());
      if (!d.ok || !Array.isArray(d.labs)) return { error: '健康後台沒有回傳檢驗資料' };
      const lim = Math.min(Math.max(Number(a.limit) || 12, 1), 50);
      const by = {};
      d.labs.forEach(l => {
        if (!l || !l.metric || l.value === '' || isNaN(Number(l.value))) return;
        if (a.metric && String(l.metric) !== String(a.metric)) return;
        const date = String(l.measuredAt || l.importedAt || '').slice(0, 10);
        (by[l.metric] = by[l.metric] || []).push({ date: date, value: Number(l.value), unit: l.unit || '' });
      });
      const out = {};
      Object.keys(by).forEach(k => {
        const seen = {};
        const rows = by[k].sort((x, y) => y.date.localeCompare(x.date)).filter(r => { if (seen[r.date + '_' + r.value]) return false; seen[r.date + '_' + r.value] = 1; return true; }).slice(0, lim);
        out[k] = { name: HEALTH_METRIC_NAMES[k] || k, records: rows };
      });
      if (!Object.keys(out).length) return { error: a.metric ? ('檢驗數據庫裡沒有 ' + a.metric) : '檢驗數據庫是空的（請在健康系統按「上傳雲端」）' };
      return { metrics: out };
    }
  },
  {
    name: 'get_health_records',
    description: '健康系統的醫藥紀錄：手術（kind=surgery）、用藥（med）、就醫診斷／病史（dx）、篩檢（screen）。問「開過什麼刀」「吃什麼藥」「有沒有高血壓診斷」「大腸癌篩檢」用這個。可用 query 做名稱關鍵字比對。',
    parameters: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['surgery', 'med', 'dx', 'screen'], description: '省略＝全部' },
        query: { type: 'string', description: '名稱、診斷或院所的關鍵字（可省略）' },
        limit: { type: 'integer', description: '最多幾筆，預設 30' }
      }
    },
    run: function (a) {
      a = a || {};
      const kv = readKv();
      if (!kv.healthBackendUrl) return { error: '還沒設定健康後台網址' };
      const r = UrlFetchApp.fetch(kv.healthBackendUrl + '?action=records' + (a.kind ? '&kind=' + encodeURIComponent(a.kind) : ''), { muteHttpExceptions: true });
      const d = JSON.parse(r.getContentText());
      if (!d.ok || !Array.isArray(d.records)) return { error: '健康後台沒有回傳醫藥紀錄（可能後台還是舊版，或 App 還沒同步）' };
      const names = { surgery: '手術', med: '用藥', dx: '診斷', screen: '篩檢' };
      let rows = d.records;
      if (a.query) { const q = String(a.query).toLowerCase(); rows = rows.filter(x => (String(x.name) + ' ' + String(x.detail) + ' ' + String(x.hospital)).toLowerCase().indexOf(q) >= 0); }
      rows.sort((x, y) => String(y.date).localeCompare(String(x.date)));
      const lim = Math.min(Math.max(Number(a.limit) || 30, 1), 100);
      if (!rows.length) return { error: '沒有符合的醫藥紀錄' + (d.records.length ? '' : '（醫藥紀錄是空的，請在健康系統按「上傳雲端」）') };
      return { count: rows.length, records: rows.slice(0, lim).map(x => ({ kind: names[x.kind] || x.kind, date: x.date, name: x.name, detail: x.detail, hospital: x.hospital })) };
    }
  },
  {
    name: 'get_school_assistant',
    description: '校長助理：今明的學校記事與代課節數。',
    parameters: { type: 'object', properties: {} },
    run: function () { return { text: paReportText() }; }
  },
  {
    name: 'remember',
    description: '把「以後每次對話都用得到」的一句話記下來（怎麼稱呼某人、偏好、固定做事方式、職務範圍）。不要拿來記一次性的待辦或當下心情。',
    parameters: { type: 'object', properties: { text: { type: 'string' } }, required: ['text'] },
    run: function (a) {
      const t = String(a.text || '').trim();
      if (!t) return { error: '內容是空的' };
      const { qa } = setupSheets();
      const exists = sheetToObjects(qa, QA_COLS).some(r => r.kind === 'profile' && String(r.a).trim() === t);
      if (exists) return { ok: true, note: '已經記過了' };
      const id = 'p' + Date.now() + Math.random().toString(36).slice(2, 7);
      upsertRow(qa, QA_COLS, id, { id: id, q: '__profile__', a: t, kind: 'profile', at: new Date().toISOString() });
      return { ok: true, id: id };
    }
  },
  {
    name: 'forget',
    description: '刪掉一條已經不成立的記憶（換職務、改了習慣）。id 來自「關於使用者」清單。',
    parameters: { type: 'object', properties: { id: { type: 'string' } }, required: ['id'] },
    run: function (a) { deleteRow(setupSheets().qa, String(a.id || '')); return { ok: true }; }
  }
];
function round2(v) { return Math.round(Number(v) * 100) / 100; }
function agentToolByName(name) { return AGENT_TOOLS.find(t => t.name === name) || null; }

// 「2026-11-20 09:00」「2026/11/20」「2026-11-20T09:00」都當台北時間。只給日期就是當天 00:00。
function agentParseTime(s) {
  if (!s) return null;
  s = String(s).trim();
  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (m) {
    const iso = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) + 'T' + ('0' + (m[4] || '0')).slice(-2) + ':' + (m[5] || '00') + ':00+08:00';
    const d = new Date(iso);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function agentFindEvent(id) {
  const cals = [CalendarApp.getDefaultCalendar()].concat(CalendarApp.getAllCalendars());
  for (let i = 0; i < cals.length; i++) {
    try { const ev = cals[i].getEventById(id); if (ev) return ev; } catch (e) { /* 下一個 */ }
  }
  return null;
}
function agentProfileLines() {
  try {
    return sheetToObjects(setupSheets().qa, QA_COLS).filter(r => r.kind === 'profile' && r.a)
      .sort((x, y) => String(y.at).localeCompare(String(x.at))).slice(0, 60)
      .map(r => '[' + r.id + '] ' + String(r.a).trim());
  } catch (e) { return []; }
}

function agentSystemPrompt(channel) {
  const now = new Date();
  const profile = agentProfileLines();
  return [
    '你是 Young 的個人助理 Ybot。Young 是教育工作者、批判教育哲學家，也寫詩與散文。你溫暖、直接、不囉嗦，像一個懂他的朋友。',
    '現在是台北時間 ' + tpeFormat(now, 'yyyy/M/d (E) HH:mm') + '。所有時間都用台北時間。',
    '',
    '你有一組工具。回答之前先想：需要哪些資料？能用工具查到的就去查，不要猜、不要編。',
    '可以連續用好幾個工具（先查持股再查單檔、先查行事曆再加行程）。資料查不到就老實說查不到，並說要去哪裡設定。',
    '使用者要你做事（加待辦、排行程、改時間、標完成、記住某件事），就直接做，做完簡短回報做了什麼。',
    '要改或完成某一筆時，先用 list 工具找到正確的 id 再動手；比對到多筆就先問清楚是哪一筆，不要亂猜。',
    '意思不清楚、或會影響既有資料（改時間、改交辦對象）而你不確定時，先問一句再做。',
    '使用者只丟一句像待辦的話（沒有問句、沒有要查東西），就幫他加成待辦；有明確時間就加成提醒。',
    '提到「明天」「下禮拜三」這類相對時間，請自己換算成確切日期再帶給工具。',
    '',
    '回答規則：繁體中文；簡短，先講結論；' + (channel === 'telegram' ? '純文字，不要用 Markdown 符號（不要 ** 和 #）；' : '可以用簡單的換行與項目符號；') +
    '金額用千分位，百分比保留一位小數；不要把工具回傳的 id 唸給使用者聽。',
    '',
    '關於使用者（之前記下來的，每條前面是記憶 id）：',
    profile.length ? profile.join('\n') : '（還沒有）'
  ].join('\n');
}

// ── 迴圈本體 ──
// history：[{role:'user'|'bot', text}]，最近幾輪。回傳 {text, steps:[{tool,args}], error}
function runAgent(question, history, opts) {
  opts = opts || {};
  const cfg = getAiConfig();
  const order = aiOrder(cfg).filter(p => p === 'gemini' || p === 'openai');   // 工具呼叫只有這兩家支援
  if (!order.length) return { text: '', error: '後台沒有可用的 AI（Gemini 或 OpenAI）金鑰，請在 Apps Script 執行 setAiConfig。' };
  const provider = order[0];
  const started = Date.now();
  const steps = [];
  const sys = agentSystemPrompt(opts.channel || 'app');
  const hist = (history || []).slice(-AGENT_HISTORY_MAX);
  const ctx = provider === 'gemini' ? agentGeminiStart(sys, hist, question) : agentOpenaiStart(sys, hist, question);
  let lastErr = '';
  for (let step = 0; step <= AGENT_MAX_STEPS; step++) {
    if (Date.now() - started > AGENT_TIME_BUDGET_MS) { lastErr = '這題花太久了，先停在這裡'; break; }
    let turn;
    try { turn = provider === 'gemini' ? agentGeminiTurn(ctx, cfg.gemini) : agentOpenaiTurn(ctx, cfg.openai); }
    catch (e) { lastErr = 'AI 呼叫失敗：' + (e && e.message || e); break; }
    if (!turn.calls.length) return { text: String(turn.text || '').trim(), steps: steps };
    if (step === AGENT_MAX_STEPS) { lastErr = '步驟太多，先停在這裡'; break; }
    const results = turn.calls.map(c => {
      const tool = agentToolByName(c.name);
      let res;
      if (!tool) res = { error: '沒有這個工具：' + c.name };
      else { try { res = tool.run(c.args || {}); } catch (e) { res = { error: String(e && e.message || e) }; } }
      steps.push({ tool: c.name, args: c.args || {}, ok: !(res && res.error) });
      return { call: c, result: res };
    });
    if (provider === 'gemini') agentGeminiFeed(ctx, turn, results); else agentOpenaiFeed(ctx, turn, results);
  }
  // 中途停下來：把做過的事告訴使用者，不要默默吞掉
  const done = steps.filter(s => s.ok).map(s => agentToolLabel(s.tool));
  return { text: (lastErr ? '⚠️ ' + lastErr : '') + (done.length ? '\n已完成：' + done.join('、') : ''), steps: steps, error: lastErr };
}
const AGENT_TOOL_LABELS = {
  get_today_overview: '看今日現況', list_notes: '查待辦筆記', add_note: '新增待辦筆記', update_note: '修改待辦筆記', complete_note: '標記完成',
  list_calendar: '查行事曆', add_calendar_event: '新增行程', update_calendar_event: '修改行程', search_gmail: '搜尋郵件', read_gmail_thread: '讀信',
  get_stock_holdings: '查持股', get_stock_quote: '查報價', get_health_summary: '查健康摘要', get_health_labs: '查檢驗紀錄', get_health_records: '查醫藥紀錄',
  get_school_assistant: '查校長助理', remember: '記住', forget: '忘記'
};
function agentToolLabel(name) { return AGENT_TOOL_LABELS[name] || name; }
function agentStepsLine(steps) {
  if (!steps || !steps.length) return '';
  const seen = {}; const names = [];
  steps.forEach(s => { const l = agentToolLabel(s.tool); if (!seen[l]) { seen[l] = 1; names.push(l); } });
  return '🔧 ' + names.join('、');
}

// ── Gemini（function calling）──
function agentGeminiStart(sys, hist, question) {
  const contents = hist.map(h => ({ role: h.role === 'user' ? 'user' : 'model', parts: [{ text: String(h.text || '').slice(0, 2000) }] }));
  contents.push({ role: 'user', parts: [{ text: String(question) }] });
  return {
    sys: sys, contents: contents,
    tools: [{ function_declarations: AGENT_TOOLS.map(t => {
      const decl = { name: t.name, description: t.description };
      // Gemini 不接受空的 properties，所以沒參數的工具就不帶 parameters；型別名稱要大寫
      if (t.parameters && t.parameters.properties && Object.keys(t.parameters.properties).length) decl.parameters = agentGeminiSchema(t.parameters);
      return decl;
    }) }]
  };
}
function agentGeminiSchema(sc) {
  if (Array.isArray(sc)) return sc.map(agentGeminiSchema);
  if (!sc || typeof sc !== 'object') return sc;
  const out = {};
  Object.keys(sc).forEach(k => {
    if (k === 'type' && typeof sc[k] === 'string') out[k] = sc[k].toUpperCase();
    else if (k === 'properties') { out[k] = {}; Object.keys(sc[k]).forEach(pk => out[k][pk] = agentGeminiSchema(sc[k][pk])); }
    else out[k] = agentGeminiSchema(sc[k]);
  });
  return out;
}
function agentGeminiTurn(ctx, key) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + key;
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({
      system_instruction: { parts: [{ text: ctx.sys }] }, contents: ctx.contents, tools: ctx.tools,
      generationConfig: { temperature: 0.4, maxOutputTokens: 2048 }
    })
  });
  const d = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error((d.error && d.error.message) || ('Gemini ' + res.getResponseCode()));
  const parts = (d.candidates && d.candidates[0] && d.candidates[0].content && d.candidates[0].content.parts) || [];
  const calls = parts.filter(p => p.functionCall).map(p => ({ name: p.functionCall.name, args: p.functionCall.args || {} }));
  const text = parts.filter(p => p.text).map(p => p.text).join('');
  return { parts: parts, calls: calls, text: text };
}
function agentGeminiFeed(ctx, turn, results) {
  ctx.contents.push({ role: 'model', parts: turn.parts });
  ctx.contents.push({ role: 'user', parts: results.map(r => ({ functionResponse: { name: r.call.name, response: { result: r.result } } })) });
}

// ── OpenAI（tools）──
function agentOpenaiStart(sys, hist, question) {
  const messages = [{ role: 'system', content: sys }];
  hist.forEach(h => messages.push({ role: h.role === 'user' ? 'user' : 'assistant', content: String(h.text || '').slice(0, 2000) }));
  messages.push({ role: 'user', content: String(question) });
  return { messages: messages, tools: AGENT_TOOLS.map(t => ({ type: 'function', function: { name: t.name, description: t.description, parameters: t.parameters } })) };
}
function agentOpenaiTurn(ctx, key) {
  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({ model: 'gpt-4o-mini', messages: ctx.messages, tools: ctx.tools, temperature: 0.4, max_tokens: 2048 })
  });
  const d = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error((d.error && d.error.message) || ('OpenAI ' + res.getResponseCode()));
  const msg = (d.choices && d.choices[0] && d.choices[0].message) || {};
  const calls = (msg.tool_calls || []).map(tc => {
    let args = {}; try { args = JSON.parse(tc.function.arguments || '{}'); } catch (e) { /* 空參數 */ }
    return { id: tc.id, name: tc.function.name, args: args };
  });
  return { msg: msg, calls: calls, text: msg.content || '' };
}
function agentOpenaiFeed(ctx, turn, results) {
  ctx.messages.push(turn.msg);
  results.forEach(r => ctx.messages.push({ role: 'tool', tool_call_id: r.call.id, content: JSON.stringify(r.result) }));
}

// ── Telegram 的對話記憶（CacheService，2 小時）──
function agentHistLoad(chatId) {
  try { return JSON.parse(CacheService.getScriptCache().get(AGENT_HIST_PROP + chatId) || '[]'); } catch (e) { return []; }
}
function agentHistSave(chatId, hist) {
  try { CacheService.getScriptCache().put(AGENT_HIST_PROP + chatId, JSON.stringify(hist.slice(-AGENT_HISTORY_MAX)), 7200); } catch (e) { /* 只是記憶 */ }
}
function agentHistClear(chatId) { try { CacheService.getScriptCache().remove(AGENT_HIST_PROP + chatId); } catch (e) { /* 忽略 */ } }
// Telegram 自由文字 → agent。回傳 true 表示已處理（已回覆）。
function tgAgentReply(text, chatId) {
  const hist = agentHistLoad(chatId);
  const r = runAgent(text, hist, { channel: 'telegram' });
  if (r.error && !r.text) { telegramSend('⚠️ ' + r.error); return true; }
  const body = String(r.text || '（沒有回覆）').slice(0, 3800);
  const foot = agentStepsLine(r.steps);
  telegramSend(body + (foot ? '\n\n' + foot : ''));
  hist.push({ role: 'user', text: text }, { role: 'bot', text: body });
  agentHistSave(chatId, hist);
  return true;
}
function agentEnabledForTelegram() {
  const kv = readKv();
  if (kv.tgAgent === 'off') return false;
  const cfg = getAiConfig();
  return !!(cfg.gemini || cfg.openai);
}
// 手動測試：在編輯器執行，看 agent 會怎麼回（結果只印在執行記錄，不會送 Telegram）
function testAgent() {
  const r = runAgent('我今天有什麼要做的？', [], { channel: 'telegram' });
  console.log(JSON.stringify(r, null, 2));
}
