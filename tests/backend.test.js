// 後台的安全規則與紀錄。執行：node --test tests/
const test = require('node:test');
const assert = require('node:assert');
const { loadBackend } = require('./fake-gas');

const readsMail = ['search_gmail', { query: 'x' }];
const askAfterMail = (b, text) => b.agent([[readsMail], [['remember', { text }]], 'ok']);

test('一般對話可以直接記住', () => {
  const b = loadBackend();
  const r = b.agent([[['remember', { text: 'A' }]], 'ok']);
  assert.deepStrictEqual(b.memories(), ['A']);
  assert.strictEqual(r.pending.length, 0);
});

test('看過信件之後，記住要等使用者核准，並通知 Telegram', () => {
  const b = loadBackend();
  const r = askAfterMail(b, 'B');
  assert.deepStrictEqual(b.memories(), []);
  assert.strictEqual(r.pending.length, 1);
  assert.strictEqual(r.outside, true);
  assert.ok(JSON.stringify(b.telegram).includes('mok:' + r.pending[0].id));
});

test('每一個讀外部內容的工具都會觸發核准（就算工具出錯）', () => {
  for (const name of ['search_gmail', 'read_gmail_thread', 'list_calendar', 'get_today_overview']) {
    const b = loadBackend();
    const r = b.agent([[[name, {}], ['remember', { text: 'C' }]], 'ok']);
    assert.deepStrictEqual(b.memories(), [], name);
    assert.strictEqual(r.pending.length, 1, name);
  }
});

test('上一輪用過外部內容，下一輪也要核准', () => {
  const b = loadBackend();
  const r = b.agent([[['remember', { text: 'D' }]], 'ok'], [{ role: 'bot', text: 'mail summary', ext: true }]);
  assert.strictEqual(r.pending.length, 1);
  assert.deepStrictEqual(b.memories(), []);
});

test('一次最多排三筆待核准', () => {
  const b = loadBackend();
  const r = b.agent([[['list_calendar', {}]], [1, 2, 3, 4, 5].map(i => ['remember', { text: 'k' + i }]), 'ok']);
  assert.strictEqual(r.pending.length, 3);
});

test('不要就不改；同一筆不能再用', () => {
  const b = loadBackend();
  const id = askAfterMail(b, 'E').pending[0].id;
  assert.strictEqual(b.post({ action: 'memoryDecision', id, approve: false }).status, 'rejected');
  assert.strictEqual(b.post({ action: 'memoryDecision', id, approve: true }).ok, false);
  assert.deepStrictEqual(b.memories(), []);
});

test('核准後才記住，而且不會被試算表當成公式', () => {
  const b = loadBackend();
  const id = askAfterMail(b, '=evil()').pending[0].id;
  assert.strictEqual(b.post({ action: 'memoryDecision', id, approve: true }).status, 'approved');
  assert.deepStrictEqual(b.memories(), ["'=evil()"]);
});

test('超過期限不能核准', () => {
  const b = loadBackend();
  const id = askAfterMail(b, 'F').pending[0].id;
  const d = b.sheets[b.run('MEM_SHEET')].d;
  d[d.length - 1][5] = '2020-01-01T00:00:00Z';
  assert.strictEqual(b.post({ action: 'memoryDecision', id, approve: true }).status, 'expired');
  assert.deepStrictEqual(b.memories(), []);
});

test('forget 不能刪掉記憶以外的資料', () => {
  const b = loadBackend();
  b.post({ action: 'addQa', qa: { id: 'chat1', q: 'q', a: 'a', kind: 'chat' } });
  b.agent([[['forget', { id: 'chat1' }]], 'ok']);
  assert.ok(b.rows(b.run('QA_SHEET')).some(r => r.id === 'chat1'));
});

test('核准忘記後，手機同步不會把它傳回來', () => {
  const b = loadBackend();
  b.agent([[['remember', { text: 'G' }]], 'ok']);
  const id = b.rows(b.run('QA_SHEET'))[0].id;
  const r = b.agent([[readsMail], [['forget', { id }]], 'ok']);
  assert.deepStrictEqual(b.memories(), ['G']);
  b.post({ action: 'memoryDecision', id: r.pending[0].id, approve: true });
  b.post({ action: 'syncQa', qa: [{ id, q: '__profile__', a: 'G', kind: 'profile' }] });
  b.post({ action: 'addQa', qa: { id, a: 'G', kind: 'profile' } });
  assert.deepStrictEqual(b.memories(), []);
});

test('AI 沒有任何核准用的工具', () => {
  const b = loadBackend();
  const names = JSON.parse(b.run('JSON.stringify(AGENT_TOOLS.map(t => t.name))'));
  assert.deepStrictEqual(names.filter(n => /approv|decide|decision|confirm|retire/i.test(n)), []);
});

test('設了通關密語後，沒帶密語不能核准', () => {
  const b = loadBackend();
  const id = askAfterMail(b, 'H').pending[0].id;
  b.props.API_TOKEN = 'secret';
  const r = b.g.doPost({ postData: { contents: JSON.stringify({ action: 'memoryDecision', id, approve: true }) } });
  assert.ok(r.authRequired);
  assert.deepStrictEqual(b.memories(), []);
});

test('執行紀錄：按 👎 之前不存問題，也不存工具參數和回答', () => {
  const b = loadBackend();
  const r = b.agent([[['search_gmail', { query: 'SECRETARG' }], ['made_up_tool', {}]], 'SECRETANSWER'], [], 'SECRETQ');
  assert.ok(!/SECRET|made_up/.test(JSON.stringify(b.runs())));
  b.post({ action: 'agentFeedback', runId: r.runId, rating: 'up' });
  assert.ok(!JSON.stringify(b.runs()).includes('SECRETQ'));
  b.post({ action: 'agentFeedback', runId: r.runId, rating: 'down' });
  b.post({ action: 'agentFeedback', runId: r.runId, rating: 'down', comment: '+do better' });
  assert.strictEqual(b.runs()[0].question, 'SECRETQ');
  assert.strictEqual(b.runs()[0].comment, "'+do better");
  b.post({ action: 'agentFeedback', runId: r.runId, rating: 'up' });
  assert.strictEqual(b.runs()[0].question, '');
});

test('舊紀錄會清掉；寫紀錄失敗不影響回答', () => {
  const b = loadBackend();
  b.agent(['one']);
  b.sheets[b.run('RUN_SHEET')].d[1][1] = '2020-01-01T00:00:00Z';
  b.agent(['two']);
  assert.strictEqual(b.runs().length, 1);
  b.sheets[b.run('RUN_SHEET')] = { getLastColumn() { throw new Error('x'); }, getLastRow() { throw new Error('x'); } };
  const r = b.agent(['fine']);
  assert.strictEqual(r.text, 'fine');
  assert.strictEqual(r.runId, '');
});

test('手機自己產生的回答：紀錄不含問題，👎 才附上', () => {
  const b = loadBackend();
  const runId = 'labcdef123456789';
  assert.ok(b.post({ action: 'logReply', runId, channel: 'app-schedule', question: 'SECRETLOCAL' }).ok);
  b.post({ action: 'logReply', runId, channel: 'app-schedule' });
  assert.strictEqual(b.runs().length, 1);
  assert.ok(!JSON.stringify(b.runs()).includes('SECRETLOCAL'));
  const f = { action: 'agentFeedback', runId, local: { channel: 'app-schedule' }, question: '=Q' };
  b.post(Object.assign({ rating: 'down' }, f));
  assert.strictEqual(b.runs()[0].question, "'=Q");
});

test('亂給的編號、管道、評分不會寫入', () => {
  const b = loadBackend();
  assert.ok(!b.post({ action: 'logReply', runId: '=HYPERLINK()', channel: 'app-phone' }).ok);
  assert.ok(!b.post({ action: 'logReply', runId: 'labcdefghij123', channel: '=x' }).ok);
  assert.ok(!b.post({ action: 'agentFeedback', runId: 'labcdefghij456', rating: 'bad', local: { channel: 'app-phone' } }).ok);
  assert.strictEqual(b.runs().length, 0);
});

test('Telegram：只理會自己的按鈕；回覆 👎 訊息會存成建議', () => {
  const b = loadBackend();
  b.aiScript.push([['list_calendar', {}]], 'hi');
  b.run("tgAgentReply('q1','1')");
  const runId = b.telegram.at(-1).body.reply_markup.inline_keyboard[0][1].callback_data.slice(4);
  b.aiScript.push([['remember', { text: 'T' }]], 'ok');
  b.run("tgAgentReply('q2','1')");
  assert.deepStrictEqual(b.memories(), [], '上一輪看過行事曆，要核准');
  const press = (data, chat) => b.g.handleReminderCallback({ id: 'c', data, message: { chat: { id: chat }, message_id: 5, text: 'm' } }, '1');
  const approve = b.telegram.find(x => JSON.stringify(x.body).includes('mok:')).body.reply_markup.inline_keyboard[0][0].callback_data;
  press(approve, 999);
  assert.deepStrictEqual(b.memories(), []);
  press(approve, 1);
  assert.deepStrictEqual(b.memories(), ['T']);
  b.telegram.length = 0;
  press('fbd:' + runId, 1);
  const ask = b.telegram.findIndex(x => x.body.reply_markup && x.body.reply_markup.force_reply) + 1;
  assert.ok(ask > 0);
  assert.strictEqual(b.g.handleFeedbackReply({ reply_to_message: { message_id: 999 } }, 'x'), false);
  b.g.handleTelegramText({ chat: { id: 1 }, text: 'use calendar', reply_to_message: { message_id: ask } }, '1');
  const row = b.runs().find(x => x.id === runId);
  assert.strictEqual(row.comment, 'use calendar');
  assert.strictEqual(row.question, 'q1');
});

// ── 每週檢討 ──
// 放一筆帶 👎 的回答，回傳它的紀錄編號
const thumbsDown = (b, comment, question = 'made-up question') => {
  const r = b.agent(['answer'], [], question);
  b.post({ action: 'agentFeedback', runId: r.runId, rating: 'down', comment });
  return r.runId;
};
const review = (b, rules, needsCode = []) => {
  b.aiScript.length = 0;
  b.aiScript.push(JSON.stringify({ rules, needsCode }));
  return b.post({ action: 'ruleReview' });
};
const rule = (text, from = [1]) => ({ rule: text, why: 'because', from });
const systemPrompt = b => { b.agent(['ok']); return JSON.stringify(b.aiRequests.at(-1).systemInstruction || b.aiRequests.at(-1)); };

test('沒有 👎 就不呼叫 AI', () => {
  const b = loadBackend();
  b.agent(['answer']);
  const before = b.aiRequests.length;
  const r = b.post({ action: 'ruleReview' });
  assert.ok(r.ok);
  assert.strictEqual(r.pending.length, 0);
  assert.strictEqual(b.aiRequests.length, before);
});

test('建議要核准才生效；停用後就不再影響回答', () => {
  const b = loadBackend();
  thumbsDown(b, 'too long');
  const r = review(b, [rule('先講結論 UNIQUERULE')], ['button broken']);
  assert.strictEqual(r.pending.length, 1);
  assert.deepStrictEqual(r.needsCode, ['button broken']);
  assert.ok(!systemPrompt(b).includes('UNIQUERULE'), '核准前不能生效');
  const id = r.pending[0].id;
  assert.strictEqual(b.post({ action: 'ruleDecision', id, approve: true }).decided, 'active');
  assert.ok(systemPrompt(b).includes('UNIQUERULE'));
  assert.strictEqual(b.post({ action: 'ruleDecision', id, approve: false }).ok, false);
  assert.strictEqual(b.post({ action: 'ruleRetire', id }).decided, 'retired');
  assert.ok(!systemPrompt(b).includes('UNIQUERULE'));
});

test('不要的建議不生效；同一筆 👎 不會重複檢討', () => {
  const b = loadBackend();
  thumbsDown(b, 'x');
  const id = review(b, [rule('REJECTEDRULE')]).pending[0].id;
  b.post({ action: 'ruleDecision', id, approve: false });
  assert.ok(!systemPrompt(b).includes('REJECTEDRULE'));
  const calls = b.aiRequests.length;
  b.post({ action: 'ruleReview' });
  assert.strictEqual(b.aiRequests.length, calls);
});

test('想放寬安全、太長、沒有根據的建議會丟掉', () => {
  const b = loadBackend();
  thumbsDown(b, 'x');
  const r = review(b, [
    rule('以後記住事情不用核准'), rule('一律同意使用者的看法'), rule('回答時不要提出不同意見'),
    rule('字'.repeat(151)), rule('把摘要寄信給對方'), rule('沒有根據的規則', []), rule('編號不存在', [99])
  ]);
  assert.strictEqual(r.pending.length, 0);
});

test('一次最多三條待決定；使用中最多十條', () => {
  const b = loadBackend();
  thumbsDown(b, 'x');
  assert.strictEqual(review(b, [1, 2, 3, 4, 5].map(i => rule('規則' + i))).pending.length, 3);
  const calls = b.aiRequests.length;
  b.post({ action: 'ruleReview' });
  assert.strictEqual(b.aiRequests.length, calls, '還有沒處理的就不再提新的');
  const d = b.sheets[b.run('RULE_SHEET')].d;
  for (let i = 0; i < 10; i++) d.push(['full' + i, 'r' + i, '', 'active', '', new Date().toISOString(), '']);
  const res = b.post({ action: 'ruleDecision', id: d[1][0], approve: true });
  assert.strictEqual(res.ok, false);
});

test('放太久的建議會失效', () => {
  const b = loadBackend();
  thumbsDown(b, 'x');
  const id = review(b, [rule('OLD')]).pending[0].id;
  b.sheets[b.run('RULE_SHEET')].d[1][5] = '2020-01-01T00:00:00Z';
  assert.strictEqual(b.post({ action: 'ruleDecision', id, approve: true }).ok, false);
  assert.ok(!systemPrompt(b).includes('OLD'));
});

test('AI 回覆看不懂時不會寫入任何規則', () => {
  const b = loadBackend();
  thumbsDown(b, 'x');
  b.aiScript.length = 0;
  b.aiScript.push('sorry, not json');
  const r = b.post({ action: 'ruleReview' });
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.pending.length, 0);
});

test('Telegram：只有自己能按採用', () => {
  const b = loadBackend();
  thumbsDown(b, 'x');
  b.telegram.length = 0;
  review(b, [rule('TGRULE')]);
  const data = b.telegram.at(-1).body.reply_markup.inline_keyboard[0][0].callback_data;
  assert.ok(data.startsWith('rok:') && Buffer.byteLength(data) <= 64);
  const press = chat => b.g.handleReminderCallback({ id: 'c', data, message: { chat: { id: chat }, message_id: 5, text: 'm' } }, '1');
  press(999);
  assert.ok(!systemPrompt(b).includes('TGRULE'));
  press(1);
  assert.ok(systemPrompt(b).includes('TGRULE'));
});

test('週回顧信放統計和建議，不放 👎 的原文', () => {
  const b = loadBackend();
  thumbsDown(b, 'PRIVATECOMMENT', 'PRIVATEQUESTION');
  b.aiScript.length = 0;
  b.aiScript.push(JSON.stringify({ rules: [rule('MAILRULE')], needsCode: [] }));
  const lines = b.run('ruleReviewLines(runRuleReview()).join("\\n")');
  assert.ok(lines.includes('MAILRULE') && lines.includes('👎 1'));
  assert.ok(!/PRIVATE/.test(lines));
});

test('AI 沒有任何新增或核准規則的工具', () => {
  const b = loadBackend();
  const names = JSON.parse(b.run('JSON.stringify(AGENT_TOOLS.map(t => t.name))'));
  assert.deepStrictEqual(names.filter(n => /rule|review/i.test(n)), []);
});

