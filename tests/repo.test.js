// 不用執行就能抓到的錯：語法、版本號、兩邊要一致的文字、不能出現的東西。
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');

const root = path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const backend = read('ybot-backend.gs');
const app = read('index.html');
const pick = (src, re, what) => { const m = src.match(re); assert.ok(m, '找不到 ' + what); return m[1]; };

test('後台和前端的程式沒有語法錯誤', () => {
  new vm.Script(backend, { filename: 'ybot-backend.gs' });
  const blocks = [...app.matchAll(/<script(?![^>]*\bsrc=)([^>]*)>([\s\S]*?)<\/script>/g)]
    .filter(m => !/type=["'](?!text\/javascript)/.test(m[1]));
  assert.ok(blocks.length > 0);
  blocks.forEach((m, i) => new vm.Script(m[2], { filename: 'index.html script ' + i }));
  new vm.Script(read('sw.js'), { filename: 'sw.js' });
});

test('版本號格式正確，APP_VERSION 等於更新紀錄最上面那筆', () => {
  const format = /^\d{4}-\d{2}-\d{2}(-\d+)?$/;
  assert.match(pick(backend, /const BACKEND_VERSION = '([^']+)'/, 'BACKEND_VERSION'), format);
  const appVersion = pick(app, /const APP_VERSION='([^']+)'/, 'APP_VERSION');
  assert.match(appVersion, format);
  const top = app.slice(app.indexOf('const CHANGELOG=')).match(/['"](\d{4}-\d{2}-\d{2}(?:-\d+)?)['"]/);
  assert.strictEqual(top && top[1], appVersion);
});

test('PR 改了後台就一定要改 BACKEND_VERSION', t => {
  const base = process.env.BASE_SHA;
  if (!base) return t.skip('只在 PR 上檢查');
  const before = execFileSync('git', ['show', base + ':ybot-backend.gs'], { cwd: root, maxBuffer: 1e8 }).toString();
  if (before === backend) return;
  const v = s => s.match(/const BACKEND_VERSION = '([^']+)'/)[1];
  assert.notStrictEqual(v(backend), v(before), '後台改了但版本號沒改，使用者不會知道要重貼');
});

test('後台和手機的領導架構、判斷規則一致', () => {
  assert.strictEqual(
    pick(app, /const LEADERSHIP_OUTLINE=`([\s\S]*?)`;/, 'LEADERSHIP_OUTLINE'),
    pick(backend, /const LEADERSHIP_FRAMEWORK = `([\s\S]*?)`;/, 'LEADERSHIP_FRAMEWORK'));
  assert.strictEqual(
    pick(app, /const LEAD_HINTS=(\/.*\/i);/, 'LEAD_HINTS'),
    pick(backend, /const LEAD_RE = (\/.*\/i);/, 'LEAD_RE'));
});

test('repo 裡沒有看起來像 API 金鑰的字串', () => {
  const files = execFileSync('git', ['ls-files'], { cwd: root }).toString().split('\n')
    .filter(f => /\.(gs|js|html|md|json|yml|py|webmanifest)$/.test(f));
  // 唯一例外：index.html 裡 MP_FB_CONFIG 的 Firebase 網頁端識別值（本來就公開，靠登入和資料庫規則把關）。
  // 用雜湊比對，只放行這一個值。
  const allowed = new Set(['f7591cab58fa06367a433672a666149ebb457ecce2d1c382eeec1a5879ed0f14']);
  const hash = v => require('crypto').createHash('sha256').update(v).digest('hex');
  const keys = /AIza[0-9A-Za-z_-]{35}|\bsk-(proj-)?[A-Za-z0-9_-]{32,}|\bnvapi-[A-Za-z0-9_-]{32,}|\b\d{8,10}:AA[A-Za-z0-9_-]{33}\b/;
  const found = files.filter(f => (read(f).match(new RegExp(keys.source, 'g')) || []).some(v => !allowed.has(hash(v))));
  assert.deepStrictEqual(found, []);
});

test('安全邊界：後台不寄別人的信、不刪信、不刪行事曆', () => {
  assert.doesNotMatch(backend, /GmailApp\.sendEmail|\.moveToTrash|\.moveThreadToTrash|\.moveThreadsToTrash|\.deleteEvent|\.deleteEventSeries/);
  assert.doesNotMatch(backend, /\.(reply|replyAll|forward)\(/);
});
