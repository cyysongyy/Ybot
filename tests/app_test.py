"""手機大小的瀏覽器測試。後台是假的，其他網路一律擋掉。只放虛構資料。
執行：pip install playwright && playwright install chromium && python tests/app_test.py"""
import json
import os
from pathlib import Path
from playwright.sync_api import sync_playwright

APP = (Path(__file__).resolve().parent.parent / 'index.html').as_uri()

# 在講事情、問意思、請它寫東西：就算句子裡有「活動」「會議」，也要交給 AI，不能回行程清單。
NOT_SCHEDULE = [
    '有位老師在活動時跟學生說學校只給十五分鐘，結果別的活動做不完。這個老師背後的意思是什麼？',
    '今天有位老師在活動時一直跟學生說時間不夠，他背後的意思是什麼？',
    '上次的活動辦得不太好，我該怎麼檢討？',
    '幫我寫一段運動會活動的開場致詞',
    '會議紀錄要怎麼寫比較清楚？',
    '家長對校外教學活動有意見，怎麼回覆比較好？',
    '研習心得幫我潤飾一下',
    '明天下午三點加一個行政會議',
]
SCHEDULE = ['明天有什麼行程？', '這週有哪些會議？', '下週的活動', '今天行程']


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 390, 'height': 760})
        errors, posts, old_backend = [], [], []
        page.on('pageerror', lambda e: errors.append(str(e)))

        stats = {'total': 3, 'up': 1, 'down': 1, 'errors': 0}
        review = {'pending': [{'id': 'r1', 'text': '<b>先講結論</b>', 'why': '虛構原因'}], 'active': [], 'stats': {'week': stats, 'last': stats}}

        def route(r):
            url = r.request.url
            if url.startswith('file:'):
                return r.continue_()
            if not url.startswith('https://script.google.com/'):
                return r.abort()
            if r.request.method == 'POST':
                body = json.loads(r.request.post_data)
                posts.append(body)
                res = {'ok': True}
                if old_backend and body['action'] == 'read':
                    res = {'ok': False, 'error': 'Unknown action: read'}
                if body['action'].startswith('rule'):
                    if body['action'] == 'ruleDecision':
                        review.update(pending=[], active=[{'id': 'r1', 'text': '<b>先講結論</b>'}])
                    res = dict(ok=True, **review)
                    if body['action'] == 'ruleTrial':
                        res['trial'] = {'ruleId': 'r1', 'rule': '<b>先講結論</b>', 'total': 1, 'withRule': 'b',
                                        'items': [{'q': '虛構題目', 'a': '<img src=x>甲的回答', 'b': '乙的回答'}]}
            else:
                res = {'ok': True, 'qa': [], 'forgotten': [], 'notes': []}
            return r.fulfill(json=res, headers={'Access-Control-Allow-Origin': '*'})

        page.route('**/*', route)
        page.goto(APP)
        page.wait_for_timeout(1500)

        for q in NOT_SCHEDULE:
            assert not page.evaluate('q => !!detectScheduleRange(q)', q), '不該回行程清單：' + q
        for q in SCHEDULE:
            assert page.evaluate('q => !!detectScheduleRange(q)', q), '應該回行程：' + q
        print('pass 行程捷徑')

        page.evaluate("""() => {
          state.gasUrl = 'https://script.google.com/macros/s/fake/exec';
          state.agentMode = false; state.chat = [];
          state.profile = [{id: 'pOLD', text: '舊記憶', at: new Date().toISOString()}];
          window.callAI = async () => '好的<<<ACTION>>>{"actions":[{"op":"remember","text":"<i>新記憶</i>"},{"op":"forget","id":"pOLD"}]}<<<END>>>';
          let e = document.getElementById('chatBox');
          while (e && e !== document.body) { if (getComputedStyle(e).display === 'none') e.style.display = 'block'; e = e.parentElement; }
        }""")

        def send(text):
            page.evaluate('t => { document.getElementById("chatInput").value = t; sendChat(); }', text)
            page.wait_for_timeout(700)

        def memories():
            return page.evaluate('state.profile.map(x => x.text)')

        send('請記住這件事')
        assert memories() == ['舊記憶'], '備援 AI 不能自己改記憶'
        waiting = page.evaluate('state.chat.map((m, i) => m.mem ? i : -1).filter(i => i >= 0)')
        assert len(waiting) == 2
        page.evaluate(f'decideMemory({waiting[0]}, true)')
        page.evaluate(f'decideMemory({waiting[1]}, false)')
        assert memories() == ['<i>新記憶</i>', '舊記憶']
        assert page.locator('#chatBox i').count() == 0, '記憶內容不能被當成 HTML'
        print('pass 手機端記憶要核准')

        send('明天有什麼行程？')
        bots = page.evaluate("state.chat.map((m, i) => m.role === 'bot' ? i : -1).filter(i => i >= 0)")
        assert page.locator('.fbRow').count() == len(bots), '每則回答都要有 👍／👎'
        page.evaluate(f"rateReply({bots[-1]}, 'down')")
        page.wait_for_timeout(500)
        assert all('question' not in x for x in posts if x['action'] == 'logReply')
        assert [x for x in posts if x['action'] == 'agentFeedback'][-1]['question'] == '明天有什麼行程？'
        print('pass 回答評分')

        page.evaluate('openReview()')
        page.wait_for_timeout(700)
        card = page.locator('#reviewCard')
        assert card.is_visible() and '等你決定' in card.inner_text()
        assert page.evaluate('rulesPrompt()') == '', '核准前不能給備用 AI'
        assert card.locator('b').count() == 0, '規則文字不能被當成 HTML'
        assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth'), '手機畫面不能左右捲'
        if os.environ.get('SHOT'):
            card.screenshot(path=os.environ['SHOT'])
        card.get_by_text('✅ 採用').click()
        page.wait_for_timeout(700)
        assert [x for x in posts if x['action'] == 'ruleDecision'] == [dict(posts[-1])]
        assert '先講結論' in page.evaluate('rulesPrompt()')
        print('pass 每週檢討')

        review.update(issues=[{'id': 'i1', 'text': '虛構問題：按鈕沒反應'}],
                      active=[{'id': 'r1', 'text': '<b>先講結論</b>', 'flags': 2, 'flagWhy': '虛構原因'}],
                      history={'adopted': 1, 'rejected': 0, 'retired': 0, 'expired': 0, 'hint': ''})
        page.evaluate('openReview()')
        page.wait_for_timeout(700)
        card.get_by_text('📋 複製給 Claude 修').click()
        box = page.locator('#repairBox')
        assert '虛構問題' in box.input_value() and '不要自己合併' in box.input_value()
        box.fill('使用者改過的內容')
        page.evaluate('window.copied = null; window.copyTextSafe = async t => { window.copied = t; return true; }')
        card.get_by_text('複製這段文字').click()
        page.wait_for_timeout(300)
        assert page.evaluate('window.copied') == '使用者改過的內容', '要複製使用者改過的版本'
        card.get_by_text('🧪 試答').click()
        page.wait_for_timeout(700)
        trial = page.locator('#trialBox')
        assert '套用規則的是' not in trial.inner_text(), '看答案前不能洩漏'
        assert page.locator('#reviewCard img, #reviewCard b:text-is("先講結論")').count() == 0
        trial.get_by_text('看哪邊有套用規則').click()
        assert '套用規則的是：乙' in trial.inner_text()
        assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth')
        if os.environ.get('SHOT'):
            card.screenshot(path=os.environ['SHOT'])
        print('pass 修正清單和試答')

        # 校務資料：敏感欄位和學生個資不能送給 AI（全部是虛構資料）
        school = page.evaluate("""() => {
          const put = (k, v) => localStorage.setItem(k, JSON.stringify(v));
          put('ck_staff', [{name:'測試老師甲', unit:'虛構處室', title:'組長', rank:'SECRET-rank', idno:'SECRET-id', salary:'SECRET-pay',
            phone:'SECRET-tel', mobile:'SECRET-mob', birth:'SECRET-birth', startDate:'SECRET-start', addr:'SECRET-addr', note:'SECRET-note'}]);
          put('ck_roster_cache', [1,2,3,4,5].map(i => ({name:'SECRET-kid'+i, cls:'虛構班', seat:'SECRET-seat', sid:'SECRET-sid'+i})));
          const g = (cls, i) => ({name:'SECRET-kid'+i, cls, subject:'數學', score:80+i, year:115, semester:1, stage:'期中', date:'2030-01-01'});
          put('cksch_grades', {data:[1,2,3,4,5].map(i => g('虛構班', i)).concat([1,2,3,4].map(i => g('小班', i)), [g('小班', 4)])});
          return buildSchoolDbContext();
        }""")
        assert 'SECRET' not in school, school
        assert '測試老師甲(虛構處室·組長)' in school and '虛構班 5人' in school
        assert '虛構班 數學 平均83.0分（5人）' in school
        assert '小班' not in school, '不到 5 位學生的平均不能送'
        print('pass 校務資料不含個資')

        # 手機端 AI 的每一種動作都要按了才做
        page.evaluate("""() => {
          state.chat = [];
          state.notes = [{id:'keep', type:'todo', content:'虛構保留', done:false, createdAt:new Date().toISOString()}];
          window.callAI = async () => '好<<<ACTION>>>' + JSON.stringify({actions:[
            {op:'delete_note', id:'keep'}, {op:'complete_note', id:'keep'}, {op:'update_note', id:'keep', content:'<i>改過</i>'},
            {op:'add_note', type:'todo', content:'虛構新增'}, {op:'add_calendar_event', title:'虛構行程', start:'2030-01-02T09:00:00+08:00'}]}) + '<<<END>>>';
        }""")
        snapshot = 'JSON.stringify(liveNotes().map(n => [n.id, n.content, !!n.done]))'
        before, sent = page.evaluate(snapshot), len(posts)
        send('幫我整理清單')
        assert page.evaluate(snapshot) == before, '按之前什麼都不能變'
        assert not [x for x in posts[sent:] if x['action'] in ('addNote', 'updateNote', 'deleteNote', 'addCalendarEvent')]
        acts = page.evaluate("state.chat.map((m, i) => m.mem && m.mem.op === 'act' ? i : -1).filter(i => i >= 0)")
        assert len(acts) == 5, acts
        assert page.locator('#chatBox').get_by_text('✅ 執行').count() == 5
        assert page.locator('#chatBox i').count() == 0
        page.evaluate(f'decideMemory({acts[0]}, false)')
        assert page.evaluate(snapshot) == before
        page.evaluate(f'Promise.all([decideMemory({acts[3]}, true), decideMemory({acts[3]}, true)])')
        assert page.evaluate("liveNotes().filter(n => n.content === '虛構新增').length") == 1, '連按只能做一次'
        page.evaluate(f'decideMemory({acts[2]}, true)')
        assert page.evaluate("liveNotes().find(n => n.id === 'keep').content") == '<i>改過</i>'
        page.evaluate(f'decideMemory({acts[1]}, true)')
        assert page.evaluate("liveNotes().find(n => n.id === 'keep').done") is True
        page.evaluate(f'decideMemory({acts[4]}, true)')
        page.wait_for_timeout(500)
        assert [x for x in posts if x['action'] == 'addCalendarEvent']
        print('pass 手機端動作要核准')

        # 不安全的連結不能點
        for bad in ['javascript:alert(1)', ' JaVaScRiPt:alert(1)', 'java\tscript:alert(1)', 'data:text/html,x', 'vbscript:x']:
            assert page.evaluate('u => safeUrl(u)', bad) == '#', bad
        assert page.evaluate("safeUrl('https://example.com/a?b=1')") == 'https://example.com/a?b=1'
        page.evaluate("state.links = [{id:'l1', name:'虛構', url:'javascript:alert(1)'}]; renderLinks()")
        assert page.locator('a[href^="javascript:alert"]').count() == 0
        print('pass 連結安全')

        # 讀取用 POST，網址裡沒有密語；舊版後台才退回 GET
        urls = []
        page.on('request', lambda q: urls.append((q.method, q.url)) if q.url.startswith('https://script.google.com/') else None)
        page.evaluate("async () => { state.gasToken = 'made-up-pass'; await (await gasGet('qa')).json(); }")
        assert [m for m, u in urls] == ['POST'] and 'made-up-pass' not in urls[0][1]
        assert posts[-1] == {'t': 'made-up-pass', 'action': 'read', 'get': 'qa', 'params': {}}
        old_backend.append(True)
        urls.clear()
        assert page.evaluate("async () => (await (await gasGet('qa')).json()).ok") is True
        assert [m for m, u in urls] == ['POST', 'GET']
        print('pass 密語不放網址')

        assert not errors, errors
        print('app tests passed')
        browser.close()


if __name__ == '__main__':
    main()
