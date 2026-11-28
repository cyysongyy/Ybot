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
        errors, posts = [], []
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
                if body['action'].startswith('rule'):
                    if body['action'] == 'ruleDecision':
                        review.update(pending=[], active=[{'id': 'r1', 'text': '<b>先講結論</b>'}])
                    res = dict(ok=True, **review)
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

        assert not errors, errors
        print('app tests passed')
        browser.close()


if __name__ == '__main__':
    main()
