// Ybot 的 service worker：目的只有一個——沒有網路的時候 App 還打得開。
//
// 為什麼需要它：錄音（getUserMedia + MediaRecorder）和存筆記（localStorage）
// 本來就完全不需要網路，真正卡住的是「離線時連頁面都跟 GitHub Pages 拿不到」。
// 把 App 本體快取起來，這層就通了。
//
// 快取策略刻意選「網路優先、離線才回退」：
// 這個 App 更新很頻繁（單檔 HTML，每次合併 PR 就是一個新版本）。如果用常見的
// 「快取優先」，使用者會卡在舊版本、得手動清快取才看得到新功能。網路優先的代價
// 只是連線時多一次請求，換來的是「線上永遠是最新版、離線永遠打得開」。
const VERSION='ybot-v1';
// 只放 App 本體。使用者資料在 localStorage / IndexedDB，不歸這裡管。
const SHELL=['./','./index.html','./ybot-manifest.webmanifest','./ybot-icon.svg'];

self.addEventListener('install',e=>{
  // 有任何一個檔案抓不到就整批不裝，免得裝出一個半殘的快取。
  e.waitUntil(caches.open(VERSION).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()));
});

self.addEventListener('activate',e=>{
  e.waitUntil((async()=>{
    const names=await caches.keys();
    await Promise.all(names.filter(n=>n!==VERSION).map(n=>caches.delete(n)));
    await self.clients.claim();
  })());
});

// 這個 sw 的 scope 是 /Ybot/，但 scope 內的頁面「送出的所有請求」都會經過這裡，
// 包括打去 Gemini／OpenAI／Apps Script 的 API，以及去隔壁 App 抓課表的請求。
// 那些一律不碰——快取 API 回應會拿到過期的答案，比沒有離線功能還糟。
function isShellRequest(req){
  if(req.method!=='GET')return false;
  const url=new URL(req.url);
  if(url.origin!==self.location.origin)return false;
  // 只認這個 sw 自己資料夾底下的東西（/Ybot/），不越界管別的 App。
  return url.pathname.startsWith(new URL('./',self.location.href).pathname);
}

// 查詢字串對這些靜態檔案沒有意義（?rec=1 拿到的還是同一份 index.html）。
// 存取都用去掉問號後的網址，才不會每來一種參數就多存一份、離線時又對不上。
function cacheKey(req){
  const u=new URL(req.url);
  return u.origin+u.pathname;
}

self.addEventListener('fetch',e=>{
  if(!isShellRequest(e.request))return;  // 不呼叫 respondWith 就是走瀏覽器預設，最安全
  e.respondWith((async()=>{
    const key=cacheKey(e.request);
    try{
      // cache:'no-cache' 是必要的：GitHub Pages 會回 Cache-Control: max-age，
      // 直接 fetch 會拿到瀏覽器 HTTP 快取裡的舊版本，「網路優先」就形同虛設。
      // 這樣會送出帶條件的請求，沒改就回 304，很省。
      const fresh=await fetch(key,{cache:'no-cache',credentials:'same-origin'});
      // 只存成功的回應，404 之類的存了會害人。
      if(fresh&&fresh.ok){
        const c=await caches.open(VERSION);
        await c.put(key,fresh.clone());
      }
      return fresh;
    }catch(err){
      // 離線了：拿快取。導覽請求（開 App）額外回退到 index.html，
      // 這樣捷徑帶著 ?rec=1 進來、或直接開資料夾網址都開得起來。
      const hit=await caches.match(key);
      if(hit)return hit;
      if(e.request.mode==='navigate'){
        const shell=await caches.match(new URL('./index.html',self.location.href).href);
        if(shell)return shell;
      }
      throw err;
    }
  })());
});
