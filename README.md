# 🤖 Ybot 個人助理

把「你做的所有 App」「Email」「生活瑣事」整合到一個可以隨時問、還會主動提醒的助理。

**線上使用：** https://cyysongyy.github.io/ybot/

單一 HTML 檔（`index.html`），瀏覽器直接開啟即可用；可選接 Google Apps Script 雲端後台（`ybot-backend.gs`）取得日曆／Gmail／每日簡報等進階功能。

> 本專案原本是 [remediation-system](https://github.com/cyysongyy/remediation-system) 的一部分，已獨立成專屬 repo。整合入口仍可從 [portal.html](https://cyysongyy.github.io/remediation-system/portal.html) 連回。

---

## ✨ 功能特色

| 功能 | 說明 |
|---|---|
| 💬 對話問答 | 接 Gemini / OpenAI / NVIDIA，隨時問行程、待辦、信件、任何生活瑣事 |
| 📅 日曆彙整 | 自動讀取 Google 日曆未來 7 天行程，回答時自動帶入 |
| 📬 Gmail 彙整 | 自動彙整近 2 天未讀信件（主旨／寄件者／摘要），可請 Ybot 幫忙摘要 |
| 📝 待辦與筆記 | 三種類型：📌 瑣事筆記／✅ 待辦／⏰ 定時提醒，本機優先、可選雲端同步 |
| ⏰ 到點提醒 | 提醒時間一到，後台自動寄 Email 提醒（每 30 分鐘檢查一次） |
| 🌅 每日簡報 | 每天早上 7 點自動寄「Ybot 每日簡報」：今日行程、🔴🟠🟢 優先分級的待辦、未讀信件重點、AI 個人化提醒 |
| 🎯 優先分級 | 待辦／提醒自動分成 🔴今天必須／🟠建議今天／🟢可延後，規則先判斷到期日，AI 只負責「升級」看起來緊急但沒填到期日的項目 |
| 🌙 晚間彙整 | 晚上 8 點檢查是否還有 🔴🟠 未處理事項，**只有真的有事才寄信**，沒事就不打擾 |
| 🔗 整合其他 App | 可串接 [remediation-system](https://github.com/cyysongyy/remediation-system) 的 `remediation-backend.gs`／`health-backend.gs`，對話與簡報一併帶入考卷批改、健康預測的最新摘要 |
| 🧩 我的其他 App | 「設定 → 我的其他 App」可自由新增任何你自建的 App（名稱＋網頁網址，選填後台 API 網址），在對話快捷列一鍵開啟，Ybot 也會知道這些 App 的存在 |

---

## 📁 檔案說明

```
index.html                  ← Ybot 個人助理主程式（單一 HTML 檔，無需安裝）
ybot-backend.gs              ← Google Apps Script 後台（選填）
ybot-icon.svg                ← App 圖示
ybot-manifest.webmanifest    ← PWA manifest
```

---

## 🔗 資料串接方式

- **行程／信件**：由 `ybot-backend.gs`（Google Apps Script）透過 `CalendarApp` / `GmailApp` 讀取你自己 Google 帳號的資料，只回傳給你自己部署的前端。
- **其他系統**：在 Ybot「設定 → 整合其他系統」填入 `remediation-backend.gs`／`health-backend.gs` 的部署網址，Ybot 後台會伺服器對伺服器彙整近期成績/迷思概念、最新健康分與生理年齡等摘要。
- **其他 GitHub repo 的自建 App**（例如 `principal-assistant`、`micro-politics`、`5-steps-rehab-for-footdrop`）：在「設定 → 我的其他 App」填入名稱與網址，預設已帶入這幾個。若該 App 之後也部署了 Apps Script 後台，把後台網址填進「後台 API 網址」欄位，Ybot 每日簡報與對話會嘗試以 `?action=all` 拉取摘要（best-effort，讀不到會自動略過，不影響其他資料）。
- **生活瑣事**：直接在「📝 待辦與筆記」輸入，Ybot 對話時會一併參考。

---

## ☁️ 雲端後台與自動化（`ybot-backend.gs`）

1. 開啟一份 Google 試算表 → **擴充功能** → **Apps Script**
2. 貼上 `ybot-backend.gs` 的內容（取代原有內容）
3. 執行一次 `setupDailyBrief()`（授權後排程每日 7 點自動簡報，含 🔴🟠🟢 優先分級）
4. 再執行一次 `setupReminderWatch()`（排程每 30 分鐘檢查一次到點提醒）
5. 再執行一次 `setupEveningDigest()`（排程晚上 8 點檢查，只有還有未處理事項才寄信）
6. **部署** → 新增部署作業 → 網頁應用程式
   - 以下列身分執行：**我（Me）**
   - 誰可以存取：**所有人（Anyone）**
7. 複製部署網址 → 回到 `index.html`：**⚙️ 設定 → 雲端後台** → 貼上網址 → **🔌 測試連線**

> ⚠️ 因為要讀取 Gmail／日曆，**重新部署後首次執行會要求你額外授權**，請同意（僅你本人帳號讀取，資料只回傳給你自己部署的前端，不經第三方伺服器）。

**AI 個人化簡報（選填）**：在 Apps Script 編輯器執行一次 `setAiConfig('gemini','你的KEY')`（或 `'openai','sk-...'`），金鑰存於 Script Properties；之後每日簡報會附上 AI 生成的「今日提醒」，AI 也會幫忙把「沒填到期日但內容看起來緊急」的項目升級到 🔴/🟠（沒把握就不會動，避免亂猜）。未設定 AI Key 時，優先分級仍會照到期日的規則正常運作，只是不會有「今日提醒」文字與升級判斷。

---

## 💾 資料儲存說明

- **本機模式**：待辦／筆記與對話紀錄預設存於瀏覽器 `localStorage`，清除瀏覽器資料會遺失
- **雲端模式**：設定 GAS 後台後可同步至 Google 試算表，跨裝置查看
- **隱私保護**：AI Key 僅存本機、直接呼叫 Gemini / OpenAI / NVIDIA API，不經過任何第三方伺服器；雲端後台為選填，資料只存在你自己的 Google 試算表

---

## 👤 關於

本系統由 Young Chen 開發 · 2026
