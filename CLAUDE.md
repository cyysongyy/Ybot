# 專案筆記（給未來處理這個 repo 的 Claude 看）

這個 repo 是從 `remediation-system` 獨立出來的 Ybot 個人助理（`index.html` + `ybot-backend.gs`）。

## Ybot 的 AI 模型設定

- `index.html` 與 `ybot-backend.gs` 呼叫 Gemini API 時，一律使用穩定別名
  `gemini-flash-latest`，**不要**寫死成 `gemini-2.5-flash` 這種帶版本號的字串。
  Google 會定期棄用舊版本（曾在官方下架日之前就開始回傳「model not found」），
  寫死版本號會讓使用者莫名其妙「AI 突然都失效」。若之後又要指定特定世代
  （例如效能/價格考量），至少也要先查證該版本目前仍受支援。
- `index.html` 裡 NVIDIA 的 `NV_DEFAULT_MODEL` 也會被 NVIDIA 常態性換掉／棄用。
  使用者回報「AI 呼叫失效」時，先懷疑是不是模型 id 過期，可上
  https://build.nvidia.com/models 查目前可用的模型 id 再更新。
- 使用者貼過來的 AI API Key（Gemini／OpenAI／NVIDIA）**絕對不要**寫進程式碼、
  commit、PR 說明或任何會進 repo 的地方。Key 只存在使用者自己瀏覽器的
  localStorage（前端設定頁）或 Apps Script 的 PropertiesService（後台，
  透過 `setAiConfig()` 由使用者自己在 Apps Script 編輯器執行）。只需要指引
  使用者去哪個欄位貼，不要代為儲存。

## 與 remediation-system 的關聯

- Ybot 可選填串接 `remediation-system` repo 的 `remediation-backend.gs`／
  `health-backend.gs` 部署網址（設定 → 整合其他系統），伺服器對伺服器彙整
  考卷批改／健康預測摘要。這兩個後台檔案仍留在 `remediation-system`，不要
  複製進本 repo。
- 頁面右上角「← 回整合入口」連回 `remediation-system` 的
  `https://cyysongyy.github.io/remediation-system/portal.html`，若該 repo
  改名或搬遷，記得回來更新這個連結。

## Git 工作流程

- 這個使用者常常直接在 GitHub 上把 PR 合併掉，不會等在這邊確認。
  每次要 push 前，先 `git fetch origin main` 再
  `git merge-base --is-ancestor <上一個commit> origin/main` 檢查是否已被合併；
  已合併的話要 `git reset --hard origin/main` 後把還沒推的變更疊上去，
  重新開一個 PR，不要假設舊 PR 還開著。
