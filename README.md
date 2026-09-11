# 營運工作台

這個 GitHub 專案用來同步工作台程式碼。Google 試算表資料仍由原本的雲端服務讀取，不會存進這個專案。

## 第一次在家裡使用

1. 安裝 Node.js 與 GitHub Desktop。
2. 在 GitHub Desktop 登入後，複製 `at19750919-star/loan-ledger` 專案。
3. 將 `.env.example` 複製一份並改名為 `.env`。
4. 打開 `.env`，填入自己的 `OPENAI_API_KEY`。
5. 雙擊 `start-boss-desk-ai.bat`，瀏覽器會開啟營運工作台。

`.env` 含有私人金鑰，已設定成不會上傳到 GitHub。

## 平常同步方式

- 開始修改前：在 GitHub Desktop 按 **Fetch origin**，有更新時再按 **Pull origin**。
- 修改完成後：填寫修改摘要，按 **Commit to main**，再按 **Push origin**。
- 換到另一台電腦時：先按 **Fetch origin / Pull origin** 取得最新版。

## 啟動方式

雙擊 `start-boss-desk-ai.bat`。工作台網址為：

`http://127.0.0.1:3218/`

這是該台電腦自己的本機網址；GitHub 負責同步程式碼，不是公開網站。
