# 霖霖的課程預約系統 (AppSheet 轉移 GitHub + Google Sheets 方案)

本專案將原本 AppSheet 的課程預約功能改寫為 **現代化響應式單頁網頁**，可直接免費託管於 **GitHub Pages**，並以 **Google Sheets（試算表）** 作為後端資料庫，實現零伺服器成本、管理者最直覺的維護體驗。

---

## 🌟 核心特色

1. **極致友善的維護方式**：
   - 管理者直接在 Google Sheets 上修改課程時間、名額、講師，或是於專屬後台界面點選操作，儲存即同步。
2. **AppSheet 經典風格介面 + 現代美化**：
   - 採用 Mobile-First 設計，支援手機、平板與電腦全尺寸自適應。
   - 頂部日期快速篩選、名額進度條、即時防重複預約、額滿自動鎖定。
3. **學員報名端 (`index.html`)**：
   - 一鍵預約、自動記憶學員聯絡資料（下次免重複輸入）、預約成功彈跳憑證。
4. **管理者維護端 (`admin.html`)**：
   - 儀表板總覽（總上架、報名人次、出席率）、關鍵字即時搜尋、一鍵「簽到」與「取消」、線上排課表單。
5. **完全免費託管**：
   - GitHub Pages 免費靜態託管 + Google Apps Script 免費 API 服務。

---

## 🚀 三分鐘上線步驟

### 步驟 1：建立 Google Sheet 與後端腳本 (GAS)
1. 前往 [Google 雲端硬碟](https://drive.google.com/)，建立一份全新的 Google 試算表（例如命名為：`霖霖課程預約資料庫`）。
2. 在試算表上方選單點選 **「擴充功能」 -> 「Apps Script」**。
3. 將本專案中的 `GoogleAppsScript_Code.js` 全部內容複製，貼入 Apps Script 的編輯器中（替換掉原本的 `Code.gs`），按儲存。
4. 在編輯器上方執行函式選擇 `setupSheets`，點擊「**執行**」（若跳出授權視窗，點選「進階」->「前往...」允許存取）。執行完畢後，回到試算表即會看到自動產生的《課程清單》、《預約記錄》、《系統設定》3 個工作表！
5. 點擊右上角 **「部署」 -> 「新建部署」**：
   - 類型齒輪選：**網路應用程式 (Web App)**
   - 說明：`課程預約 API`
   - 執行身分：`我`
   - **誰可以存取：`任何人 (Anyone)`** （重要！確保學員免登入 Google 即可報名）
6. 點擊部署，複製產生的 **網路應用程式網址 (Web App URL)**。

---

### 步驟 2：設定網頁並部署到 GitHub Pages
1. 將本專案（`index.html`、`admin.html`、`README.md`）上傳至您的 GitHub 儲存庫（Repository）。
2. 在 `index.html` 與 `admin.html` 中：
   - 您可以直接在管理端網頁點擊右上角 **「設定 API 網址」** 將步驟 1 的網址貼上並儲存（會記錄於瀏覽器中）。
   - 或者直接編輯 `index.html` 第 184 行的 `API_URL = "YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL"` 替換為您的網址。
3. 進入 GitHub Repository 的 **Settings -> Pages**：
   - Source 選擇 `Deploy from a branch`。
   - Branch 選擇 `main`（或 `master`），資料夾選 `/(root)`，按 Save。
4. 稍等約 1 分鐘，GitHub 會給您一個公開網址（例如 `https://<帳號>.github.io/<專案名>/`），即刻正式上線！

---

## 📂 檔案清單

- `index.html`：學員線上報名首頁（手機優先、卡片式檢視、日期篩選、即時額滿保護）
- `admin.html`：管理員維護後台（即時簽到、名冊過濾、一鍵排課）
- `GoogleAppsScript_Code.js`：Google Sheets 後端 API 程式碼（自動建表、CORS 支援、防重複登記鎖）
