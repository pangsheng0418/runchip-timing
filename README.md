# 智慧路跑 NFC 晶片計時系統

支援 **單機離線運作 (PWA + IndexedDB)** 與 **多機即時連線同步 (Google Sheets + Apps Script)** 的網頁版 NFC 路跑計時系統，免編譯即可直接部署於 GitHub Pages、Vercel 或 Cloudflare Pages。

## 檔案結構

| 檔案路徑 | 功能說明 |
|---|---|
| `index.html` | 工作人員與後台主系統（賽道感應計時、晶片寫入報到、成績修正/匯出、賽事設定） |
| `public.html` | 公開選手動態即時查詢網頁（免登入、自動隱碼身份證字號） |
| `gas/Code.gs` | Google Apps Script 後端 API（含自動建表、併發鎖定、防連點去重、隱碼雜湊） |
| `js/timing-engine.js` | 大會時間、淨時間、分段配速、ETA 預估、總名次與分組名次運算核心 |
| `js/db.js` | 本地 IndexedDB 與 LocalStorage 離線儲存層 |
| `js/nfc.js` | Web NFC API 讀取/寫入/自動回讀驗證、Web Audio 提示音與模擬器 |
| `js/sync.js` | Google Apps Script 雲端同步與離線佇列補傳模組 |
| `js/excel.js` | SheetJS (`.xlsx`) 多工作表成績匯出模組 |
| `manifest.json` / `sw.js` | PWA 安裝與離線快取設定 |

---

## 一、Google 試算表工作表欄位結構

執行 `gas/Code.gs` 中的 `setupSheets()` 會自動建立以下 5 張工作表：

1. **`Config`**（賽事全域設定）：`key` | `value` | `description`
2. **`Categories`**（分組與起跑時間）：`id` | `name` | `distanceKm` | `gunTimeMs` | `gunTimeIso` | `color`
3. **`Checkpoints`**（感應站點設定）：`id` | `name` | `order` | `distanceKm` | `type` (`start`/`checkpoint`/`finish`)
4. **`Runners`**（選手名單與罰時）：`bib` | `name` | `idNumber` | `categoryId` | `nfcUid` | `penaltySec` | `checkedInAt` | `updatedAt`
5. **`Scans`**（晶片感應紀錄）：`id` | `bib` | `checkpointId` | `timestampMs` | `timeStr` | `isoTime` | `source` | `deviceId` | `deleted` | `note` | `updatedAt`

---

## 二、Google Apps Script (GAS) 部署步驟

1. 建立一份新的 **Google 試算表**。
2. 點選上方選單「**擴充功能**」→「**Apps Script**」。
3. 將 `gas/Code.gs` 內容完整貼入編輯器並儲存。
4. 在工具列選擇函式 **`setupSheets`**，點擊「**執行**」並完成 Google 帳號權限授權（工作表將自動建立完成）。
5. 點選右上角「**部署**」→「**新增部署作業**」：
   - 類型選擇「**網頁應用程式**」
   - 執行身分：**我**
   - 誰可以存取：**所有人 (Anyone)**
6. 複製產生的 `https://script.google.com/macros/s/....../exec` 網址。
7. 開啟 `index.html` →「**賽事設定與同步**」（預設管理員密碼：`1234`）→ 貼入 GAS 網址 → 點擊「**儲存並測試連線**」並切換為「**多機即時連線模式**」。

---

## 三、前端部署方式（需 HTTPS 才支援 Web NFC）

本系統無須 Node.js 編譯，直接上傳整個資料夾即可：

- **GitHub Pages**：推送至 GitHub 儲存庫 → Settings → Pages → 選擇 `main` 分支根目錄 → 取得 `https://<帳號>.github.io/<repo>/`。
- **Cloudflare Pages / Vercel**：直接匯入儲存庫或拖曳資料夾即可自動配發 HTTPS 網址。

---

## 四、Web NFC 操作與手機注意事項

1. **硬體與瀏覽器限制**：Web NFC API 目前僅支援 **Android 手機上的 Google Chrome（v89 以上）**，且網頁必須透過 **HTTPS**（或本機 `localhost`）開啟。iPhone (iOS Safari) 與桌機無法讀取實體 NFC 晶片，但可使用內建的「**模擬 NFC 感應測試器**」與「**手動補登**」。
2. **手機端事前設定**：
   - 開啟 Android「設定 → 連線裝置 → NFC」並確認開啟。
   - 建議將網頁「加入主畫面」以 PWA 全螢幕模式操作，避免誤觸瀏覽器返回鍵。
   - 賽事現場建議將手機螢幕休眠時間調長（或保持充電），並開啟頁面上方的「**強光高對比**」與「**音效開**」。
   - 確認各台計時手機皆開啟「自動設定日期與時間（網路校時）」，確保多機時間一致。
3. **建議使用之 NFC 晶片**：支援 NDEF 格式之 **NTAG213 / NTAG215 / NTAG216** 抗金屬貼紙或矽膠手環（容量需 ≥ 144 bytes）。
4. **公開查詢頁分享**：於後台點擊「**複製公開查詢頁分享連結**」，即可取得 `public.html?api=<GAS_URL>` 網址或做成 QR Code 張貼於現場供選手與親友掃碼查詢。
