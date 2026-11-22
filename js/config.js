/**
 * ============================================================================
 * 智慧路跑 NFC 多場次計時平台 — 全裝置共用雲端預設設定 (js/config.js)
 * ============================================================================
 * 強烈建議：
 * 將您的 Google Apps Script 網頁應用程式網址貼入下方 DEFAULT_GAS_URL 後上傳至 GitHub，
 * 如此一來，無論是主辦電腦、各站工作人員手機，還是民眾開啟公開查詢頁，
 * 進入系統只要有網路就會「自動連線並立即同步同一份雲端資料」，絕不產生各裝置資料不一致！
 *
 * 範例：
 * DEFAULT_GAS_URL: 'https://script.google.com/macros/s/AKfycbz7ALJAQ_cXYa2gRdITvXI2HlCeWBrsqhQGIcRtIL4nRCBskXIB3fiIRev5DspwPGEgzw/exec',
 * ============================================================================
 */
window.APP_CONFIG = {
  // 請將部署完成的 GAS 網頁應用程式網址貼在引號內（留空則由各裝置於後台或 ?api= 網址設定）
  DEFAULT_GAS_URL: 'https://script.google.com/macros/s/AKfycbz7ALJAQ_cXYa2gRdITvXI2HlCeWBrsqhQGIcRtIL4nRCBskXIB3fiIRev5DspwPGEgzw/exec',

  // 當設定了雲端網址且偵測到網路正常時，進入系統自動切換連線並同步雲端最新資料
  AUTO_SYNC_ON_START: true
};
