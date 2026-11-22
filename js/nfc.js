/**
 * 智慧路跑 NFC 晶片計時系統 — Web NFC API + 音效/震動回饋 + 模擬器模組
 */
(function (global) {
  'use strict';

  /* ---------------- Web Audio API 提示音合成（零外部檔案、離線即響） ---------------- */
  class SoundFX {
    constructor() {
      this.ctx = null;
      this.enabled = true;
    }

    _ensureCtx() {
      if (!this.enabled) return null;
      const AudioCtx = global.AudioContext || global.webkitAudioContext;
      if (!AudioCtx) return null;
      if (!this.ctx) this.ctx = new AudioCtx();
      if (this.ctx.state === 'suspended') {
        this.ctx.resume().catch(() => {});
      }
      return this.ctx;
    }

    _tone(freq, durationMs, type = 'sine', gainVal = 0.25, delaySec = 0) {
      const ctx = this._ensureCtx();
      if (!ctx) return;
      const now = ctx.currentTime + delaySec;
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = type;
      osc.frequency.setValueAtTime(freq, now);
      gain.gain.setValueAtTime(gainVal, now);
      gain.gain.exponentialRampToValueAtTime(0.001, now + durationMs / 1000);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(now);
      osc.stop(now + durationMs / 1000 + 0.02);
    }

    /** 清脆成功感應雙音 */
    beepSuccess() {
      this._tone(1046.5, 90, 'sine', 0.28, 0);
      this._tone(1568.0, 140, 'sine', 0.30, 0.09);
      this.vibrate([80]);
    }

    /** 晶片寫入並驗證完成三連音 */
    beepWriteVerified() {
      this._tone(880, 80, 'triangle', 0.25, 0);
      this._tone(1174.6, 80, 'triangle', 0.25, 0.09);
      this._tone(1760, 180, 'triangle', 0.30, 0.18);
      this.vibrate([60, 50, 120]);
    }

    /** 防連點重複感應警示音 */
    beepWarning() {
      this._tone(520, 110, 'square', 0.18, 0);
      this._tone(420, 160, 'square', 0.18, 0.13);
      this.vibrate([120, 60, 120]);
    }

    /** 錯誤提示音 */
    beepError() {
      this._tone(260, 260, 'sawtooth', 0.22, 0);
      this.vibrate([220]);
    }

    /** 鳴槍起跑音效 */
    beepGun() {
      this._tone(180, 180, 'sawtooth', 0.35, 0);
      this._tone(900, 320, 'triangle', 0.32, 0.05);
      this.vibrate([200, 80, 200]);
    }

    vibrate(pattern) {
      if (navigator.vibrate) {
        try {
          navigator.vibrate(pattern);
        } catch {}
      }
    }
  }

  /* ---------------- Web NFC API 封裝 ---------------- */
  class NFCManager {
    constructor() {
      this.scanAbortController = null;
      this.isScanning = false;
      this.sound = new SoundFX();
    }

    /**
     * 檢查目前瀏覽器環境是否支援 Web NFC
     */
    getDiagnostics() {
      const isSecure = global.isSecureContext || location.hostname === 'localhost' || location.hostname === '127.0.0.1';
      const hasNDEF = 'NDEFReader' in global;
      const ua = navigator.userAgent || '';
      const isAndroid = /Android/i.test(ua);
      const isChrome = /Chrome/i.test(ua) && !/Edg|OPR|SamsungBrowser/i.test(ua);

      let statusMessage = '支援 Web NFC（可直接感應與寫入晶片）';
      if (!isSecure) {
        statusMessage = '需透過 HTTPS 安全連線才能啟用 Web NFC API（目前為非安全連線，可使用模擬測試功能）';
      } else if (!hasNDEF) {
        statusMessage = isAndroid
          ? '目前瀏覽器不支援 Web NFC，請改用 Android 版 Google Chrome (v89+)'
          : '偵測到非 Android Chrome 裝置（桌機或 iOS 不支援 Web NFC，可使用內建「模擬感應」完整測試）';
      }

      return {
        supported: isSecure && hasNDEF,
        isSecure,
        hasNDEF,
        isAndroid,
        isChrome,
        statusMessage
      };
    }

    /**
     * 將選手物件編碼為寫入 NFC NDEF 的精簡 JSON 字串
     */
    encodeRunnerPayload(runner) {
      return JSON.stringify({
        v: 1,
        bib: String(runner.bib || '').trim(),
        name: String(runner.name || '').trim(),
        id: String(runner.idNumber || '').trim().toUpperCase(),
        cat: String(runner.categoryId || '').trim()
      });
    }

    /**
     * 解析 NFC 標籤內容（同時支援 JSON 與管線分隔 BIB|NAME|ID|CAT）
     */
    decodeNdefMessage(message) {
      const decoder = new TextDecoder();
      for (const record of message.records || []) {
        let text = '';
        if (record.recordType === 'text' || record.recordType === 'mime' || record.recordType === 'unknown') {
          text = decoder.decode(record.data);
        }
        text = text.trim();
        if (!text) continue;

        // 1. JSON 格式
        if (text.startsWith('{')) {
          try {
            const obj = JSON.parse(text);
            if (obj.bib) {
              return {
                bib: String(obj.bib).trim(),
                name: String(obj.name || '').trim(),
                idNumber: String(obj.id || obj.idNumber || '').trim().toUpperCase(),
                categoryId: String(obj.cat || obj.categoryId || '').trim(),
                raw: text
              };
            }
          } catch {}
        }

        // 2. 分隔字串格式 BIB|NAME|ID|CAT
        if (text.includes('|')) {
          const [bib, name, idNumber, categoryId] = text.split('|');
          if (bib && bib.trim()) {
            return {
              bib: bib.trim(),
              name: (name || '').trim(),
              idNumber: (idNumber || '').trim().toUpperCase(),
              categoryId: (categoryId || '').trim(),
              raw: text
            };
          }
        }

        // 3. 純參賽編號
        return {
          bib: text,
          name: '',
          idNumber: '',
          categoryId: '',
          raw: text
        };
      }
      return null;
    }

    /**
     * 啟動賽道極速感應監聽
     */
    async startScanning(onTagRead, onError) {
      const diag = this.getDiagnostics();
      if (!diag.supported) {
        throw new Error(diag.statusMessage);
      }

      this.stopScanning();
      this.scanAbortController = new AbortController();
      const ndef = new global.NDEFReader();

      await ndef.scan({ signal: this.scanAbortController.signal });
      this.isScanning = true;

      ndef.onreadingerror = () => {
        this.sound.beepError();
        if (onError) onError(new Error('無法讀取此 NFC 晶片，請將晶片貼近手機感應區重試'));
      };

      ndef.onreading = (event) => {
        const parsed = this.decodeNdefMessage(event.message);
        if (!parsed) {
          this.sound.beepError();
          if (onError) onError(new Error('晶片內容為空或格式無法辨識'));
          return;
        }
        onTagRead({
          ...parsed,
          nfcUid: event.serialNumber || '',
          source: 'nfc'
        });
      };
    }

    stopScanning() {
      if (this.scanAbortController) {
        this.scanAbortController.abort();
        this.scanAbortController = null;
      }
      this.isScanning = false;
    }

    /**
     * 寫入 NFC 晶片並自動執行一次讀取驗證（防呆確保無誤）
     * @param {Object} runner { bib, name, idNumber, categoryId }
     * @param {Function} onStep 狀態回報 callback(stepText)
     */
    async writeAndVerify(runner, onStep) {
      const diag = this.getDiagnostics();
      if (!diag.supported) {
        throw new Error(diag.statusMessage);
      }

      // 寫入前先暫停賽道掃描避免資源衝突
      const wasScanning = this.isScanning;
      if (wasScanning) this.stopScanning();

      const payloadStr = this.encodeRunnerPayload(runner);
      const ndef = new global.NDEFReader();

      if (onStep) onStep('請將空白 NFC 貼紙 / 手環緊貼手機背面感應區...');

      const writeAbort = new AbortController();
      const writeTimer = setTimeout(() => writeAbort.abort(), 15000);
      try {
        await ndef.write(
          {
            records: [
              {
                recordType: 'text',
                lang: 'zh-TW',
                data: payloadStr
              }
            ]
          },
          { overwrite: true, signal: writeAbort.signal }
        );
      } finally {
        clearTimeout(writeTimer);
      }

      if (onStep) onStep('寫入完成！正在自動讀取驗證晶片內容，請暫勿移開...');

      // 自動讀取驗證（限時 8 秒）
      const verifyResult = await new Promise(async (resolve, reject) => {
        const verifyAbort = new AbortController();
        const timer = setTimeout(() => {
          verifyAbort.abort();
          reject(new Error('讀取驗證逾時，請保持晶片貼近手機感應區並重新驗證'));
        }, 8000);

        try {
          await ndef.scan({ signal: verifyAbort.signal });
          ndef.onreading = (event) => {
            clearTimeout(timer);
            verifyAbort.abort();
            const readData = this.decodeNdefMessage(event.message);
            if (!readData || readData.bib !== String(runner.bib).trim()) {
              reject(new Error(`驗證失敗：讀回編號 (${readData?.bib || '空'}) 與寫入 (${runner.bib}) 不符`));
              return;
            }
            resolve({
              verified: true,
              nfcUid: event.serialNumber || '',
              data: readData
            });
          };
          ndef.onreadingerror = () => {
            clearTimeout(timer);
            verifyAbort.abort();
            reject(new Error('驗證讀取發生錯誤，請確認晶片未受損'));
          };
        } catch (err) {
          clearTimeout(timer);
          reject(err);
        }
      });

      this.sound.beepWriteVerified();
      return verifyResult;
    }

    /**
     * 模擬 NFC 寫入與驗證（供無實體晶片或桌機測試使用）
     */
    async mockWriteAndVerify(runner, onStep) {
      if (onStep) onStep('[模擬模式] 正在模擬寫入 NDEF JSON 資料至虛擬晶片...');
      await new Promise((r) => setTimeout(r, 450));
      if (onStep) onStep('[模擬模式] 寫入完成，正在執行自動讀取驗證...');
      await new Promise((r) => setTimeout(r, 450));
      const payloadStr = this.encodeRunnerPayload(runner);
      const parsed = JSON.parse(payloadStr);
      this.sound.beepWriteVerified();
      return {
        verified: true,
        nfcUid: 'MOCK-' + Math.random().toString(16).slice(2, 10).toUpperCase(),
        data: {
          bib: parsed.bib,
          name: parsed.name,
          idNumber: parsed.id,
          categoryId: parsed.cat,
          raw: payloadStr
        }
      };
    }
  }

  global.NFCManager = new NFCManager();
})(window);
