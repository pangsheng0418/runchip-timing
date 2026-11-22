/**
 * 智慧路跑 NFC 晶片計時平台 v3.1 — 共用晶片 Web NFC 讀寫與 UID 雙模核心
 * 修正重點 (v3.1)：
 * 1. 徹底分離寫入 (ndef.write) 與掃描 (ndef.scan)，消除 Android Chrome 同實例 scan+write 衝突與點擊手勢失效問題
 * 2. 點擊寫入鈕當下零前置 await，直接同步呼叫 ndef.write({ records: [{ recordType: 'text', encoding: 'utf-8', lang: 'zh', data }] }, { overwrite: true })
 * 3. 寫入內容僅存「姓名|身份證字號」（約 20 bytes），並於錯誤訊息附上瀏覽器原始錯誤碼以便精準排錯
 * 4. 支援獨立一步點擊「回讀驗證」與「僅綁定晶片 UID」（適用唯讀卡 / MIFARE Classic）
 */
(function (global) {
  'use strict';

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
      if (this.ctx.state === 'suspended') this.ctx.resume().catch(() => {});
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

    beepSuccess() {
      this._tone(1046.5, 90, 'sine', 0.28, 0);
      this._tone(1568.0, 140, 'sine', 0.30, 0.09);
      this.vibrate([80]);
    }

    beepWriteVerified() {
      this._tone(880, 80, 'triangle', 0.25, 0);
      this._tone(1174.6, 80, 'triangle', 0.25, 0.09);
      this._tone(1760, 180, 'triangle', 0.30, 0.18);
      this.vibrate([60, 50, 120]);
    }

    beepWarning() {
      this._tone(520, 110, 'square', 0.18, 0);
      this._tone(420, 160, 'square', 0.18, 0.13);
      this.vibrate([120, 60, 120]);
    }

    beepError() {
      this._tone(260, 260, 'sawtooth', 0.22, 0);
      this.vibrate([220]);
    }

    beepGun() {
      this._tone(180, 180, 'sawtooth', 0.35, 0);
      this._tone(900, 320, 'triangle', 0.32, 0.05);
      this.vibrate([200, 80, 200]);
    }

    vibrate(pattern) {
      if (navigator.vibrate) {
        try { navigator.vibrate(pattern); } catch {}
      }
    }
  }

  function translateNfcError(err) {
    if (!err) return '未知的 NFC 錯誤';
    const name = err.name || 'Error';
    const rawMsg = err.message || String(err);
    const rawTag = ` [原始訊息: ${name} - ${rawMsg}]`;

    if (name === 'NotAllowedError') {
      return 'NFC 權限遭拒或未開啟：請點選 Chrome 網址列左側圖示 → 權限 → 允許「NFC 裝置」，並確認手機系統 NFC 開關已開啟。' + rawTag;
    }
    if (name === 'NotSupportedError') {
      return '此晶片不支援標準 NDEF 寫入（可能為 MIFARE Classic、悠遊卡/門禁卡、未格式化或唯讀鎖定晶片）。請改點「🔗 僅綁定晶片 UID」免寫入直接綁定，或更換 NTAG213/215/216 晶片。' + rawTag;
    }
    if (name === 'NetworkError' || name === 'IOError') {
      return '晶片傳輸中斷：請將晶片緊貼手機背面 NFC 天線並保持靜止 1～2 秒（勿過早移開）；若反覆出現代表此晶片可能為唯讀或非 NTAG 晶片，請改用「🔗 僅綁定晶片 UID」。' + rawTag;
    }
    if (name === 'NotReadableError') {
      return '無法讀取此 NFC 晶片（距離過遠、移開太快或晶片損壞），請緊貼背面 NFC 天線停留 1 秒重試。' + rawTag;
    }
    if (name === 'AbortError') {
      return '已取消或等待感應逾時（30 秒內未感應到晶片），請重新點擊按鈕後再靠近晶片。';
    }
    if (name === 'InvalidStateError') {
      return '手機 NFC 正忙線中，請稍候 1 秒後重試。' + rawTag;
    }
    return rawMsg + rawTag;
  }

  class NFCManager {
    constructor() {
      this.reader = null;
      this.scanAbortController = null;
      this.opAbortController = null;
      this.isScanning = false;
      this.isBusy = false;
      this.sound = new SoundFX();
    }

    getDiagnostics() {
      const isSecure = global.isSecureContext || location.hostname === 'localhost' || location.hostname === '127.0.0.1';
      const hasNDEF = 'NDEFReader' in global;
      const ua = navigator.userAgent || '';
      const isAndroid = /Android/i.test(ua);
      const isChrome = /Chrome/i.test(ua) && !/Edg|OPR|SamsungBrowser/i.test(ua);

      let statusMessage = '✅ Web NFC 就緒 (v3.1) — 共用晶片僅寫入「姓名｜身份證字號」，跨場次通用';
      if (!isSecure) {
        statusMessage = '❌ 目前非 HTTPS 安全連線，瀏覽器會停用 Web NFC（請透過 GitHub Pages 等 HTTPS 網址開啟）';
      } else if (!hasNDEF) {
        statusMessage = isAndroid
          ? '⚠️ 目前瀏覽器不支援 Web NFC（請改用 Android 版 Google Chrome v89+，不可用 LINE/FB 內建瀏覽器）'
          : '⚠️ 目前為非 Android Chrome 裝置（電腦或 iPhone 不支援 Web NFC，請改用 Android Chrome 或使用模擬測試）';
      }
      return { supported: isSecure && hasNDEF, isSecure, hasNDEF, isAndroid, isChrome, statusMessage };
    }

    async queryPermissionState() {
      try {
        if (navigator.permissions && navigator.permissions.query) {
          const st = await navigator.permissions.query({ name: 'nfc' });
          return st.state;
        }
      } catch {}
      return 'unknown';
    }

    /**
     * 編碼共用晶片內容：僅「姓名|身份證字號」（約 20 bytes）
     */
    encodeChipPayload({ name, idNumber }) {
      const cleanName = String(name || '').replace(/\|/g, ' ').trim();
      const cleanId = String(idNumber || '').replace(/\|/g, '').trim().toUpperCase();
      return `${cleanName}|${cleanId}`;
    }

    decodeNdefMessage(message) {
      const records = (message && message.records) || [];
      for (const record of records) {
        if (!record || !record.data || record.data.byteLength === 0) continue;

        let text = '';
        try {
          text = new TextDecoder(record.encoding || 'utf-8').decode(record.data).trim();
        } catch {
          try {
            text = new TextDecoder('utf-8').decode(record.data).trim();
          } catch {
            continue;
          }
        }
        if (!text) continue;

        if (text.startsWith('{')) {
          try {
            const obj = JSON.parse(text);
            const idNumber = String(obj.id || obj.idNumber || '').trim().toUpperCase();
            const name = String(obj.name || '').trim();
            if (idNumber || name) {
              return { name, idNumber, bib: String(obj.bib || '').trim(), raw: text, recordType: record.recordType };
            }
          } catch {}
        }

        if (text.includes('|')) {
          const parts = text.split('|').map((s) => s.trim());
          if (parts.length === 2) {
            return { name: parts[0], idNumber: parts[1].toUpperCase(), bib: '', raw: text, recordType: record.recordType };
          }
          if (parts.length >= 3) {
            return { bib: parts[0], name: parts[1], idNumber: parts[2].toUpperCase(), raw: text, recordType: record.recordType };
          }
        }

        if (record.recordType === 'text' || record.recordType === 'mime' || record.recordType === 'unknown') {
          const looksLikeId = /^[A-Z][0-9A-Z]{6,12}$/i.test(text);
          return {
            name: looksLikeId ? '' : text,
            idNumber: looksLikeId ? text.toUpperCase() : '',
            bib: '',
            raw: text,
            recordType: record.recordType
          };
        }
      }
      return null;
    }

    cancelActiveOperation() {
      if (this.opAbortController) {
        this.opAbortController.abort();
        this.opAbortController = null;
      }
      this.isBusy = false;
    }

    /**
     * 賽道連續感應監聽（同時支援讀取「姓名|身份證」與「硬體 UID」）
     */
    async startScanning(onTagRead, onError) {
      const diag = this.getDiagnostics();
      if (!diag.supported) throw new Error(diag.statusMessage);

      this.cancelActiveOperation();
      this.stopScanning();

      this.scanAbortController = new AbortController();
      const ndef = new global.NDEFReader();
      this.reader = ndef;

      ndef.addEventListener('readingerror', () => {
        this.sound.beepError();
        if (onError) {
          onError(new Error('無法讀取此 NFC 晶片（移開太快或非標準 NDEF 晶片），請緊貼手機背面 NFC 天線 1 秒重試'));
        }
      });

      ndef.addEventListener('reading', (event) => {
        const uid = String(event.serialNumber || '').trim().toUpperCase();
        const parsed = this.decodeNdefMessage(event.message);

        if (parsed && (parsed.idNumber || parsed.name || parsed.bib)) {
          onTagRead({ ...parsed, nfcUid: uid, uidOnly: false, source: 'nfc' });
          return;
        }
        if (uid) {
          onTagRead({ name: '', idNumber: '', bib: '', raw: '', nfcUid: uid, uidOnly: true, source: 'nfc' });
          return;
        }
        this.sound.beepError();
        if (onError) onError(new Error('感應到空白晶片且無可辨識 UID，請先寫入晶片或綁定 UID'));
      });

      try {
        await ndef.scan({ signal: this.scanAbortController.signal });
        this.isScanning = true;
      } catch (err) {
        this.stopScanning();
        throw new Error(translateNfcError(err));
      }
    }

    stopScanning() {
      if (this.scanAbortController) {
        this.scanAbortController.abort();
        this.scanAbortController = null;
      }
      this.reader = null;
      this.isScanning = false;
    }

    /**
     * 純淨寫入共用晶片（僅寫入「姓名|身份證字號」）
     * 重要：零前置 await、零 scan() 並發衝突，使用獨立全新 NDEFReader 直接呼叫 write()！
     */
    async writeChip({ name, idNumber }, { onStep } = {}) {
      const diag = this.getDiagnostics();
      if (!diag.supported) throw new Error(diag.statusMessage);

      // 先同步關閉任何背景掃描與舊作業，避免 Android Chrome NFC 硬體資源衝突
      this.stopScanning();
      this.cancelActiveOperation();

      const targetId = String(idNumber || '').trim().toUpperCase();
      const targetName = String(name || '').trim();
      const payloadStr = this.encodeChipPayload({ name: targetName, idNumber: targetId });

      const ac = new AbortController();
      this.opAbortController = ac;
      this.isBusy = true;

      const ndef = new global.NDEFReader();
      this.reader = ndef;

      if (onStep) {
        onStep(`📡 請將 NFC 晶片緊貼手機背面感應區並靜止停留 1～2 秒...（寫入內容：${payloadStr}）`);
      }

      const timer = setTimeout(() => ac.abort(), 30000);

      try {
        // 同步進入 ndef.write()，完整保留使用者點擊手勢 (Transient User Activation)
        await ndef.write(
          {
            records: [
              {
                recordType: 'text',
                encoding: 'utf-8',
                lang: 'zh',
                data: payloadStr
              }
            ]
          },
          { overwrite: true, signal: ac.signal }
        );

        clearTimeout(timer);
        this.opAbortController = null;
        this.reader = null;
        this.isBusy = false;
        this.sound.beepWriteVerified();

        return {
          written: true,
          name: targetName,
          idNumber: targetId,
          raw: payloadStr
        };
      } catch (err) {
        clearTimeout(timer);
        this.opAbortController = null;
        this.reader = null;
        this.isBusy = false;
        const translated = new Error(translateNfcError(err));
        translated.originalName = err?.name || '';
        throw translated;
      }
    }

    /**
     * 獨立單次讀取晶片（用於「讀取驗證」、「測試晶片」與「UID 綁定」）
     * 全新獨立點擊手勢，絕無 scan/write 衝突！
     */
    async readSingleTag(onStep, timeoutMs = 25000) {
      const diag = this.getDiagnostics();
      if (!diag.supported) throw new Error(diag.statusMessage);

      this.stopScanning();
      this.cancelActiveOperation();

      const ac = new AbortController();
      this.opAbortController = ac;
      this.isBusy = true;

      if (onStep) onStep('📡 請將 NFC 晶片緊貼手機背面感應區（等待讀取中...）');

      const ndef = new global.NDEFReader();
      this.reader = ndef;

      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          cleanup();
          reject(new Error('等待讀取逾時（25 秒），請確認手機已開啟 NFC 並將晶片貼近背面感應天線。'));
        }, timeoutMs);

        const cleanup = () => {
          clearTimeout(timer);
          if (!ac.signal.aborted) ac.abort();
          this.opAbortController = null;
          this.reader = null;
          this.isBusy = false;
        };

        ac.signal.addEventListener('abort', () => {
          clearTimeout(timer);
          this.isBusy = false;
        });

        ndef.addEventListener('reading', (event) => {
          const uid = String(event.serialNumber || '').trim().toUpperCase();
          const parsed = this.decodeNdefMessage(event.message);
          const recordTypes = ((event.message && event.message.records) || []).map((r) => r.recordType);
          cleanup();
          this.sound.beepSuccess();
          resolve({ nfcUid: uid, parsed, recordTypes, raw: parsed ? parsed.raw : '' });
        }, { once: true });

        ndef.addEventListener('readingerror', () => {
          cleanup();
          this.sound.beepError();
          reject(new Error('讀取晶片失敗：請將晶片貼緊手機背面 NFC 天線停留 1 秒；若為加密門禁卡/悠遊卡則無法讀取 NDEF。'));
        }, { once: true });

        // 直接於點擊手勢內同步呼叫 scan()
        ndef.scan({ signal: ac.signal }).catch((err) => {
          cleanup();
          reject(new Error(translateNfcError(err)));
        });
      });
    }

    async mockWriteChip({ name, idNumber }, { onStep } = {}) {
      const targetId = String(idNumber || '').trim().toUpperCase();
      const targetName = String(name || '').trim();
      const payloadStr = this.encodeChipPayload({ name: targetName, idNumber: targetId });
      if (onStep) onStep(`[模擬模式] 正在寫入共用晶片（${payloadStr}）...`);
      await new Promise((r) => setTimeout(r, 380));
      this.sound.beepWriteVerified();
      return {
        written: true,
        verified: true,
        nfcUid: 'MOCK-' + Math.random().toString(16).slice(2, 10).toUpperCase(),
        name: targetName,
        idNumber: targetId,
        raw: payloadStr
      };
    }
  }

  global.NFCManager = new NFCManager();
})(window);
