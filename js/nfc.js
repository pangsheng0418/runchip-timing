/**
 * 智慧路跑 NFC 晶片計時系統 v3.0 — 共用晶片 Web NFC 讀寫與 UID 雙模核心
 * 晶片寫入格式（跨場次共用）：僅寫入「姓名|身份證字號」（約 22 bytes）
 * 優點：
 * 1. 選手持同一張晶片即可參加不同活動與不同組別，不需每次活動重新寫卡
 * 2. 體積極小，連小容量 NTAG203 / Ultralight / NTAG213/215/216 均可穩定寫入
 * 3. 亦相容「僅綁定晶片 UID」模式（唯讀貼紙/手環/MIFARE 卡亦可使用）
 */
(function (global) {
  'use strict';

  /* ---------------- Web Audio API 提示音合成 ---------------- */
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
    const name = err.name || '';
    const msg = err.message || String(err);

    if (name === 'NotAllowedError') {
      return 'NFC 權限遭拒或未開啟：請點選 Chrome 網址列左側圖示允許「NFC 裝置」權限，並確認手機系統已開啟 NFC 開關。';
    }
    if (name === 'NotSupportedError') {
      return '此晶片不支援 NDEF 寫入（可能為唯讀卡、MIFARE Classic 或未格式化晶片）。建議改點「🔗 僅綁定晶片 UID」免寫入直接綁定，或更換 NTAG213/215/216 晶片。';
    }
    if (name === 'NetworkError' || name === 'IOError') {
      return '晶片通訊中斷或容量不足：寫入時請將晶片緊貼手機背面 NFC 天線並保持靜止 1～2 秒（勿過早移開）；若晶片為唯讀請改用「UID 綁定模式」。';
    }
    if (name === 'NotReadableError') {
      return '無法讀取此 NFC 晶片（晶片可能損壞、已鎖定或距離過遠），請重新緊貼感應區。';
    }
    if (name === 'AbortError') {
      return '已取消或等待感應逾時，請重新點擊按鈕後再靠近晶片。';
    }
    if (name === 'InvalidStateError') {
      return '手機 NFC 正在執行其他感應作業中，請稍候 1 秒再試。';
    }
    return `${msg}${name ? ` (${name})` : ''}`;
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

      let statusMessage = '✅ 支援 Web NFC（共用晶片格式：僅寫入「姓名｜身份證字號」，跨場次通用）';
      if (!isSecure) {
        statusMessage = '❌ 目前非 HTTPS 安全連線，瀏覽器會停用 Web NFC（請使用 HTTPS 網址開啟）';
      } else if (!hasNDEF) {
        statusMessage = isAndroid
          ? '⚠️ 目前瀏覽器不支援 Web NFC（請改用 Android 版 Google Chrome v89+，勿用 LINE/FB 內建瀏覽器）'
          : '⚠️ 目前為非 Android Chrome 裝置（電腦或 iPhone 不支援 Web NFC，可使用內建「模擬感應/寫入」測試）';
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
     * 編碼共用晶片內容：僅包含「姓名|身份證字號」（約 22 bytes，跨場次重複使用）
     */
    encodeChipPayload({ name, idNumber }) {
      const cleanName = String(name || '').replace(/\|/g, ' ').trim();
      const cleanId = String(idNumber || '').replace(/\|/g, '').trim().toUpperCase();
      return `${cleanName}|${cleanId}`;
    }

    /**
     * 解析 NFC 標籤內容：
     * - v3.0 預設 2 欄格式：NAME|IDNUMBER（例如：林志明|A123456789）
     * - 亦向下相容 JSON {"name":"...","id":"..."} 與舊版 4 欄格式 BIB|NAME|ID|CAT
     */
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

        // 1. JSON 格式相容
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

        // 2. 管線分隔格式
        if (text.includes('|')) {
          const parts = text.split('|').map((s) => s.trim());
          // v3.0 標準格式：姓名|身份證字號 (2 欄)
          if (parts.length === 2) {
            return {
              name: parts[0],
              idNumber: parts[1].toUpperCase(),
              bib: '',
              raw: text,
              recordType: record.recordType
            };
          }
          // 向下相容舊版 4 欄格式：BIB|NAME|ID|CAT
          if (parts.length >= 3) {
            return {
              bib: parts[0],
              name: parts[1],
              idNumber: parts[2].toUpperCase(),
              raw: text,
              recordType: record.recordType
            };
          }
        }

        // 3. 純文字（視為身份證字號或姓名）
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
          onTagRead({
            ...parsed,
            nfcUid: uid,
            uidOnly: false,
            source: 'nfc'
          });
          return;
        }

        if (uid) {
          onTagRead({
            name: '',
            idNumber: '',
            bib: '',
            raw: '',
            nfcUid: uid,
            uidOnly: true,
            source: 'nfc'
          });
          return;
        }

        this.sound.beepError();
        if (onError) onError(new Error('感應到空白晶片且無可辨識 UID，請先進行晶片寫入或 UID 綁定'));
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

      return new Promise(async (resolve, reject) => {
        const timer = setTimeout(() => {
          ac.abort();
          this.isBusy = false;
          reject(new Error('等待讀取逾時（25 秒），請重新點擊並將晶片緊貼手機背面 NFC 天線。'));
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
          reject(new Error('讀取晶片失敗：請將晶片貼緊手機背面 NFC 天線停留 1 秒，或確認此晶片非加密門禁卡。'));
        }, { once: true });

        try {
          await ndef.scan({ signal: ac.signal });
        } catch (err) {
          cleanup();
          reject(new Error(translateNfcError(err)));
        }
      });
    }

    /**
     * 寫入共用晶片（僅寫入「姓名|身份證字號」）並自動引導回讀驗證
     */
    async writeChipAndVerify({ name, idNumber }, { onStep, onWriteDone, skipVerify = false } = {}) {
      const diag = this.getDiagnostics();
      if (!diag.supported) throw new Error(diag.statusMessage);

      this.stopScanning();
      this.cancelActiveOperation();

      const ac = new AbortController();
      this.opAbortController = ac;
      this.isBusy = true;

      const targetId = String(idNumber || '').trim().toUpperCase();
      const targetName = String(name || '').trim();
      const payloadStr = this.encodeChipPayload({ name: targetName, idNumber: targetId });
      const ndef = new global.NDEFReader();
      this.reader = ndef;

      const writeTimer = setTimeout(() => ac.abort(), 30000);

      try {
        let writeCompleted = false;
        let verifyResolve = null;
        let verifyReject = null;

        const verifyPromise = new Promise((res, rej) => {
          verifyResolve = res;
          verifyReject = rej;
        });

        if (!skipVerify) {
          ndef.addEventListener('reading', (event) => {
            if (!writeCompleted) return;
            const uid = String(event.serialNumber || '').trim().toUpperCase();
            const readData = this.decodeNdefMessage(event.message);
            if (!readData || readData.idNumber !== targetId) {
              verifyReject(new Error(`回讀驗證不符：讀回身份證「${readData?.idNumber || '空'}」，預期「${targetId}」`));
              return;
            }
            verifyResolve({ verified: true, nfcUid: uid, data: readData });
          });

          ndef.addEventListener('readingerror', () => {
            if (!writeCompleted) return;
            verifyReject(new Error('回讀驗證感應失敗，請再重新貼近一次'));
          });

          try {
            await ndef.scan({ signal: ac.signal });
          } catch (scanErr) {
            console.warn('預啟 scan 略過:', scanErr);
          }
        }

        if (onStep) {
          onStep(`📡 請將 NFC 晶片緊貼手機背面感應區並停留 1～2 秒...（寫入內容：${payloadStr}）`);
        }

        await ndef.write(payloadStr, { overwrite: true, signal: ac.signal });
        clearTimeout(writeTimer);
        writeCompleted = true;

        this.sound.beepSuccess();
        if (onWriteDone) await onWriteDone({ written: true, raw: payloadStr });

        if (skipVerify) {
          ac.abort();
          this.opAbortController = null;
          this.reader = null;
          this.isBusy = false;
          this.sound.beepWriteVerified();
          return { verified: false, written: true, nfcUid: '', data: { name: targetName, idNumber: targetId, raw: payloadStr } };
        }

        if (onStep) {
          onStep('✅ 晶片已寫入成功！請將晶片【移開手機 1 秒後，再重新貼近一次】進行自動讀取驗證...');
        }

        const verifyTimer = setTimeout(() => {
          verifyResolve({
            verified: false,
            written: true,
            verifyTimeout: true,
            nfcUid: '',
            data: { name: targetName, idNumber: targetId, raw: payloadStr }
          });
        }, 20000);

        const onVerifyAbort = () => {
          clearTimeout(verifyTimer);
          verifyResolve({
            verified: false,
            written: true,
            verifySkipped: true,
            nfcUid: '',
            data: { name: targetName, idNumber: targetId, raw: payloadStr }
          });
        };
        ac.signal.addEventListener('abort', onVerifyAbort, { once: true });

        try {
          const vRes = await verifyPromise;
          clearTimeout(verifyTimer);
          ac.signal.removeEventListener('abort', onVerifyAbort);
          if (!ac.signal.aborted) ac.abort();
          this.opAbortController = null;
          this.reader = null;
          this.isBusy = false;
          this.sound.beepWriteVerified();
          return { ...vRes, written: true };
        } catch (vErr) {
          clearTimeout(verifyTimer);
          ac.signal.removeEventListener('abort', onVerifyAbort);
          if (!ac.signal.aborted) ac.abort();
          this.opAbortController = null;
          this.reader = null;
          this.isBusy = false;
          throw vErr;
        }
      } catch (err) {
        clearTimeout(writeTimer);
        if (!ac.signal.aborted) ac.abort();
        this.opAbortController = null;
        this.reader = null;
        this.isBusy = false;
        throw new Error(translateNfcError(err));
      }
    }

    async mockWriteChipAndVerify({ name, idNumber }, { onStep } = {}) {
      const targetId = String(idNumber || '').trim().toUpperCase();
      const targetName = String(name || '').trim();
      const payloadStr = this.encodeChipPayload({ name: targetName, idNumber: targetId });
      if (onStep) onStep(`[模擬模式] 正在寫入共用晶片資料（${payloadStr}）...`);
      await new Promise((r) => setTimeout(r, 380));
      if (onStep) onStep('[模擬模式] 寫入完成，正在模擬移開重刷讀取驗證...');
      await new Promise((r) => setTimeout(r, 380));
      this.sound.beepWriteVerified();
      return {
        written: true,
        verified: true,
        nfcUid: 'MOCK-' + Math.random().toString(16).slice(2, 10).toUpperCase(),
        data: { name: targetName, idNumber: targetId, raw: payloadStr }
      };
    }
  }

  global.NFCManager = new NFCManager();
})(window);
