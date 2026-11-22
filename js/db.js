/**
 * 智慧路跑 NFC 晶片計時系統 v2.0 — 本地 IndexedDB + LocalStorage 離線儲存層
 * 支援離線帳號快取、權限鎖定站點、完整 JSON 備份還原
 */
(function (global) {
  'use strict';

  const DB_NAME = 'SmartRaceTimingDB';
  const DB_VERSION = 1;
  const SETTINGS_KEY = 'smart_race_device_settings_v2';
  const SESSION_KEY = 'smart_race_auth_session_v2';

  function getTodayStr() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function createDefaultStaffAccounts() {
    const now = Date.now();
    return [
      { username: 'writer01', password: '1111', displayName: '報到晶片寫入組', role: 'writer', checkpointId: '', enabled: true, updatedAt: now },
      { username: 'start01', password: '1111', displayName: '起點感應組', role: 'scanner', checkpointId: 'START', enabled: true, updatedAt: now },
      { username: 'cp1', password: '1111', displayName: '補給站 1 感應組', role: 'scanner', checkpointId: 'CP1', enabled: true, updatedAt: now },
      { username: 'cp2', password: '1111', displayName: '折返點/補給站 2 組', role: 'scanner', checkpointId: 'CP2', enabled: true, updatedAt: now },
      { username: 'cp3', password: '1111', displayName: '補給站 3 感應組', role: 'scanner', checkpointId: 'CP3', enabled: true, updatedAt: now },
      { username: 'finish01', password: '1111', displayName: '終點衝線感應組', role: 'scanner', checkpointId: 'FINISH', enabled: true, updatedAt: now }
    ];
  }

  function createDefaultRaceConfig() {
    return {
      id: 'main',
      raceName: '2026 智慧城市馬拉松',
      raceDate: getTodayStr(),
      debounceSeconds: 30,
      adminUsername: 'admin',
      adminPassword: 'admin1234',
      rankBy: 'gun',
      categories: [
        { id: '42K', name: '全馬組 42K', distanceKm: 42.195, gunTimeMs: null, gunTimeIso: '', color: '#ef4444' },
        { id: '21K', name: '半馬組 21K', distanceKm: 21.0975, gunTimeMs: null, gunTimeIso: '', color: '#3b82f6' },
        { id: '10K', name: '健康組 10K', distanceKm: 10, gunTimeMs: null, gunTimeIso: '', color: '#10b981' }
      ],
      checkpoints: [
        { id: 'START', name: '起點感應門', order: 1, distanceKm: 0, type: 'start' },
        { id: 'CP1', name: '補給站 1 (5K)', order: 2, distanceKm: 5, type: 'checkpoint' },
        { id: 'CP2', name: '折返點 / 補給站 2 (10K)', order: 3, distanceKm: 10, type: 'checkpoint' },
        { id: 'CP3', name: '補給站 3 (21K)', order: 4, distanceKm: 21.1, type: 'checkpoint' },
        { id: 'FINISH', name: '終點衝線門', order: 5, distanceKm: 42.195, type: 'finish' }
      ],
      staffAccounts: createDefaultStaffAccounts(),
      updatedAt: Date.now()
    };
  }

  function generateDeviceId() {
    return 'DEV-' + Math.random().toString(36).slice(2, 8).toUpperCase();
  }

  function generateId(prefix = 'SCAN') {
    if (global.crypto && global.crypto.randomUUID) {
      return `${prefix}-${global.crypto.randomUUID()}`;
    }
    return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  }

  class RaceDB {
    constructor() {
      this.db = null;
    }

    async init() {
      if (this.db) return this.db;
      this.db = await new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('config')) {
            db.createObjectStore('config', { keyPath: 'id' });
          }
          if (!db.objectStoreNames.contains('runners')) {
            const store = db.createObjectStore('runners', { keyPath: 'bib' });
            store.createIndex('categoryId', 'categoryId', { unique: false });
            store.createIndex('synced', 'synced', { unique: false });
          }
          if (!db.objectStoreNames.contains('scans')) {
            const store = db.createObjectStore('scans', { keyPath: 'id' });
            store.createIndex('bib', 'bib', { unique: false });
            store.createIndex('checkpointId', 'checkpointId', { unique: false });
            store.createIndex('synced', 'synced', { unique: false });
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });

      const existing = await this.getConfig();
      if (!existing) {
        await this.saveConfig(createDefaultRaceConfig());
      } else {
        // 向下相容舊版設定，補齊管理員帳密與預設作業員帳號
        let changed = false;
        if (!existing.adminUsername) {
          existing.adminUsername = 'admin';
          changed = true;
        }
        if (!existing.adminPassword) {
          existing.adminPassword = existing.adminPin || 'admin1234';
          changed = true;
        }
        if (!Array.isArray(existing.staffAccounts) || existing.staffAccounts.length === 0) {
          existing.staffAccounts = createDefaultStaffAccounts();
          changed = true;
        }
        if (changed) {
          await this.saveConfig(existing);
        }
      }
      return this.db;
    }

    _tx(storeName, mode = 'readonly') {
      return this.db.transaction(storeName, mode).objectStore(storeName);
    }

    _req(request) {
      return new Promise((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    }

    /* ---------------- 裝置設定與登入階段 (LocalStorage / SessionStorage) ---------------- */

    getDeviceSettings() {
      const defaults = {
        mode: 'offline', // 'offline' | 'online'
        gasUrl: '',
        deviceId: generateDeviceId(),
        highContrast: false,
        soundEnabled: true,
        autoSync: true,
        lastSyncAt: null
      };
      try {
        const raw = localStorage.getItem(SETTINGS_KEY) || localStorage.getItem('smart_race_device_settings_v1');
        if (!raw) {
          localStorage.setItem(SETTINGS_KEY, JSON.stringify(defaults));
          return defaults;
        }
        return { ...defaults, ...JSON.parse(raw) };
      } catch {
        return defaults;
      }
    }

    saveDeviceSettings(patch) {
      const merged = { ...this.getDeviceSettings(), ...patch };
      try {
        localStorage.setItem(SETTINGS_KEY, JSON.stringify(merged));
      } catch (e) {
        console.warn('LocalStorage 寫入失敗:', e);
      }
      return merged;
    }

    getSession() {
      try {
        const raw = sessionStorage.getItem(SESSION_KEY);
        return raw ? JSON.parse(raw) : null;
      } catch {
        return null;
      }
    }

    saveSession(sessionObj) {
      try {
        if (!sessionObj) {
          sessionStorage.removeItem(SESSION_KEY);
        } else {
          sessionStorage.setItem(SESSION_KEY, JSON.stringify(sessionObj));
        }
      } catch {}
    }

    /* ---------------- 賽事設定與帳號驗證 (Config) ---------------- */

    async getConfig() {
      await this.init();
      return this._req(this._tx('config').get('main'));
    }

    async saveConfig(cfg) {
      await this.init();
      const record = { ...cfg, id: 'main', updatedAt: Date.now() };
      await this._req(this._tx('config', 'readwrite').put(record));
      return record;
    }

    /**
     * 本機驗證登入（無網路離線時使用）
     * @param {'staff'|'admin'} mode
     */
    async verifyLocalLogin(mode, username, password) {
      const cfg = await this.getConfig();
      const u = String(username || '').trim();
      const p = String(password || '');

      if (mode === 'admin') {
        const expectedUser = String(cfg.adminUsername || 'admin').trim();
        const expectedPass = String(cfg.adminPassword || 'admin1234');
        if (u === expectedUser && p === expectedPass) {
          return {
            ok: true,
            account: { username: expectedUser, password: expectedPass, displayName: '最高系統管理員', role: 'admin' }
          };
        }
        return { ok: false, error: '管理員帳號或密碼錯誤' };
      }

      const list = cfg.staffAccounts || [];
      const matched = list.find((s) => s.enabled !== false && s.username === u && s.password === p);
      if (!matched) {
        return { ok: false, error: '作業員帳號或密碼錯誤（若為新帳號請先在有網路狀態登入同步一次）' };
      }
      return { ok: true, account: matched };
    }

    /**
     * 將線上驗證成功的單一作業員帳號快取至本機，確保後續無網路也能登入同一站點
     */
    async cacheStaffAccount(account) {
      if (!account || !account.username) return;
      const cfg = await this.getConfig();
      const list = [...(cfg.staffAccounts || [])];
      const idx = list.findIndex((s) => s.username === account.username);
      if (idx >= 0) {
        list[idx] = { ...list[idx], ...account };
      } else {
        list.push(account);
      }
      await this.saveConfig({ ...cfg, staffAccounts: list });
    }

    /* ---------------- 選手資料 (Runners) ---------------- */

    async getAllRunners() {
      await this.init();
      return this._req(this._tx('runners').getAll());
    }

    async getRunner(bib) {
      await this.init();
      return this._req(this._tx('runners').get(String(bib).trim()));
    }

    async findRunnerByBibOrId(query) {
      const q = String(query || '').trim().toUpperCase();
      if (!q) return null;
      const all = await this.getAllRunners();
      return (
        all.find((r) => String(r.bib).trim().toUpperCase() === q) ||
        all.find((r) => String(r.idNumber || '').trim().toUpperCase() === q) ||
        null
      );
    }

    async upsertRunner(runner, markSynced = false) {
      await this.init();
      const bib = String(runner.bib || '').trim();
      if (!bib) throw new Error('參賽編號不可為空');
      const existing = await this.getRunner(bib);
      const idNumber = String(runner.idNumber ?? existing?.idNumber ?? '').trim().toUpperCase();
      const idHash = idNumber ? await global.TimingEngine.sha256Hex(idNumber) : (runner.idHash || existing?.idHash || '');
      const record = {
        bib,
        name: String(runner.name ?? existing?.name ?? '').trim(),
        idNumber,
        idMasked: runner.idMasked || global.TimingEngine.maskIdNumber(idNumber),
        idHash,
        categoryId: String(runner.categoryId ?? existing?.categoryId ?? '').trim(),
        nfcUid: String(runner.nfcUid ?? existing?.nfcUid ?? ''),
        penaltySec: Number(runner.penaltySec ?? existing?.penaltySec ?? 0),
        checkedInAt: runner.checkedInAt ?? existing?.checkedInAt ?? null,
        updatedAt: Date.now(),
        synced: markSynced ? 1 : 0
      };
      await this._req(this._tx('runners', 'readwrite').put(record));
      return record;
    }

    async upsertRunnersBulk(list, markSynced = false) {
      const results = [];
      for (const r of list) {
        if (!r || !r.bib) continue;
        results.push(await this.upsertRunner(r, markSynced));
      }
      return results;
    }

    async deleteRunner(bib) {
      await this.init();
      await this._req(this._tx('runners', 'readwrite').delete(String(bib).trim()));
    }

    /* ---------------- 感應紀錄 (Scans) ---------------- */

    async getAllScans() {
      await this.init();
      return this._req(this._tx('scans').getAll());
    }

    async getPendingSyncCounts() {
      const [runners, scans] = await Promise.all([this.getAllRunners(), this.getAllScans()]);
      const unsyncedRunners = runners.filter((r) => !r.synced);
      const unsyncedScans = scans.filter((s) => !s.synced);
      return {
        runners: unsyncedRunners.length,
        scans: unsyncedScans.length,
        total: unsyncedRunners.length + unsyncedScans.length,
        unsyncedRunners,
        unsyncedScans
      };
    }

    async addScan({ bib, checkpointId, timestampMs = Date.now(), source = 'nfc', operator = '', note = '', bypassDebounce = false }) {
      await this.init();
      const cleanBib = String(bib || '').trim();
      const cleanCp = String(checkpointId || '').trim();
      if (!cleanBib || !cleanCp) {
        throw new Error('缺少參賽編號或感應站點');
      }

      const config = await this.getConfig();
      const debounceSec = Number(config?.debounceSeconds ?? 30);

      if (!bypassDebounce && debounceSec > 0) {
        const allScans = await this.getAllScans();
        const recent = allScans
          .filter((s) => !s.deleted && s.bib === cleanBib && s.checkpointId === cleanCp)
          .sort((a, b) => b.timestampMs - a.timestampMs)[0];

        if (recent) {
          const diffMs = Math.abs(timestampMs - recent.timestampMs);
          if (diffMs < debounceSec * 1000) {
            return {
              ok: false,
              duplicate: true,
              elapsedSec: Math.round(diffMs / 1000),
              waitRemainingSec: Math.ceil((debounceSec * 1000 - diffMs) / 1000),
              previousScan: recent
            };
          }
        }
      }

      const settings = this.getDeviceSettings();
      const scan = {
        id: generateId('SCAN'),
        bib: cleanBib,
        checkpointId: cleanCp,
        timestampMs: Number(timestampMs),
        timeStr: global.TimingEngine.formatClockMs(timestampMs),
        isoTime: new Date(Number(timestampMs)).toISOString(),
        source,
        deviceId: settings.deviceId,
        operator: String(operator || ''),
        deleted: false,
        note: String(note || ''),
        synced: 0,
        updatedAt: Date.now()
      };

      await this._req(this._tx('scans', 'readwrite').put(scan));
      return { ok: true, scan };
    }

    async putScan(scan) {
      await this.init();
      const record = {
        ...scan,
        timeStr: global.TimingEngine.formatClockMs(scan.timestampMs),
        isoTime: new Date(Number(scan.timestampMs)).toISOString(),
        updatedAt: Date.now()
      };
      await this._req(this._tx('scans', 'readwrite').put(record));
      return record;
    }

    async markScansSynced(ids) {
      await this.init();
      const idSet = new Set(ids);
      const all = await this.getAllScans();
      const tx = this.db.transaction('scans', 'readwrite');
      const store = tx.objectStore('scans');
      for (const s of all) {
        if (idSet.has(s.id) && !s.synced) {
          s.synced = 1;
          store.put(s);
        }
      }
      return new Promise((resolve, reject) => {
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    }

    async markRunnersSynced(bibs) {
      await this.init();
      const bibSet = new Set(bibs);
      const all = await this.getAllRunners();
      const tx = this.db.transaction('runners', 'readwrite');
      const store = tx.objectStore('runners');
      for (const r of all) {
        if (bibSet.has(r.bib) && !r.synced) {
          r.synced = 1;
          store.put(r);
        }
      }
      return new Promise((resolve, reject) => {
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    }

    async mergeCloudData(cloudData) {
      await this.init();
      if (!cloudData) return;

      const currentCfg = await this.getConfig();
      const mergedCfg = {
        ...currentCfg,
        raceName: cloudData.config?.raceName || currentCfg.raceName,
        raceDate: cloudData.config?.raceDate || currentCfg.raceDate,
        debounceSeconds: cloudData.config?.debounceSeconds ?? currentCfg.debounceSeconds,
        rankBy: cloudData.config?.rankBy || currentCfg.rankBy,
        adminUsername: cloudData.config?.adminUsername || currentCfg.adminUsername,
        adminPassword: cloudData.config?.adminPassword || currentCfg.adminPassword,
        categories: cloudData.categories?.length ? cloudData.categories : currentCfg.categories,
        checkpoints: cloudData.checkpoints?.length ? cloudData.checkpoints : currentCfg.checkpoints,
        staffAccounts: Array.isArray(cloudData.staffAccounts) && cloudData.staffAccounts.length
          ? cloudData.staffAccounts
          : currentCfg.staffAccounts
      };
      await this.saveConfig(mergedCfg);

      if (Array.isArray(cloudData.runners)) {
        for (const r of cloudData.runners) {
          const local = await this.getRunner(r.bib);
          if (local && !local.synced && local.updatedAt > (r.updatedAt || 0)) {
            continue;
          }
          await this.upsertRunner(
            {
              ...r,
              idNumber: r.idNumber || local?.idNumber || ''
            },
            true
          );
        }
      }

      if (Array.isArray(cloudData.scans)) {
        const localScans = await this.getAllScans();
        const localById = new Map(localScans.map((s) => [s.id, s]));
        for (const s of cloudData.scans) {
          const local = localById.get(s.id);
          if (local && !local.synced && local.updatedAt > (s.updatedAt || 0)) {
            continue;
          }
          await this.putScan({ ...s, synced: 1 });
        }
      }
    }

    /* ---------------- 完整備份匯出 / 還原匯入 (JSON) ---------------- */

    async exportFullBackup() {
      const [config, runners, scans] = await Promise.all([
        this.getConfig(),
        this.getAllRunners(),
        this.getAllScans()
      ]);
      return {
        schemaVersion: '2.0',
        exportedAt: new Date().toISOString(),
        config,
        runners,
        scans
      };
    }

    /**
     * 還原或合併 JSON 備份檔
     * @param {Object} backupObj
     * @param {'replace'|'merge'} strategy replace=完全覆蓋還原, merge=合併感應與選手紀錄（適用彙整多支離線手機）
     */
    async importFullBackup(backupObj, strategy = 'replace') {
      if (!backupObj || (!backupObj.config && !Array.isArray(backupObj.runners) && !Array.isArray(backupObj.scans))) {
        throw new Error('無效的系統備份檔格式');
      }

      if (strategy === 'replace') {
        await this.clearAllData(false);
        if (backupObj.config) {
          await this.saveConfig({ ...createDefaultRaceConfig(), ...backupObj.config });
        }
      } else if (backupObj.config) {
        const cur = await this.getConfig();
        await this.saveConfig({
          ...cur,
          ...backupObj.config,
          categories: backupObj.config.categories || cur.categories,
          checkpoints: backupObj.config.checkpoints || cur.checkpoints,
          staffAccounts: backupObj.config.staffAccounts || cur.staffAccounts
        });
      }

      let runnerCount = 0;
      if (Array.isArray(backupObj.runners)) {
        for (const r of backupObj.runners) {
          if (!r || !r.bib) continue;
          await this.upsertRunner(r, false);
          runnerCount++;
        }
      }

      let scanCount = 0;
      if (Array.isArray(backupObj.scans)) {
        for (const s of backupObj.scans) {
          if (!s || !s.id || !s.bib) continue;
          await this.putScan({ ...s, synced: 0 });
          scanCount++;
        }
      }

      return { runnerCount, scanCount };
    }

    async loadDemoData() {
      await this.init();
      const prevCfg = await this.getConfig();
      await this.clearAllData(false);

      const now = Date.now();
      const gun42 = now - 3 * 3600 * 1000 - 45 * 60 * 1000;
      const gun21 = now - 2 * 3600 * 1000 - 10 * 60 * 1000;
      const gun10 = now - 1 * 3600 * 1000 - 15 * 60 * 1000;

      const cfg = createDefaultRaceConfig();
      // 保留使用者已設定的管理員帳密與作業員帳號
      cfg.adminUsername = prevCfg?.adminUsername || 'admin';
      cfg.adminPassword = prevCfg?.adminPassword || 'admin1234';
      if (prevCfg?.staffAccounts?.length) cfg.staffAccounts = prevCfg.staffAccounts;

      cfg.categories[0].gunTimeMs = gun42;
      cfg.categories[0].gunTimeIso = new Date(gun42).toISOString();
      cfg.categories[1].gunTimeMs = gun21;
      cfg.categories[1].gunTimeIso = new Date(gun21).toISOString();
      cfg.categories[2].gunTimeMs = gun10;
      cfg.categories[2].gunTimeIso = new Date(gun10).toISOString();
      await this.saveConfig(cfg);

      const demoRunners = [
        { bib: 'A001', name: '林志明', idNumber: 'A123456789', categoryId: '42K', checkedInAt: gun42 - 1800000 },
        { bib: 'A002', name: '陳怡君', idNumber: 'B234567891', categoryId: '42K', checkedInAt: gun42 - 1700000 },
        { bib: 'A003', name: '張家豪', idNumber: 'C145678912', categoryId: '42K', checkedInAt: gun42 - 1600000 },
        { bib: 'B101', name: '王雅婷', idNumber: 'F256789123', categoryId: '21K', checkedInAt: gun21 - 1500000 },
        { bib: 'B102', name: '李俊傑', idNumber: 'H167891234', categoryId: '21K', checkedInAt: gun21 - 1400000 },
        { bib: 'B103', name: '黃淑芬', idNumber: 'E278912345', categoryId: '21K', checkedInAt: gun21 - 1300000 },
        { bib: 'C201', name: '吳冠宇', idNumber: 'T189123456', categoryId: '10K', checkedInAt: gun10 - 1200000 },
        { bib: 'C202', name: '劉佳穎', idNumber: 'N291234567', categoryId: '10K', checkedInAt: gun10 - 1100000 },
        { bib: 'C203', name: '蔡宗翰', idNumber: 'K112345678', categoryId: '10K', checkedInAt: gun10 - 1000000 }
      ];
      await this.upsertRunnersBulk(demoRunners, false);

      const addDemoScan = async (bib, cpId, tsMs) => {
        await this.addScan({ bib, checkpointId: cpId, timestampMs: tsMs, source: 'mock', operator: 'demo', bypassDebounce: true });
      };

      const min = 60 * 1000;
      await addDemoScan('A001', 'START', gun42 + 8200);
      await addDemoScan('A001', 'CP1', gun42 + 23 * min + 15200);
      await addDemoScan('A001', 'CP2', gun42 + 47 * min + 5400);
      await addDemoScan('A001', 'CP3', gun42 + 99 * min + 41000);
      await addDemoScan('A001', 'FINISH', gun42 + 202 * min + 18450);

      await addDemoScan('A002', 'START', gun42 + 14500);
      await addDemoScan('A002', 'CP1', gun42 + 25 * min + 10100);
      await addDemoScan('A002', 'CP2', gun42 + 51 * min + 32000);
      await addDemoScan('A002', 'CP3', gun42 + 108 * min + 22000);
      await addDemoScan('A002', 'FINISH', gun42 + 218 * min + 42310);

      await addDemoScan('A003', 'START', gun42 + 21000);
      await addDemoScan('A003', 'CP1', gun42 + 29 * min + 5000);
      await addDemoScan('A003', 'CP2', gun42 + 60 * min + 12000);
      await addDemoScan('A003', 'CP3', gun42 + 131 * min + 44000);

      await addDemoScan('B101', 'START', gun21 + 6100);
      await addDemoScan('B101', 'CP1', gun21 + 24 * min + 1200);
      await addDemoScan('B101', 'CP2', gun21 + 48 * min + 21500);
      await addDemoScan('B101', 'FINISH', gun21 + 101 * min + 9320);

      await addDemoScan('B102', 'START', gun21 + 18900);
      await addDemoScan('B102', 'CP1', gun21 + 26 * min + 35000);
      await addDemoScan('B102', 'CP2', gun21 + 54 * min + 10000);
      await addDemoScan('B102', 'FINISH', gun21 + 114 * min + 26780);

      await addDemoScan('B103', 'START', gun21 + 25000);
      await addDemoScan('B103', 'CP1', gun21 + 31 * min + 14000);
      await addDemoScan('B103', 'CP2', gun21 + 64 * min + 8000);

      await addDemoScan('C201', 'START', gun10 + 4200);
      await addDemoScan('C201', 'CP1', gun10 + 23 * min + 11000);
      await addDemoScan('C201', 'FINISH', gun10 + 46 * min + 55120);

      await addDemoScan('C202', 'START', gun10 + 12500);
      await addDemoScan('C202', 'CP1', gun10 + 29 * min + 40200);

      return true;
    }

    async clearAllData(resetConfig = true) {
      await this.init();
      await Promise.all([
        this._req(this._tx('runners', 'readwrite').clear()),
        this._req(this._tx('scans', 'readwrite').clear())
      ]);
      if (resetConfig) {
        await this.saveConfig(createDefaultRaceConfig());
      }
    }
  }

  global.RaceDB = new RaceDB();
})(window);
