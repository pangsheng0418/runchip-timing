/**
 * 智慧路跑 NFC 晶片計時平台 v3.0 — 多場次共用晶片 IndexedDB 離線儲存層
 * Stores:
 * - platform      : 全域設定 (adminUsername, adminPassword, activeEventId, events[], staffAccounts[])
 * - chips         : 跨活動共用晶片主檔 (key: idNumber -> { idNumber, name, nfcUid, writtenAt, synced })
 * - registrations : 各活動事前匯入參賽名單 (key: regId = eventId__idNumberOrBib)
 * - scans         : 各活動感應紀錄 (key: id,含 eventId, bib, idNumber, checkpointId)
 */
(function (global) {
  'use strict';

  const DB_NAME = 'SmartRacePlatformDB_v3';
  const DB_VERSION = 1;
  const SETTINGS_KEY = 'smart_race_platform_settings_v3';
  const SESSION_KEY = 'smart_race_platform_session_v3';

  function getTodayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function createDefaultEvent(eventId = 'EVT-2026-01', name = '2026 智慧城市馬拉松', date = getTodayStr(), location = '台北市政府廣場', ownerAdmin = 'admin') {
    return {
      eventId,
      name,
      date,
      location,
      status: 'active',
      ownerAdmin: String(ownerAdmin || 'admin').trim() || 'admin',
      eventType: 'distance', // 'distance' (一般距離賽) | 'lap_timed' (繞圈計時賽)
      lapDistanceKm: 0,
      debounceSeconds: 30,
      rankBy: 'gun',
      unifiedGunTimeMs: null,
      categories: [
        { id: '42K', name: '全馬組 42K', distanceKm: 42.195, gunTimeMs: null, gunTimeIso: '', color: '#ef4444' },
        { id: '21K', name: '半馬組 21K', distanceKm: 21.0975, gunTimeMs: null, gunTimeIso: '', color: '#3b82f6' },
        { id: '10K', name: '健康組 10K', distanceKm: 10, gunTimeMs: null, gunTimeIso: '', color: '#10b981' }
      ],
      checkpoints: [
        { id: 'CP1', name: '補給站 1 (5K)', order: 1, distanceKm: 5, type: 'checkpoint' },
        { id: 'CP2', name: '折返點 / 補給站 2 (10K)', order: 2, distanceKm: 10, type: 'checkpoint' },
        { id: 'CP3', name: '補給站 3 (21K)', order: 3, distanceKm: 21.1, type: 'checkpoint' },
        { id: 'FINISH', name: '終點衝線門', order: 4, distanceKm: 42.195, type: 'finish' }
      ],
      updatedAt: Date.now()
    };
  }

  function createLapTimedEvent(eventId = 'EVT-2026-LAP', name = '2026 超馬繞圈計時賽', date = getTodayStr(), location = '大安森林公園環道', lapDistanceKm = 2.54, ownerAdmin = 'admin') {
    const lapKm = Number(lapDistanceKm || 2.54);
    return {
      eventId,
      name,
      date,
      location,
      status: 'active',
      ownerAdmin: String(ownerAdmin || 'admin').trim() || 'admin',
      eventType: 'lap_timed',
      lapDistanceKm: lapKm,
      debounceSeconds: 30,
      rankBy: 'distance',
      unifiedGunTimeMs: null,
      categories: [
        { id: '6H', name: '6小時繞圈組', durationHours: 6, distanceKm: 6, gunTimeMs: null, gunTimeIso: '', color: '#10b981' },
        { id: '12H', name: '12小時繞圈組', durationHours: 12, distanceKm: 12, gunTimeMs: null, gunTimeIso: '', color: '#3b82f6' },
        { id: '24H', name: '24小時繞圈組', durationHours: 24, distanceKm: 24, gunTimeMs: null, gunTimeIso: '', color: '#ef4444' }
      ],
      checkpoints: [
        { id: 'LAP', name: '繞圈計次感應區', order: 1, distanceKm: lapKm, type: 'lap' }
      ],
      updatedAt: Date.now()
    };
  }

  function createDefaultAdminAccounts(adminUsername = 'admin', adminPassword = 'admin1234') {
    const now = Date.now();
    return [
      {
        username: String(adminUsername || 'admin').trim() || 'admin',
        password: String(adminPassword || 'admin1234'),
        displayName: '最高系統管理員',
        isSuperAdmin: true,
        enabled: true,
        updatedAt: now
      }
    ];
  }

  function createDefaultStaffAccounts() {
    const now = Date.now();
    return [
      { username: 'writer01', password: '1111', displayName: '晶片寫入發卡組', role: 'writer', eventId: 'ALL', checkpointId: '', ownerAdmin: '', enabled: true, updatedAt: now },
      { username: 'checkin01', password: '1111', displayName: '選手報到感應組', role: 'scanner', eventId: 'ALL', checkpointId: 'CHECKIN', ownerAdmin: '', enabled: true, updatedAt: now },
      { username: 'lap01', password: '1111', displayName: '繞圈計次感應組', role: 'scanner', eventId: 'ALL', checkpointId: 'LAP', ownerAdmin: '', enabled: true, updatedAt: now },
      { username: 'cp1', password: '1111', displayName: '補給站 1 感應組', role: 'scanner', eventId: 'ALL', checkpointId: 'CP1', ownerAdmin: '', enabled: true, updatedAt: now },
      { username: 'cp2', password: '1111', displayName: '補給站 2 / 折返點組', role: 'scanner', eventId: 'ALL', checkpointId: 'CP2', ownerAdmin: '', enabled: true, updatedAt: now },
      { username: 'cp3', password: '1111', displayName: '補給站 3 感應組', role: 'scanner', eventId: 'ALL', checkpointId: 'CP3', ownerAdmin: '', enabled: true, updatedAt: now },
      { username: 'finish01', password: '1111', displayName: '終點衝線感應組', role: 'scanner', eventId: 'ALL', checkpointId: 'FINISH', ownerAdmin: '', enabled: true, updatedAt: now }
    ];
  }

  function normalizePlatform(raw) {
    const base = raw || createDefaultPlatform();
    const adminUsername = String(base.adminUsername || 'admin').trim() || 'admin';
    const adminPassword = String(base.adminPassword || 'admin1234');

    let adminAccounts = Array.isArray(base.adminAccounts)
      ? base.adminAccounts
          .filter((a) => a && String(a.username || '').trim())
          .map((a) => ({
            username: String(a.username).trim(),
            password: String(a.password ?? ''),
            displayName: String(a.displayName || a.username).trim(),
            isSuperAdmin: Boolean(a.isSuperAdmin),
            enabled: a.enabled !== false,
            updatedAt: Number(a.updatedAt || Date.now())
          }))
      : [];

    if (adminAccounts.length === 0) {
      adminAccounts = createDefaultAdminAccounts(adminUsername, adminPassword);
    } else if (!adminAccounts.some((a) => a.isSuperAdmin && a.enabled !== false)) {
      const idx = adminAccounts.findIndex((a) => a.username === adminUsername);
      if (idx >= 0) {
        adminAccounts[idx] = { ...adminAccounts[idx], isSuperAdmin: true, enabled: true };
      } else {
        adminAccounts.unshift(createDefaultAdminAccounts(adminUsername, adminPassword)[0]);
      }
    }

    const primarySuper = adminAccounts.find((a) => a.isSuperAdmin && a.enabled !== false) || adminAccounts[0];
    const defaultOwner = primarySuper.username;

    const events = (Array.isArray(base.events) ? base.events : []).map((ev) => {
      const fallbackGun = (ev.categories || []).find((c) => c && c.gunTimeMs)?.gunTimeMs ?? null;
      return {
        ...ev,
        ownerAdmin: String(ev.ownerAdmin || '').trim() || defaultOwner,
        unifiedGunTimeMs: ev.unifiedGunTimeMs !== undefined && ev.unifiedGunTimeMs !== null
          ? Number(ev.unifiedGunTimeMs)
          : (fallbackGun ? Number(fallbackGun) : null)
      };
    });

    const staffAccounts = (Array.isArray(base.staffAccounts) ? base.staffAccounts : createDefaultStaffAccounts()).map((s) => ({
      ...s,
      ownerAdmin: String(s.ownerAdmin || '').trim()
    }));

    // 升級相容：若既有資料庫尚無任何報到感應帳號 (CHECKIN)，自動補建預設報到感應組 (checkin01 / 1111)
    if (staffAccounts.length > 0 && !staffAccounts.some((s) => String(s.checkpointId || '').toUpperCase() === 'CHECKIN' || s.username === 'checkin01')) {
      staffAccounts.splice(1, 0, {
        username: 'checkin01',
        password: '1111',
        displayName: '選手報到感應組',
        role: 'scanner',
        eventId: 'ALL',
        checkpointId: 'CHECKIN',
        ownerAdmin: '',
        enabled: true,
        updatedAt: Date.now()
      });
    }

    return {
      ...base,
      id: 'main',
      adminUsername: primarySuper.username,
      adminPassword: primarySuper.password,
      adminAccounts,
      events,
      staffAccounts
    };
  }

  function createDefaultPlatform() {
    const defaultEvt = createDefaultEvent('EVT-2026-01', '2026 智慧城市馬拉松', getTodayStr(), '台北市政府廣場', 'admin');
    return {
      id: 'main',
      adminUsername: 'admin',
      adminPassword: 'admin1234',
      adminAccounts: createDefaultAdminAccounts('admin', 'admin1234'),
      activeEventId: defaultEvt.eventId,
      events: [defaultEvt],
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

  function makeRegId(eventId, idNumber, bib) {
    const cleanId = String(idNumber || '').trim().toUpperCase();
    const cleanBib = String(bib || '').trim();
    return `${String(eventId).trim()}__${cleanId || cleanBib}`;
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
          if (!db.objectStoreNames.contains('platform')) {
            db.createObjectStore('platform', { keyPath: 'id' });
          }
          if (!db.objectStoreNames.contains('chips')) {
            const store = db.createObjectStore('chips', { keyPath: 'idNumber' });
            store.createIndex('nfcUid', 'nfcUid', { unique: false });
          }
          if (!db.objectStoreNames.contains('registrations')) {
            const store = db.createObjectStore('registrations', { keyPath: 'regId' });
            store.createIndex('eventId', 'eventId', { unique: false });
            store.createIndex('idNumber', 'idNumber', { unique: false });
          }
          if (!db.objectStoreNames.contains('scans')) {
            const store = db.createObjectStore('scans', { keyPath: 'id' });
            store.createIndex('eventId', 'eventId', { unique: false });
            store.createIndex('bib', 'bib', { unique: false });
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });

      const existing = await this.getPlatform();
      if (!existing) {
        await this.savePlatform(createDefaultPlatform(), true);
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

    /* ---------------- 裝置設定與登入階段 ---------------- */

    getDeviceSettings() {
      const cfgGasUrl = (global.APP_CONFIG && global.APP_CONFIG.DEFAULT_GAS_URL ? String(global.APP_CONFIG.DEFAULT_GAS_URL).trim() : '');
      const defaults = {
        mode: cfgGasUrl ? 'online' : 'offline',
        gasUrl: cfgGasUrl,
        configAppliedUrl: cfgGasUrl,
        deviceId: generateDeviceId(),
        highContrast: false,
        soundEnabled: true,
        selectedEventId: '',
        lastSyncAt: null
      };
      try {
        const raw =
          localStorage.getItem(SETTINGS_KEY) ||
          localStorage.getItem('smart_race_device_settings_v2') ||
          localStorage.getItem('smart_race_device_settings_v1');
        if (!raw) {
          localStorage.setItem(SETTINGS_KEY, JSON.stringify(defaults));
          return defaults;
        }
        const parsed = JSON.parse(raw);
        const merged = { ...defaults, ...parsed };
        // 若 js/config.js 設定了共用雲端網址，且與上次套用過的不同（或本機尚未設定），自動更新並開啟連線
        if (cfgGasUrl && (!merged.gasUrl || merged.configAppliedUrl !== cfgGasUrl)) {
          merged.gasUrl = cfgGasUrl;
          merged.configAppliedUrl = cfgGasUrl;
          merged.mode = 'online';
          try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(merged)); } catch {}
        }
        return merged;
      } catch {
        return defaults;
      }
    }

    saveDeviceSettings(patch) {
      const merged = { ...this.getDeviceSettings(), ...patch };
      try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(merged)); } catch {}
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
        if (!sessionObj) sessionStorage.removeItem(SESSION_KEY);
        else sessionStorage.setItem(SESSION_KEY, JSON.stringify(sessionObj));
      } catch {}
    }

    /* ---------------- 平台設定與多場次活動 (Platform & Events) ---------------- */

    async getPlatform() {
      await this.init();
      const raw = await this._req(this._tx('platform').get('main'));
      return raw ? normalizePlatform(raw) : null;
    }

    async savePlatform(platform, markSynced = false) {
      await this.init();
      const normalized = normalizePlatform(platform);
      const record = { ...normalized, id: 'main', synced: markSynced ? 1 : 0, updatedAt: Date.now() };
      await this._req(this._tx('platform', 'readwrite').put(record));
      return record;
    }

    createEventTemplate(name, date, location, eventType = 'distance', lapDistanceKm = 2.54, ownerAdmin = 'admin') {
      const id = 'EVT-' + Date.now().toString(36).toUpperCase().slice(-6);
      if (eventType === 'lap_timed') {
        return createLapTimedEvent(id, name || '新繞圈計時賽', date || getTodayStr(), location || '', lapDistanceKm, ownerAdmin);
      }
      return createDefaultEvent(id, name || '新路跑活動', date || getTodayStr(), location || '', ownerAdmin);
    }

    async verifyLocalLogin(mode, username, password) {
      const p = await this.getPlatform();
      const u = String(username || '').trim();
      const pw = String(password || '');

      if (mode === 'admin') {
        const matchedAdmin = (p.adminAccounts || []).find((a) => a.username === u && a.password === pw);
        if (matchedAdmin) {
          if (matchedAdmin.enabled === false) {
            return { ok: false, error: '此管理者帳號已停用' };
          }
          return {
            ok: true,
            account: {
              username: matchedAdmin.username,
              password: matchedAdmin.password,
              displayName: matchedAdmin.displayName || matchedAdmin.username,
              isSuperAdmin: Boolean(matchedAdmin.isSuperAdmin),
              role: 'admin'
            }
          };
        }
        return { ok: false, error: '管理員帳號或密碼錯誤' };
      }

      const matched = (p.staffAccounts || []).find((s) => s.enabled !== false && s.username === u && s.password === pw);
      if (!matched) {
        return { ok: false, error: '作業員帳號或密碼錯誤（若為新帳號請先在有網路狀態登入同步一次）' };
      }
      return { ok: true, account: matched };
    }

    async cacheAdminAccount(account) {
      if (!account || !account.username) return;
      const p = await this.getPlatform();
      const list = [...(p.adminAccounts || [])];
      const idx = list.findIndex((a) => a.username === account.username);
      const entry = {
        username: String(account.username).trim(),
        password: String(account.password ?? (idx >= 0 ? list[idx].password : '')),
        displayName: String(account.displayName || account.username).trim(),
        isSuperAdmin: Boolean(account.isSuperAdmin),
        enabled: account.enabled !== false,
        updatedAt: Date.now()
      };
      if (idx >= 0) list[idx] = { ...list[idx], ...entry };
      else list.push(entry);
      await this.savePlatform({ ...p, adminAccounts: list }, true);
    }

    async cacheStaffAccount(account) {
      if (!account || !account.username) return;
      const p = await this.getPlatform();
      const list = [...(p.staffAccounts || [])];
      const idx = list.findIndex((s) => s.username === account.username);
      if (idx >= 0) list[idx] = { ...list[idx], ...account };
      else list.push(account);
      await this.savePlatform({ ...p, staffAccounts: list });
    }

    /* ---------------- 共用晶片主檔 (Chips — 姓名 + 身份證字號 + UID，跨活動通用) ---------------- */

    async getAllChips() {
      await this.init();
      return this._req(this._tx('chips').getAll());
    }

    async getChipById(idNumber) {
      await this.init();
      const key = String(idNumber || '').trim().toUpperCase();
      if (!key) return null;
      return this._req(this._tx('chips').get(key));
    }

    async findChipByUid(nfcUid) {
      const target = String(nfcUid || '').trim().toUpperCase();
      if (!target) return null;
      const all = await this.getAllChips();
      return all.find((c) => String(c.nfcUid || '').trim().toUpperCase() === target) || null;
    }

    async upsertChip({ idNumber, name, nfcUid = '', writtenAt = Date.now() }, markSynced = false) {
      await this.init();
      const cleanId = String(idNumber || '').trim().toUpperCase();
      if (!cleanId) throw new Error('身份證字號不可為空');
      const existing = await this.getChipById(cleanId);
      const record = {
        idNumber: cleanId,
        name: String(name ?? existing?.name ?? '').trim(),
        nfcUid: String(nfcUid || existing?.nfcUid || '').trim().toUpperCase(),
        writtenAt: writtenAt || existing?.writtenAt || Date.now(),
        updatedAt: Date.now(),
        synced: markSynced ? 1 : 0
      };
      await this._req(this._tx('chips', 'readwrite').put(record));
      return record;
    }

    async deleteChip(idNumber) {
      await this.init();
      await this._req(this._tx('chips', 'readwrite').delete(String(idNumber).trim().toUpperCase()));
    }

    /* ---------------- 各活動選手明細 / 報名名單 (Registrations) ---------------- */

    async getAllRegistrations() {
      await this.init();
      return this._req(this._tx('registrations').getAll());
    }

    async getEventRegistrations(eventId) {
      const all = await this.getAllRegistrations();
      return all.filter((r) => r.eventId === eventId);
    }

    async upsertRegistration(reg, markSynced = false) {
      await this.init();
      const eventId = String(reg.eventId || '').trim();
      const bib = String(reg.bib || '').trim();
      const idNumber = String(reg.idNumber || '').trim().toUpperCase();
      if (!eventId || !bib) throw new Error('活動代碼與參賽編號不可為空');

      const regId = String(reg.regId || makeRegId(eventId, idNumber, bib)).trim();
      const existing = await this._req(this._tx('registrations').get(regId));
      const finalId = idNumber || existing?.idNumber || '';
      const idHash = finalId ? await global.TimingEngine.sha256Hex(finalId) : (reg.idHash || existing?.idHash || '');

      const record = {
        regId,
        eventId,
        bib,
        name: String(reg.name ?? existing?.name ?? '').trim(),
        idNumber: finalId,
        idMasked: reg.idMasked || global.TimingEngine.maskIdNumber(finalId),
        idHash,
        categoryId: String(reg.categoryId ?? existing?.categoryId ?? '').trim(),
        penaltySec: Number(reg.penaltySec ?? existing?.penaltySec ?? 0),
        updatedAt: Date.now(),
        synced: markSynced ? 1 : 0
      };
      await this._req(this._tx('registrations', 'readwrite').put(record));
      return record;
    }

    async upsertRegistrationsBulk(list, markSynced = false) {
      const results = [];
      for (const item of list) {
        if (!item || !item.eventId || !item.bib) continue;
        results.push(await this.upsertRegistration(item, markSynced));
      }
      return results;
    }

    async deleteRegistration(regId) {
      await this.init();
      await this._req(this._tx('registrations', 'readwrite').delete(String(regId)));
    }

    /**
     * 核心：依晶片資料（身份證字號 / 硬體 UID / 姓名）比對「指定活動」的選手明細！
     */
    async matchRegistrationForEvent(eventId, { idNumber = '', name = '', nfcUid = '', bib = '' }) {
      const regs = await this.getEventRegistrations(eventId);
      let resolvedId = String(idNumber || '').trim().toUpperCase();
      let resolvedName = String(name || '').trim();

      // 若為純 UID 感應，先查跨活動共用晶片主檔 Chips 取得身份證與姓名
      if (!resolvedId && nfcUid) {
        const chip = await this.findChipByUid(nfcUid);
        if (chip) {
          resolvedId = chip.idNumber;
          resolvedName = resolvedName || chip.name;
        }
      }

      // 1. 最精準比對：該活動內比對「身份證字號」
      if (resolvedId) {
        const byId = regs.find((r) => String(r.idNumber || '').trim().toUpperCase() === resolvedId);
        if (byId) return { matched: true, registration: byId, resolvedId, resolvedName: resolvedName || byId.name };
      }

      // 2. 次要比對：相容手動補登或舊晶片帶入 bib
      const cleanBib = String(bib || '').trim().toUpperCase();
      if (cleanBib) {
        const byBib = regs.find((r) => String(r.bib).trim().toUpperCase() === cleanBib);
        if (byBib) return { matched: true, registration: byBib, resolvedId: resolvedId || byBib.idNumber, resolvedName: resolvedName || byBib.name };
      }

      // 3. 若晶片僅有姓名且該活動內僅一人同名，做輔助比對
      if (!resolvedId && resolvedName) {
        const byNameList = regs.filter((r) => r.name === resolvedName);
        if (byNameList.length === 1) {
          return { matched: true, registration: byNameList[0], resolvedId: byNameList[0].idNumber, resolvedName };
        }
      }

      return { matched: false, registration: null, resolvedId, resolvedName };
    }

    async findRegistrationByQuery(eventId, query) {
      const q = String(query || '').trim().toUpperCase();
      if (!q) return null;
      const regs = await this.getEventRegistrations(eventId);
      return (
        regs.find((r) => String(r.bib).trim().toUpperCase() === q) ||
        regs.find((r) => String(r.idNumber || '').trim().toUpperCase() === q) ||
        regs.find((r) => String(r.name || '').trim().toUpperCase() === q) ||
        null
      );
    }

    /* ---------------- 各活動感應紀錄 (Scans) ---------------- */

    async getAllScans() {
      await this.init();
      return this._req(this._tx('scans').getAll());
    }

    async getEventScans(eventId) {
      const all = await this.getAllScans();
      return all.filter((s) => s.eventId === eventId);
    }

    async addScan({ eventId, bib, idNumber = '', checkpointId, timestampMs = Date.now(), source = 'nfc', operator = '', note = '', bypassDebounce = false }) {
      await this.init();
      const cleanEvt = String(eventId || '').trim();
      const cleanBib = String(bib || '').trim();
      const cleanCp = String(checkpointId || '').trim();
      if (!cleanEvt || !cleanBib || !cleanCp) throw new Error('缺少活動、編號或感應站點');

      const platform = await this.getPlatform();
      const evt = (platform.events || []).find((e) => e.eventId === cleanEvt);
      const debounceSec = Number(evt?.debounceSeconds ?? 30);

      if (!bypassDebounce && debounceSec > 0) {
        const evtScans = await this.getEventScans(cleanEvt);
        const recent = evtScans
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
        eventId: cleanEvt,
        bib: cleanBib,
        idNumber: String(idNumber || '').trim().toUpperCase(),
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

    /* ---------------- 待同步狀態與批次標記 ---------------- */

    async getPendingSyncCounts() {
      const [platform, chips, regs, scans] = await Promise.all([
        this.getPlatform(),
        this.getAllChips(),
        this.getAllRegistrations(),
        this.getAllScans()
      ]);
      const platformDirty = platform && platform.synced === 0 ? 1 : 0;
      const unsyncedChips = chips.filter((c) => !c.synced);
      const unsyncedRegs = regs.filter((r) => !r.synced);
      const unsyncedScans = scans.filter((s) => !s.synced);
      return {
        platformDirty,
        chips: unsyncedChips.length,
        registrations: unsyncedRegs.length,
        scans: unsyncedScans.length,
        total: platformDirty + unsyncedChips.length + unsyncedRegs.length + unsyncedScans.length,
        unsyncedChips,
        unsyncedRegs,
        unsyncedScans
      };
    }

    async markSynced({ platform = false, chipIds = [], regIds = [], scanIds = [] }) {
      await this.init();
      if (platform) {
        const cur = await this.getPlatform();
        if (cur && !cur.synced) {
          cur.synced = 1;
          await this._req(this._tx('platform', 'readwrite').put(cur));
        }
      }
      if (chipIds.length) {
        const set = new Set(chipIds);
        const all = await this.getAllChips();
        const tx = this.db.transaction('chips', 'readwrite');
        for (const c of all) {
          if (set.has(c.idNumber) && !c.synced) { c.synced = 1; tx.objectStore('chips').put(c); }
        }
        await new Promise((r, j) => { tx.oncomplete = r; tx.onerror = () => j(tx.error); });
      }
      if (regIds.length) {
        const set = new Set(regIds);
        const all = await this.getAllRegistrations();
        const tx = this.db.transaction('registrations', 'readwrite');
        for (const reg of all) {
          if (set.has(reg.regId) && !reg.synced) { reg.synced = 1; tx.objectStore('registrations').put(reg); }
        }
        await new Promise((r, j) => { tx.oncomplete = r; tx.onerror = () => j(tx.error); });
      }
      if (scanIds.length) {
        const set = new Set(scanIds);
        const all = await this.getAllScans();
        const tx = this.db.transaction('scans', 'readwrite');
        for (const s of all) {
          if (set.has(s.id) && !s.synced) { s.synced = 1; tx.objectStore('scans').put(s); }
        }
        await new Promise((r, j) => { tx.oncomplete = r; tx.onerror = () => j(tx.error); });
      }
    }

    /**
     * 將雲端快照合併至本機 IndexedDB（支援多管理員活動隔離，不覆蓋其他管理員的本機活動與未上傳紀錄）
     */
    async mergeCloudData(cloudData, { scopedOwnerAdmin = '', scopedAdminOnly = false } = {}) {
      await this.init();
      if (!cloudData) return;

      const cur = await this.getPlatform();
      const defaultOwner = cloudData.globalConfig?.adminUsername || cur.adminUsername || 'admin';
      const localEventMap = new Map((cur.events || []).map((e) => [e.eventId, e]));

      // 1. 合併活動清單
      let mergedEvents = cur.events || [];
      if (Array.isArray(cloudData.events)) {
        const normalizedCloudEvents = cloudData.events.map((ev) => ({
          ...ev,
          ownerAdmin: String(ev.ownerAdmin || '').trim() || localEventMap.get(ev.eventId)?.ownerAdmin || scopedOwnerAdmin || defaultOwner
        }));
        if (scopedOwnerAdmin) {
          // 僅更新此管理員名下的活動，保留其他管理員的本機活動
          const otherEvents = (cur.events || []).filter((e) => (e.ownerAdmin || defaultOwner) !== scopedOwnerAdmin);
          mergedEvents = [...otherEvents, ...normalizedCloudEvents];
        } else if (normalizedCloudEvents.length > 0) {
          mergedEvents = normalizedCloudEvents;
        }
      }

      // 2. 合併管理者帳號清單
      let mergedAdmins = cur.adminAccounts || [];
      if (Array.isArray(cloudData.adminAccounts) && cloudData.adminAccounts.length > 0) {
        if (scopedAdminOnly) {
          const map = new Map(mergedAdmins.map((a) => [a.username, a]));
          cloudData.adminAccounts.forEach((a) => {
            if (a && a.username) map.set(a.username, { ...(map.get(a.username) || {}), ...a });
          });
          mergedAdmins = Array.from(map.values());
        } else {
          mergedAdmins = cloudData.adminAccounts;
        }
      } else if (cloudData.globalConfig?.adminUsername && cloudData.globalConfig?.adminPassword) {
        const map = new Map(mergedAdmins.map((a) => [a.username, a]));
        const u = String(cloudData.globalConfig.adminUsername).trim();
        map.set(u, {
          ...(map.get(u) || {}),
          username: u,
          password: String(cloudData.globalConfig.adminPassword),
          displayName: map.get(u)?.displayName || '最高系統管理員',
          isSuperAdmin: true,
          enabled: true,
          updatedAt: Date.now()
        });
        mergedAdmins = Array.from(map.values());
      }

      // 3. 合併作業員帳號清單
      let mergedStaff = cur.staffAccounts || [];
      if (Array.isArray(cloudData.staffAccounts)) {
        if (scopedOwnerAdmin) {
          const keepOtherStaff = (cur.staffAccounts || []).filter((s) => {
            const o = String(s.ownerAdmin || '').trim();
            if (scopedAdminOnly) return o !== scopedOwnerAdmin;
            return o && o !== scopedOwnerAdmin;
          });
          mergedStaff = [...keepOtherStaff, ...cloudData.staffAccounts];
        } else if (cloudData.staffAccounts.length > 0) {
          mergedStaff = cloudData.staffAccounts;
        }
      }

      const mergedPlatform = {
        ...cur,
        adminUsername: cloudData.globalConfig?.adminUsername || cur.adminUsername,
        adminPassword: cloudData.globalConfig?.adminPassword || cur.adminPassword,
        adminAccounts: mergedAdmins,
        activeEventId: cloudData.globalConfig?.activeEventId || (mergedEvents[0]?.eventId || cur.activeEventId),
        events: mergedEvents,
        staffAccounts: mergedStaff
      };
      await this.savePlatform(mergedPlatform, true);

      // 計算目前受此雲端同步管轄的活動集合（避免刪除其他管理員的報名與感應快取）
      const scopedEventIds = scopedOwnerAdmin
        ? new Set(mergedEvents.filter((e) => (e.ownerAdmin || defaultOwner) === scopedOwnerAdmin).map((e) => e.eventId))
        : null;

      if (Array.isArray(cloudData.chips)) {
        const cloudChipSet = new Set(cloudData.chips.map((c) => String(c.idNumber || '').trim().toUpperCase()));
        const localChips = await this.getAllChips();
        for (const lc of localChips) {
          if (lc.synced && !cloudChipSet.has(lc.idNumber)) {
            await this.deleteChip(lc.idNumber);
          }
        }
        for (const c of cloudData.chips) {
          const local = await this.getChipById(c.idNumber);
          if (local && !local.synced && local.updatedAt > (c.updatedAt || 0)) continue;
          await this.upsertChip(c, true);
        }
      }

      if (Array.isArray(cloudData.registrations)) {
        const localAll = await this.getAllRegistrations();
        const localMap = new Map(localAll.map((r) => [r.regId, r]));
        const cloudRegSet = new Set(cloudData.registrations.map((r) => r.regId));
        for (const lr of localAll) {
          if (scopedEventIds && !scopedEventIds.has(lr.eventId)) continue;
          if (lr.synced && !cloudRegSet.has(lr.regId)) {
            await this.deleteRegistration(lr.regId);
          }
        }
        for (const r of cloudData.registrations) {
          const local = localMap.get(r.regId);
          if (local && !local.synced && local.updatedAt > (r.updatedAt || 0)) continue;
          await this.upsertRegistration({ ...r, idNumber: r.idNumber || local?.idNumber || '' }, true);
        }
      }

      if (Array.isArray(cloudData.scans)) {
        const localAll = await this.getAllScans();
        const localMap = new Map(localAll.map((s) => [s.id, s]));
        const cloudScanSet = new Set(cloudData.scans.map((s) => s.id));
        for (const ls of localAll) {
          if (scopedEventIds && !scopedEventIds.has(ls.eventId)) continue;
          if (ls.synced && !cloudScanSet.has(ls.id)) {
            await this._req(this._tx('scans', 'readwrite').delete(ls.id));
          }
        }
        for (const s of cloudData.scans) {
          const local = localMap.get(s.id);
          if (local && !local.synced && local.updatedAt > (s.updatedAt || 0)) continue;
          await this.putScan({ ...s, synced: 1 });
        }
      }
    }

    /* ---------------- 完整系統備份 / 還原合併 (JSON) ---------------- */

    async exportFullBackup({ ownerAdmin = '', isSuperAdmin = true } = {}) {
      const [platform, chips, registrations, scans] = await Promise.all([
        this.getPlatform(),
        this.getAllChips(),
        this.getAllRegistrations(),
        this.getAllScans()
      ]);

      if (ownerAdmin) {
        const ownEvents = (platform.events || []).filter((e) => (e.ownerAdmin || platform.adminUsername) === ownerAdmin);
        const ownEventIds = new Set(ownEvents.map((e) => e.eventId));
        const ownStaff = (platform.staffAccounts || []).filter((s) => {
          const o = String(s.ownerAdmin || '').trim();
          return isSuperAdmin ? (!o || o === ownerAdmin) : o === ownerAdmin;
        });
        return {
          schemaVersion: '3.4',
          exportedAt: new Date().toISOString(),
          scopedOwnerAdmin: ownerAdmin,
          platform: {
            ...platform,
            adminAccounts: isSuperAdmin ? platform.adminAccounts : (platform.adminAccounts || []).filter((a) => a.username === ownerAdmin),
            events: ownEvents,
            staffAccounts: ownStaff
          },
          chips,
          registrations: registrations.filter((r) => ownEventIds.has(r.eventId)),
          scans: scans.filter((s) => ownEventIds.has(s.eventId))
        };
      }

      return {
        schemaVersion: '3.4',
        exportedAt: new Date().toISOString(),
        platform,
        chips,
        registrations,
        scans
      };
    }

    async importFullBackup(backupObj, strategy = 'replace', { ownerAdmin = '', isSuperAdmin = true } = {}) {
      if (!backupObj || (!backupObj.platform && !Array.isArray(backupObj.scans))) {
        throw new Error('無效的系統備份檔格式');
      }
      if (strategy === 'replace' && isSuperAdmin) {
        await this.clearAllData(false);
        if (backupObj.platform) {
          await this.savePlatform({ ...createDefaultPlatform(), ...backupObj.platform });
        }
      } else if (backupObj.platform) {
        const cur = await this.getPlatform();
        const eventMap = new Map((cur.events || []).map((e) => [e.eventId, e]));
        (backupObj.platform.events || []).forEach((e) => {
          const existing = eventMap.get(e.eventId);
          // 非最高管理員合併備份時，絕不覆蓋其他管理員的活動
          if (ownerAdmin && existing && (existing.ownerAdmin || cur.adminUsername) !== ownerAdmin) return;
          eventMap.set(e.eventId, ownerAdmin ? { ...e, ownerAdmin } : e);
        });
        const staffMap = new Map((cur.staffAccounts || []).map((s) => [s.username, s]));
        (backupObj.platform.staffAccounts || []).forEach((s) => {
          if (!s?.username) return;
          const existing = staffMap.get(s.username);
          if (ownerAdmin && existing && existing.ownerAdmin && existing.ownerAdmin !== ownerAdmin) return;
          staffMap.set(s.username, ownerAdmin && !isSuperAdmin ? { ...s, ownerAdmin } : s);
        });
        await this.savePlatform({
          ...cur,
          events: Array.from(eventMap.values()),
          staffAccounts: Array.from(staffMap.values())
        });
      }

      const latestPlatform = await this.getPlatform();
      const allowedEventIds = ownerAdmin
        ? new Set((latestPlatform.events || []).filter((e) => (e.ownerAdmin || latestPlatform.adminUsername) === ownerAdmin).map((e) => e.eventId))
        : null;

      let chipCount = 0, regCount = 0, scanCount = 0;
      for (const c of backupObj.chips || []) {
        if (!c?.idNumber) continue;
        await this.upsertChip(c, false);
        chipCount++;
      }
      for (const r of backupObj.registrations || []) {
        if (!r?.eventId || !r?.bib) continue;
        if (allowedEventIds && !allowedEventIds.has(r.eventId)) continue;
        await this.upsertRegistration(r, false);
        regCount++;
      }
      for (const s of backupObj.scans || []) {
        if (!s?.id || !s?.eventId || !s?.bib) continue;
        if (allowedEventIds && !allowedEventIds.has(s.eventId)) continue;
        await this.putScan({ ...s, synced: 0 });
        scanCount++;
      }
      return { chipCount, regCount, scanCount };
    }

    /**
     * 載入「雙場次共用晶片」完整模擬資料（展示同一選手持同一晶片參加不同活動與組別）
     */
    async loadDemoData(ownerAdmin = '') {
      await this.init();
      const prev = await this.getPlatform();
      const targetOwner = String(ownerAdmin || prev?.adminUsername || 'admin').trim() || 'admin';
      await this.clearAllData(false);

      const now = Date.now();
      const min = 60 * 1000;
      const gunSpring = now - 225 * min;
      const gunCoast = now - 85 * min;
      const gunUltra = now - 250 * min;

      const evt1 = createDefaultEvent('EVT-2026-SPRING', '2026 春季城市馬拉松', getTodayStr(), '台北市政府廣場', targetOwner);
      evt1.unifiedGunTimeMs = gunSpring;
      evt1.categories.forEach((c) => {
        c.gunTimeMs = gunSpring;
        c.gunTimeIso = new Date(gunSpring).toISOString();
      });

      const evt2 = {
        eventId: 'EVT-2026-COAST',
        name: '2026 秋季海岸公益路跑',
        date: getTodayStr(),
        location: '淡水漁人碼頭',
        status: 'active',
        ownerAdmin: targetOwner,
        eventType: 'distance',
        lapDistanceKm: 0,
        debounceSeconds: 30,
        rankBy: 'gun',
        unifiedGunTimeMs: gunCoast,
        categories: [
          { id: '12K', name: '挑戰組 12K', distanceKm: 12, gunTimeMs: gunCoast, gunTimeIso: new Date(gunCoast).toISOString(), color: '#8b5cf6' },
          { id: '5K', name: '樂活組 5K', distanceKm: 5, gunTimeMs: gunCoast, gunTimeIso: new Date(gunCoast).toISOString(), color: '#f59e0b' }
        ],
        checkpoints: [
          { id: 'CP1', name: '海岸折返點 (6K)', order: 1, distanceKm: 6, type: 'checkpoint' },
          { id: 'FINISH', name: '終點衝線門', order: 2, distanceKm: 12, type: 'finish' }
        ],
        updatedAt: now
      };

      const evt3 = createLapTimedEvent('EVT-2026-ULTRA', '2026 環湖超級馬拉松繞圈賽', getTodayStr(), '碧潭環湖賽道 (單圈2.54K)', 2.54, targetOwner);
      evt3.unifiedGunTimeMs = gunUltra;
      evt3.categories.forEach((c) => {
        c.gunTimeMs = gunUltra;
        c.gunTimeIso = new Date(gunUltra).toISOString();
      });

      const platform = createDefaultPlatform();
      platform.adminUsername = prev?.adminUsername || 'admin';
      platform.adminPassword = prev?.adminPassword || 'admin1234';
      if (prev?.adminAccounts?.length) platform.adminAccounts = prev.adminAccounts;
      if (prev?.staffAccounts?.length) platform.staffAccounts = prev.staffAccounts;
      platform.activeEventId = evt3.eventId;
      platform.events = [evt1, evt2, evt3];
      await this.savePlatform(platform);

      // 共用晶片主檔（每人僅一張晶片：姓名 + 身份證字號，可跨場參加多項活動！）
      const people = [
        { idNumber: 'A123456789', name: '林志明', nfcUid: '04:A1:10:01' },
        { idNumber: 'B234567891', name: '陳怡君', nfcUid: '04:A1:10:02' },
        { idNumber: 'C145678912', name: '張家豪', nfcUid: '04:A1:10:03' },
        { idNumber: 'F256789123', name: '王雅婷', nfcUid: '04:A1:10:04' },
        { idNumber: 'H167891234', name: '李俊傑', nfcUid: '04:A1:10:05' },
        { idNumber: 'T189123456', name: '吳冠宇', nfcUid: '04:A1:10:06' },
        { idNumber: 'N291234567', name: '劉佳穎', nfcUid: '04:A1:10:07' }
      ];
      for (const p of people) {
        await this.upsertChip(p, false);
      }

      // 活動 1 名單（春季馬拉松）
      await this.upsertRegistrationsBulk([
        { eventId: evt1.eventId, bib: 'A001', name: '林志明', idNumber: 'A123456789', categoryId: '42K' },
        { eventId: evt1.eventId, bib: 'A002', name: '陳怡君', idNumber: 'B234567891', categoryId: '42K' },
        { eventId: evt1.eventId, bib: 'A003', name: '張家豪', idNumber: 'C145678912', categoryId: '42K' },
        { eventId: evt1.eventId, bib: 'B101', name: '王雅婷', idNumber: 'F256789123', categoryId: '21K' },
        { eventId: evt1.eventId, bib: 'B102', name: '李俊傑', idNumber: 'H167891234', categoryId: '21K' },
        { eventId: evt1.eventId, bib: 'C201', name: '吳冠宇', idNumber: 'T189123456', categoryId: '10K' }
      ], false);

      // 活動 2 名單（秋季海岸路跑）
      await this.upsertRegistrationsBulk([
        { eventId: evt2.eventId, bib: 'S101', name: '林志明', idNumber: 'A123456789', categoryId: '12K' },
        { eventId: evt2.eventId, bib: 'S102', name: '陳怡君', idNumber: 'B234567891', categoryId: '12K' },
        { eventId: evt2.eventId, bib: 'L201', name: '王雅婷', idNumber: 'F256789123', categoryId: '5K' },
        { eventId: evt2.eventId, bib: 'L202', name: '劉佳穎', idNumber: 'N291234567', categoryId: '5K' }
      ], false);

      // 活動 3 名單（繞圈計時賽 — 同一批選手持同一晶片挑戰 6H / 12H / 24H 繞圈！）
      await this.upsertRegistrationsBulk([
        { eventId: evt3.eventId, bib: 'U001', name: '林志明', idNumber: 'A123456789', categoryId: '12H' },
        { eventId: evt3.eventId, bib: 'U002', name: '陳怡君', idNumber: 'B234567891', categoryId: '12H' },
        { eventId: evt3.eventId, bib: 'U101', name: '張家豪', idNumber: 'C145678912', categoryId: '24H' },
        { eventId: evt3.eventId, bib: 'U201', name: '王雅婷', idNumber: 'F256789123', categoryId: '6H' },
        { eventId: evt3.eventId, bib: 'U202', name: '吳冠宇', idNumber: 'T189123456', categoryId: '6H' }
      ], false);

      const addS = (eventId, bib, idNumber, checkpointId, timestampMs) =>
        this.addScan({ eventId, bib, idNumber, checkpointId, timestampMs, source: 'mock', operator: 'demo', bypassDebounce: true });

      // 活動 1 報到與賽道感應紀錄（採統一律定起跑時間，無起點感應；B102 尚未報到以示範報到掌握）
      await addS(evt1.eventId, 'A001', 'A123456789', 'CHECKIN', gunSpring - 32 * min);
      await addS(evt1.eventId, 'A001', 'A123456789', 'CP1', gunSpring + 23 * min + 15000);
      await addS(evt1.eventId, 'A001', 'A123456789', 'CP2', gunSpring + 47 * min + 5000);
      await addS(evt1.eventId, 'A001', 'A123456789', 'CP3', gunSpring + 99 * min + 41000);
      await addS(evt1.eventId, 'A001', 'A123456789', 'FINISH', gunSpring + 202 * min + 18450);

      await addS(evt1.eventId, 'A002', 'B234567891', 'CHECKIN', gunSpring - 28 * min);
      await addS(evt1.eventId, 'A002', 'B234567891', 'CP1', gunSpring + 25 * min + 10000);
      await addS(evt1.eventId, 'A002', 'B234567891', 'CP2', gunSpring + 51 * min + 32000);
      await addS(evt1.eventId, 'A002', 'B234567891', 'CP3', gunSpring + 108 * min + 22000);
      await addS(evt1.eventId, 'A002', 'B234567891', 'FINISH', gunSpring + 218 * min + 42310);

      await addS(evt1.eventId, 'A003', 'C145678912', 'CHECKIN', gunSpring - 25 * min);
      await addS(evt1.eventId, 'A003', 'C145678912', 'CP1', gunSpring + 29 * min + 5000);
      await addS(evt1.eventId, 'A003', 'C145678912', 'CP2', gunSpring + 60 * min + 12000);

      await addS(evt1.eventId, 'B101', 'F256789123', 'CHECKIN', gunSpring - 21 * min);
      await addS(evt1.eventId, 'B101', 'F256789123', 'CP1', gunSpring + 24 * min + 1200);
      await addS(evt1.eventId, 'B101', 'F256789123', 'CP2', gunSpring + 48 * min + 21500);
      await addS(evt1.eventId, 'B101', 'F256789123', 'FINISH', gunSpring + 101 * min + 9320);

      await addS(evt1.eventId, 'C201', 'T189123456', 'CHECKIN', gunSpring - 18 * min);
      await addS(evt1.eventId, 'C201', 'T189123456', 'CP1', gunSpring + 23 * min + 11000);
      await addS(evt1.eventId, 'C201', 'T189123456', 'FINISH', gunSpring + 46 * min + 55120);

      // 活動 2 報到與賽道感應紀錄（L202 尚未報到以示範）
      await addS(evt2.eventId, 'S101', 'A123456789', 'CHECKIN', gunCoast - 30 * min);
      await addS(evt2.eventId, 'S101', 'A123456789', 'CP1', gunCoast + 28 * min + 12000);
      await addS(evt2.eventId, 'S101', 'A123456789', 'FINISH', gunCoast + 57 * min + 34120);

      await addS(evt2.eventId, 'S102', 'B234567891', 'CHECKIN', gunCoast - 22 * min);
      await addS(evt2.eventId, 'S102', 'B234567891', 'CP1', gunCoast + 32 * min + 5000);

      await addS(evt2.eventId, 'L201', 'F256789123', 'CHECKIN', gunCoast - 15 * min);
      await addS(evt2.eventId, 'L201', 'F256789123', 'FINISH', gunCoast + 26 * min + 48200);

      // 活動 3 繞圈計時賽：先報到感應 (CHECKIN，不計入圈數)，再繞圈計次 (LAP)
      await addS(evt3.eventId, 'U101', 'C145678912', 'CHECKIN', gunUltra - 35 * min);
      await addS(evt3.eventId, 'U001', 'A123456789', 'CHECKIN', gunUltra - 30 * min);
      await addS(evt3.eventId, 'U002', 'B234567891', 'CHECKIN', gunUltra - 26 * min);
      await addS(evt3.eventId, 'U201', 'F256789123', 'CHECKIN', gunUltra - 20 * min);
      await addS(evt3.eventId, 'U202', 'T189123456', 'CHECKIN', gunUltra - 18 * min);

      const addLaps = async (bib, idNum, startMs, count, avgLapMin) => {
        let cur = startMs;
        for (let i = 1; i <= count; i++) {
          cur += Math.round((avgLapMin + ((i % 3) - 1) * 0.6) * min);
          await addS(evt3.eventId, bib, idNum, 'LAP', cur);
        }
      };
      await addLaps('U101', 'C145678912', gunUltra, 16, 15.2); // 16圈 (24H組)
      await addLaps('U001', 'A123456789', gunUltra, 15, 14.5); // 15圈 (12H組)
      await addLaps('U002', 'B234567891', gunUltra, 13, 15.8); // 13圈 (12H組)
      await addLaps('U201', 'F256789123', gunUltra, 10, 14.2); // 10圈 (6H組)
      await addLaps('U202', 'T189123456', gunUltra, 10, 14.9); // 10圈 (6H組，同里程比先後！)

      this.saveDeviceSettings({ selectedEventId: evt3.eventId });
      return true;
    }

    async clearAllData(resetPlatform = true) {
      await this.init();
      await Promise.all([
        this._req(this._tx('chips', 'readwrite').clear()),
        this._req(this._tx('registrations', 'readwrite').clear()),
        this._req(this._tx('scans', 'readwrite').clear())
      ]);
      if (resetPlatform) {
        await this.savePlatform(createDefaultPlatform());
      }
    }
  }

  global.RaceDB = new RaceDB();
})(window);
