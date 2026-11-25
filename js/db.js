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
  const DB_VERSION = 2;
  const SETTINGS_KEY = 'smart_race_platform_settings_v3';
  const SESSION_KEY = 'smart_race_platform_session_v3';

  function getTodayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function createDefaultEvent(eventId = 'EVT-2026-01', name = '2026 TUR 超級馬拉松挑戰賽', date = getTodayStr(), location = '台北市政府廣場', ownerAdmin = 'admin') {
    return {
      eventId,
      name,
      englishName: '',
      organization: 'Taiwan Ultra Runners',
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
      englishName: '',
      organization: 'Taiwan Ultra Runners',
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

  function normalizeGenderValue(rawGender, idNumber = '') {
    const g = String(rawGender ?? '').trim().toUpperCase();
    if (g === 'M' || g === 'MALE' || g === '男' || g === 'BOY' || g === 'MAN') return 'M';
    if (g === 'F' || g === 'FEMALE' || g === '女' || g === 'GIRL' || g === 'WOMAN') return 'F';
    if (g === '') {
      const id = String(idNumber || '').trim().toUpperCase();
      if (/^[A-Z][1289]\d{8}$/.test(id)) {
        return (id[1] === '1' || id[1] === '8') ? 'M' : 'F';
      }
    }
    return '';
  }

  function normalizeDobValue(raw) {
    const s = String(raw ?? '').trim();
    if (!s) return '';
    const m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    const d = new Date(s);
    if (!Number.isNaN(d.getTime()) && d.getFullYear() >= 1900 && d.getFullYear() <= 2100) {
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
    return s;
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
    const defaultEvt = createDefaultEvent('EVT-2026-01', '2026 TUR 超級馬拉松挑戰賽', getTodayStr(), '台北市政府廣場', 'admin');
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

  function normalizeHistoryDate(rawDate, rawYear) {
    const d = String(rawDate || '').trim();
    if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(d)) {
      const parts = d.split(/[-/.]/);
      return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }
    const yMatch = (d || String(rawYear || '')).match(/(\d{4})/);
    if (yMatch) return `${yMatch[1]}-01-01`;
    return getTodayStr();
  }

  function makeHistoryId(eventDate, eventName, idNumber, bib, name, categoryName) {
    const cleanId = String(idNumber || '').trim().toUpperCase();
    const cleanBib = String(bib || '').trim().toUpperCase();
    const cleanName = String(name || '').trim();
    const cleanEvt = String(eventName || '').trim();
    const cleanCat = String(categoryName || '').trim();
    const personKey = cleanId || `${cleanName}_${cleanBib || cleanCat}`;
    return `HIST__${eventDate}__${cleanEvt}__${personKey}`.replace(/\s+/g, '_');
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
          if (!db.objectStoreNames.contains('history_records')) {
            const store = db.createObjectStore('history_records', { keyPath: 'id' });
            store.createIndex('year', 'year', { unique: false });
            store.createIndex('eventName', 'eventName', { unique: false });
            store.createIndex('idNumber', 'idNumber', { unique: false });
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

    _runBatchTx(storeName, { deletes = [], puts = [] } = {}) {
      if (!deletes.length && !puts.length) return Promise.resolve();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(storeName, 'readwrite');
        const store = tx.objectStore(storeName);
        for (const key of deletes) store.delete(key);
        for (const item of puts) store.put(item);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error || new Error('IndexedDB 批次寫入中止'));
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
        // 安全防護：若先前曾殘留任何管理員登入紀錄於 localStorage / sessionStorage，一律立即清除，絕不自動登入
        const lsRaw = localStorage.getItem(SESSION_KEY);
        if (lsRaw) localStorage.removeItem(SESSION_KEY);

        const raw = sessionStorage.getItem(SESSION_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (!parsed || parsed.mode === 'admin') {
          sessionStorage.removeItem(SESSION_KEY);
          return null;
        }
        return parsed;
      } catch {
        try {
          localStorage.removeItem(SESSION_KEY);
          sessionStorage.removeItem(SESSION_KEY);
        } catch {}
        return null;
      }
    }

    saveSession(sessionObj) {
      try {
        localStorage.removeItem(SESSION_KEY);
        // 管理員 (最高管理員／活動承辦人) 登入狀態僅存於當頁記憶體，絕不寫入 localStorage 或 sessionStorage，防止外人藉登入紀錄進入
        if (!sessionObj || sessionObj.mode === 'admin') {
          sessionStorage.removeItem(SESSION_KEY);
        } else {
          sessionStorage.setItem(SESSION_KEY, JSON.stringify(sessionObj));
        }
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

      const rawGender = reg.gender !== undefined ? reg.gender : (existing?.gender ?? '');
      const finalGender = normalizeGenderValue(rawGender, finalId);
      const finalDob = reg.dob !== undefined ? normalizeDobValue(reg.dob) : (existing?.dob || '');

      const record = {
        regId,
        eventId,
        bib,
        name: String(reg.name ?? existing?.name ?? '').trim(),
        lastName: String(reg.lastName ?? existing?.lastName ?? '').trim(),
        firstName: String(reg.firstName ?? existing?.firstName ?? '').trim(),
        gender: finalGender,
        dob: finalDob,
        city: String(reg.city ?? existing?.city ?? '').trim(),
        country: String(reg.country ?? existing?.country ?? '').trim(),
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
      await this.init();
      const valid = (list || []).filter((item) => item && String(item.eventId || '').trim() && String(item.bib || '').trim());
      if (!valid.length) return [];

      const existingAll = await this.getAllRegistrations();
      const existingMap = new Map(existingAll.map((r) => [r.regId, r]));
      const hashCache = new Map();
      const now = Date.now();

      const records = await Promise.all(valid.map(async (reg) => {
        const eventId = String(reg.eventId || '').trim();
        const bib = String(reg.bib || '').trim();
        const idNumber = String(reg.idNumber || '').trim().toUpperCase();
        const regId = String(reg.regId || makeRegId(eventId, idNumber, bib)).trim();
        const existing = existingMap.get(regId);
        const finalId = idNumber || existing?.idNumber || '';

        let idHash = reg.idHash || (existing && existing.idNumber === finalId ? existing.idHash : '') || '';
        if (!idHash && finalId) {
          if (!hashCache.has(finalId)) {
            hashCache.set(finalId, global.TimingEngine.sha256Hex(finalId));
          }
          idHash = await hashCache.get(finalId);
        }

        const rawGender = reg.gender !== undefined ? reg.gender : (existing?.gender ?? '');
        const finalGender = normalizeGenderValue(rawGender, finalId);
        const finalDob = reg.dob !== undefined ? normalizeDobValue(reg.dob) : (existing?.dob || '');

        return {
          regId,
          eventId,
          bib,
          name: String(reg.name ?? existing?.name ?? '').trim(),
          lastName: String(reg.lastName ?? existing?.lastName ?? '').trim(),
          firstName: String(reg.firstName ?? existing?.firstName ?? '').trim(),
          gender: finalGender,
          dob: finalDob,
          city: String(reg.city ?? existing?.city ?? '').trim(),
          country: String(reg.country ?? existing?.country ?? '').trim(),
          idNumber: finalId,
          idMasked: reg.idMasked || global.TimingEngine.maskIdNumber(finalId),
          idHash,
          categoryId: String(reg.categoryId ?? existing?.categoryId ?? '').trim(),
          penaltySec: Number(reg.penaltySec ?? existing?.penaltySec ?? 0),
          updatedAt: Number(reg.updatedAt || now),
          synced: markSynced ? 1 : 0
        };
      }));

      await this._runBatchTx('registrations', { puts: records });
      return records;
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

    /**
     * 刪除指定活動的所有本機明細資料（參賽明細 Registrations 與 感應紀錄 Scans）
     * 跨活動共用晶片主檔 (Chips) 與其他活動資料均維持不變
     */
    async deleteEventData(eventId) {
      await this.init();
      const targetId = String(eventId || '').trim();
      if (!targetId) return { deletedRegistrations: 0, deletedScans: 0 };

      const [regs, scans] = await Promise.all([
        this.getEventRegistrations(targetId),
        this.getEventScans(targetId)
      ]);

      let deletedRegistrations = 0;
      if (regs.length > 0) {
        const tx = this.db.transaction('registrations', 'readwrite');
        const store = tx.objectStore('registrations');
        for (const r of regs) {
          store.delete(r.regId);
          deletedRegistrations++;
        }
        await new Promise((resolve, reject) => {
          tx.oncomplete = resolve;
          tx.onerror = () => reject(tx.error);
        });
      }

      let deletedScans = 0;
      if (scans.length > 0) {
        const tx = this.db.transaction('scans', 'readwrite');
        const store = tx.objectStore('scans');
        for (const s of scans) {
          store.delete(s.id);
          deletedScans++;
        }
        await new Promise((resolve, reject) => {
          tx.oncomplete = resolve;
          tx.onerror = () => reject(tx.error);
        });
      }

      return { deletedRegistrations, deletedScans };
    }

    /* ---------------- 歷年活動及成績匯入 (History Records) ---------------- */

    async getAllHistoryRecords() {
      await this.init();
      return this._req(this._tx('history_records').getAll());
    }

    async normalizeHistoryItem(item, fromCloud = false) {
      const TE = global.TimingEngine;
      const eventDate = normalizeHistoryDate(item.eventDate || item.date, item.year);
      const yearMatch = eventDate.match(/^(\d{4})/);
      const year = String(item.year || (yearMatch ? yearMatch[1] : '')).trim() || eventDate.slice(0, 4);
      const eventName = String(item.eventName || '').trim();
      const categoryName = String(item.categoryName || item.category || '').trim() || '一般組';
      const bib = String(item.bib || '').trim();
      const name = String(item.name || '').trim();
      const idNumber = String(item.idNumber || '').trim().toUpperCase();
      const idMasked = item.idMasked || (idNumber ? TE.maskIdNumber(idNumber) : '');
      const idHash = item.idHash || (idNumber ? await TE.sha256Hex(idNumber) : '');
      const distanceKm = Math.max(0, Number(item.distanceKm || 0) || 0);
      const laps = Math.max(0, parseInt(item.laps || 0, 10) || 0);
      const overallRank = item.overallRank !== undefined && item.overallRank !== null && String(item.overallRank).trim() !== ''
        ? (parseInt(item.overallRank, 10) || null)
        : null;
      const categoryRank = item.categoryRank !== undefined && item.categoryRank !== null && String(item.categoryRank).trim() !== ''
        ? (parseInt(item.categoryRank, 10) || null)
        : null;
      const finishTime = global.TimingEngine?.normalizeTimeStr
        ? global.TimingEngine.normalizeTimeStr(item.finishTime)
        : String(item.finishTime || '').trim().replace(/^(-?\d{1,3}:\d{2}:\d{2})\.\d+$/, '$1');
      const rawStatus = String(item.status || '').trim();
      const status = rawStatus || ((finishTime || distanceKm > 0 || laps > 0) ? '完賽' : '未完賽');
      const id = String(item.id || makeHistoryId(eventDate, eventName, idNumber, bib, name, categoryName));

      return {
        id,
        year,
        eventDate,
        eventName,
        categoryName,
        bib,
        name,
        idNumber,
        idMasked,
        idHash,
        finishTime,
        distanceKm,
        laps,
        overallRank,
        categoryRank,
        status,
        note: String(item.note || '').trim(),
        ownerAdmin: String(item.ownerAdmin || 'admin').trim() || 'admin',
        synced: fromCloud ? 1 : 0,
        updatedAt: Number(item.updatedAt || Date.now())
      };
    }

    async upsertHistoryRecordsBulk(list, fromCloud = false) {
      await this.init();
      const normalized = [];
      for (const item of list || []) {
        if (!item || !String(item.eventName || '').trim() || !String(item.name || '').trim()) continue;
        normalized.push(await this.normalizeHistoryItem(item, fromCloud));
      }
      if (!normalized.length) return [];
      const tx = this.db.transaction('history_records', 'readwrite');
      const store = tx.objectStore('history_records');
      for (const rec of normalized) {
        store.put(rec);
      }
      await new Promise((resolve, reject) => {
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
      return normalized;
    }

    async deleteHistoryRecord(id) {
      await this.init();
      await this._req(this._tx('history_records', 'readwrite').delete(id));
    }

    async deleteHistoryEvent(eventDate, eventName) {
      await this.init();
      const all = await this.getAllHistoryRecords();
      const targets = all.filter((h) => h.eventDate === eventDate && h.eventName === eventName);
      if (!targets.length) return [];
      const tx = this.db.transaction('history_records', 'readwrite');
      const store = tx.objectStore('history_records');
      for (const t of targets) store.delete(t.id);
      await new Promise((resolve, reject) => {
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
      return targets.map((t) => t.id);
    }

    /* ---------------- 待同步狀態與批次標記 ---------------- */

    async getPendingSyncCounts() {
      const [platform, chips, regs, scans, history] = await Promise.all([
        this.getPlatform(),
        this.getAllChips(),
        this.getAllRegistrations(),
        this.getAllScans(),
        this.getAllHistoryRecords()
      ]);
      const platformDirty = platform && platform.synced === 0 ? 1 : 0;
      const unsyncedChips = chips.filter((c) => !c.synced);
      const unsyncedRegs = regs.filter((r) => !r.synced);
      const unsyncedScans = scans.filter((s) => !s.synced);
      const unsyncedHistory = history.filter((h) => !h.synced);
      return {
        platformDirty,
        chips: unsyncedChips.length,
        registrations: unsyncedRegs.length,
        scans: unsyncedScans.length,
        history: unsyncedHistory.length,
        total: platformDirty + unsyncedChips.length + unsyncedRegs.length + unsyncedScans.length + unsyncedHistory.length,
        unsyncedChips,
        unsyncedRegs,
        unsyncedScans,
        unsyncedHistory
      };
    }

    async markSynced({ platform = false, chipIds = [], regIds = [], scanIds = [], historyIds = [] }) {
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
      if (historyIds.length) {
        const set = new Set(historyIds);
        const all = await this.getAllHistoryRecords();
        const tx = this.db.transaction('history_records', 'readwrite');
        for (const h of all) {
          if (set.has(h.id) && !h.synced) { h.synced = 1; tx.objectStore('history_records').put(h); }
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
      const removedScopedEventIds = new Set();
      if (Array.isArray(cloudData.events)) {
        const normalizedCloudEvents = cloudData.events.map((ev) => ({
          ...ev,
          ownerAdmin: String(ev.ownerAdmin || '').trim() || localEventMap.get(ev.eventId)?.ownerAdmin || scopedOwnerAdmin || defaultOwner
        }));
        if (scopedOwnerAdmin) {
          // 僅更新此管理員名下的活動，保留其他管理員的本機活動，並記錄已被刪除的活動以便清除本機明細快取
          const nextOwnSet = new Set(normalizedCloudEvents.map((e) => e.eventId));
          (cur.events || []).forEach((e) => {
            if ((e.ownerAdmin || defaultOwner) === scopedOwnerAdmin && !nextOwnSet.has(e.eventId)) {
              removedScopedEventIds.add(e.eventId);
            }
          });
          const otherEvents = (cur.events || []).filter((e) => (e.ownerAdmin || defaultOwner) !== scopedOwnerAdmin);
          mergedEvents = [...otherEvents, ...normalizedCloudEvents];
        } else if (
          normalizedCloudEvents.length > 0 ||
          (Array.isArray(cloudData.registrations) && cloudData.registrations.length === 0 && Array.isArray(cloudData.scans) && cloudData.scans.length === 0)
        ) {
          const nextAllSet = new Set(normalizedCloudEvents.map((e) => e.eventId));
          (cur.events || []).forEach((e) => {
            if (!nextAllSet.has(e.eventId)) removedScopedEventIds.add(e.eventId);
          });
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
        if (scopedOwnerAdmin && scopedAdminOnly) {
          const keepOtherStaff = (cur.staffAccounts || []).filter((s) => {
            const o = String(s.ownerAdmin || '').trim();
            return o !== scopedOwnerAdmin;
          });
          mergedStaff = [...keepOtherStaff, ...cloudData.staffAccounts];
        } else if (scopedOwnerAdmin || cloudData.staffAccounts.length > 0) {
          // 最高管理員接收雲端完整作業員清單，以支援全平台帳號管理
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

      // 計算目前受此雲端同步管轄的活動集合（包含已被雲端刪除的本人活動，確保對應本機明細快取一併清除）
      const scopedEventIds = scopedOwnerAdmin
        ? new Set([
            ...mergedEvents.filter((e) => (e.ownerAdmin || defaultOwner) === scopedOwnerAdmin).map((e) => e.eventId),
            ...removedScopedEventIds
          ])
        : null;

      // 平行一次讀出各表本機資料，並以「單一 IndexedDB 批次交易」取代逐筆讀寫（大幅加速同步）
      const [localChips, localRegs, localScans, localHist] = await Promise.all([
        Array.isArray(cloudData.chips) ? this.getAllChips() : Promise.resolve(null),
        Array.isArray(cloudData.registrations) ? this.getAllRegistrations() : Promise.resolve(null),
        Array.isArray(cloudData.scans) ? this.getAllScans() : Promise.resolve(null),
        Array.isArray(cloudData.historyRecords) ? this.getAllHistoryRecords() : Promise.resolve(null)
      ]);

      const now = Date.now();
      const batchPromises = [];
      const sharedHashCache = new Map();
      const getHashCached = async (idStr) => {
        if (!idStr) return '';
        if (!sharedHashCache.has(idStr)) {
          sharedHashCache.set(idStr, global.TimingEngine.sha256Hex(idStr));
        }
        return sharedHashCache.get(idStr);
      };

      if (Array.isArray(cloudData.chips) && localChips) {
        const localMap = new Map(localChips.map((c) => [c.idNumber, c]));
        const cloudChipSet = new Set();
        const deletes = [];
        const puts = [];

        for (const c of cloudData.chips) {
          const cleanId = String(c.idNumber || '').trim().toUpperCase();
          if (!cleanId) continue;
          cloudChipSet.add(cleanId);
          const local = localMap.get(cleanId);
          if (local && !local.synced && local.updatedAt > (c.updatedAt || 0)) continue;
          puts.push({
            idNumber: cleanId,
            name: String(c.name ?? local?.name ?? '').trim(),
            nfcUid: String(c.nfcUid || local?.nfcUid || '').trim().toUpperCase(),
            writtenAt: c.writtenAt || local?.writtenAt || now,
            updatedAt: Number(c.updatedAt || now),
            synced: 1
          });
        }
        for (const lc of localChips) {
          if (lc.synced && !cloudChipSet.has(lc.idNumber)) deletes.push(lc.idNumber);
        }
        batchPromises.push(this._runBatchTx('chips', { deletes, puts }));
      }

      if (Array.isArray(cloudData.registrations) && localRegs) {
        const localMap = new Map(localRegs.map((r) => [r.regId, r]));
        const cloudRegSet = new Set(cloudData.registrations.map((r) => r.regId));
        const deletes = [];
        for (const lr of localRegs) {
          if (scopedEventIds && !scopedEventIds.has(lr.eventId)) continue;
          if ((lr.synced || removedScopedEventIds.has(lr.eventId)) && !cloudRegSet.has(lr.regId)) {
            deletes.push(lr.regId);
          }
        }
        const puts = [];
        for (const r of cloudData.registrations) {
          const eventId = String(r.eventId || '').trim();
          const bib = String(r.bib || '').trim();
          if (!eventId || !bib) continue;
          const local = localMap.get(r.regId);
          if (local && !local.synced && local.updatedAt > (r.updatedAt || 0)) continue;
          const finalId = String(r.idNumber || local?.idNumber || '').trim().toUpperCase();
          const regId = String(r.regId || makeRegId(eventId, finalId, bib)).trim();
          const rawGender = r.gender !== undefined ? r.gender : (local?.gender ?? '');
          const finalDob = r.dob !== undefined ? normalizeDobValue(r.dob) : (local?.dob || '');
          const existingHash = r.idHash || (local && local.idNumber === finalId ? local.idHash : '') || '';
          puts.push({
            regId,
            eventId,
            bib,
            name: String(r.name ?? local?.name ?? '').trim(),
            lastName: String(r.lastName ?? local?.lastName ?? '').trim(),
            firstName: String(r.firstName ?? local?.firstName ?? '').trim(),
            gender: normalizeGenderValue(rawGender, finalId),
            dob: finalDob,
            city: String(r.city ?? local?.city ?? '').trim(),
            country: String(r.country ?? local?.country ?? '').trim(),
            idNumber: finalId,
            idMasked: r.idMasked || global.TimingEngine.maskIdNumber(finalId),
            idHash: existingHash,
            categoryId: String(r.categoryId ?? local?.categoryId ?? '').trim(),
            penaltySec: Number(r.penaltySec ?? local?.penaltySec ?? 0),
            updatedAt: Number(r.updatedAt || now),
            synced: 1
          });
        }
        batchPromises.push((async () => {
          // 僅對缺少 idHash 的紀錄平行計算 SHA-256
          await Promise.all(puts.map(async (rec) => {
            if (!rec.idHash && rec.idNumber) rec.idHash = await getHashCached(rec.idNumber);
          }));
          await this._runBatchTx('registrations', { deletes, puts });
        })());
      }

      if (Array.isArray(cloudData.scans) && localScans) {
        const localMap = new Map(localScans.map((s) => [s.id, s]));
        const cloudScanSet = new Set(cloudData.scans.map((s) => s.id));
        const deletes = [];
        for (const ls of localScans) {
          if (scopedEventIds && !scopedEventIds.has(ls.eventId)) continue;
          if ((ls.synced || removedScopedEventIds.has(ls.eventId)) && !cloudScanSet.has(ls.id)) {
            deletes.push(ls.id);
          }
        }
        const puts = [];
        for (const s of cloudData.scans) {
          const local = localMap.get(s.id);
          if (local && !local.synced && local.updatedAt > (s.updatedAt || 0)) continue;
          const ts = Number(s.timestampMs || 0);
          puts.push({
            ...s,
            timeStr: s.timeStr || global.TimingEngine.formatClockMs(ts),
            isoTime: s.isoTime || (ts ? new Date(ts).toISOString() : ''),
            synced: 1,
            updatedAt: Number(s.updatedAt || now)
          });
        }
        batchPromises.push(this._runBatchTx('scans', { deletes, puts }));
      }

      if (Array.isArray(cloudData.historyRecords) && localHist) {
        const localMap = new Map(localHist.map((h) => [h.id, h]));
        const cloudHistSet = new Set(cloudData.historyRecords.map((h) => h.id));
        const deletes = [];
        for (const lh of localHist) {
          if (lh.synced && !cloudHistSet.has(lh.id)) deletes.push(lh.id);
        }
        const toPut = [];
        for (const h of cloudData.historyRecords) {
          if (!h || !String(h.eventName || '').trim() || !String(h.name || '').trim()) continue;
          const local = localMap.get(h.id);
          if (local && !local.synced && local.updatedAt > (h.updatedAt || 0)) continue;
          toPut.push({
            ...h,
            idNumber: h.idNumber || local?.idNumber || '',
            idHash: h.idHash || (local && local.idNumber === (h.idNumber || local.idNumber) ? local.idHash : '') || ''
          });
        }
        batchPromises.push((async () => {
          const normalized = await Promise.all(toPut.map((item) => this.normalizeHistoryItem(item, true)));
          await this._runBatchTx('history_records', { deletes, puts: normalized });
        })());
      }

      await Promise.all(batchPromises);
    }

    /* ---------------- 完整系統備份 / 還原合併 (JSON) ---------------- */

    async exportFullBackup({ ownerAdmin = '', isSuperAdmin = true } = {}) {
      const [platform, chips, registrations, scans, historyRecords] = await Promise.all([
        this.getPlatform(),
        this.getAllChips(),
        this.getAllRegistrations(),
        this.getAllScans(),
        this.getAllHistoryRecords()
      ]);

      // 非最高管理權責時，僅匯出該承辦人名下資料；最高管理權責匯出全機完整資料
      if (ownerAdmin && !isSuperAdmin) {
        const ownEvents = (platform.events || []).filter((e) => (e.ownerAdmin || platform.adminUsername) === ownerAdmin);
        const ownEventIds = new Set(ownEvents.map((e) => e.eventId));
        const ownStaff = (platform.staffAccounts || []).filter((s) => String(s.ownerAdmin || '').trim() === ownerAdmin);
        return {
          schemaVersion: '3.13',
          exportedAt: new Date().toISOString(),
          scopedOwnerAdmin: ownerAdmin,
          platform: {
            ...platform,
            adminAccounts: (platform.adminAccounts || []).filter((a) => a.username === ownerAdmin),
            events: ownEvents,
            staffAccounts: ownStaff
          },
          chips,
          registrations: registrations.filter((r) => ownEventIds.has(r.eventId)),
          scans: scans.filter((s) => ownEventIds.has(s.eventId)),
          historyRecords: historyRecords.filter((h) => h.ownerAdmin === ownerAdmin)
        };
      }

      return {
        schemaVersion: '3.13',
        exportedAt: new Date().toISOString(),
        platform,
        chips,
        registrations,
        scans,
        historyRecords
      };
    }

    async importFullBackup(backupObj, strategy = 'replace', { ownerAdmin = '', isSuperAdmin = true } = {}) {
      if (!backupObj || typeof backupObj !== 'object' || (!backupObj.platform && !Array.isArray(backupObj.scans) && !Array.isArray(backupObj.registrations) && !Array.isArray(backupObj.chips) && !Array.isArray(backupObj.historyRecords))) {
        throw new Error('無效的系統備份檔格式（請選擇本系統下載的 .json 備份檔）');
      }
      await this.init();

      let importedEventCount = 0;
      if (strategy === 'replace' && isSuperAdmin) {
        await this.clearAllData(false);
        if (backupObj.platform) {
          const saved = await this.savePlatform({ ...createDefaultPlatform(), ...backupObj.platform });
          importedEventCount = (saved.events || []).length;
        }
      } else if (backupObj.platform) {
        const cur = await this.getPlatform();
        const eventMap = new Map((cur.events || []).map((e) => [e.eventId, e]));
        (backupObj.platform.events || []).forEach((e) => {
          if (!e?.eventId) return;
          const existing = eventMap.get(e.eventId);
          // 非最高管理員合併備份時，絕不覆蓋其他管理員的活動
          if (!isSuperAdmin && ownerAdmin && existing && (existing.ownerAdmin || cur.adminUsername) !== ownerAdmin) return;
          eventMap.set(e.eventId, !isSuperAdmin && ownerAdmin ? { ...e, ownerAdmin } : e);
          importedEventCount++;
        });
        const staffMap = new Map((cur.staffAccounts || []).map((s) => [s.username, s]));
        (backupObj.platform.staffAccounts || []).forEach((s) => {
          if (!s?.username) return;
          const existing = staffMap.get(s.username);
          if (!isSuperAdmin && ownerAdmin && existing && existing.ownerAdmin && existing.ownerAdmin !== ownerAdmin) return;
          staffMap.set(s.username, !isSuperAdmin && ownerAdmin ? { ...s, ownerAdmin } : s);
        });
        const adminMap = new Map((cur.adminAccounts || []).map((a) => [a.username, a]));
        if (isSuperAdmin && Array.isArray(backupObj.platform.adminAccounts)) {
          backupObj.platform.adminAccounts.forEach((a) => {
            if (!a?.username) return;
            // 合併還原保留目前既有帳密，僅補入備份中新增的管理員帳號
            if (!adminMap.has(a.username)) adminMap.set(a.username, a);
          });
        }
        await this.savePlatform({
          ...cur,
          adminAccounts: Array.from(adminMap.values()),
          events: Array.from(eventMap.values()),
          staffAccounts: Array.from(staffMap.values())
        });
      }

      const latestPlatform = await this.getPlatform();
      const allowedEventIds = (!isSuperAdmin && ownerAdmin)
        ? new Set((latestPlatform.events || []).filter((e) => (e.ownerAdmin || latestPlatform.adminUsername) === ownerAdmin).map((e) => e.eventId))
        : null;

      const now = Date.now();
      const tasks = [];
      let chipCount = 0, regCount = 0, scanCount = 0, historyCount = 0;

      // 1. 批次還原共用晶片
      const validChips = (backupObj.chips || []).filter((c) => c && String(c.idNumber || '').trim());
      if (validChips.length) {
        tasks.push((async () => {
          const existingChips = await this.getAllChips();
          const existMap = new Map(existingChips.map((c) => [c.idNumber, c]));
          const puts = validChips.map((c) => {
            const cleanId = String(c.idNumber).trim().toUpperCase();
            const existing = existMap.get(cleanId);
            return {
              idNumber: cleanId,
              name: String(c.name ?? existing?.name ?? '').trim(),
              nfcUid: String(c.nfcUid || existing?.nfcUid || '').trim().toUpperCase(),
              writtenAt: c.writtenAt || existing?.writtenAt || now,
              updatedAt: Number(c.updatedAt || now),
              synced: 0
            };
          });
          await this._runBatchTx('chips', { puts });
          chipCount = puts.length;
        })());
      }

      // 2. 批次還原報名名單
      const validRegs = (backupObj.registrations || []).filter((r) => {
        if (!r?.eventId || !r?.bib) return false;
        if (allowedEventIds && !allowedEventIds.has(r.eventId)) return false;
        return true;
      });
      if (validRegs.length) {
        tasks.push((async () => {
          const saved = await this.upsertRegistrationsBulk(validRegs, false);
          regCount = saved.length;
        })());
      }

      // 3. 批次還原感應紀錄
      const validScans = (backupObj.scans || []).filter((s) => {
        if (!s?.id || !s?.eventId || !s?.bib) return false;
        if (allowedEventIds && !allowedEventIds.has(s.eventId)) return false;
        return true;
      });
      if (validScans.length) {
        tasks.push((async () => {
          const puts = validScans.map((s) => {
            const ts = Number(s.timestampMs || 0);
            return {
              ...s,
              timeStr: s.timeStr || global.TimingEngine.formatClockMs(ts),
              isoTime: s.isoTime || (ts ? new Date(ts).toISOString() : ''),
              synced: 0,
              updatedAt: Number(s.updatedAt || now)
            };
          });
          await this._runBatchTx('scans', { puts });
          scanCount = puts.length;
        })());
      }

      // 4. 批次還原歷年紀錄
      if (Array.isArray(backupObj.historyRecords) && backupObj.historyRecords.length) {
        tasks.push((async () => {
          const savedHist = await this.upsertHistoryRecordsBulk(backupObj.historyRecords, false);
          historyCount = savedHist.length;
        })());
      }

      await Promise.all(tasks);
      return {
        eventCount: importedEventCount || (latestPlatform.events || []).length,
        chipCount,
        regCount,
        scanCount,
        historyCount
      };
    }

    /**
     * 載入 TUR (Taiwan Ultra Runners) 全情境完整測試資料（5 場不同性質／狀態活動、12 位跨場選手、DUV 欄位、分段與繞圈逾時測試、2023–2025 歷年與八百壯士英雄榜）
     */
    async loadDemoData(ownerAdmin = '') {
      await this.init();
      const prev = await this.getPlatform();
      const targetOwner = String(ownerAdmin || prev?.adminUsername || 'admin').trim() || 'admin';
      await this.clearAllData(false);

      const now = Date.now();
      const min = 60 * 1000;
      const gunEast = now - 660 * min;     // 進行中超馬距離賽（11小時前起跑）
      const gunUltra = now - 390 * min;    // 進行中繞圈賽（6.5小時前起跑，讓 6H 組示範逾時圈判定！）
      const gunWin25 = new Date('2025-11-16T05:00:00+08:00').getTime();
      const gunLap25 = new Date('2025-08-10T06:00:00+08:00').getTime();

      // 活動 1：進行中一般超馬距離賽（含 101.3K / 54.2K / 21K、4 站分段成績、外籍與複姓、罰時與完整 DUV 欄位）
      const evt1 = {
        eventId: 'EVT-2026-800-EAST',
        name: '2026 TUR 八百壯士極東霸主超級馬拉松',
        englishName: '800 Heroes Ultra Marathon - Overlord of easternmost point',
        organization: 'Taiwan Ultra Runners',
        date: getTodayStr(),
        location: '新北貢寮三貂角燈塔廣場',
        status: 'active',
        ownerAdmin: targetOwner,
        eventType: 'distance',
        lapDistanceKm: 0,
        debounceSeconds: 30,
        rankBy: 'gun',
        unifiedGunTimeMs: gunEast,
        categories: [
          { id: '101.3K', name: '超馬組 101.3K', distanceKm: 101.3, gunTimeMs: gunEast, gunTimeIso: new Date(gunEast).toISOString(), color: '#ef4444' },
          { id: '54.2K', name: '挑戰組 54.2K', distanceKm: 54.2, gunTimeMs: gunEast, gunTimeIso: new Date(gunEast).toISOString(), color: '#3b82f6' },
          { id: '21K', name: '半馬體驗組 21K', distanceKm: 21.1, gunTimeMs: gunEast, gunTimeIso: new Date(gunEast).toISOString(), color: '#10b981' }
        ],
        checkpoints: [
          { id: 'CP1', name: '補給站 1 (20K)', order: 1, distanceKm: 20, type: 'checkpoint' },
          { id: 'CP2', name: '極東折返點 (50K)', order: 2, distanceKm: 50, type: 'checkpoint' },
          { id: 'CP3', name: '補給站 3 (75K)', order: 3, distanceKm: 75, type: 'checkpoint' },
          { id: 'FINISH', name: '終點衝線門', order: 4, distanceKm: 101.3, type: 'finish' }
        ],
        updatedAt: now
      };

      // 活動 2：進行中繞圈計時賽（單圈 2.54km，6H/12H/24H，含逐圈明細與 6H 逾時圈示範）
      const evt2 = createLapTimedEvent('EVT-2026-800-ULTRA', '2026 TUR 八百壯士環湖繞圈計時賽', getTodayStr(), '碧潭環湖賽道 (單圈 2.54K)', 2.54, targetOwner);
      evt2.englishName = '800 Heroes Ultra Marathon - Bitan Timed Loop';
      evt2.organization = 'Taiwan Ultra Runners';
      evt2.unifiedGunTimeMs = gunUltra;
      evt2.categories.forEach((c) => {
        c.gunTimeMs = gunUltra;
        c.gunTimeIso = new Date(gunUltra).toISOString();
      });

      // 活動 3：籌備中一般距離賽（尚未鳴槍，含部分預先報到，方便測試賽前報到與一鍵鳴槍）
      const evt3 = {
        eventId: 'EVT-2026-COAST-PREP',
        name: '2026 TUR 北海岸公益馬拉松（籌備中）',
        englishName: 'TUR North Coast Charity Marathon',
        organization: 'Taiwan Ultra Runners',
        date: getTodayStr(),
        location: '淡水漁人碼頭觀海廣場',
        status: 'upcoming',
        ownerAdmin: targetOwner,
        eventType: 'distance',
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
          { id: 'CP1', name: '淺水灣補給站 (10K)', order: 1, distanceKm: 10, type: 'checkpoint' },
          { id: 'CP2', name: '富貴角折返點 (21K)', order: 2, distanceKm: 21.1, type: 'checkpoint' },
          { id: 'FINISH', name: '終點衝線門', order: 3, distanceKm: 42.195, type: 'finish' }
        ],
        updatedAt: now
      };

      // 活動 4：已結束距離賽（鎖定狀態＋自動歸檔至歷年成績、八百壯士英雄榜與證書列印）
      const evt4 = {
        eventId: 'EVT-2025-800-WINTER',
        name: '2025 TUR 八百壯士冬季百公里挑戰賽',
        englishName: '2025 800 Heroes Winter 100K Challenge',
        organization: 'Taiwan Ultra Runners',
        date: '2025-11-16',
        location: '花蓮七星潭海岸步道',
        status: 'finished',
        ownerAdmin: targetOwner,
        eventType: 'distance',
        lapDistanceKm: 0,
        debounceSeconds: 30,
        rankBy: 'gun',
        unifiedGunTimeMs: gunWin25,
        categories: [
          { id: '100K', name: '百公里組 100K', distanceKm: 100, gunTimeMs: gunWin25, gunTimeIso: new Date(gunWin25).toISOString(), color: '#ef4444' },
          { id: '50K', name: '五十公里組 50K', distanceKm: 50, gunTimeMs: gunWin25, gunTimeIso: new Date(gunWin25).toISOString(), color: '#3b82f6' }
        ],
        checkpoints: [
          { id: 'CP1', name: '海濱折返點 (25K)', order: 1, distanceKm: 25, type: 'checkpoint' },
          { id: 'CP2', name: '鯉魚潭補給站 (50K)', order: 2, distanceKm: 50, type: 'checkpoint' },
          { id: 'CP3', name: '溪口補給站 (75K)', order: 3, distanceKm: 75, type: 'checkpoint' },
          { id: 'FINISH', name: '終點衝線門', order: 4, distanceKm: 100, type: 'finish' }
        ],
        updatedAt: now - 1000
      };

      // 活動 5：已結束繞圈計時賽（鎖定狀態＋自動歸檔至歷年成績與證書列印）
      const evt5 = createLapTimedEvent('EVT-2025-TAIPEI-12H', '2025 TUR 台北河濱 12H 繞圈經典賽', '2025-08-10', '台北大佳河濱公園 (單圈 2.00K)', 2.0, targetOwner);
      evt5.englishName = '2025 TUR Taipei Riverside 12H Classic';
      evt5.organization = 'Taiwan Ultra Runners';
      evt5.status = 'finished';
      evt5.unifiedGunTimeMs = gunLap25;
      evt5.categories = [
        { id: '12H', name: '12小時繞圈組', durationHours: 12, distanceKm: 12, gunTimeMs: gunLap25, gunTimeIso: new Date(gunLap25).toISOString(), color: '#3b82f6' },
        { id: '6H', name: '6小時繞圈組', durationHours: 6, distanceKm: 6, gunTimeMs: gunLap25, gunTimeIso: new Date(gunLap25).toISOString(), color: '#10b981' }
      ];

      const platform = createDefaultPlatform();
      platform.adminUsername = prev?.adminUsername || 'admin';
      platform.adminPassword = prev?.adminPassword || 'admin1234';
      if (prev?.adminAccounts?.length) platform.adminAccounts = [...prev.adminAccounts];
      if (!platform.adminAccounts.some((a) => a && !a.isSuperAdmin && a.enabled !== false)) {
        platform.adminAccounts.push({
          username: 'host01',
          password: '1234',
          displayName: '活動承辦人（範例）',
          isSuperAdmin: false,
          enabled: true,
          updatedAt: Date.now()
        });
      }
      if (prev?.staffAccounts?.length) platform.staffAccounts = prev.staffAccounts;
      platform.activeEventId = evt1.eventId;
      platform.events = [evt1, evt2, evt3, evt4, evt5];
      await this.savePlatform(platform);

      // 12 位共用晶片選手主檔（涵蓋台灣男女、複姓、外籍選手與 1 位測試缺 DOB 選手）
      const people = [
        { idNumber: 'A123456789', name: '林志明', nfcUid: '04:A1:10:01', gender: 'M', dob: '1984-01-18', city: 'Taipei', country: 'Taiwan' },
        { idNumber: 'B234567891', name: '陳潔馨', nfcUid: '04:A1:10:02', gender: 'F', dob: '1980-02-07', city: 'Taipei', country: 'Taiwan' },
        { idNumber: 'C145678912', name: '張家豪', nfcUid: '04:A1:10:03', gender: 'M', dob: '1975-09-13', city: 'New Taipei', country: 'Taiwan' },
        { idNumber: 'F256789123', name: '王雅婷', nfcUid: '04:A1:10:04', gender: 'F', dob: '1990-03-15', city: 'Taichung', country: 'Taiwan' },
        { idNumber: 'H167891234', name: '鄭順元', nfcUid: '04:A1:10:05', gender: 'M', dob: '1977-06-28', city: 'Taoyuan', country: 'Taiwan' },
        { idNumber: 'T189123456', name: '吳冠宇', nfcUid: '04:A1:10:06', gender: 'M', dob: '1983-08-28', city: 'Kaohsiung', country: 'Taiwan' },
        { idNumber: 'N291234567', name: '劉佳穎', nfcUid: '04:A1:10:07', gender: 'F', dob: '1988-11-22', city: 'Tainan', country: 'Taiwan' },
        { idNumber: 'D155667788', name: '歐陽志遠', nfcUid: '04:A1:10:08', gender: 'M', dob: '1979-05-12', city: 'Taipei', country: 'Taiwan' },
        { idNumber: 'E800012345', name: 'Emmanuel Lemire', lastName: 'Lemire', firstName: 'Emmanuel', nfcUid: '04:A1:10:09', gender: 'M', dob: '1993-07-15', city: 'Taipei', country: 'France' },
        { idNumber: 'K277889900', name: '黃瓊華', nfcUid: '04:A1:10:10', gender: 'F', dob: '1980-08-28', city: 'Hsinchu', country: 'Taiwan' },
        { idNumber: 'M133445566', name: '陳偉庭', nfcUid: '04:A1:10:11', gender: 'M', dob: '1984-01-18', city: 'Taipei', country: 'Taiwan' },
        { idNumber: 'P199887766', name: '林柏翰', nfcUid: '04:A1:10:12', gender: 'M', dob: '', city: 'Keelung', country: 'Taiwan' }
      ];
      const peopleMap = new Map(people.map((p) => [p.idNumber, p]));
      for (const p of people) {
        await this.upsertChip(p, false);
      }
      const withDemoBio = (reg) => {
        const p = peopleMap.get(reg.idNumber);
        return p
          ? { ...reg, gender: p.gender, dob: p.dob, city: p.city, country: p.country, lastName: p.lastName || '', firstName: p.firstName || '' }
          : reg;
      };

      // 活動 1（進行中 101.3K/54.2K/21K）：10 位選手（完賽＋賽道上＋已報到＋未報到＋罰時＋缺DOB提醒）
      await this.upsertRegistrationsBulk([
        withDemoBio({ eventId: evt1.eventId, bib: 'A01', name: '陳偉庭', idNumber: 'M133445566', categoryId: '101.3K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'A02', name: '林志明', idNumber: 'A123456789', categoryId: '101.3K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'A03', name: '歐陽志遠', idNumber: 'D155667788', categoryId: '101.3K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'A04', name: '陳潔馨', idNumber: 'B234567891', categoryId: '101.3K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'A05', name: '張家豪', idNumber: 'C145678912', categoryId: '101.3K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'B01', name: 'Emmanuel Lemire', idNumber: 'E800012345', categoryId: '54.2K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'B02', name: '黃瓊華', idNumber: 'K277889900', categoryId: '54.2K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'B03', name: '王雅婷', idNumber: 'F256789123', categoryId: '54.2K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'C01', name: '吳冠宇', idNumber: 'T189123456', categoryId: '21K', penaltySec: 30 }),
        withDemoBio({ eventId: evt1.eventId, bib: 'C02', name: '林柏翰', idNumber: 'P199887766', categoryId: '21K' }),
        withDemoBio({ eventId: evt1.eventId, bib: 'C03', name: '劉佳穎', idNumber: 'N291234567', categoryId: '21K' })
      ], false);

      // 活動 2（進行中繞圈賽 6H/12H/24H）：7 位選手
      await this.upsertRegistrationsBulk([
        withDemoBio({ eventId: evt2.eventId, bib: 'U001', name: '林志明', idNumber: 'A123456789', categoryId: '12H' }),
        withDemoBio({ eventId: evt2.eventId, bib: 'U002', name: '陳潔馨', idNumber: 'B234567891', categoryId: '12H' }),
        withDemoBio({ eventId: evt2.eventId, bib: 'U101', name: '張家豪', idNumber: 'C145678912', categoryId: '24H' }),
        withDemoBio({ eventId: evt2.eventId, bib: 'U102', name: '鄭順元', idNumber: 'H167891234', categoryId: '24H' }),
        withDemoBio({ eventId: evt2.eventId, bib: 'U201', name: '王雅婷', idNumber: 'F256789123', categoryId: '6H' }),
        withDemoBio({ eventId: evt2.eventId, bib: 'U202', name: '吳冠宇', idNumber: 'T189123456', categoryId: '6H' }),
        withDemoBio({ eventId: evt2.eventId, bib: 'U203', name: '劉佳穎', idNumber: 'N291234567', categoryId: '6H' })
      ], false);

      // 活動 3（籌備中距離賽）：5 位選手（未開跑，2 人已預先報到）
      await this.upsertRegistrationsBulk([
        withDemoBio({ eventId: evt3.eventId, bib: 'M001', name: '林志明', idNumber: 'A123456789', categoryId: '42K' }),
        withDemoBio({ eventId: evt3.eventId, bib: 'M002', name: '陳偉庭', idNumber: 'M133445566', categoryId: '42K' }),
        withDemoBio({ eventId: evt3.eventId, bib: 'H101', name: '黃瓊華', idNumber: 'K277889900', categoryId: '21K' }),
        withDemoBio({ eventId: evt3.eventId, bib: 'H102', name: 'Emmanuel Lemire', idNumber: 'E800012345', categoryId: '21K' }),
        withDemoBio({ eventId: evt3.eventId, bib: 'K201', name: '劉佳穎', idNumber: 'N291234567', categoryId: '10K' })
      ], false);

      // 活動 4（已結束 2025 八百壯士百公里）：6 位完賽選手
      await this.upsertRegistrationsBulk([
        withDemoBio({ eventId: evt4.eventId, bib: 'W001', name: '林志明', idNumber: 'A123456789', categoryId: '100K' }),
        withDemoBio({ eventId: evt4.eventId, bib: 'W002', name: '張家豪', idNumber: 'C145678912', categoryId: '100K' }),
        withDemoBio({ eventId: evt4.eventId, bib: 'W003', name: '歐陽志遠', idNumber: 'D155667788', categoryId: '100K' }),
        withDemoBio({ eventId: evt4.eventId, bib: 'W101', name: '陳潔馨', idNumber: 'B234567891', categoryId: '50K' }),
        withDemoBio({ eventId: evt4.eventId, bib: 'W102', name: '王雅婷', idNumber: 'F256789123', categoryId: '50K' }),
        withDemoBio({ eventId: evt4.eventId, bib: 'W103', name: '鄭順元', idNumber: 'H167891234', categoryId: '50K' })
      ], false);

      // 活動 5（已結束 2025 台北河濱 12H 繞圈）：4 位完賽選手
      await this.upsertRegistrationsBulk([
        withDemoBio({ eventId: evt5.eventId, bib: 'R01', name: '林志明', idNumber: 'A123456789', categoryId: '12H' }),
        withDemoBio({ eventId: evt5.eventId, bib: 'R02', name: '張家豪', idNumber: 'C145678912', categoryId: '12H' }),
        withDemoBio({ eventId: evt5.eventId, bib: 'R11', name: '王雅婷', idNumber: 'F256789123', categoryId: '6H' }),
        withDemoBio({ eventId: evt5.eventId, bib: 'R12', name: '吳冠宇', idNumber: 'T189123456', categoryId: '6H' })
      ], false);

      const addS = (eventId, bib, idNumber, checkpointId, timestampMs) =>
        this.addScan({ eventId, bib, idNumber, checkpointId, timestampMs, source: 'mock', operator: 'demo', bypassDebounce: true });

      // --- 活動 1 感應資料（101.3K / 54.2K / 21K 分段完整） ---
      const addDistRun = async (evId, gun, bib, idNum, cpOffsets) => {
        await addS(evId, bib, idNum, 'CHECKIN', gun - 30 * min);
        for (const [cpId, offsetMs] of cpOffsets) {
          await addS(evId, bib, idNum, cpId, gun + offsetMs);
        }
      };
      await addDistRun(evt1.eventId, gunEast, 'A01', 'M133445566', [
        ['CP1', 112 * min + 14000], ['CP2', 289 * min + 35000], ['CP3', 442 * min + 10000], ['FINISH', 609 * min + 28000]
      ]);
      await addDistRun(evt1.eventId, gunEast, 'A02', 'A123456789', [
        ['CP1', 118 * min + 22000], ['CP2', 302 * min + 18000], ['CP3', 462 * min + 40000], ['FINISH', 636 * min + 15000]
      ]);
      await addDistRun(evt1.eventId, gunEast, 'A03', 'D155667788', [
        ['CP1', 120 * min + 5000], ['CP2', 309 * min + 42000], ['CP3', 475 * min + 12000], ['FINISH', 648 * min + 29000]
      ]);
      await addDistRun(evt1.eventId, gunEast, 'A04', 'B234567891', [
        ['CP1', 126 * min + 31000], ['CP2', 322 * min + 10000], ['CP3', 494 * min + 55000], ['FINISH', 655 * min + 42000]
      ]);
      // A05 賽道上（已通過 CP1、CP2，展示即時賽道動態與 ETA！）
      await addDistRun(evt1.eventId, gunEast, 'A05', 'C145678912', [
        ['CP1', 135 * min + 12000], ['CP2', 348 * min + 20000]
      ]);
      // 54.2K 完賽（外籍＋女選手）與已報到未通過
      await addDistRun(evt1.eventId, gunEast, 'B01', 'E800012345', [
        ['CP1', 114 * min + 8000], ['CP2', 298 * min + 44000], ['FINISH', 333 * min + 16000]
      ]);
      await addDistRun(evt1.eventId, gunEast, 'B02', 'K277889900', [
        ['CP1', 138 * min + 15000], ['CP2', 378 * min + 20000], ['FINISH', 415 * min + 36000]
      ]);
      await addS(evt1.eventId, 'B03', 'F256789123', 'CHECKIN', gunEast - 20 * min);
      // 21K 完賽（C01 含 +30s 罰時、C02 缺 DOB 測試 DUV 提醒、C03 未報到）
      await addDistRun(evt1.eventId, gunEast, 'C01', 'T189123456', [
        ['CP1', 96 * min + 20000], ['FINISH', 102 * min + 15000]
      ]);
      await addDistRun(evt1.eventId, gunEast, 'C02', 'P199887766', [
        ['CP1', 108 * min + 40000], ['FINISH', 115 * min + 8000]
      ]);

      // --- 活動 2 感應資料（繞圈計時賽，含 6H 組 1 筆逾時圈示範！） ---
      for (const [bib, idNum] of [
        ['U101', 'C145678912'], ['U102', 'H167891234'], ['U001', 'A123456789'],
        ['U002', 'B234567891'], ['U201', 'F256789123'], ['U202', 'T189123456'], ['U203', 'N291234567']
      ]) {
        await addS(evt2.eventId, bib, idNum, 'CHECKIN', gunUltra - 25 * min);
      }
      const addLaps = async (evId, bib, idNum, startMs, count, avgLapMin) => {
        let cur = startMs;
        for (let i = 1; i <= count; i++) {
          cur += Math.round((avgLapMin + ((i % 3) - 1) * 0.6) * min);
          await addS(evId, bib, idNum, 'LAP', cur);
        }
        return cur;
      };
      await addLaps(evt2.eventId, 'U101', 'C145678912', gunUltra, 18, 15.1);
      await addLaps(evt2.eventId, 'U102', 'H167891234', gunUltra, 15, 16.0);
      await addLaps(evt2.eventId, 'U001', 'A123456789', gunUltra, 17, 14.4);
      await addLaps(evt2.eventId, 'U002', 'B234567891', gunUltra, 14, 15.7);
      // U201 (6H 組)：跑完 16 圈在 6H 內，再補 1 筆第 6H12m 的「超時不計圈」示範！
      await addLaps(evt2.eventId, 'U201', 'F256789123', gunUltra, 16, 21.5);
      await addS(evt2.eventId, 'U201', 'F256789123', 'LAP', gunUltra + 372 * min);
      await addLaps(evt2.eventId, 'U202', 'T189123456', gunUltra, 14, 22.8);

      // --- 活動 3 感應資料（籌備中：僅預先報到 2 位） ---
      await addS(evt3.eventId, 'M001', 'A123456789', 'CHECKIN', now - 12 * min);
      await addS(evt3.eventId, 'H101', 'K277889900', 'CHECKIN', now - 8 * min);

      // --- 活動 4 感應資料（已結束 2025 八百壯士冬季百公里） ---
      await addDistRun(evt4.eventId, gunWin25, 'W001', 'A123456789', [
        ['CP1', 136 * min], ['CP2', 280 * min], ['CP3', 428 * min], ['FINISH', 578 * min + 15000]
      ]);
      await addDistRun(evt4.eventId, gunWin25, 'W002', 'C145678912', [
        ['CP1', 142 * min], ['CP2', 296 * min], ['CP3', 454 * min], ['FINISH', 612 * min + 40000]
      ]);
      await addDistRun(evt4.eventId, gunWin25, 'W003', 'D155667788', [
        ['CP1', 148 * min], ['CP2', 308 * min], ['CP3', 470 * min], ['FINISH', 635 * min + 12000]
      ]);
      await addDistRun(evt4.eventId, gunWin25, 'W101', 'B234567891', [
        ['CP1', 132 * min], ['CP2', 274 * min], ['FINISH', 282 * min + 10000]
      ]);
      await addDistRun(evt4.eventId, gunWin25, 'W102', 'F256789123', [
        ['CP1', 140 * min], ['CP2', 295 * min], ['FINISH', 305 * min + 22000]
      ]);
      await addDistRun(evt4.eventId, gunWin25, 'W103', 'H167891234', [
        ['CP1', 145 * min], ['CP2', 304 * min], ['FINISH', 318 * min + 45000]
      ]);

      // --- 活動 5 感應資料（已結束 2025 台北河濱 12H 繞圈） ---
      await addS(evt5.eventId, 'R01', 'A123456789', 'CHECKIN', gunLap25 - 25 * min);
      await addS(evt5.eventId, 'R02', 'C145678912', 'CHECKIN', gunLap25 - 22 * min);
      await addS(evt5.eventId, 'R11', 'F256789123', 'CHECKIN', gunLap25 - 18 * min);
      await addS(evt5.eventId, 'R12', 'T189123456', 'CHECKIN', gunLap25 - 15 * min);
      await addLaps(evt5.eventId, 'R01', 'A123456789', gunLap25, 52, 13.4);
      await addLaps(evt5.eventId, 'R02', 'C145678912', gunLap25, 48, 14.5);
      await addLaps(evt5.eventId, 'R11', 'F256789123', gunLap25, 26, 13.5);
      await addLaps(evt5.eventId, 'R12', 'T189123456', gunLap25, 24, 14.2);

      // 歷年匯入成績（2023–2025 跨多年度八百壯士與一般超馬，完整展示個人歷年合併與英雄榜！）
      await this.upsertHistoryRecordsBulk([
        { eventDate: '2025-05-10', eventName: '2025 TUR 陽明山百公里越野賽', categoryName: '50K', bib: 'Y010', name: '林志明', idNumber: 'A123456789', finishTime: '05:18:40', distanceKm: 50, laps: 0, overallRank: 3, categoryRank: 2, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2025-05-10', eventName: '2025 TUR 陽明山百公里越野賽', categoryName: '50K', bib: 'Y011', name: '王雅婷', idNumber: 'F256789123', finishTime: '05:45:12', distanceKm: 50, laps: 0, overallRank: 6, categoryRank: 1, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2025-05-10', eventName: '2025 TUR 陽明山百公里越野賽', categoryName: '100K', bib: 'Y001', name: '陳偉庭', idNumber: 'M133445566', finishTime: '11:02:18', distanceKm: 100, laps: 0, overallRank: 1, categoryRank: 1, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2024-11-17', eventName: '2024 TUR 八百壯士經典挑戰賽', categoryName: '100K', bib: 'H001', name: '林志明', idNumber: 'A123456789', finishTime: '09:55:08', distanceKm: 100, laps: 0, overallRank: 1, categoryRank: 1, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2024-11-17', eventName: '2024 TUR 八百壯士經典挑戰賽', categoryName: '100K', bib: 'H002', name: '張家豪', idNumber: 'C145678912', finishTime: '10:25:33', distanceKm: 100, laps: 0, overallRank: 2, categoryRank: 2, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2024-11-17', eventName: '2024 TUR 八百壯士經典挑戰賽', categoryName: '100K', bib: 'H003', name: '陳偉庭', idNumber: 'M133445566', finishTime: '10:31:05', distanceKm: 100, laps: 0, overallRank: 3, categoryRank: 3, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2024-11-17', eventName: '2024 TUR 八百壯士經典挑戰賽', categoryName: '50K', bib: 'H101', name: '陳潔馨', idNumber: 'B234567891', finishTime: '04:51:19', distanceKm: 50, laps: 0, overallRank: 1, categoryRank: 1, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2024-11-17', eventName: '2024 TUR 八百壯士經典挑戰賽', categoryName: '50K', bib: 'H102', name: '歐陽志遠', idNumber: 'D155667788', finishTime: '05:08:44', distanceKm: 50, laps: 0, overallRank: 2, categoryRank: 1, status: '完賽', ownerAdmin: targetOwner },
        { eventDate: '2023-11-19', eventName: '2023 TUR 八百壯士創始紀念超馬', categoryName: '24H', bib: 'G001', name: '林志明', idNumber: 'A123456789', finishTime: '23:48:10', distanceKm: 162.56, laps: 64, overallRank: 1, categoryRank: 1, status: '完賽', note: '單圈 2.54km', ownerAdmin: targetOwner },
        { eventDate: '2023-11-19', eventName: '2023 TUR 八百壯士創始紀念超馬', categoryName: '24H', bib: 'G002', name: '張家豪', idNumber: 'C145678912', finishTime: '23:52:40', distanceKm: 154.94, laps: 61, overallRank: 2, categoryRank: 2, status: '完賽', note: '單圈 2.54km', ownerAdmin: targetOwner },
        { eventDate: '2023-11-19', eventName: '2023 TUR 八百壯士創始紀念超馬', categoryName: '12H', bib: 'G101', name: '鄭順元', idNumber: 'H167891234', finishTime: '11:45:22', distanceKm: 96.52, laps: 38, overallRank: 1, categoryRank: 1, status: '完賽', note: '單圈 2.54km', ownerAdmin: targetOwner },
        { eventDate: '2023-11-19', eventName: '2023 TUR 八百壯士創始紀念超馬', categoryName: '12H', bib: 'G102', name: '陳潔馨', idNumber: 'B234567891', finishTime: '11:50:15', distanceKm: 88.9, laps: 35, overallRank: 2, categoryRank: 1, status: '完賽', note: '單圈 2.54km', ownerAdmin: targetOwner }
      ], false);

      this.saveDeviceSettings({ selectedEventId: evt1.eventId });
      return true;
    }

    async clearAllData(resetPlatform = true) {
      await this.init();
      await Promise.all([
        this._req(this._tx('chips', 'readwrite').clear()),
        this._req(this._tx('registrations', 'readwrite').clear()),
        this._req(this._tx('scans', 'readwrite').clear()),
        this._req(this._tx('history_records', 'readwrite').clear())
      ]);
      if (resetPlatform) {
        await this.savePlatform(createDefaultPlatform());
      }
    }
  }

  global.RaceDB = new RaceDB();
})(window);
