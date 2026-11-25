/**
 * 智慧路跑 NFC 晶片計時平台 v3.2 — 全自動多裝置雲端同步模組
 * 修正重點 (v3.2)：
 * 1. GET 請求全面加入 cache: 'no-store' 與 _t 時間戳防快取，確保各裝置永遠拿到最新雲端資料
 * 2. 新增 syncPublicSnapshot()：未登入前進入大廳只要網路正常即自動同步雲端所有活動清單
 * 3. syncPendingToCloud(session) 若為最高管理員自動一併上傳所有活動/分組/鳴槍/站點與作業員帳號
 * 4. 新增 fullBidirectionalSync(session)：先推本機未同步異動，再拉最新雲端完整快照，確保所有平台資料 100% 一致
 */
(function (global) {
  'use strict';

  class CloudSync {
    getUrl() {
      return (global.RaceDB.getDeviceSettings().gasUrl || '').trim();
    }

    hasUrl() {
      return Boolean(this.getUrl());
    }

    isOnlineEnabled() {
      const s = global.RaceDB.getDeviceSettings();
      return s.mode === 'online' && Boolean(s.gasUrl && s.gasUrl.trim());
    }

    async _get(url, params = {}, timeoutMs = 18000) {
      const target = new URL(url);
      Object.entries({ ...params, _t: Date.now() }).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') target.searchParams.set(k, String(v));
      });
      const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
      const timer = ctrl ? setTimeout(() => ctrl.abort(), timeoutMs) : null;
      try {
        const resp = await fetch(target.toString(), {
          method: 'GET',
          redirect: 'follow',
          cache: 'no-store',
          signal: ctrl ? ctrl.signal : undefined
        });
        if (!resp.ok) throw new Error(`HTTP 錯誤 (${resp.status})`);
        const json = await resp.json();
        if (!json.ok) throw new Error(json.error || '雲端 API 回傳錯誤');
        return json;
      } catch (err) {
        if (err && err.name === 'AbortError') throw new Error('雲端連線逾時，請稍後重試');
        throw err;
      } finally {
        if (timer) clearTimeout(timer);
      }
    }

    async _post(url, body, timeoutMs = 20000) {
      const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
      const timer = ctrl ? setTimeout(() => ctrl.abort(), timeoutMs) : null;
      try {
        const resp = await fetch(url, {
          method: 'POST',
          redirect: 'follow',
          cache: 'no-store',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(body),
          signal: ctrl ? ctrl.signal : undefined
        });
        if (!resp.ok) throw new Error(`HTTP 錯誤 (${resp.status})`);
        const json = await resp.json();
        if (!json.ok) {
          const err = new Error(json.error || '雲端 API 寫入失敗');
          err.multiAdmin = Boolean(json.multiAdmin);
          throw err;
        }
        return json;
      } catch (err) {
        if (err && err.name === 'AbortError') throw new Error('雲端請求逾時，請稍後重試');
        throw err;
      } finally {
        if (timer) clearTimeout(timer);
      }
    }

    _buildScopedPlatformPayload(platform, auth, extra = {}) {
      const owner = String(extra?.targetOwner || auth?.username || platform.adminUsername || 'admin').trim();
      const defaultOwner = String(platform.adminUsername || 'admin').trim();
      const adminList = platform.adminAccounts || [];
      const callerAcc = adminList.find((a) => a.username === owner);
      const isSuper = Boolean(callerAcc ? callerAcc.isSuperAdmin : owner === defaultOwner);

      const scopedEvents = (platform.events || []).filter(
        (e) => (String(e.ownerAdmin || '').trim() || defaultOwner) === owner
      );
      const scopedStaff = isSuper
        ? (platform.staffAccounts || [])
        : (platform.staffAccounts || []).filter((s) => String(s.ownerAdmin || '').trim() === owner);
      const scopedAdmins = isSuper
        ? adminList
        : adminList.filter((a) => a.username === owner);

      const restExtra = { ...(extra || {}) };
      delete restExtra.targetOwner;

      return {
        scopedAdminSync: true,
        manageAllStaff: isSuper,
        globalConfig: {
          adminUsername: platform.adminUsername,
          adminPassword: platform.adminPassword,
          activeEventId: platform.activeEventId
        },
        adminAccounts: scopedAdmins,
        events: scopedEvents,
        staffAccounts: scopedStaff,
        ...restExtra
      };
    }

    async ping(customUrl) {
      const url = (customUrl || this.getUrl()).trim();
      if (!url) throw new Error('請先輸入 Google Apps Script 網頁應用程式 URL');
      return this._get(url, { action: 'ping' });
    }

    async fetchPublicData(customUrl) {
      const url = (customUrl || this.getUrl()).trim();
      if (!url) throw new Error('未設定雲端 API 網址');
      return this._get(url, { action: 'publicData' });
    }

    /**
     * 進入系統大廳時自動同步雲端公開活動清單、分組與最新紀錄（免登入即可保持活動清單一致）
     */
    async syncPublicSnapshot() {
      const url = this.getUrl();
      if (!url || !navigator.onLine) return null;
      const res = await this._get(url, { action: 'publicData' });
      if (res.data) {
        await global.RaceDB.mergeCloudData(res.data);
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      }
      return res;
    }

    async loginCloud(mode, username, password, { lite = false } = {}) {
      const url = this.getUrl();
      if (!url) throw new Error('尚未設定雲端 API 網址');
      const res = await this._post(url, { action: 'login', mode, username, password, lite: Boolean(lite) });
      if (mode === 'admin') {
        const acc = res.account || { username, isSuperAdmin: true, displayName: username, role: 'admin' };
        if (res.data) {
          const opts = res.multiAdmin
            ? { scopedOwnerAdmin: acc.username, scopedAdminOnly: !acc.isSuperAdmin }
            : {};
          await global.RaceDB.mergeCloudData(res.data, opts);
        }
        await global.RaceDB.cacheAdminAccount({ ...acc, password });
      } else {
        if (res.data) await global.RaceDB.mergeCloudData(res.data);
        if (res.account) await global.RaceDB.cacheStaffAccount(res.account);
      }
      if (!lite || res.data) {
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      }
      return res;
    }

    async pullFromCloud(auth) {
      const url = this.getUrl();
      if (!url) throw new Error('請先設定 Google Apps Script 網址');
      const res = await this._post(url, { action: 'fullSync', auth });
      const opts = res.isAdmin && res.multiAdmin && res.adminUsername
        ? { scopedOwnerAdmin: res.adminUsername, scopedAdminOnly: !res.isSuperAdmin }
        : {};
      await global.RaceDB.mergeCloudData(res.data, opts);
      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return res;
    }

    /**
     * 上傳本機所有待同步資料（含管理員修改的活動/站點/作業員帳號 + 晶片 + 報名 + 感應）
     */
    async syncPendingToCloud(session = null) {
      const url = this.getUrl();
      if (!url) throw new Error('請先設定 Google Apps Script 網址');

      const pending = await global.RaceDB.getPendingSyncCounts();
      const isAdmin = session && session.mode === 'admin' && session.auth;
      const includePlatform = Boolean(isAdmin && pending.platformDirty);

      let regsToUpload = pending.unsyncedRegs;
      let scansToUpload = pending.unsyncedScans;
      let historyToUpload = isAdmin ? (pending.unsyncedHistory || []) : [];
      let platformPayload = null;

      if (isAdmin) {
        const platform = await global.RaceDB.getPlatform();
        const owner = String(session.auth.username || '').trim();
        const defaultOwner = String(platform.adminUsername || 'admin').trim();
        const ownEventIds = new Set(
          (platform.events || [])
            .filter((e) => (String(e.ownerAdmin || '').trim() || defaultOwner) === owner)
            .map((e) => e.eventId)
        );
        regsToUpload = pending.unsyncedRegs.filter((r) => ownEventIds.has(r.eventId));
        scansToUpload = pending.unsyncedScans.filter((s) => ownEventIds.has(s.eventId));
        if (includePlatform) {
          platformPayload = this._buildScopedPlatformPayload(platform, session.auth);
        }
      }

      if (!includePlatform && pending.chips === 0 && regsToUpload.length === 0 && scansToUpload.length === 0 && historyToUpload.length === 0) {
        return { platformSynced: false, uploadedChips: 0, uploadedRegs: 0, uploadedScans: 0, uploadedHistory: 0, skippedDuplicate: 0 };
      }

      const body = {
        action: 'batchSync',
        chips: pending.unsyncedChips,
        registrations: regsToUpload,
        scans: scansToUpload,
        historyRecords: historyToUpload
      };

      if (session && session.auth) {
        body.auth = session.auth;
      }
      if (platformPayload) {
        body.platform = platformPayload;
      }

      const res = await this._post(url, body);

      const scanDoneIds = [
        ...(res.result?.scans?.acceptedIds || []),
        ...(res.result?.scans?.duplicateIds || [])
      ];

      await global.RaceDB.markSynced({
        platform: Boolean(res.result?.platformSynced),
        chipIds: pending.unsyncedChips.map((c) => c.idNumber),
        regIds: regsToUpload.map((r) => r.regId),
        scanIds: scanDoneIds,
        historyIds: historyToUpload.map((h) => h.id)
      });

      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return {
        platformSynced: Boolean(res.result?.platformSynced),
        purgedRegistrations: res.result?.purge?.deletedRegistrations || 0,
        purgedScans: res.result?.purge?.deletedScans || 0,
        uploadedChips: pending.unsyncedChips.length,
        uploadedRegs: regsToUpload.length,
        uploadedScans: res.result?.scans?.inserted || 0,
        uploadedHistory: historyToUpload.length,
        skippedDuplicate: res.result?.scans?.skippedDuplicate || 0
      };
    }

    /**
     * 完整雙向同步：先上傳本機未同步資料，再拉取雲端最新全平台資料
     * 確保電腦、手機與公開端資料完全一致！
     */
    async fullBidirectionalSync(session = null) {
      if (!this.hasUrl() || !navigator.onLine) {
        throw new Error('無網路或尚未設定雲端網址');
      }
      const uploadRes = await this.syncPendingToCloud(session);
      if (session && session.auth) {
        await this.pullFromCloud(session.auth);
      } else {
        await this.syncPublicSnapshot();
      }
      return uploadRes;
    }

    async pushSingleScan(scan) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      try {
        const res = await this._post(this.getUrl(), { action: 'pushScans', scans: [scan] });
        const doneIds = [
          ...(res.result?.acceptedIds || []),
          ...(res.result?.duplicateIds || [])
        ];
        if (doneIds.length) await global.RaceDB.markSynced({ scanIds: doneIds });
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return { queued: false };
      } catch (err) {
        console.warn('即時上傳失敗，已保留至離線佇列:', err);
        return { queued: true, error: err.message };
      }
    }

    async pushChips(chips) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      try {
        await this._post(this.getUrl(), { action: 'upsertChips', chips });
        await global.RaceDB.markSynced({ chipIds: chips.map((c) => c.idNumber) });
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return { queued: false };
      } catch (err) {
        return { queued: true, error: err.message };
      }
    }

    async pushRegistrations(registrations, auth = null) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      try {
        const body = { action: 'upsertRegistrations', registrations };
        if (auth) body.auth = auth;
        await this._post(this.getUrl(), body);
        await global.RaceDB.markSynced({ regIds: registrations.map((r) => r.regId) });
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return { queued: false };
      } catch (err) {
        return { queued: true, error: err.message };
      }
    }

    async pushPlatform(platform, auth, extra = {}) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      const scoped = this._buildScopedPlatformPayload(platform, auth, extra);
      const res = await this._post(this.getUrl(), {
        action: 'savePlatform',
        auth,
        ...scoped
      });
      await global.RaceDB.markSynced({ platform: true });
      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return res;
    }

    async pushAdminEditScan(scan, auth) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      const res = await this._post(this.getUrl(), { action: 'editScan', auth, scan });
      await global.RaceDB.markSynced({ scanIds: [scan.id] });
      return res;
    }

    async pushHistoryRecords(historyRecords, auth = null) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      try {
        const body = { action: 'upsertHistory', historyRecords };
        if (auth) body.auth = auth;
        await this._post(this.getUrl(), body);
        await global.RaceDB.markSynced({ historyIds: historyRecords.map((h) => h.id) });
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return { queued: false };
      } catch (err) {
        return { queued: true, error: err.message };
      }
    }

    async deleteHistoryRecordsCloud(ids, auth = null) {
      if (!this.isOnlineEnabled() || !navigator.onLine || !ids?.length) return { queued: true };
      try {
        const body = { action: 'deleteHistory', ids };
        if (auth) body.auth = auth;
        await this._post(this.getUrl(), body);
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return { queued: false };
      } catch (err) {
        return { queued: true, error: err.message };
      }
    }
  }

  global.CloudSync = new CloudSync();
})(window);
