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

    async _get(url, params = {}) {
      const target = new URL(url);
      Object.entries({ ...params, _t: Date.now() }).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') target.searchParams.set(k, String(v));
      });
      const resp = await fetch(target.toString(), { method: 'GET', redirect: 'follow', cache: 'no-store' });
      if (!resp.ok) throw new Error(`HTTP 錯誤 (${resp.status})`);
      const json = await resp.json();
      if (!json.ok) throw new Error(json.error || '雲端 API 回傳錯誤');
      return json;
    }

    async _post(url, body) {
      const resp = await fetch(url, {
        method: 'POST',
        redirect: 'follow',
        cache: 'no-store',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body)
      });
      if (!resp.ok) throw new Error(`HTTP 錯誤 (${resp.status})`);
      const json = await resp.json();
      if (!json.ok) throw new Error(json.error || '雲端 API 寫入失敗');
      return json;
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

    async loginCloud(mode, username, password) {
      const url = this.getUrl();
      if (!url) throw new Error('尚未設定雲端 API 網址');
      const res = await this._post(url, { action: 'login', mode, username, password });
      if (res.data) await global.RaceDB.mergeCloudData(res.data);
      if (mode === 'staff' && res.account) await global.RaceDB.cacheStaffAccount(res.account);
      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return res;
    }

    async pullFromCloud(auth) {
      const url = this.getUrl();
      if (!url) throw new Error('請先設定 Google Apps Script 網址');
      const res = await this._post(url, { action: 'fullSync', auth });
      await global.RaceDB.mergeCloudData(res.data);
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

      if (!includePlatform && pending.chips === 0 && pending.registrations === 0 && pending.scans === 0) {
        return { platformSynced: false, uploadedChips: 0, uploadedRegs: 0, uploadedScans: 0, skippedDuplicate: 0 };
      }

      const body = {
        action: 'batchSync',
        chips: pending.unsyncedChips,
        registrations: pending.unsyncedRegs,
        scans: pending.unsyncedScans
      };

      if (includePlatform) {
        const platform = await global.RaceDB.getPlatform();
        body.auth = session.auth;
        body.platform = {
          globalConfig: {
            adminUsername: platform.adminUsername,
            adminPassword: platform.adminPassword,
            activeEventId: platform.activeEventId
          },
          events: platform.events,
          staffAccounts: platform.staffAccounts
        };
      }

      const res = await this._post(url, body);

      const scanDoneIds = [
        ...(res.result?.scans?.acceptedIds || []),
        ...(res.result?.scans?.duplicateIds || [])
      ];

      await global.RaceDB.markSynced({
        platform: Boolean(res.result?.platformSynced),
        chipIds: pending.unsyncedChips.map((c) => c.idNumber),
        regIds: pending.unsyncedRegs.map((r) => r.regId),
        scanIds: scanDoneIds
      });

      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return {
        platformSynced: Boolean(res.result?.platformSynced),
        uploadedChips: pending.unsyncedChips.length,
        uploadedRegs: pending.unsyncedRegs.length,
        uploadedScans: res.result?.scans?.inserted || 0,
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

    async pushRegistrations(registrations) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      try {
        await this._post(this.getUrl(), { action: 'upsertRegistrations', registrations });
        await global.RaceDB.markSynced({ regIds: registrations.map((r) => r.regId) });
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return { queued: false };
      } catch (err) {
        return { queued: true, error: err.message };
      }
    }

    async pushPlatform(platform, auth) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      const res = await this._post(this.getUrl(), {
        action: 'savePlatform',
        auth,
        globalConfig: {
          adminUsername: platform.adminUsername,
          adminPassword: platform.adminPassword,
          activeEventId: platform.activeEventId
        },
        events: platform.events,
        staffAccounts: platform.staffAccounts
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
  }

  global.CloudSync = new CloudSync();
})(window);
