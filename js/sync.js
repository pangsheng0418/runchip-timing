/**
 * 智慧路跑 NFC 晶片計時平台 v3.0 — 多場次雲端同步與帳密驗證模組
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
      Object.entries(params).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') target.searchParams.set(k, String(v));
      });
      const resp = await fetch(target.toString(), { method: 'GET', redirect: 'follow' });
      if (!resp.ok) throw new Error(`HTTP 錯誤 (${resp.status})`);
      const json = await resp.json();
      if (!json.ok) throw new Error(json.error || '雲端 API 回傳錯誤');
      return json;
    }

    async _post(url, body) {
      const resp = await fetch(url, {
        method: 'POST',
        redirect: 'follow',
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

    async syncPendingToCloud() {
      const url = this.getUrl();
      if (!url) throw new Error('請先設定 Google Apps Script 網址');

      const pending = await global.RaceDB.getPendingSyncCounts();
      if (pending.total === 0) {
        return { uploadedChips: 0, uploadedRegs: 0, uploadedScans: 0, skippedDuplicate: 0 };
      }

      const res = await this._post(url, {
        action: 'batchSync',
        chips: pending.unsyncedChips,
        registrations: pending.unsyncedRegs,
        scans: pending.unsyncedScans
      });

      const scanDoneIds = [
        ...(res.result?.scans?.acceptedIds || []),
        ...(res.result?.scans?.duplicateIds || [])
      ];

      await global.RaceDB.markSynced({
        chipIds: pending.unsyncedChips.map((c) => c.idNumber),
        regIds: pending.unsyncedRegs.map((r) => r.regId),
        scanIds: scanDoneIds
      });

      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return {
        uploadedChips: pending.unsyncedChips.length,
        uploadedRegs: pending.unsyncedRegs.length,
        uploadedScans: res.result?.scans?.inserted || 0,
        skippedDuplicate: res.result?.scans?.skippedDuplicate || 0
      };
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
      return this._post(this.getUrl(), {
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
