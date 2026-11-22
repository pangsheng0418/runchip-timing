/**
 * 智慧路跑 NFC 晶片計時系統 v2.0 — Google Apps Script (GAS) 雲端同步與帳密驗證模組
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
        if (v !== undefined && v !== null && v !== '') {
          target.searchParams.set(k, String(v));
        }
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

    /**
     * 線上帳密驗證並自動快取至本機 IndexedDB（讓後續離線登入無縫運作）
     */
    async loginCloud(mode, username, password) {
      const url = this.getUrl();
      if (!url) throw new Error('尚未設定雲端 API 網址');
      const res = await this._post(url, {
        action: 'login',
        mode,
        username,
        password
      });
      if (res.data) {
        await global.RaceDB.mergeCloudData(res.data);
      }
      if (mode === 'staff' && res.account) {
        await global.RaceDB.cacheStaffAccount(res.account);
      }
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
     * 統一將本機所有待同步的選手與感應紀錄批次上傳至 Google Sheets
     */
    async syncPendingToCloud() {
      const url = this.getUrl();
      if (!url) throw new Error('請先設定 Google Apps Script 網址');

      const pending = await global.RaceDB.getPendingSyncCounts();
      if (pending.total === 0) {
        return { uploadedRunners: 0, uploadedScans: 0, skippedDuplicate: 0 };
      }

      const res = await this._post(url, {
        action: 'batchSync',
        runners: pending.unsyncedRunners,
        scans: pending.unsyncedScans
      });

      const syncedRunnerBibs = pending.unsyncedRunners.map((r) => r.bib);
      if (syncedRunnerBibs.length) {
        await global.RaceDB.markRunnersSynced(syncedRunnerBibs);
      }

      const acceptedIds = [
        ...(res.result?.scans?.acceptedIds || []),
        ...(res.result?.scans?.duplicateIds || [])
      ];
      if (acceptedIds.length) {
        await global.RaceDB.markScansSynced(acceptedIds);
      }

      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return {
        uploadedRunners: syncedRunnerBibs.length,
        uploadedScans: res.result?.scans?.inserted || 0,
        skippedDuplicate: res.result?.scans?.skippedDuplicate || 0
      };
    }

    async pushSingleScan(scan) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      try {
        const res = await this._post(this.getUrl(), {
          action: 'pushScans',
          scans: [scan]
        });
        const doneIds = [
          ...(res.result?.acceptedIds || []),
          ...(res.result?.duplicateIds || [])
        ];
        if (doneIds.length) {
          await global.RaceDB.markScansSynced(doneIds);
        }
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return {
          queued: false,
          serverDuplicate: (res.result?.duplicateIds || []).includes(scan.id)
        };
      } catch (err) {
        console.warn('即時上傳失敗，已保留至離線佇列:', err);
        return { queued: true, error: err.message };
      }
    }

    async pushRunners(runners) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      try {
        await this._post(this.getUrl(), {
          action: 'upsertRunners',
          runners
        });
        await global.RaceDB.markRunnersSynced(runners.map((r) => r.bib));
        global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
        return { queued: false };
      } catch (err) {
        console.warn('選手上傳失敗，已保留於離線佇列:', err);
        return { queued: true, error: err.message };
      }
    }

    async pushConfig(config, auth) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      return this._post(this.getUrl(), {
        action: 'saveConfig',
        auth,
        config: {
          raceName: config.raceName,
          raceDate: config.raceDate,
          debounceSeconds: config.debounceSeconds,
          rankBy: config.rankBy,
          adminUsername: config.adminUsername,
          adminPassword: config.adminPassword
        },
        categories: config.categories,
        checkpoints: config.checkpoints,
        staffAccounts: config.staffAccounts
      });
    }

    async pushAdminEditScan(scan, auth) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      const res = await this._post(this.getUrl(), {
        action: 'editScan',
        auth,
        scan
      });
      await global.RaceDB.markScansSynced([scan.id]);
      return res;
    }

    async pushRestoreAll(backupObj, auth) {
      const url = this.getUrl();
      if (!url) throw new Error('請先設定 Google Apps Script 網址');
      const cfg = backupObj.config || {};
      return this._post(url, {
        action: 'restoreAll',
        auth,
        config: {
          raceName: cfg.raceName,
          raceDate: cfg.raceDate,
          debounceSeconds: cfg.debounceSeconds,
          rankBy: cfg.rankBy,
          adminUsername: cfg.adminUsername,
          adminPassword: cfg.adminPassword
        },
        categories: cfg.categories,
        checkpoints: cfg.checkpoints,
        staffAccounts: cfg.staffAccounts,
        runners: backupObj.runners || [],
        scans: backupObj.scans || []
      });
    }
  }

  global.CloudSync = new CloudSync();
})(window);
