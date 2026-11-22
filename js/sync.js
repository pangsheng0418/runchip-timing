/**
 * 智慧路跑 NFC 晶片計時系統 — Google Apps Script (GAS) 雲端同步模組
 * 說明：POST 一律使用 text/plain;charset=utf-8，避免 GAS Web App 觸發 CORS preflight 失敗
 */
(function (global) {
  'use strict';

  class CloudSync {
    getUrl() {
      return (global.RaceDB.getDeviceSettings().gasUrl || '').trim();
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

    /** 測試 GAS 連線 */
    async ping(customUrl) {
      const url = (customUrl || this.getUrl()).trim();
      if (!url) throw new Error('請先輸入 Google Apps Script 網頁應用程式 URL');
      return this._get(url, { action: 'ping' });
    }

    /** 公開查詢端拉取（已隱碼身份證字號） */
    async fetchPublicData(customUrl) {
      const url = (customUrl || this.getUrl()).trim();
      if (!url) throw new Error('未設定雲端 API 網址');
      return this._get(url, { action: 'publicData' });
    }

    /** 工作人員端從雲端拉取最新資料並合併至本地 IndexedDB */
    async pullFromCloud(adminPin = '') {
      const url = this.getUrl();
      if (!url) throw new Error('請先於後台設定 Google Apps Script 網址');
      const res = await this._get(url, { action: 'fullSync', pin: adminPin });
      await global.RaceDB.mergeCloudData(res.data);
      global.RaceDB.saveDeviceSettings({ lastSyncAt: Date.now() });
      return res;
    }

    /** 一鍵將本地所有待同步的選手與感應紀錄批次上傳至 Google Sheets */
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

    /** 即時推送單筆感應（連線模式下自動觸發，失敗時保留於離線佇列） */
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

    /** 推送選手資料更新 */
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

    /** 同步賽事設定、組別起跑時間與站點至 Google Sheets */
    async pushConfig(config, adminPin) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      return this._post(this.getUrl(), {
        action: 'saveConfig',
        adminPin,
        config: {
          raceName: config.raceName,
          raceDate: config.raceDate,
          debounceSeconds: config.debounceSeconds,
          rankBy: config.rankBy,
          newAdminPin: config.adminPin
        },
        categories: config.categories,
        checkpoints: config.checkpoints
      });
    }

    /** 管理員人工修正感應時間同步至雲端 */
    async pushAdminEditScan(scan, adminPin) {
      if (!this.isOnlineEnabled() || !navigator.onLine) return { queued: true };
      const res = await this._post(this.getUrl(), {
        action: 'editScan',
        adminPin,
        scan
      });
      await global.RaceDB.markScansSynced([scan.id]);
      return res;
    }
  }

  global.CloudSync = new CloudSync();
})(window);
