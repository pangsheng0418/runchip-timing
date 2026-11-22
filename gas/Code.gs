/**
 * ============================================================================
 * TUR (Taiwan Ultra Runners) 路跑晶片計時平台 v3.10 — Google Apps Script (GAS) 後端
 * ============================================================================
 * 核心設計：
 * - 平台支援建立多場次路跑活動 (Events)，每場擁有獨立分組、鳴槍時間、站點與參賽名單
 * - 共用晶片主檔 (Chips)：晶片內僅存「姓名｜身份證字號」，同一晶片可重複參加不同場次與組別
 * - 參賽明細 (Registrations)：支援選填 DUV 欄位（性別、出生年月日、城市、國家、姓／名）
 * - 歷年活動成績 (HistoryResults)：支援匯入歷年成績並整合八百壯士英雄榜
 * - 三大權限分區：公開查詢（隱碼免登入）、作業員帳密鎖定站點、多管理者獨立隔離
 *
 * 部署步驟：
 * 1. 開啟 Google 試算表 →「擴充功能」→「Apps Script」
 * 2. 貼入本檔案全部內容並儲存
 * 3. 選擇函式「setupSheets」並點選「執行」（自動建立 8 張工作表）
 * 4. 點選「部署」→「新增部署作業」（或「管理部署作業」→「建立新版本」）
 *    - 類型：網頁應用程式 ｜ 執行身分：我 ｜ 誰可以存取：所有人 (Anyone)
 * ============================================================================
 */

const SHEET_NAMES = {
  GLOBAL_CONFIG: 'GlobalConfig',
  ADMINS: 'AdminAccounts',
  EVENTS: 'Events',
  STAFF: 'StaffAccounts',
  CHIPS: 'Chips',
  REGISTRATIONS: 'Registrations',
  SCANS: 'Scans',
  HISTORY: 'HistoryResults'
};

/**
 * 首次初始化：自動建立 8 張工作表與預設範例活動/帳號
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const now = Date.now();

  // 1. GlobalConfig（平台主管理員帳密與預設活動）
  const cfgSheet = getOrCreateSheet_(ss, SHEET_NAMES.GLOBAL_CONFIG, ['key', 'value', 'description']);
  if (cfgSheet.getLastRow() <= 1) {
    cfgSheet.getRange(2, 1, 3, 3).setValues([
      ['adminUsername', 'admin', '後台最高管理員帳號'],
      ['adminPassword', 'admin1234', '後台最高管理員密碼'],
      ['activeEventId', 'EVT-2026-01', '預設顯示活動代碼']
    ]);
  }

  // 1.5 AdminAccounts（多管理者帳號列表，各管理者活動資料獨立隔離）
  const adminSheet = getOrCreateSheet_(ss, SHEET_NAMES.ADMINS, [
    'username', 'password', 'displayName', 'isSuperAdmin', 'enabled', 'updatedAt'
  ]);
  if (adminSheet.getLastRow() <= 1) {
    const gCfg = readGlobalConfigMap_();
    adminSheet.getRange(2, 1, 1, 6).setValues([[
      String(gCfg.adminUsername || 'admin').trim(),
      String(gCfg.adminPassword || 'admin1234'),
      '最高系統管理員',
      1,
      1,
      now
    ]]);
  }

  // 2. Events（多場次路跑活動清單：支援一般距離賽 distance 與繞圈計時賽 lap_timed，含建立者 ownerAdmin 與 DUV 設定）
  const evtSheet = getOrCreateSheet_(ss, SHEET_NAMES.EVENTS, [
    'eventId', 'name', 'date', 'location', 'status', 'debounceSeconds',
    'rankBy', 'categoriesJson', 'checkpointsJson', 'updatedAt', 'eventType', 'lapDistanceKm', 'ownerAdmin',
    'englishName', 'organization'
  ]);
  if (evtSheet.getLastRow() <= 1) {
    const defaultCats = JSON.stringify([
      { id: '42K', name: '全馬組 42K', distanceKm: 42.195, gunTimeMs: null, gunTimeIso: '', color: '#ef4444' },
      { id: '21K', name: '半馬組 21K', distanceKm: 21.0975, gunTimeMs: null, gunTimeIso: '', color: '#3b82f6' },
      { id: '10K', name: '健康組 10K', distanceKm: 10, gunTimeMs: null, gunTimeIso: '', color: '#10b981' }
    ]);
    const defaultCps = JSON.stringify([
      { id: 'CP1', name: '補給站 1 (5K)', order: 1, distanceKm: 5, type: 'checkpoint' },
      { id: 'CP2', name: '折返點 / 補給站 2 (10K)', order: 2, distanceKm: 10, type: 'checkpoint' },
      { id: 'CP3', name: '補給站 3 (21K)', order: 3, distanceKm: 21.1, type: 'checkpoint' },
      { id: 'FINISH', name: '終點衝線門', order: 4, distanceKm: 42.195, type: 'finish' }
    ]);
    evtSheet.getRange(2, 3, 1, 1).setNumberFormat('@');
    evtSheet.getRange(2, 1, 1, 15).setValues([[
      'EVT-2026-01', '2026 TUR 超級馬拉松挑戰賽', today, '台北市政府廣場', 'active',
      30, 'gun', defaultCats, defaultCps, now, 'distance', 0, 'admin',
      '', 'Taiwan Ultra Runners'
    ]]);
  }

  // 3. StaffAccounts（作業人員帳密、權限與鎖定活動/站點，含所屬管理者 ownerAdmin）
  const staffSheet = getOrCreateSheet_(ss, SHEET_NAMES.STAFF, [
    'username', 'password', 'displayName', 'role', 'eventId', 'checkpointId', 'enabled', 'updatedAt', 'ownerAdmin'
  ]);
  if (staffSheet.getLastRow() <= 1) {
    staffSheet.getRange(2, 1, 7, 9).setValues([
      ['writer01', '1111', '晶片寫入發卡組', 'writer', 'ALL', '', 1, now, ''],
      ['checkin01', '1111', '選手報到感應組', 'scanner', 'ALL', 'CHECKIN', 1, now, ''],
      ['lap01', '1111', '繞圈計次感應組', 'scanner', 'ALL', 'LAP', 1, now, ''],
      ['cp1', '1111', '補給站 1 感應組', 'scanner', 'ALL', 'CP1', 1, now, ''],
      ['cp2', '1111', '補給站 2 / 折返點組', 'scanner', 'ALL', 'CP2', 1, now, ''],
      ['cp3', '1111', '補給站 3 感應組', 'scanner', 'ALL', 'CP3', 1, now, ''],
      ['finish01', '1111', '終點衝線感應組', 'scanner', 'ALL', 'FINISH', 1, now, '']
    ]);
  }

  // 4. Chips（跨活動共用晶片主檔：僅存姓名、身份證字號與晶片硬體 UID）
  getOrCreateSheet_(ss, SHEET_NAMES.CHIPS, [
    'idNumber', 'name', 'nfcUid', 'writtenAt', 'updatedAt'
  ]);

  // 5. Registrations（各活動參賽選手明細、組別與選填 DUV 個人欄位）
  getOrCreateSheet_(ss, SHEET_NAMES.REGISTRATIONS, [
    'regId', 'eventId', 'bib', 'name', 'idNumber', 'categoryId', 'penaltySec', 'updatedAt',
    'gender', 'dob', 'city', 'country', 'lastName', 'firstName'
  ]);

  // 6. Scans（各活動晶片感應紀錄）
  getOrCreateSheet_(ss, SHEET_NAMES.SCANS, [
    'id', 'eventId', 'bib', 'idNumber', 'checkpointId', 'timestampMs', 'timeStr', 'isoTime',
    'source', 'deviceId', 'operator', 'deleted', 'note', 'updatedAt'
  ]);

  // 7. HistoryResults（以前歷年活動與個人成績匯入）
  getOrCreateSheet_(ss, SHEET_NAMES.HISTORY, [
    'id', 'year', 'eventDate', 'eventName', 'categoryName', 'bib', 'name', 'idNumber',
    'finishTime', 'distanceKm', 'laps', 'overallRank', 'categoryRank', 'status', 'note', 'ownerAdmin', 'updatedAt'
  ]);

  Logger.log('TUR (Taiwan Ultra Runners) 計時平台 v3.10 共 8 張工作表已初始化完成！');
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * HTTP GET 入口（公開查詢與連線測試，密碼驗證一律走 POST）
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || 'publicData';

    if (action === 'ping') {
      return jsonResponse_({
        ok: true,
        apiVersion: '3.10',
        multiAdmin: true,
        serverTimeMs: Date.now(),
        message: 'TUR (Taiwan Ultra Runners) 路跑計時平台 API v3.10 運作正常'
      });
    }

    if (action === 'publicData') {
      const data = readAllData_(true, null);
      return jsonResponse_({
        ok: true,
        apiVersion: '3.9',
        multiAdmin: true,
        serverTimeMs: Date.now(),
        data: data
      });
    }

    return jsonResponse_({ ok: false, error: '未知的 GET action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) });
  }
}

/**
 * HTTP POST 入口（使用 text/plain;charset=utf-8 避免 CORS 預檢）
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(rawBody);
    const action = payload.action;
    if (!action) return jsonResponse_({ ok: false, error: '缺少 action 參數' });

    const gCfg = readGlobalConfigMap_();

    // 1. 登入驗證（作業人員或各管理者帳號）
    if (action === 'login') {
      const mode = payload.mode || 'staff';
      const u = String(payload.username || '').trim();
      const p = String(payload.password || '');

      if (mode === 'admin') {
        const matchedAdmin = verifyAdminAccount_({ username: u, password: p }, gCfg);
        if (!matchedAdmin) {
          return jsonResponse_({ ok: false, multiAdmin: true, error: '管理員帳號或密碼錯誤，或帳號已停用' });
        }
        return jsonResponse_({
          ok: true,
          apiVersion: '3.5',
          multiAdmin: true,
          serverTimeMs: Date.now(),
          account: {
            username: matchedAdmin.username,
            displayName: matchedAdmin.displayName || matchedAdmin.username,
            isSuperAdmin: matchedAdmin.isSuperAdmin,
            role: 'admin'
          },
          data: readAllData_(false, matchedAdmin)
        });
      }

      const staffList = readStaffList_();
      const matched = staffList.find(function (s) {
        return s.enabled && s.username === u && s.password === p;
      });
      if (!matched) {
        return jsonResponse_({ ok: false, multiAdmin: true, error: '作業員帳號或密碼錯誤，或帳號已停用' });
      }
      return jsonResponse_({
        ok: true,
        apiVersion: '3.5',
        multiAdmin: true,
        serverTimeMs: Date.now(),
        account: matched,
        data: readAllData_(false, null)
      });
    }

    // 2. 授權同步拉取（管理員僅回傳該管理員本人建立的活動，避免看見或誤改他人活動）
    if (action === 'fullSync') {
      const callerAdmin = verifyAdminAccount_(payload.auth, gCfg);
      const isStaff = !callerAdmin && verifyStaff_(payload.auth);
      if (!callerAdmin && !isStaff) {
        return jsonResponse_({ ok: false, multiAdmin: true, error: '身分驗證失敗，請重新登入' });
      }
      return jsonResponse_({
        ok: true,
        apiVersion: '3.5',
        multiAdmin: true,
        serverTimeMs: Date.now(),
        isAdmin: Boolean(callerAdmin),
        isSuperAdmin: Boolean(callerAdmin && callerAdmin.isSuperAdmin),
        adminUsername: callerAdmin ? callerAdmin.username : '',
        data: readAllData_(false, callerAdmin)
      });
    }

    // 3. 作業員與管理員均可上傳：感應紀錄、共用晶片主檔、批次離線同步
    if (action === 'pushScans') {
      const res = handlePushScans_(payload.scans || []);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: res });
    }

    if (action === 'upsertChips') {
      const res = handleUpsertChips_(payload.chips || []);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: res });
    }

    if (action === 'upsertRegistrations') {
      const callerAdmin = payload.auth ? verifyAdminAccount_(payload.auth, gCfg) : null;
      if (payload.auth && !callerAdmin) {
        return jsonResponse_({ ok: false, multiAdmin: true, error: '管理員身分驗證失敗' });
      }
      const regs = payload.registrations || [];
      if (callerAdmin && regs.length) {
        assertEventsOwnedByAdmin_(regs.map(function (r) { return r.eventId; }), callerAdmin, gCfg);
      }
      const res = handleUpsertRegistrations_(regs);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: res });
    }

    if (action === 'batchSync') {
      let platformSynced = false;
      let purgeRes = { eventIds: [], deletedRegistrations: 0, deletedScans: 0 };
      const callerAdmin = payload.auth ? verifyAdminAccount_(payload.auth, gCfg) : null;
      if (payload.platform && callerAdmin) {
        const saveRes = handleSavePlatform_(payload.platform, callerAdmin, gCfg);
        platformSynced = true;
        if (saveRes && saveRes.purge) purgeRes = saveRes.purge;
      }
      const chipRes = handleUpsertChips_(payload.chips || []);
      const regs = payload.registrations || [];
      if (callerAdmin && regs.length) {
        assertEventsOwnedByAdmin_(regs.map(function (r) { return r.eventId; }), callerAdmin, gCfg);
      }
      const regRes = handleUpsertRegistrations_(regs);
      const scanRes = handlePushScans_(payload.scans || []);
      const histRes = (callerAdmin && Array.isArray(payload.historyRecords) && payload.historyRecords.length)
        ? handleUpsertHistory_(payload.historyRecords, callerAdmin)
        : { inserted: 0, updated: 0 };
      return jsonResponse_({
        ok: true,
        multiAdmin: true,
        serverTimeMs: Date.now(),
        result: { platformSynced: platformSynced, purge: purgeRes, chips: chipRes, registrations: regRes, scans: scanRes, history: histRes }
      });
    }

    // 以下需管理員帳號驗證
    const callerAdmin = verifyAdminAccount_(payload.auth, gCfg);
    if (!callerAdmin) {
      return jsonResponse_({ ok: false, multiAdmin: true, error: '權限不足：需要管理員帳號密碼驗證' });
    }

    if (action === 'upsertHistory') {
      const res = handleUpsertHistory_(payload.historyRecords || [], callerAdmin);
      return jsonResponse_({ ok: true, multiAdmin: true, serverTimeMs: Date.now(), result: res });
    }

    if (action === 'deleteHistory') {
      const res = handleDeleteHistory_(payload.ids || []);
      return jsonResponse_({ ok: true, multiAdmin: true, serverTimeMs: Date.now(), result: res });
    }

    if (action === 'savePlatform') {
      const saveRes = handleSavePlatform_(payload, callerAdmin, gCfg);
      return jsonResponse_({
        ok: true,
        multiAdmin: true,
        serverTimeMs: Date.now(),
        message: '您的活動設定與帳號已同步至雲端',
        purge: saveRes && saveRes.purge ? saveRes.purge : { eventIds: [], deletedRegistrations: 0, deletedScans: 0 }
      });
    }

    if (action === 'editScan') {
      const res = handleEditScan_(payload.scan, callerAdmin, gCfg);
      return jsonResponse_({ ok: true, multiAdmin: true, serverTimeMs: Date.now(), result: res });
    }

    return jsonResponse_({ ok: false, error: '不支援的 POST action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, multiAdmin: true, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function readAdminAccountsList_(gCfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ADMINS);
  const cfg = gCfg || readGlobalConfigMap_();
  const primaryUser = String(cfg.adminUsername || 'admin').trim() || 'admin';
  const primaryPass = String(cfg.adminPassword || 'admin1234');

  const list = readTableObjects_(sheet).map(function (a) {
    return {
      username: String(a.username || '').trim(),
      password: String(a.password || ''),
      displayName: String(a.displayName || a.username || '').trim(),
      isSuperAdmin: String(a.isSuperAdmin) === '1' || String(a.isSuperAdmin).toLowerCase() === 'true',
      enabled: String(a.enabled) !== '0' && String(a.enabled).toLowerCase() !== 'false',
      updatedAt: a.updatedAt ? Number(a.updatedAt) : Date.now()
    };
  }).filter(function (a) { return a.username; });

  if (!list.length) {
    return [{
      username: primaryUser,
      password: primaryPass,
      displayName: '最高系統管理員',
      isSuperAdmin: true,
      enabled: true,
      updatedAt: Date.now()
    }];
  }

  if (!list.some(function (a) { return a.isSuperAdmin && a.enabled; })) {
    const idx = list.findIndex(function (a) { return a.username === primaryUser; });
    if (idx >= 0) {
      list[idx].isSuperAdmin = true;
      list[idx].enabled = true;
    } else {
      list.unshift({
        username: primaryUser,
        password: primaryPass,
        displayName: '最高系統管理員',
        isSuperAdmin: true,
        enabled: true,
        updatedAt: Date.now()
      });
    }
  }
  return list;
}

function verifyAdminAccount_(auth, gCfg) {
  if (!auth || !auth.username) return null;
  const u = String(auth.username || '').trim();
  const p = String(auth.password || '');
  const list = readAdminAccountsList_(gCfg);
  return list.find(function (a) {
    return a.enabled && a.username === u && a.password === p;
  }) || null;
}

function assertEventsOwnedByAdmin_(eventIds, callerAdmin, gCfg) {
  const defaultOwner = String((gCfg && gCfg.adminUsername) || 'admin').trim() || 'admin';
  const events = readEventsList_(defaultOwner);
  const ownerMap = {};
  events.forEach(function (ev) { ownerMap[ev.eventId] = ev.ownerAdmin; });

  eventIds.forEach(function (eid) {
    const clean = String(eid || '').trim();
    if (!clean) return;
    if (ownerMap[clean] && ownerMap[clean] !== callerAdmin.username) {
      throw new Error('權限不足：活動「' + clean + '」屬於其他管理員，不可修改');
    }
  });
}

function verifyStaff_(auth) {
  if (!auth || !auth.username) return false;
  return readStaffList_().some(function (s) {
    return s.enabled && s.username === String(auth.username).trim() && s.password === String(auth.password || '');
  });
}

function readGlobalConfigMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GLOBAL_CONFIG);
  const map = {};
  if (!sheet || sheet.getLastRow() <= 1) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0]) map[String(r[0]).trim()] = r[1];
  });
  return map;
}

function readStaffList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.STAFF);
  return readTableObjects_(sheet).map(function (s) {
    return {
      username: String(s.username || '').trim(),
      password: String(s.password || ''),
      displayName: String(s.displayName || '').trim(),
      role: String(s.role || 'scanner').trim(),
      eventId: String(s.eventId || 'ALL').trim() || 'ALL',
      checkpointId: String(s.checkpointId || '').trim(),
      ownerAdmin: String(s.ownerAdmin || '').trim(),
      enabled: String(s.enabled) !== '0' && String(s.enabled).toLowerCase() !== 'false',
      updatedAt: s.updatedAt ? Number(s.updatedAt) : Date.now()
    };
  }).filter(function (s) { return s.username; });
}

function formatDateCell_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

function formatTimeCell_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    const tz = Session.getScriptTimeZone();
    // Google Sheets 將 HH:mm:ss 存入時會以 1899-12-30 為基準日期，超過 24 小時則會跨日
    if (val.getFullYear() <= 1900) {
      const epoch = new Date(1899, 11, 30, 0, 0, 0, 0);
      const diffSec = Math.max(0, Math.round((val.getTime() - epoch.getTime()) / 1000));
      const h = Math.floor(diffSec / 3600);
      const m = Math.floor((diffSec % 3600) / 60);
      const s = diffSec % 60;
      const pad = function (n) { return (n < 10 ? '0' : '') + n; };
      return pad(h) + ':' + pad(m) + ':' + pad(s);
    }
    return Utilities.formatDate(val, tz, 'HH:mm:ss');
  }
  return String(val).trim();
}

function readEventsList_(defaultOwner) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.EVENTS);
  const fallbackOwner = String(defaultOwner || 'admin').trim() || 'admin';
  return readTableObjects_(sheet).map(function (ev) {
    let cats = [];
    let cps = [];
    try { cats = JSON.parse(ev.categoriesJson || '[]'); } catch (e) {}
    try { cps = JSON.parse(ev.checkpointsJson || '[]'); } catch (e) {}
    const eventType = String(ev.eventType || 'distance').trim() || 'distance';
    return {
      eventId: String(ev.eventId || '').trim(),
      name: String(ev.name || '').trim(),
      englishName: String(ev.englishName || '').trim(),
      organization: String(ev.organization || '').trim(),
      date: formatDateCell_(ev.date),
      location: String(ev.location || '').trim(),
      status: String(ev.status || 'active').trim(),
      ownerAdmin: String(ev.ownerAdmin || '').trim() || fallbackOwner,
      eventType: eventType,
      lapDistanceKm: Number(ev.lapDistanceKm || 0),
      debounceSeconds: Number(ev.debounceSeconds || 30),
      rankBy: String(ev.rankBy || 'gun').trim(),
      categories: Array.isArray(cats) ? cats : [],
      checkpoints: Array.isArray(cps) ? cps : [],
      updatedAt: ev.updatedAt ? Number(ev.updatedAt) : Date.now()
    };
  }).filter(function (ev) { return ev.eventId; });
}

function readAllData_(maskSensitive, callerAdmin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gCfg = readGlobalConfigMap_();
  const defaultOwner = String(gCfg.adminUsername || 'admin').trim() || 'admin';
  const allEvents = readEventsList_(defaultOwner);

  // 若為管理員登入，僅回傳該管理員本人建立的活動、報名、感應與作業員，落實多管理者資料隔離！
  const events = callerAdmin
    ? allEvents.filter(function (ev) { return ev.ownerAdmin === callerAdmin.username; })
    : allEvents;

  const allowedEventIdSet = callerAdmin ? {} : null;
  if (callerAdmin) {
    events.forEach(function (ev) { allowedEventIdSet[ev.eventId] = true; });
  }

  const globalConfig = {
    activeEventId: String(gCfg.activeEventId || (allEvents[0] ? allEvents[0].eventId : '')).trim()
  };
  if (callerAdmin && callerAdmin.isSuperAdmin) {
    globalConfig.adminUsername = defaultOwner;
    globalConfig.adminPassword = String(gCfg.adminPassword || 'admin1234');
  }

  const registrations = readTableObjects_(ss.getSheetByName(SHEET_NAMES.REGISTRATIONS)).map(function (r) {
    const rawId = String(r.idNumber || '').trim().toUpperCase();
    const obj = {
      regId: String(r.regId || (String(r.eventId).trim() + '__' + (rawId || String(r.bib).trim()))),
      eventId: String(r.eventId || '').trim(),
      bib: String(r.bib || '').trim(),
      name: String(r.name || '').trim(),
      lastName: String(r.lastName || '').trim(),
      firstName: String(r.firstName || '').trim(),
      gender: String(r.gender || '').trim().toUpperCase(),
      city: String(r.city || '').trim(),
      country: String(r.country || '').trim(),
      idMasked: maskIdNumber_(rawId),
      idHash: rawId ? sha256Hex_(rawId) : '',
      categoryId: String(r.categoryId || '').trim(),
      penaltySec: Number(r.penaltySec || 0),
      updatedAt: r.updatedAt ? Number(r.updatedAt) : Date.now()
    };
    if (!maskSensitive) {
      obj.idNumber = rawId;
      obj.dob = formatDateCell_(r.dob);
    }
    return obj;
  }).filter(function (r) {
    if (!r.eventId || !r.bib) return false;
    if (allowedEventIdSet && !allowedEventIdSet[r.eventId]) return false;
    return true;
  });

  const scans = readTableObjects_(ss.getSheetByName(SHEET_NAMES.SCANS)).map(function (s) {
    const obj = {
      id: String(s.id || ''),
      eventId: String(s.eventId || '').trim(),
      bib: String(s.bib || '').trim(),
      checkpointId: String(s.checkpointId || '').trim(),
      timestampMs: Number(s.timestampMs || 0),
      timeStr: String(s.timeStr || ''),
      isoTime: String(s.isoTime || ''),
      source: String(s.source || 'nfc'),
      deviceId: String(s.deviceId || ''),
      operator: String(s.operator || ''),
      deleted: String(s.deleted) === 'true' || Number(s.deleted) === 1,
      note: String(s.note || ''),
      updatedAt: s.updatedAt ? Number(s.updatedAt) : Number(s.timestampMs || 0)
    };
    if (!maskSensitive) obj.idNumber = String(s.idNumber || '').trim().toUpperCase();
    return obj;
  }).filter(function (s) {
    if (!s.id || !s.eventId || !s.bib || s.deleted) return false;
    if (allowedEventIdSet && !allowedEventIdSet[s.eventId]) return false;
    return true;
  });

  const historyRecords = readTableObjects_(ss.getSheetByName(SHEET_NAMES.HISTORY)).map(function (h) {
    const rawId = String(h.idNumber || '').trim().toUpperCase();
    const evDate = formatDateCell_(h.eventDate);
    const yr = String(h.year || (evDate ? evDate.slice(0, 4) : '')).trim();
    const obj = {
      id: String(h.id || '').trim(),
      year: yr,
      eventDate: evDate,
      eventName: String(h.eventName || '').trim(),
      categoryName: String(h.categoryName || '').trim(),
      bib: String(h.bib || '').trim(),
      name: String(h.name || '').trim(),
      idMasked: maskIdNumber_(rawId),
      idHash: rawId ? sha256Hex_(rawId) : '',
      finishTime: formatTimeCell_(h.finishTime),
      distanceKm: Number(h.distanceKm || 0),
      laps: parseInt(h.laps || 0, 10) || 0,
      overallRank: h.overallRank !== '' && h.overallRank !== null && h.overallRank !== undefined ? (parseInt(h.overallRank, 10) || null) : null,
      categoryRank: h.categoryRank !== '' && h.categoryRank !== null && h.categoryRank !== undefined ? (parseInt(h.categoryRank, 10) || null) : null,
      status: String(h.status || '完賽').trim(),
      note: String(h.note || '').trim(),
      ownerAdmin: String(h.ownerAdmin || defaultOwner).trim(),
      updatedAt: h.updatedAt ? Number(h.updatedAt) : Date.now()
    };
    if (!maskSensitive) obj.idNumber = rawId;
    return obj;
  }).filter(function (h) { return h.id && h.eventName && h.name; });

  const result = {
    globalConfig: globalConfig,
    events: events,
    registrations: registrations,
    scans: scans,
    historyRecords: historyRecords
  };

  if (!maskSensitive) {
    result.chips = readTableObjects_(ss.getSheetByName(SHEET_NAMES.CHIPS)).map(function (c) {
      return {
        idNumber: String(c.idNumber || '').trim().toUpperCase(),
        name: String(c.name || '').trim(),
        nfcUid: String(c.nfcUid || '').trim().toUpperCase(),
        writtenAt: c.writtenAt ? Number(c.writtenAt) : null,
        updatedAt: c.updatedAt ? Number(c.updatedAt) : Date.now()
      };
    }).filter(function (c) { return c.idNumber; });
  }

  if (callerAdmin) {
    const allStaff = readStaffList_();
    result.staffAccounts = allStaff.filter(function (s) {
      if (callerAdmin.isSuperAdmin) return !s.ownerAdmin || s.ownerAdmin === callerAdmin.username;
      return s.ownerAdmin === callerAdmin.username;
    });
    const allAdmins = readAdminAccountsList_(gCfg);
    result.adminAccounts = callerAdmin.isSuperAdmin
      ? allAdmins
      : allAdmins.filter(function (a) { return a.username === callerAdmin.username; });
  }

  return result;
}

function handleUpsertChips_(incomingChips) {
  if (!incomingChips || !incomingChips.length) return { inserted: 0, updated: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.CHIPS, ['idNumber', 'name', 'nfcUid', 'writtenAt', 'updatedAt']);

  const values = sheet.getDataRange().getValues();
  const rowById = {};
  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][0] || '').trim().toUpperCase();
    if (id) rowById[id] = i + 1;
  }

  let inserted = 0, updated = 0;
  const toAppend = [];
  const pendingIdxById = {};

  incomingChips.forEach(function (c) {
    const id = String(c.idNumber || '').trim().toUpperCase();
    if (!id) return;
    const row = [
      id,
      String(c.name || '').trim(),
      String(c.nfcUid || '').trim().toUpperCase(),
      c.writtenAt ? Number(c.writtenAt) : Date.now(),
      Number(c.updatedAt || Date.now())
    ];
    if (rowById[id] > 0) {
      sheet.getRange(rowById[id], 1, 1, 5).setValues([row]);
      updated++;
    } else if (pendingIdxById[id] !== undefined) {
      toAppend[pendingIdxById[id]] = row;
      updated++;
    } else {
      pendingIdxById[id] = toAppend.length;
      toAppend.push(row);
      inserted++;
    }
  });

  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 5).setValues(toAppend);
  }
  return { inserted: inserted, updated: updated };
}

function handleUpsertRegistrations_(incomingRegs) {
  if (!incomingRegs || !incomingRegs.length) return { inserted: 0, updated: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regHeaders = [
    'regId', 'eventId', 'bib', 'name', 'idNumber', 'categoryId', 'penaltySec', 'updatedAt',
    'gender', 'dob', 'city', 'country', 'lastName', 'firstName'
  ];
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.REGISTRATIONS, regHeaders);
  if (sheet.getLastColumn() < regHeaders.length) {
    sheet.getRange(1, 1, 1, regHeaders.length).setValues([regHeaders])
      .setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }

  const values = sheet.getDataRange().getValues();
  const rowByKey = {};
  for (let i = 1; i < values.length; i++) {
    const rk = String(values[i][0] || '').trim();
    if (rk) rowByKey[rk] = i + 1;
  }

  let inserted = 0, updated = 0;
  const toAppend = [];
  const pendingIdxByKey = {};

  incomingRegs.forEach(function (r) {
    const eventId = String(r.eventId || '').trim();
    const bib = String(r.bib || '').trim();
    const idNumber = String(r.idNumber || '').trim().toUpperCase();
    if (!eventId || !bib) return;
    const regId = String(r.regId || (eventId + '__' + (idNumber || bib))).trim();
    const genderRaw = String(r.gender || '').trim().toUpperCase();
    const gender = (genderRaw === 'M' || genderRaw === 'MALE' || genderRaw === '男') ? 'M'
      : ((genderRaw === 'F' || genderRaw === 'FEMALE' || genderRaw === '女') ? 'F' : '');

    const row = [
      regId,
      eventId,
      bib,
      String(r.name || '').trim(),
      idNumber,
      String(r.categoryId || '').trim(),
      Number(r.penaltySec || 0),
      Number(r.updatedAt || Date.now()),
      gender,
      formatDateCell_(r.dob),
      String(r.city || '').trim(),
      String(r.country || '').trim(),
      String(r.lastName || '').trim(),
      String(r.firstName || '').trim()
    ];

    if (rowByKey[regId] > 0) {
      sheet.getRange(rowByKey[regId], 3, 1, 1).setNumberFormat('@');
      sheet.getRange(rowByKey[regId], 10, 1, 1).setNumberFormat('@');
      sheet.getRange(rowByKey[regId], 1, 1, regHeaders.length).setValues([row]);
      updated++;
    } else if (pendingIdxByKey[regId] !== undefined) {
      toAppend[pendingIdxByKey[regId]] = row;
      updated++;
    } else {
      pendingIdxByKey[regId] = toAppend.length;
      toAppend.push(row);
      inserted++;
    }
  });

  if (toAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 3, toAppend.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 10, toAppend.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, toAppend.length, regHeaders.length).setValues(toAppend);
  }
  return { inserted: inserted, updated: updated };
}

function handlePushScans_(incomingScans) {
  if (!incomingScans || !incomingScans.length) {
    return { inserted: 0, skippedDuplicate: 0, acceptedIds: [], duplicateIds: [] };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.SCANS, [
    'id', 'eventId', 'bib', 'idNumber', 'checkpointId', 'timestampMs', 'timeStr', 'isoTime',
    'source', 'deviceId', 'operator', 'deleted', 'note', 'updatedAt'
  ]);

  const eventDebounceMap = {};
  readEventsList_().forEach(function (ev) {
    eventDebounceMap[ev.eventId] = Number(ev.debounceSeconds || 30) * 1000;
  });

  const existing = readTableObjects_(sheet);
  const existingIdSet = {};
  const recentMap = {};

  existing.forEach(function (row) {
    const id = String(row.id || '');
    if (id) existingIdSet[id] = true;
    const isDeleted = String(row.deleted) === 'true' || Number(row.deleted) === 1;
    if (!isDeleted && row.eventId && row.bib && row.checkpointId && row.timestampMs) {
      const key = String(row.eventId).trim() + '::' + String(row.bib).trim() + '::' + String(row.checkpointId).trim();
      if (!recentMap[key]) recentMap[key] = [];
      recentMap[key].push(Number(row.timestampMs));
    }
  });

  const toAppend = [];
  const acceptedIds = [];
  const duplicateIds = [];

  const sorted = incomingScans.slice().sort(function (a, b) {
    return Number(a.timestampMs || 0) - Number(b.timestampMs || 0);
  });

  sorted.forEach(function (scan) {
    const id = String(scan.id || Utilities.getUuid());
    const eventId = String(scan.eventId || '').trim();
    const bib = String(scan.bib || '').trim();
    const cpId = String(scan.checkpointId || '').trim();
    const ts = Number(scan.timestampMs || Date.now());
    const isManual = scan.source === 'manual' || scan.source === 'admin';

    if (!eventId || !bib || !cpId) return;

    if (existingIdSet[id]) {
      acceptedIds.push(id);
      return;
    }

    const debounceMs = eventDebounceMap[eventId] !== undefined ? eventDebounceMap[eventId] : 30000;
    const key = eventId + '::' + bib + '::' + cpId;
    if (!isManual && debounceMs > 0 && recentMap[key]) {
      const tooClose = recentMap[key].some(function (prevTs) {
        return Math.abs(ts - prevTs) < debounceMs;
      });
      if (tooClose) {
        duplicateIds.push(id);
        return;
      }
    }

    existingIdSet[id] = true;
    if (!recentMap[key]) recentMap[key] = [];
    recentMap[key].push(ts);

    toAppend.push([
      id,
      eventId,
      bib,
      String(scan.idNumber || '').trim().toUpperCase(),
      cpId,
      ts,
      String(scan.timeStr || ''),
      String(scan.isoTime || new Date(ts).toISOString()),
      String(scan.source || 'nfc'),
      String(scan.deviceId || ''),
      String(scan.operator || ''),
      scan.deleted ? 1 : 0,
      String(scan.note || ''),
      Date.now()
    ]);
    acceptedIds.push(id);
  });

  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 14).setValues(toAppend);
  }
  return {
    inserted: toAppend.length,
    skippedDuplicate: duplicateIds.length,
    acceptedIds: acceptedIds,
    duplicateIds: duplicateIds
  };
}

function handleSavePlatform_(payload, callerAdmin, gCfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = gCfg || readGlobalConfigMap_();
  const defaultOwner = String(cfg.adminUsername || 'admin').trim() || 'admin';
  const isSuper = Boolean(callerAdmin && callerAdmin.isSuperAdmin);

  // 若最高管理員修改了管理者帳號名稱，同步搬移對應活動與作業員的 ownerAdmin
  const renamedAdmins = (isSuper && payload.renamedAdmins && typeof payload.renamedAdmins === 'object')
    ? payload.renamedAdmins
    : {};
  const renameOwner_ = function (name) {
    const clean = String(name || '').trim();
    return renamedAdmins[clean] ? String(renamedAdmins[clean]).trim() : clean;
  };
  const effectiveCaller = renameOwner_(callerAdmin.username);

  // 1. 更新管理者帳號列表 (AdminAccounts)
  const adminHeaders = ['username', 'password', 'displayName', 'isSuperAdmin', 'enabled', 'updatedAt'];
  let currentAdmins = readAdminAccountsList_(cfg);
  let primarySuper = currentAdmins.find(function (a) { return a.isSuperAdmin && a.enabled; }) || currentAdmins[0];

  if (Array.isArray(payload.adminAccounts)) {
    const adminSheet = getOrCreateSheet_(ss, SHEET_NAMES.ADMINS, adminHeaders);
    let nextAdmins = currentAdmins.slice();

    if (isSuper) {
      const seen = {};
      nextAdmins = [];
      payload.adminAccounts.forEach(function (a) {
        const u = String(a.username || '').trim();
        if (!u || seen[u]) return;
        seen[u] = true;
        nextAdmins.push({
          username: u,
          password: String(a.password || ''),
          displayName: String(a.displayName || u).trim(),
          isSuperAdmin: Boolean(a.isSuperAdmin),
          enabled: a.enabled !== false,
          updatedAt: Number(a.updatedAt || Date.now())
        });
      });
      if (!nextAdmins.some(function (a) { return a.isSuperAdmin && a.enabled; })) {
        throw new Error('系統至少須保留一位啟用中的最高管理員帳號');
      }
    } else {
      // 一般活動管理者僅能修改自己的密碼與名稱
      const selfIncoming = payload.adminAccounts.find(function (a) {
        return String(a.username || '').trim() === callerAdmin.username;
      });
      if (selfIncoming) {
        nextAdmins = nextAdmins.map(function (a) {
          if (a.username !== callerAdmin.username) return a;
          return {
            username: a.username,
            password: String(selfIncoming.password || a.password),
            displayName: String(selfIncoming.displayName || a.displayName || a.username).trim(),
            isSuperAdmin: a.isSuperAdmin,
            enabled: a.enabled,
            updatedAt: Date.now()
          };
        });
      }
    }

    if (adminSheet.getLastRow() > 1) adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, 6).clearContent();
    if (nextAdmins.length > 0) {
      const adminRows = nextAdmins.map(function (a) {
        return [a.username, a.password, a.displayName, a.isSuperAdmin ? 1 : 0, a.enabled === false ? 0 : 1, a.updatedAt];
      });
      adminSheet.getRange(2, 1, adminRows.length, 6).setValues(adminRows);
    }
    currentAdmins = nextAdmins;
    primarySuper = currentAdmins.find(function (a) { return a.isSuperAdmin && a.enabled; }) || currentAdmins[0];
  }

  // 2. 更新活動清單（嚴格依 ownerAdmin 隔離合併，絕不覆蓋或刪除其他管理員的活動！）
  let finalAllEvents = readEventsList_(defaultOwner).map(function (ev) {
    ev.ownerAdmin = renameOwner_(ev.ownerAdmin);
    return ev;
  });
  let purgeResult = { eventIds: [], deletedRegistrations: 0, deletedScans: 0 };

  if (Array.isArray(payload.events) || Object.keys(renamedAdmins).length > 0) {
    const evtHeaders = [
      'eventId', 'name', 'date', 'location', 'status', 'debounceSeconds',
      'rankBy', 'categoriesJson', 'checkpointsJson', 'updatedAt', 'eventType', 'lapDistanceKm', 'ownerAdmin',
      'englishName', 'organization'
    ];
    const evtSheet = getOrCreateSheet_(ss, SHEET_NAMES.EVENTS, evtHeaders);
    // 確保舊版工作表自動補齊完整欄位標題 (ownerAdmin / englishName / organization)
    if (evtSheet.getLastColumn() < evtHeaders.length) {
      evtSheet.getRange(1, 1, 1, evtHeaders.length).setValues([evtHeaders])
        .setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }

    if (Array.isArray(payload.events)) {
      const prevCallerEventIds = finalAllEvents
        .filter(function (ev) { return ev.ownerAdmin === effectiveCaller; })
        .map(function (ev) { return ev.eventId; });
      const otherEvents = finalAllEvents.filter(function (ev) { return ev.ownerAdmin !== effectiveCaller; });
      const otherIdSet = {};
      otherEvents.forEach(function (ev) { otherIdSet[ev.eventId] = true; });

      const callerEvents = [];
      const nextCallerIdSet = {};
      payload.events.forEach(function (ev) {
        const eid = String(ev.eventId || '').trim();
        if (!eid || otherIdSet[eid]) return; // 嚴禁寫入/篡改其他管理員的活動 ID
        nextCallerIdSet[eid] = true;
        callerEvents.push({
          eventId: eid,
          name: String(ev.name || '').trim(),
          englishName: String(ev.englishName || '').trim(),
          organization: String(ev.organization || '').trim(),
          date: String(ev.date || '').trim(),
          location: String(ev.location || '').trim(),
          status: String(ev.status || 'active').trim(),
          debounceSeconds: Number(ev.debounceSeconds || 30),
          rankBy: String(ev.rankBy || 'gun').trim(),
          categories: Array.isArray(ev.categories) ? ev.categories : [],
          checkpoints: Array.isArray(ev.checkpoints) ? ev.checkpoints : [],
          updatedAt: Number(ev.updatedAt || Date.now()),
          eventType: String(ev.eventType || 'distance').trim(),
          lapDistanceKm: Number(ev.lapDistanceKm || 0),
          ownerAdmin: effectiveCaller
        });
      });

      finalAllEvents = otherEvents.concat(callerEvents);

      // 偵測本次被刪除的活動 ID，並連動清除 Registrations 與 Scans 中該活動的明細資料以減輕資料庫負荷
      const removedIdSet = {};
      prevCallerEventIds.forEach(function (eid) {
        if (!nextCallerIdSet[eid]) removedIdSet[eid] = true;
      });
      if (Array.isArray(payload.deletedEventIds)) {
        payload.deletedEventIds.forEach(function (rawId) {
          const eid = String(rawId || '').trim();
          if (eid && !otherIdSet[eid] && !nextCallerIdSet[eid]) {
            removedIdSet[eid] = true;
          }
        });
      }
      const removedEventIds = Object.keys(removedIdSet);
      if (removedEventIds.length > 0) {
        purgeResult = purgeEventsDetailData_(ss, removedEventIds);
      }
    }

    const prevEvtLastRow = evtSheet.getLastRow();
    const evtCols = Math.max(evtHeaders.length, evtSheet.getLastColumn());
    if (finalAllEvents.length > 0) {
      const rows = finalAllEvents.map(function (ev) {
        return [
          ev.eventId,
          ev.name,
          ev.date,
          ev.location,
          ev.status,
          ev.debounceSeconds,
          ev.rankBy,
          JSON.stringify(ev.categories || []),
          JSON.stringify(ev.checkpoints || []),
          ev.updatedAt,
          ev.eventType,
          ev.lapDistanceKm,
          ev.ownerAdmin,
          String(ev.englishName || '').trim(),
          String(ev.organization || '').trim()
        ];
      });
      const range = evtSheet.getRange(2, 1, rows.length, evtHeaders.length);
      evtSheet.getRange(2, 3, rows.length, 1).setNumberFormat('@');
      range.setValues(rows);
    }
    const newEvtLastRow = finalAllEvents.length + 1;
    if (prevEvtLastRow > newEvtLastRow) {
      evtSheet.getRange(newEvtLastRow + 1, 1, prevEvtLastRow - newEvtLastRow, evtCols).clearContent();
      const maxDeletable = Math.max(0, evtSheet.getMaxRows() - Math.max(newEvtLastRow + 1, 2));
      const rowsToDelete = Math.min(prevEvtLastRow - newEvtLastRow, maxDeletable);
      if (rowsToDelete > 0) {
        evtSheet.deleteRows(newEvtLastRow + 1, rowsToDelete);
      }
    }
  }

  // 3. 更新作業員帳號（嚴格依 ownerAdmin 隔離合併，防止跨管理員覆蓋或帳號衝突）
  if (Array.isArray(payload.staffAccounts) || Object.keys(renamedAdmins).length > 0) {
    const staffHeaders = [
      'username', 'password', 'displayName', 'role', 'eventId', 'checkpointId', 'enabled', 'updatedAt', 'ownerAdmin'
    ];
    const staffSheet = getOrCreateSheet_(ss, SHEET_NAMES.STAFF, staffHeaders);
    if (staffSheet.getLastColumn() < 9) {
      staffSheet.getRange(1, 1, 1, 9).setValues([staffHeaders])
        .setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }

    const existingStaff = readStaffList_().map(function (s) {
      s.ownerAdmin = s.ownerAdmin ? renameOwner_(s.ownerAdmin) : '';
      return s;
    });

    let finalStaff = existingStaff;
    if (Array.isArray(payload.staffAccounts)) {
      const otherStaff = existingStaff.filter(function (s) {
        if (isSuper) return s.ownerAdmin && s.ownerAdmin !== effectiveCaller;
        return s.ownerAdmin !== effectiveCaller;
      });
      const otherUserSet = {};
      otherStaff.forEach(function (s) { otherUserSet[s.username] = true; });

      const callerStaff = [];
      payload.staffAccounts.forEach(function (s) {
        const u = String(s.username || '').trim();
        if (!u) return;
        if (otherUserSet[u]) {
          throw new Error('作業員帳號「' + u + '」已被其他管理員使用，請更換其他帳號名稱');
        }
        const rawOwner = String(s.ownerAdmin || '').trim();
        callerStaff.push({
          username: u,
          password: String(s.password || ''),
          displayName: String(s.displayName || '').trim(),
          role: String(s.role || 'scanner').trim(),
          eventId: String(s.eventId || 'ALL').trim() || 'ALL',
          checkpointId: String(s.checkpointId || '').trim(),
          enabled: s.enabled !== false,
          updatedAt: Number(s.updatedAt || Date.now()),
          ownerAdmin: isSuper ? (rawOwner === '' ? '' : effectiveCaller) : effectiveCaller
        });
      });
      finalStaff = otherStaff.concat(callerStaff);
    }

    if (staffSheet.getLastRow() > 1) {
      staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, Math.max(9, staffSheet.getLastColumn())).clearContent();
    }
    if (finalStaff.length > 0) {
      const rows = finalStaff.map(function (s) {
        return [
          s.username,
          s.password,
          s.displayName,
          s.role,
          s.eventId,
          s.checkpointId,
          s.enabled === false ? 0 : 1,
          s.updatedAt,
          s.ownerAdmin
        ];
      });
      staffSheet.getRange(2, 1, rows.length, 9).setValues(rows);
    }
  }

  // 4. 更新 GlobalConfig
  if (payload.globalConfig || primarySuper) {
    const cfgSheet = getOrCreateSheet_(ss, SHEET_NAMES.GLOBAL_CONFIG, ['key', 'value', 'description']);
    const existing = readGlobalConfigMap_();
    let nextActiveEvt = String(existing.activeEventId || '');
    const reqActiveEvt = payload.globalConfig && payload.globalConfig.activeEventId
      ? String(payload.globalConfig.activeEventId).trim()
      : '';
    if (reqActiveEvt) {
      const targetEv = finalAllEvents.find(function (ev) { return ev.eventId === reqActiveEvt; });
      if (isSuper || (targetEv && targetEv.ownerAdmin === effectiveCaller)) {
        nextActiveEvt = reqActiveEvt;
      }
    }
    if (nextActiveEvt && !finalAllEvents.some(function (ev) { return ev.eventId === nextActiveEvt; })) {
      nextActiveEvt = finalAllEvents[0] ? finalAllEvents[0].eventId : '';
    }
    const merged = {
      adminUsername: primarySuper ? primarySuper.username : (existing.adminUsername || 'admin'),
      adminPassword: primarySuper ? primarySuper.password : (existing.adminPassword || 'admin1234'),
      activeEventId: nextActiveEvt
    };
    if (cfgSheet.getLastRow() > 1) cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 3).clearContent();
    cfgSheet.getRange(2, 1, 3, 3).setValues([
      ['adminUsername', String(merged.adminUsername), '後台最高管理員帳號'],
      ['adminPassword', String(merged.adminPassword), '後台最高管理員密碼'],
      ['activeEventId', String(merged.activeEventId), '預設顯示活動代碼']
    ]);
  }

  return { purge: purgeResult };
}

/**
 * 批次刪除指定工作表中符合 eventId 的所有資料列，並收縮多餘空白列以減輕資料庫負荷
 */
function deleteSheetRowsByEventIds_(sheet, eventIdSet) {
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return 0;

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function (h) { return String(h || '').trim(); });
  const evtCol = headers.indexOf('eventId');
  if (evtCol < 0) return 0;

  const keptRows = [];
  let deletedCount = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const hasData = row.some(function (cell) { return cell !== '' && cell !== null; });
    if (!hasData) continue;
    const rowEvtId = String(row[evtCol] || '').trim();
    if (rowEvtId && eventIdSet[rowEvtId]) {
      deletedCount++;
    } else {
      keptRows.push(row);
    }
  }

  if (deletedCount === 0) return 0;

  if (keptRows.length > 0) {
    sheet.getRange(2, 1, keptRows.length, lastCol).setValues(keptRows);
  }
  const newLastRow = keptRows.length + 1;
  if (lastRow > newLastRow) {
    sheet.getRange(newLastRow + 1, 1, lastRow - newLastRow, lastCol).clearContent();
    const maxDeletable = Math.max(0, sheet.getMaxRows() - Math.max(newLastRow + 1, 2));
    const rowsToDelete = Math.min(lastRow - newLastRow, maxDeletable);
    if (rowsToDelete > 0) {
      sheet.deleteRows(newLastRow + 1, rowsToDelete);
    }
  }
  return deletedCount;
}

/**
 * 連動清除已刪除活動的明細資料（參賽明細 Registrations ＋ 感應紀錄 Scans）
 * 跨活動共用晶片主檔 Chips 不受影響
 */
function purgeEventsDetailData_(ss, eventIds) {
  const cleanIds = (eventIds || [])
    .map(function (id) { return String(id || '').trim(); })
    .filter(function (id) { return Boolean(id); });
  if (!cleanIds.length) {
    return { eventIds: [], deletedRegistrations: 0, deletedScans: 0 };
  }
  const idSet = {};
  cleanIds.forEach(function (id) { idSet[id] = true; });

  const deletedRegistrations = deleteSheetRowsByEventIds_(ss.getSheetByName(SHEET_NAMES.REGISTRATIONS), idSet);
  const deletedScans = deleteSheetRowsByEventIds_(ss.getSheetByName(SHEET_NAMES.SCANS), idSet);

  return {
    eventIds: cleanIds,
    deletedRegistrations: deletedRegistrations,
    deletedScans: deletedScans
  };
}

function handleEditScan_(scan, callerAdmin, gCfg) {
  if (!scan || !scan.eventId || !scan.bib || !scan.checkpointId) {
    throw new Error('缺少必要的感應修正欄位 (eventId, bib, checkpointId)');
  }
  if (callerAdmin) {
    assertEventsOwnedByAdmin_([scan.eventId], callerAdmin, gCfg);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.SCANS, [
    'id', 'eventId', 'bib', 'idNumber', 'checkpointId', 'timestampMs', 'timeStr', 'isoTime',
    'source', 'deviceId', 'operator', 'deleted', 'note', 'updatedAt'
  ]);

  const values = sheet.getDataRange().getValues();
  let targetRow = -1;
  if (scan.id) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(scan.id)) {
        targetRow = i + 1;
        break;
      }
    }
  }

  const id = String(scan.id || Utilities.getUuid());
  const ts = Number(scan.timestampMs || Date.now());
  const rowData = [
    id,
    String(scan.eventId).trim(),
    String(scan.bib).trim(),
    String(scan.idNumber || '').trim().toUpperCase(),
    String(scan.checkpointId).trim(),
    ts,
    String(scan.timeStr || ''),
    String(scan.isoTime || new Date(ts).toISOString()),
    String(scan.source || 'admin'),
    String(scan.deviceId || 'admin-console'),
    String(scan.operator || (callerAdmin ? callerAdmin.username : 'admin')),
    scan.deleted ? 1 : 0,
    String(scan.note || '後台人工修正'),
    Date.now()
  ];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, 14).setValues([rowData]);
    return { action: 'updated', id: id };
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 14).setValues([rowData]);
    return { action: 'inserted', id: id };
  }
}

function handleUpsertHistory_(incomingList, callerAdmin) {
  if (!incomingList || !incomingList.length) return { inserted: 0, updated: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = [
    'id', 'year', 'eventDate', 'eventName', 'categoryName', 'bib', 'name', 'idNumber',
    'finishTime', 'distanceKm', 'laps', 'overallRank', 'categoryRank', 'status', 'note', 'ownerAdmin', 'updatedAt'
  ];
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.HISTORY, headers);

  const values = sheet.getDataRange().getValues();
  const rowById = {};
  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][0] || '').trim();
    if (id) rowById[id] = i + 1;
  }

  let inserted = 0;
  let updated = 0;
  const rowsToAppend = [];
  const pendingIdxById = {};
  const owner = callerAdmin ? callerAdmin.username : 'admin';

  incomingList.forEach(function (h) {
    const id = String(h.id || '').trim();
    const eventName = String(h.eventName || '').trim();
    const name = String(h.name || '').trim();
    if (!id || !eventName || !name) return;

    const evDate = formatDateCell_(h.eventDate);
    const yr = String(h.year || (evDate ? evDate.slice(0, 4) : '')).trim();
    const rowData = [
      id,
      yr,
      evDate,
      eventName,
      String(h.categoryName || '').trim(),
      String(h.bib || '').trim(),
      name,
      String(h.idNumber || '').trim().toUpperCase(),
      formatTimeCell_(h.finishTime),
      Number(h.distanceKm || 0),
      parseInt(h.laps || 0, 10) || 0,
      h.overallRank !== null && h.overallRank !== undefined && h.overallRank !== '' ? Number(h.overallRank) : '',
      h.categoryRank !== null && h.categoryRank !== undefined && h.categoryRank !== '' ? Number(h.categoryRank) : '',
      String(h.status || '完賽').trim(),
      String(h.note || '').trim(),
      String(h.ownerAdmin || owner).trim(),
      Number(h.updatedAt || Date.now())
    ];

    if (rowById[id] > 0) {
      const targetRow = rowById[id];
      sheet.getRange(targetRow, 2, 1, 2).setNumberFormat('@');
      sheet.getRange(targetRow, 6, 1, 1).setNumberFormat('@');
      sheet.getRange(targetRow, 9, 1, 1).setNumberFormat('@');
      sheet.getRange(targetRow, 1, 1, 17).setValues([rowData]);
      updated++;
    } else if (pendingIdxById[id] !== undefined) {
      rowsToAppend[pendingIdxById[id]] = rowData;
      updated++;
    } else {
      pendingIdxById[id] = rowsToAppend.length;
      rowsToAppend.push(rowData);
      inserted++;
    }
  });

  if (rowsToAppend.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 2, rowsToAppend.length, 2).setNumberFormat('@');
    sheet.getRange(startRow, 6, rowsToAppend.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 9, rowsToAppend.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, rowsToAppend.length, 17).setValues(rowsToAppend);
  }

  return { inserted: inserted, updated: updated };
}

function handleDeleteHistory_(ids) {
  if (!ids || !ids.length) return { deleted: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
  if (!sheet || sheet.getLastRow() <= 1) return { deleted: 0 };

  const delSet = {};
  ids.forEach(function (id) {
    const clean = String(id || '').trim();
    if (clean) delSet[clean] = true;
  });

  const data = sheet.getDataRange().getValues();
  const kept = [data[0]];
  let deleted = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    if (delSet[id]) deleted++;
    else kept.push(data[i]);
  }
  if (deleted === 0) return { deleted: 0 };

  sheet.clearContents();
  if (kept.length > 0) {
    if (kept.length > 1) {
      sheet.getRange(2, 2, kept.length - 1, 2).setNumberFormat('@');
      sheet.getRange(2, 6, kept.length - 1, 1).setNumberFormat('@');
      sheet.getRange(2, 9, kept.length - 1, 1).setNumberFormat('@');
    }
    sheet.getRange(1, 1, kept.length, kept[0].length).setValues(kept);
  }
  const maxRows = sheet.getMaxRows();
  const keepRows = Math.max(kept.length, 2);
  if (maxRows > keepRows) sheet.deleteRows(keepRows + 1, maxRows - keepRows);
  return { deleted: deleted };
}

/* ========================= 內部工具函式 ========================= */

function readTableObjects_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h).trim(); });
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row.some(function (cell) { return cell !== '' && cell !== null; })) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    result.push(obj);
  }
  return result;
}

function maskIdNumber_(idStr) {
  const s = String(idStr || '').trim().toUpperCase();
  if (!s) return '';
  if (s.length <= 4) return s.slice(0, 1) + '***';
  return s.slice(0, 4) + '****' + s.slice(-2);
}

function sha256Hex_(text) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    const v = b < 0 ? b + 256 : b;
    return (v < 16 ? '0' : '') + v.toString(16);
  }).join('');
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
