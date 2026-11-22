/**
 * ============================================================================
 * 智慧路跑 NFC 晶片計時平台 v3.0 — 多場次共用晶片 Google Apps Script (GAS) 後端
 * ============================================================================
 * 核心設計：
 * - 平台支援建立多場次路跑活動 (Events)，每場擁有獨立分組、鳴槍時間、站點與參賽名單
 * - 共用晶片主檔 (Chips)：晶片內僅存「姓名｜身份證字號」，同一晶片可重複參加不同場次與組別
 * - 參賽明細 (Registrations)：事前匯入各活動選手名單與組別，感應時依晶片身份證自動比對
 * - 三大權限分區：公開查詢（隱碼免登入）、作業員帳密鎖定站點、最高管理員獨立帳密
 *
 * 部署步驟：
 * 1. 開啟 Google 試算表 →「擴充功能」→「Apps Script」
 * 2. 貼入本檔案全部內容並儲存
 * 3. 選擇函式「setupSheets」並點選「執行」（自動建立 6 張工作表）
 * 4. 點選「部署」→「新增部署作業」（或「管理部署作業」→「建立新版本」）
 *    - 類型：網頁應用程式 ｜ 執行身分：我 ｜ 誰可以存取：所有人 (Anyone)
 * ============================================================================
 */

const SHEET_NAMES = {
  GLOBAL_CONFIG: 'GlobalConfig',
  EVENTS: 'Events',
  STAFF: 'StaffAccounts',
  CHIPS: 'Chips',
  REGISTRATIONS: 'Registrations',
  SCANS: 'Scans'
};

/**
 * 首次初始化：自動建立 6 張工作表與預設範例活動/帳號
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const now = Date.now();

  // 1. GlobalConfig（平台管理員帳密與預設活動）
  const cfgSheet = getOrCreateSheet_(ss, SHEET_NAMES.GLOBAL_CONFIG, ['key', 'value', 'description']);
  if (cfgSheet.getLastRow() <= 1) {
    cfgSheet.getRange(2, 1, 3, 3).setValues([
      ['adminUsername', 'admin', '後台最高管理員帳號'],
      ['adminPassword', 'admin1234', '後台最高管理員密碼'],
      ['activeEventId', 'EVT-2026-01', '預設顯示活動代碼']
    ]);
  }

  // 2. Events（多場次路跑活動清單：支援一般距離賽 distance 與繞圈計時賽 lap_timed）
  const evtSheet = getOrCreateSheet_(ss, SHEET_NAMES.EVENTS, [
    'eventId', 'name', 'date', 'location', 'status', 'debounceSeconds',
    'rankBy', 'categoriesJson', 'checkpointsJson', 'updatedAt', 'eventType', 'lapDistanceKm'
  ]);
  if (evtSheet.getLastRow() <= 1) {
    const defaultCats = JSON.stringify([
      { id: '42K', name: '全馬組 42K', distanceKm: 42.195, gunTimeMs: null, gunTimeIso: '', color: '#ef4444' },
      { id: '21K', name: '半馬組 21K', distanceKm: 21.0975, gunTimeMs: null, gunTimeIso: '', color: '#3b82f6' },
      { id: '10K', name: '健康組 10K', distanceKm: 10, gunTimeMs: null, gunTimeIso: '', color: '#10b981' }
    ]);
    const defaultCps = JSON.stringify([
      { id: 'START', name: '起點感應門', order: 1, distanceKm: 0, type: 'start' },
      { id: 'CP1', name: '補給站 1 (5K)', order: 2, distanceKm: 5, type: 'checkpoint' },
      { id: 'CP2', name: '折返點 / 補給站 2 (10K)', order: 3, distanceKm: 10, type: 'checkpoint' },
      { id: 'CP3', name: '補給站 3 (21K)', order: 4, distanceKm: 21.1, type: 'checkpoint' },
      { id: 'FINISH', name: '終點衝線門', order: 5, distanceKm: 42.195, type: 'finish' }
    ]);
    evtSheet.getRange(2, 3, 1, 1).setNumberFormat('@');
    evtSheet.getRange(2, 1, 1, 12).setValues([[
      'EVT-2026-01', '2026 智慧城市馬拉松', today, '台北市政府廣場', 'active',
      30, 'gun', defaultCats, defaultCps, now, 'distance', 0
    ]]);
  }

  // 3. StaffAccounts（作業人員帳密、權限與鎖定活動/站點）
  const staffSheet = getOrCreateSheet_(ss, SHEET_NAMES.STAFF, [
    'username', 'password', 'displayName', 'role', 'eventId', 'checkpointId', 'enabled', 'updatedAt'
  ]);
  if (staffSheet.getLastRow() <= 1) {
    staffSheet.getRange(2, 1, 7, 8).setValues([
      ['writer01', '1111', '晶片寫入報到組', 'writer', 'ALL', '', 1, now],
      ['lap01', '1111', '繞圈計次感應組', 'scanner', 'ALL', 'LAP', 1, now],
      ['start01', '1111', '起點感應組', 'scanner', 'ALL', 'START', 1, now],
      ['cp1', '1111', '補給站 1 感應組', 'scanner', 'ALL', 'CP1', 1, now],
      ['cp2', '1111', '補給站 2 / 折返點組', 'scanner', 'ALL', 'CP2', 1, now],
      ['cp3', '1111', '補給站 3 感應組', 'scanner', 'ALL', 'CP3', 1, now],
      ['finish01', '1111', '終點衝線感應組', 'scanner', 'ALL', 'FINISH', 1, now]
    ]);
  }

  // 4. Chips（跨活動共用晶片主檔：僅存姓名、身份證字號與晶片硬體 UID）
  getOrCreateSheet_(ss, SHEET_NAMES.CHIPS, [
    'idNumber', 'name', 'nfcUid', 'writtenAt', 'updatedAt'
  ]);

  // 5. Registrations（各活動參賽選手明細與組別）
  getOrCreateSheet_(ss, SHEET_NAMES.REGISTRATIONS, [
    'regId', 'eventId', 'bib', 'name', 'idNumber', 'categoryId', 'penaltySec', 'updatedAt'
  ]);

  // 6. Scans（各活動晶片感應紀錄）
  getOrCreateSheet_(ss, SHEET_NAMES.SCANS, [
    'id', 'eventId', 'bib', 'idNumber', 'checkpointId', 'timestampMs', 'timeStr', 'isoTime',
    'source', 'deviceId', 'operator', 'deleted', 'note', 'updatedAt'
  ]);

  Logger.log('v3.0 多場次共用晶片計時平台 6 張工作表已初始化完成！');
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
        serverTimeMs: Date.now(),
        message: '智慧路跑多場次計時平台 API v3.0 運作正常'
      });
    }

    if (action === 'publicData') {
      const data = readAllData_(true, false);
      return jsonResponse_({
        ok: true,
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
    const adminUser = String(gCfg.adminUsername || 'admin').trim();
    const adminPass = String(gCfg.adminPassword || 'admin1234');

    // 1. 登入驗證（作業人員或最高管理員）
    if (action === 'login') {
      const mode = payload.mode || 'staff';
      const u = String(payload.username || '').trim();
      const p = String(payload.password || '');

      if (mode === 'admin') {
        if (u !== adminUser || p !== adminPass) {
          return jsonResponse_({ ok: false, error: '管理員帳號或密碼錯誤' });
        }
        return jsonResponse_({
          ok: true,
          serverTimeMs: Date.now(),
          account: { username: adminUser, displayName: '最高系統管理員', role: 'admin' },
          data: readAllData_(false, true)
        });
      }

      const staffList = readStaffList_();
      const matched = staffList.find(function (s) {
        return s.enabled && s.username === u && s.password === p;
      });
      if (!matched) {
        return jsonResponse_({ ok: false, error: '作業員帳號或密碼錯誤，或帳號已停用' });
      }
      return jsonResponse_({
        ok: true,
        serverTimeMs: Date.now(),
        account: matched,
        data: readAllData_(false, false)
      });
    }

    // 2. 授權同步拉取
    if (action === 'fullSync') {
      const isAdmin = verifyAdmin_(payload.auth, adminUser, adminPass);
      const isStaff = !isAdmin && verifyStaff_(payload.auth);
      if (!isAdmin && !isStaff) {
        return jsonResponse_({ ok: false, error: '身分驗證失敗，請重新登入' });
      }
      return jsonResponse_({
        ok: true,
        serverTimeMs: Date.now(),
        isAdmin: isAdmin,
        data: readAllData_(false, isAdmin)
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
      const res = handleUpsertRegistrations_(payload.registrations || []);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: res });
    }

    if (action === 'batchSync') {
      // 若為最高管理員帶入 auth + platform 資料，同步更新所有活動、分組鳴槍與作業員帳號
      let platformSynced = false;
      if (payload.platform && verifyAdmin_(payload.auth, adminUser, adminPass)) {
        handleSavePlatform_(payload.platform);
        platformSynced = true;
      }
      const chipRes = handleUpsertChips_(payload.chips || []);
      const regRes = handleUpsertRegistrations_(payload.registrations || []);
      const scanRes = handlePushScans_(payload.scans || []);
      return jsonResponse_({
        ok: true,
        serverTimeMs: Date.now(),
        result: { platformSynced: platformSynced, chips: chipRes, registrations: regRes, scans: scanRes }
      });
    }

    // 以下需最高管理員驗證
    if (!verifyAdmin_(payload.auth, adminUser, adminPass)) {
      return jsonResponse_({ ok: false, error: '權限不足：需要最高管理員帳號密碼驗證' });
    }

    if (action === 'savePlatform') {
      handleSavePlatform_(payload);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), message: '平台設定、活動與帳號已同步至雲端' });
    }

    if (action === 'editScan') {
      const res = handleEditScan_(payload.scan);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: res });
    }

    return jsonResponse_({ ok: false, error: '不支援的 POST action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function verifyAdmin_(auth, adminUser, adminPass) {
  if (!auth) return false;
  return String(auth.username || '').trim() === adminUser && String(auth.password || '') === adminPass;
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

function readEventsList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.EVENTS);
  return readTableObjects_(sheet).map(function (ev) {
    let cats = [];
    let cps = [];
    try { cats = JSON.parse(ev.categoriesJson || '[]'); } catch (e) {}
    try { cps = JSON.parse(ev.checkpointsJson || '[]'); } catch (e) {}
    const eventType = String(ev.eventType || 'distance').trim() || 'distance';
    return {
      eventId: String(ev.eventId || '').trim(),
      name: String(ev.name || '').trim(),
      date: formatDateCell_(ev.date),
      location: String(ev.location || '').trim(),
      status: String(ev.status || 'active').trim(),
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

function readAllData_(maskSensitive, includeAdminSecrets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gCfg = readGlobalConfigMap_();
  const events = readEventsList_();

  const globalConfig = {
    activeEventId: String(gCfg.activeEventId || (events[0] ? events[0].eventId : '')).trim()
  };
  if (includeAdminSecrets) {
    globalConfig.adminUsername = String(gCfg.adminUsername || 'admin');
    globalConfig.adminPassword = String(gCfg.adminPassword || 'admin1234');
  }

  const registrations = readTableObjects_(ss.getSheetByName(SHEET_NAMES.REGISTRATIONS)).map(function (r) {
    const rawId = String(r.idNumber || '').trim().toUpperCase();
    const obj = {
      regId: String(r.regId || (String(r.eventId).trim() + '__' + (rawId || String(r.bib).trim()))),
      eventId: String(r.eventId || '').trim(),
      bib: String(r.bib || '').trim(),
      name: String(r.name || '').trim(),
      idMasked: maskIdNumber_(rawId),
      idHash: rawId ? sha256Hex_(rawId) : '',
      categoryId: String(r.categoryId || '').trim(),
      penaltySec: Number(r.penaltySec || 0),
      updatedAt: r.updatedAt ? Number(r.updatedAt) : Date.now()
    };
    if (!maskSensitive) obj.idNumber = rawId;
    return obj;
  }).filter(function (r) { return r.eventId && r.bib; });

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
  }).filter(function (s) { return s.id && s.eventId && s.bib && !s.deleted; });

  const result = {
    globalConfig: globalConfig,
    events: events,
    registrations: registrations,
    scans: scans
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

  if (includeAdminSecrets) {
    result.staffAccounts = readStaffList_();
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
    if (rowById[id]) {
      sheet.getRange(rowById[id], 1, 1, 5).setValues([row]);
      updated++;
    } else {
      toAppend.push(row);
      rowById[id] = values.length + toAppend.length;
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
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.REGISTRATIONS, [
    'regId', 'eventId', 'bib', 'name', 'idNumber', 'categoryId', 'penaltySec', 'updatedAt'
  ]);

  const values = sheet.getDataRange().getValues();
  const rowByKey = {};
  for (let i = 1; i < values.length; i++) {
    const rk = String(values[i][0] || '').trim();
    if (rk) rowByKey[rk] = i + 1;
  }

  let inserted = 0, updated = 0;
  const toAppend = [];

  incomingRegs.forEach(function (r) {
    const eventId = String(r.eventId || '').trim();
    const bib = String(r.bib || '').trim();
    const idNumber = String(r.idNumber || '').trim().toUpperCase();
    if (!eventId || !bib) return;
    const regId = String(r.regId || (eventId + '__' + (idNumber || bib))).trim();

    const row = [
      regId,
      eventId,
      bib,
      String(r.name || '').trim(),
      idNumber,
      String(r.categoryId || '').trim(),
      Number(r.penaltySec || 0),
      Number(r.updatedAt || Date.now())
    ];

    if (rowByKey[regId]) {
      sheet.getRange(rowByKey[regId], 1, 1, 8).setValues([row]);
      updated++;
    } else {
      toAppend.push(row);
      rowByKey[regId] = values.length + toAppend.length;
      inserted++;
    }
  });

  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 8).setValues(toAppend);
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

function handleSavePlatform_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (payload.globalConfig) {
    const cfgSheet = getOrCreateSheet_(ss, SHEET_NAMES.GLOBAL_CONFIG, ['key', 'value', 'description']);
    const existing = readGlobalConfigMap_();
    const merged = {
      adminUsername: payload.globalConfig.adminUsername || existing.adminUsername || 'admin',
      adminPassword: payload.globalConfig.adminPassword || existing.adminPassword || 'admin1234',
      activeEventId: payload.globalConfig.activeEventId || existing.activeEventId || ''
    };
    if (cfgSheet.getLastRow() > 1) cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 3).clearContent();
    cfgSheet.getRange(2, 1, 3, 3).setValues([
      ['adminUsername', String(merged.adminUsername), '後台最高管理員帳號'],
      ['adminPassword', String(merged.adminPassword), '後台最高管理員密碼'],
      ['activeEventId', String(merged.activeEventId), '預設顯示活動代碼']
    ]);
  }

  if (Array.isArray(payload.events)) {
    const evtHeaders = [
      'eventId', 'name', 'date', 'location', 'status', 'debounceSeconds',
      'rankBy', 'categoriesJson', 'checkpointsJson', 'updatedAt', 'eventType', 'lapDistanceKm'
    ];
    const evtSheet = getOrCreateSheet_(ss, SHEET_NAMES.EVENTS, evtHeaders);
    // 確保舊版 10 欄工作表自動補齊第 11、12 欄標題
    if (evtSheet.getLastColumn() < 12) {
      evtSheet.getRange(1, 1, 1, 12).setValues([evtHeaders])
        .setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }
    if (evtSheet.getLastRow() > 1) evtSheet.getRange(2, 1, evtSheet.getLastRow() - 1, 12).clearContent();
    if (payload.events.length > 0) {
      const rows = payload.events.map(function (ev) {
        return [
          String(ev.eventId || '').trim(),
          String(ev.name || '').trim(),
          String(ev.date || '').trim(),
          String(ev.location || '').trim(),
          String(ev.status || 'active').trim(),
          Number(ev.debounceSeconds || 30),
          String(ev.rankBy || 'gun').trim(),
          JSON.stringify(ev.categories || []),
          JSON.stringify(ev.checkpoints || []),
          Number(ev.updatedAt || Date.now()),
          String(ev.eventType || 'distance').trim(),
          Number(ev.lapDistanceKm || 0)
        ];
      });
      const range = evtSheet.getRange(2, 1, rows.length, 12);
      evtSheet.getRange(2, 3, rows.length, 1).setNumberFormat('@');
      range.setValues(rows);
    }
  }

  if (Array.isArray(payload.staffAccounts)) {
    const staffSheet = getOrCreateSheet_(ss, SHEET_NAMES.STAFF, [
      'username', 'password', 'displayName', 'role', 'eventId', 'checkpointId', 'enabled', 'updatedAt'
    ]);
    if (staffSheet.getLastRow() > 1) staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 8).clearContent();
    if (payload.staffAccounts.length > 0) {
      const rows = payload.staffAccounts.map(function (s) {
        return [
          String(s.username || '').trim(),
          String(s.password || ''),
          String(s.displayName || '').trim(),
          String(s.role || 'scanner').trim(),
          String(s.eventId || 'ALL').trim() || 'ALL',
          String(s.checkpointId || '').trim(),
          s.enabled === false ? 0 : 1,
          Number(s.updatedAt || Date.now())
        ];
      });
      staffSheet.getRange(2, 1, rows.length, 8).setValues(rows);
    }
  }
}

function handleEditScan_(scan) {
  if (!scan || !scan.eventId || !scan.bib || !scan.checkpointId) {
    throw new Error('缺少必要的感應修正欄位 (eventId, bib, checkpointId)');
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
    String(scan.operator || 'admin'),
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
