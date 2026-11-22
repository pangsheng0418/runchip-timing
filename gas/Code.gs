/**
 * ============================================================================
 * 智慧路跑 NFC 晶片計時系統 — Google Apps Script (GAS) 後端 API
 * ============================================================================
 * 部署步驟：
 * 1. 建立一個新的 Google 試算表，點選「擴充功能」>「Apps Script」
 * 2. 將本檔案全部貼入 Code.gs 並儲存
 * 3. 先在編輯器上方選擇函式「setupSheets」並點選「執行」，完成授權與工作表初始化
 * 4. 點選右上角「部署」>「新增部署作業」> 類型選「網頁應用程式」
 *    - 執行身分：我（試算表擁有者）
 *    - 誰可以存取：所有人（Anyone）
 * 5. 複製部署後的「網頁應用程式 URL」，貼入前端系統設定頁即可開始多機即時同步！
 * ============================================================================
 */

const SHEET_NAMES = {
  CONFIG: 'Config',
  CATEGORIES: 'Categories',
  CHECKPOINTS: 'Checkpoints',
  RUNNERS: 'Runners',
  SCANS: 'Scans'
};

/**
 * 首次初始化：自動建立所有工作表、標題列、預設賽事資料與欄位格式
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Config 工作表 (Key-Value)
  const configSheet = getOrCreateSheet_(ss, SHEET_NAMES.CONFIG, ['key', 'value', 'description']);
  if (configSheet.getLastRow() <= 1) {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    configSheet.getRange(2, 1, 5, 3).setValues([
      ['raceName', '2026 智慧城市馬拉松', '賽事名稱'],
      ['raceDate', today, '賽事日期 (YYYY-MM-DD)'],
      ['debounceSeconds', '30', '防連點冷卻秒數（同一選手同站點）'],
      ['adminPin', '1234', '後台管理與成績修正密碼'],
      ['rankBy', 'gun', '預設排名依據 (gun=大會時間, chip=個人淨時間)']
    ]);
  }

  // 2. Categories 工作表（組別與起跑時間）
  const catSheet = getOrCreateSheet_(ss, SHEET_NAMES.CATEGORIES, [
    'id', 'name', 'distanceKm', 'gunTimeMs', 'gunTimeIso', 'color'
  ]);
  if (catSheet.getLastRow() <= 1) {
    catSheet.getRange(2, 1, 3, 6).setValues([
      ['42K', '全馬組 42K', 42.195, '', '', '#ef4444'],
      ['21K', '半馬組 21K', 21.0975, '', '', '#3b82f6'],
      ['10K', '健康組 10K', 10, '', '', '#10b981']
    ]);
  }

  // 3. Checkpoints 工作表（感應站點設定）
  const cpSheet = getOrCreateSheet_(ss, SHEET_NAMES.CHECKPOINTS, [
    'id', 'name', 'order', 'distanceKm', 'type'
  ]);
  if (cpSheet.getLastRow() <= 1) {
    cpSheet.getRange(2, 1, 5, 5).setValues([
      ['START', '起點感應門', 1, 0, 'start'],
      ['CP1', '補給站 1 (5K)', 2, 5, 'checkpoint'],
      ['CP2', '折返點 / 補給站 2 (10K)', 3, 10, 'checkpoint'],
      ['CP3', '補給站 3 (21K)', 4, 21.1, 'checkpoint'],
      ['FINISH', '終點衝線門', 5, 42.195, 'finish']
    ]);
  }

  // 4. Runners 工作表（選手名單與罰時）
  getOrCreateSheet_(ss, SHEET_NAMES.RUNNERS, [
    'bib', 'name', 'idNumber', 'categoryId', 'nfcUid', 'penaltySec', 'checkedInAt', 'updatedAt'
  ]);

  // 5. Scans 工作表（晶片感應紀錄）
  getOrCreateSheet_(ss, SHEET_NAMES.SCANS, [
    'id', 'bib', 'checkpointId', 'timestampMs', 'timeStr', 'isoTime',
    'source', 'deviceId', 'deleted', 'note', 'updatedAt'
  ]);

  Logger.log('所有工作表已初始化完成！');
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
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
 * HTTP GET 入口
 * 參數：
 * - action=ping       : 測試連線
 * - action=publicData : 公開查詢用（自動隱碼身份證字號，回傳 SHA-256 雜湊供精準比對）
 * - action=fullSync   : 工作人員端完整拉取（需帶 pin 或為管理同步）
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || 'publicData';

    if (action === 'ping') {
      return jsonResponse_({
        ok: true,
        serverTimeMs: Date.now(),
        message: '智慧路跑計時 API 運作正常'
      });
    }

    if (action === 'publicData') {
      const data = readAllData_(true);
      return jsonResponse_({
        ok: true,
        serverTimeMs: Date.now(),
        data: data
      });
    }

    if (action === 'fullSync') {
      const config = readConfigMap_();
      const pin = params.pin || '';
      const includeSensitive = pin && String(pin) === String(config.adminPin || '1234');
      const data = readAllData_(!includeSensitive);
      return jsonResponse_({
        ok: true,
        serverTimeMs: Date.now(),
        authorized: includeSensitive,
        data: data
      });
    }

    return jsonResponse_({ ok: false, error: '未知的 action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) });
  }
}

/**
 * HTTP POST 入口
 * 前端請以 Content-Type: text/plain 傳送 JSON 字串以避免 CORS 預檢問題
 * 支援 action：
 * - pushScans    : 批次上傳感應紀錄（含伺服器端 30 秒防連點去重）
 * - upsertRunner : 新增/更新單一或多名選手（晶片寫入/報到/匯入）
 * - saveConfig   : 更新賽事設定、組別起跑時間、站點清單（需驗證 adminPin）
 * - editScan     : 管理員人工修正/補登/刪除感應時間（需驗證 adminPin）
 * - batchSync    : 離線恢復後一鍵同步（同時送出 runners + scans）
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // 最多等待 15 秒取得寫入鎖，避免多支手機同時感應造成資料錯亂
    lock.waitLock(15000);

    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(rawBody);
    const action = payload.action;

    if (!action) {
      return jsonResponse_({ ok: false, error: '缺少 action 參數' });
    }

    const config = readConfigMap_();
    const expectedPin = String(config.adminPin || '1234');

    if (action === 'pushScans') {
      const result = handlePushScans_(payload.scans || [], Number(config.debounceSeconds || 30));
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: result });
    }

    if (action === 'upsertRunners') {
      const result = handleUpsertRunners_(payload.runners || []);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: result });
    }

    if (action === 'batchSync') {
      const runnerRes = handleUpsertRunners_(payload.runners || []);
      const scanRes = handlePushScans_(payload.scans || [], Number(config.debounceSeconds || 30));
      return jsonResponse_({
        ok: true,
        serverTimeMs: Date.now(),
        result: { runners: runnerRes, scans: scanRes }
      });
    }

    if (action === 'saveConfig') {
      if (String(payload.adminPin || '') !== expectedPin) {
        return jsonResponse_({ ok: false, error: '管理員密碼錯誤' });
      }
      handleSaveConfig_(payload);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), message: '賽事設定已同步至雲端' });
    }

    if (action === 'editScan') {
      if (String(payload.adminPin || '') !== expectedPin) {
        return jsonResponse_({ ok: false, error: '管理員密碼錯誤' });
      }
      const result = handleEditScan_(payload.scan);
      return jsonResponse_({ ok: true, serverTimeMs: Date.now(), result: result });
    }

    return jsonResponse_({ ok: false, error: '不支援的 POST action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 讀取全部賽事資料
 * @param {boolean} maskSensitive 是否隱碼身份證字號（公開查詢頁為 true）
 */
function readAllData_(maskSensitive) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfigMap_();

  // 不將 adminPin 暴露給公開端
  const safeConfig = {
    raceName: config.raceName || '智慧路跑計時賽',
    raceDate: config.raceDate || '',
    debounceSeconds: Number(config.debounceSeconds || 30),
    rankBy: config.rankBy || 'gun'
  };

  const categories = readTableObjects_(ss.getSheetByName(SHEET_NAMES.CATEGORIES)).map(function (c) {
    return {
      id: String(c.id || ''),
      name: String(c.name || ''),
      distanceKm: Number(c.distanceKm || 0),
      gunTimeMs: c.gunTimeMs !== '' && c.gunTimeMs !== null ? Number(c.gunTimeMs) : null,
      gunTimeIso: String(c.gunTimeIso || ''),
      color: String(c.color || '#3b82f6')
    };
  }).filter(function (c) { return c.id; });

  const checkpoints = readTableObjects_(ss.getSheetByName(SHEET_NAMES.CHECKPOINTS)).map(function (cp) {
    return {
      id: String(cp.id || ''),
      name: String(cp.name || ''),
      order: Number(cp.order || 0),
      distanceKm: Number(cp.distanceKm || 0),
      type: String(cp.type || 'checkpoint')
    };
  }).filter(function (cp) { return cp.id; })
    .sort(function (a, b) { return a.order - b.order; });

  const runners = readTableObjects_(ss.getSheetByName(SHEET_NAMES.RUNNERS)).map(function (r) {
    const rawId = String(r.idNumber || '').trim().toUpperCase();
    const obj = {
      bib: String(r.bib || '').trim(),
      name: String(r.name || '').trim(),
      idMasked: maskIdNumber_(rawId),
      idHash: rawId ? sha256Hex_(rawId) : '',
      categoryId: String(r.categoryId || '').trim(),
      nfcUid: String(r.nfcUid || ''),
      penaltySec: Number(r.penaltySec || 0),
      checkedInAt: r.checkedInAt ? Number(r.checkedInAt) : null,
      updatedAt: r.updatedAt ? Number(r.updatedAt) : null
    };
    if (!maskSensitive) {
      obj.idNumber = rawId;
    }
    return obj;
  }).filter(function (r) { return r.bib; });

  const scans = readTableObjects_(ss.getSheetByName(SHEET_NAMES.SCANS)).map(function (s) {
    return {
      id: String(s.id || ''),
      bib: String(s.bib || '').trim(),
      checkpointId: String(s.checkpointId || '').trim(),
      timestampMs: Number(s.timestampMs || 0),
      timeStr: String(s.timeStr || ''),
      isoTime: String(s.isoTime || ''),
      source: String(s.source || 'nfc'),
      deviceId: String(s.deviceId || ''),
      deleted: String(s.deleted) === 'true' || Number(s.deleted) === 1,
      note: String(s.note || ''),
      updatedAt: s.updatedAt ? Number(s.updatedAt) : Number(s.timestampMs || 0)
    };
  }).filter(function (s) { return s.id && s.bib && !s.deleted; });

  return {
    config: safeConfig,
    categories: categories,
    checkpoints: checkpoints,
    runners: runners,
    scans: scans
  };
}

/**
 * 處理感應批次寫入（含雙重去重：相同 scan.id 去重 + 同一選手同站點防連點冷卻去重）
 */
function handlePushScans_(incomingScans, debounceSec) {
  if (!incomingScans || !incomingScans.length) {
    return { inserted: 0, skippedDuplicate: 0, acceptedIds: [], duplicateIds: [] };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.SCANS, [
    'id', 'bib', 'checkpointId', 'timestampMs', 'timeStr', 'isoTime',
    'source', 'deviceId', 'deleted', 'note', 'updatedAt'
  ]);

  const existing = readTableObjects_(sheet);
  const existingIdSet = {};
  // key: bib + '::' + checkpointId -> array of active timestamps
  const recentMap = {};

  existing.forEach(function (row) {
    const id = String(row.id || '');
    if (id) existingIdSet[id] = true;
    const isDeleted = String(row.deleted) === 'true' || Number(row.deleted) === 1;
    if (!isDeleted && row.bib && row.checkpointId && row.timestampMs) {
      const key = String(row.bib).trim() + '::' + String(row.checkpointId).trim();
      if (!recentMap[key]) recentMap[key] = [];
      recentMap[key].push(Number(row.timestampMs));
    }
  });

  const debounceMs = Math.max(0, Number(debounceSec || 30)) * 1000;
  const rowsToAppend = [];
  const acceptedIds = [];
  const duplicateIds = [];

  // 先依時間排序，確保批次上傳時防連點邏輯一致
  const sorted = incomingScans.slice().sort(function (a, b) {
    return Number(a.timestampMs || 0) - Number(b.timestampMs || 0);
  });

  sorted.forEach(function (scan) {
    const id = String(scan.id || Utilities.getUuid());
    const bib = String(scan.bib || '').trim();
    const cpId = String(scan.checkpointId || '').trim();
    const ts = Number(scan.timestampMs || Date.now());
    const isManual = scan.source === 'manual' || scan.source === 'admin';

    if (!bib || !cpId) return;

    // 1. 若 scan.id 已存在，視為已同步成功（冪等性）
    if (existingIdSet[id]) {
      acceptedIds.push(id);
      return;
    }

    // 2. 防連點檢查（人工補登/修正不受限）
    const key = bib + '::' + cpId;
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

    rowsToAppend.push([
      id,
      bib,
      cpId,
      ts,
      String(scan.timeStr || ''),
      String(scan.isoTime || new Date(ts).toISOString()),
      String(scan.source || 'nfc'),
      String(scan.deviceId || ''),
      scan.deleted ? 1 : 0,
      String(scan.note || ''),
      Date.now()
    ]);
    acceptedIds.push(id);
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 11).setValues(rowsToAppend);
  }

  return {
    inserted: rowsToAppend.length,
    skippedDuplicate: duplicateIds.length,
    acceptedIds: acceptedIds,
    duplicateIds: duplicateIds
  };
}

/**
 * 新增或更新選手名單（以 bib 為主鍵）
 */
function handleUpsertRunners_(incomingRunners) {
  if (!incomingRunners || !incomingRunners.length) {
    return { inserted: 0, updated: 0 };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.RUNNERS, [
    'bib', 'name', 'idNumber', 'categoryId', 'nfcUid', 'penaltySec', 'checkedInAt', 'updatedAt'
  ]);

  const values = sheet.getDataRange().getValues();
  const rowIndexByBib = {};
  for (let i = 1; i < values.length; i++) {
    const bib = String(values[i][0] || '').trim();
    if (bib) rowIndexByBib[bib] = i + 1; // 1-based row number
  }

  let inserted = 0;
  let updated = 0;
  const rowsToAppend = [];

  incomingRunners.forEach(function (r) {
    const bib = String(r.bib || '').trim();
    if (!bib) return;

    const rowData = [
      bib,
      String(r.name || '').trim(),
      String(r.idNumber || '').trim().toUpperCase(),
      String(r.categoryId || '').trim(),
      String(r.nfcUid || ''),
      Number(r.penaltySec || 0),
      r.checkedInAt ? Number(r.checkedInAt) : '',
      Number(r.updatedAt || Date.now())
    ];

    if (rowIndexByBib[bib]) {
      sheet.getRange(rowIndexByBib[bib], 1, 1, 8).setValues([rowData]);
      updated++;
    } else {
      rowsToAppend.push(rowData);
      rowIndexByBib[bib] = values.length + rowsToAppend.length;
      inserted++;
    }
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
  }

  return { inserted: inserted, updated: updated };
}

/**
 * 儲存賽事基本設定、組別起跑時間與站點
 */
function handleSaveConfig_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (payload.config) {
    const configSheet = getOrCreateSheet_(ss, SHEET_NAMES.CONFIG, ['key', 'value', 'description']);
    const existing = readConfigMap_();
    const merged = {
      raceName: payload.config.raceName !== undefined ? payload.config.raceName : existing.raceName,
      raceDate: payload.config.raceDate !== undefined ? payload.config.raceDate : existing.raceDate,
      debounceSeconds: payload.config.debounceSeconds !== undefined ? payload.config.debounceSeconds : existing.debounceSeconds,
      adminPin: payload.config.newAdminPin ? payload.config.newAdminPin : (existing.adminPin || '1234'),
      rankBy: payload.config.rankBy !== undefined ? payload.config.rankBy : (existing.rankBy || 'gun')
    };

    if (configSheet.getLastRow() > 1) {
      configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 3).clearContent();
    }
    configSheet.getRange(2, 1, 5, 3).setValues([
      ['raceName', String(merged.raceName || ''), '賽事名稱'],
      ['raceDate', String(merged.raceDate || ''), '賽事日期'],
      ['debounceSeconds', String(merged.debounceSeconds || 30), '防連點冷卻秒數'],
      ['adminPin', String(merged.adminPin || '1234'), '管理員密碼'],
      ['rankBy', String(merged.rankBy || 'gun'), '預設排名依據']
    ]);
  }

  if (Array.isArray(payload.categories)) {
    const catSheet = getOrCreateSheet_(ss, SHEET_NAMES.CATEGORIES, [
      'id', 'name', 'distanceKm', 'gunTimeMs', 'gunTimeIso', 'color'
    ]);
    if (catSheet.getLastRow() > 1) {
      catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 6).clearContent();
    }
    if (payload.categories.length > 0) {
      const catRows = payload.categories.map(function (c) {
        return [
          String(c.id || ''),
          String(c.name || ''),
          Number(c.distanceKm || 0),
          c.gunTimeMs ? Number(c.gunTimeMs) : '',
          c.gunTimeIso ? String(c.gunTimeIso) : '',
          String(c.color || '#3b82f6')
        ];
      });
      catSheet.getRange(2, 1, catRows.length, 6).setValues(catRows);
    }
  }

  if (Array.isArray(payload.checkpoints)) {
    const cpSheet = getOrCreateSheet_(ss, SHEET_NAMES.CHECKPOINTS, [
      'id', 'name', 'order', 'distanceKm', 'type'
    ]);
    if (cpSheet.getLastRow() > 1) {
      cpSheet.getRange(2, 1, cpSheet.getLastRow() - 1, 5).clearContent();
    }
    if (payload.checkpoints.length > 0) {
      const cpRows = payload.checkpoints.map(function (cp) {
        return [
          String(cp.id || ''),
          String(cp.name || ''),
          Number(cp.order || 0),
          Number(cp.distanceKm || 0),
          String(cp.type || 'checkpoint')
        ];
      });
      cpSheet.getRange(2, 1, cpRows.length, 5).setValues(cpRows);
    }
  }
}

/**
 * 管理員人工修正/補登/刪除單筆感應紀錄
 */
function handleEditScan_(scan) {
  if (!scan || !scan.bib || !scan.checkpointId) {
    throw new Error('缺少必要的感應修正欄位 (bib, checkpointId)');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.SCANS, [
    'id', 'bib', 'checkpointId', 'timestampMs', 'timeStr', 'isoTime',
    'source', 'deviceId', 'deleted', 'note', 'updatedAt'
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
    String(scan.bib).trim(),
    String(scan.checkpointId).trim(),
    ts,
    String(scan.timeStr || ''),
    String(scan.isoTime || new Date(ts).toISOString()),
    String(scan.source || 'admin'),
    String(scan.deviceId || 'admin-console'),
    scan.deleted ? 1 : 0,
    String(scan.note || '後台人工修正'),
    Date.now()
  ];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, 11).setValues([rowData]);
    return { action: 'updated', id: id };
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 11).setValues([rowData]);
    return { action: 'inserted', id: id };
  }
}

/* ========================= 內部工具函式 ========================= */

function readConfigMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const map = {};
  if (!sheet || sheet.getLastRow() <= 1) return map;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  rows.forEach(function (r) {
    if (r[0]) map[String(r[0]).trim()] = r[1];
  });
  return map;
}

function readTableObjects_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h).trim(); });
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const hasValue = row.some(function (cell) { return cell !== '' && cell !== null; });
    if (!hasValue) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
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
