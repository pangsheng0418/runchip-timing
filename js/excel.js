/**
 * 智慧路跑 NFC 晶片計時系統 — Excel (.xlsx) 一鍵匯出模組 (SheetJS)
 */
(function (global) {
  'use strict';

  const STATUS_LABELS = {
    finished: '已完賽',
    on_course: '賽道上',
    not_started: '未出發'
  };

  function exportResultsToExcel({ config, computed, scans, maskId = false }) {
    if (typeof global.XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) 套件尚未載入，請確認網路連線或重新整理頁面');
    }

    const TE = global.TimingEngine;
    const wb = global.XLSX.utils.book_new();
    const checkpoints = computed.checkpoints || [];

    /* ==================== 工作表 1：成績總表 ==================== */
    const resultRows = computed.rows.map((row) => {
      const obj = {
        總名次: row.overallRank ?? '',
        分組名次: row.categoryRank ?? '',
        參賽編號: row.bib,
        選手姓名: row.name,
        身份證字號: maskId ? row.idMasked : (row.idNumber || row.idMasked),
        參賽分組: row.categoryName,
        狀態: STATUS_LABELS[row.status] || row.status,
        分組起跑時間: row.gunTimeMs ? TE.formatClockSec(row.gunTimeMs) : ''
      };

      // 各站點感應時間
      checkpoints.forEach((cp) => {
        const s = row.cpScanMap[cp.id];
        obj[`[站點] ${cp.name}`] = s ? TE.formatClockMs(s.timestampMs) : '';
      });

      // 各分段耗時與配速
      row.splits.forEach((sp, idx) => {
        obj[`[分段${idx + 1}] ${sp.fromName}→${sp.toName}`] =
          sp.durationMs !== null ? TE.formatDuration(sp.durationMs, false) : '';
        obj[`[配速${idx + 1}]`] = sp.pace !== '--' ? sp.pace : '';
      });

      obj['判罰加時(秒)'] = row.penaltySec || 0;
      obj['大會時間 (Gun Time)'] = row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, true) : '';
      obj['個人淨時間 (Chip Time)'] = row.chipDurationMs !== null ? TE.formatDuration(row.chipDurationMs, true) : '';
      obj['平均配速'] = row.avgPace !== '--' ? row.avgPace : '';

      return obj;
    });

    const ws1 = global.XLSX.utils.json_to_sheet(resultRows);
    ws1['!cols'] = Object.keys(resultRows[0] || { 參賽編號: '' }).map((k) => ({
      wch: Math.max(14, k.length * 2 + 2)
    }));
    global.XLSX.utils.book_append_sheet(wb, ws1, '成績總表');

    /* ==================== 工作表 2：各分組統計 ==================== */
    const statRows = (computed.summary.byCategory || []).map((c) => ({
      分組代碼: c.id,
      分組名稱: c.name,
      距離_km: c.distanceKm,
      鳴槍時間: c.gunTimeMs ? TE.formatClockSec(c.gunTimeMs) : '尚未鳴槍',
      總人數: c.total,
      已完賽: c.finished,
      賽道上: c.onCourse,
      未出發: c.notStarted,
      完賽率: c.total ? `${((c.finished / c.total) * 100).toFixed(1)}%` : '0%',
      最佳成績: c.bestMs !== null ? TE.formatDuration(c.bestMs, true) : '--',
      平均成績: c.avgMs !== null ? TE.formatDuration(c.avgMs, false) : '--'
    }));

    const ws2 = global.XLSX.utils.json_to_sheet(statRows);
    ws2['!cols'] = Object.keys(statRows[0] || { 分組名稱: '' }).map(() => ({ wch: 16 }));
    global.XLSX.utils.book_append_sheet(wb, ws2, '各分組統計');

    /* ==================== 工作表 3：原始感應紀錄 ==================== */
    const cpNameMap = new Map(checkpoints.map((cp) => [cp.id, cp.name]));
    const rawRows = (scans || [])
      .filter((s) => !s.deleted)
      .sort((a, b) => a.timestampMs - b.timestampMs)
      .map((s) => ({
        紀錄ID: s.id,
        參賽編號: s.bib,
        站點代碼: s.checkpointId,
        站點名稱: cpNameMap.get(s.checkpointId) || s.checkpointId,
        感應時間: TE.formatClockMs(s.timestampMs),
        ISO時間: s.isoTime,
        來源: s.source,
        操作員帳號: s.operator || '',
        裝置ID: s.deviceId || '',
        雲端同步: s.synced ? '已同步' : '待同步',
        備註: s.note || ''
      }));

    const ws3 = global.XLSX.utils.json_to_sheet(rawRows);
    ws3['!cols'] = Object.keys(rawRows[0] || { 紀錄ID: '' }).map(() => ({ wch: 18 }));
    global.XLSX.utils.book_append_sheet(wb, ws3, '原始感應紀錄');

    const safeRaceName = (config?.raceName || '路跑成績').replace(/[\\/:*?"<>|]/g, '_');
    const dateTag = config?.raceDate || new Date().toISOString().slice(0, 10);
    const fileName = `${safeRaceName}_成績報表_${dateTag}.xlsx`;

    global.XLSX.writeFile(wb, fileName);
    return fileName;
  }

  global.ExcelExporter = { exportResultsToExcel };
})(window);
