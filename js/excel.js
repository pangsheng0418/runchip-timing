/**
 * 智慧路跑 NFC 晶片計時平台 v3.3 — 支援「一般距離賽」與「繞圈計時賽」雙模式 Excel (.xlsx) 匯出
 */
(function (global) {
  'use strict';

  const STATUS_LABELS = {
    finished: '已完賽',
    on_course: '賽道上',
    checked_in: '已報到',
    not_checked_in: '未報到',
    not_started: '未出發'
  };

  function exportLapTimedToExcel({ event, computed, scans, maskId = false }) {
    const TE = global.TimingEngine;
    const wb = global.XLSX.utils.book_new();
    const maxLaps = computed.rows.reduce((m, r) => Math.max(m, r.lapCount || 0), 0);

    // 工作表 1：繞圈成績總表（依總里程排名，含報到狀態）
    const rows1 = computed.rows.map((row) => {
      const obj = {
        活動名稱: event?.name || '',
        總名次: row.overallRank ?? '',
        分組名次: row.categoryRank ?? '',
        參賽編號: row.bib,
        選手姓名: row.name,
        身份證字號: maskId ? row.idMasked : (row.idNumber || row.idMasked),
        限時分組: `${row.categoryName} (${row.durationHours}H)`,
        報到狀態: row.checkedIn ? '已報到' : '未報到',
        報到時間: row.checkInTimeMs ? TE.formatClockSec(row.checkInTimeMs) : '',
        累積圈數: row.lapCount,
        單圈距離_km: row.lapDistanceKm,
        總里程_km: Number(row.totalDistanceKm.toFixed(3)),
        累積耗時: row.totalElapsedMs !== null ? TE.formatDuration(row.totalElapsedMs, true) : '',
        最後通過時間: row.lastLapTimeMs ? TE.formatClockMs(row.lastLapTimeMs) : '',
        最速單圈: row.bestLapMs !== null ? TE.formatDuration(row.bestLapMs, false) : '',
        平均單圈: row.avgLapMs !== null ? TE.formatDuration(row.avgLapMs, false) : '',
        平均配速: row.avgPace !== '--' ? row.avgPace : '',
        超時圈數: row.overtimeCount || 0
      };
      for (let i = 0; i < maxLaps; i++) {
        const lap = row.laps.filter((l) => !l.overtime)[i];
        obj[`第${i + 1}圈`] = lap ? TE.formatClockSec(lap.timestampMs) : '';
      }
      return obj;
    });

    const ws1 = global.XLSX.utils.json_to_sheet(rows1);
    ws1['!cols'] = Object.keys(rows1[0] || { 參賽編號: '' }).map((k) => ({ wch: Math.max(13, k.length * 2 + 2) }));
    global.XLSX.utils.book_append_sheet(wb, ws1, '繞圈成績總表');

    // 工作表 2：各限時分組里程統計
    const rows2 = (computed.summary.byCategory || []).map((c) => ({
      活動名稱: event?.name || '',
      分組代碼: c.id,
      分組名稱: c.name,
      限時小時: c.durationHours,
      律定起跑時間: c.gunTimeMs ? TE.formatClockSec(c.gunTimeMs) : '尚未設定',
      截止時間: c.cutoffTimeMs ? TE.formatClockSec(c.cutoffTimeMs) : '--',
      參賽人數: c.total,
      已報到人數: c.checkedIn ?? 0,
      未報到人數: c.notCheckedIn ?? 0,
      已有圈數人數: c.withLaps,
      全組總圈數: c.sumLaps,
      最高里程_km: c.topKm,
      平均里程_km: c.avgKm
    }));
    const ws2 = global.XLSX.utils.json_to_sheet(rows2);
    ws2['!cols'] = Object.keys(rows2[0] || { 分組名稱: '' }).map(() => ({ wch: 16 }));
    global.XLSX.utils.book_append_sheet(wb, ws2, '各分組里程統計');

    // 工作表 3：逐圈原始感應明細
    const rows3 = [];
    for (const row of computed.rows) {
      for (const lap of row.laps) {
        rows3.push({
          活動代碼: event?.eventId || '',
          參賽編號: row.bib,
          選手姓名: row.name,
          分組: row.categoryName,
          圈次: lap.overtime ? '超時不計' : lap.lapNumber,
          感應時間: TE.formatClockMs(lap.timestampMs),
          單圈耗時: lap.lapDurationMs !== null ? TE.formatDuration(lap.lapDurationMs, false) : '',
          單圈配速: lap.lapPace !== '--' ? lap.lapPace : '',
          累積里程_km: lap.cumulativeKm !== null ? lap.cumulativeKm : '',
          來源: lap.source || '',
          備註: lap.note || ''
        });
      }
    }
    const ws3 = global.XLSX.utils.json_to_sheet(rows3);
    ws3['!cols'] = Object.keys(rows3[0] || { 參賽編號: '' }).map(() => ({ wch: 16 }));
    global.XLSX.utils.book_append_sheet(wb, ws3, '逐圈原始感應紀錄');

    const safeName = (event?.name || '繞圈計時賽').replace(/[\\/:*?"<>|]/g, '_');
    const dateTag = event?.date || new Date().toISOString().slice(0, 10);
    const suffix = maskId ? '繞圈公開隱碼版' : '繞圈完整報表';
    const fileName = `${safeName}_${suffix}_${dateTag}.xlsx`;
    global.XLSX.writeFile(wb, fileName);
    return fileName;
  }

  function exportResultsToExcel({ event, computed, scans, maskId = false }) {
    if (typeof global.XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) 套件尚未載入，請確認網路連線或重新整理頁面');
    }

    if (event?.eventType === 'lap_timed' || computed?.eventType === 'lap_timed') {
      return exportLapTimedToExcel({ event, computed, scans, maskId });
    }

    const TE = global.TimingEngine;
    const wb = global.XLSX.utils.book_new();
    const checkpoints = computed.checkpoints || [];

    const resultRows = computed.rows.map((row) => {
      const obj = {
        活動名稱: event?.name || '',
        總名次: row.overallRank ?? '',
        分組名次: row.categoryRank ?? '',
        參賽編號: row.bib,
        選手姓名: row.name,
        身份證字號: maskId ? row.idMasked : (row.idNumber || row.idMasked),
        參賽分組: row.categoryName,
        報到狀態: row.checkedIn ? '已報到' : '未報到',
        報到時間: row.checkInTimeMs ? TE.formatClockSec(row.checkInTimeMs) : '',
        狀態: STATUS_LABELS[row.status] || row.status,
        律定起跑時間: row.gunTimeMs ? TE.formatClockSec(row.gunTimeMs) : ''
      };

      checkpoints.forEach((cp) => {
        const s = row.cpScanMap[cp.id];
        obj[`[站點] ${cp.name}`] = s ? TE.formatClockMs(s.timestampMs) : '';
      });

      row.splits.forEach((sp, idx) => {
        obj[`[分段${idx + 1}] ${sp.fromName}→${sp.toName}`] =
          sp.durationMs !== null ? TE.formatDuration(sp.durationMs, false) : '';
        obj[`[配速${idx + 1}]`] = sp.pace !== '--' ? sp.pace : '';
      });

      obj['判罰加時(秒)'] = row.penaltySec || 0;
      obj['完賽時間'] = row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, true) : '';
      obj['平均配速'] = row.avgPace !== '--' ? row.avgPace : '';
      return obj;
    });

    const ws1 = global.XLSX.utils.json_to_sheet(resultRows);
    ws1['!cols'] = Object.keys(resultRows[0] || { 參賽編號: '' }).map((k) => ({ wch: Math.max(14, k.length * 2 + 2) }));
    global.XLSX.utils.book_append_sheet(wb, ws1, '成績總表');

    const statRows = (computed.summary.byCategory || []).map((c) => ({
      活動名稱: event?.name || '',
      分組代碼: c.id,
      分組名稱: c.name,
      距離_km: c.distanceKm,
      律定起跑時間: c.gunTimeMs ? TE.formatClockSec(c.gunTimeMs) : '尚未設定',
      總人數: c.total,
      已報到: c.checkedIn ?? 0,
      未報到: c.notCheckedIn ?? 0,
      已完賽: c.finished,
      賽道上: c.onCourse,
      完賽率: c.total ? `${((c.finished / c.total) * 100).toFixed(1)}%` : '0%',
      最佳成績: c.bestMs !== null ? TE.formatDuration(c.bestMs, true) : '--',
      平均成績: c.avgMs !== null ? TE.formatDuration(c.avgMs, false) : '--'
    }));

    const ws2 = global.XLSX.utils.json_to_sheet(statRows);
    ws2['!cols'] = Object.keys(statRows[0] || { 分組名稱: '' }).map(() => ({ wch: 16 }));
    global.XLSX.utils.book_append_sheet(wb, ws2, '各分組統計');

    const cpNameMap = new Map([
      ['CHECKIN', '選手報到感應區'],
      ...checkpoints.map((cp) => [cp.id, cp.name])
    ]);
    const rawRows = (scans || [])
      .filter((s) => !s.deleted)
      .sort((a, b) => a.timestampMs - b.timestampMs)
      .map((s) => ({
        紀錄ID: s.id,
        活動代碼: s.eventId || event?.eventId || '',
        參賽編號: s.bib,
        身份證字號: maskId ? TE.maskIdNumber(s.idNumber) : (s.idNumber || ''),
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

    const safeName = (event?.name || '路跑成績').replace(/[\\/:*?"<>|]/g, '_');
    const dateTag = event?.date || new Date().toISOString().slice(0, 10);
    const suffix = maskId ? '公開隱碼版' : '完整報表';
    const fileName = `${safeName}_${suffix}_${dateTag}.xlsx`;

    global.XLSX.writeFile(wb, fileName);
    return fileName;
  }

  global.ExcelExporter = { exportResultsToExcel };
})(window);
