/**
 * TUR (Taiwan Ultra Runners) 路跑晶片計時平台 v3.10 — Excel (.xlsx) / DUV 國際超馬協會成績匯出模組
 */
(function (global) {
  'use strict';

  const MONTHS_EN = [
    'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
    'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'
  ];

  const COMPOUND_SURNAMES = new Set([
    '歐陽', '司馬', '諸葛', '上官', '司徒', '皇甫', '尉遲', '公孫', '慕容',
    '長孫', '端木', '宇文', '鮮于', '軒轅', '令狐', '東方', '獨孤', '南宮',
    '范姜', '張簡', '鍾離', '申屠', '聞人', '赫連', '呼延', '百里', '西門', '東門'
  ]);

  function formatDuvDate(dateStr) {
    const m = String(dateStr || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!m) return String(dateStr || '').trim().toUpperCase();
    const y = m[1];
    const mon = MONTHS_EN[(parseInt(m[2], 10) || 1) - 1] || m[2];
    const d = parseInt(m[3], 10);
    return `${mon} ${d}, ${y}`;
  }

  function splitRunnerName(fullName, explicitLast = '', explicitFirst = '') {
    const l = String(explicitLast || '').trim();
    const f = String(explicitFirst || '').trim();
    if (l || f) return { lastName: l, firstName: f };

    const clean = String(fullName || '').trim();
    if (!clean) return { lastName: '', firstName: '' };

    // 含空白（多半為外籍姓名，如 Emmanuel Lemire）
    if (/\s+/.test(clean)) {
      const parts = clean.split(/\s+/).filter(Boolean);
      if (parts.length >= 2) {
        return {
          lastName: parts[parts.length - 1],
          firstName: parts.slice(0, -1).join(' ')
        };
      }
    }

    // 中日韓姓名：支援常見複姓（3字以上），其餘預設第 1 字為姓、其餘為名
    if (clean.length >= 3 && COMPOUND_SURNAMES.has(clean.slice(0, 2))) {
      return { lastName: clean.slice(0, 2), firstName: clean.slice(2) };
    }
    if (clean.length >= 2) {
      return { lastName: clean.slice(0, 1), firstName: clean.slice(1) };
    }
    return { lastName: clean, firstName: '' };
  }

  function resolveDuvGender(row) {
    const g = String(row.gender || row.runner?.gender || '').trim().toUpperCase();
    if (g === 'M' || g === 'MALE' || g === '男') return 'M';
    if (g === 'F' || g === 'FEMALE' || g === '女') return 'F';
    const id = String(row.idNumber || row.runner?.idNumber || '').trim().toUpperCase();
    if (/^[A-Z][1289]\d{8}$/.test(id)) {
      return (id[1] === '1' || id[1] === '8') ? 'M' : 'F';
    }
    return '';
  }

  function csvEscape(val) {
    const s = String(val ?? '');
    if (/[",\r\n]/.test(s)) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  }

  function sanitizeSheetName(name, usedSet) {
    let clean = String(name || 'Sheet').replace(/[\\/*?:[\]]/g, '_').trim() || 'Sheet';
    if (clean.length > 28) clean = clean.slice(0, 28);
    let candidate = clean;
    let idx = 2;
    while (usedSet.has(candidate)) {
      candidate = `${clean.slice(0, 25)}_${idx++}`;
    }
    usedSet.add(candidate);
    return candidate;
  }

  function buildDuvCategoryBlocks({ event, computed, options = {} }) {
    const TE = global.TimingEngine;
    const isLap = event?.eventType === 'lap_timed' || computed?.eventType === 'lap_timed';
    const organization = String(options.organization ?? event?.organization ?? 'Taiwan Ultra Runners').trim() || 'Taiwan Ultra Runners';
    const eventName = String(options.eventName ?? (event?.englishName || event?.name || '')).trim();
    const defaultCity = String(options.defaultCity ?? 'Taipei').trim();
    const defaultCountry = String(options.defaultCountry ?? 'Taiwan').trim() || 'Taiwan';
    const lapMetric = options.lapMetric === 'time' ? 'time' : 'distance';
    const duvDate = formatDuvDate(event?.date || new Date().toISOString().slice(0, 10));
    const eventNameAndDate = duvDate ? `${eventName}\n${duvDate}` : eventName;

    const categories = Array.isArray(event?.categories) ? [...event.categories] : [];
    const catIds = new Set(categories.map((c) => c.id));
    for (const r of computed?.rows || []) {
      if (r.categoryId && !catIds.has(r.categoryId)) {
        catIds.add(r.categoryId);
        categories.push({ id: r.categoryId, name: r.categoryName || r.categoryId, distanceKm: r.distanceKm || 0, durationHours: r.durationHours || 0 });
      }
    }

    const perfHeader = (isLap && lapMetric === 'distance') ? 'Distance (km)' : 'Total Time';
    const blocks = [];
    let totalFinishers = 0;
    let missingDobCount = 0;
    let missingGenderCount = 0;

    for (const cat of categories) {
      let catRows = (computed?.rows || []).filter((r) => r.categoryId === cat.id);
      if (isLap) {
        catRows = catRows
          .filter((r) => Number(r.lapCount || 0) > 0)
          .sort((a, b) => {
            if (b.totalDistanceKm !== a.totalDistanceKm) return b.totalDistanceKm - a.totalDistanceKm;
            return (a.totalElapsedMs ?? Number.MAX_SAFE_INTEGER) - (b.totalElapsedMs ?? Number.MAX_SAFE_INTEGER);
          });
      } else {
        catRows = catRows
          .filter((r) => r.status === 'finished' && r.gunDurationMs !== null)
          .sort((a, b) => (a.gunDurationMs ?? Number.MAX_SAFE_INTEGER) - (b.gunDurationMs ?? Number.MAX_SAFE_INTEGER));
      }

      if (!catRows.length) continue;

      const catTitle = isLap
        ? (cat.durationHours ? `Category ${cat.durationHours}H` : `Category ${cat.name || cat.id}`)
        : (Number(cat.distanceKm || 0) > 0 ? `Category ${Number(cat.distanceKm)} KM` : `Category ${cat.name || cat.id}`);

      const sheetLabel = isLap
        ? (cat.durationHours ? `${cat.durationHours}H` : (cat.id || cat.name))
        : (Number(cat.distanceKm || 0) > 0 ? `${Number(cat.distanceKm)} KM` : (cat.id || cat.name));

      const aoa = [
        ['Event organization', '', organization, '', '', '', '', '', '', '', ''],
        ['Event name and date', '', eventNameAndDate, '', '', '', '', '', '', '', ''],
        [catTitle, '', '', '', '', '', '', '', '', '', ''],
        ['Sr. No.', 'Last Name', '1st Name', 'Gender', 'BIB No.', 'DOB', 'City', 'Country', perfHeader, 'Rank', 'Remarks']
      ];

      catRows.forEach((row, idx) => {
        const srNo = idx + 1;
        const { lastName, firstName } = splitRunnerName(
          row.name,
          row.lastName || row.runner?.lastName,
          row.firstName || row.runner?.firstName
        );
        const gender = resolveDuvGender(row);
        const dob = String(row.dob || row.runner?.dob || '').trim();
        const city = String(row.city || row.runner?.city || defaultCity).trim();
        const country = String(row.country || row.runner?.country || defaultCountry).trim();

        if (!dob) missingDobCount++;
        if (!gender) missingGenderCount++;

        let perfVal = '';
        let remarks = '';
        if (isLap) {
          if (lapMetric === 'distance') {
            perfVal = Number(row.totalDistanceKm || 0).toFixed(3);
            const tStr = row.totalElapsedMs !== null ? TE.formatDuration(row.totalElapsedMs, false) : '';
            remarks = `${row.lapCount} laps${tStr ? ` (${tStr})` : ''}`;
          } else {
            perfVal = row.totalElapsedMs !== null ? TE.formatDuration(row.totalElapsedMs, false) : '';
            remarks = `${Number(row.totalDistanceKm || 0).toFixed(3)} km (${row.lapCount} laps)`;
          }
        } else {
          perfVal = row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, false) : '';
          remarks = '';
        }

        aoa.push([
          srNo,
          lastName,
          firstName,
          gender,
          String(row.bib || '').trim(),
          dob,
          city,
          country,
          perfVal,
          srNo,
          remarks
        ]);
      });

      totalFinishers += catRows.length;
      blocks.push({
        categoryId: cat.id,
        categoryTitle: catTitle,
        sheetLabel,
        count: catRows.length,
        aoa
      });
    }

    return {
      isLap,
      organization,
      eventName,
      duvDate,
      perfHeader,
      blocks,
      totalFinishers,
      missingDobCount,
      missingGenderCount
    };
  }

  function exportDuvResults({ event, computed, options = {} }) {
    const format = options.format || 'xlsx_stacked'; // 'xlsx_stacked' | 'xlsx_sheets' | 'csv'
    const built = buildDuvCategoryBlocks({ event, computed, options });
    if (!built.blocks.length || built.totalFinishers === 0) {
      throw new Error('本活動目前尚無已完賽（或有效圈數）之選手成績可供匯出 DUV 報表');
    }

    const safeName = (built.eventName || event?.name || 'TUR_Event').replace(/[\\/:*?"<>|]/g, '_');
    const dateTag = event?.date || new Date().toISOString().slice(0, 10);

    if (format === 'csv') {
      const allRows = [];
      built.blocks.forEach((b) => {
        b.aoa.forEach((row) => allRows.push(row.map(csvEscape).join(',')));
      });
      const csvContent = '\uFEFF' + allRows.join('\r\n');
      const fileName = `${safeName}_DUV_${dateTag}.csv`;
      if (typeof global.Blob !== 'undefined' && typeof global.document !== 'undefined') {
        const blob = new global.Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = global.URL.createObjectURL(blob);
        const a = global.document.createElement('a');
        a.href = url;
        a.download = fileName;
        global.document.body.appendChild(a);
        a.click();
        a.remove();
        global.URL.revokeObjectURL(url);
      }
      return { fileName, ...built, csvContent };
    }

    if (typeof global.XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) 套件尚未載入，請確認網路連線或重新整理頁面');
    }

    const wb = global.XLSX.utils.book_new();
    const colWidths = [
      { wch: 8 },  // Sr. No.
      { wch: 12 }, // Last Name
      { wch: 14 }, // 1st Name
      { wch: 8 },  // Gender
      { wch: 10 }, // BIB No.
      { wch: 13 }, // DOB
      { wch: 14 }, // City
      { wch: 12 }, // Country
      { wch: 15 }, // Total Time / Distance (km)
      { wch: 8 },  // Rank
      { wch: 20 }  // Remarks
    ];

    if (format === 'xlsx_sheets') {
      const usedNames = new Set();
      built.blocks.forEach((b) => {
        const ws = global.XLSX.utils.aoa_to_sheet(b.aoa);
        ws['!cols'] = colWidths;
        global.XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(b.sheetLabel, usedNames));
      });
      const fileName = `${safeName}_DUV_分組工作表_${dateTag}.xlsx`;
      global.XLSX.writeFile(wb, fileName);
      return { fileName, ...built };
    }

    // 預設：xlsx_stacked（全部組別依範例堆疊於同一工作表）
    const stackedAoa = [];
    built.blocks.forEach((b) => {
      b.aoa.forEach((row) => stackedAoa.push(row));
    });
    const ws = global.XLSX.utils.aoa_to_sheet(stackedAoa);
    ws['!cols'] = colWidths;
    global.XLSX.utils.book_append_sheet(wb, ws, 'DUV_Results');
    const fileName = `${safeName}_DUV_${dateTag}.xlsx`;
    global.XLSX.writeFile(wb, fileName);
    return { fileName, ...built };
  }

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

  function downloadHistoryTemplate() {
    if (typeof global.XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) 套件尚未載入，請確認網路連線或重新整理頁面');
    }
    const wb = global.XLSX.utils.book_new();
    const sampleRows = [
      {
        活動日期: '2025-11-16',
        活動名稱: '2025 TUR 八百壯士超級馬拉松',
        組別: '100K',
        參賽編號: 'W001',
        姓名: '林志明',
        身份證字號: 'A123456789',
        完賽時間: '09:38:15',
        完賽里程km: 100,
        圈數: 0,
        總名次: 1,
        組名次: 1,
        完賽狀態: '完賽',
        備註: ''
      },
      {
        活動日期: '2024-11-17',
        活動名稱: '2024 TUR 八百壯士經典挑戰賽',
        組別: '24H',
        參賽編號: 'H002',
        姓名: '張家豪',
        身份證字號: 'C145678912',
        完賽時間: '23:42:10',
        完賽里程km: 162.56,
        圈數: 64,
        總名次: 1,
        組名次: 1,
        完賽狀態: '完賽',
        備註: '單圈2.54km'
      },
      {
        活動日期: '2025-05-10',
        活動名稱: '2025 TUR 陽明山越野超馬',
        組別: '50K',
        參賽編號: 'Y010',
        姓名: '陳怡君',
        身份證字號: 'B234567891',
        完賽時間: '05:22:30',
        完賽里程km: 50,
        圈數: 0,
        總名次: 2,
        組名次: 1,
        完賽狀態: '完賽',
        備註: ''
      }
    ];
    const ws = global.XLSX.utils.json_to_sheet(sampleRows);
    ws['!cols'] = [
      { wch: 13 }, { wch: 30 }, { wch: 10 }, { wch: 11 }, { wch: 12 },
      { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 8 },
      { wch: 8 }, { wch: 10 }, { wch: 18 }
    ];
    global.XLSX.utils.book_append_sheet(wb, ws, '歷年活動成績匯入範本');
    const fileName = 'TUR_歷年活動成績匯入範本.xlsx';
    global.XLSX.writeFile(wb, fileName);
    return fileName;
  }

  global.ExcelExporter = {
    exportResultsToExcel,
    downloadHistoryTemplate,
    buildDuvCategoryBlocks,
    exportDuvResults,
    splitRunnerName,
    formatDuvDate
  };
})(window);
