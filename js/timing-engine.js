/**
 * 智慧路跑 NFC 晶片計時系統 — 成績計算、分段配速、名次排序與隱碼核心引擎
 */
(function (global) {
  'use strict';

  function pad2(n) {
    return String(Math.floor(n)).padStart(2, '0');
  }

  function pad3(n) {
    return String(Math.floor(n)).padStart(3, '0');
  }

  /**
   * 格式化系統時間為 HH:mm:ss.SSS
   */
  function formatClockMs(tsMs) {
    if (tsMs === null || tsMs === undefined || Number.isNaN(Number(tsMs))) return '--:--:--.---';
    const d = new Date(Number(tsMs));
    return `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}.${pad3(d.getMilliseconds())}`;
  }

  /**
   * 格式化系統時間為 HH:mm:ss
   */
  function formatClockSec(tsMs) {
    if (tsMs === null || tsMs === undefined || Number.isNaN(Number(tsMs))) return '--:--:--';
    const d = new Date(Number(tsMs));
    return `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  }

  /**
   * 格式化時間長度 (毫秒) 為 HH:mm:ss.SSS 或 HH:mm:ss
   */
  function formatDuration(durationMs, includeMs = true) {
    if (durationMs === null || durationMs === undefined || Number.isNaN(Number(durationMs))) {
      return includeMs ? '--:--:--.---' : '--:--:--';
    }
    const neg = durationMs < 0;
    const abs = Math.abs(Math.round(durationMs));
    const hours = Math.floor(abs / 3600000);
    const minutes = Math.floor((abs % 3600000) / 60000);
    const seconds = Math.floor((abs % 60000) / 1000);
    const ms = abs % 1000;
    const prefix = neg ? '-' : '';
    return includeMs
      ? `${prefix}${pad2(hours)}:${pad2(minutes)}:${pad2(seconds)}.${pad3(ms)}`
      : `${prefix}${pad2(hours)}:${pad2(minutes)}:${pad2(seconds)}`;
  }

  /**
   * 計算配速 (每公里 分'秒"/km)
   */
  function formatPace(durationMs, distanceKm) {
    if (!durationMs || durationMs <= 0 || !distanceKm || distanceKm <= 0) return '--';
    const secPerKm = (durationMs / 1000) / distanceKm;
    if (!Number.isFinite(secPerKm) || secPerKm > 3600) return '--';
    const min = Math.floor(secPerKm / 60);
    const sec = Math.round(secPerKm % 60);
    if (sec === 60) return `${min + 1}'00"/km`;
    return `${min}'${pad2(sec)}"/km`;
  }

  /**
   * 解析日期 (YYYY-MM-DD) 與時間字串 (HH:mm:ss 或 HH:mm:ss.SSS) 為毫秒時間戳
   */
  function parseDateTimeToMs(dateStr, timeStr) {
    if (!timeStr) return null;
    const cleanTime = String(timeStr).trim();
    const match = cleanTime.match(/^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?$/);
    if (!match) return null;

    const baseDate = dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)
      ? dateStr.split('-').map(Number)
      : (() => {
          const now = new Date();
          return [now.getFullYear(), now.getMonth() + 1, now.getDate()];
        })();

    const h = Number(match[1]);
    const m = Number(match[2]);
    const s = Number(match[3] || 0);
    const ms = Number((match[4] || '0').padEnd(3, '0'));
    if (h > 23 || m > 59 || s > 59) return null;

    const d = new Date(baseDate[0], baseDate[1] - 1, baseDate[2], h, m, s, ms);
    return d.getTime();
  }

  /**
   * 身份證字號隱碼 (例如 A123456789 -> A123****89)
   */
  function maskIdNumber(idStr) {
    const s = String(idStr || '').trim().toUpperCase();
    if (!s) return '';
    if (s.length <= 4) return s.slice(0, 1) + '***';
    return s.slice(0, 4) + '****' + s.slice(-2);
  }

  /**
   * SHA-256 雜湊（公開查詢頁安全比對身份證字號用，含純 JS 降級實作）
   */
  async function sha256Hex(text) {
    const normalized = String(text || '').trim().toUpperCase();
    if (!normalized) return '';
    if (global.crypto && global.crypto.subtle && global.TextEncoder) {
      const buf = await global.crypto.subtle.digest(
        'SHA-256',
        new TextEncoder().encode(normalized)
      );
      return Array.from(new Uint8Array(buf))
        .map((b) => b.toString(16).padStart(2, '0'))
        .join('');
    }
    return sha256Fallback_(normalized);
  }

  // 純 JS SHA-256 fallback (無 HTTPS 環境也能正常運作)
  function sha256Fallback_(ascii) {
    function rightRotate(value, amount) {
      return (value >>> amount) | (value << (32 - amount));
    }
    const maxWord = Math.pow(2, 32);
    let result = '';
    const words = [];
    const asciiBitLength = ascii.length * 8;
    const hash = [];
    const k = [];
    let primeCounter = 0;
    const isComposite = {};

    for (let candidate = 2; primeCounter < 64; candidate++) {
      if (!isComposite[candidate]) {
        for (let i = 0; i < 313; i += candidate) isComposite[i] = candidate;
        hash[primeCounter] = (Math.pow(candidate, 0.5) * maxWord) | 0;
        k[primeCounter++] = (Math.pow(candidate, 1 / 3) * maxWord) | 0;
      }
    }

    ascii += '\x80';
    while ((ascii.length % 64) - 56) ascii += '\x00';
    for (let i = 0; i < ascii.length; i++) {
      const j = ascii.charCodeAt(i);
      words[i >> 2] |= j << (((3 - i) % 4) * 8);
    }
    words[words.length] = (asciiBitLength / maxWord) | 0;
    words[words.length] = asciiBitLength;

    for (let j = 0; j < words.length;) {
      const w = words.slice(j, (j += 16));
      const oldHash = hash.slice(0);
      for (let i = 0; i < 64; i++) {
        const w15 = w[i - 15], w2 = w[i - 2];
        const a = hash[0], e = hash[4];
        const temp1 =
          hash[7] +
          (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25)) +
          ((e & hash[5]) ^ (~e & hash[6])) +
          k[i] +
          (w[i] =
            i < 16
              ? w[i]
              : (w[i - 16] +
                  (rightRotate(w15, 7) ^ rightRotate(w15, 18) ^ (w15 >>> 3)) +
                  w[i - 7] +
                  (rightRotate(w2, 17) ^ rightRotate(w2, 19) ^ (w2 >>> 10))) |
                0);
        const temp2 =
          (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22)) +
          ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
        hash.unshift((temp1 + temp2) | 0);
        hash[4] = (hash[4] + temp1) | 0;
        hash.pop();
      }
      for (let i = 0; i < 8; i++) hash[i] = (hash[i] + oldHash[i]) | 0;
    }

    for (let i = 0; i < 8; i++) {
      for (let j = 3; j + 1; j--) {
        const b = (hash[i] >> (j * 8)) & 255;
        result += (b < 16 ? '0' : '') + b.toString(16);
      }
    }
    return result;
  }

  const CHECKIN_CP_ID = 'CHECKIN';

  function isCheckInScan(scan) {
    return String(scan?.checkpointId || '').trim().toUpperCase() === CHECKIN_CP_ID;
  }

  function pickPreferredScan(list) {
    if (!list || !list.length) return null;
    return list
      .slice()
      .sort((a, b) => {
        if (a.source === 'admin' && b.source !== 'admin') return -1;
        if (b.source === 'admin' && a.source !== 'admin') return 1;
        return Number(a.timestampMs) - Number(b.timestampMs);
      })[0];
  }

  /**
   * 完整計算所有選手報到狀態、成績、分段、配速、預估完賽時間、總名次與分組名次
   * 本系統統一採「律定起跑時間」計時，不設起跑感應亦不另計個人淨時間。
   */
  function computeResults({ runners = [], scans = [], categories = [], checkpoints = [], unifiedGunTimeMs = null }) {
    // 排除報到點 (CHECKIN) 及舊版遺留之起點站 (START / type==='start')，僅保留賽道站點與終點
    const sortedCheckpoints = [...checkpoints]
      .filter((cp) => {
        const id = String(cp?.id || '').trim().toUpperCase();
        return cp && cp.type !== 'start' && cp.type !== 'checkin' && id !== 'START' && id !== CHECKIN_CP_ID;
      })
      .sort((a, b) => Number(a.order) - Number(b.order));

    const catMap = new Map(categories.map((c) => [c.id, c]));
    const finishCp =
      [...sortedCheckpoints].reverse().find((cp) => cp.type === 'finish') ||
      sortedCheckpoints[sortedCheckpoints.length - 1] ||
      null;

    // 彙整有效感應紀錄
    const activeScans = scans.filter((s) => !s.deleted);
    const scansByBib = new Map();
    for (const s of activeScans) {
      if (!scansByBib.has(s.bib)) scansByBib.set(s.bib, []);
      scansByBib.get(s.bib).push(s);
    }

    // 確保即使有「未事先建檔但已感應」的選手也能納入計算
    const runnerMap = new Map();
    for (const r of runners) {
      runnerMap.set(r.bib, r);
    }
    for (const bib of scansByBib.keys()) {
      if (!runnerMap.has(bib)) {
        runnerMap.set(bib, {
          bib,
          name: `未登錄選手 (${bib})`,
          idNumber: '',
          idMasked: '',
          categoryId: categories[0]?.id || '',
          penaltySec: 0
        });
      }
    }

    const now = Date.now();
    const rows = [];

    for (const runner of runnerMap.values()) {
      const cat = catMap.get(runner.categoryId) || null;
      const gunTimeMs = cat && cat.gunTimeMs
        ? Number(cat.gunTimeMs)
        : (unifiedGunTimeMs ? Number(unifiedGunTimeMs) : null);
      const penaltyMs = Number(runner.penaltySec || 0) * 1000;
      const runnerScans = scansByBib.get(runner.bib) || [];

      // 報到感應紀錄 (CHECKIN)
      const checkInScan = pickPreferredScan(runnerScans.filter((s) => isCheckInScan(s)));
      const checkedIn = Boolean(checkInScan);

      // 各賽道站點感應紀錄
      const cpScanMap = {};
      for (const cp of sortedCheckpoints) {
        const picked = pickPreferredScan(runnerScans.filter((s) => s.checkpointId === cp.id));
        if (picked) cpScanMap[cp.id] = picked;
      }

      const finishScan = finishCp ? cpScanMap[finishCp.id] : null;

      // 找出最後通過的賽道站點
      let lastCheckpoint = null;
      let lastScan = null;
      for (const cp of sortedCheckpoints) {
        if (cpScanMap[cp.id]) {
          lastCheckpoint = cp;
          lastScan = cpScanMap[cp.id];
        }
      }

      // 計算分段成績 (一律以統一律定起跑時間 gunTimeMs 為起算基準)
      const splits = [];
      let prevTimeMs = gunTimeMs;
      let prevDistKm = 0;
      let prevName = '統一開跑';

      for (const cp of sortedCheckpoints) {
        const scan = cpScanMap[cp.id];
        const targetDist = cp.type === 'finish' && cat && cat.distanceKm
          ? Number(cat.distanceKm)
          : Number(cp.distanceKm || 0);
        const segDistKm = Math.max(0, targetDist - prevDistKm);

        if (scan && prevTimeMs !== null) {
          const durationMs = Number(scan.timestampMs) - prevTimeMs;
          splits.push({
            fromName: prevName,
            toId: cp.id,
            toName: cp.name,
            durationMs,
            distanceKm: segDistKm,
            pace: formatPace(durationMs, segDistKm)
          });
          prevTimeMs = Number(scan.timestampMs);
          prevDistKm = targetDist;
          prevName = cp.name;
        } else {
          splits.push({
            fromName: prevName,
            toId: cp.id,
            toName: cp.name,
            durationMs: null,
            distanceKm: segDistKm,
            pace: '--'
          });
        }
      }

      // 計算統一成績 (完賽時間 = 終點感應時間 - 律定起跑時間 + 罰時)
      const gunDurationMs =
        finishScan && gunTimeMs !== null
          ? Number(finishScan.timestampMs) - gunTimeMs + penaltyMs
          : null;

      const totalDistanceKm = cat ? Number(cat.distanceKm || 0) : 0;
      const primaryDurationMs = gunDurationMs;
      const avgPace = formatPace(primaryDurationMs, totalDistanceKm);

      // 判斷選手狀態：finished(已完賽) / on_course(賽道上) / checked_in(已報到待起跑) / not_checked_in(未報到)
      let status = 'not_checked_in';
      if (finishScan) {
        status = 'finished';
      } else if (lastScan) {
        status = 'on_course';
      } else if (checkedIn) {
        status = (gunTimeMs !== null && now >= gunTimeMs) ? 'on_course' : 'checked_in';
      }

      // 預估完賽時間 (ETA)：在賽道上且有已知距離 > 0
      let etaTimestampMs = null;
      let etaDurationMs = null;
      if (status === 'on_course' && lastScan && lastCheckpoint && totalDistanceKm > 0 && gunTimeMs !== null) {
        const coveredKm = Number(lastCheckpoint.distanceKm || 0);
        if (coveredKm > 0 && coveredKm < totalDistanceKm) {
          const elapsed = Number(lastScan.timestampMs) - gunTimeMs;
          if (elapsed > 0) {
            const estTotal = (elapsed / coveredKm) * totalDistanceKm;
            etaDurationMs = estTotal;
            etaTimestampMs = gunTimeMs + estTotal;
          }
        }
      }

      rows.push({
        runner,
        bib: runner.bib,
        name: runner.name,
        idNumber: runner.idNumber || '',
        idMasked: runner.idMasked || maskIdNumber(runner.idNumber),
        idHash: runner.idHash || '',
        categoryId: runner.categoryId,
        categoryName: cat ? cat.name : runner.categoryId || '未分組',
        categoryColor: cat ? cat.color : '#64748b',
        distanceKm: totalDistanceKm,
        gunTimeMs,
        penaltySec: Number(runner.penaltySec || 0),
        checkedIn,
        checkInScan,
        checkInTimeMs: checkInScan ? Number(checkInScan.timestampMs) : null,
        cpScanMap,
        finishScan,
        lastCheckpoint,
        lastScan,
        splits,
        gunDurationMs,
        primaryDurationMs,
        avgPace,
        status,
        etaTimestampMs,
        etaDurationMs,
        overallRank: null,
        categoryRank: null
      });
    }

    // 排序與名次計算（已完賽者依完賽時間升冪排名）
    const finishedRows = rows
      .filter((r) => r.status === 'finished' && r.primaryDurationMs !== null)
      .sort((a, b) => a.primaryDurationMs - b.primaryDurationMs);

    finishedRows.forEach((r, idx) => {
      r.overallRank = idx + 1;
    });

    // 分組排名
    const catCounter = {};
    finishedRows.forEach((r) => {
      const cid = r.categoryId || '__NONE__';
      catCounter[cid] = (catCounter[cid] || 0) + 1;
      r.categoryRank = catCounter[cid];
    });

    // 總表預設排序：已完賽 -> 賽道上 -> 已報到待起跑 -> 未報到
    rows.sort((a, b) => {
      const rankOrder = { finished: 0, on_course: 1, checked_in: 2, not_checked_in: 3 };
      if (rankOrder[a.status] !== rankOrder[b.status]) {
        return rankOrder[a.status] - rankOrder[b.status];
      }
      if (a.status === 'finished') {
        return (a.overallRank || 999999) - (b.overallRank || 999999);
      }
      if (a.status === 'on_course') {
        const orderA = a.lastCheckpoint ? a.lastCheckpoint.order : -1;
        const orderB = b.lastCheckpoint ? b.lastCheckpoint.order : -1;
        if (orderA !== orderB) return orderB - orderA;
        return (a.lastScan?.timestampMs || 0) - (b.lastScan?.timestampMs || 0);
      }
      return String(a.bib).localeCompare(String(b.bib), 'zh-Hant', { numeric: true });
    });

    // 統計看板資料（含報到掌握統計）
    const checkedInCount = rows.filter((r) => r.checkedIn || r.status === 'finished' || Boolean(r.lastScan)).length;
    const summary = {
      totalRunners: rows.length,
      checkedInCount,
      notCheckedInCount: rows.length - checkedInCount,
      finishedCount: rows.filter((r) => r.status === 'finished').length,
      onCourseCount: rows.filter((r) => r.status === 'on_course').length,
      checkedInWaitingCount: rows.filter((r) => r.status === 'checked_in').length,
      notStartedCount: rows.filter((r) => r.status === 'checked_in' || r.status === 'not_checked_in').length,
      byCategory: categories.map((cat) => {
        const list = rows.filter((r) => r.categoryId === cat.id);
        const fin = list.filter((r) => r.status === 'finished');
        const onC = list.filter((r) => r.status === 'on_course');
        const chk = list.filter((r) => r.checkedIn || r.status === 'finished' || Boolean(r.lastScan));
        const times = fin.map((r) => r.primaryDurationMs).filter((t) => t !== null);
        const bestMs = times.length ? Math.min(...times) : null;
        const avgMs = times.length ? times.reduce((a, b) => a + b, 0) / times.length : null;
        return {
          ...cat,
          total: list.length,
          checkedIn: chk.length,
          notCheckedIn: list.length - chk.length,
          finished: fin.length,
          onCourse: onC.length,
          notStarted: list.length - fin.length - onC.length,
          bestMs,
          avgMs
        };
      }),
      byCheckpoint: sortedCheckpoints.map((cp) => {
        const passed = rows.filter((r) => Boolean(r.cpScanMap[cp.id])).length;
        return {
          ...cp,
          passedCount: passed
        };
      })
    };

    return {
      eventType: 'distance',
      rows,
      summary,
      checkpoints: sortedCheckpoints,
      finishCp
    };
  }

  /**
   * 繞圈計時賽專用運算引擎 (6H / 12H / 24H 或自訂時數)
   * 規則：
   * - 報到感應 (CHECKIN) 僅記錄到場狀態，不計入繞圈次數
   * - 賽道感應 (LAP) 每次計 1 圈，累積里程 = 有效圈數 × 單圈公里數 (lapDistanceKm)
   * - 統一依律定起跑時間計算截止時間與累積耗時
   */
  function computeLapTimedResults({
    runners = [],
    scans = [],
    categories = [],
    checkpoints = [],
    lapDistanceKm = 2.54,
    unifiedGunTimeMs = null
  }) {
    const lapKm = Math.max(0.001, Number(lapDistanceKm || 2.54));
    const catMap = new Map(categories.map((c) => [c.id, c]));
    const activeScans = scans.filter((s) => !s.deleted);

    const scansByBib = new Map();
    for (const s of activeScans) {
      if (!scansByBib.has(s.bib)) scansByBib.set(s.bib, []);
      scansByBib.get(s.bib).push(s);
    }

    const runnerMap = new Map();
    for (const r of runners) runnerMap.set(r.bib, r);
    for (const bib of scansByBib.keys()) {
      if (!runnerMap.has(bib)) {
        runnerMap.set(bib, {
          bib,
          name: `未登錄選手 (${bib})`,
          idNumber: '',
          idMasked: '',
          categoryId: categories[0]?.id || '',
          penaltySec: 0
        });
      }
    }

    const now = Date.now();
    const rows = [];

    for (const runner of runnerMap.values()) {
      const cat = catMap.get(runner.categoryId) || null;
      const gunTimeMs = cat && cat.gunTimeMs
        ? Number(cat.gunTimeMs)
        : (unifiedGunTimeMs ? Number(unifiedGunTimeMs) : null);
      const durationHours = cat ? Number(cat.durationHours || cat.distanceKm || 6) : 6;
      const cutoffTimeMs = gunTimeMs !== null && durationHours > 0
        ? gunTimeMs + durationHours * 3600 * 1000
        : null;

      const allRunnerScans = scansByBib.get(runner.bib) || [];
      const checkInScan = pickPreferredScan(allRunnerScans.filter((s) => isCheckInScan(s)));
      const checkedIn = Boolean(checkInScan);

      // 報到感應不計入繞圈次數
      const sortedScans = allRunnerScans
        .filter((s) => !isCheckInScan(s))
        .slice()
        .sort((a, b) => Number(a.timestampMs) - Number(b.timestampMs));

      const laps = [];
      let prevTs = gunTimeMs;
      let validLapNum = 0;

      for (const s of sortedScans) {
        const ts = Number(s.timestampMs);
        const overtime = cutoffTimeMs !== null && ts > cutoffTimeMs;
        if (!overtime) validLapNum++;
        const lapDurationMs = prevTs !== null && ts >= prevTs ? ts - prevTs : null;
        laps.push({
          lapNumber: overtime ? null : validLapNum,
          scanId: s.id,
          timestampMs: ts,
          lapDurationMs,
          lapPace: formatPace(lapDurationMs, lapKm),
          cumulativeKm: overtime ? null : Number((validLapNum * lapKm).toFixed(3)),
          overtime,
          source: s.source,
          note: s.note || ''
        });
        if (!overtime) prevTs = ts;
      }

      const validLaps = laps.filter((l) => !l.overtime);
      const lapCount = validLaps.length;
      const totalDistanceKm = Number((lapCount * lapKm).toFixed(3));
      const lastValidLap = lapCount > 0 ? validLaps[lapCount - 1] : null;
      const lastLapTimeMs = lastValidLap ? lastValidLap.timestampMs : null;

      const validDurations = validLaps.map((l) => l.lapDurationMs).filter((d) => d !== null && d > 0);
      const bestLapMs = validDurations.length ? Math.min(...validDurations) : null;
      const avgLapMs = validDurations.length
        ? validDurations.reduce((a, b) => a + b, 0) / validDurations.length
        : null;

      const totalElapsedMs =
        lastLapTimeMs !== null && gunTimeMs !== null
          ? lastLapTimeMs - gunTimeMs
          : (validDurations.length ? validDurations.reduce((a, b) => a + b, 0) : null);

      const avgPace = formatPace(totalElapsedMs, totalDistanceKm);

      let status = 'not_checked_in';
      if (cutoffTimeMs !== null && now >= cutoffTimeMs) {
        status = lapCount > 0 ? 'finished' : (checkedIn ? 'checked_in' : 'not_checked_in');
      } else if (lapCount > 0) {
        status = 'on_course';
      } else if (checkedIn) {
        status = (gunTimeMs !== null && now >= gunTimeMs) ? 'on_course' : 'checked_in';
      }

      rows.push({
        runner,
        bib: runner.bib,
        name: runner.name,
        idNumber: runner.idNumber || '',
        idMasked: runner.idMasked || maskIdNumber(runner.idNumber),
        idHash: runner.idHash || '',
        categoryId: runner.categoryId,
        categoryName: cat ? cat.name : runner.categoryId || '未分組',
        categoryColor: cat ? cat.color : '#64748b',
        durationHours,
        gunTimeMs,
        cutoffTimeMs,
        lapDistanceKm: lapKm,
        checkedIn,
        checkInScan,
        checkInTimeMs: checkInScan ? Number(checkInScan.timestampMs) : null,
        laps,
        lapCount,
        overtimeCount: laps.length - lapCount,
        totalDistanceKm,
        lastLapTimeMs,
        totalElapsedMs,
        bestLapMs,
        avgLapMs,
        avgPace,
        status,
        overallRank: null,
        categoryRank: null
      });
    }

    // 排名：依總里程（圈數）降冪；同里程則先達成者（累積耗時小 / 最後一圈時間早）名次在前
    const ranked = rows
      .filter((r) => r.lapCount > 0)
      .sort((a, b) => {
        if (b.totalDistanceKm !== a.totalDistanceKm) return b.totalDistanceKm - a.totalDistanceKm;
        const tA = a.totalElapsedMs !== null ? a.totalElapsedMs : (a.lastLapTimeMs || Infinity);
        const tB = b.totalElapsedMs !== null ? b.totalElapsedMs : (b.lastLapTimeMs || Infinity);
        return tA - tB;
      });

    ranked.forEach((r, idx) => { r.overallRank = idx + 1; });

    const catCounter = {};
    ranked.forEach((r) => {
      const cid = r.categoryId || '__NONE__';
      catCounter[cid] = (catCounter[cid] || 0) + 1;
      r.categoryRank = catCounter[cid];
    });

    rows.sort((a, b) => {
      if (a.overallRank !== null && b.overallRank !== null) return a.overallRank - b.overallRank;
      if (a.overallRank !== null) return -1;
      if (b.overallRank !== null) return 1;
      if (a.checkedIn !== b.checkedIn) return a.checkedIn ? -1 : 1;
      return String(a.bib).localeCompare(String(b.bib), 'zh-Hant', { numeric: true });
    });

    const totalLapsAll = rows.reduce((sum, r) => sum + r.lapCount, 0);
    const totalKmAll = Number(rows.reduce((sum, r) => sum + r.totalDistanceKm, 0).toFixed(2));
    const topDistanceKm = rows.length ? Math.max(...rows.map((r) => r.totalDistanceKm)) : 0;
    const checkedInCount = rows.filter((r) => r.checkedIn || r.lapCount > 0).length;

    const summary = {
      totalRunners: rows.length,
      checkedInCount,
      notCheckedInCount: rows.length - checkedInCount,
      activeWithLaps: rows.filter((r) => r.lapCount > 0).length,
      finishedCount: rows.filter((r) => r.status === 'finished').length,
      onCourseCount: rows.filter((r) => r.status === 'on_course').length,
      notStartedCount: rows.filter((r) => r.status === 'checked_in' || r.status === 'not_checked_in').length,
      totalLapsAll,
      totalKmAll,
      topDistanceKm,
      lapDistanceKm: lapKm,
      byCategory: categories.map((cat) => {
        const list = rows.filter((r) => r.categoryId === cat.id);
        const gunMs = cat.gunTimeMs ? Number(cat.gunTimeMs) : (unifiedGunTimeMs ? Number(unifiedGunTimeMs) : null);
        const hrs = Number(cat.durationHours || cat.distanceKm || 6);
        const cutMs = gunMs !== null ? gunMs + hrs * 3600 * 1000 : null;
        const remainingMs = cutMs !== null ? Math.max(0, cutMs - now) : null;
        const kms = list.map((r) => r.totalDistanceKm);
        const topKm = kms.length ? Math.max(...kms) : 0;
        const sumLaps = list.reduce((a, b) => a + b.lapCount, 0);
        const avgKm = list.length ? Number((kms.reduce((a, b) => a + b, 0) / list.length).toFixed(2)) : 0;
        const chkCount = list.filter((r) => r.checkedIn || r.lapCount > 0).length;
        return {
          ...cat,
          durationHours: hrs,
          cutoffTimeMs: cutMs,
          remainingMs,
          total: list.length,
          checkedIn: chkCount,
          notCheckedIn: list.length - chkCount,
          withLaps: list.filter((r) => r.lapCount > 0).length,
          sumLaps,
          topKm,
          avgKm
        };
      }),
      topLeaders: rows.filter((r) => r.lapCount > 0).slice(0, 6)
    };

    return {
      eventType: 'lap_timed',
      lapDistanceKm: lapKm,
      rows,
      summary,
      checkpoints
    };
  }

  /**
   * 統一入口：依活動性質自動分派一般距離賽或繞圈計時賽引擎
   */
  function computeEventResults({ event, runners = [], scans = [] }) {
    const unifiedGunTimeMs = event?.unifiedGunTimeMs ? Number(event.unifiedGunTimeMs) : null;
    if (event && event.eventType === 'lap_timed') {
      return computeLapTimedResults({
        runners,
        scans,
        categories: event.categories || [],
        checkpoints: event.checkpoints || [],
        lapDistanceKm: Number(event.lapDistanceKm || 2.54),
        unifiedGunTimeMs
      });
    }
    return computeResults({
      runners,
      scans,
      categories: event?.categories || [],
      checkpoints: event?.checkpoints || [],
      unifiedGunTimeMs
    });
  }

  global.TimingEngine = {
    CHECKIN_CP_ID,
    formatClockMs,
    formatClockSec,
    formatDuration,
    formatPace,
    parseDateTimeToMs,
    maskIdNumber,
    sha256Hex,
    computeResults,
    computeLapTimedResults,
    computeEventResults
  };
})(window);
