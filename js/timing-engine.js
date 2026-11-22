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

  /**
   * 完整計算所有選手成績、分段、配速、預估完賽時間、總名次與分組名次
   * @param {Object} params
   * @param {Array} params.runners
   * @param {Array} params.scans
   * @param {Array} params.categories
   * @param {Array} params.checkpoints
   * @param {'gun'|'chip'} params.rankBy
   */
  function computeResults({ runners = [], scans = [], categories = [], checkpoints = [], rankBy = 'gun' }) {
    const sortedCheckpoints = [...checkpoints].sort((a, b) => Number(a.order) - Number(b.order));
    const catMap = new Map(categories.map((c) => [c.id, c]));

    const startCp = sortedCheckpoints.find((cp) => cp.type === 'start') || sortedCheckpoints[0] || null;
    const finishCp =
      [...sortedCheckpoints].reverse().find((cp) => cp.type === 'finish') ||
      sortedCheckpoints[sortedCheckpoints.length - 1] ||
      null;

    // 彙整有效感應紀錄：每位選手在每個站點取最早的有效感應（或被人工編輯過最新一筆）
    const activeScans = scans.filter((s) => !s.deleted);
    const scansByBib = new Map();
    for (const s of activeScans) {
      if (!scansByBib.has(s.bib)) scansByBib.set(s.bib, []);
      scansByBib.get(s.bib).push(s);
    }

    // 確保即使有「未事先建檔但已感應」的選手也能納入成績計算
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

    const rows = [];

    for (const runner of runnerMap.values()) {
      const cat = catMap.get(runner.categoryId) || null;
      const gunTimeMs = cat && cat.gunTimeMs ? Number(cat.gunTimeMs) : null;
      const penaltyMs = Number(runner.penaltySec || 0) * 1000;
      const runnerScans = scansByBib.get(runner.bib) || [];

      // 每個站點取該站點最早的感應（若有人工修正 source==='admin' 優先）
      const cpScanMap = {};
      for (const cp of sortedCheckpoints) {
        const list = runnerScans
          .filter((s) => s.checkpointId === cp.id)
          .sort((a, b) => {
            if (a.source === 'admin' && b.source !== 'admin') return -1;
            if (b.source === 'admin' && a.source !== 'admin') return 1;
            return Number(a.timestampMs) - Number(b.timestampMs);
          });
        if (list.length > 0) {
          cpScanMap[cp.id] = list[0];
        }
      }

      const startScan = startCp ? cpScanMap[startCp.id] : null;
      const finishScan = finishCp ? cpScanMap[finishCp.id] : null;

      // 找出最後通過的站點
      let lastCheckpoint = null;
      let lastScan = null;
      for (const cp of sortedCheckpoints) {
        if (cpScanMap[cp.id]) {
          lastCheckpoint = cp;
          lastScan = cpScanMap[cp.id];
        }
      }

      // 計算分段成績 (Split Times)
      const splits = [];
      let prevTimeMs = startScan ? Number(startScan.timestampMs) : gunTimeMs;
      let prevDistKm = startCp ? Number(startCp.distanceKm || 0) : 0;
      let prevName = startCp ? startCp.name : '起跑';

      for (const cp of sortedCheckpoints) {
        if (startCp && cp.id === startCp.id) continue;
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

      // 計算大會時間 (Gun Time) 與 個人淨時間 (Chip Time)
      const gunDurationMs =
        finishScan && gunTimeMs !== null
          ? Number(finishScan.timestampMs) - gunTimeMs + penaltyMs
          : null;

      const chipDurationMs =
        finishScan && startScan
          ? Number(finishScan.timestampMs) - Number(startScan.timestampMs) + penaltyMs
          : gunDurationMs;

      // 總距離與平均配速
      const totalDistanceKm = cat ? Number(cat.distanceKm || 0) : 0;
      const primaryDurationMs = rankBy === 'chip'
        ? (chipDurationMs !== null ? chipDurationMs : gunDurationMs)
        : (gunDurationMs !== null ? gunDurationMs : chipDurationMs);

      const avgPace = formatPace(primaryDurationMs, totalDistanceKm);

      // 判斷選手狀態：finished / on_course / not_started
      let status = 'not_started';
      if (finishScan) {
        status = 'finished';
      } else if (lastScan || (gunTimeMs !== null && Date.now() >= gunTimeMs && startScan)) {
        status = 'on_course';
      } else if (gunTimeMs !== null && Date.now() >= gunTimeMs) {
        // 已鳴槍且有起點感應門以外的感應，或無起點感應門時已鳴槍視為賽道上
        status = startCp ? 'not_started' : 'on_course';
      }

      // 預估完賽時間 (ETA)：在賽道上且有已知距離 > 0
      let etaTimestampMs = null;
      let etaDurationMs = null;
      if (status === 'on_course' && lastScan && lastCheckpoint && totalDistanceKm > 0) {
        const coveredKm = Number(lastCheckpoint.distanceKm || 0);
        const refStartMs = startScan ? Number(startScan.timestampMs) : gunTimeMs;
        if (coveredKm > 0 && coveredKm < totalDistanceKm && refStartMs !== null) {
          const elapsed = Number(lastScan.timestampMs) - refStartMs;
          if (elapsed > 0) {
            const estTotal = (elapsed / coveredKm) * totalDistanceKm;
            etaDurationMs = estTotal;
            etaTimestampMs = refStartMs + estTotal;
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
        cpScanMap,
        startScan,
        finishScan,
        lastCheckpoint,
        lastScan,
        splits,
        gunDurationMs,
        chipDurationMs,
        primaryDurationMs,
        avgPace,
        status,
        etaTimestampMs,
        etaDurationMs,
        overallRank: null,
        categoryRank: null
      });
    }

    // 排序與名次計算（已完賽者依 primaryDurationMs 升冪排名）
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

    // 總表預設排序：已完賽（依名次） -> 賽道上（依最後站點順序降冪、最後通過時間升冪） -> 未出發（依編號）
    rows.sort((a, b) => {
      const rankOrder = { finished: 0, on_course: 1, not_started: 2 };
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

    // 統計看板資料
    const summary = {
      totalRunners: rows.length,
      finishedCount: rows.filter((r) => r.status === 'finished').length,
      onCourseCount: rows.filter((r) => r.status === 'on_course').length,
      notStartedCount: rows.filter((r) => r.status === 'not_started').length,
      byCategory: categories.map((cat) => {
        const list = rows.filter((r) => r.categoryId === cat.id);
        const fin = list.filter((r) => r.status === 'finished');
        const onC = list.filter((r) => r.status === 'on_course');
        const times = fin.map((r) => r.primaryDurationMs).filter((t) => t !== null);
        const bestMs = times.length ? Math.min(...times) : null;
        const avgMs = times.length ? times.reduce((a, b) => a + b, 0) / times.length : null;
        return {
          ...cat,
          total: list.length,
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
      rows,
      summary,
      checkpoints: sortedCheckpoints,
      startCp,
      finishCp
    };
  }

  global.TimingEngine = {
    formatClockMs,
    formatClockSec,
    formatDuration,
    formatPace,
    parseDateTimeToMs,
    maskIdNumber,
    sha256Hex,
    computeResults
  };
})(window);
