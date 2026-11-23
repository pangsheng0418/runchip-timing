/**
 * TUR (Taiwan Ultra Runners) 路跑晶片計時平台 v3.11 — 官方完賽證書繪製、自訂底圖套印與離線 PDF/PNG 匯出模組
 * 支援兩種模式：
 * 1. 官方公版（中英文對照）：高質感燙金邊框、TUR 官方 LOGO、台灣山稜浮水印、官方金色完賽鋼印、活動日期中英對照
 * 2. 自訂底圖模式：承辦人上傳自訂底圖（PNG/JPG），自由開關、拖曳或輸入座標微調各項成績欄位並輸出
 * 輸出格式：高解析度 PNG 圖檔、單張 PDF 檔、多頁合併批次 PDF 檔（100% 純 JS 離線生成）
 */
(function (global) {
  'use strict';

  const MONTHS_EN = [
    'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
    'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'
  ];

  const STORAGE_PREFIX = 'tur_cert_custom_cfg_v1_';

  const imageCache = new Map();

  function loadImage(src) {
    if (!src) return Promise.resolve(null);
    if (imageCache.has(src)) return Promise.resolve(imageCache.get(src));
    return new Promise((resolve) => {
      if (typeof global.Image === 'undefined') {
        resolve(null);
        return;
      }
      const img = new global.Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        if (!src.startsWith('data:')) imageCache.set(src, img);
        resolve(img);
      };
      img.onerror = () => resolve(null);
      img.src = src;
    });
  }

  function formatBilingualEventDate(dateStr) {
    const clean = String(dateStr || '').trim();
    const m = clean.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/);
    if (!m) {
      const today = new Date();
      const y = today.getFullYear();
      const mo = today.getMonth() + 1;
      const d = today.getDate();
      return {
        iso: `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
        zh: `${y} 年 ${mo} 月 ${d} 日`,
        en: `${MONTHS_EN[mo - 1]} ${d}, ${y}`,
        year: String(y)
      };
    }
    const y = parseInt(m[1], 10);
    const mo = Math.min(12, Math.max(1, parseInt(m[2], 10)));
    const d = Math.min(31, Math.max(1, parseInt(m[3], 10)));
    return {
      iso: `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      zh: `${y} 年 ${mo} 月 ${d} 日`,
      en: `${MONTHS_EN[mo - 1]} ${d}, ${y}`,
      year: String(y)
    };
  }

  function extractRunnerMetrics(event, row) {
    const TE = global.TimingEngine;
    const isLap = event?.eventType === 'lap_timed' || row?.lapCount !== undefined;
    const dateInfo = formatBilingualEventDate(event?.date);
    const bib = String(row?.bib || '--').trim();
    const name = String(row?.name || '示範選手').trim();
    const catName = String(row?.categoryName || row?.categoryId || '--').trim();
    const overallRank = row?.overallRank ? `#${row.overallRank}` : '--';
    const categoryRank = row?.categoryRank ? `#${row.categoryRank}` : '--';
    const avgPace = row?.avgPace && row.avgPace !== '--' ? row.avgPace : '--';

    if (isLap) {
      const totalKm = Number(row?.totalDistanceKm || 0);
      const lapCount = Number(row?.lapCount || 0);
      const lapDist = Number(row?.lapDistanceKm || event?.lapDistanceKm || 2.54);
      const elapsedStr = row?.totalElapsedMs !== null && row?.totalElapsedMs !== undefined
        ? TE.formatDuration(row.totalElapsedMs, false)
        : '--:--:--';

      const validLaps = Array.isArray(row?.laps) ? row.laps.filter((l) => !l.overtime) : [];
      const splitItems = [];
      if (validLaps.length > 0 && validLaps.length <= 6) {
        validLaps.forEach((l) => {
          splitItems.push({
            title: `第 ${l.lapNumber} 圈`,
            time: l.lapDurationMs !== null ? TE.formatDuration(l.lapDurationMs, false) : '--:--:--',
            sub: l.lapPace && l.lapPace !== '--' ? l.lapPace : `${l.cumulativeKm ?? ''} km`
          });
        });
      } else if (validLaps.length > 6) {
        const firstLap = validLaps[0];
        const lastLap = validLaps[validLaps.length - 1];
        const bestLap = validLaps.reduce((b, l) => (l.lapDurationMs !== null && (!b || l.lapDurationMs < b.lapDurationMs) ? l : b), null);
        splitItems.push(
          {
            title: '首圈 FIRST LAP',
            time: firstLap.lapDurationMs !== null ? TE.formatDuration(firstLap.lapDurationMs, false) : '--:--:--',
            sub: firstLap.lapPace && firstLap.lapPace !== '--' ? firstLap.lapPace : ''
          },
          {
            title: '最速單圈 BEST LAP',
            time: row?.bestLapMs !== null && row?.bestLapMs !== undefined ? TE.formatDuration(row.bestLapMs, false) : '--:--:--',
            sub: bestLap?.lapPace && bestLap.lapPace !== '--' ? bestLap.lapPace : ''
          },
          {
            title: '平均單圈 AVG LAP',
            time: row?.avgLapMs !== null && row?.avgLapMs !== undefined ? TE.formatDuration(row.avgLapMs, false) : '--:--:--',
            sub: avgPace !== '--' ? avgPace : ''
          },
          {
            title: `末圈 LAP ${lastLap.lapNumber}`,
            time: lastLap.lapDurationMs !== null ? TE.formatDuration(lastLap.lapDurationMs, false) : '--:--:--',
            sub: lastLap.lapPace && lastLap.lapPace !== '--' ? lastLap.lapPace : ''
          }
        );
      }

      return {
        isLap: true,
        dateInfo,
        bib,
        name,
        categoryName: catName,
        overallRank,
        categoryRank,
        avgPace,
        primaryLabelZh: '累積總里程',
        primaryLabelEn: 'TOTAL DISTANCE',
        primaryValue: `${totalKm.toFixed(2)} km`,
        secondaryLabelZh: '累積圈數',
        secondaryLabelEn: 'TOTAL LAPS',
        secondaryValue: `${lapCount} Laps`,
        thirdLabelZh: '累積耗時',
        thirdLabelEn: 'ELAPSED TIME',
        thirdValue: elapsedStr,
        fourthLabelZh: '總排 / 組排',
        fourthLabelEn: 'OVERALL / CAT RANK',
        fourthValue: `${overallRank} / ${categoryRank}`,
        subNote: `單圈距離 LAP ${lapDist} km   •   平均配速 AVG PACE ${avgPace}`,
        splitTitleZh: '繞圈分段摘要',
        splitTitleEn: 'LAP SPLITS',
        splitItems,
        customFields: {
          name,
          bib,
          categoryName: catName,
          primaryResult: `${totalKm.toFixed(2)} km`,
          secondaryResult: `${lapCount} 圈 (${elapsedStr})`,
          overallRank,
          categoryRank,
          eventName: String(event?.name || ''),
          eventDate: dateInfo.iso
        }
      };
    }

    const timeStr = row?.gunDurationMs !== null && row?.gunDurationMs !== undefined
      ? TE.formatDuration(row.gunDurationMs, false)
      : (row?.finishTimeText ? (TE.normalizeTimeStr ? TE.normalizeTimeStr(row.finishTimeText) : String(row.finishTimeText).replace(/\.\d+$/, '')) : '--:--:--');
    const distKm = Number(row?.distanceKm || 0);
    const distStr = distKm > 0 ? `${distKm} km` : catName;

    const splitItems = (Array.isArray(row?.splits) ? row.splits : [])
      .filter((sp) => sp && sp.durationMs !== null && sp.durationMs !== undefined)
      .map((sp) => ({
        title: String(sp.toName || sp.toId || '').trim(),
        time: TE.formatDuration(sp.durationMs, false),
        sub: sp.pace && sp.pace !== '--' ? sp.pace : (sp.distanceKm > 0 ? `${Number(sp.distanceKm).toFixed(1)} km` : '')
      }));

    return {
      isLap: false,
      dateInfo,
      bib,
      name,
      categoryName: catName,
      overallRank,
      categoryRank,
      avgPace,
      primaryLabelZh: '完賽時間',
      primaryLabelEn: 'OFFICIAL TIME',
      primaryValue: timeStr,
      secondaryLabelZh: '完賽距離',
      secondaryLabelEn: 'DISTANCE',
      secondaryValue: distStr,
      thirdLabelZh: '總名次',
      thirdLabelEn: 'OVERALL RANK',
      thirdValue: overallRank,
      fourthLabelZh: '分組名次',
      fourthLabelEn: 'CATEGORY RANK',
      fourthValue: categoryRank,
      subNote: `平均配速 AVG PACE ${avgPace}`,
      splitTitleZh: '分段成績',
      splitTitleEn: 'SPLIT TIMES',
      splitItems,
      customFields: {
        name,
        bib,
        categoryName: catName,
        primaryResult: timeStr,
        secondaryResult: `${distStr} / ${avgPace}`,
        overallRank,
        categoryRank,
        eventName: String(event?.name || ''),
        eventDate: dateInfo.iso
      }
    };
  }

  function getDefaultCustomLayout() {
    return [
      { key: 'eventName', label: '活動名稱', enabled: false, x: 50, y: 16, fontSize: 52, color: '#1e293b', bold: true, align: 'center', prefix: '' },
      { key: 'name', label: '選手姓名', enabled: true, x: 50, y: 42, fontSize: 112, color: '#0f172a', bold: true, align: 'center', prefix: '' },
      { key: 'bib', label: '參賽編號', enabled: true, x: 34, y: 52, fontSize: 42, color: '#334155', bold: true, align: 'center', prefix: 'BIB: ' },
      { key: 'categoryName', label: '參賽組別', enabled: true, x: 66, y: 52, fontSize: 42, color: '#334155', bold: true, align: 'center', prefix: '' },
      { key: 'primaryResult', label: '主要成績 (時間/里程)', enabled: true, x: 50, y: 64, fontSize: 82, color: '#0f172a', bold: true, align: 'center', prefix: '' },
      { key: 'secondaryResult', label: '次要成績 (距離/圈數)', enabled: false, x: 50, y: 72, fontSize: 38, color: '#475569', bold: false, align: 'center', prefix: '' },
      { key: 'overallRank', label: '總名次', enabled: true, x: 35, y: 76, fontSize: 44, color: '#854d0e', bold: true, align: 'center', prefix: '總名次 ' },
      { key: 'categoryRank', label: '分組名次', enabled: true, x: 65, y: 76, fontSize: 44, color: '#854d0e', bold: true, align: 'center', prefix: '分組名次 ' },
      { key: 'eventDate', label: '活動日期', enabled: true, x: 50, y: 86, fontSize: 36, color: '#334155', bold: false, align: 'center', prefix: '' }
    ];
  }

  function loadSavedCustomConfig(eventId) {
    try {
      if (!eventId || typeof global.localStorage === 'undefined') return null;
      const raw = global.localStorage.getItem(STORAGE_PREFIX + eventId);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch {
      return null;
    }
  }

  function saveCustomConfig(eventId, cfg) {
    try {
      if (!eventId || typeof global.localStorage === 'undefined') return false;
      global.localStorage.setItem(STORAGE_PREFIX + eventId, JSON.stringify(cfg));
      return true;
    } catch {
      // 若底圖過大超過 localStorage 限額，僅儲存座標版面
      try {
        const { bgDataUrl, ...rest } = cfg || {};
        global.localStorage.setItem(STORAGE_PREFIX + eventId, JSON.stringify(rest));
      } catch {}
      return false;
    }
  }

  /* ========================================================================
   * 繪圖輔助工具
   * ======================================================================== */

  function drawRoundedRect(ctx, x, y, w, h, r) {
    const radius = Math.min(r, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.lineTo(x + w - radius, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
    ctx.lineTo(x + w, y + h - radius);
    ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
    ctx.lineTo(x + radius, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
    ctx.lineTo(x, y + radius);
    ctx.quadraticCurveTo(x, y, x + radius, y);
    ctx.closePath();
  }

  function fitTextFont(ctx, text, maxWidth, baseSize, minSize, fontTemplate) {
    let size = baseSize;
    while (size > minSize) {
      ctx.font = fontTemplate.replace('{size}', String(size));
      if (ctx.measureText(text).width <= maxWidth) break;
      size -= 2;
    }
    ctx.font = fontTemplate.replace('{size}', String(size));
    return size;
  }

  function drawCornerOrnament(ctx, cx, cy, sx, sy, size, colorMain, colorAccent) {
    ctx.save();
    ctx.translate(cx, cy);
    ctx.scale(sx, sy);

    // 外層折角線
    ctx.strokeStyle = colorMain;
    ctx.lineWidth = 4;
    ctx.beginPath();
    ctx.moveTo(0, size);
    ctx.lineTo(0, 14);
    ctx.lineTo(14, 0);
    ctx.lineTo(size, 0);
    ctx.stroke();

    // 內層幾何雙線
    ctx.strokeStyle = colorAccent;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(14, size * 0.78);
    ctx.lineTo(14, 24);
    ctx.lineTo(24, 14);
    ctx.lineTo(size * 0.78, 14);
    ctx.stroke();

    // 角落菱形徽章
    ctx.fillStyle = colorMain;
    ctx.beginPath();
    ctx.moveTo(26, 16);
    ctx.lineTo(36, 26);
    ctx.lineTo(26, 36);
    ctx.lineTo(16, 26);
    ctx.closePath();
    ctx.fill();

    // 延伸細線與點綴
    ctx.fillStyle = colorAccent;
    ctx.beginPath();
    ctx.arc(size + 14, 0, 4, 0, Math.PI * 2);
    ctx.arc(0, size + 14, 4, 0, Math.PI * 2);
    ctx.fill();

    ctx.restore();
  }

  function drawOfficialSeal(ctx, cx, cy, radius, yearStr, isDark) {
    ctx.save();
    ctx.translate(cx, cy);

    // 外圈波浪齒輪印章外緣
    const teeth = 36;
    const outerR = radius;
    const innerToothR = radius * 0.92;
    const grad = ctx.createLinearGradient(-outerR, -outerR, outerR, outerR);
    grad.addColorStop(0, '#f3d778');
    grad.addColorStop(0.35, '#d4a634');
    grad.addColorStop(0.7, '#b5841d');
    grad.addColorStop(1, '#f0d06b');

    ctx.beginPath();
    for (let i = 0; i < teeth * 2; i++) {
      const angle = (i * Math.PI) / teeth;
      const r = i % 2 === 0 ? outerR : innerToothR;
      const x = Math.cos(angle) * r;
      const y = Math.sin(angle) * r;
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    }
    ctx.closePath();
    ctx.fillStyle = grad;
    ctx.shadowColor = isDark ? 'rgba(0,0,0,0.45)' : 'rgba(133,96,20,0.28)';
    ctx.shadowBlur = 18;
    ctx.shadowOffsetY = 6;
    ctx.fill();
    ctx.shadowColor = 'transparent';

    // 內圈深金圓底
    ctx.beginPath();
    ctx.arc(0, 0, radius * 0.82, 0, Math.PI * 2);
    ctx.fillStyle = isDark ? '#142642' : '#fcf6e3';
    ctx.fill();
    ctx.lineWidth = 3.5;
    ctx.strokeStyle = '#b8860b';
    ctx.stroke();

    // 細內框線
    ctx.beginPath();
    ctx.arc(0, 0, radius * 0.72, 0, Math.PI * 2);
    ctx.lineWidth = 1.6;
    ctx.setLineDash([4, 3]);
    ctx.strokeStyle = isDark ? '#e6c465' : '#9a7016';
    ctx.stroke();
    ctx.setLineDash([]);

    // 鋼印文字
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const ink = isDark ? '#f3d778' : '#7c580f';
    const subInk = isDark ? '#d9b54a' : '#5d8226';

    ctx.fillStyle = subInk;
    ctx.font = `800 ${Math.round(radius * 0.105)}px "Segoe UI", Arial, sans-serif`;
    ctx.fillText('TAIWAN ULTRA RUNNERS', 0, -radius * 0.34);

    ctx.fillStyle = ink;
    ctx.font = `900 ${Math.round(radius * 0.30)}px Georgia, "Times New Roman", serif`;
    ctx.fillText('TUR', 0, -radius * 0.04);

    ctx.fillStyle = ink;
    ctx.font = `700 ${Math.round(radius * 0.118)}px "Noto Sans TC", "Microsoft JhengHei", sans-serif`;
    ctx.fillText(`官方完賽認證 • ${yearStr}`, 0, radius * 0.25);

    ctx.fillStyle = subInk;
    ctx.font = `700 ${Math.round(radius * 0.095)}px "Segoe UI", Arial, sans-serif`;
    ctx.fillText('★ OFFICIAL FINISHER ★', 0, radius * 0.44);

    ctx.restore();
  }

  /* ========================================================================
   * 模式 1：官方公版證書（中英文對照・高質感官方模式・含 TUR LOGO）
   * ======================================================================== */

  async function renderOfficialCertificate(canvas, { event, row, options = {} }) {
    const orientation = options.orientation === 'portrait' ? 'portrait' : 'landscape';
    const theme = options.theme === 'navy_gold' ? 'navy_gold' : 'ivory_gold';
    const isDark = theme === 'navy_gold';
    const W = orientation === 'portrait' ? 1754 : 2480;
    const H = orientation === 'portrait' ? 2480 : 1754;
    canvas.width = W;
    canvas.height = H;

    const ctx = canvas.getContext('2d');
    const metrics = extractRunnerMetrics(event, row);
    const logoImg = await loadImage('./tur-logo.jpg');

    const orgName = String(options.organization || event?.organization || 'Taiwan Ultra Runners').trim() || 'Taiwan Ultra Runners';
    const eventNameZh = String(event?.name || 'TUR 路跑挑戰賽').trim();
    const eventNameEn = String(options.englishName ?? event?.englishName ?? '').trim();
    const showSplits = options.showSplits !== false;
    const locationText = String(event?.location || '').trim();

    // 配色表
    const palette = isDark
      ? {
          bgInner: '#132645',
          bgMid: '#0c1a31',
          bgOuter: '#060e1c',
          goldMain: '#e5c158',
          goldSoft: '#b88f2d',
          turGreen: '#8ec641',
          textPrimary: '#ffffff',
          textSecondary: '#cbd5e1',
          textMuted: '#94a3b8',
          cardBg: 'rgba(15, 29, 54, 0.82)',
          cardBorder: 'rgba(229, 193, 88, 0.48)',
          divider: 'rgba(229, 193, 88, 0.4)',
          watermark: 'rgba(229, 193, 88, 0.045)',
          ridge: 'rgba(142, 198, 65, 0.09)'
        }
      : {
          bgInner: '#fffefa',
          bgMid: '#faf4e6',
          bgOuter: '#efe5cc',
          goldMain: '#b8860b',
          goldSoft: '#cfa33a',
          turGreen: '#638c26',
          textPrimary: '#1a2332',
          textSecondary: '#3b4a5e',
          textMuted: '#64748b',
          cardBg: 'rgba(255, 255, 255, 0.82)',
          cardBorder: 'rgba(184, 134, 11, 0.38)',
          divider: 'rgba(184, 134, 11, 0.35)',
          watermark: 'rgba(99, 140, 38, 0.045)',
          ridge: 'rgba(99, 140, 38, 0.08)'
        };

    // 1. 宣紙／深藍質感底色
    const bgGrad = ctx.createRadialGradient(W * 0.5, H * 0.42, W * 0.05, W * 0.5, H * 0.5, Math.max(W, H) * 0.72);
    bgGrad.addColorStop(0, palette.bgInner);
    bgGrad.addColorStop(0.65, palette.bgMid);
    bgGrad.addColorStop(1, palette.bgOuter);
    ctx.fillStyle = bgGrad;
    ctx.fillRect(0, 0, W, H);

    // 2. 細膩幾何網紋與台灣中央山脈剪影浮水印
    ctx.save();
    ctx.strokeStyle = palette.ridge;
    ctx.lineWidth = 2.5;
    const ridgeY = H * 0.76;
    ctx.beginPath();
    ctx.moveTo(W * 0.06, ridgeY + 60);
    ctx.lineTo(W * 0.18, ridgeY - 35);
    ctx.lineTo(W * 0.28, ridgeY + 20);
    ctx.lineTo(W * 0.41, ridgeY - 88);
    ctx.lineTo(W * 0.51, ridgeY - 12);
    ctx.lineTo(W * 0.62, ridgeY - 102);
    ctx.lineTo(W * 0.74, ridgeY - 18);
    ctx.lineTo(W * 0.84, ridgeY - 52);
    ctx.lineTo(W * 0.94, ridgeY + 55);
    ctx.stroke();

    // 淡雅中央 TUR 浮水印
    ctx.fillStyle = palette.watermark;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = `900 ${Math.round(Math.min(W, H) * 0.34)}px Georgia, serif`;
    ctx.fillText('TUR', W * 0.5, H * 0.53);
    ctx.restore();

    // 3. 官方多層燙金與 TUR 橄欖綠典雅邊框
    const pad1 = Math.round(Math.min(W, H) * 0.036);
    const pad2 = pad1 + 18;
    const pad3 = pad2 + 12;

    // 外細綠金裝飾線
    ctx.strokeStyle = palette.turGreen;
    ctx.lineWidth = 4;
    ctx.strokeRect(pad1, pad1, W - pad1 * 2, H - pad1 * 2);

    // 主粗燙金框
    const frameGrad = ctx.createLinearGradient(0, 0, W, H);
    frameGrad.addColorStop(0, '#d8b44a');
    frameGrad.addColorStop(0.5, palette.goldMain);
    frameGrad.addColorStop(1, '#e0be59');
    ctx.strokeStyle = frameGrad;
    ctx.lineWidth = 9;
    ctx.strokeRect(pad2, pad2, W - pad2 * 2, H - pad2 * 2);

    // 內細線框
    ctx.strokeStyle = palette.goldSoft;
    ctx.lineWidth = 2;
    ctx.strokeRect(pad3, pad3, W - pad3 * 2, H - pad3 * 2);

    // 四角古典飾紋
    const ornSize = Math.round(Math.min(W, H) * 0.068);
    drawCornerOrnament(ctx, pad3 + 12, pad3 + 12, 1, 1, ornSize, palette.goldMain, palette.turGreen);
    drawCornerOrnament(ctx, W - pad3 - 12, pad3 + 12, -1, 1, ornSize, palette.goldMain, palette.turGreen);
    drawCornerOrnament(ctx, pad3 + 12, H - pad3 - 12, 1, -1, ornSize, palette.goldMain, palette.turGreen);
    drawCornerOrnament(ctx, W - pad3 - 12, H - pad3 - 12, -1, -1, ornSize, palette.goldMain, palette.turGreen);

    // 佈局縱向座標（LOGO 移至下方與日期、鋼印並列；上方放大活動名稱與字體，並加入分段成績卡）
    const isPort = orientation === 'portrait';
    const contentW = W - (pad3 + 90) * 2;
    const centerX = W * 0.5;
    const activeSplits = showSplits ? (metrics.splitItems || []) : [];
    const hasSplits = activeSplits.length > 0;

    // 4. 官方中英文抬頭（字體適度放大）
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    const titleEnY = pad3 + (isPort ? 105 : 78);
    ctx.fillStyle = palette.turGreen;
    ctx.font = `700 ${isPort ? 36 : 34}px "Segoe UI", "Helvetica Neue", Arial, sans-serif`;
    ctx.fillText('OFFICIAL FINISHER CERTIFICATE', centerX, titleEnY);

    const titleZhY = titleEnY + (isPort ? 86 : 74);
    ctx.fillStyle = palette.textPrimary;
    ctx.font = `900 ${isPort ? 102 : 94}px "Noto Serif TC", "PMingLiU", "DFKai-SB", "Microsoft JhengHei", serif`;
    ctx.fillText('官 方 完 賽 證 書', centerX, titleZhY);

    // 金色中軸飾線
    const divY = titleZhY + (isPort ? 66 : 56);
    ctx.save();
    ctx.strokeStyle = palette.goldMain;
    ctx.lineWidth = 3;
    ctx.beginPath();
    ctx.moveTo(centerX - 360, divY);
    ctx.lineTo(centerX - 26, divY);
    ctx.moveTo(centerX + 26, divY);
    ctx.lineTo(centerX + 360, divY);
    ctx.stroke();
    ctx.fillStyle = palette.goldMain;
    ctx.beginPath();
    ctx.moveTo(centerX, divY - 11);
    ctx.lineTo(centerX + 11, divY);
    ctx.lineTo(centerX, divY + 11);
    ctx.lineTo(centerX - 11, divY);
    ctx.closePath();
    ctx.fill();
    ctx.restore();

    // 5. 活動名稱（中英文對照・特別加大字體！）
    const evZhY = divY + (isPort ? 88 : 70);
    ctx.fillStyle = palette.textPrimary;
    fitTextFont(ctx, eventNameZh, contentW, isPort ? 72 : 76, 38, '900 {size}px "Noto Sans TC", "Microsoft JhengHei", sans-serif');
    ctx.fillText(eventNameZh, centerX, evZhY);

    let nextY = evZhY + (isPort ? 68 : 60);
    if (eventNameEn) {
      ctx.fillStyle = palette.goldMain;
      fitTextFont(ctx, eventNameEn, contentW, isPort ? 40 : 40, 22, '700 {size}px Georgia, "Times New Roman", serif');
      ctx.fillText(eventNameEn, centerX, nextY);
      nextY += isPort ? 70 : 54;
    } else {
      nextY += isPort ? 26 : 16;
    }

    // 6. 茲證明 / THIS IS TO CERTIFY THAT
    ctx.fillStyle = palette.textMuted;
    ctx.font = `600 ${isPort ? 30 : 28}px "Noto Sans TC", "Segoe UI", sans-serif`;
    ctx.fillText('茲 證 明 選 手 順 利 完 成 本 屆 賽 事   •   THIS IS TO CERTIFY THAT', centerX, nextY);

    // 7. 選手姓名主焦點（放大呈現）
    const nameY = nextY + (isPort ? 128 : (hasSplits ? 98 : 114));
    ctx.fillStyle = palette.textPrimary;
    fitTextFont(ctx, metrics.name, contentW * 0.85, isPort ? 150 : 144, 72, '900 {size}px "Noto Serif TC", "Microsoft JhengHei", Georgia, serif');
    ctx.fillText(metrics.name, centerX, nameY);

    // 姓名下方細金線
    const nameLineY = nameY + (isPort ? 90 : 76);
    ctx.strokeStyle = palette.divider;
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    ctx.moveTo(centerX - 320, nameLineY);
    ctx.lineTo(centerX + 320, nameLineY);
    ctx.stroke();

    // 參賽編號 & 組別（放大）
    const bibLineY = nameLineY + (isPort ? 56 : 46);
    ctx.fillStyle = palette.textSecondary;
    ctx.font = `700 ${isPort ? 40 : 38}px "Noto Sans TC", "Segoe UI", sans-serif`;
    ctx.fillText(`參賽編號 BIB NO.  ${metrics.bib}      |      參賽組別 CATEGORY  ${metrics.categoryName}`, centerX, bibLineY);

    // 8. 官方中英文對照主成績卡（4 欄・放大數值與標籤）
    const cardTopY = bibLineY + (isPort ? 66 : (hasSplits ? 44 : 62));
    const cardW = isPort ? contentW : Math.min(contentW, 2060);
    const cardH = isPort ? 460 : 272;
    const cardX = centerX - cardW / 2;

    ctx.save();
    drawRoundedRect(ctx, cardX, cardTopY, cardW, cardH, 24);
    ctx.fillStyle = palette.cardBg;
    ctx.fill();
    ctx.lineWidth = 2.5;
    ctx.strokeStyle = palette.cardBorder;
    ctx.stroke();

    const items = [
      { zh: metrics.primaryLabelZh, en: metrics.primaryLabelEn, val: metrics.primaryValue, highlight: true },
      { zh: metrics.secondaryLabelZh, en: metrics.secondaryLabelEn, val: metrics.secondaryValue, highlight: false },
      { zh: metrics.thirdLabelZh, en: metrics.thirdLabelEn, val: metrics.thirdValue, highlight: false },
      { zh: metrics.fourthLabelZh, en: metrics.fourthLabelEn, val: metrics.fourthValue, highlight: false }
    ];

    if (isPort) {
      // 直式 2x2 格
      const cellW = cardW / 2;
      const cellH = (cardH - 76) / 2;
      ctx.strokeStyle = palette.divider;
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(cardX + cellW, cardTopY + 24);
      ctx.lineTo(cardX + cellW, cardTopY + cardH - 80);
      ctx.moveTo(cardX + 36, cardTopY + cellH);
      ctx.lineTo(cardX + cardW - 36, cardTopY + cellH);
      ctx.stroke();

      items.forEach((it, idx) => {
        const col = idx % 2;
        const r = Math.floor(idx / 2);
        const cx = cardX + col * cellW + cellW / 2;
        const cy = cardTopY + r * cellH;
        ctx.fillStyle = palette.textMuted;
        ctx.font = '700 30px "Noto Sans TC", "Microsoft JhengHei", sans-serif';
        ctx.fillText(it.zh, cx, cy + 44);
        ctx.font = '600 21px "Segoe UI", Arial, sans-serif';
        ctx.fillText(it.en, cx, cy + 74);
        ctx.fillStyle = it.highlight ? palette.turGreen : palette.textPrimary;
        fitTextFont(ctx, it.val, cellW - 48, 70, 34, '900 {size}px ui-monospace, "Segoe UI", monospace');
        ctx.fillText(it.val, cx, cy + 138);
      });
    } else {
      // 橫式 4x1 格
      const cellW = cardW / 4;
      const bodyH = cardH - 64;
      ctx.strokeStyle = palette.divider;
      ctx.lineWidth = 1.5;
      for (let i = 1; i < 4; i++) {
        ctx.beginPath();
        ctx.moveTo(cardX + cellW * i, cardTopY + 26);
        ctx.lineTo(cardX + cellW * i, cardTopY + bodyH - 10);
        ctx.stroke();
      }
      items.forEach((it, idx) => {
        const cx = cardX + idx * cellW + cellW / 2;
        ctx.fillStyle = palette.textMuted;
        ctx.font = '700 30px "Noto Sans TC", "Microsoft JhengHei", sans-serif';
        ctx.fillText(it.zh, cx, cardTopY + 46);
        ctx.font = '600 21px "Segoe UI", Arial, sans-serif';
        ctx.fillText(it.en, cx, cardTopY + 76);
        ctx.fillStyle = it.highlight ? palette.turGreen : palette.textPrimary;
        fitTextFont(ctx, it.val, cellW - 40, 72, 34, '900 {size}px ui-monospace, "Segoe UI", monospace');
        ctx.fillText(it.val, cx, cardTopY + 146);
      });
    }

    // 卡片底部配速附註條
    const stripY = cardTopY + cardH - 56;
    ctx.strokeStyle = palette.divider;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(cardX + 36, stripY - 12);
    ctx.lineTo(cardX + cardW - 36, stripY - 12);
    ctx.stroke();

    ctx.fillStyle = palette.textSecondary;
    ctx.font = '700 28px "Noto Sans TC", "Segoe UI", sans-serif';
    ctx.fillText(metrics.subNote, centerX, stripY + 20);
    ctx.restore();

    // 9. 分段成績卡（分段成績 SPLIT TIMES / 繞圈分段摘要 LAP SPLITS）
    if (hasSplits) {
      const shownSplits = activeSplits.slice(0, 6);
      const splitTopY = cardTopY + cardH + (isPort ? 34 : 26);
      const splitH = isPort ? (shownSplits.length > 4 ? 372 : 284) : 196;

      ctx.save();
      drawRoundedRect(ctx, cardX, splitTopY, cardW, splitH, 20);
      ctx.fillStyle = palette.cardBg;
      ctx.fill();
      ctx.lineWidth = 2;
      ctx.strokeStyle = palette.cardBorder;
      ctx.stroke();

      // 分段卡標題條
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillStyle = palette.goldMain;
      ctx.font = `800 ${isPort ? 26 : 24}px "Noto Sans TC", "Segoe UI", sans-serif`;
      ctx.fillText(`${metrics.splitTitleZh}   •   ${metrics.splitTitleEn}`, centerX, splitTopY + 30);

      ctx.strokeStyle = palette.divider;
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(cardX + 32, splitTopY + 54);
      ctx.lineTo(cardX + cardW - 32, splitTopY + 54);
      ctx.stroke();

      if (isPort) {
        // 直式採 2 欄網格排列分段
        const cols = Math.min(2, shownSplits.length);
        const rowsCount = Math.ceil(shownSplits.length / cols);
        const cellW = cardW / cols;
        const cellH = (splitH - 64) / rowsCount;

        if (cols === 2) {
          ctx.beginPath();
          ctx.moveTo(cardX + cellW, splitTopY + 66);
          ctx.lineTo(cardX + cellW, splitTopY + splitH - 16);
          ctx.stroke();
        }
        for (let r = 1; r < rowsCount; r++) {
          ctx.beginPath();
          ctx.moveTo(cardX + 32, splitTopY + 58 + cellH * r);
          ctx.lineTo(cardX + cardW - 32, splitTopY + 58 + cellH * r);
          ctx.stroke();
        }

        shownSplits.forEach((sp, idx) => {
          const col = idx % cols;
          const r = Math.floor(idx / cols);
          const cx = cardX + col * cellW + cellW / 2;
          const cy = splitTopY + 58 + r * cellH;
          ctx.fillStyle = palette.textSecondary;
          fitTextFont(ctx, sp.title, cellW - 36, 25, 18, '700 {size}px "Noto Sans TC", "Microsoft JhengHei", sans-serif');
          ctx.fillText(sp.title, cx, cy + cellH * 0.25);

          ctx.fillStyle = palette.textPrimary;
          fitTextFont(ctx, sp.time, cellW - 36, 38, 24, '900 {size}px ui-monospace, "Segoe UI", monospace');
          ctx.fillText(sp.time, cx, cy + cellH * 0.58);

          if (sp.sub) {
            ctx.fillStyle = palette.textMuted;
            ctx.font = '600 21px "Segoe UI", ui-monospace, sans-serif';
            ctx.fillText(sp.sub, cx, cy + cellH * 0.85);
          }
        });
      } else {
        // 橫式採單列等寬欄位（最多 6 站）
        const cols = shownSplits.length;
        const cellW = cardW / cols;
        for (let i = 1; i < cols; i++) {
          ctx.beginPath();
          ctx.moveTo(cardX + cellW * i, splitTopY + 68);
          ctx.lineTo(cardX + cellW * i, splitTopY + splitH - 18);
          ctx.stroke();
        }
        shownSplits.forEach((sp, idx) => {
          const cx = cardX + idx * cellW + cellW / 2;
          ctx.fillStyle = palette.textSecondary;
          fitTextFont(ctx, sp.title, cellW - 30, 26, 17, '700 {size}px "Noto Sans TC", "Microsoft JhengHei", sans-serif');
          ctx.fillText(sp.title, cx, splitTopY + 86);

          ctx.fillStyle = palette.textPrimary;
          fitTextFont(ctx, sp.time, cellW - 28, 40, 24, '900 {size}px ui-monospace, "Segoe UI", monospace');
          ctx.fillText(sp.time, cx, splitTopY + 132);

          if (sp.sub) {
            ctx.fillStyle = palette.textMuted;
            ctx.font = '600 22px "Segoe UI", ui-monospace, sans-serif';
            ctx.fillText(sp.sub, cx, splitTopY + 168);
          }
        });
      }
      ctx.restore();
    }

    // 10. 底部三欄並列官方落款區（左：活動日期與地點｜中：官方金色鋼印｜右：TUR 官方 LOGO ＋ 主辦單位；無賽事總監字樣）
    const footerY = H - pad3 - (isPort ? 245 : 178);
    const leftX = pad3 + (isPort ? 285 : 400);
    const rightX = W - pad3 - (isPort ? 285 : 400);

    // 左欄：活動日期（以活動日期為主，字體適度放大）
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillStyle = palette.textMuted;
    ctx.font = '700 25px "Noto Sans TC", "Segoe UI", sans-serif';
    ctx.fillText('活動日期  EVENT DATE', leftX, footerY - 62);

    ctx.fillStyle = palette.textPrimary;
    ctx.font = '800 40px "Noto Sans TC", "Microsoft JhengHei", sans-serif';
    ctx.fillText(metrics.dateInfo.zh, leftX, footerY - 10);

    ctx.fillStyle = palette.goldMain;
    ctx.font = '700 27px Georgia, "Times New Roman", serif';
    ctx.fillText(metrics.dateInfo.en, leftX, footerY + 34);

    if (locationText) {
      ctx.fillStyle = palette.textSecondary;
      fitTextFont(ctx, locationText, 520, 25, 18, '600 {size}px "Noto Sans TC", "Microsoft JhengHei", sans-serif');
      ctx.fillText(locationText, leftX, footerY + 76);
    }

    // 中欄：官方金色完賽鋼印
    drawOfficialSeal(ctx, centerX, footerY + 4, isPort ? 124 : 112, metrics.dateInfo.year, isDark);

    // 右欄：TUR 官方 LOGO 置於下方並列 ＋ 主辦單位名稱（取消賽事總監字樣）
    const footLogoH = isPort ? 92 : 94;
    if (logoImg && logoImg.naturalWidth > 0) {
      const ratio = logoImg.naturalWidth / logoImg.naturalHeight;
      const drawH = footLogoH;
      const drawW = Math.min(drawH * ratio, isPort ? 360 : 420);
      const finalH = drawW / ratio;
      const logoY = footerY - 68;
      if (isDark) {
        const bx = rightX - drawW / 2 - 20;
        const by = logoY - 12;
        const bw = drawW + 40;
        const bh = finalH + 24;
        ctx.save();
        drawRoundedRect(ctx, bx, by, bw, bh, 14);
        ctx.fillStyle = '#ffffff';
        ctx.fill();
        ctx.lineWidth = 2.5;
        ctx.strokeStyle = palette.goldMain;
        ctx.stroke();
        ctx.drawImage(logoImg, rightX - drawW / 2, logoY, drawW, finalH);
        ctx.restore();
      } else {
        ctx.save();
        ctx.globalCompositeOperation = 'multiply';
        ctx.drawImage(logoImg, rightX - drawW / 2, logoY, drawW, finalH);
        ctx.restore();
      }
    }

    ctx.textAlign = 'center';
    ctx.fillStyle = palette.textMuted;
    ctx.font = '700 22px "Noto Sans TC", "Segoe UI", sans-serif';
    ctx.fillText('主辦單位  ORGANIZER', rightX, footerY + 44);

    ctx.fillStyle = palette.textPrimary;
    fitTextFont(ctx, orgName, 520, 32, 22, '800 {size}px Georgia, "Times New Roman", serif');
    ctx.fillText(orgName, rightX, footerY + 80);

    // 右下角官方流水編號（避開右下角古典角花飾紋）
    ctx.textAlign = 'right';
    ctx.fillStyle = palette.textMuted;
    ctx.font = '600 18px ui-monospace, monospace';
    ctx.fillText(`CERT NO. TUR-${metrics.dateInfo.year}-${metrics.bib}`, W - pad3 - ornSize - 42, H - pad3 - 26);

    return { width: W, height: H, metrics };
  }

  /* ========================================================================
   * 模式 2：承辦人自訂底圖模式（上傳底圖＋成績套印＋可拖曳預覽框）
   * ======================================================================== */

  async function renderCustomBgCertificate(canvas, {
    event,
    row,
    bgImage = null,
    layout = [],
    options = {},
    selectedKey = '',
    showGuides = false
  }) {
    const metrics = extractRunnerMetrics(event, row);
    const maxEdge = 2480;
    let W = 2480;
    let H = 1754;

    if (bgImage && bgImage.naturalWidth > 0 && bgImage.naturalHeight > 0) {
      const nw = bgImage.naturalWidth;
      const nh = bgImage.naturalHeight;
      const scale = Math.min(maxEdge / Math.max(nw, nh), 2);
      W = Math.round(nw * scale);
      H = Math.round(nh * scale);
    }

    canvas.width = W;
    canvas.height = H;
    const ctx = canvas.getContext('2d');

    // 1. 繪製承辦人上傳的底圖（或未上傳時的引導預覽底）
    if (bgImage && bgImage.naturalWidth > 0) {
      ctx.drawImage(bgImage, 0, 0, W, H);
    } else {
      ctx.fillStyle = '#f8fafc';
      ctx.fillRect(0, 0, W, H);
      ctx.strokeStyle = '#cbd5e1';
      ctx.lineWidth = 6;
      ctx.setLineDash([20, 14]);
      ctx.strokeRect(48, 48, W - 96, H - 96);
      ctx.setLineDash([]);

      ctx.fillStyle = '#64748b';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.font = '800 54px "Noto Sans TC", "Microsoft JhengHei", sans-serif';
      ctx.fillText('尚未上傳自訂證書底圖', W / 2, H * 0.22);
      ctx.font = '600 32px "Noto Sans TC", "Microsoft JhengHei", sans-serif';
      ctx.fillText('請於左側點擊「📤 上傳自訂底圖 (JPG / PNG)」，並可直接拖曳下方成績文字對齊位置', W / 2, H * 0.27);
    }

    // 2. 可選疊加 TUR LOGO（左上角或置頂）
    if (options.showTurLogo) {
      const logoImg = await loadImage('./tur-logo.jpg');
      if (logoImg && logoImg.naturalWidth > 0) {
        const lh = Math.round(H * 0.068);
        const lw = Math.round(lh * (logoImg.naturalWidth / logoImg.naturalHeight));
        const lx = Math.round(W * 0.05);
        const ly = Math.round(H * 0.05);
        ctx.save();
        drawRoundedRect(ctx, lx - 14, ly - 10, lw + 28, lh + 20, 14);
        ctx.fillStyle = 'rgba(255,255,255,0.92)';
        ctx.fill();
        ctx.drawImage(logoImg, lx, ly, lw, lh);
        ctx.restore();
      }
    }

    // 3. 套印所有啟用中的成績欄位（並回傳命中框供拖曳定位）
    const scaleFactor = W / 2480;
    const hitBoxes = [];

    for (const item of layout) {
      if (!item || !item.enabled) continue;
      const rawVal = metrics.customFields[item.key] ?? '';
      const text = `${item.prefix || ''}${rawVal}`;
      if (!text.trim()) continue;

      const px = (Number(item.x || 50) / 100) * W;
      const py = (Number(item.y || 50) / 100) * H;
      const fSize = Math.max(14, Math.round(Number(item.fontSize || 48) * scaleFactor));
      const weight = item.bold ? '800' : '500';
      const align = item.align || 'center';

      ctx.save();
      ctx.font = `${weight} ${fSize}px "Noto Sans TC", "Microsoft JhengHei", "Segoe UI", sans-serif`;
      ctx.fillStyle = item.color || '#0f172a';
      ctx.textAlign = align;
      ctx.textBaseline = 'middle';
      ctx.fillText(text, px, py);

      const textW = Math.max(60, ctx.measureText(text).width);
      const textH = fSize * 1.25;
      let boxX = px - textW / 2;
      if (align === 'left') boxX = px;
      else if (align === 'right') boxX = px - textW;
      const boxY = py - textH / 2;

      hitBoxes.push({ key: item.key, x: boxX - 12, y: boxY - 8, w: textW + 24, h: textH + 16 });

      if (showGuides) {
        const isSel = item.key === selectedKey;
        ctx.strokeStyle = isSel ? '#f59e0b' : 'rgba(56, 189, 248, 0.55)';
        ctx.lineWidth = isSel ? Math.max(4, 5 * scaleFactor) : Math.max(2, 2.5 * scaleFactor);
        ctx.setLineDash(isSel ? [14 * scaleFactor, 8 * scaleFactor] : [8 * scaleFactor, 8 * scaleFactor]);
        ctx.strokeRect(boxX - 12, boxY - 8, textW + 24, textH + 16);
        ctx.setLineDash([]);
      }
      ctx.restore();
    }

    return { width: W, height: H, metrics, hitBoxes };
  }

  /* ========================================================================
   * 模式 3：選手客製化個人證書（上傳底圖 ＋ 個人照片整合・公開查詢頁專用）
   * ======================================================================== */

  function drawCoverImage(ctx, img, x, y, w, h, zoom = 1, offsetXPct = 0, offsetYPct = 0) {
    const iw = img.naturalWidth || img.width;
    const ih = img.naturalHeight || img.height;
    if (!iw || !ih) return;
    const baseScale = Math.max(w / iw, h / ih);
    const scale = baseScale * Math.max(0.5, Number(zoom || 1));
    const dw = iw * scale;
    const dh = ih * scale;
    const maxPanX = Math.max(0, (dw - w) / 2);
    const maxPanY = Math.max(0, (dh - h) / 2);
    const dx = x + (w - dw) / 2 + (Number(offsetXPct || 0) / 50) * Math.max(maxPanX, w * 0.25);
    const dy = y + (h - dh) / 2 + (Number(offsetYPct || 0) / 50) * Math.max(maxPanY, h * 0.25);
    ctx.drawImage(img, dx, dy, dw, dh);
  }

  async function renderPersonalizedCertificate(canvas, {
    event,
    row,
    bgImage = null,
    photoImage = null,
    options = {}
  }) {
    const orientation = options.orientation === 'portrait' ? 'portrait' : 'landscape';
    const isPort = orientation === 'portrait';
    const isDarkTextStyle = options.cardStyle === 'dark_glass';
    const photoPosition = options.photoPosition === 'right' ? 'right' : 'left';
    const photoShape = ['rounded', 'circle', 'polaroid'].includes(options.photoShape) ? options.photoShape : 'rounded';
    const showLogo = options.showLogo !== false;
    const showSeal = options.showSeal !== false;
    const showSplits = options.showSplits !== false;
    const photoZoom = Number(options.photoZoom || 1);
    const photoOffsetX = Number(options.photoOffsetX || 0);
    const photoOffsetY = Number(options.photoOffsetY || 0);
    const customQuote = String(options.customQuote || '').trim();

    const W = isPort ? 1754 : 2480;
    const H = isPort ? 2480 : 1754;
    canvas.width = W;
    canvas.height = H;

    const ctx = canvas.getContext('2d');
    const metrics = extractRunnerMetrics(event, row);
    const logoImg = showLogo ? await loadImage('./tur-logo.jpg') : null;

    const orgName = String(options.organization || event?.organization || 'Taiwan Ultra Runners').trim() || 'Taiwan Ultra Runners';
    const eventNameZh = String(event?.name || 'TUR 路跑挑戰賽').trim();
    const eventNameEn = String(options.englishName ?? event?.englishName ?? '').trim();
    const locationText = String(event?.location || '').trim();

    const palette = isDarkTextStyle
      ? {
          bgInner: '#132645',
          bgMid: '#0c1a31',
          bgOuter: '#060e1c',
          scrim: 'rgba(6, 14, 28, 0.50)',
          panelBg: 'rgba(10, 22, 42, 0.80)',
          panelBorder: 'rgba(229, 193, 88, 0.55)',
          goldMain: '#e5c158',
          goldSoft: '#b88f2d',
          turGreen: '#8ec641',
          textPrimary: '#ffffff',
          textSecondary: '#e2e8f0',
          textMuted: '#cbd5e1',
          divider: 'rgba(229, 193, 88, 0.42)',
          cardBg: 'rgba(15, 31, 58, 0.72)'
        }
      : {
          bgInner: '#fffefa',
          bgMid: '#faf4e6',
          bgOuter: '#efe5cc',
          scrim: 'rgba(255, 252, 245, 0.24)',
          panelBg: 'rgba(255, 253, 248, 0.88)',
          panelBorder: 'rgba(184, 134, 11, 0.48)',
          goldMain: '#b8860b',
          goldSoft: '#cfa33a',
          turGreen: '#638c26',
          textPrimary: '#1a2332',
          textSecondary: '#334155',
          textMuted: '#64748b',
          divider: 'rgba(184, 134, 11, 0.35)',
          cardBg: 'rgba(255, 255, 255, 0.88)'
        };

    // 1. 底圖層：選手上傳自訂底圖（等比滿版裁切）或優雅官方預設底紋
    if (bgImage && bgImage.naturalWidth > 0) {
      drawCoverImage(ctx, bgImage, 0, 0, W, H, 1, 0, 0);
      ctx.fillStyle = palette.scrim;
      ctx.fillRect(0, 0, W, H);
    } else {
      const bgGrad = ctx.createRadialGradient(W * 0.5, H * 0.42, W * 0.05, W * 0.5, H * 0.5, Math.max(W, H) * 0.72);
      bgGrad.addColorStop(0, palette.bgInner);
      bgGrad.addColorStop(0.65, palette.bgMid);
      bgGrad.addColorStop(1, palette.bgOuter);
      ctx.fillStyle = bgGrad;
      ctx.fillRect(0, 0, W, H);
    }

    // 2. 典雅外框＋四角飾紋
    const pad1 = Math.round(Math.min(W, H) * 0.032);
    const pad2 = pad1 + 16;
    const pad3 = pad2 + 10;

    ctx.strokeStyle = palette.turGreen;
    ctx.lineWidth = 4;
    ctx.strokeRect(pad1, pad1, W - pad1 * 2, H - pad1 * 2);

    ctx.strokeStyle = palette.goldMain;
    ctx.lineWidth = 8;
    ctx.strokeRect(pad2, pad2, W - pad2 * 2, H - pad2 * 2);

    ctx.strokeStyle = palette.goldSoft;
    ctx.lineWidth = 2;
    ctx.strokeRect(pad3, pad3, W - pad3 * 2, H - pad3 * 2);

    const ornSize = Math.round(Math.min(W, H) * 0.058);
    drawCornerOrnament(ctx, pad3 + 10, pad3 + 10, 1, 1, ornSize, palette.goldMain, palette.turGreen);
    drawCornerOrnament(ctx, W - pad3 - 10, pad3 + 10, -1, 1, ornSize, palette.goldMain, palette.turGreen);
    drawCornerOrnament(ctx, pad3 + 10, H - pad3 - 10, 1, -1, ornSize, palette.goldMain, palette.turGreen);
    drawCornerOrnament(ctx, W - pad3 - 10, H - pad3 - 10, -1, -1, ornSize, palette.goldMain, palette.turGreen);

    // 3. 計算「個人照片區」與「成績證書資訊區」版位
    const innerX = pad3 + 64;
    const innerY = pad3 + 64;
    const innerW = W - innerX * 2;
    const innerH = H - innerY * 2;

    let photoBox;
    let infoBox;
    const gap = 52;

    if (isPort) {
      // 直式：上方個人照片、下方毛玻璃成績主卡
      const phH = Math.round(innerH * 0.34);
      const phW = Math.round(innerW * 0.56);
      photoBox = {
        x: Math.round(innerX + (innerW - phW) / 2),
        y: innerY + 12,
        w: phW,
        h: phH
      };
      infoBox = {
        x: innerX,
        y: photoBox.y + photoBox.h + 38,
        w: innerW,
        h: innerH - (photoBox.h + 50)
      };
    } else {
      // 橫式：左圖右文 或 左文右圖
      const phW = Math.round(innerW * 0.35);
      const infW = innerW - phW - gap;
      if (photoPosition === 'right') {
        infoBox = { x: innerX, y: innerY, w: infW, h: innerH };
        photoBox = { x: innerX + infW + gap, y: innerY + 28, w: phW, h: innerH - 56 };
      } else {
        photoBox = { x: innerX, y: innerY + 28, w: phW, h: innerH - 56 };
        infoBox = { x: innerX + phW + gap, y: innerY, w: infW, h: innerH };
      }
    }

    // 4. 繪製個人照片區（圓角金邊 / 圓形徽章 / 拍立得白框）
    ctx.save();
    if (photoShape === 'circle') {
      const diameter = Math.min(photoBox.w, photoBox.h);
      const cx = photoBox.x + photoBox.w / 2;
      const cy = photoBox.y + photoBox.h / 2;
      const r = diameter / 2;

      ctx.shadowColor = 'rgba(0,0,0,0.28)';
      ctx.shadowBlur = 32;
      ctx.shadowOffsetY = 12;
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
      ctx.fillStyle = isDarkTextStyle ? '#0f172a' : '#f8fafc';
      ctx.fill();
      ctx.shadowColor = 'transparent';

      ctx.save();
      ctx.beginPath();
      ctx.arc(cx, cy, r - 10, 0, Math.PI * 2);
      ctx.clip();
      if (photoImage && photoImage.naturalWidth > 0) {
        drawCoverImage(ctx, photoImage, cx - r, cy - r, diameter, diameter, photoZoom, photoOffsetX, photoOffsetY);
      } else {
        ctx.fillStyle = isDarkTextStyle ? 'rgba(15, 29, 54, 0.8)' : 'rgba(241, 245, 249, 0.9)';
        ctx.fillRect(cx - r, cy - r, diameter, diameter);
        ctx.fillStyle = palette.textMuted;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.font = '800 34px "Noto Sans TC", sans-serif';
        ctx.fillText('📸 個人照片區', cx, cy - 22);
        ctx.font = '600 24px "Noto Sans TC", sans-serif';
        ctx.fillText('請由左側上傳照片', cx, cy + 28);
      }
      ctx.restore();

      ctx.beginPath();
      ctx.arc(cx, cy, r - 4, 0, Math.PI * 2);
      ctx.lineWidth = 10;
      ctx.strokeStyle = palette.goldMain;
      ctx.stroke();

      ctx.beginPath();
      ctx.arc(cx, cy, r - 15, 0, Math.PI * 2);
      ctx.lineWidth = 2.5;
      ctx.strokeStyle = palette.goldSoft;
      ctx.stroke();
    } else if (photoShape === 'polaroid') {
      const border = 26;
      const bottomLip = isPort ? 86 : 118;
      ctx.shadowColor = 'rgba(0,0,0,0.32)';
      ctx.shadowBlur = 34;
      ctx.shadowOffsetY = 14;
      drawRoundedRect(ctx, photoBox.x, photoBox.y, photoBox.w, photoBox.h, 12);
      ctx.fillStyle = '#ffffff';
      ctx.fill();
      ctx.shadowColor = 'transparent';
      ctx.lineWidth = 2;
      ctx.strokeStyle = 'rgba(184, 134, 11, 0.45)';
      ctx.stroke();

      const imgX = photoBox.x + border;
      const imgY = photoBox.y + border;
      const imgW = photoBox.w - border * 2;
      const imgH = photoBox.h - border - bottomLip;

      ctx.save();
      ctx.beginPath();
      ctx.rect(imgX, imgY, imgW, imgH);
      ctx.clip();
      if (photoImage && photoImage.naturalWidth > 0) {
        drawCoverImage(ctx, photoImage, imgX, imgY, imgW, imgH, photoZoom, photoOffsetX, photoOffsetY);
      } else {
        ctx.fillStyle = '#f1f5f9';
        ctx.fillRect(imgX, imgY, imgW, imgH);
        ctx.fillStyle = '#64748b';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.font = '800 34px "Noto Sans TC", sans-serif';
        ctx.fillText('📸 個人完賽照片', imgX + imgW / 2, imgY + imgH / 2 - 20);
        ctx.font = '600 24px "Noto Sans TC", sans-serif';
        ctx.fillText('請由左側上傳照片（可拖曳調整位置）', imgX + imgW / 2, imgY + imgH / 2 + 28);
      }
      ctx.restore();

      // 拍立得下方紀念字樣
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillStyle = '#334155';
      ctx.font = `800 ${isPort ? 26 : 30}px "Noto Serif TC", Georgia, serif`;
      ctx.fillText(`${metrics.name}  •  #${metrics.bib}`, photoBox.x + photoBox.w / 2, photoBox.y + photoBox.h - bottomLip * 0.58);
      ctx.fillStyle = '#b8860b';
      ctx.font = `700 ${isPort ? 20 : 22}px ui-monospace, monospace`;
      ctx.fillText(metrics.dateInfo.en, photoBox.x + photoBox.w / 2, photoBox.y + photoBox.h - bottomLip * 0.26);
    } else {
      // 圓角金邊相框（預設）
      ctx.shadowColor = 'rgba(0,0,0,0.28)';
      ctx.shadowBlur = 32;
      ctx.shadowOffsetY = 12;
      drawRoundedRect(ctx, photoBox.x, photoBox.y, photoBox.w, photoBox.h, 28);
      ctx.fillStyle = isDarkTextStyle ? '#0f172a' : '#f8fafc';
      ctx.fill();
      ctx.shadowColor = 'transparent';

      ctx.save();
      drawRoundedRect(ctx, photoBox.x + 8, photoBox.y + 8, photoBox.w - 16, photoBox.h - 16, 22);
      ctx.clip();
      if (photoImage && photoImage.naturalWidth > 0) {
        drawCoverImage(ctx, photoImage, photoBox.x + 8, photoBox.y + 8, photoBox.w - 16, photoBox.h - 16, photoZoom, photoOffsetX, photoOffsetY);
      } else {
        ctx.fillStyle = isDarkTextStyle ? 'rgba(15, 29, 54, 0.8)' : 'rgba(241, 245, 249, 0.9)';
        ctx.fillRect(photoBox.x + 8, photoBox.y + 8, photoBox.w - 16, photoBox.h - 16);
        ctx.strokeStyle = palette.goldSoft;
        ctx.lineWidth = 4;
        ctx.setLineDash([14, 10]);
        ctx.strokeRect(photoBox.x + 34, photoBox.y + 34, photoBox.w - 68, photoBox.h - 68);
        ctx.setLineDash([]);
        ctx.fillStyle = palette.textMuted;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.font = '800 34px "Noto Sans TC", sans-serif';
        ctx.fillText('📸 個人完賽照片區', photoBox.x + photoBox.w / 2, photoBox.y + photoBox.h / 2 - 22);
        ctx.font = '600 24px "Noto Sans TC", sans-serif';
        ctx.fillText('上傳照片後可直接拖曳調整構圖', photoBox.x + photoBox.w / 2, photoBox.y + photoBox.h / 2 + 28);
      }
      ctx.restore();

      drawRoundedRect(ctx, photoBox.x + 4, photoBox.y + 4, photoBox.w - 8, photoBox.h - 8, 25);
      ctx.lineWidth = 8;
      ctx.strokeStyle = palette.goldMain;
      ctx.stroke();
    }
    ctx.restore();

    // 5. 繪製成績證書主資訊面板（底圖存在時套用毛玻璃半透明底板確保 100% 清晰）
    ctx.save();
    drawRoundedRect(ctx, infoBox.x, infoBox.y, infoBox.w, infoBox.h, 28);
    ctx.fillStyle = palette.panelBg;
    ctx.fill();
    ctx.lineWidth = 2.5;
    ctx.strokeStyle = palette.panelBorder;
    ctx.stroke();
    ctx.restore();

    const cx = infoBox.x + infoBox.w / 2;
    const innerPad = 48;
    const textMaxW = infoBox.w - innerPad * 2;
    let curY = infoBox.y + (isPort ? 54 : 66);

    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    // 抬頭
    ctx.fillStyle = palette.turGreen;
    ctx.font = `700 ${isPort ? 28 : 28}px "Segoe UI", Arial, sans-serif`;
    ctx.fillText('OFFICIAL FINISHER CERTIFICATE', cx, curY);
    curY += isPort ? 62 : 66;

    ctx.fillStyle = palette.textPrimary;
    ctx.font = `900 ${isPort ? 76 : 80}px "Noto Serif TC", "Microsoft JhengHei", serif`;
    ctx.fillText('官 方 完 賽 證 書', cx, curY);
    curY += isPort ? 80 : 88;

    // 活動中英文名稱（大字）
    ctx.fillStyle = palette.textPrimary;
    fitTextFont(ctx, eventNameZh, textMaxW, isPort ? 54 : 58, 30, '900 {size}px "Noto Sans TC", "Microsoft JhengHei", sans-serif');
    ctx.fillText(eventNameZh, cx, curY);
    curY += isPort ? 54 : 58;

    if (eventNameEn) {
      ctx.fillStyle = palette.goldMain;
      fitTextFont(ctx, eventNameEn, textMaxW, isPort ? 30 : 32, 18, '700 {size}px Georgia, serif');
      ctx.fillText(eventNameEn, cx, curY);
      curY += isPort ? 44 : 46;
    }

    // 細金線
    ctx.strokeStyle = palette.divider;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(cx - textMaxW * 0.34, curY);
    ctx.lineTo(cx + textMaxW * 0.34, curY);
    ctx.stroke();
    curY += isPort ? 76 : 84;

    // 選手姓名
    ctx.fillStyle = palette.textPrimary;
    fitTextFont(ctx, metrics.name, textMaxW * 0.9, isPort ? 114 : 122, 58, '900 {size}px "Noto Serif TC", "Microsoft JhengHei", Georgia, serif');
    ctx.fillText(metrics.name, cx, curY);
    curY += isPort ? 74 : 78;

    // 編號與組別
    ctx.fillStyle = palette.textSecondary;
    ctx.font = `700 ${isPort ? 32 : 34}px "Noto Sans TC", "Segoe UI", sans-serif`;
    ctx.fillText(`BIB NO. ${metrics.bib}   |   ${metrics.categoryName}`, cx, curY);
    curY += isPort ? 46 : 50;

    // 2x2 成績資訊格
    const gX = infoBox.x + innerPad;
    const gW = textMaxW;
    const gH = isPort ? 300 : 320;
    ctx.save();
    drawRoundedRect(ctx, gX, curY, gW, gH, 20);
    ctx.fillStyle = palette.cardBg;
    ctx.fill();
    ctx.lineWidth = 2;
    ctx.strokeStyle = palette.panelBorder;
    ctx.stroke();

    const cellW = gW / 2;
    const cellH = (gH - 54) / 2;
    ctx.strokeStyle = palette.divider;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(gX + cellW, curY + 18);
    ctx.lineTo(gX + cellW, curY + gH - 60);
    ctx.moveTo(gX + 24, curY + cellH);
    ctx.lineTo(gX + gW - 24, curY + cellH);
    ctx.moveTo(gX + 24, curY + gH - 54);
    ctx.lineTo(gX + gW - 24, curY + gH - 54);
    ctx.stroke();

    const gridItems = [
      { zh: metrics.primaryLabelZh, en: metrics.primaryLabelEn, val: metrics.primaryValue, hi: true },
      { zh: metrics.secondaryLabelZh, en: metrics.secondaryLabelEn, val: metrics.secondaryValue, hi: false },
      { zh: metrics.thirdLabelZh, en: metrics.thirdLabelEn, val: metrics.thirdValue, hi: false },
      { zh: metrics.fourthLabelZh, en: metrics.fourthLabelEn, val: metrics.fourthValue, hi: false }
    ];
    gridItems.forEach((it, idx) => {
      const col = idx % 2;
      const r = Math.floor(idx / 2);
      const cellCx = gX + col * cellW + cellW / 2;
      const cellTop = curY + r * cellH;
      ctx.fillStyle = palette.textMuted;
      ctx.font = '700 24px "Noto Sans TC", sans-serif';
      ctx.fillText(`${it.zh} ${it.en}`, cellCx, cellTop + 34);
      ctx.fillStyle = it.hi ? palette.turGreen : palette.textPrimary;
      fitTextFont(ctx, it.val, cellW - 36, 54, 28, '900 {size}px ui-monospace, "Segoe UI", monospace');
      ctx.fillText(it.val, cellCx, cellTop + 86);
    });

    ctx.fillStyle = palette.textSecondary;
    ctx.font = '700 24px "Noto Sans TC", "Segoe UI", sans-serif';
    ctx.fillText(metrics.subNote, cx, curY + gH - 26);
    ctx.restore();
    curY += gH + 22;

    // 精簡分段成績條（若有且開啟）
    const splits = showSplits ? (metrics.splitItems || []).slice(0, 4) : [];
    if (splits.length > 0) {
      const spH = 112;
      ctx.save();
      drawRoundedRect(ctx, gX, curY, gW, spH, 16);
      ctx.fillStyle = palette.cardBg;
      ctx.fill();
      ctx.lineWidth = 1.5;
      ctx.strokeStyle = palette.panelBorder;
      ctx.stroke();

      const sw = gW / splits.length;
      ctx.strokeStyle = palette.divider;
      for (let i = 1; i < splits.length; i++) {
        ctx.beginPath();
        ctx.moveTo(gX + sw * i, curY + 14);
        ctx.lineTo(gX + sw * i, curY + spH - 14);
        ctx.stroke();
      }
      splits.forEach((sp, idx) => {
        const scx = gX + idx * sw + sw / 2;
        ctx.fillStyle = palette.textMuted;
        fitTextFont(ctx, sp.title, sw - 20, 21, 14, '700 {size}px "Noto Sans TC", sans-serif');
        ctx.fillText(sp.title, scx, curY + 30);
        ctx.fillStyle = palette.textPrimary;
        fitTextFont(ctx, sp.time, sw - 20, 32, 20, '900 {size}px ui-monospace, monospace');
        ctx.fillText(sp.time, scx, curY + 66);
        if (sp.sub) {
          ctx.fillStyle = palette.textMuted;
          ctx.font = '600 18px ui-monospace, sans-serif';
          ctx.fillText(sp.sub, scx, curY + 94);
        }
      });
      ctx.restore();
      curY += spH + 18;
    }

    // 選手個人感言／座右銘（選填）
    if (customQuote) {
      ctx.fillStyle = palette.goldMain;
      fitTextFont(ctx, `「${customQuote}」`, textMaxW, 30, 18, 'italic 700 {size}px "Noto Serif TC", Georgia, serif');
      ctx.fillText(`「${customQuote}」`, cx, curY + 20);
    }

    // 6. 面板底部官方落款列（左：活動日期｜中：金色鋼印｜右：TUR LOGO ＋ 主辦單位）
    const footY = infoBox.y + infoBox.h - (isPort ? 105 : 114);
    const fLeftX = infoBox.x + infoBox.w * 0.22;
    const fRightX = infoBox.x + infoBox.w * 0.78;

    ctx.textAlign = 'center';
    ctx.fillStyle = palette.textMuted;
    ctx.font = '700 20px "Noto Sans TC", sans-serif';
    ctx.fillText('活動日期 EVENT DATE', fLeftX, footY - 40);
    ctx.fillStyle = palette.textPrimary;
    ctx.font = '800 30px "Noto Sans TC", sans-serif';
    ctx.fillText(metrics.dateInfo.zh, fLeftX, footY - 2);
    ctx.fillStyle = palette.goldMain;
    ctx.font = '700 21px Georgia, serif';
    ctx.fillText(metrics.dateInfo.en, fLeftX, footY + 32);
    if (locationText) {
      ctx.fillStyle = palette.textSecondary;
      fitTextFont(ctx, locationText, infoBox.w * 0.36, 20, 14, '600 {size}px "Noto Sans TC", sans-serif');
      ctx.fillText(locationText, fLeftX, footY + 62);
    }

    if (showSeal) {
      drawOfficialSeal(ctx, cx, footY + 6, isPort ? 86 : 92, metrics.dateInfo.year, isDarkTextStyle);
    }

    if (logoImg && logoImg.naturalWidth > 0) {
      const ratio = logoImg.naturalWidth / logoImg.naturalHeight;
      const drawW = Math.min(72 * ratio, infoBox.w * 0.30);
      const drawH = drawW / ratio;
      const ly = footY - 50;
      if (isDarkTextStyle) {
        ctx.save();
        drawRoundedRect(ctx, fRightX - drawW / 2 - 14, ly - 8, drawW + 28, drawH + 16, 12);
        ctx.fillStyle = '#ffffff';
        ctx.fill();
        ctx.drawImage(logoImg, fRightX - drawW / 2, ly, drawW, drawH);
        ctx.restore();
      } else {
        ctx.save();
        ctx.globalCompositeOperation = 'multiply';
        ctx.drawImage(logoImg, fRightX - drawW / 2, ly, drawW, drawH);
        ctx.restore();
      }
    }
    ctx.fillStyle = palette.textPrimary;
    fitTextFont(ctx, orgName, infoBox.w * 0.36, 25, 16, '800 {size}px Georgia, serif');
    ctx.fillText(orgName, fRightX, footY + 46);

    return { width: W, height: H, metrics, photoBox };
  }

  function printCanvasCertificate(canvas, title = 'TUR 官方完賽證書') {
    const dataUrl = canvas.toDataURL('image/png');
    const win = global.open('', '_blank');
    if (!win) throw new Error('瀏覽器封鎖了列印視窗，請允許彈出視窗或改用「下載 PDF」列印');
    win.document.write(
      `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title>` +
      `<style>@page{margin:0;size:auto;}body{margin:0;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#fff;}img{max-width:100%;max-height:100vh;object-fit:contain;}</style>` +
      `</head><body><img src="${dataUrl}" onload="window.focus();window.print();" /></body></html>`
    );
    win.document.close();
  }

  /* ========================================================================
   * 純 JS 離線 PDF 1.4 產生器（支援單頁與多頁批次合併，免外部 CDN）
   * ======================================================================== */

  function dataUrlToUint8Array(dataUrl) {
    const base64 = String(dataUrl || '').split(',')[1] || '';
    const binary = global.atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  function canvasToJpegPage(canvas, quality = 0.92) {
    const dataUrl = canvas.toDataURL('image/jpeg', quality);
    return {
      jpegBytes: dataUrlToUint8Array(dataUrl),
      widthPx: canvas.width,
      heightPx: canvas.height
    };
  }

  function encodeAscii(str) {
    const arr = new Uint8Array(str.length);
    for (let i = 0; i < str.length; i++) arr[i] = str.charCodeAt(i) & 0xff;
    return arr;
  }

  function concatUint8Arrays(chunks) {
    const total = chunks.reduce((sum, c) => sum + c.length, 0);
    const out = new Uint8Array(total);
    let offset = 0;
    for (const c of chunks) {
      out.set(c, offset);
      offset += c.length;
    }
    return out;
  }

  /**
   * 將多張 JPEG 頁面打包為標準 PDF 1.4 檔案（A4 寬/高比例自動對齊）
   * @param {Array<{jpegBytes: Uint8Array, widthPx: number, heightPx: number}>} pages
   * @returns {Uint8Array}
   */
  function buildPdfFromJpegs(pages) {
    if (!Array.isArray(pages) || pages.length === 0) {
      throw new Error('沒有可供產生 PDF 的證書頁面');
    }

    // 物件配置：
    // 1: Catalog
    // 2: Pages
    // 每頁佔用 3 個物件：pageObjId, imgObjId, contentObjId
    const pageObjIds = [];
    const objects = []; // { id, chunks: Uint8Array[] }

    let nextId = 3;
    const pageSpecs = pages.map((p) => {
      const pageId = nextId++;
      const imgId = nextId++;
      const contentId = nextId++;
      pageObjIds.push(pageId);

      // 依原圖比例對應至 A4 尺寸（pt）：橫式以 841.89 pt 為寬，直式以 595.28 pt 為寬
      const isLandscape = p.widthPx >= p.heightPx;
      const pageW = isLandscape ? 841.89 : 595.28;
      const pageH = Number(((pageW * p.heightPx) / p.widthPx).toFixed(2));
      return { ...p, pageId, imgId, contentId, pageW, pageH };
    });

    objects.push({
      id: 1,
      chunks: [encodeAscii('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n')]
    });

    objects.push({
      id: 2,
      chunks: [
        encodeAscii(
          `2 0 obj\n<< /Type /Pages /Kids [${pageObjIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageObjIds.length} >>\nendobj\n`
        )
      ]
    });

    for (const spec of pageSpecs) {
      const streamCmd = `q\n${spec.pageW.toFixed(2)} 0 0 ${spec.pageH.toFixed(2)} 0 0 cm\n/Im0 Do\nQ\n`;
      const streamBytes = encodeAscii(streamCmd);

      objects.push({
        id: spec.pageId,
        chunks: [
          encodeAscii(
            `${spec.pageId} 0 obj\n` +
            `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${spec.pageW.toFixed(2)} ${spec.pageH.toFixed(2)}] ` +
            `/Resources << /XObject << /Im0 ${spec.imgId} 0 R >> >> /Contents ${spec.contentId} 0 R >>\nendobj\n`
          )
        ]
      });

      objects.push({
        id: spec.imgId,
        chunks: [
          encodeAscii(
            `${spec.imgId} 0 obj\n` +
            `<< /Type /XObject /Subtype /Image /Width ${spec.widthPx} /Height ${spec.heightPx} ` +
            `/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${spec.jpegBytes.length} >>\nstream\n`
          ),
          spec.jpegBytes,
          encodeAscii('\nendstream\nendobj\n')
        ]
      });

      objects.push({
        id: spec.contentId,
        chunks: [
          encodeAscii(`${spec.contentId} 0 obj\n<< /Length ${streamBytes.length} >>\nstream\n`),
          streamBytes,
          encodeAscii('endstream\nendobj\n')
        ]
      });
    }

    const finalChunks = [];
    // PDF 1.4 檔頭＋二進位標記註解
    const header = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34, 0x0a, 0x25, 0xff, 0xff, 0xff, 0xff, 0x0a]);
    finalChunks.push(header);
    let offset = header.length;

    const offsets = new Array(objects.length + 1).fill(0);
    for (const obj of objects) {
      offsets[obj.id] = offset;
      for (const ch of obj.chunks) {
        finalChunks.push(ch);
        offset += ch.length;
      }
    }

    const xrefOffset = offset;
    let xrefStr = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (let id = 1; id <= objects.length; id++) {
      xrefStr += `${String(offsets[id]).padStart(10, '0')} 00000 n \n`;
    }
    xrefStr += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;
    finalChunks.push(encodeAscii(xrefStr));

    return concatUint8Arrays(finalChunks);
  }

  function triggerBlobDownload(blob, fileName) {
    const url = global.URL.createObjectURL(blob);
    const a = global.document.createElement('a');
    a.href = url;
    a.download = fileName;
    global.document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => global.URL.revokeObjectURL(url), 1500);
  }

  async function downloadPng(canvas, fileName) {
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!blob) throw new Error('無法產生 PNG 圖檔');
    triggerBlobDownload(blob, fileName);
    return fileName;
  }

  function downloadPdfFromCanvas(canvas, fileName) {
    const page = canvasToJpegPage(canvas, 0.93);
    const pdfBytes = buildPdfFromJpegs([page]);
    const blob = new global.Blob([pdfBytes], { type: 'application/pdf' });
    triggerBlobDownload(blob, fileName);
    return fileName;
  }

  function downloadPdfFromPages(pages, fileName) {
    const pdfBytes = buildPdfFromJpegs(pages);
    const blob = new global.Blob([pdfBytes], { type: 'application/pdf' });
    triggerBlobDownload(blob, fileName);
    return fileName;
  }

  global.CertificateGenerator = {
    loadImage,
    formatBilingualEventDate,
    extractRunnerMetrics,
    getDefaultCustomLayout,
    loadSavedCustomConfig,
    saveCustomConfig,
    renderOfficialCertificate,
    renderCustomBgCertificate,
    renderPersonalizedCertificate,
    printCanvasCertificate,
    canvasToJpegPage,
    buildPdfFromJpegs,
    downloadPng,
    downloadPdfFromCanvas,
    downloadPdfFromPages
  };
})(window);
