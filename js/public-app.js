/**
 * 智慧路跑 NFC 晶片計時平台 v3.3 — 公開選手動態即時查詢頁面控制器
 * 支援「一般距離路跑」與「繞圈計時賽 (6H/12H/24H，依總里程排名)」雙模式專屬畫面
 */
(function () {
  'use strict';

  const DB = window.RaceDB;
  const TE = window.TimingEngine;
  const Sync = window.CloudSync;

  const state = {
    gasUrl: '',
    events: [],
    allRegistrations: [],
    allScans: [],
    selectedEventId: '',
    computed: null,
    catFilter: 'ALL',
    rankBy: 'gun',
    searchQuery: ''
  };

  function $(id) {
    return document.getElementById(id);
  }

  function esc(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  async function initPublicPage() {
    const params = new URLSearchParams(location.search);
    const paramApi = (params.get('api') || '').trim();
    const paramEvent = (params.get('event') || '').trim();
    const cfgApi = (window.APP_CONFIG && window.APP_CONFIG.DEFAULT_GAS_URL ? String(window.APP_CONFIG.DEFAULT_GAS_URL).trim() : '');
    const saved = DB.getDeviceSettings();

    // 優先順序：URL ?api= 參數 > js/config.js 共用網址 > 本機記憶網址
    state.gasUrl = paramApi || cfgApi || (saved.gasUrl || '').trim();
    state.selectedEventId = paramEvent || (saved.selectedEventId || '');

    bindEvents();
    await loadData(false);
    setInterval(() => loadData(true), 20000);
  }

  /**
   * 確保無論 GAS 回傳 v3 多活動格式、或舊版 v1/v2 單活動格式，都能正確轉換成標準活動清單
   */
  function normalizeCloudPayload(d) {
    if (!d) return { events: [], registrations: [], scans: [], activeEventId: '', isLegacy: false };
    if (Array.isArray(d.events) && d.events.length > 0) {
      return {
        events: d.events,
        registrations: Array.isArray(d.registrations) ? d.registrations : [],
        scans: Array.isArray(d.scans) ? d.scans : [],
        activeEventId: d.globalConfig?.activeEventId || d.events[0].eventId,
        isLegacy: false
      };
    }
    // 相容尚未升級至 v3 的舊版 GAS (v1/v2 回傳 config + categories + checkpoints + runners)
    if (d.config && (Array.isArray(d.categories) || Array.isArray(d.runners))) {
      const legacyId = 'EVT-LEGACY-01';
      return {
        events: [{
          eventId: legacyId,
          name: d.config.raceName || '路跑活動',
          date: d.config.raceDate || '',
          location: '',
          status: 'active',
          eventType: 'distance',
          lapDistanceKm: 0,
          categories: d.categories || [],
          checkpoints: d.checkpoints || []
        }],
        registrations: (d.runners || []).map((r) => ({ ...r, eventId: legacyId, regId: `${legacyId}__${r.bib}` })),
        scans: (d.scans || []).map((s) => ({ ...s, eventId: legacyId })),
        activeEventId: legacyId,
        isLegacy: true
      };
    }
    return {
      events: [],
      registrations: Array.isArray(d.registrations) ? d.registrations : [],
      scans: Array.isArray(d.scans) ? d.scans : [],
      activeEventId: '',
      isLegacy: false
    };
  }

  function ensureValidSelectedEvent(preferredDefaultId = '') {
    if (!state.events.length) {
      state.selectedEventId = '';
      return;
    }
    if (state.selectedEventId && state.events.some((e) => e.eventId === state.selectedEventId)) {
      return;
    }
    if (preferredDefaultId && state.events.some((e) => e.eventId === preferredDefaultId)) {
      state.selectedEventId = preferredDefaultId;
      return;
    }
    state.selectedEventId = state.events[0].eventId;
  }

  async function loadData(silent = false) {
    const badge = $('pub-source-badge');
    if (!silent) badge.textContent = '資料載入中...';

    try {
      if (state.gasUrl) {
        const res = await Sync.fetchPublicData(state.gasUrl);
        const norm = normalizeCloudPayload(res.data);

        if (norm.events.length > 0) {
          state.events = norm.events;
          state.allRegistrations = norm.registrations;
          state.allScans = norm.scans;
          ensureValidSelectedEvent(norm.activeEventId);
          badge.textContent = norm.isLegacy
            ? 'Google 雲端試算表即時同步（舊版 GAS 相容模式）'
            : `Google 雲端試算表即時同步（共 ${state.events.length} 場活動）`;

          // 背景快取至本機 IndexedDB，確保手機 PWA 斷網也能看到最新活動清單
          if (!norm.isLegacy && res.data) {
            DB.mergeCloudData(res.data).catch(() => {});
          }
        } else {
          // 雲端 Events 工作表目前無資料列 -> 回退本機快取並提示管理者同步
          await loadFromLocalIndexedDB();
          badge.textContent = state.events.length
            ? '雲端 Events 工作表尚無資料，顯示本機快取（請由後台點「一鍵完整雙向同步」）'
            : '雲端與本機尚無路跑活動（請由後台建立活動並同步）';
        }
      } else {
        await loadFromLocalIndexedDB();
        badge.textContent = '本機預覽模式（未指定雲端 API）';
      }
    } catch (err) {
      console.warn('雲端讀取失敗，切換本機快取:', err);
      await loadFromLocalIndexedDB();
      badge.textContent = `雲端連線中斷，顯示本機快取 (${err.message})`;
    }

    renderEventSelector();
    recomputeAndRender();
    $('pub-last-updated').textContent = `更新時間：${TE.formatClockSec(Date.now())}`;
  }

  async function loadFromLocalIndexedDB() {
    await DB.init();
    const [platform, regs, scans] = await Promise.all([
      DB.getPlatform(),
      DB.getAllRegistrations(),
      DB.getAllScans()
    ]);

    state.events = platform.events || [];
    ensureValidSelectedEvent(platform.activeEventId || '');

    state.allRegistrations = await Promise.all(
      regs.map(async (r) => ({
        regId: r.regId,
        eventId: r.eventId,
        bib: r.bib,
        name: r.name,
        idMasked: r.idMasked || TE.maskIdNumber(r.idNumber),
        idHash: r.idHash || (r.idNumber ? await TE.sha256Hex(r.idNumber) : ''),
        categoryId: r.categoryId,
        penaltySec: r.penaltySec || 0
      }))
    );
    state.allScans = scans.filter((s) => !s.deleted);
  }

  function getActiveEvent() {
    return state.events.find((e) => e.eventId === state.selectedEventId) || state.events[0] || null;
  }

  function renderEventSelector() {
    const sel = $('pub-event-select');
    const pills = $('pub-event-pills');
    if (state.events.length === 0) {
      sel.innerHTML = '<option value="">尚無路跑活動（請確認後台已建立並同步）</option>';
      if (pills) pills.innerHTML = '';
      return;
    }

    ensureValidSelectedEvent();
    const statusTag = { active: '進行中', upcoming: '籌備中', finished: '已結束' };
    sel.innerHTML = state.events
      .map((ev) => {
        const typeTag = ev.eventType === 'lap_timed' ? `🔁繞圈(${ev.lapDistanceKm}K/圈)` : '🏃距離賽';
        return `<option value="${esc(ev.eventId)}" style="background-color:#0f172a;color:#f8fafc;">${typeTag} ${esc(ev.name)} (${esc(ev.date || '')}) [${statusTag[ev.status] || ev.status}]</option>`;
      })
      .join('');
    sel.value = state.selectedEventId;

    // 同步顯示一鍵活動快捷按鈕，解決部分手機原生下拉選單不易辨識的問題
    if (pills) {
      pills.innerHTML = state.events
        .map((ev) => {
          const active = ev.eventId === state.selectedEventId;
          const icon = ev.eventType === 'lap_timed' ? '🔁' : '🏃';
          return `<button type="button" data-pill-event="${esc(ev.eventId)}"
            class="px-3 py-1.5 rounded-xl text-xs font-black transition ${
              active
                ? 'bg-sky-500 text-slate-950 shadow'
                : 'bg-slate-900 hover:bg-slate-800 border border-slate-700 text-slate-300'
            }">${icon} ${esc(ev.name)}</button>`;
        })
        .join('');
    }
  }

  function recomputeAndRender() {
    const evt = getActiveEvent();
    if (!evt) return;

    const eventRegs = state.allRegistrations.filter((r) => r.eventId === evt.eventId);
    const eventScans = state.allScans.filter((s) => s.eventId === evt.eventId && !s.deleted);

    state.computed = TE.computeEventResults({
      event: evt,
      runners: eventRegs,
      scans: eventScans,
      rankBy: state.rankBy
    });

    const isLap = evt.eventType === 'lap_timed';
    $('pub-race-name').textContent = (isLap ? '🔁 ' : '') + evt.name;
    $('pub-race-meta').textContent = isLap
      ? `${evt.date || '--'}${evt.location ? ` ｜ ${evt.location}` : ''} ｜ 繞圈計時賽（單圈 ${evt.lapDistanceKm} km • 依累積總里程排名）`
      : `${evt.date || '--'}${evt.location ? ` ｜ ${evt.location}` : ''} ｜ 活動代碼：${evt.eventId}`;

    // 繞圈賽依里程排名，不需 Gun/Chip 排序切換
    $('pub-rankby').classList.toggle('hidden', isLap);
    $('pub-lb-title').textContent = isLap ? '🏆 繞圈計時賽即時里程排行榜' : '🏆 本場即時成績排行榜';

    renderOverview(evt);
    renderLeaderboard(evt);
    if (state.searchQuery) runSearch(state.searchQuery);
  }

  function renderOverview(evt) {
    const s = state.computed.summary;
    const isLap = evt.eventType === 'lap_timed';

    if (isLap) {
      $('pub-kpi-label-1').textContent = '本場報名人數';
      $('pub-kpi-label-2').textContent = '全場累積總圈數';
      $('pub-kpi-label-3').textContent = '全場累積總里程 (km)';
      $('pub-kpi-label-4').textContent = '目前最高里程 (km)';

      $('pub-kpi-total').textContent = String(s.totalRunners);
      $('pub-kpi-finished').textContent = `${s.totalLapsAll} 圈`;
      $('pub-kpi-oncourse').textContent = `${s.totalKmAll.toFixed(2)}`;
      $('pub-kpi-notstarted').textContent = `${s.topDistanceKm.toFixed(2)}`;

      $('pub-cat-card-title').textContent = '⏱️ 各限時組別計時進度與領先里程';
      $('pub-right-card-title').textContent = '🏅 目前總里程領先榜 (TOP 6)';

      $('pub-category-stats').innerHTML = !s.byCategory.length
        ? '<div class="text-xs text-slate-500">尚無組別</div>'
        : s.byCategory
            .map((c) => {
              const remainText = c.gunTimeMs === null
                ? '尚未鳴槍'
                : c.remainingMs > 0
                  ? `剩餘 ${TE.formatDuration(c.remainingMs, false)}`
                  : '限時已結束';
              return `
                <div class="p-3 rounded-xl bg-slate-950 border border-slate-800 space-y-1">
                  <div class="flex items-center justify-between text-xs">
                    <span class="font-black text-white flex items-center gap-1.5">
                      <span class="w-2.5 h-2.5 rounded-full" style="background:${esc(c.color)}"></span>
                      ${esc(c.name)}（限時 ${c.durationHours} 小時）
                    </span>
                    <span class="font-mono-num font-bold ${c.remainingMs > 0 ? 'text-amber-300' : 'text-slate-400'}">${remainText}</span>
                  </div>
                  <div class="flex items-center justify-between text-[11px] font-mono-num text-slate-300">
                    <span>參賽 ${c.total} 人 ｜ 累積 ${c.sumLaps} 圈</span>
                    <span>領先 <strong class="text-emerald-400">${c.topKm.toFixed(2)} km</strong> ｜ 平均 ${c.avgKm.toFixed(2)} km</span>
                  </div>
                </div>
              `;
            })
            .join('');

      const leaders = s.topLeaders || [];
      const maxKm = Math.max(1, s.topDistanceKm || 1);
      $('pub-checkpoint-stats').innerHTML = !leaders.length
        ? '<div class="text-xs text-slate-500">尚無選手完成圈數</div>'
        : leaders
            .map((r) => {
              const pct = Math.min(100, Math.round((r.totalDistanceKm / maxKm) * 100));
              return `
                <div class="space-y-1">
                  <div class="flex items-center justify-between text-xs">
                    <span class="font-bold text-white">#${r.overallRank} ${esc(r.name)} <span class="text-slate-400 font-mono-num">(${esc(r.bib)} • ${esc(r.categoryName)})</span></span>
                    <span class="font-mono-num font-black text-emerald-400">${r.totalDistanceKm.toFixed(2)} km (${r.lapCount} 圈)</span>
                  </div>
                  <div class="h-2 rounded-full bg-slate-800 overflow-hidden">
                    <div class="h-full rounded-full transition-all duration-500" style="width:${pct}%; background:${esc(r.categoryColor)}"></div>
                  </div>
                </div>
              `;
            })
            .join('');
    } else {
      $('pub-kpi-label-1').textContent = '本場報名人數';
      $('pub-kpi-label-2').textContent = '已完賽人數';
      $('pub-kpi-label-3').textContent = '賽道上人數';
      $('pub-kpi-label-4').textContent = '未出發';

      $('pub-kpi-total').textContent = String(s.totalRunners);
      $('pub-kpi-finished').textContent = String(s.finishedCount);
      $('pub-kpi-oncourse').textContent = String(s.onCourseCount);
      $('pub-kpi-notstarted').textContent = String(s.notStartedCount);

      $('pub-cat-card-title').textContent = '🏃 各分組完賽進度';
      $('pub-right-card-title').textContent = '📍 各補給站通過人數統計';

      $('pub-category-stats').innerHTML = !s.byCategory.length
        ? '<div class="text-xs text-slate-500">本場活動尚無分組</div>'
        : s.byCategory
            .map((c) => {
              const pct = c.total > 0 ? Math.round((c.finished / c.total) * 100) : 0;
              return `
                <div class="space-y-1.5">
                  <div class="flex items-center justify-between text-xs">
                    <span class="font-bold text-white flex items-center gap-1.5">
                      <span class="w-2.5 h-2.5 rounded-full" style="background:${esc(c.color)}"></span>
                      ${esc(c.name)}
                    </span>
                    <span class="font-mono-num text-slate-300">
                      完賽 <strong class="text-emerald-400">${c.finished}</strong> /
                      賽道 <strong class="text-amber-400">${c.onCourse}</strong> /
                      總數 ${c.total} (${pct}%)
                    </span>
                  </div>
                  <div class="h-2.5 rounded-full bg-slate-800 overflow-hidden">
                    <div class="h-full rounded-full transition-all duration-500" style="width:${pct}%; background:${esc(c.color)}"></div>
                  </div>
                </div>
              `;
            })
            .join('');

      const maxPassed = Math.max(1, s.totalRunners);
      $('pub-checkpoint-stats').innerHTML = !s.byCheckpoint.length
        ? '<div class="text-xs text-slate-500">本場活動尚無站點</div>'
        : s.byCheckpoint
            .map((cp) => {
              const pct = Math.min(100, Math.round((cp.passedCount / maxPassed) * 100));
              return `
                <div class="space-y-1">
                  <div class="flex items-center justify-between text-xs">
                    <span class="font-bold text-slate-200">#${cp.order} ${esc(cp.name)}</span>
                    <span class="font-mono-num font-bold text-sky-400">${cp.passedCount} 人通過</span>
                  </div>
                  <div class="h-2 rounded-full bg-slate-800 overflow-hidden">
                    <div class="h-full rounded-full bg-sky-500 transition-all duration-500" style="width:${pct}%"></div>
                  </div>
                </div>
              `;
            })
            .join('');
    }

    const catSelect = $('pub-cat-filter');
    const cats = evt.categories || [];
    catSelect.innerHTML =
      '<option value="ALL">全部組別</option>' +
      cats.map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    if (!cats.some((c) => c.id === state.catFilter)) state.catFilter = 'ALL';
    catSelect.value = state.catFilter;
  }

  async function runSearch(rawQuery) {
    const q = String(rawQuery || '').trim();
    state.searchQuery = q;
    const container = $('pub-search-results');

    if (!q || !state.computed) {
      container.classList.add('hidden');
      container.innerHTML = '';
      return;
    }

    const qLower = q.toLowerCase();
    const qHash = await TE.sha256Hex(q.toUpperCase());

    const matches = state.computed.rows.filter((row) => {
      if (row.bib.toLowerCase().includes(qLower)) return true;
      if (row.name.toLowerCase().includes(qLower)) return true;
      if (row.idHash && row.idHash === qHash) return true;
      if (row.idMasked && row.idMasked.toLowerCase() === qLower) return true;
      return false;
    });

    container.classList.remove('hidden');

    if (matches.length === 0) {
      const otherEvtIds = new Set(
        state.allRegistrations
          .filter((r) => r.eventId !== state.selectedEventId && (
            r.name.toLowerCase().includes(qLower) || (r.idHash && r.idHash === qHash)
          ))
          .map((r) => r.eventId)
      );
      const otherHint = otherEvtIds.size > 0
        ? `<div class="mt-2 flex flex-wrap justify-center gap-2">${Array.from(otherEvtIds).map((eid) => {
            const ev = state.events.find((x) => x.eventId === eid);
            return `<button type="button" data-switch-pub-event="${esc(eid)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-sky-600 hover:bg-sky-500 text-white">切換至「${esc(ev ? ev.name : eid)}」查看 →</button>`;
          }).join('')}</div>`
        : '';

      container.innerHTML = `
        <div class="rounded-2xl bg-slate-950 border border-slate-800 p-5 text-center text-sm text-slate-400">
          本場活動查無符合「<strong class="text-white">${esc(q)}</strong>」的選手資料。
          ${otherHint}
        </div>
      `;
      return;
    }

    const isLap = state.computed.eventType === 'lap_timed';

    if (isLap) {
      container.innerHTML = matches.slice(0, 10).map((row) => renderLapSearchCard(row)).join('');
    } else {
      container.innerHTML = matches.slice(0, 10).map((row) => renderDistanceSearchCard(row)).join('');
    }
  }

  function renderOtherEventChips(row) {
    const otherRegs = row.idHash
      ? state.allRegistrations.filter((r) => r.idHash === row.idHash && r.eventId !== state.selectedEventId)
      : [];
    if (!otherRegs.length) return '';
    return `<div class="flex flex-wrap items-center gap-1.5 text-xs pt-1">
      <span class="text-slate-400">同晶片參加其他場次：</span>
      ${otherRegs.map((or) => {
        const ev = state.events.find((e) => e.eventId === or.eventId);
        return `<button type="button" data-switch-pub-event="${esc(or.eventId)}" class="px-2 py-0.5 rounded bg-slate-800 hover:bg-sky-700 text-sky-300 font-bold">${esc(ev ? ev.name : or.eventId)} (${esc(or.categoryId)})</button>`;
      }).join('')}
    </div>`;
  }

  function renderLapSearchCard(row) {
    const lapRowsHtml = row.laps.length
      ? row.laps.map((l) => `
          <tr class="${l.overtime ? 'text-rose-400/80' : 'text-slate-200'}">
            <td class="py-1.5 pr-3 font-mono-num font-bold">${l.overtime ? '超時' : `第 ${l.lapNumber} 圈`}</td>
            <td class="py-1.5 pr-3 font-mono-num text-sky-300">${TE.formatClockMs(l.timestampMs)}</td>
            <td class="py-1.5 pr-3 font-mono-num">${l.lapDurationMs !== null ? TE.formatDuration(l.lapDurationMs, false) : '--'}</td>
            <td class="py-1.5 pr-3 font-mono-num text-slate-400">${esc(l.lapPace)}</td>
            <td class="py-1.5 text-right font-mono-num font-bold text-emerald-400">${l.cumulativeKm !== null ? `${l.cumulativeKm.toFixed(2)} km` : '不計'}</td>
          </tr>
        `).join('')
      : '<tr><td colspan="5" class="py-4 text-center text-slate-500">尚未完成任何圈數</td></tr>';

    return `
      <div class="rounded-2xl bg-slate-950 border border-purple-500/40 p-4 sm:p-5 space-y-4">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <div>
            <div class="flex items-center gap-2 flex-wrap">
              <span class="font-mono-num text-lg font-black text-amber-400">${esc(row.bib)}</span>
              <span class="text-xl font-black text-white">${esc(row.name)}</span>
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-purple-500/20 text-purple-300 border border-purple-500/40">
                🔁 ${esc(row.categoryName)} (${row.durationHours}H)
              </span>
            </div>
            <div class="text-xs text-slate-400 mt-1">
              單圈 ${row.lapDistanceKm} km ${row.idMasked ? ` ｜ 身份證：<code class="text-sky-300 font-mono-num">${esc(row.idMasked)}</code>` : ''}
            </div>
            ${renderOtherEventChips(row)}
          </div>
          <div class="text-right">
            <div class="text-[11px] text-slate-400">最後一圈通過時間</div>
            <div class="text-sm font-mono-num font-bold text-sky-300">${row.lastLapTimeMs ? TE.formatClockSec(row.lastLapTimeMs) : '尚未感應'}</div>
          </div>
        </div>

        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div class="bg-slate-900 border border-purple-500/30 rounded-xl p-3">
            <div class="text-[11px] text-purple-300 font-bold">累積圈數</div>
            <div class="text-2xl font-black font-mono-num text-white mt-0.5">${row.lapCount} 圈</div>
          </div>
          <div class="bg-slate-900 border border-emerald-500/30 rounded-xl p-3">
            <div class="text-[11px] text-emerald-400 font-bold">累積總里程</div>
            <div class="text-2xl font-black font-mono-num text-emerald-400 mt-0.5">${row.totalDistanceKm.toFixed(2)} km</div>
          </div>
          <div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
            <div class="text-[11px] text-slate-400 font-bold">總名次 / 分組名次</div>
            <div class="text-lg font-black font-mono-num text-amber-400 mt-0.5">
              ${row.overallRank ? `總 #${row.overallRank}` : '--'} / ${row.categoryRank ? `組 #${row.categoryRank}` : '--'}
            </div>
          </div>
          <div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
            <div class="text-[11px] text-slate-400 font-bold">平均單圈 / 配速</div>
            <div class="text-sm font-black font-mono-num text-sky-300 mt-1">
              ${row.avgLapMs !== null ? TE.formatDuration(row.avgLapMs, false) : '--'} • ${esc(row.avgPace)}
            </div>
          </div>
        </div>

        <div class="pt-1 space-y-2">
          <div class="text-xs font-extrabold text-slate-300">🔁 逐圈感應成績明細</div>
          <div class="overflow-x-auto max-h-60 border border-slate-800 rounded-xl px-3">
            <table class="w-full text-xs text-left whitespace-nowrap">
              <thead class="text-slate-400 border-b border-slate-800 sticky top-0 bg-slate-950">
                <tr>
                  <th class="py-2 pr-3">圈次</th>
                  <th class="py-2 pr-3">通過時間</th>
                  <th class="py-2 pr-3">單圈耗時</th>
                  <th class="py-2 pr-3">單圈配速</th>
                  <th class="py-2 text-right">累積里程</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-800/60">${lapRowsHtml}</tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  }

  function renderDistanceSearchCard(row) {
    const checkpoints = state.computed.checkpoints;
    const statusBanner = {
      finished: { text: '🏁 已完賽', cls: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40' },
      on_course: { text: '🏃 賽道進行中', cls: 'bg-amber-500/20 text-amber-300 border-amber-500/40' },
      not_started: { text: '⏳ 尚未出發', cls: 'bg-slate-800 text-slate-300 border-slate-700' }
    };
    const st = statusBanner[row.status] || statusBanner.not_started;
    const lastLoc = row.lastCheckpoint
      ? `${row.lastCheckpoint.name}（${TE.formatClockSec(row.lastScan.timestampMs)}）`
      : '尚未通過任何感應點';

    const etaBlock =
      row.status === 'on_course' && row.etaTimestampMs
        ? `<div class="bg-amber-950/40 border border-amber-500/40 rounded-xl p-3">
             <div class="text-[11px] text-amber-300 font-bold">預估完賽時間 (ETA)</div>
             <div class="text-lg font-black font-mono-num text-amber-300 mt-0.5">${TE.formatClockSec(row.etaTimestampMs)}</div>
             <div class="text-[11px] text-amber-200/80 font-mono-num">預估總耗時 ${TE.formatDuration(row.etaDurationMs, false)}</div>
           </div>`
        : `<div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
             <div class="text-[11px] text-slate-400 font-bold">總名次 / 分組名次</div>
             <div class="text-lg font-black font-mono-num text-emerald-400 mt-0.5">
               ${row.overallRank ? `總 #${row.overallRank}` : '--'} / ${row.categoryRank ? `組 #${row.categoryRank}` : '--'}
             </div>
             <div class="text-[11px] text-slate-400 font-mono-num">平均配速 ${esc(row.avgPace)}</div>
           </div>`;

    const timelineHtml = checkpoints
      .map((cp) => {
        const scan = row.cpScanMap[cp.id];
        const split = row.splits.find((sp) => sp.toId === cp.id);
        const passed = Boolean(scan);
        return `
          <div class="flex items-start gap-3">
            <div class="mt-1 w-3.5 h-3.5 rounded-full shrink-0 ${passed ? 'bg-emerald-400 ring-4 ring-emerald-400/20' : 'bg-slate-700'}"></div>
            <div class="flex-1 flex flex-wrap items-center justify-between gap-1 pb-2.5 border-b border-slate-800/60">
              <div>
                <div class="text-xs font-bold ${passed ? 'text-white' : 'text-slate-500'}">${esc(cp.name)}</div>
                ${split && split.durationMs !== null ? `<div class="text-[11px] font-mono-num text-sky-400">分段耗時 ${TE.formatDuration(split.durationMs, false)} • 配速 ${esc(split.pace)}</div>` : ''}
              </div>
              <div class="font-mono-num text-xs font-bold ${passed ? 'text-emerald-300' : 'text-slate-600'}">
                ${passed ? TE.formatClockMs(scan.timestampMs) : '未通過'}
              </div>
            </div>
          </div>
        `;
      })
      .join('');

    return `
      <div class="rounded-2xl bg-slate-950 border border-slate-700 p-4 sm:p-5 space-y-4">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <div>
            <div class="flex items-center gap-2 flex-wrap">
              <span class="font-mono-num text-lg font-black text-amber-400">${esc(row.bib)}</span>
              <span class="text-xl font-black text-white">${esc(row.name)}</span>
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold border ${st.cls}">${st.text}</span>
            </div>
            <div class="text-xs text-slate-400 mt-1">
              分組：<strong class="text-slate-200">${esc(row.categoryName)}</strong>
              ${row.idMasked ? ` ｜ 身份證：<code class="text-sky-300 font-mono-num">${esc(row.idMasked)}</code>` : ''}
            </div>
            ${renderOtherEventChips(row)}
          </div>
          <div class="text-right">
            <div class="text-[11px] text-slate-400">最後通過位置</div>
            <div class="text-xs sm:text-sm font-bold text-sky-300">${esc(lastLoc)}</div>
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
            <div class="text-[11px] text-slate-400 font-bold">大會時間 (Gun Time)</div>
            <div class="text-lg font-black font-mono-num text-white mt-0.5">
              ${row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, true) : '--:--:--.---'}
            </div>
          </div>
          <div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
            <div class="text-[11px] text-slate-400 font-bold">個人淨時間 (Chip Time)</div>
            <div class="text-lg font-black font-mono-num text-sky-300 mt-0.5">
              ${row.chipDurationMs !== null ? TE.formatDuration(row.chipDurationMs, true) : '--:--:--.---'}
            </div>
          </div>
          ${etaBlock}
        </div>

        <div class="pt-1 space-y-2">
          <div class="text-xs font-extrabold text-slate-300">📍 各感應站點與分段成績紀錄</div>
          <div class="space-y-2">${timelineHtml}</div>
        </div>
      </div>
    `;
  }

  function renderLeaderboard(evt) {
    if (!state.computed) return;
    const isLap = evt.eventType === 'lap_timed';
    const rows = state.computed.rows.filter((r) => state.catFilter === 'ALL' || r.categoryId === state.catFilter);
    const thead = $('pub-leaderboard-thead');
    const tbody = $('pub-leaderboard-tbody');

    if (isLap) {
      thead.innerHTML = `
        <tr>
          <th class="py-2.5 pr-3">總名次</th>
          <th class="py-2.5 pr-3">組名次</th>
          <th class="py-2.5 pr-3">編號</th>
          <th class="py-2.5 pr-3">姓名</th>
          <th class="py-2.5 pr-3">身份證 (隱碼)</th>
          <th class="py-2.5 pr-3">限時組別</th>
          <th class="py-2.5 pr-3">累積圈數</th>
          <th class="py-2.5 pr-3">總里程 (km)</th>
          <th class="py-2.5 pr-3">累積耗時</th>
          <th class="py-2.5 pr-3">平均單圈</th>
          <th class="py-2.5">平均配速</th>
        </tr>
      `;
      if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="11" class="py-8 text-center text-slate-500">本場活動尚無繞圈成績資料</td></tr>';
        return;
      }
      tbody.innerHTML = rows
        .map((row) => `
          <tr class="hover:bg-slate-800/40 cursor-pointer" data-pick-bib="${esc(row.bib)}">
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-400">${row.overallRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-sky-300">${row.categoryRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-white">${esc(row.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(row.name)}</td>
            <td class="py-2.5 pr-3 font-mono-num text-slate-400">${esc(row.idMasked || '--')}</td>
            <td class="py-2.5 pr-3">
              <span class="inline-flex items-center gap-1.5 text-xs font-bold">
                <span class="w-2 h-2 rounded-full" style="background:${esc(row.categoryColor)}"></span>
                ${esc(row.categoryName)}
              </span>
            </td>
            <td class="py-2.5 pr-3 font-mono-num font-black text-purple-300">${row.lapCount} 圈</td>
            <td class="py-2.5 pr-3 font-mono-num font-black text-emerald-400">${row.totalDistanceKm.toFixed(2)} km</td>
            <td class="py-2.5 pr-3 font-mono-num text-sky-300">${row.totalElapsedMs !== null ? TE.formatDuration(row.totalElapsedMs, false) : '--'}</td>
            <td class="py-2.5 pr-3 font-mono-num text-slate-300">${row.avgLapMs !== null ? TE.formatDuration(row.avgLapMs, false) : '--'}</td>
            <td class="py-2.5 font-mono-num text-slate-300">${esc(row.avgPace)}</td>
          </tr>
        `)
        .join('');
      return;
    }

    thead.innerHTML = `
      <tr>
        <th class="py-2.5 pr-3">總名次</th>
        <th class="py-2.5 pr-3">組名次</th>
        <th class="py-2.5 pr-3">編號</th>
        <th class="py-2.5 pr-3">姓名</th>
        <th class="py-2.5 pr-3">身份證 (隱碼)</th>
        <th class="py-2.5 pr-3">組別</th>
        <th class="py-2.5 pr-3">最新動態 / 站點</th>
        <th class="py-2.5 pr-3">大會時間</th>
        <th class="py-2.5 pr-3">淨時間</th>
        <th class="py-2.5">平均配速 / ETA</th>
      </tr>
    `;

    if (rows.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" class="py-8 text-center text-slate-500">本場活動尚無選手報名或成績資料</td></tr>';
      return;
    }

    tbody.innerHTML = rows
      .map((row) => {
        const locStr =
          row.status === 'finished'
            ? '<span class="text-emerald-400 font-bold">🏁 已完賽</span>'
            : row.lastCheckpoint
              ? `<span class="text-amber-300 font-bold">🏃 ${esc(row.lastCheckpoint.name)}</span>`
              : '<span class="text-slate-500">未出發</span>';

        const paceOrEta =
          row.status === 'finished'
            ? esc(row.avgPace)
            : row.etaTimestampMs
              ? `<span class="text-amber-300">預估 ${TE.formatClockSec(row.etaTimestampMs)}</span>`
              : '--';

        return `
          <tr class="hover:bg-slate-800/40 cursor-pointer" data-pick-bib="${esc(row.bib)}">
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-400">${row.overallRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-sky-300">${row.categoryRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-white">${esc(row.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(row.name)}</td>
            <td class="py-2.5 pr-3 font-mono-num text-slate-400">${esc(row.idMasked || '--')}</td>
            <td class="py-2.5 pr-3">
              <span class="inline-flex items-center gap-1.5 text-xs font-bold">
                <span class="w-2 h-2 rounded-full" style="background:${esc(row.categoryColor)}"></span>
                ${esc(row.categoryName)}
              </span>
            </td>
            <td class="py-2.5 pr-3">${locStr}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-emerald-400">
              ${row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, true) : '--'}
            </td>
            <td class="py-2.5 pr-3 font-mono-num text-sky-300">
              ${row.chipDurationMs !== null ? TE.formatDuration(row.chipDurationMs, true) : '--'}
            </td>
            <td class="py-2.5 font-mono-num text-slate-300">${paceOrEta}</td>
          </tr>
        `;
      })
      .join('');
  }

  function bindEvents() {
    $('btn-pub-refresh').addEventListener('click', () => loadData(false));

    $('pub-event-select').addEventListener('change', (e) => {
      state.selectedEventId = e.target.value;
      state.catFilter = 'ALL';
      renderEventSelector();
      recomputeAndRender();
    });

    $('pub-event-pills').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-pill-event]');
      if (!btn) return;
      state.selectedEventId = btn.dataset.pillEvent;
      state.catFilter = 'ALL';
      renderEventSelector();
      recomputeAndRender();
    });

    $('form-pub-search').addEventListener('submit', (e) => {
      e.preventDefault();
      runSearch($('pub-search-input').value);
    });

    $('pub-search-input').addEventListener('input', (e) => {
      if (!e.target.value.trim()) runSearch('');
    });

    $('pub-cat-filter').addEventListener('change', (e) => {
      state.catFilter = e.target.value;
      renderLeaderboard(getActiveEvent());
    });

    $('pub-rankby').addEventListener('change', (e) => {
      state.rankBy = e.target.value;
      recomputeAndRender();
    });

    $('pub-leaderboard-tbody').addEventListener('click', (e) => {
      const tr = e.target.closest('[data-pick-bib]');
      if (!tr) return;
      $('pub-search-input').value = tr.dataset.pickBib;
      runSearch(tr.dataset.pickBib);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    $('pub-search-results').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-switch-pub-event]');
      if (!btn) return;
      state.selectedEventId = btn.dataset.switchPubEvent;
      state.catFilter = 'ALL';
      renderEventSelector();
      recomputeAndRender();
    });
  }

  document.addEventListener('DOMContentLoaded', initPublicPage);
})();
