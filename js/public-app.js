/**
 * 智慧路跑 NFC 晶片計時系統 — 公開選手動態即時查詢頁面控制器 (免登入)
 * 資料來源優先順序：
 * 1. URL 參數 ?api=<GAS_URL>
 * 2. 同源裝置設定的 gasUrl
 * 3. 本機 IndexedDB 預覽模式（無需網路即可即時展示或測試）
 */
(function () {
  'use strict';

  const DB = window.RaceDB;
  const TE = window.TimingEngine;
  const Sync = window.CloudSync;

  const state = {
    gasUrl: '',
    isCloud: false,
    config: null,
    categories: [],
    checkpoints: [],
    runners: [],
    scans: [],
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
    const savedApi = (DB.getDeviceSettings().gasUrl || '').trim();
    state.gasUrl = paramApi || savedApi;

    bindEvents();
    await loadData();

    // 每 20 秒自動刷新最新賽道動態
    setInterval(() => loadData(true), 20000);
  }

  async function loadData(silent = false) {
    const badge = $('pub-source-badge');
    if (!silent) badge.textContent = '資料載入中...';

    try {
      if (state.gasUrl) {
        const res = await Sync.fetchPublicData(state.gasUrl);
        const d = res.data;
        state.isCloud = true;
        state.config = d.config;
        state.categories = d.categories || [];
        state.checkpoints = d.checkpoints || [];
        state.runners = d.runners || [];
        state.scans = d.scans || [];
        badge.textContent = 'Google 雲端試算表即時同步';
      } else {
        await loadFromLocalIndexedDB();
        badge.textContent = '本機預覽模式（未指定雲端 API）';
      }
    } catch (err) {
      console.warn('雲端讀取失敗，切換回本機預覽資料:', err);
      await loadFromLocalIndexedDB();
      badge.textContent = `雲端連線中斷，顯示本機快取 (${err.message})`;
    }

    recomputeAndRender();
    $('pub-last-updated').textContent = `更新時間：${TE.formatClockSec(Date.now())}`;
  }

  async function loadFromLocalIndexedDB() {
    await DB.init();
    const [cfg, runners, scans] = await Promise.all([
      DB.getConfig(),
      DB.getAllRunners(),
      DB.getAllScans()
    ]);

    state.isCloud = false;
    state.config = {
      raceName: cfg.raceName,
      raceDate: cfg.raceDate,
      rankBy: cfg.rankBy || 'gun'
    };
    state.categories = cfg.categories || [];
    state.checkpoints = cfg.checkpoints || [];
    // 公開頁面嚴格隱碼身份證字號，僅保留 idMasked 與 idHash
    state.runners = await Promise.all(
      runners.map(async (r) => ({
        bib: r.bib,
        name: r.name,
        idMasked: r.idMasked || TE.maskIdNumber(r.idNumber),
        idHash: r.idHash || (r.idNumber ? await TE.sha256Hex(r.idNumber) : ''),
        categoryId: r.categoryId,
        penaltySec: r.penaltySec || 0
      }))
    );
    state.scans = scans.filter((s) => !s.deleted);
  }

  function recomputeAndRender() {
    state.computed = TE.computeResults({
      runners: state.runners,
      scans: state.scans,
      categories: state.categories,
      checkpoints: state.checkpoints,
      rankBy: state.rankBy
    });

    $('pub-race-name').textContent = state.config?.raceName || '智慧路跑即時成績查詢';
    $('pub-race-date').textContent = state.config?.raceDate ? `賽事日期：${state.config.raceDate}` : '';

    renderOverview();
    renderLeaderboard();
    if (state.searchQuery) {
      runSearch(state.searchQuery);
    }
  }

  function renderOverview() {
    const s = state.computed.summary;
    $('pub-kpi-total').textContent = String(s.totalRunners);
    $('pub-kpi-finished').textContent = String(s.finishedCount);
    $('pub-kpi-oncourse').textContent = String(s.onCourseCount);
    $('pub-kpi-notstarted').textContent = String(s.notStartedCount);

    // 各分組完賽進度條
    const catContainer = $('pub-category-stats');
    if (!s.byCategory.length) {
      catContainer.innerHTML = '<div class="text-xs text-slate-500">尚無分組資料</div>';
    } else {
      catContainer.innerHTML = s.byCategory
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
    }

    // 各站點通過人數統計
    const cpContainer = $('pub-checkpoint-stats');
    const maxPassed = Math.max(1, s.totalRunners);
    if (!s.byCheckpoint.length) {
      cpContainer.innerHTML = '<div class="text-xs text-slate-500">尚無站點資料</div>';
    } else {
      cpContainer.innerHTML = s.byCheckpoint
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

    // 排行榜分組下拉選單
    const catSelect = $('pub-cat-filter');
    catSelect.innerHTML =
      '<option value="ALL">全部組別</option>' +
      state.categories.map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    catSelect.value = state.catFilter;
  }

  /**
   * 選手搜尋：支援 姓名、參賽編號、完整身份證字號（SHA-256 隱私安全比對）
   */
  async function runSearch(rawQuery) {
    const q = String(rawQuery || '').trim();
    state.searchQuery = q;
    const container = $('pub-search-results');

    if (!q) {
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
      container.innerHTML = `
        <div class="rounded-2xl bg-slate-950 border border-slate-800 p-5 text-center text-sm text-slate-400">
          查無符合「<strong class="text-white">${esc(q)}</strong>」的選手資料，請確認編號、姓名或完整身份證字號是否正確。
        </div>
      `;
      return;
    }

    const checkpoints = state.computed.checkpoints;
    const statusBanner = {
      finished: { text: '🏁 已完賽', cls: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40' },
      on_course: { text: '🏃 賽道進行中', cls: 'bg-amber-500/20 text-amber-300 border-amber-500/40' },
      not_started: { text: '⏳ 尚未出發', cls: 'bg-slate-800 text-slate-300 border-slate-700' }
    };

    container.innerHTML = matches
      .slice(0, 10)
      .map((row) => {
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

        // 視覺化站點進度時間軸
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
                  ${row.idMasked ? ` ｜ 身份證字號：<code class="text-sky-300 font-mono-num">${esc(row.idMasked)}</code>` : ''}
                </div>
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
      })
      .join('');
  }

  function renderLeaderboard() {
    const rows = state.computed.rows.filter((r) => state.catFilter === 'ALL' || r.categoryId === state.catFilter);
    const tbody = $('pub-leaderboard-tbody');

    if (rows.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" class="py-8 text-center text-slate-500">目前尚無選手成績資料</td></tr>';
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

    $('form-pub-search').addEventListener('submit', (e) => {
      e.preventDefault();
      runSearch($('pub-search-input').value);
    });

    $('pub-search-input').addEventListener('input', (e) => {
      if (!e.target.value.trim()) runSearch('');
    });

    $('pub-cat-filter').addEventListener('change', (e) => {
      state.catFilter = e.target.value;
      renderLeaderboard();
    });

    $('pub-rankby').addEventListener('change', (e) => {
      state.rankBy = e.target.value;
      recomputeAndRender();
    });

    // 點擊排行榜任一列直接展開該選手完整分段與動態卡片
    $('pub-leaderboard-tbody').addEventListener('click', (e) => {
      const tr = e.target.closest('[data-pick-bib]');
      if (!tr) return;
      const bib = tr.dataset.pickBib;
      $('pub-search-input').value = bib;
      runSearch(bib);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }

  document.addEventListener('DOMContentLoaded', initPublicPage);
})();
