/**
 * 智慧路跑 NFC 晶片計時系統 — 主工作人員端與後台管理控制器
 */
(function () {
  'use strict';

  const DB = window.RaceDB;
  const TE = window.TimingEngine;
  const NFC = window.NFCManager;
  const Sync = window.CloudSync;
  const Excel = window.ExcelExporter;

  const state = {
    activeTab: 'scan',
    config: null,
    runners: [],
    scans: [],
    computed: null,
    adminUnlocked: false,
    adminPinSession: '',
    rankBy: 'gun',
    resultsCatFilter: 'ALL',
    resultsSearch: '',
    rosterSearch: '',
    editMode: 'scan' // 'scan' | 'penalty'
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

  function showToast(message, type = 'info') {
    const colors = {
      success: 'bg-emerald-950/95 border-emerald-500 text-emerald-200',
      warn: 'bg-amber-950/95 border-amber-500 text-amber-200',
      error: 'bg-rose-950/95 border-rose-500 text-rose-200',
      info: 'bg-slate-900/95 border-sky-500 text-sky-200'
    };
    const el = document.createElement('div');
    el.className = `pointer-events-auto border rounded-xl px-4 py-3 text-sm font-bold shadow-xl transition ${colors[type] || colors.info}`;
    el.textContent = message;
    $('toast-container').appendChild(el);
    setTimeout(() => {
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 300);
    }, 3200);
  }

  /* ========================================================================
   * 初始化與資料載入
   * ======================================================================== */
  async function initApp() {
    registerServiceWorker();
    await DB.init();

    // 載入裝置偏好設定
    const settings = DB.getDeviceSettings();
    NFC.sound.enabled = settings.soundEnabled !== false;
    if (settings.highContrast) document.body.classList.add('high-contrast');

    bindEvents();
    startLiveClock();
    renderNfcDiagnostics();
    await reloadState();

    // 網路狀態監聽（恢復連線時在連線模式下自動上傳排隊紀錄）
    window.addEventListener('online', async () => {
      updateNetworkBadge();
      if (Sync.isOnlineEnabled()) {
        const pending = await DB.getPendingSyncCounts();
        if (pending.total > 0) {
          showToast('網路已恢復，正在自動同步離線紀錄...', 'info');
          await handleUploadPendingSync(true);
        }
      }
    });
    window.addEventListener('offline', () => {
      updateNetworkBadge();
      showToast('已進入離線狀態，感應資料將安全暫存於本機 IndexedDB', 'warn');
    });
  }

  function registerServiceWorker() {
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
      navigator.serviceWorker.register('./sw.js').catch((err) => {
        console.warn('Service Worker 註冊失敗:', err);
      });
    }
  }

  async function reloadState() {
    const [config, runners, scans] = await Promise.all([
      DB.getConfig(),
      DB.getAllRunners(),
      DB.getAllScans()
    ]);

    state.config = config;
    state.runners = runners;
    state.scans = scans;
    state.computed = TE.computeResults({
      runners,
      scans,
      categories: config.categories || [],
      checkpoints: config.checkpoints || [],
      rankBy: state.rankBy
    });

    renderHeader();
    renderScanTab();
    renderWriterTab();
    renderResultsTab();
    renderAdminTab();
    await updatePendingSyncBadge();
  }

  function startLiveClock() {
    const clockEl = $('live-clock');
    function tick() {
      if (clockEl) {
        clockEl.textContent = TE.formatClockMs(Date.now());
      }
      requestAnimationFrame(tick);
    }
    requestAnimationFrame(tick);
  }

  function renderNfcDiagnostics() {
    const diag = NFC.getDiagnostics();
    const scanHint = $('scan-nfc-env-hint');
    const writerBanner = $('writer-nfc-banner');
    const writerText = $('writer-nfc-status-text');

    if (scanHint) {
      scanHint.textContent = diag.supported
        ? '✅ Web NFC 就緒（Android Chrome + HTTPS）'
        : '⚠️ 非 Web NFC 環境（請用下方模擬感應或手動補登）';
    }
    if (writerText) {
      writerText.textContent = diag.statusMessage;
    }
    if (writerBanner && !diag.supported) {
      writerBanner.className = 'rounded-2xl border border-amber-500/40 bg-amber-950/30 p-4 text-sm text-amber-200 flex items-start gap-3';
    }
  }

  /* ========================================================================
   * 頂部狀態與分頁切換
   * ======================================================================== */
  function renderHeader() {
    const settings = DB.getDeviceSettings();
    $('header-race-name').textContent = state.config.raceName || '智慧路跑 NFC 晶片計時系統';
    $('header-race-date').textContent = state.config.raceDate || '--';
    $('header-device-id').textContent = settings.deviceId;

    // 公開查詢頁連結自動帶入 api 參數（若已設定 GAS URL）
    const pubLink = $('link-public-page');
    if (settings.gasUrl) {
      pubLink.href = `./public.html?api=${encodeURIComponent(settings.gasUrl)}`;
    } else {
      pubLink.href = './public.html';
    }

    updateNetworkBadge();
    $('btn-toggle-sound').textContent = NFC.sound.enabled ? '🔊 音效開' : '🔇 音效關';
    $('btn-toggle-contrast').textContent = settings.highContrast ? '🌙 一般模式' : '☀️ 強光高對比';
  }

  function updateNetworkBadge() {
    const settings = DB.getDeviceSettings();
    const badge = $('badge-network');
    const isOnlineNet = navigator.onLine;
    if (settings.mode === 'online' && settings.gasUrl) {
      if (isOnlineNet) {
        badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-sky-400 animate-pulse"></span><span>多機雲端連線</span>';
        badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-sky-950 border border-sky-500/50 text-sky-300';
      } else {
        badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-amber-400"></span><span>斷網暫存中</span>';
        badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-amber-950 border border-amber-500/50 text-amber-300';
      }
    } else {
      badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-emerald-400"></span><span>單機離線模式</span>';
      badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-slate-800 border border-slate-700 text-slate-300';
    }
  }

  async function updatePendingSyncBadge() {
    const pending = await DB.getPendingSyncCounts();
    const btn = $('btn-quick-sync');
    const countEl = $('pending-sync-count');
    const settings = DB.getDeviceSettings();
    countEl.textContent = String(pending.total);

    if (pending.total > 0 && settings.gasUrl) {
      btn.classList.remove('hidden');
    } else {
      btn.classList.add('hidden');
    }

    const syncInfo = $('sync-status-info');
    if (syncInfo) {
      const lastStr = settings.lastSyncAt ? TE.formatClockSec(settings.lastSyncAt) : '尚未同步';
      syncInfo.textContent = `待同步資料：選手 ${pending.runners} 筆、感應 ${pending.scans} 筆 ｜ 最後同步時間：${lastStr}`;
    }
  }

  function switchTab(tabName) {
    state.activeTab = tabName;
    document.querySelectorAll('.tab-panel').forEach((p) => p.classList.add('hidden'));
    const target = $(`tab-${tabName}`);
    if (target) target.classList.remove('hidden');

    document.querySelectorAll('.nav-tab-btn').forEach((btn) => {
      const active = btn.dataset.tab === tabName;
      btn.className = active
        ? 'nav-tab-btn px-5 py-3 text-sm font-bold border-b-2 border-emerald-400 text-emerald-300'
        : 'nav-tab-btn px-5 py-3 text-sm font-bold border-b-2 border-transparent text-slate-400 hover:text-white';
    });

    document.querySelectorAll('.mobile-tab-btn').forEach((btn) => {
      const active = btn.dataset.tab === tabName;
      btn.className = active
        ? 'mobile-tab-btn py-2.5 flex flex-col items-center gap-0.5 text-emerald-400 font-bold text-[11px]'
        : 'mobile-tab-btn py-2.5 flex flex-col items-center gap-0.5 text-slate-400 font-bold text-[11px]';
    });
  }

  /* ========================================================================
   * TAB 1：賽道感應計時終端
   * ======================================================================== */
  function renderScanTab() {
    const checkpoints = [...(state.config.checkpoints || [])].sort((a, b) => a.order - b.order);
    const settings = DB.getDeviceSettings();
    const select = $('scan-checkpoint-select');

    const prevVal = settings.selectedCheckpointId || checkpoints[0]?.id || '';
    select.innerHTML = checkpoints
      .map((cp) => `<option value="${esc(cp.id)}">#${cp.order} ${esc(cp.name)} (${cp.distanceKm} km)</option>`)
      .join('');
    if (checkpoints.some((c) => c.id === prevVal)) {
      select.value = prevVal;
    } else if (checkpoints[0]) {
      select.value = checkpoints[0].id;
      DB.saveDeviceSettings({ selectedCheckpointId: checkpoints[0].id });
    }

    // 模擬器選手下拉選單
    const mockSelect = $('mock-runner-select');
    const prevMock = mockSelect.value;
    mockSelect.innerHTML =
      '<option value="">-- 隨機挑選一名選手 --</option>' +
      state.runners
        .slice()
        .sort((a, b) => a.bib.localeCompare(b.bib, 'zh-Hant', { numeric: true }))
        .map((r) => `<option value="${esc(r.bib)}">${esc(r.bib)} - ${esc(r.name)} (${esc(r.categoryId)})</option>`)
        .join('');
    if (prevMock && state.runners.some((r) => r.bib === prevMock)) {
      mockSelect.value = prevMock;
    }

    renderRecentCheckpointScans();
  }

  function renderRecentCheckpointScans() {
    const cpId = $('scan-checkpoint-select').value;
    const runnerMap = new Map(state.runners.map((r) => [r.bib, r]));
    const catMap = new Map((state.config.categories || []).map((c) => [c.id, c]));

    const list = state.scans
      .filter((s) => !s.deleted && s.checkpointId === cpId)
      .sort((a, b) => b.timestampMs - a.timestampMs);

    $('checkpoint-scan-count').textContent = `${list.length} 筆`;

    const tbody = $('recent-scans-tbody');
    if (list.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="py-6 text-center text-slate-500">本站尚無感應紀錄，請啟動感應或使用模擬器測試</td></tr>';
      return;
    }

    const sourceBadges = {
      nfc: '<span class="px-2 py-0.5 rounded bg-emerald-900/60 text-emerald-300 text-xs font-bold">NFC</span>',
      manual: '<span class="px-2 py-0.5 rounded bg-sky-900/60 text-sky-300 text-xs font-bold">手動</span>',
      mock: '<span class="px-2 py-0.5 rounded bg-indigo-900/60 text-indigo-300 text-xs font-bold">模擬</span>',
      admin: '<span class="px-2 py-0.5 rounded bg-amber-900/60 text-amber-300 text-xs font-bold">修正</span>'
    };

    tbody.innerHTML = list
      .slice(0, 15)
      .map((s) => {
        const r = runnerMap.get(s.bib);
        const cat = r ? catMap.get(r.categoryId) : null;
        return `
          <tr class="hover:bg-slate-800/40">
            <td class="py-2.5 pr-3 font-mono-num font-bold text-sky-300">${esc(TE.formatClockMs(s.timestampMs))}</td>
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-300">${esc(s.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(r ? r.name : '未登錄選手')}</td>
            <td class="py-2.5 pr-3 text-slate-300">${esc(cat ? cat.name : (r?.categoryId || '--'))}</td>
            <td class="py-2.5 pr-3">${sourceBadges[s.source] || esc(s.source)}</td>
            <td class="py-2.5 text-right text-xs font-bold ${s.synced ? 'text-emerald-400' : 'text-amber-400'}">
              ${s.synced ? '☁️ 已同步' : '💾 本機暫存'}
            </td>
          </tr>
        `;
      })
      .join('');
  }

  /**
   * 處理感應事件（NFC 實機 / 手動補登 / Mock 模擬共用同一核心流程）
   */
  async function processIncomingScan(tagData, { bypassDebounce = false } = {}) {
    const scanTs = Date.now();
    const cpId = $('scan-checkpoint-select').value;
    const cp = (state.config.checkpoints || []).find((c) => c.id === cpId);
    if (!cpId) {
      showToast('請先選擇感應站點', 'error');
      return;
    }

    // 查找選手資料；若本地未建檔但晶片內含完整資訊則自動補建檔
    let runner = await DB.getRunner(tagData.bib);
    if (!runner && tagData.name) {
      runner = await DB.upsertRunner(
        {
          bib: tagData.bib,
          name: tagData.name,
          idNumber: tagData.idNumber || '',
          categoryId: tagData.categoryId || state.config.categories?.[0]?.id || '',
          nfcUid: tagData.nfcUid || ''
        },
        false
      );
    }

    const res = await DB.addScan({
      bib: tagData.bib,
      checkpointId: cpId,
      timestampMs: scanTs,
      source: tagData.source || 'nfc',
      note: tagData.note || '',
      bypassDebounce
    });

    const cat = (state.config.categories || []).find((c) => c.id === (runner?.categoryId || tagData.categoryId));
    const displayName = runner?.name || tagData.name || `選手 ${tagData.bib}`;
    const displayCat = cat ? cat.name : (runner?.categoryId || tagData.categoryId || '未分組');

    const heroCard = $('scan-hero-card');
    heroCard.classList.remove('flash-success', 'flash-warn', 'flash-error');
    void heroCard.offsetWidth; // 重新觸發動畫

    if (!res.ok && res.duplicate) {
      NFC.sound.beepWarning();
      heroCard.classList.add('flash-warn');
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-amber-500 text-slate-950 mb-3';
      $('scan-status-pill').textContent = `⚠️ 防連點略過 — ${res.elapsedSec} 秒前已感應過（冷卻剩 ${res.waitRemainingSec} 秒）`;
      $('hero-bib-cat').textContent = `${tagData.bib} • ${displayCat}`;
      $('hero-name').textContent = displayName;
      $('hero-subinfo').textContent = `本站重複感應不重複計時 ｜ 上次紀錄：${TE.formatClockMs(res.previousScan.timestampMs)}`;
      return;
    }

    // 計算即時經過時間（優先顯示大會經過時間，並附上淨時間）
    const startCp = [...(state.config.checkpoints || [])].sort((a, b) => a.order - b.order).find((c) => c.type === 'start');
    const startScan = startCp
      ? state.scans.find((s) => !s.deleted && s.bib === tagData.bib && s.checkpointId === startCp.id)
      : null;

    let elapsedStr = '尚未鳴槍';
    if (cat && cat.gunTimeMs) {
      elapsedStr = TE.formatDuration(scanTs - Number(cat.gunTimeMs), true);
    } else if (startScan) {
      elapsedStr = TE.formatDuration(scanTs - Number(startScan.timestampMs), true);
    }

    NFC.sound.beepSuccess();
    heroCard.classList.add('flash-success');
    $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-emerald-500 text-slate-950 mb-3';
    $('scan-status-pill').textContent = `✅ 感應成功 — ${cp ? cp.name : cpId}`;
    $('hero-bib-cat').textContent = `${tagData.bib} • ${displayCat}`;
    $('hero-name').textContent = displayName;
    $('hero-elapsed').textContent = elapsedStr;
    $('hero-subinfo').textContent = `感應時間：${TE.formatClockMs(scanTs)} ｜ 來源：${(tagData.source || 'nfc').toUpperCase()}`;

    // 若開啟連線模式，非同步即時推送至 Google Sheets（不阻塞畫面速度）
    if (Sync.isOnlineEnabled()) {
      Sync.pushSingleScan(res.scan).then(() => reloadState());
    }
    await reloadState();
  }

  async function toggleNfcScanning() {
    const btn = $('btn-toggle-nfc-scan');
    if (NFC.isScanning) {
      NFC.stopScanning();
      btn.textContent = '📡 啟動極速 NFC 感應';
      btn.className = 'no-select w-full sm:w-auto sm:min-w-[280px] py-5 px-8 rounded-2xl text-xl sm:text-2xl font-black bg-emerald-500 hover:bg-emerald-400 active:scale-[0.98] text-slate-950 shadow-lg transition';
      $('scan-status-pill').textContent = '已暫停 NFC 感應';
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-slate-800 text-slate-300 mb-3';
      return;
    }

    try {
      await NFC.startScanning(
        (tag) => processIncomingScan(tag),
        (err) => showToast(err.message, 'error')
      );
      btn.textContent = '🛑 感應中（點此停止）';
      btn.className = 'no-select nfc-active-pulse w-full sm:w-auto sm:min-w-[280px] py-5 px-8 rounded-2xl text-xl sm:text-2xl font-black bg-rose-500 hover:bg-rose-400 text-white shadow-lg transition';
      $('scan-status-pill').textContent = '🟢 NFC 感應運作中 — 請將晶片貼近手機感應區';
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-emerald-950 border border-emerald-500 text-emerald-300 mb-3';
      showToast('已啟動 Web NFC 極速感應模式', 'success');
    } catch (err) {
      showToast(err.message, 'warn');
    }
  }

  /* ========================================================================
   * TAB 2：晶片寫入與選手報到
   * ======================================================================== */
  function renderWriterTab() {
    const catSelect = $('writer-category');
    const prev = catSelect.value;
    catSelect.innerHTML = (state.config.categories || [])
      .map((c) => `<option value="${esc(c.id)}">${esc(c.name)} (${esc(c.id)})</option>`)
      .join('');
    if (prev && (state.config.categories || []).some((c) => c.id === prev)) {
      catSelect.value = prev;
    }

    const q = state.rosterSearch.trim().toLowerCase();
    const filtered = state.runners
      .filter((r) => {
        if (!q) return true;
        return (
          r.bib.toLowerCase().includes(q) ||
          r.name.toLowerCase().includes(q) ||
          (r.idNumber || '').toLowerCase().includes(q)
        );
      })
      .sort((a, b) => a.bib.localeCompare(b.bib, 'zh-Hant', { numeric: true }));

    $('roster-count').textContent = String(state.runners.length);
    const catMap = new Map((state.config.categories || []).map((c) => [c.id, c.name]));

    const tbody = $('roster-tbody');
    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="py-6 text-center text-slate-500">尚無選手資料，請由上方新增或批次匯入</td></tr>';
      return;
    }

    tbody.innerHTML = filtered
      .map(
        (r) => `
        <tr class="hover:bg-slate-800/40">
          <td class="py-2.5 pr-3 font-mono-num font-black text-amber-300">${esc(r.bib)}</td>
          <td class="py-2.5 pr-3 font-bold text-white">${esc(r.name)}</td>
          <td class="py-2.5 pr-3 font-mono-num text-slate-300">${esc(r.idNumber || r.idMasked || '--')}</td>
          <td class="py-2.5 pr-3 text-slate-300">${esc(catMap.get(r.categoryId) || r.categoryId)}</td>
          <td class="py-2.5 pr-3 font-mono-num text-xs ${r.nfcUid ? 'text-emerald-400' : 'text-slate-500'}">
            ${esc(r.nfcUid || '尚未寫入')}
          </td>
          <td class="py-2.5 text-right space-x-1.5">
            <button type="button" data-load-runner="${esc(r.bib)}"
              class="px-2.5 py-1 rounded-lg text-xs font-bold bg-sky-600/30 hover:bg-sky-600/50 text-sky-300">
              帶入寫入
            </button>
            <button type="button" data-delete-runner="${esc(r.bib)}"
              class="px-2 py-1 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">
              刪除
            </button>
          </td>
        </tr>
      `
      )
      .join('');
  }

  function getWriterFormRunner() {
    return {
      bib: $('writer-bib').value.trim(),
      name: $('writer-name').value.trim(),
      idNumber: $('writer-idnumber').value.trim().toUpperCase(),
      categoryId: $('writer-category').value
    };
  }

  function setWriterVerifyBox(status, title, detail) {
    const box = $('writer-verify-box');
    const styles = {
      working: 'rounded-xl border border-sky-500/50 bg-sky-950/40 p-4 text-sm text-sky-200 space-y-1.5',
      success: 'rounded-xl border border-emerald-500/50 bg-emerald-950/40 p-4 text-sm text-emerald-200 space-y-1.5',
      error: 'rounded-xl border border-rose-500/50 bg-rose-950/40 p-4 text-sm text-rose-200 space-y-1.5'
    };
    box.className = styles[status] || styles.working;
    $('writer-verify-title').textContent = title;
    $('writer-verify-detail').textContent = detail;
  }

  async function executeNfcWrite(isMock = false) {
    const runner = getWriterFormRunner();
    if (!runner.bib || !runner.name) {
      showToast('請填寫參賽編號與選手姓名', 'warn');
      return;
    }

    setWriterVerifyBox('working', isMock ? '🧪 模擬寫入中...' : '📡 等待貼近 NFC 晶片...', '準備寫入 NDEF 格式資料...');

    try {
      const fn = isMock ? NFC.mockWriteAndVerify.bind(NFC) : NFC.writeAndVerify.bind(NFC);
      const result = await fn(runner, (step) => {
        setWriterVerifyBox('working', '⏳ 晶片寫入與自動驗證進行中', step);
      });

      const saved = await DB.upsertRunner(
        {
          ...runner,
          nfcUid: result.nfcUid,
          checkedInAt: Date.now()
        },
        false
      );

      setWriterVerifyBox(
        'success',
        `✅ 晶片寫入並自動讀取驗證成功！（UID: ${result.nfcUid || '標準 NDEF'}）`,
        `讀回驗證資料：${result.data.raw}`
      );
      showToast(`選手 ${runner.bib} ${runner.name} 報到寫入完成！`, 'success');

      if (Sync.isOnlineEnabled()) {
        Sync.pushRunners([saved]).then(() => updatePendingSyncBadge());
      }
      await reloadState();
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ 寫入或驗證失敗', err.message);
      showToast(err.message, 'error');
    }
  }

  async function importRunnersFromRows(rows) {
    const valid = [];
    const defaultCat = state.config.categories?.[0]?.id || '10K';
    for (const cols of rows) {
      if (!cols || cols.length < 2) continue;
      const bib = String(cols[0] ?? '').trim();
      const name = String(cols[1] ?? '').trim();
      if (!bib || !name || bib === '參賽編號' || bib.toLowerCase() === 'bib') continue;
      valid.push({
        bib,
        name,
        idNumber: String(cols[2] ?? '').trim().toUpperCase(),
        categoryId: String(cols[3] ?? defaultCat).trim() || defaultCat
      });
    }
    if (valid.length === 0) {
      showToast('未偵測到有效的選手名單資料', 'warn');
      return;
    }
    const saved = await DB.upsertRunnersBulk(valid, false);
    if (Sync.isOnlineEnabled()) {
      Sync.pushRunners(saved).then(() => updatePendingSyncBadge());
    }
    await reloadState();
    showToast(`已成功匯入 ${valid.length} 位選手名單！`, 'success');
  }

  /* ========================================================================
   * TAB 3：成績總覽、人工修正與 Excel 匯出
   * ======================================================================== */
  function renderResultsTab() {
    const computed = state.computed;
    if (!computed) return;

    // KPI
    $('kpi-total').textContent = String(computed.summary.totalRunners);
    $('kpi-finished').textContent = String(computed.summary.finishedCount);
    $('kpi-oncourse').textContent = String(computed.summary.onCourseCount);
    $('kpi-notstarted').textContent = String(computed.summary.notStartedCount);

    // 組別篩選下拉
    const catFilter = $('results-category-filter');
    catFilter.innerHTML =
      '<option value="ALL">全部組別</option>' +
      (state.config.categories || []).map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    catFilter.value = state.resultsCatFilter;
    $('results-rankby-select').value = state.rankBy;

    // 管理員修正按鈕與提示
    const unlockBtn = $('btn-admin-unlock-results');
    const editHint = $('results-edit-hint');
    if (state.adminUnlocked) {
      unlockBtn.textContent = '🔓 人工修正模式已啟用（點此鎖定）';
      unlockBtn.className = 'px-3.5 py-2 rounded-xl text-xs font-extrabold bg-emerald-500/20 border border-emerald-500/50 text-emerald-300';
      editHint.classList.remove('hidden');
    } else {
      unlockBtn.textContent = '🔒 解鎖人工成績修正';
      unlockBtn.className = 'px-3.5 py-2 rounded-xl text-xs font-extrabold bg-amber-500/20 border border-amber-500/50 text-amber-300 hover:bg-amber-500/30';
      editHint.classList.add('hidden');
    }

    const checkpoints = computed.checkpoints;
    $('results-thead').innerHTML = `
      <tr>
        <th class="px-3 py-3">總名次</th>
        <th class="px-3 py-3">組名次</th>
        <th class="px-3 py-3">編號</th>
        <th class="px-3 py-3">姓名</th>
        <th class="px-3 py-3">分組</th>
        <th class="px-3 py-3">狀態</th>
        ${checkpoints.map((cp) => `<th class="px-3 py-3">${esc(cp.name)}</th>`).join('')}
        <th class="px-3 py-3">分段耗時 / 配速</th>
        <th class="px-3 py-3">罰時</th>
        <th class="px-3 py-3">大會時間</th>
        <th class="px-3 py-3">個人淨時間</th>
        <th class="px-3 py-3">平均配速</th>
      </tr>
    `;

    const q = state.resultsSearch.trim().toLowerCase();
    const filtered = computed.rows.filter((r) => {
      if (state.resultsCatFilter !== 'ALL' && r.categoryId !== state.resultsCatFilter) return false;
      if (q) {
        return (
          r.bib.toLowerCase().includes(q) ||
          r.name.toLowerCase().includes(q) ||
          (r.idNumber || '').toLowerCase().includes(q)
        );
      }
      return true;
    });

    const tbody = $('results-tbody');
    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${11 + checkpoints.length}" class="py-8 text-center text-slate-500">目前無符合條件的選手成績</td></tr>`;
      return;
    }

    const statusMap = {
      finished: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-500/20 text-emerald-300">已完賽</span>',
      on_course: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-amber-500/20 text-amber-300">賽道上</span>',
      not_started: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-slate-800 text-slate-400">未出發</span>'
    };

    tbody.innerHTML = filtered
      .map((row) => {
        const cpCells = checkpoints
          .map((cp) => {
            const s = row.cpScanMap[cp.id];
            const label = s ? TE.formatClockMs(s.timestampMs) : '--';
            const editedTag = s && s.source === 'admin' ? ' <span class="text-amber-400">*</span>' : '';
            if (state.adminUnlocked) {
              return `
                <td class="px-3 py-2.5 font-mono-num">
                  <button type="button"
                    data-edit-scan-bib="${esc(row.bib)}"
                    data-edit-scan-cp="${esc(cp.id)}"
                    data-edit-scan-id="${esc(s ? s.id : '')}"
                    class="px-2 py-1 rounded-lg border border-dashed ${s ? 'border-sky-500/50 text-sky-300 hover:bg-sky-900/40' : 'border-slate-700 text-slate-500 hover:border-amber-400 hover:text-amber-300'} transition">
                    ${label}${editedTag}
                  </button>
                </td>
              `;
            }
            return `<td class="px-3 py-2.5 font-mono-num ${s ? 'text-slate-200' : 'text-slate-600'}">${label}${editedTag}</td>`;
          })
          .join('');

        const splitSummary = row.splits
          .filter((sp) => sp.durationMs !== null)
          .map((sp) => `<div class="text-[11px] font-mono-num text-slate-300">${esc(sp.toName)}: <span class="text-sky-300">${TE.formatDuration(sp.durationMs, false)}</span> (${sp.pace})</div>`)
          .join('') || '<span class="text-slate-600">--</span>';

        const penaltyCell = state.adminUnlocked
          ? `<button type="button" data-edit-penalty-bib="${esc(row.bib)}" class="px-2 py-1 rounded border border-dashed border-amber-500/50 text-amber-300 font-mono-num hover:bg-amber-900/30">${row.penaltySec ? `+${row.penaltySec}s` : '0s'}</button>`
          : `<span class="font-mono-num ${row.penaltySec ? 'text-rose-400 font-bold' : 'text-slate-500'}">${row.penaltySec ? `${row.penaltySec > 0 ? '+' : ''}${row.penaltySec}s` : '0s'}</span>`;

        return `
          <tr class="hover:bg-slate-800/40 align-top">
            <td class="px-3 py-2.5 font-mono-num font-black text-amber-400">${row.overallRank ?? '-'}</td>
            <td class="px-3 py-2.5 font-mono-num font-bold text-sky-300">${row.categoryRank ?? '-'}</td>
            <td class="px-3 py-2.5 font-mono-num font-bold text-white">${esc(row.bib)}</td>
            <td class="px-3 py-2.5 font-bold text-white">${esc(row.name)}</td>
            <td class="px-3 py-2.5">
              <span class="inline-flex items-center gap-1.5 text-xs font-bold">
                <span class="w-2 h-2 rounded-full" style="background:${esc(row.categoryColor)}"></span>
                ${esc(row.categoryName)}
              </span>
            </td>
            <td class="px-3 py-2.5">${statusMap[row.status] || row.status}</td>
            ${cpCells}
            <td class="px-3 py-2.5">${splitSummary}</td>
            <td class="px-3 py-2.5">${penaltyCell}</td>
            <td class="px-3 py-2.5 font-mono-num font-bold ${row.gunDurationMs !== null ? 'text-emerald-400' : 'text-slate-600'}">
              ${row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, true) : '--'}
            </td>
            <td class="px-3 py-2.5 font-mono-num font-bold ${row.chipDurationMs !== null ? 'text-sky-300' : 'text-slate-600'}">
              ${row.chipDurationMs !== null ? TE.formatDuration(row.chipDurationMs, true) : '--'}
            </td>
            <td class="px-3 py-2.5 font-mono-num text-slate-300">${esc(row.avgPace)}</td>
          </tr>
        `;
      })
      .join('');
  }

  function openEditScanModal(bib, cpId, scanId) {
    const runner = state.runners.find((r) => r.bib === bib);
    const cp = (state.config.checkpoints || []).find((c) => c.id === cpId);
    const scan = scanId ? state.scans.find((s) => s.id === scanId) : null;

    state.editMode = 'scan';
    $('modal-edit-title').textContent = `人工修正時間 — ${bib} ${runner?.name || ''} @ ${cp?.name || cpId}`;
    $('edit-bib').value = bib;
    $('edit-cp-id').value = cpId;
    $('edit-scan-id').value = scanId || '';
    $('edit-time-str').value = scan ? TE.formatClockMs(scan.timestampMs) : TE.formatClockMs(Date.now());
    $('edit-note').value = scan?.note || '後台人工修正';

    $('edit-scan-fields').classList.remove('hidden');
    $('edit-penalty-fields').classList.add('hidden');
    $('btn-delete-scan').classList.toggle('hidden', !scan);
    $('modal-edit-scan').classList.remove('hidden');
  }

  function openEditPenaltyModal(bib) {
    const runner = state.runners.find((r) => r.bib === bib);
    state.editMode = 'penalty';
    $('modal-edit-title').textContent = `判罰加時設定 — ${bib} ${runner?.name || ''}`;
    $('edit-bib').value = bib;
    $('edit-penalty-sec').value = String(runner?.penaltySec || 0);

    $('edit-scan-fields').classList.add('hidden');
    $('edit-penalty-fields').classList.remove('hidden');
    $('btn-delete-scan').classList.add('hidden');
    $('modal-edit-scan').classList.remove('hidden');
  }

  function closeEditModal() {
    $('modal-edit-scan').classList.add('hidden');
  }

  /* ========================================================================
   * TAB 4：後台管理（賽事設定、分組鳴槍、站點設定、雲端同步）
   * ======================================================================== */
  function renderAdminTab() {
    if (!state.adminUnlocked) {
      $('admin-lock-panel').classList.remove('hidden');
      $('admin-content-panel').classList.add('hidden');
      return;
    }

    $('admin-lock-panel').classList.add('hidden');
    $('admin-content-panel').classList.remove('hidden');

    const settings = DB.getDeviceSettings();
    document.querySelectorAll('input[name="sync-mode"]').forEach((radio) => {
      radio.checked = radio.value === settings.mode;
    });
    $('input-gas-url').value = settings.gasUrl || '';

    // 基本資料
    $('cfg-race-name').value = state.config.raceName || '';
    $('cfg-race-date').value = state.config.raceDate || '';
    $('cfg-debounce-sec').value = String(state.config.debounceSeconds ?? 30);
    $('cfg-admin-pin').value = state.config.adminPin || '1234';

    // 分組與起跑時間
    const catList = $('categories-list');
    catList.innerHTML = (state.config.categories || [])
      .map(
        (c, idx) => `
        <div class="p-4 rounded-xl bg-slate-950 border border-slate-800 space-y-3" data-cat-index="${idx}">
          <div class="grid grid-cols-1 sm:grid-cols-12 gap-2.5 items-end">
            <div class="sm:col-span-2">
              <label class="block text-[11px] text-slate-400 mb-1">代碼</label>
              <input type="text" value="${esc(c.id)}" data-cat-field="id" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-sm font-mono-num text-white" />
            </div>
            <div class="sm:col-span-4">
              <label class="block text-[11px] text-slate-400 mb-1">分組名稱</label>
              <input type="text" value="${esc(c.name)}" data-cat-field="name" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-sm font-bold text-white" />
            </div>
            <div class="sm:col-span-2">
              <label class="block text-[11px] text-slate-400 mb-1">距離 (km)</label>
              <input type="number" step="0.001" value="${esc(c.distanceKm)}" data-cat-field="distanceKm" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-sm font-mono-num text-white" />
            </div>
            <div class="sm:col-span-4 flex flex-wrap gap-1.5">
              <button type="button" data-save-cat="${idx}" class="px-3 py-2 rounded-lg text-xs font-bold bg-sky-600 hover:bg-sky-500 text-white">儲存分組</button>
              <button type="button" data-del-cat="${idx}" class="px-3 py-2 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>
            </div>
          </div>

          <!-- 獨立鳴槍起跑控制列 -->
          <div class="pt-2 border-t border-slate-800/80 flex flex-wrap items-center justify-between gap-2.5">
            <div class="flex items-center gap-2">
              <span class="text-xs text-slate-400 font-bold">鳴槍起跑時間：</span>
              <span class="font-mono-num text-base font-black ${c.gunTimeMs ? 'text-emerald-400' : 'text-slate-500'}">
                ${c.gunTimeMs ? TE.formatClockSec(c.gunTimeMs) : '尚未鳴槍'}
              </span>
            </div>

            <div class="flex flex-wrap items-center gap-2">
              <button type="button" data-gun-now="${idx}"
                class="px-4 py-2 rounded-xl text-xs font-black bg-emerald-500 hover:bg-emerald-400 text-slate-950 shadow">
                🔫 一鍵鳴槍起跑（記錄當下）
              </button>
              <input type="text" placeholder="HH:mm:ss" value="${c.gunTimeMs ? TE.formatClockSec(c.gunTimeMs) : ''}" data-gun-manual-input="${idx}"
                class="w-28 font-mono-num text-xs bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-white" />
              <button type="button" data-gun-manual-set="${idx}"
                class="px-3 py-2 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-sky-300">
                手動套用
              </button>
              <button type="button" data-gun-clear="${idx}"
                class="px-2.5 py-2 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-400">
                清除
              </button>
            </div>
          </div>
        </div>
      `
      )
      .join('');

    // 站點設定清單
    const cpList = $('checkpoints-list');
    const sortedCps = [...(state.config.checkpoints || [])].sort((a, b) => a.order - b.order);
    cpList.innerHTML = sortedCps
      .map(
        (cp, idx) => `
        <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-center p-3 rounded-xl bg-slate-950 border border-slate-800" data-cp-row="${idx}">
          <div class="sm:col-span-1">
            <label class="sm:hidden text-[11px] text-slate-400">順序</label>
            <input type="number" value="${esc(cp.order)}" data-cp-field="order" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-sm font-mono-num text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="sm:hidden text-[11px] text-slate-400">代碼</label>
            <input type="text" value="${esc(cp.id)}" data-cp-field="id" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-mono-num text-white" />
          </div>
          <div class="sm:col-span-4">
            <label class="sm:hidden text-[11px] text-slate-400">站點名稱</label>
            <input type="text" value="${esc(cp.name)}" data-cp-field="name" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-bold text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="sm:hidden text-[11px] text-slate-400">累計距離 (km)</label>
            <input type="number" step="0.1" value="${esc(cp.distanceKm)}" data-cp-field="distanceKm" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-mono-num text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="sm:hidden text-[11px] text-slate-400">站點類型</label>
            <select data-cp-field="type" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-xs font-bold text-white">
              <option value="start" ${cp.type === 'start' ? 'selected' : ''}>起點 (Start)</option>
              <option value="checkpoint" ${cp.type === 'checkpoint' ? 'selected' : ''}>中途補給/折返</option>
              <option value="finish" ${cp.type === 'finish' ? 'selected' : ''}>終點 (Finish)</option>
            </select>
          </div>
          <div class="sm:col-span-1 text-right">
            <button type="button" data-del-cp="${idx}" class="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>
          </div>
        </div>
      `
      )
      .join('');
  }

  async function syncConfigIfOnline(cfg) {
    if (Sync.isOnlineEnabled()) {
      try {
        await Sync.pushConfig(cfg, state.adminPinSession || cfg.adminPin);
        showToast('設定已同步至 Google 試算表', 'success');
      } catch (err) {
        showToast(`本機已儲存，但雲端同步失敗：${err.message}`, 'warn');
      }
    }
  }

  async function handleUploadPendingSync(silent = false) {
    try {
      const res = await Sync.syncPendingToCloud();
      await reloadState();
      if (!silent || res.uploadedScans > 0 || res.uploadedRunners > 0) {
        showToast(
          `同步完成！上傳選手 ${res.uploadedRunners} 筆、感應 ${res.uploadedScans} 筆` +
            (res.skippedDuplicate ? `（略過雲端防連點重複 ${res.skippedDuplicate} 筆）` : ''),
          'success'
        );
      }
    } catch (err) {
      showToast(`雲端上傳失敗：${err.message}`, 'error');
    }
  }

  /* ========================================================================
   * 事件綁定
   * ======================================================================== */
  function bindEvents() {
    // 分頁切換
    document.querySelectorAll('[data-tab]').forEach((btn) => {
      btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });

    // 站點選擇持久化
    $('scan-checkpoint-select').addEventListener('change', (e) => {
      DB.saveDeviceSettings({ selectedCheckpointId: e.target.value });
      renderRecentCheckpointScans();
    });

    // 音效開關 & 強光高對比開關
    $('btn-toggle-sound').addEventListener('click', () => {
      NFC.sound.enabled = !NFC.sound.enabled;
      DB.saveDeviceSettings({ soundEnabled: NFC.sound.enabled });
      renderHeader();
    });

    $('btn-toggle-contrast').addEventListener('click', () => {
      const next = !DB.getDeviceSettings().highContrast;
      DB.saveDeviceSettings({ highContrast: next });
      document.body.classList.toggle('high-contrast', next);
      renderHeader();
    });

    // 啟動/停止 NFC 感應
    $('btn-toggle-nfc-scan').addEventListener('click', toggleNfcScanning);

    // 頂部快速同步按鈕
    $('btn-quick-sync').addEventListener('click', () => handleUploadPendingSync(false));

    // 手動補感應表單
    $('form-manual-scan').addEventListener('submit', async (e) => {
      e.preventDefault();
      const input = $('manual-scan-input');
      const query = input.value.trim();
      if (!query) return;

      const found = await DB.findRunnerByBibOrId(query);
      if (found) {
        await processIncomingScan(
          { bib: found.bib, name: found.name, categoryId: found.categoryId, source: 'manual', note: '現場手動補登' },
          { bypassDebounce: true }
        );
      } else {
        await processIncomingScan(
          { bib: query.toUpperCase(), name: '', source: 'manual', note: '手動輸入未登錄編號' },
          { bypassDebounce: true }
        );
      }
      input.value = '';
    });

    // 模擬器：載入模擬選手捷徑
    $('btn-load-demo-shortcut').addEventListener('click', async () => {
      await DB.loadDemoData();
      await reloadState();
      showToast('已載入 9 位模擬選手與示範賽道紀錄，可直接點擊「模擬晶片刷過」！', 'success');
    });

    // 模擬器：模擬 NFC 晶片刷過
    $('btn-mock-scan').addEventListener('click', async () => {
      if (state.runners.length === 0) {
        await DB.loadDemoData();
        await reloadState();
      }
      const selectedBib = $('mock-runner-select').value;
      const runner = selectedBib
        ? state.runners.find((r) => r.bib === selectedBib)
        : state.runners[Math.floor(Math.random() * state.runners.length)];

      if (!runner) return;
      await processIncomingScan({
        bib: runner.bib,
        name: runner.name,
        idNumber: runner.idNumber,
        categoryId: runner.categoryId,
        source: 'mock',
        note: '模擬器測試感應'
      });
    });

    // TAB 2：寫入 NFC 晶片
    $('form-nfc-write').addEventListener('submit', (e) => {
      e.preventDefault();
      executeNfcWrite(false);
    });
    $('btn-mock-write-nfc').addEventListener('click', () => executeNfcWrite(true));
    $('btn-save-runner-only').addEventListener('click', async () => {
      const r = getWriterFormRunner();
      if (!r.bib || !r.name) return showToast('請輸入參賽編號與姓名', 'warn');
      const saved = await DB.upsertRunner(r, false);
      if (Sync.isOnlineEnabled()) Sync.pushRunners([saved]);
      await reloadState();
      showToast(`已儲存選手 ${r.bib} ${r.name}`, 'success');
    });

    // CSV 文字批次匯入
    $('btn-import-csv-text').addEventListener('click', async () => {
      const text = $('textarea-bulk-csv').value.trim();
      if (!text) return showToast('請先貼上 CSV 名單內容', 'warn');
      const rows = text.split(/\r?\n/).map((line) => line.split(/[,\t]/));
      await importRunnersFromRows(rows);
      $('textarea-bulk-csv').value = '';
    });

    // 檔案上傳 CSV / Excel 匯入
    $('input-import-file').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const buf = await file.arrayBuffer();
        const wb = window.XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
        await importRunnersFromRows(rows);
      } catch (err) {
        showToast(`檔案解析失敗：${err.message}`, 'error');
      }
      e.target.value = '';
    });

    // 選手名單搜尋與操作
    $('roster-search').addEventListener('input', (e) => {
      state.rosterSearch = e.target.value;
      renderWriterTab();
    });

    $('roster-tbody').addEventListener('click', async (e) => {
      const loadBtn = e.target.closest('[data-load-runner]');
      if (loadBtn) {
        const r = state.runners.find((x) => x.bib === loadBtn.dataset.loadRunner);
        if (r) {
          $('writer-bib').value = r.bib;
          $('writer-name').value = r.name;
          $('writer-idnumber').value = r.idNumber || '';
          $('writer-category').value = r.categoryId;
          window.scrollTo({ top: 0, behavior: 'smooth' });
          showToast(`已帶入 ${r.bib} ${r.name} 至寫入表單`, 'info');
        }
        return;
      }
      const delBtn = e.target.closest('[data-delete-runner]');
      if (delBtn) {
        const bib = delBtn.dataset.deleteRunner;
        if (confirm(`確定要刪除選手 ${bib} 嗎？`)) {
          await DB.deleteRunner(bib);
          await reloadState();
          showToast(`已刪除選手 ${bib}`, 'info');
        }
      }
    });

    // TAB 3：成績篩選、排序、解鎖修正、Excel 匯出
    $('results-category-filter').addEventListener('change', (e) => {
      state.resultsCatFilter = e.target.value;
      renderResultsTab();
    });
    $('results-rankby-select').addEventListener('change', async (e) => {
      state.rankBy = e.target.value;
      await reloadState();
    });
    $('results-search').addEventListener('input', (e) => {
      state.resultsSearch = e.target.value;
      renderResultsTab();
    });

    $('btn-admin-unlock-results').addEventListener('click', () => {
      if (state.adminUnlocked) {
        state.adminUnlocked = false;
        renderResultsTab();
        renderAdminTab();
        showToast('已鎖定管理員權限', 'info');
        return;
      }
      const pin = prompt('請輸入管理員密碼（預設 1234）：');
      if (pin === null) return;
      if (String(pin) === String(state.config.adminPin || '1234')) {
        state.adminUnlocked = true;
        state.adminPinSession = pin;
        renderResultsTab();
        renderAdminTab();
        showToast('已解鎖管理員成績修正模式', 'success');
      } else {
        showToast('管理員密碼錯誤', 'error');
      }
    });

    $('btn-export-excel').addEventListener('click', () => {
      try {
        const fileName = Excel.exportResultsToExcel({
          config: state.config,
          computed: state.computed,
          scans: state.scans,
          maskId: false
        });
        showToast(`已匯出 Excel 報表：${fileName}`, 'success');
      } catch (err) {
        showToast(err.message, 'error');
      }
    });

    // 點擊成績表站點時間 / 罰時開啟修正 Modal
    $('results-tbody').addEventListener('click', (e) => {
      const scanBtn = e.target.closest('[data-edit-scan-bib]');
      if (scanBtn) {
        openEditScanModal(
          scanBtn.dataset.editScanBib,
          scanBtn.dataset.editScanCp,
          scanBtn.dataset.editScanId
        );
        return;
      }
      const penBtn = e.target.closest('[data-edit-penalty-bib]');
      if (penBtn) {
        openEditPenaltyModal(penBtn.dataset.editPenaltyBib);
      }
    });

    // 修正 Modal 關閉與提交
    $('btn-close-edit-modal').addEventListener('click', closeEditModal);
    $('btn-cancel-edit-modal').addEventListener('click', closeEditModal);

    $('form-edit-scan').addEventListener('submit', async (e) => {
      e.preventDefault();
      const bib = $('edit-bib').value;

      if (state.editMode === 'penalty') {
        const sec = Number($('edit-penalty-sec').value || 0);
        const updated = await DB.upsertRunner({ bib, penaltySec: sec }, false);
        if (Sync.isOnlineEnabled()) Sync.pushRunners([updated]);
        closeEditModal();
        await reloadState();
        showToast(`已更新選手 ${bib} 判罰加時 (${sec}s) 並重算名次`, 'success');
        return;
      }

      const cpId = $('edit-cp-id').value;
      const scanId = $('edit-scan-id').value;
      const timeStr = $('edit-time-str').value.trim();
      const note = $('edit-note').value.trim() || '後台人工修正';

      const tsMs = TE.parseDateTimeToMs(state.config.raceDate, timeStr);
      if (tsMs === null) {
        showToast('時間格式錯誤，請使用 HH:mm:ss 或 HH:mm:ss.SSS', 'error');
        return;
      }

      let savedScan;
      if (scanId) {
        const existing = state.scans.find((s) => s.id === scanId);
        savedScan = await DB.putScan({
          ...existing,
          timestampMs: tsMs,
          source: 'admin',
          note,
          deleted: false,
          synced: 0
        });
      } else {
        const res = await DB.addScan({
          bib,
          checkpointId: cpId,
          timestampMs: tsMs,
          source: 'admin',
          note,
          bypassDebounce: true
        });
        savedScan = res.scan;
      }

      if (Sync.isOnlineEnabled()) {
        try {
          await Sync.pushAdminEditScan(savedScan, state.adminPinSession || state.config.adminPin);
        } catch (err) {
          showToast(`本機已修正，雲端同步待重試：${err.message}`, 'warn');
        }
      }

      closeEditModal();
      await reloadState();
      showToast('成績時間已更新並自動重算總名次與分組名次！', 'success');
    });

    $('btn-delete-scan').addEventListener('click', async () => {
      const scanId = $('edit-scan-id').value;
      if (!scanId) return;
      if (!confirm('確定要刪除此站點感應紀錄嗎？')) return;

      const existing = state.scans.find((s) => s.id === scanId);
      if (!existing) return;

      const updated = await DB.putScan({
        ...existing,
        deleted: true,
        source: 'admin',
        note: '管理員刪除誤感應',
        synced: 0
      });

      if (Sync.isOnlineEnabled()) {
        try {
          await Sync.pushAdminEditScan(updated, state.adminPinSession || state.config.adminPin);
        } catch {}
      }

      closeEditModal();
      await reloadState();
      showToast('已刪除此站感應紀錄並重新計算名次', 'info');
    });

    // TAB 4：管理員解鎖 / 鎖定
    $('form-admin-login').addEventListener('submit', (e) => {
      e.preventDefault();
      const pin = $('input-admin-pin').value.trim();
      if (pin === String(state.config.adminPin || '1234')) {
        state.adminUnlocked = true;
        state.adminPinSession = pin;
        $('input-admin-pin').value = '';
        renderAdminTab();
        renderResultsTab();
        showToast('後台管理已解鎖', 'success');
      } else {
        NFC.sound.beepError();
        showToast('密碼錯誤（預設密碼為 1234）', 'error');
      }
    });

    $('btn-admin-lock-now').addEventListener('click', () => {
      state.adminUnlocked = false;
      renderAdminTab();
      renderResultsTab();
      showToast('後台已鎖定', 'info');
    });

    // 系統模式切換
    document.querySelectorAll('input[name="sync-mode"]').forEach((radio) => {
      radio.addEventListener('change', (e) => {
        DB.saveDeviceSettings({ mode: e.target.value });
        renderHeader();
        showToast(e.target.value === 'online' ? '已切換為「多機即時連線模式」' : '已切換為「單機離線模式」', 'info');
      });
    });

    // 儲存並測試 GAS URL
    $('btn-save-gas-url').addEventListener('click', async () => {
      const url = $('input-gas-url').value.trim();
      DB.saveDeviceSettings({ gasUrl: url });
      renderHeader();
      if (!url) return showToast('已清除 GAS URL', 'info');
      try {
        const res = await Sync.ping(url);
        showToast(`✅ 連線成功！${res.message || ''}`, 'success');
      } catch (err) {
        showToast(`❌ 連線測試失敗：${err.message}`, 'error');
      }
    });

    $('btn-sync-upload').addEventListener('click', () => handleUploadPendingSync(false));

    $('btn-sync-pull').addEventListener('click', async () => {
      try {
        await Sync.pullFromCloud(state.adminPinSession || state.config.adminPin);
        await reloadState();
        showToast('已從 Google 試算表拉取最新賽事與感應資料！', 'success');
      } catch (err) {
        showToast(`雲端拉取失敗：${err.message}`, 'error');
      }
    });

    $('btn-copy-public-link').addEventListener('click', async () => {
      const gasUrl = DB.getDeviceSettings().gasUrl;
      const base = new URL('./public.html', location.href);
      if (gasUrl) base.searchParams.set('api', gasUrl);
      try {
        await navigator.clipboard.writeText(base.toString());
        showToast('已複製公開查詢頁連結！', 'success');
      } catch {
        prompt('請複製以下公開查詢頁連結：', base.toString());
      }
    });

    // 儲存賽事基本資料
    $('form-race-basic').addEventListener('submit', async (e) => {
      e.preventDefault();
      const updated = await DB.saveConfig({
        ...state.config,
        raceName: $('cfg-race-name').value.trim(),
        raceDate: $('cfg-race-date').value,
        debounceSeconds: Number($('cfg-debounce-sec').value || 30),
        adminPin: $('cfg-admin-pin').value.trim() || '1234'
      });
      state.adminPinSession = updated.adminPin;
      await syncConfigIfOnline(updated);
      await reloadState();
      showToast('賽事基本設定已儲存', 'success');
    });

    // 分組管理與鳴槍起跑
    $('btn-add-category').addEventListener('click', async () => {
      const cats = [...(state.config.categories || [])];
      const nextNum = cats.length + 1;
      cats.push({
        id: `CAT${nextNum}`,
        name: `新分組 ${nextNum}`,
        distanceKm: 5,
        gunTimeMs: null,
        gunTimeIso: '',
        color: '#a855f7'
      });
      const updated = await DB.saveConfig({ ...state.config, categories: cats });
      await syncConfigIfOnline(updated);
      await reloadState();
    });

    $('categories-list').addEventListener('click', async (e) => {
      const cats = [...(state.config.categories || [])];

      const saveBtn = e.target.closest('[data-save-cat]');
      if (saveBtn) {
        const idx = Number(saveBtn.dataset.saveCat);
        const card = saveBtn.closest('[data-cat-index]');
        cats[idx] = {
          ...cats[idx],
          id: card.querySelector('[data-cat-field="id"]').value.trim(),
          name: card.querySelector('[data-cat-field="name"]').value.trim(),
          distanceKm: Number(card.querySelector('[data-cat-field="distanceKm"]').value || 0)
        };
        const updated = await DB.saveConfig({ ...state.config, categories: cats });
        await syncConfigIfOnline(updated);
        await reloadState();
        showToast('分組資料已更新', 'success');
        return;
      }

      const delBtn = e.target.closest('[data-del-cat]');
      if (delBtn) {
        const idx = Number(delBtn.dataset.delCat);
        if (!confirm(`確定刪除分組「${cats[idx]?.name}」？`)) return;
        cats.splice(idx, 1);
        const updated = await DB.saveConfig({ ...state.config, categories: cats });
        await syncConfigIfOnline(updated);
        await reloadState();
        return;
      }

      // 一鍵鳴槍起跑
      const gunNowBtn = e.target.closest('[data-gun-now]');
      if (gunNowBtn) {
        const idx = Number(gunNowBtn.dataset.gunNow);
        const nowSecMs = Math.floor(Date.now() / 1000) * 1000;
        cats[idx].gunTimeMs = nowSecMs;
        cats[idx].gunTimeIso = new Date(nowSecMs).toISOString();
        NFC.sound.beepGun();
        const updated = await DB.saveConfig({ ...state.config, categories: cats });
        await syncConfigIfOnline(updated);
        await reloadState();
        showToast(`🔫 「${cats[idx].name}」已鳴槍起跑！時間：${TE.formatClockSec(nowSecMs)}`, 'success');
        return;
      }

      // 手動輸入起跑時間
      const gunManualBtn = e.target.closest('[data-gun-manual-set]');
      if (gunManualBtn) {
        const idx = Number(gunManualBtn.dataset.gunManualSet);
        const val = document.querySelector(`[data-gun-manual-input="${idx}"]`).value.trim();
        const tsMs = TE.parseDateTimeToMs(state.config.raceDate, val);
        if (tsMs === null) {
          return showToast('時間格式請輸入 HH:mm:ss（例如 06:30:00）', 'error');
        }
        cats[idx].gunTimeMs = tsMs;
        cats[idx].gunTimeIso = new Date(tsMs).toISOString();
        const updated = await DB.saveConfig({ ...state.config, categories: cats });
        await syncConfigIfOnline(updated);
        await reloadState();
        showToast(`已設定「${cats[idx].name}」起跑時間為 ${TE.formatClockSec(tsMs)}`, 'success');
        return;
      }

      // 清除鳴槍
      const gunClearBtn = e.target.closest('[data-gun-clear]');
      if (gunClearBtn) {
        const idx = Number(gunClearBtn.dataset.gunClear);
        cats[idx].gunTimeMs = null;
        cats[idx].gunTimeIso = '';
        const updated = await DB.saveConfig({ ...state.config, categories: cats });
        await syncConfigIfOnline(updated);
        await reloadState();
        showToast(`已清除「${cats[idx].name}」鳴槍時間`, 'info');
      }
    });

    // 站點管理
    $('btn-add-checkpoint').addEventListener('click', async () => {
      const cps = [...(state.config.checkpoints || [])];
      const maxOrder = cps.reduce((m, c) => Math.max(m, Number(c.order || 0)), 0);
      cps.push({
        id: `CP${maxOrder + 1}`,
        name: `補給站 ${maxOrder + 1}`,
        order: maxOrder + 1,
        distanceKm: 15,
        type: 'checkpoint'
      });
      const updated = await DB.saveConfig({ ...state.config, checkpoints: cps });
      await reloadState();
      showToast('已新增站點，請編輯內容後點擊「儲存站點設定」', 'info');
    });

    $('checkpoints-list').addEventListener('click', async (e) => {
      const delBtn = e.target.closest('[data-del-cp]');
      if (!delBtn) return;
      const idx = Number(delBtn.dataset.delCp);
      const sorted = [...(state.config.checkpoints || [])].sort((a, b) => a.order - b.order);
      sorted.splice(idx, 1);
      const updated = await DB.saveConfig({ ...state.config, checkpoints: sorted });
      await syncConfigIfOnline(updated);
      await reloadState();
    });

    $('btn-save-checkpoints').addEventListener('click', async () => {
      const rows = Array.from(document.querySelectorAll('#checkpoints-list [data-cp-row]'));
      const cps = rows
        .map((row) => ({
          order: Number(row.querySelector('[data-cp-field="order"]').value || 0),
          id: row.querySelector('[data-cp-field="id"]').value.trim(),
          name: row.querySelector('[data-cp-field="name"]').value.trim(),
          distanceKm: Number(row.querySelector('[data-cp-field="distanceKm"]').value || 0),
          type: row.querySelector('[data-cp-field="type"]').value
        }))
        .filter((cp) => cp.id && cp.name)
        .sort((a, b) => a.order - b.order);

      const updated = await DB.saveConfig({ ...state.config, checkpoints: cps });
      await syncConfigIfOnline(updated);
      await reloadState();
      showToast('感應站點設定已儲存！', 'success');
    });

    // 載入模擬資料 / 清空資料
    $('btn-load-demo-data').addEventListener('click', async () => {
      if (!confirm('載入模擬賽事資料將會覆蓋目前本機的測試選手與感應紀錄，確定繼續嗎？')) return;
      await DB.loadDemoData();
      await reloadState();
      showToast('已載入完整模擬賽事資料（9 名選手、3 分組、多站感應）！', 'success');
    });

    $('btn-clear-all-data').addEventListener('click', async () => {
      if (!confirm('警告：這將清空本機所有選手名單與感應紀錄（不影響已上傳至 Google 試算表的雲端資料），確定清空嗎？')) return;
      await DB.clearAllData(true);
      await reloadState();
      showToast('本機資料已全部清空並重設為預設賽事結構', 'info');
    });
  }

  document.addEventListener('DOMContentLoaded', initApp);
})();
