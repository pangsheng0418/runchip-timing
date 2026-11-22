/**
 * 智慧路跑 NFC 晶片計時系統 v2.0 — 三大作業區權限管制 + 離線/連線無縫同步主控制器
 */
(function () {
  'use strict';

  const DB = window.RaceDB;
  const TE = window.TimingEngine;
  const NFC = window.NFCManager;
  const Sync = window.CloudSync;
  const Excel = window.ExcelExporter;

  const state = {
    view: 'portal', // 'portal' | 'staff-login' | 'admin-login' | 'staff-workspace' | 'admin-workspace'
    session: null,  // { mode: 'staff'|'admin', account: {...}, auth: { username, password } }
    staffSubTab: 'scan', // 'scan' | 'writer'
    adminTab: 'results', // 'results' | 'staff' | 'race' | 'backup' | 'sync'
    config: null,
    runners: [],
    scans: [],
    computed: null,
    rankBy: 'gun',
    resultsCatFilter: 'ALL',
    resultsSearch: '',
    rosterSearch: '',
    editMode: 'scan'
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
    }, 3400);
  }

  /* ========================================================================
   * 初始化與路由切換
   * ======================================================================== */
  async function initApp() {
    registerServiceWorker();
    await DB.init();

    // 若網址帶入 ?api=<GAS_URL>，自動寫入本機裝置設定並開啟多機連線
    const params = new URLSearchParams(location.search);
    const paramApi = (params.get('api') || '').trim();
    if (paramApi) {
      DB.saveDeviceSettings({ gasUrl: paramApi, mode: 'online' });
      showToast('已自動套用雲端同步網址，請登入您的作業帳號', 'info');
    }

    const settings = DB.getDeviceSettings();
    NFC.sound.enabled = settings.soundEnabled !== false;
    if (settings.highContrast) document.body.classList.add('high-contrast');

    bindEvents();
    startLiveClock();
    renderNfcDiagnostics();

    // 還原同分頁登入階段（防止賽道現場誤重新整理遺失鎖定站點）
    const savedSession = DB.getSession();
    if (savedSession && savedSession.mode && savedSession.account) {
      state.session = savedSession;
      if (savedSession.mode === 'admin') {
        state.view = 'admin-workspace';
      } else {
        state.view = 'staff-workspace';
        state.staffSubTab = savedSession.account.role === 'writer' ? 'writer' : 'scan';
      }
    } else {
      state.view = 'portal';
    }

    await reloadState();
    switchView(state.view);

    // 網路恢復與斷線事件監聽
    window.addEventListener('online', async () => {
      updateNetworkIndicators();
      const pending = await DB.getPendingSyncCounts();
      if (Sync.isOnlineEnabled() && pending.total > 0) {
        showToast(`偵測到網路已恢復！自動上傳 ${pending.total} 筆離線資料中...`, 'info');
        await handleUploadPendingSync(false);
      } else if (pending.total > 0 && Sync.hasUrl()) {
        showToast(`網路已恢復！您有 ${pending.total} 筆離線資料可點擊「統一上傳同步」`, 'info');
      }
    });

    window.addEventListener('offline', () => {
      updateNetworkIndicators();
      showToast('已進入無網路狀態！晶片計時作業可繼續離線執行，資料安全暫存本機', 'warn');
    });
  }

  function registerServiceWorker() {
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
      navigator.serviceWorker.register('./sw.js').catch((err) => {
        console.warn('Service Worker 註冊失敗:', err);
      });
    }
  }

  function switchView(viewName) {
    // 離開作業區時停止實體 NFC 掃描
    if (viewName !== 'staff-workspace' && NFC.isScanning) {
      NFC.stopScanning();
      resetScanButtonUi();
    }

    state.view = viewName;
    document.querySelectorAll('.app-view').forEach((v) => v.classList.add('hidden'));
    const target = $(`view-${viewName}`);
    if (target) target.classList.remove('hidden');

    renderHeader();
    updateNetworkIndicators();
    if (viewName === 'staff-workspace') renderStaffWorkspace();
    if (viewName === 'admin-workspace') renderAdminWorkspace();
    window.scrollTo({ top: 0, behavior: 'instant' });
  }

  function switchAdminTab(tabName) {
    state.adminTab = tabName;
    document.querySelectorAll('.admin-panel').forEach((p) => p.classList.add('hidden'));
    const target = $(`admin-tab-${tabName}`);
    if (target) target.classList.remove('hidden');

    document.querySelectorAll('.admin-tab-btn').forEach((btn) => {
      const active = btn.dataset.adminTab === tabName;
      btn.className = active
        ? 'admin-tab-btn px-4 py-2.5 rounded-xl text-xs sm:text-sm font-black bg-amber-500 text-slate-950'
        : 'admin-tab-btn px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800';
    });
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
    updateNetworkIndicators();
    if (state.view === 'staff-workspace') renderStaffWorkspace();
    if (state.view === 'admin-workspace') renderAdminWorkspace();
    await updatePendingSyncBadge();
  }

  function startLiveClock() {
    const clockEl = $('live-clock');
    function tick() {
      if (clockEl) clockEl.textContent = TE.formatClockMs(Date.now());
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
        : '⚠️ 非 Web NFC 環境（可使用下方模擬感應或手動補登）';
    }
    if (writerText) writerText.textContent = diag.statusMessage;
    if (writerBanner && !diag.supported) {
      writerBanner.className = 'rounded-2xl border border-amber-500/40 bg-amber-950/30 p-4 text-sm text-amber-200 flex items-start gap-3';
    }
  }

  /* ========================================================================
   * 頂部與網路狀態渲染
   * ======================================================================== */
  function renderHeader() {
    const settings = DB.getDeviceSettings();
    $('header-race-name').textContent = state.config?.raceName || '智慧路跑 NFC 晶片計時系統';
    $('header-race-date').textContent = state.config?.raceDate || '--';
    $('header-device-id').textContent = settings.deviceId;

    // 公開查詢頁連結自動附帶 api
    const pubLink = $('portal-link-public');
    if (pubLink) {
      pubLink.href = settings.gasUrl
        ? `./public.html?api=${encodeURIComponent(settings.gasUrl)}`
        : './public.html';
    }

    // 登入身分顯示
    const userBox = $('header-user-box');
    const userLabel = $('header-user-label');
    if (state.session) {
      userBox.classList.remove('hidden');
      const acc = state.session.account;
      userLabel.textContent = state.session.mode === 'admin'
        ? `🛡️ ${acc.username} (最高管理員)`
        : `👷 ${acc.displayName || acc.username}`;
    } else {
      userBox.classList.add('hidden');
    }

    const soundBtn = $('btn-toggle-sound');
    const contrastBtn = $('btn-toggle-contrast');
    if (soundBtn) soundBtn.textContent = NFC.sound.enabled ? '🔊 音效開' : '🔇 音效關';
    if (contrastBtn) contrastBtn.textContent = settings.highContrast ? '🌙 一般模式' : '☀️ 強光高對比';
  }

  function updateNetworkIndicators() {
    const settings = DB.getDeviceSettings();
    const onlineNet = navigator.onLine;
    const badge = $('badge-network');

    if (settings.mode === 'online' && settings.gasUrl) {
      if (onlineNet) {
        badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-sky-400 animate-pulse"></span><span>多機雲端連線</span>';
        badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-sky-950 border border-sky-500/50 text-sky-300';
      } else {
        badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-amber-400"></span><span>無網路・離線暫存</span>';
        badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-amber-950 border border-amber-500/50 text-amber-300';
      }
    } else {
      badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-emerald-400"></span><span>單機離線模式</span>';
      badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-slate-800 border border-slate-700 text-slate-300';
    }

    // 大廳網路橫幅
    const portalBanner = $('portal-net-banner');
    if (portalBanner) {
      if (onlineNet) {
        portalBanner.className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-xs font-extrabold bg-emerald-950/70 border border-emerald-500/40 text-emerald-300';
        portalBanner.textContent = settings.gasUrl
          ? '🟢 網路連線正常 • 已綁定 Google 雲端試算表，亦支援無網路離線作業'
          : '🟢 網路連線正常 • 目前為本機離線資料庫（可由後台設定 Google 試算表同步）';
      } else {
        portalBanner.className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-xs font-extrabold bg-amber-950/80 border border-amber-500/50 text-amber-300';
        portalBanner.textContent = '📴 目前無網路連線 • 晶片計時作業仍可正常離線登入與感應，有網後統一上傳同步！';
      }
    }

    // 作業員登入頁網路說明
    const loginHint = $('staff-login-offline-hint');
    if (loginHint) {
      if (!onlineNet) {
        loginHint.className = 'rounded-xl border border-amber-500/40 bg-amber-950/40 p-3 text-xs text-amber-200 font-bold';
        loginHint.textContent = '📴 無網路離線模式：將自動使用本機快取帳號驗證登入，感應紀錄將安全暫存於本機 IndexedDB，有網路後可一鍵上傳同步。';
      } else if (settings.gasUrl) {
        loginHint.className = 'rounded-xl border border-sky-500/40 bg-sky-950/40 p-3 text-xs text-sky-200';
        loginHint.textContent = '🌐 雲端驗證就緒：登入時將自動核對最新站點指派，並快取至本機供後續賽道斷訊離線使用。';
      } else {
        loginHint.className = 'rounded-xl border border-emerald-500/30 bg-emerald-950/30 p-3 text-xs text-emerald-200';
        loginHint.textContent = '💾 本機單機模式：使用本機資料庫驗證作業員帳號並鎖定對應站點。';
      }
    }

    // 作業員工作區離線/連線切換鈕
    const toggleBtn = $('btn-staff-toggle-mode');
    if (toggleBtn) {
      if (settings.mode === 'online' && settings.gasUrl) {
        toggleBtn.textContent = onlineNet
          ? '🌐 目前：雲端即時連線（點此切離線）'
          : '⚠️ 連線模式（等待網路訊號中）';
        toggleBtn.className = onlineNet
          ? 'px-3.5 py-2 rounded-xl text-xs font-black bg-sky-600/30 hover:bg-sky-600/50 border border-sky-500/60 text-sky-300 transition'
          : 'px-3.5 py-2 rounded-xl text-xs font-black bg-amber-600/30 border border-amber-500/60 text-amber-300 transition';
      } else {
        toggleBtn.textContent = '📴 目前：離線作業（點此切雲端同步）';
        toggleBtn.className = 'px-3.5 py-2 rounded-xl text-xs font-black bg-slate-800 hover:bg-slate-700 border border-slate-600 text-slate-200 transition';
      }
    }
  }

  async function updatePendingSyncBadge() {
    const pending = await DB.getPendingSyncCounts();
    const settings = DB.getDeviceSettings();

    $('pending-sync-count').textContent = String(pending.total);
    const staffCountEl = $('staff-pending-count');
    if (staffCountEl) staffCountEl.textContent = String(pending.total);

    const quickBtn = $('btn-quick-sync');
    if (pending.total > 0 && settings.gasUrl) {
      quickBtn.classList.remove('hidden');
    } else {
      quickBtn.classList.add('hidden');
    }

    const syncInfo = $('sync-status-info');
    if (syncInfo) {
      const lastStr = settings.lastSyncAt ? TE.formatClockSec(settings.lastSyncAt) : '尚未同步';
      syncInfo.textContent = `待同步資料：選手 ${pending.runners} 筆、感應 ${pending.scans} 筆 ｜ 最後同步時間：${lastStr}`;
    }
  }

  /* ========================================================================
   * 登入與登出流程（線上驗證＋快取，斷網無縫切換離線驗證）
   * ======================================================================== */
  async function handleLogin(mode, username, password) {
    const u = String(username || '').trim();
    const p = String(password || '');
    if (!u || !p) return showToast('請輸入帳號與密碼', 'warn');

    let account = null;
    let usedCloud = false;

    // 1. 有網路且有設 GAS URL 時，優先透過雲端驗證並同步最新設定
    if (navigator.onLine && Sync.hasUrl()) {
      try {
        const res = await Sync.loginCloud(mode, u, p);
        account = { ...res.account, password: p };
        usedCloud = true;
      } catch (err) {
        // 若明確回傳帳密錯誤就不回退；若為網路逾時/連線失敗則降級至離線驗證
        if (String(err.message).includes('錯誤') || String(err.message).includes('停用')) {
          NFC.sound.beepError();
          showToast(err.message, 'error');
          return;
        }
        console.warn('雲端登入連線失敗，改採本機離線驗證:', err);
      }
    }

    // 2. 無網路、單機模式、或雲端暫時連不上時 -> 本機 IndexedDB 離線驗證
    if (!account) {
      const localRes = await DB.verifyLocalLogin(mode, u, p);
      if (!localRes.ok) {
        NFC.sound.beepError();
        showToast(localRes.error, 'error');
        return;
      }
      account = { ...localRes.account, password: p };
    }

    state.session = {
      mode,
      account,
      auth: { username: u, password: p },
      loginAt: Date.now()
    };
    DB.saveSession(state.session);
    await reloadState();

    if (mode === 'admin') {
      showToast(usedCloud ? '🛡️ 已通過雲端驗證並同步最新賽事資料' : '🛡️ 已登入最高權限後台（本機模式）', 'success');
      switchView('admin-workspace');
      switchAdminTab('results');
    } else {
      state.staffSubTab = account.role === 'writer' ? 'writer' : 'scan';
      const cp = (state.config.checkpoints || []).find((c) => c.id === account.checkpointId);
      const zoneName = account.role === 'writer'
        ? 'NFC 晶片寫入與報到區'
        : (cp ? cp.name : account.checkpointId || '未指定站點');
      showToast(`✅ 登入成功！已鎖定進入「${zoneName}」${usedCloud ? '（已同步最新設定）' : '（離線可用）'}`, 'success');
      switchView('staff-workspace');
    }
  }

  function handleLogout() {
    if (NFC.isScanning) {
      NFC.stopScanning();
      resetScanButtonUi();
    }
    state.session = null;
    DB.saveSession(null);
    showToast('已登出並返回作業區大廳', 'info');
    switchView('portal');
  }

  /* ========================================================================
   * VIEW 4：晶片計時作業區（鎖定站點 / 鎖定寫入區）
   * ======================================================================== */
  function getLockedCheckpoint() {
    const acc = state.session?.account;
    if (!acc || !acc.checkpointId) return null;
    return (state.config?.checkpoints || []).find((c) => c.id === acc.checkpointId) || {
      id: acc.checkpointId,
      name: acc.checkpointId,
      order: 0,
      distanceKm: 0,
      type: 'checkpoint'
    };
  }

  function renderStaffWorkspace() {
    if (!state.session || state.session.mode !== 'staff') return;
    const acc = state.session.account;
    const role = acc.role || 'scanner';
    const lockedCp = getLockedCheckpoint();

    // 頂部鎖定身分標題
    let zoneTitle = '';
    if (role === 'writer') {
      zoneTitle = 'NFC 晶片寫入與報到區';
    } else if (role === 'both') {
      zoneTitle = `晶片寫入 ＋ 鎖定感應站點：${lockedCp ? lockedCp.name : '未指定'}`;
    } else {
      zoneTitle = `鎖定感應站點：${lockedCp ? lockedCp.name : '未指派站點（請聯繫管理員）'}`;
    }
    $('staff-locked-zone-title').textContent = zoneTitle;
    $('staff-operator-meta').textContent = `操作員：${acc.displayName || acc.username} (${acc.username}) ｜ 權限不可自行切換站點`;

    // 依權限強制顯示對應面板
    const bothTabs = $('staff-both-tabs');
    const panelScan = $('staff-panel-scan');
    const panelWriter = $('staff-panel-writer');

    if (role === 'writer') {
      bothTabs.classList.add('hidden');
      panelScan.classList.add('hidden');
      panelWriter.classList.remove('hidden');
      renderWriterPanel();
    } else if (role === 'scanner') {
      bothTabs.classList.add('hidden');
      panelWriter.classList.add('hidden');
      panelScan.classList.remove('hidden');
      renderScanPanel(lockedCp);
    } else {
      // role === 'both'
      bothTabs.classList.remove('hidden');
      document.querySelectorAll('.staff-sub-btn').forEach((b) => {
        const active = b.dataset.staffSub === state.staffSubTab;
        b.className = active
          ? 'staff-sub-btn flex-1 py-3 rounded-xl font-black text-sm bg-emerald-500 text-slate-950'
          : 'staff-sub-btn flex-1 py-3 rounded-xl font-black text-sm bg-slate-800 text-slate-300';
      });
      panelScan.classList.toggle('hidden', state.staffSubTab !== 'scan');
      panelWriter.classList.toggle('hidden', state.staffSubTab !== 'writer');
      if (state.staffSubTab === 'scan') renderScanPanel(lockedCp);
      else renderWriterPanel();
    }
  }

  function renderScanPanel(lockedCp) {
    $('locked-cp-name').textContent = lockedCp ? `#${lockedCp.order || '-'} ${lockedCp.name}` : '未指派站點';
    $('locked-cp-meta').textContent = lockedCp ? `代碼：${lockedCp.id} ｜ 累計距離：${lockedCp.distanceKm} km` : '請聯繫後台管理員指派站點';

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

    renderRecentCheckpointScans(lockedCp?.id || '');
  }

  function renderRecentCheckpointScans(cpId) {
    const runnerMap = new Map(state.runners.map((r) => [r.bib, r]));
    const catMap = new Map((state.config.categories || []).map((c) => [c.id, c]));

    const list = state.scans
      .filter((s) => !s.deleted && s.checkpointId === cpId)
      .sort((a, b) => b.timestampMs - a.timestampMs);

    $('checkpoint-scan-count').textContent = `${list.length} 筆`;
    const tbody = $('recent-scans-tbody');
    if (list.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="py-6 text-center text-slate-500">本鎖定站點尚無感應紀錄，請啟動感應或使用模擬器測試</td></tr>';
      return;
    }

    const sourceBadges = {
      nfc: '<span class="px-2 py-0.5 rounded bg-emerald-900/60 text-emerald-300 text-xs font-bold">NFC</span>',
      manual: '<span class="px-2 py-0.5 rounded bg-sky-900/60 text-sky-300 text-xs font-bold">手動</span>',
      mock: '<span class="px-2 py-0.5 rounded bg-indigo-900/60 text-indigo-300 text-xs font-bold">模擬</span>',
      admin: '<span class="px-2 py-0.5 rounded bg-amber-900/60 text-amber-300 text-xs font-bold">修正</span>'
    };

    tbody.innerHTML = list
      .slice(0, 20)
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
              ${s.synced ? '☁️ 已同步' : '💾 離線暫存'}
            </td>
          </tr>
        `;
      })
      .join('');
  }

  async function processIncomingScan(tagData, { bypassDebounce = false } = {}) {
    const lockedCp = getLockedCheckpoint();
    if (!lockedCp || !lockedCp.id) {
      NFC.sound.beepError();
      showToast('此帳號尚未指派鎖定站點，無法執行感應，請聯繫後台管理員', 'error');
      return;
    }

    const scanTs = Date.now();
    const operator = state.session?.account?.username || '';

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
      checkpointId: lockedCp.id,
      timestampMs: scanTs,
      source: tagData.source || 'nfc',
      operator,
      note: tagData.note || '',
      bypassDebounce
    });

    const cat = (state.config.categories || []).find((c) => c.id === (runner?.categoryId || tagData.categoryId));
    const displayName = runner?.name || tagData.name || `選手 ${tagData.bib}`;
    const displayCat = cat ? cat.name : (runner?.categoryId || tagData.categoryId || '未分組');

    const heroCard = $('scan-hero-card');
    heroCard.classList.remove('flash-success', 'flash-warn', 'flash-error');
    void heroCard.offsetWidth;

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
    $('scan-status-pill').textContent = `✅ 感應成功 — ${lockedCp.name}`;
    $('hero-bib-cat').textContent = `${tagData.bib} • ${displayCat}`;
    $('hero-name').textContent = displayName;
    $('hero-elapsed').textContent = elapsedStr;
    $('hero-subinfo').textContent = `感應時間：${TE.formatClockMs(scanTs)} ｜ 操作員：${operator} ｜ 來源：${(tagData.source || 'nfc').toUpperCase()}`;

    if (Sync.isOnlineEnabled() && navigator.onLine) {
      Sync.pushSingleScan(res.scan).then(() => reloadState());
    }
    await reloadState();
  }

  function resetScanButtonUi() {
    const btn = $('btn-toggle-nfc-scan');
    if (!btn) return;
    btn.textContent = '📡 啟動極速 NFC 感應';
    btn.className = 'no-select w-full sm:w-auto sm:min-w-[280px] py-5 px-8 rounded-2xl text-xl sm:text-2xl font-black bg-emerald-500 hover:bg-emerald-400 active:scale-[0.98] text-slate-950 shadow-lg transition';
  }

  async function toggleNfcScanning() {
    const btn = $('btn-toggle-nfc-scan');
    if (NFC.isScanning) {
      NFC.stopScanning();
      resetScanButtonUi();
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

  /* ---------------- 晶片寫入面板 ---------------- */
  function renderWriterPanel() {
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
      tbody.innerHTML = '<tr><td colspan="5" class="py-6 text-center text-slate-500">尚無選手資料，請由左側表單報到寫入</td></tr>';
      return;
    }

    tbody.innerHTML = filtered
      .map(
        (r) => `
        <tr class="hover:bg-slate-800/40">
          <td class="py-2 pr-2 font-mono-num font-black text-amber-300">${esc(r.bib)}</td>
          <td class="py-2 pr-2 font-bold text-white">${esc(r.name)}</td>
          <td class="py-2 pr-2 text-slate-300">${esc(catMap.get(r.categoryId) || r.categoryId)}</td>
          <td class="py-2 pr-2 font-mono-num text-[11px] ${r.nfcUid ? 'text-emerald-400' : 'text-slate-500'}">
            ${esc(r.nfcUid ? '✅ 已寫入' : '未寫入')}
          </td>
          <td class="py-2 text-right">
            <button type="button" data-load-runner="${esc(r.bib)}"
              class="px-2.5 py-1 rounded-lg text-xs font-bold bg-sky-600/30 hover:bg-sky-600/50 text-sky-300">
              帶入寫入
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
        { ...runner, nfcUid: result.nfcUid, checkedInAt: Date.now() },
        false
      );

      setWriterVerifyBox(
        'success',
        `✅ 晶片寫入並自動讀取驗證成功！（UID: ${result.nfcUid || '標準 NDEF'}）`,
        `讀回驗證資料：${result.data.raw}`
      );
      showToast(`選手 ${runner.bib} ${runner.name} 報到寫入完成！`, 'success');

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        Sync.pushRunners([saved]).then(() => updatePendingSyncBadge());
      }
      await reloadState();
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ 寫入或驗證失敗', err.message);
      showToast(err.message, 'error');
    }
  }

  /* ========================================================================
   * VIEW 5：後台管理區（最高權限）
   * ======================================================================== */
  function renderAdminWorkspace() {
    if (!state.session || state.session.mode !== 'admin') return;
    renderAdminResults();
    renderAdminStaffAccounts();
    renderAdminRaceSettings();
    renderAdminSyncPanel();
  }

  function renderAdminResults() {
    const computed = state.computed;
    if (!computed) return;

    $('kpi-total').textContent = String(computed.summary.totalRunners);
    $('kpi-finished').textContent = String(computed.summary.finishedCount);
    $('kpi-oncourse').textContent = String(computed.summary.onCourseCount);
    $('kpi-notstarted').textContent = String(computed.summary.notStartedCount);

    const catFilter = $('results-category-filter');
    catFilter.innerHTML =
      '<option value="ALL">全部組別</option>' +
      (state.config.categories || []).map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    catFilter.value = state.resultsCatFilter;
    $('results-rankby-select').value = state.rankBy;

    const checkpoints = computed.checkpoints;
    $('results-thead').innerHTML = `
      <tr>
        <th class="px-3 py-3">總名次</th>
        <th class="px-3 py-3">組名次</th>
        <th class="px-3 py-3">編號</th>
        <th class="px-3 py-3">姓名 (點此編輯)</th>
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
      tbody.innerHTML = `<tr><td colspan="${11 + checkpoints.length}" class="py-8 text-center text-slate-500">目前尚無符合條件的選手成績，可點選「建立選手」或於「備份還原」載入模擬資料</td></tr>`;
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
            const label = s ? TE.formatClockMs(s.timestampMs) : '+ 補登';
            const editedTag = s && s.source === 'admin' ? ' <span class="text-amber-400">*</span>' : '';
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
          })
          .join('');

        const splitSummary = row.splits
          .filter((sp) => sp.durationMs !== null)
          .map((sp) => `<div class="text-[11px] font-mono-num text-slate-300">${esc(sp.toName)}: <span class="text-sky-300">${TE.formatDuration(sp.durationMs, false)}</span> (${sp.pace})</div>`)
          .join('') || '<span class="text-slate-600">--</span>';

        return `
          <tr class="hover:bg-slate-800/40 align-top">
            <td class="px-3 py-2.5 font-mono-num font-black text-amber-400">${row.overallRank ?? '-'}</td>
            <td class="px-3 py-2.5 font-mono-num font-bold text-sky-300">${row.categoryRank ?? '-'}</td>
            <td class="px-3 py-2.5 font-mono-num font-bold text-white">${esc(row.bib)}</td>
            <td class="px-3 py-2.5">
              <button type="button" data-admin-edit-runner="${esc(row.bib)}" class="font-bold text-white hover:text-amber-300 underline decoration-dotted">
                ${esc(row.name)}
              </button>
            </td>
            <td class="px-3 py-2.5">
              <span class="inline-flex items-center gap-1.5 text-xs font-bold">
                <span class="w-2 h-2 rounded-full" style="background:${esc(row.categoryColor)}"></span>
                ${esc(row.categoryName)}
              </span>
            </td>
            <td class="px-3 py-2.5">${statusMap[row.status] || row.status}</td>
            ${cpCells}
            <td class="px-3 py-2.5">${splitSummary}</td>
            <td class="px-3 py-2.5">
              <button type="button" data-edit-penalty-bib="${esc(row.bib)}" class="px-2 py-1 rounded border border-dashed border-amber-500/50 text-amber-300 font-mono-num hover:bg-amber-900/30">
                ${row.penaltySec ? `${row.penaltySec > 0 ? '+' : ''}${row.penaltySec}s` : '0s'}
              </button>
            </td>
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

  function renderAdminStaffAccounts() {
    const list = state.config.staffAccounts || [];
    const cps = [...(state.config.checkpoints || [])].sort((a, b) => a.order - b.order);
    const container = $('staff-accounts-list');

    if (list.length === 0) {
      container.innerHTML = '<div class="text-center text-sm text-slate-500 py-6">尚無作業人員帳號，請點擊上方「＋ 新增作業員帳號」</div>';
      return;
    }

    container.innerHTML = list
      .map(
        (s, idx) => `
        <div class="grid grid-cols-1 sm:grid-cols-12 gap-2.5 items-end p-3.5 rounded-xl bg-slate-950 border border-slate-800" data-staff-row="${idx}">
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">登入帳號</label>
            <input type="text" value="${esc(s.username)}" data-staff-field="username" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-mono-num font-bold text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">登入密碼</label>
            <input type="text" value="${esc(s.password)}" data-staff-field="password" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-mono-num text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">人員/組別名稱</label>
            <input type="text" value="${esc(s.displayName)}" data-staff-field="displayName" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-bold text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">作業權限類型</label>
            <select data-staff-field="role" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-bold text-white">
              <option value="scanner" ${s.role === 'scanner' ? 'selected' : ''}>鎖定指定感應站點</option>
              <option value="writer" ${s.role === 'writer' ? 'selected' : ''}>僅限 NFC 晶片寫入</option>
              <option value="both" ${s.role === 'both' ? 'selected' : ''}>晶片寫入＋鎖定站點</option>
            </select>
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">鎖定感應站點</label>
            <select data-staff-field="checkpointId" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-bold text-sky-300">
              <option value="">（無／僅寫入）</option>
              ${cps.map((cp) => `<option value="${esc(cp.id)}" ${s.checkpointId === cp.id ? 'selected' : ''}>#${cp.order} ${esc(cp.name)}</option>`).join('')}
            </select>
          </div>
          <div class="sm:col-span-2 flex items-center justify-between gap-2 pb-1">
            <label class="inline-flex items-center gap-1.5 text-xs font-bold text-slate-300 cursor-pointer">
              <input type="checkbox" data-staff-field="enabled" ${s.enabled !== false ? 'checked' : ''} class="accent-emerald-500 w-4 h-4" />
              <span>啟用</span>
            </label>
            <button type="button" data-del-staff="${idx}" class="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>
          </div>
        </div>
      `
      )
      .join('');
  }

  function renderAdminRaceSettings() {
    $('cfg-race-name').value = state.config.raceName || '';
    $('cfg-race-date').value = state.config.raceDate || '';
    $('cfg-debounce-sec').value = String(state.config.debounceSeconds ?? 30);
    $('cfg-admin-user').value = state.config.adminUsername || 'admin';
    $('cfg-admin-pass').value = state.config.adminPassword || 'admin1234';

    // 分組鳴槍列表
    $('categories-list').innerHTML = (state.config.categories || [])
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
                🔫 一鍵鳴槍起跑
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

    // 站點清單
    const sortedCps = [...(state.config.checkpoints || [])].sort((a, b) => a.order - b.order);
    $('checkpoints-list').innerHTML = sortedCps
      .map(
        (cp, idx) => `
        <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-center p-3 rounded-xl bg-slate-950 border border-slate-800" data-cp-row="${idx}">
          <div class="sm:col-span-1">
            <input type="number" value="${esc(cp.order)}" data-cp-field="order" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-sm font-mono-num text-white" />
          </div>
          <div class="sm:col-span-2">
            <input type="text" value="${esc(cp.id)}" data-cp-field="id" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-mono-num text-white" />
          </div>
          <div class="sm:col-span-4">
            <input type="text" value="${esc(cp.name)}" data-cp-field="name" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-bold text-white" />
          </div>
          <div class="sm:col-span-2">
            <input type="number" step="0.1" value="${esc(cp.distanceKm)}" data-cp-field="distanceKm" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-mono-num text-white" />
          </div>
          <div class="sm:col-span-2">
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

  function renderAdminSyncPanel() {
    const settings = DB.getDeviceSettings();
    document.querySelectorAll('input[name="sync-mode"]').forEach((radio) => {
      radio.checked = radio.value === settings.mode;
    });
    $('input-gas-url').value = settings.gasUrl || '';
  }

  /* ========================================================================
   * 同步、備份與匯入匯出工具函式
   * ======================================================================== */
  async function syncConfigIfOnline(cfg) {
    if (Sync.isOnlineEnabled() && navigator.onLine && state.session?.mode === 'admin') {
      try {
        await Sync.pushConfig(cfg, state.session.auth);
        showToast('設定已即時同步至 Google 試算表', 'success');
      } catch (err) {
        showToast(`本機已儲存，但雲端同步失敗：${err.message}`, 'warn');
      }
    }
  }

  async function handleUploadPendingSync(silent = false) {
    if (!navigator.onLine) {
      return showToast('目前仍無網路連線，請待連上網路後再統一上傳', 'warn');
    }
    if (!Sync.hasUrl()) {
      return showToast('尚未設定 Google Apps Script 雲端網址，請由後台管理區設定', 'warn');
    }
    try {
      const res = await Sync.syncPendingToCloud();
      await reloadState();
      if (!silent || res.uploadedScans > 0 || res.uploadedRunners > 0) {
        showToast(
          `☁️ 統一上傳同步完成！選手 ${res.uploadedRunners} 筆、感應 ${res.uploadedScans} 筆` +
            (res.skippedDuplicate ? `（雲端防連點略過 ${res.skippedDuplicate} 筆）` : ''),
          'success'
        );
      } else {
        showToast('目前沒有待上傳的離線紀錄，資料已是最新狀態', 'info');
      }
    } catch (err) {
      showToast(`雲端上傳失敗：${err.message}`, 'error');
    }
  }

  async function downloadJsonBackup(prefix = '智慧路跑完整備份') {
    const backup = await DB.exportFullBackup();
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    a.href = url;
    a.download = `${prefix}_${ts}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast(`已下載備份檔（含 ${backup.runners.length} 位選手、${backup.scans.length} 筆感應）`, 'success');
  }

  async function handleImportBackupFile(file, strategy) {
    if (!file) return;
    try {
      const text = await file.text();
      const json = JSON.parse(text);
      const res = await DB.importFullBackup(json, strategy);
      await reloadState();
      showToast(
        `${strategy === 'replace' ? '覆蓋還原' : '合併彙整'}完成！匯入選手 ${res.runnerCount} 筆、感應 ${res.scanCount} 筆`,
        'success'
      );
    } catch (err) {
      showToast(`備份檔匯入失敗：${err.message}`, 'error');
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
    if (valid.length === 0) return showToast('未偵測到有效的選手名單資料', 'warn');
    const saved = await DB.upsertRunnersBulk(valid, false);
    if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRunners(saved).then(() => updatePendingSyncBadge());
    await reloadState();
    showToast(`已成功匯入 ${valid.length} 位選手名單！`, 'success');
  }

  /* ---------------- Modal：人工修正感應時間 / 罰時 / 選手資料 ---------------- */
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

  function openAdminRunnerModal(bib = '') {
    const existing = bib ? state.runners.find((r) => r.bib === bib) : null;
    $('modal-runner-title').textContent = existing ? `編輯選手：${existing.bib}` : '建立新選手資料';
    $('admin-runner-bib').value = existing?.bib || '';
    $('admin-runner-bib').readOnly = Boolean(existing);
    $('admin-runner-name').value = existing?.name || '';
    $('admin-runner-id').value = existing?.idNumber || '';
    $('admin-runner-cat').innerHTML = (state.config.categories || [])
      .map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`)
      .join('');
    if (existing?.categoryId) $('admin-runner-cat').value = existing.categoryId;
    $('btn-admin-delete-runner').classList.toggle('hidden', !existing);
    $('modal-edit-runner').classList.remove('hidden');
  }

  /* ========================================================================
   * 全部互動事件綁定
   * ======================================================================== */
  function bindEvents() {
    // 大廳導覽按鈕
    $('btn-portal-staff').addEventListener('click', () => switchView('staff-login'));
    $('btn-portal-admin').addEventListener('click', () => switchView('admin-login'));
    $('btn-logout').addEventListener('click', handleLogout);
    $('header-brand-home').addEventListener('click', () => {
      if (!state.session) switchView('portal');
    });
    document.querySelectorAll('[data-goto]').forEach((btn) => {
      btn.addEventListener('click', () => switchView(btn.dataset.goto));
    });

    // 頂部待同步快速上傳
    $('btn-quick-sync').addEventListener('click', () => handleUploadPendingSync(false));

    // 作業員登入頁快速填入測試帳號
    document.querySelectorAll('[data-quick-staff]').forEach((btn) => {
      btn.addEventListener('click', () => {
        $('staff-login-user').value = btn.dataset.quickStaff;
        $('staff-login-pass').value = '1111';
      });
    });

    $('form-staff-login').addEventListener('submit', (e) => {
      e.preventDefault();
      handleLogin('staff', $('staff-login-user').value, $('staff-login-pass').value);
    });

    $('form-admin-login').addEventListener('submit', (e) => {
      e.preventDefault();
      handleLogin('admin', $('admin-login-user').value, $('admin-login-pass').value);
    });

    // 作業員工作區：離線/連線切換、統一上傳、匯出離線備份檔
    $('btn-staff-toggle-mode').addEventListener('click', async () => {
      const settings = DB.getDeviceSettings();
      if (settings.mode === 'offline') {
        if (!settings.gasUrl) {
          return showToast('本機尚未設定雲端網址，請由後台管理區設定或透過工作人員連線連結開啟', 'warn');
        }
        DB.saveDeviceSettings({ mode: 'online' });
        updateNetworkIndicators();
        showToast('已切換為「雲端即時連線模式」', 'success');
        const pending = await DB.getPendingSyncCounts();
        if (pending.total > 0 && navigator.onLine) {
          await handleUploadPendingSync(false);
        }
      } else {
        DB.saveDeviceSettings({ mode: 'offline' });
        updateNetworkIndicators();
        showToast('已切換為「單機離線作業模式」（所有感應安全暫存本機）', 'info');
      }
    });

    $('btn-staff-upload-sync').addEventListener('click', () => handleUploadPendingSync(false));

    $('btn-staff-export-offline-json').addEventListener('click', () => {
      const op = state.session?.account?.username || 'staff';
      const cp = state.session?.account?.checkpointId || 'writer';
      downloadJsonBackup(`離線感應備份_${op}_${cp}`);
    });

    // 若 role==='both' 切換子頁籤
    document.querySelectorAll('.staff-sub-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.staffSubTab = btn.dataset.staffSub;
        renderStaffWorkspace();
      });
    });

    // 音效與強光高對比
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

    // NFC 感應啟動/停止
    $('btn-toggle-nfc-scan').addEventListener('click', toggleNfcScanning);

    // 手動補感應
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

    // 模擬器
    $('btn-load-demo-shortcut').addEventListener('click', async () => {
      await DB.loadDemoData();
      await reloadState();
      showToast('已載入 9 位模擬選手資料，可直接點擊「模擬晶片刷過」！', 'success');
    });
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

    // 晶片寫入表單
    $('form-nfc-write').addEventListener('submit', (e) => {
      e.preventDefault();
      executeNfcWrite(false);
    });
    $('btn-mock-write-nfc').addEventListener('click', () => executeNfcWrite(true));
    $('btn-save-runner-only').addEventListener('click', async () => {
      const r = getWriterFormRunner();
      if (!r.bib || !r.name) return showToast('請輸入參賽編號與姓名', 'warn');
      const saved = await DB.upsertRunner(r, false);
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRunners([saved]);
      await reloadState();
      showToast(`已儲存選手 ${r.bib} ${r.name}`, 'success');
    });

    $('roster-search').addEventListener('input', (e) => {
      state.rosterSearch = e.target.value;
      renderWriterPanel();
    });
    $('roster-tbody').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-load-runner]');
      if (!btn) return;
      const r = state.runners.find((x) => x.bib === btn.dataset.loadRunner);
      if (!r) return;
      $('writer-bib').value = r.bib;
      $('writer-name').value = r.name;
      $('writer-idnumber').value = r.idNumber || '';
      $('writer-category').value = r.categoryId;
      showToast(`已帶入 ${r.bib} ${r.name} 至晶片寫入表單`, 'info');
    });

    /* ==================== 後台管理區事件 ==================== */
    document.querySelectorAll('.admin-tab-btn').forEach((btn) => {
      btn.addEventListener('click', () => switchAdminTab(btn.dataset.adminTab));
    });

    // 成績表篩選與修正
    $('results-category-filter').addEventListener('change', (e) => {
      state.resultsCatFilter = e.target.value;
      renderAdminResults();
    });
    $('results-rankby-select').addEventListener('change', async (e) => {
      state.rankBy = e.target.value;
      await reloadState();
    });
    $('results-search').addEventListener('input', (e) => {
      state.resultsSearch = e.target.value;
      renderAdminResults();
    });

    $('btn-admin-add-runner').addEventListener('click', () => openAdminRunnerModal(''));
    $('btn-export-excel-quick').addEventListener('click', () => {
      const f = Excel.exportResultsToExcel({ config: state.config, computed: state.computed, scans: state.scans, maskId: false });
      showToast(`已匯出 Excel 報表：${f}`, 'success');
    });

    $('results-tbody').addEventListener('click', (e) => {
      const scanBtn = e.target.closest('[data-edit-scan-bib]');
      if (scanBtn) {
        openEditScanModal(scanBtn.dataset.editScanBib, scanBtn.dataset.editScanCp, scanBtn.dataset.editScanId);
        return;
      }
      const penBtn = e.target.closest('[data-edit-penalty-bib]');
      if (penBtn) {
        openEditPenaltyModal(penBtn.dataset.editPenaltyBib);
        return;
      }
      const editRunnerBtn = e.target.closest('[data-admin-edit-runner]');
      if (editRunnerBtn) {
        openAdminRunnerModal(editRunnerBtn.dataset.adminEditRunner);
      }
    });

    // Modal 1：站點時間 / 罰時修正
    const closeEditModal = () => $('modal-edit-scan').classList.add('hidden');
    $('btn-close-edit-modal').addEventListener('click', closeEditModal);
    $('btn-cancel-edit-modal').addEventListener('click', closeEditModal);

    $('form-edit-scan').addEventListener('submit', async (e) => {
      e.preventDefault();
      const bib = $('edit-bib').value;

      if (state.editMode === 'penalty') {
        const sec = Number($('edit-penalty-sec').value || 0);
        const updated = await DB.upsertRunner({ bib, penaltySec: sec }, false);
        if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRunners([updated]);
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
      if (tsMs === null) return showToast('時間格式錯誤，請輸入 HH:mm:ss 或 HH:mm:ss.SSS', 'error');

      let savedScan;
      if (scanId) {
        const existing = state.scans.find((s) => s.id === scanId);
        savedScan = await DB.putScan({
          ...existing,
          timestampMs: tsMs,
          source: 'admin',
          operator: state.session?.account?.username || 'admin',
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
          operator: state.session?.account?.username || 'admin',
          note,
          bypassDebounce: true
        });
        savedScan = res.scan;
      }

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try {
          await Sync.pushAdminEditScan(savedScan, state.session.auth);
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
      if (!scanId || !confirm('確定要刪除此站點感應紀錄嗎？')) return;
      const existing = state.scans.find((s) => s.id === scanId);
      if (!existing) return;
      const updated = await DB.putScan({
        ...existing,
        deleted: true,
        source: 'admin',
        operator: state.session?.account?.username || 'admin',
        note: '管理員刪除誤感應',
        synced: 0
      });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try {
          await Sync.pushAdminEditScan(updated, state.session.auth);
        } catch {}
      }
      closeEditModal();
      await reloadState();
      showToast('已刪除該站感應並重新計算名次', 'info');
    });

    // Modal 2：最高管理員建立/編輯選手
    const closeRunnerModal = () => $('modal-edit-runner').classList.add('hidden');
    $('btn-close-runner-modal').addEventListener('click', closeRunnerModal);
    $('btn-cancel-runner-modal').addEventListener('click', closeRunnerModal);

    $('form-admin-runner').addEventListener('submit', async (e) => {
      e.preventDefault();
      const saved = await DB.upsertRunner(
        {
          bib: $('admin-runner-bib').value.trim(),
          name: $('admin-runner-name').value.trim(),
          idNumber: $('admin-runner-id').value.trim().toUpperCase(),
          categoryId: $('admin-runner-cat').value
        },
        false
      );
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRunners([saved]);
      closeRunnerModal();
      await reloadState();
      showToast(`已儲存選手 ${saved.bib} ${saved.name}`, 'success');
    });

    $('btn-admin-delete-runner').addEventListener('click', async () => {
      const bib = $('admin-runner-bib').value.trim();
      if (!bib || !confirm(`確定要刪除選手 ${bib} 嗎？`)) return;
      await DB.deleteRunner(bib);
      closeRunnerModal();
      await reloadState();
      showToast(`已刪除選手 ${bib}`, 'info');
    });

    // 後台分頁 2：作業人員帳號與站點管制 CRUD
    $('btn-add-staff').addEventListener('click', async () => {
      const list = [...(state.config.staffAccounts || [])];
      const n = list.length + 1;
      list.push({
        username: `staff${n}`,
        password: '1111',
        displayName: `工作人員 ${n}`,
        role: 'scanner',
        checkpointId: state.config.checkpoints?.[0]?.id || '',
        enabled: true,
        updatedAt: Date.now()
      });
      await DB.saveConfig({ ...state.config, staffAccounts: list });
      await reloadState();
      showToast('已新增一筆作業員帳號，請編輯後點擊「儲存作業員帳號」', 'info');
    });

    $('staff-accounts-list').addEventListener('click', async (e) => {
      const delBtn = e.target.closest('[data-del-staff]');
      if (!delBtn) return;
      const idx = Number(delBtn.dataset.delStaff);
      const list = [...(state.config.staffAccounts || [])];
      if (!confirm(`確定刪除作業員帳號「${list[idx]?.username}」嗎？`)) return;
      list.splice(idx, 1);
      const updated = await DB.saveConfig({ ...state.config, staffAccounts: list });
      await syncConfigIfOnline(updated);
      await reloadState();
    });

    $('btn-save-staff-accounts').addEventListener('click', async () => {
      const rows = Array.from(document.querySelectorAll('#staff-accounts-list [data-staff-row]'));
      const seen = new Set();
      const staffAccounts = [];
      for (const row of rows) {
        const username = row.querySelector('[data-staff-field="username"]').value.trim();
        const password = row.querySelector('[data-staff-field="password"]').value;
        if (!username) continue;
        if (seen.has(username)) {
          return showToast(`作業員帳號「${username}」重複，請修正後再儲存`, 'error');
        }
        seen.add(username);
        const role = row.querySelector('[data-staff-field="role"]').value;
        const cpId = row.querySelector('[data-staff-field="checkpointId"]').value;
        if (role !== 'writer' && !cpId) {
          return showToast(`帳號「${username}」屬於感應權限，必須指派鎖定站點！`, 'error');
        }
        staffAccounts.push({
          username,
          password: password || '1111',
          displayName: row.querySelector('[data-staff-field="displayName"]').value.trim() || username,
          role,
          checkpointId: role === 'writer' ? '' : cpId,
          enabled: row.querySelector('[data-staff-field="enabled"]').checked,
          updatedAt: Date.now()
        });
      }
      const updated = await DB.saveConfig({ ...state.config, staffAccounts });
      await syncConfigIfOnline(updated);
      await reloadState();
      showToast(`已儲存 ${staffAccounts.length} 組作業員帳號與鎖定站點設定！`, 'success');
    });

    // 後台分頁 3：賽事基本設定 + 管理員獨立帳密
    $('form-race-basic').addEventListener('submit', async (e) => {
      e.preventDefault();
      const newAdminUser = $('cfg-admin-user').value.trim() || 'admin';
      const newAdminPass = $('cfg-admin-pass').value || 'admin1234';
      const oldAuth = { ...state.session.auth };

      const updated = await DB.saveConfig({
        ...state.config,
        raceName: $('cfg-race-name').value.trim(),
        raceDate: $('cfg-race-date').value,
        debounceSeconds: Number($('cfg-debounce-sec').value || 30),
        adminUsername: newAdminUser,
        adminPassword: newAdminPass
      });

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try {
          await Sync.pushConfig(updated, oldAuth);
          showToast('賽事設定與管理員帳密已同步至雲端', 'success');
        } catch (err) {
          showToast(`本機已更新，但雲端同步失敗：${err.message}`, 'warn');
        }
      } else {
        showToast('賽事設定與管理員帳密已儲存', 'success');
      }

      // 更新目前登入階段的新帳密
      state.session.auth = { username: newAdminUser, password: newAdminPass };
      state.session.account.username = newAdminUser;
      DB.saveSession(state.session);
      await reloadState();
    });

    // 分組鳴槍與 CRUD
    $('btn-add-category').addEventListener('click', async () => {
      const cats = [...(state.config.categories || [])];
      const n = cats.length + 1;
      cats.push({ id: `CAT${n}`, name: `新分組 ${n}`, distanceKm: 5, gunTimeMs: null, gunTimeIso: '', color: '#a855f7' });
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

      const gunManualBtn = e.target.closest('[data-gun-manual-set]');
      if (gunManualBtn) {
        const idx = Number(gunManualBtn.dataset.gunManualSet);
        const val = document.querySelector(`[data-gun-manual-input="${idx}"]`).value.trim();
        const tsMs = TE.parseDateTimeToMs(state.config.raceDate, val);
        if (tsMs === null) return showToast('時間格式請輸入 HH:mm:ss（例如 06:30:00）', 'error');
        cats[idx].gunTimeMs = tsMs;
        cats[idx].gunTimeIso = new Date(tsMs).toISOString();
        const updated = await DB.saveConfig({ ...state.config, categories: cats });
        await syncConfigIfOnline(updated);
        await reloadState();
        showToast(`已設定「${cats[idx].name}」起跑時間為 ${TE.formatClockSec(tsMs)}`, 'success');
        return;
      }

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

    // 站點設定
    $('btn-add-checkpoint').addEventListener('click', async () => {
      const cps = [...(state.config.checkpoints || [])];
      const maxOrder = cps.reduce((m, c) => Math.max(m, Number(c.order || 0)), 0);
      cps.push({ id: `CP${maxOrder + 1}`, name: `補給站 ${maxOrder + 1}`, order: maxOrder + 1, distanceKm: 15, type: 'checkpoint' });
      await DB.saveConfig({ ...state.config, checkpoints: cps });
      await reloadState();
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

    // 後台分頁 4：Excel 匯出、JSON 備份還原、批次匯入、模擬資料
    $('btn-export-excel-full').addEventListener('click', () => {
      const f = Excel.exportResultsToExcel({ config: state.config, computed: state.computed, scans: state.scans, maskId: false });
      showToast(`已匯出完整 Excel 報表：${f}`, 'success');
    });
    $('btn-export-excel-masked').addEventListener('click', () => {
      const f = Excel.exportResultsToExcel({ config: state.config, computed: state.computed, scans: state.scans, maskId: true });
      showToast(`已匯出隱碼公開 Excel 報表：${f}`, 'success');
    });

    $('btn-backup-download').addEventListener('click', () => downloadJsonBackup('智慧路跑完整備份'));
    $('input-backup-merge').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      await handleImportBackupFile(file, 'merge');
      e.target.value = '';
    });
    $('input-backup-replace').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (file && confirm('確定要以備份檔完全覆蓋還原本機所有賽事設定與紀錄嗎？')) {
        await handleImportBackupFile(file, 'replace');
      }
      e.target.value = '';
    });

    $('btn-import-csv-text').addEventListener('click', async () => {
      const text = $('textarea-bulk-csv').value.trim();
      if (!text) return showToast('請先貼上 CSV 名單內容', 'warn');
      const rows = text.split(/\r?\n/).map((line) => line.split(/[,\t]/));
      await importRunnersFromRows(rows);
      $('textarea-bulk-csv').value = '';
    });
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

    $('btn-load-demo-data').addEventListener('click', async () => {
      if (!confirm('確定載入模擬賽事資料（將覆蓋目前本機測試選手與感應紀錄）？')) return;
      await DB.loadDemoData();
      await reloadState();
      showToast('已載入完整模擬賽事資料！', 'success');
    });
    $('btn-clear-all-data').addEventListener('click', async () => {
      if (!confirm('確定清空本機所有選手與感應紀錄嗎？（不影響已上傳 Google 試算表的資料）')) return;
      await DB.clearAllData(true);
      await reloadState();
      showToast('本機資料已清空重設', 'info');
    });

    // 後台分頁 5：雲端同步與分享連結
    document.querySelectorAll('input[name="sync-mode"]').forEach((radio) => {
      radio.addEventListener('change', (e) => {
        DB.saveDeviceSettings({ mode: e.target.value });
        updateNetworkIndicators();
        showToast(e.target.value === 'online' ? '已切換為「多機即時連線模式」' : '已切換為「單機離線模式」', 'info');
      });
    });

    $('btn-save-gas-url').addEventListener('click', async () => {
      const url = $('input-gas-url').value.trim();
      DB.saveDeviceSettings({ gasUrl: url });
      renderHeader();
      updateNetworkIndicators();
      if (!url) return showToast('已清除 GAS 網址', 'info');
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
        await Sync.pullFromCloud(state.session.auth);
        await reloadState();
        showToast('已從 Google 試算表拉取最新賽事、帳號與感應資料！', 'success');
      } catch (err) {
        showToast(`雲端拉取失敗：${err.message}`, 'error');
      }
    });

    $('btn-sync-push-all-config').addEventListener('click', async () => {
      if (!Sync.hasUrl()) return showToast('請先設定 GAS 網址', 'warn');
      try {
        await Sync.pushConfig(state.config, state.session.auth);
        showToast('已將本機全部賽事設定、站點與作業員帳號強制推送至雲端！', 'success');
      } catch (err) {
        showToast(`推送失敗：${err.message}`, 'error');
      }
    });

    const copyLink = async (path, label) => {
      const gasUrl = DB.getDeviceSettings().gasUrl;
      const target = new URL(path, location.href);
      if (gasUrl) target.searchParams.set('api', gasUrl);
      try {
        await navigator.clipboard.writeText(target.toString());
        showToast(`已複製「${label}」！`, 'success');
      } catch {
        prompt(`請複製以下${label}：`, target.toString());
      }
    };

    $('btn-copy-staff-link').addEventListener('click', () => copyLink('./index.html', '工作人員端連線網址'));
    $('btn-copy-public-link').addEventListener('click', () => copyLink('./public.html', '公開查詢頁分享連結'));
  }

  document.addEventListener('DOMContentLoaded', initApp);
})();
