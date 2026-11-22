/**
 * 智慧路跑 NFC 晶片計時平台 v3.0 — 多場次共用晶片主控制器
 * 核心運作流程：
 * 1. 平台可建立多項路跑活動，各有獨立分組、鳴槍時間、站點與事前匯入選手明細
 * 2. 共用 NFC 晶片僅寫入「姓名|身份證字號」，同一選手可持同一晶片參加多場不同活動與組別
 * 3. 作業員選取活動登入後鎖定該活動與指定站點；感應晶片時自動以身份證（或 UID）比對該活動選手明細
 * 4. 支援無網路離線登入與感應，連網後一鍵統一上傳雲端同步
 */
(function () {
  'use strict';

  const DB = window.RaceDB;
  const TE = window.TimingEngine;
  const NFC = window.NFCManager;
  const Sync = window.CloudSync;
  const Excel = window.ExcelExporter;

  const state = {
    view: 'portal',
    session: null,      // { mode: 'staff'|'admin', account, auth, lockedEventId }
    selectedEventId: '',
    staffSubTab: 'scan',
    adminTab: 'events',
    platform: null,
    chips: [],
    allRegistrations: [],
    allScans: [],
    computed: null,
    rankBy: 'gun',
    resultsCatFilter: 'ALL',
    resultsSearch: '',
    rosterSearch: '',
    writerEventFilter: 'ALL',
    adminRegSearch: '',
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
    }, 3600);
  }

  function getActiveEvent() {
    const events = state.platform?.events || [];
    const targetId =
      state.view === 'staff-workspace' && state.session?.lockedEventId
        ? state.session.lockedEventId
        : state.selectedEventId;
    return events.find((e) => e.eventId === targetId) || events[0] || null;
  }

  /* ========================================================================
   * 初始化與資料重新載入
   * ======================================================================== */
  async function initApp() {
    registerServiceWorker();
    await DB.init();

    const params = new URLSearchParams(location.search);
    const paramApi = (params.get('api') || '').trim();
    const paramEvent = (params.get('event') || '').trim();
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

    state.selectedEventId = paramEvent || settings.selectedEventId || '';
    await reloadState();
    switchView(state.view);

    window.addEventListener('online', async () => {
      updateNetworkIndicators();
      const pending = await DB.getPendingSyncCounts();
      if (Sync.isOnlineEnabled() && pending.total > 0) {
        showToast(`網路已恢復！自動上傳 ${pending.total} 筆離線紀錄中...`, 'info');
        await handleUploadPendingSync(false);
      } else if (pending.total > 0 && Sync.hasUrl()) {
        showToast(`網路已恢復！您有 ${pending.total} 筆離線紀錄可點擊「統一上傳同步」`, 'info');
      }
    });

    window.addEventListener('offline', () => {
      updateNetworkIndicators();
      showToast('已進入無網路狀態！晶片計時作業仍可離線執行，有網後再統一上傳', 'warn');
    });
  }

  function registerServiceWorker() {
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
      navigator.serviceWorker.register('./sw.js').catch((err) => console.warn('SW 註冊失敗:', err));
    }
  }

  async function reloadState() {
    const [platform, chips, regs, scans] = await Promise.all([
      DB.getPlatform(),
      DB.getAllChips(),
      DB.getAllRegistrations(),
      DB.getAllScans()
    ]);

    state.platform = platform;
    state.chips = chips;
    state.allRegistrations = regs;
    state.allScans = scans;

    const events = platform.events || [];
    if (!state.selectedEventId || !events.some((e) => e.eventId === state.selectedEventId)) {
      state.selectedEventId = platform.activeEventId || events[0]?.eventId || '';
      DB.saveDeviceSettings({ selectedEventId: state.selectedEventId });
    }

    const evt = getActiveEvent();
    if (evt) {
      const evtRegs = regs.filter((r) => r.eventId === evt.eventId);
      const evtScans = scans.filter((s) => s.eventId === evt.eventId);
      state.computed = TE.computeResults({
        runners: evtRegs,
        scans: evtScans,
        categories: evt.categories || [],
        checkpoints: evt.checkpoints || [],
        rankBy: state.rankBy
      });
    } else {
      state.computed = null;
    }

    renderHeader();
    renderEventDropdowns();
    updateNetworkIndicators();
    if (state.view === 'staff-workspace') renderStaffWorkspace();
    if (state.view === 'admin-workspace') renderAdminWorkspace();
    await updatePendingSyncBadge();
  }

  function switchView(viewName) {
    if (viewName !== 'staff-workspace') {
      NFC.cancelActiveOperation();
      if (NFC.isScanning) {
        NFC.stopScanning();
        resetScanButtonUi();
      }
    }
    state.view = viewName;
    document.querySelectorAll('.app-view').forEach((v) => v.classList.add('hidden'));
    const target = $(`view-${viewName}`);
    if (target) target.classList.remove('hidden');

    renderHeader();
    renderEventDropdowns();
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
        ? 'admin-tab-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-black bg-amber-500 text-slate-950'
        : 'admin-tab-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800';
    });
  }

  function startLiveClock() {
    const el = $('live-clock');
    function tick() {
      if (el) el.textContent = TE.formatClockMs(Date.now());
      requestAnimationFrame(tick);
    }
    requestAnimationFrame(tick);
  }

  async function renderNfcDiagnostics() {
    const diag = NFC.getDiagnostics();
    const perm = diag.supported ? await NFC.queryPermissionState() : 'unsupported';
    const scanHint = $('scan-nfc-env-hint');
    const writerBanner = $('writer-nfc-banner');
    const writerText = $('writer-nfc-status-text');

    let finalMsg = diag.statusMessage;
    if (diag.supported && perm === 'denied') {
      finalMsg = '❌ Chrome 的 NFC 權限已被拒絕：請點網址列左側圖示 → 權限 → 允許「NFC 裝置」後重新整理';
    }
    if (scanHint) {
      scanHint.textContent = diag.supported
        ? (perm === 'denied' ? '❌ NFC 權限遭拒' : '✅ Web NFC 就緒（自動讀取晶片姓名/身份證或 UID 比對本場明細）')
        : '⚠️ 非 Web NFC 環境（可使用模擬晶片或手動補登）';
    }
    if (writerText) writerText.textContent = finalMsg;
    if (writerBanner && (!diag.supported || perm === 'denied')) {
      writerBanner.className = 'rounded-2xl border border-amber-500/40 bg-amber-950/30 p-4 text-sm text-amber-200 flex items-start gap-3';
    }
  }

  /* ========================================================================
   * 頂部狀態、活動選單與網路燈號
   * ======================================================================== */
  function renderHeader() {
    const settings = DB.getDeviceSettings();
    const evt = getActiveEvent();
    $('header-event-badge').textContent = evt ? `${evt.name} (${evt.date || ''})` : '尚無活動';
    $('header-device-id').textContent = settings.deviceId;

    const pubLink = $('portal-link-public');
    if (pubLink) {
      const target = new URL('./public.html', location.href);
      if (evt) target.searchParams.set('event', evt.eventId);
      if (settings.gasUrl) target.searchParams.set('api', settings.gasUrl);
      pubLink.href = target.toString();
    }

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

  function renderEventDropdowns() {
    const events = state.platform?.events || [];
    const statusLabel = { active: '進行中', upcoming: '籌備中', finished: '已結束' };
    const optionsHtml = events.length
      ? events.map((ev) => `<option value="${esc(ev.eventId)}">${esc(ev.name)} (${esc(ev.date || '')}) [${statusLabel[ev.status] || ev.status}]</option>`).join('')
      : '<option value="">尚無活動</option>';

    ['portal-event-select', 'staff-login-event', 'admin-event-select'].forEach((id) => {
      const el = $(id);
      if (!el) return;
      el.innerHTML = optionsHtml;
      if (events.some((e) => e.eventId === state.selectedEventId)) {
        el.value = state.selectedEventId;
      }
    });

    const wf = $('writer-event-filter');
    if (wf) {
      const prev = state.writerEventFilter;
      wf.innerHTML = '<option value="ALL">所有活動名單</option>' + optionsHtml;
      wf.value = (prev === 'ALL' || events.some((e) => e.eventId === prev)) ? prev : 'ALL';
    }
  }

  function updateNetworkIndicators() {
    const settings = DB.getDeviceSettings();
    const onlineNet = navigator.onLine;
    const badge = $('badge-network');

    if (settings.mode === 'online' && settings.gasUrl) {
      badge.innerHTML = onlineNet
        ? '<span class="w-2 h-2 rounded-full bg-sky-400 animate-pulse"></span><span>多機雲端連線</span>'
        : '<span class="w-2 h-2 rounded-full bg-amber-400"></span><span>無網路・離線暫存</span>';
      badge.className = onlineNet
        ? 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-sky-950 border border-sky-500/50 text-sky-300'
        : 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-amber-950 border border-amber-500/50 text-amber-300';
    } else {
      badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-emerald-400"></span><span>單機離線模式</span>';
      badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-slate-800 border border-slate-700 text-slate-300';
    }

    const portalBanner = $('portal-net-banner');
    if (portalBanner) {
      if (onlineNet) {
        portalBanner.className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-xs font-extrabold bg-emerald-950/70 border border-emerald-500/40 text-emerald-300';
        portalBanner.textContent = settings.gasUrl
          ? '🟢 網路連線正常 • 已連接 Google 雲端試算表，亦支援無網路離線作業'
          : '🟢 網路連線正常 • 本機離線模式運作中（可至後台設定 Google 試算表同步）';
      } else {
        portalBanner.className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-xs font-extrabold bg-amber-950/80 border border-amber-500/50 text-amber-300';
        portalBanner.textContent = '📴 目前無網路連線 • 晶片計時作業仍可正常離線登入與感應，連網後統一上傳同步！';
      }
    }

    const loginHint = $('staff-login-offline-hint');
    if (loginHint) {
      if (!onlineNet) {
        loginHint.className = 'rounded-xl border border-amber-500/40 bg-amber-950/40 p-3 text-xs text-amber-200 font-bold';
        loginHint.textContent = '📴 無網路離線模式：使用本機快取帳號與各活動名單進行比對感應，恢復網路後一鍵統一上傳！';
      } else if (settings.gasUrl) {
        loginHint.className = 'rounded-xl border border-sky-500/40 bg-sky-950/40 p-3 text-xs text-sky-200';
        loginHint.textContent = '🌐 雲端驗證就緒：登入時會自動同步最新活動名單與晶片主檔至本機，確保賽道斷訊也能離線比對。';
      } else {
        loginHint.className = 'rounded-xl border border-emerald-500/30 bg-emerald-950/30 p-3 text-xs text-emerald-200';
        loginHint.textContent = '💾 本機單機模式：驗證本機作業員帳號並鎖定活動與感應站點。';
      }
    }

    const toggleBtn = $('btn-staff-toggle-mode');
    if (toggleBtn) {
      if (settings.mode === 'online' && settings.gasUrl) {
        toggleBtn.textContent = onlineNet ? '🌐 目前：雲端即時連線（點此切離線）' : '⚠️ 連線模式（等待訊號中）';
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

    $('btn-quick-sync').classList.toggle('hidden', !(pending.total > 0 && settings.gasUrl));

    const syncInfo = $('sync-status-info');
    if (syncInfo) {
      const lastStr = settings.lastSyncAt ? TE.formatClockSec(settings.lastSyncAt) : '尚未同步';
      syncInfo.textContent = `待同步：共用晶片 ${pending.chips} 筆、活動報名 ${pending.registrations} 筆、感應紀錄 ${pending.scans} 筆 ｜ 最後同步：${lastStr}`;
    }
  }

  /* ========================================================================
   * 登入與登出流程（含指定活動鎖定）
   * ======================================================================== */
  async function handleLogin(mode, username, password, chosenEventId = '') {
    const u = String(username || '').trim();
    const p = String(password || '');
    if (!u || !p) return showToast('請輸入帳號與密碼', 'warn');

    let account = null;
    let usedCloud = false;

    if (navigator.onLine && Sync.hasUrl()) {
      try {
        const res = await Sync.loginCloud(mode, u, p);
        account = { ...res.account, password: p };
        usedCloud = true;
      } catch (err) {
        if (String(err.message).includes('錯誤') || String(err.message).includes('停用')) {
          NFC.sound.beepError();
          showToast(err.message, 'error');
          return;
        }
      }
    }

    if (!account) {
      const localRes = await DB.verifyLocalLogin(mode, u, p);
      if (!localRes.ok) {
        NFC.sound.beepError();
        showToast(localRes.error, 'error');
        return;
      }
      account = { ...localRes.account, password: p };
    }

    // 決定作業員鎖定的活動：若帳號本身有綁定特定 eventId（非 ALL）則強制鎖定該活動；否則鎖定登入時選取的活動
    const platform = await DB.getPlatform();
    const events = platform.events || [];
    let lockedEventId = '';
    if (mode === 'staff') {
      if (account.eventId && account.eventId !== 'ALL') {
        lockedEventId = account.eventId;
      } else {
        lockedEventId = chosenEventId || state.selectedEventId || events[0]?.eventId || '';
      }
      state.selectedEventId = lockedEventId;
      DB.saveDeviceSettings({ selectedEventId: lockedEventId });
    }

    state.session = {
      mode,
      account,
      auth: { username: u, password: p },
      lockedEventId,
      loginAt: Date.now()
    };
    DB.saveSession(state.session);
    await reloadState();

    if (mode === 'admin') {
      showToast(usedCloud ? '🛡️ 已驗證最高管理員並同步最新全平台資料' : '🛡️ 已登入最高管理員後台', 'success');
      switchView('admin-workspace');
      switchAdminTab('events');
    } else {
      state.staffSubTab = account.role === 'writer' ? 'writer' : 'scan';
      const evt = getActiveEvent();
      const cp = (evt?.checkpoints || []).find((c) => c.id === account.checkpointId);
      const zoneName = account.role === 'writer'
        ? '共用 NFC 晶片寫入區'
        : `${evt ? evt.name : lockedEventId} ／ ${cp ? cp.name : account.checkpointId}`;
      showToast(`✅ 登入成功！已鎖定進入「${zoneName}」`, 'success');
      switchView('staff-workspace');
    }
  }

  function handleLogout() {
    NFC.cancelActiveOperation();
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
   * VIEW 4：晶片計時作業區（鎖定活動＋站點感應 / 共用晶片寫入）
   * ======================================================================== */
  function getLockedCheckpoint() {
    const evt = getActiveEvent();
    const acc = state.session?.account;
    if (!acc || !acc.checkpointId) return null;
    return (evt?.checkpoints || []).find((c) => c.id === acc.checkpointId) || {
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
    const evt = getActiveEvent();
    const lockedCp = getLockedCheckpoint();

    $('staff-locked-event-pill').textContent = role === 'writer'
      ? '🌐 共用晶片跨活動製卡'
      : `🏟️ 鎖定活動：${evt ? evt.name : '未選定'}`;

    let zoneTitle = '';
    if (role === 'writer') zoneTitle = '共用 NFC 晶片寫入與製卡區（僅姓名＋身份證）';
    else if (role === 'both') zoneTitle = `晶片寫入 ＋ 鎖定站點：${lockedCp ? lockedCp.name : '未指派'}`;
    else zoneTitle = `鎖定感應站點：${lockedCp ? lockedCp.name : '未指派（請洽管理員）'}`;

    $('staff-locked-zone-title').textContent = zoneTitle;
    $('staff-operator-meta').textContent = `操作員：${acc.displayName || acc.username} (${acc.username}) ｜ 權限與站點已鎖定，防止選錯站`;

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
      renderScanPanel(evt, lockedCp);
    } else {
      bothTabs.classList.remove('hidden');
      document.querySelectorAll('.staff-sub-btn').forEach((b) => {
        const active = b.dataset.staffSub === state.staffSubTab;
        b.className = active
          ? 'staff-sub-btn flex-1 py-3 rounded-xl font-black text-sm bg-emerald-500 text-slate-950'
          : 'staff-sub-btn flex-1 py-3 rounded-xl font-black text-sm bg-slate-800 text-slate-300';
      });
      panelScan.classList.toggle('hidden', state.staffSubTab !== 'scan');
      panelWriter.classList.toggle('hidden', state.staffSubTab !== 'writer');
      if (state.staffSubTab === 'scan') renderScanPanel(evt, lockedCp);
      else renderWriterPanel();
    }
  }

  function renderScanPanel(evt, lockedCp) {
    $('locked-event-title').textContent = evt ? `🏟️ 鎖定活動：${evt.name}（${evt.date || ''}）` : '尚未選定活動';
    $('locked-cp-name').textContent = lockedCp ? `#${lockedCp.order || '-'} ${lockedCp.name}` : '未指派站點';
    $('locked-cp-meta').textContent = lockedCp ? `代碼：${lockedCp.id} ｜ 累計距離：${lockedCp.distanceKm} km` : '請確認本活動有對應站點代碼';

    // 模擬器列出所有「共用晶片」（並標示該晶片是否已報名本活動，方便測試比對與未報名攔截！）
    const evtRegIdSet = new Set(
      state.allRegistrations
        .filter((r) => evt && r.eventId === evt.eventId)
        .map((r) => String(r.idNumber || '').toUpperCase())
    );
    const mockSel = $('mock-chip-select');
    const prevVal = mockSel.value;

    // 合併共用晶片主檔 + 本場已報名人員（即使還沒寫卡也能模擬刷過測試）
    const chipPoolMap = new Map();
    state.chips.forEach((c) => chipPoolMap.set(c.idNumber, { name: c.name, idNumber: c.idNumber, nfcUid: c.nfcUid }));
    state.allRegistrations.forEach((r) => {
      if (r.idNumber && !chipPoolMap.has(r.idNumber)) {
        chipPoolMap.set(r.idNumber, { name: r.name, idNumber: r.idNumber, nfcUid: '' });
      }
    });

    const pool = Array.from(chipPoolMap.values());
    mockSel.innerHTML =
      '<option value="">-- 隨機刷過一張晶片 --</option>' +
      pool
        .map((c) => {
          const inEvt = evtRegIdSet.has(c.idNumber);
          return `<option value="${esc(c.idNumber)}">${esc(c.name)} (${esc(c.idNumber)}) ${inEvt ? '✅已報本場' : '⚠️未報本場'}</option>`;
        })
        .join('');
    if (prevVal && chipPoolMap.has(prevVal)) mockSel.value = prevVal;

    renderRecentCheckpointScans(evt?.eventId || '', lockedCp?.id || '');
  }

  function renderRecentCheckpointScans(eventId, cpId) {
    const evt = getActiveEvent();
    const regMap = new Map(
      state.allRegistrations.filter((r) => r.eventId === eventId).map((r) => [r.bib, r])
    );
    const catMap = new Map((evt?.categories || []).map((c) => [c.id, c]));

    const list = state.allScans
      .filter((s) => !s.deleted && s.eventId === eventId && s.checkpointId === cpId)
      .sort((a, b) => b.timestampMs - a.timestampMs);

    $('checkpoint-scan-count').textContent = `${list.length} 筆`;
    const tbody = $('recent-scans-tbody');
    if (list.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="py-6 text-center text-slate-500">本活動鎖定站點尚無感應紀錄</td></tr>';
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
        const reg = regMap.get(s.bib);
        const cat = reg ? catMap.get(reg.categoryId) : null;
        return `
          <tr class="hover:bg-slate-800/40">
            <td class="py-2.5 pr-3 font-mono-num font-bold text-sky-300">${esc(TE.formatClockMs(s.timestampMs))}</td>
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-300">${esc(s.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(reg ? reg.name : '未登錄')}</td>
            <td class="py-2.5 pr-3 text-slate-300">${esc(cat ? cat.name : (reg?.categoryId || '--'))}</td>
            <td class="py-2.5 pr-3">${sourceBadges[s.source] || esc(s.source)}</td>
            <td class="py-2.5 text-right text-xs font-bold ${s.synced ? 'text-emerald-400' : 'text-amber-400'}">
              ${s.synced ? '☁️ 已同步' : '💾 離線暫存'}
            </td>
          </tr>
        `;
      })
      .join('');
  }

  /**
   * 核心：感應晶片後，依晶片內「姓名＋身份證字號（或 UID）」比對【本活動選手明細】！
   */
  async function processIncomingScan(tagData, { bypassDebounce = false } = {}) {
    const evt = getActiveEvent();
    const lockedCp = getLockedCheckpoint();
    if (!evt || !lockedCp || !lockedCp.id) {
      NFC.sound.beepError();
      showToast('尚未鎖定活動或站點，無法執行感應', 'error');
      return;
    }

    const scanTs = Date.now();
    const operator = state.session?.account?.username || '';
    const heroCard = $('scan-hero-card');
    heroCard.classList.remove('flash-success', 'flash-warn', 'flash-error');
    void heroCard.offsetWidth;

    // 1. 依晶片資料比對「目前鎖定活動」的事前匯入選手明細
    const matchRes = await DB.matchRegistrationForEvent(evt.eventId, {
      idNumber: tagData.idNumber || '',
      name: tagData.name || '',
      nfcUid: tagData.nfcUid || '',
      bib: tagData.bib || ''
    });

    if (!matchRes.matched || !matchRes.registration) {
      NFC.sound.beepError();
      heroCard.classList.add('flash-error');
      const idMask = matchRes.resolvedId ? TE.maskIdNumber(matchRes.resolvedId) : '';
      const who = matchRes.resolvedName || idMask || (tagData.nfcUid ? `UID:${tagData.nfcUid}` : '未知晶片');

      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-rose-500 text-white mb-3';
      $('scan-status-pill').textContent = `🚫 非本場活動報名選手（已攔截，不計入成績）`;
      $('hero-bib-cat').textContent = '無本場資格';
      $('hero-name').textContent = matchRes.resolvedName || '未登錄晶片';
      $('hero-elapsed').textContent = '--:--:--.---';
      $('hero-subinfo').textContent = `晶片識別：${who}${idMask ? ` (${idMask})` : ''} ｜ 未出現在「${evt.name}」事前匯入名單中`;
      showToast(`攔截：${who} 未報名「${evt.name}」`, 'warn');
      return;
    }

    const reg = matchRes.registration;
    const cat = (evt.categories || []).find((c) => c.id === reg.categoryId);
    const displayCat = cat ? cat.name : reg.categoryId || '未分組';

    // 2. 寫入本活動感應紀錄（含防連點檢查）
    const res = await DB.addScan({
      eventId: evt.eventId,
      bib: reg.bib,
      idNumber: reg.idNumber,
      checkpointId: lockedCp.id,
      timestampMs: scanTs,
      source: tagData.source || 'nfc',
      operator,
      note: tagData.note || (tagData.uidOnly ? `UID感應(${tagData.nfcUid})` : ''),
      bypassDebounce
    });

    if (!res.ok && res.duplicate) {
      NFC.sound.beepWarning();
      heroCard.classList.add('flash-warn');
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-amber-500 text-slate-950 mb-3';
      $('scan-status-pill').textContent = `⚠️ 防連點略過 — ${res.elapsedSec} 秒前已感應過（剩 ${res.waitRemainingSec} 秒）`;
      $('hero-bib-cat').textContent = `${reg.bib} • ${displayCat}`;
      $('hero-name').textContent = reg.name;
      $('hero-subinfo').textContent = `本站重複感應不重複計時 ｜ 上次紀錄：${TE.formatClockMs(res.previousScan.timestampMs)}`;
      return;
    }

    // 3. 計算本場經過時間
    const startCp = [...(evt.checkpoints || [])].sort((a, b) => a.order - b.order).find((c) => c.type === 'start');
    const startScan = startCp
      ? state.allScans.find((s) => !s.deleted && s.eventId === evt.eventId && s.bib === reg.bib && s.checkpointId === startCp.id)
      : null;

    let elapsedStr = '尚未鳴槍';
    if (cat && cat.gunTimeMs) elapsedStr = TE.formatDuration(scanTs - Number(cat.gunTimeMs), true);
    else if (startScan) elapsedStr = TE.formatDuration(scanTs - Number(startScan.timestampMs), true);

    NFC.sound.beepSuccess();
    heroCard.classList.add('flash-success');
    $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-emerald-500 text-slate-950 mb-3';
    $('scan-status-pill').textContent = `✅ 比對成功並計時 — ${lockedCp.name}`;
    $('hero-bib-cat').textContent = `${reg.bib} • ${displayCat}`;
    $('hero-name').textContent = reg.name;
    $('hero-elapsed').textContent = elapsedStr;
    $('hero-subinfo').textContent = `感應時間：${TE.formatClockMs(scanTs)} ｜ 活動：${evt.name} ｜ 操作員：${operator}`;

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
      $('scan-status-pill').textContent = '🟢 NFC 感應運作中 — 選手晶片靠近即自動比對本活動明細';
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-emerald-950 border border-emerald-500 text-emerald-300 mb-3';
      showToast('已啟動極速 NFC 感應模式', 'success');
    } catch (err) {
      showToast(err.message, 'warn');
    }
  }

  /* ---------------- 共用晶片寫入面板（僅姓名＋身份證字號） ---------------- */
  function renderWriterPanel() {
    const chipMap = new Map(state.chips.map((c) => [c.idNumber, c]));
    const peopleMap = new Map();

    // 彙整所有已匯入活動名單的選手（依身份證去重，計算報名幾場活動）
    for (const r of state.allRegistrations) {
      if (state.writerEventFilter !== 'ALL' && r.eventId !== state.writerEventFilter) continue;
      const id = String(r.idNumber || '').trim().toUpperCase();
      if (!id) continue;
      if (!peopleMap.has(id)) {
        peopleMap.set(id, { idNumber: id, name: r.name, eventIds: new Set() });
      }
      peopleMap.get(id).eventIds.add(r.eventId);
    }

    // 若篩選為 ALL，也補上已寫卡但尚未報名的晶片主檔
    if (state.writerEventFilter === 'ALL') {
      for (const c of state.chips) {
        if (!peopleMap.has(c.idNumber)) {
          peopleMap.set(c.idNumber, { idNumber: c.idNumber, name: c.name, eventIds: new Set() });
        }
      }
    }

    const q = state.rosterSearch.trim().toLowerCase();
    const list = Array.from(peopleMap.values())
      .filter((p) => !q || p.name.toLowerCase().includes(q) || p.idNumber.toLowerCase().includes(q))
      .sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));

    $('writer-people-count').textContent = String(list.length);
    const tbody = $('roster-tbody');
    if (list.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" class="py-6 text-center text-slate-500">尚無選手資料，請由後台匯入活動名單或於左側手動填寫製卡</td></tr>';
      return;
    }

    tbody.innerHTML = list
      .map((p) => {
        const chip = chipMap.get(p.idNumber);
        const statusHtml = chip
          ? `<span class="text-emerald-400 font-bold">✅ 已製卡${chip.nfcUid ? ` (${esc(chip.nfcUid.slice(0, 11))})` : ''}</span>`
          : '<span class="text-slate-500">⏳ 尚未製卡</span>';
        return `
          <tr class="hover:bg-slate-800/40">
            <td class="py-2 pr-2 font-bold text-white">${esc(p.name)}</td>
            <td class="py-2 pr-2 font-mono-num text-sky-300">${esc(p.idNumber)}</td>
            <td class="py-2 pr-2 font-mono-num text-slate-300">${p.eventIds.size} 場</td>
            <td class="py-2 pr-2 text-xs">${statusHtml}</td>
            <td class="py-2 text-right">
              <button type="button" data-fill-chip-name="${esc(p.name)}" data-fill-chip-id="${esc(p.idNumber)}"
                class="px-2.5 py-1 rounded-lg text-xs font-bold bg-sky-600/30 hover:bg-sky-600/50 text-sky-300">
                帶入寫卡
              </button>
            </td>
          </tr>
        `;
      })
      .join('');
  }

  function getWriterChipInput() {
    return {
      name: $('writer-name').value.trim(),
      idNumber: $('writer-idnumber').value.trim().toUpperCase()
    };
  }

  function setWriterVerifyBox(status, title, detail, showCancel = false) {
    const box = $('writer-verify-box');
    const styles = {
      working: 'rounded-xl border border-sky-500/50 bg-sky-950/40 p-4 text-sm text-sky-200 space-y-2',
      success: 'rounded-xl border border-emerald-500/50 bg-emerald-950/40 p-4 text-sm text-emerald-200 space-y-2',
      warn: 'rounded-xl border border-amber-500/50 bg-amber-950/40 p-4 text-sm text-amber-200 space-y-2',
      error: 'rounded-xl border border-rose-500/50 bg-rose-950/40 p-4 text-sm text-rose-200 space-y-2'
    };
    box.className = styles[status] || styles.working;
    $('writer-verify-title').textContent = title;
    $('writer-verify-detail').textContent = detail;
    $('btn-cancel-nfc-op').classList.toggle('hidden', !showCancel);
  }

  async function executeNfcWrite(isMock = false) {
    const person = getWriterChipInput();
    if (!person.name || !person.idNumber) {
      return showToast('請填寫「選手姓名」與「身份證字號」', 'warn');
    }

    setWriterVerifyBox(
      'working',
      isMock ? '🧪 模擬寫入共用晶片...' : '📡 等待貼近 NFC 晶片...',
      `寫入內容（僅姓名｜身份證）：${NFC.encodeChipPayload(person)}`,
      !isMock
    );

    try {
      const fn = isMock ? NFC.mockWriteChipAndVerify.bind(NFC) : NFC.writeChipAndVerify.bind(NFC);
      let savedChip = null;

      const result = await fn(person, {
        onStep: (step) => setWriterVerifyBox('working', '⏳ 共用晶片寫入與自動驗證', step, !isMock),
        onWriteDone: async () => {
          savedChip = await DB.upsertChip({ ...person, nfcUid: 'NDEF-WRITTEN', writtenAt: Date.now() }, false);
          if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushChips([savedChip]).then(() => updatePendingSyncBadge());
          await reloadState();
        }
      });

      const finalChip = await DB.upsertChip(
        { ...person, nfcUid: result.nfcUid || savedChip?.nfcUid || 'NDEF-WRITTEN', writtenAt: Date.now() },
        false
      );

      if (result.verified) {
        setWriterVerifyBox(
          'success',
          `✅ 共用晶片寫入並自動驗證成功！（UID: ${result.nfcUid || '標準 NDEF'}）`,
          `晶片內容：${result.data.raw} ｜ 同一晶片可直接參加所有已報名活動！`,
          false
        );
        showToast(`${person.name} (${person.idNumber}) 共用晶片寫入驗證完成！`, 'success');
      } else {
        setWriterVerifyBox(
          'warn',
          '✅ 共用晶片已寫入成功（略過或未重刷驗證）',
          `寫入內容：${result.data.raw} ｜ 可隨時點「🔍 讀取/測試晶片」檢查`,
          false
        );
        showToast(`${person.name} 共用晶片已寫入成功`, 'success');
      }

      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushChips([finalChip]).then(() => updatePendingSyncBadge());
      await reloadState();
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ 共用晶片寫入失敗', err.message, false);
      showToast(err.message, 'error');
    }
  }

  async function executeBindNfcUid() {
    const person = getWriterChipInput();
    if (!person.name || !person.idNumber) {
      return showToast('請先填寫「選手姓名」與「身份證字號」再綁定晶片 UID', 'warn');
    }

    setWriterVerifyBox('working', '🔗 UID 綁定模式：請貼近 NFC 晶片...', '直接綁定晶片唯一序號 UID 至此身份證（跨活動通用）。', true);
    try {
      const tag = await NFC.readSingleTag((step) => setWriterVerifyBox('working', '🔗 讀取晶片 UID 中...', step, true));
      if (!tag.nfcUid) throw new Error('此晶片未回傳硬體 UID，請改用標準「寫入共用晶片」');

      const saved = await DB.upsertChip({ ...person, nfcUid: tag.nfcUid, writtenAt: Date.now() }, false);
      NFC.sound.beepWriteVerified();
      setWriterVerifyBox(
        'success',
        `✅ 晶片 UID 綁定成功！（UID: ${tag.nfcUid}）`,
        `${person.name} (${person.idNumber}) 持此晶片可於本平台所有已報名活動直接感應計時！`,
        false
      );
      showToast(`已綁定晶片 UID (${tag.nfcUid}) 給 ${person.name}`, 'success');
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushChips([saved]).then(() => updatePendingSyncBadge());
      await reloadState();
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ UID 綁定失敗', err.message, false);
      showToast(err.message, 'error');
    }
  }

  async function executeReadTestNfc() {
    setWriterVerifyBox('working', '🔍 讀取測試晶片：請貼近手機背面感應區...', '正在檢測晶片內容、UID 與所有已報名活動...', true);
    try {
      const tag = await NFC.readSingleTag((step) => setWriterVerifyBox('working', '🔍 讀取測試中...', step, true));
      const chipByUid = tag.nfcUid ? await DB.findChipByUid(tag.nfcUid) : null;
      const detectedId = tag.parsed?.idNumber || chipByUid?.idNumber || '';
      const detectedName = tag.parsed?.name || chipByUid?.name || '';

      const evtMap = new Map((state.platform?.events || []).map((e) => [e.eventId, e.name]));
      const joinedRegs = detectedId
        ? state.allRegistrations.filter((r) => String(r.idNumber || '').toUpperCase() === detectedId)
        : [];
      const joinedText = joinedRegs.length
        ? joinedRegs.map((r) => `${evtMap.get(r.eventId) || r.eventId} [${r.categoryId} / #${r.bib}]`).join('、')
        : '尚未匯入任何活動名單';

      const detail = [
        `UID：${tag.nfcUid || '無'}`,
        `原始內容：${tag.raw || '（空）'}`,
        `識別身分：${detectedName || '未知'} (${detectedId || '無身份證'})`,
        `已報名場次：${joinedText}`
      ].join(' ｜ ');

      setWriterVerifyBox('success', '✅ 晶片讀取成功！NFC 功能正常', detail, false);
      showToast('晶片讀取測試成功！', 'success');
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ 晶片讀取失敗', err.message, false);
      showToast(err.message, 'error');
    }
  }

  /* ========================================================================
   * VIEW 5：後台管理區（6 大分頁渲染）
   * ======================================================================== */
  function renderAdminWorkspace() {
    if (!state.session || state.session.mode !== 'admin') return;
    renderAdminEventsTab();
    renderAdminRosterTab();
    renderAdminResultsTab();
    renderAdminRaceTab();
    renderAdminStaffTab();
    renderAdminSyncTab();
  }

  function renderAdminEventsTab() {
    const events = state.platform.events || [];
    const currentEvt = getActiveEvent();
    const statusBadge = {
      active: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-500/20 text-emerald-300">進行中</span>',
      upcoming: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-sky-500/20 text-sky-300">籌備中</span>',
      finished: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-slate-700 text-slate-300">已結束</span>'
    };

    $('admin-events-list').innerHTML = events
      .map((ev) => {
        const regCount = state.allRegistrations.filter((r) => r.eventId === ev.eventId).length;
        const scanCount = state.allScans.filter((s) => !s.deleted && s.eventId === ev.eventId).length;
        const isSelected = currentEvt && ev.eventId === currentEvt.eventId;
        const isDefault = ev.eventId === state.platform.activeEventId;
        return `
          <div class="p-4 rounded-2xl border ${isSelected ? 'border-amber-400 bg-amber-950/20' : 'border-slate-800 bg-slate-950'} space-y-2.5">
            <div class="flex items-start justify-between gap-2">
              <div>
                <div class="flex items-center gap-2 flex-wrap">
                  <span class="font-black text-white text-base">${esc(ev.name)}</span>
                  ${statusBadge[ev.status] || ''}
                  ${isDefault ? '<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-sky-600 text-white">公開預設</span>' : ''}
                </div>
                <div class="text-xs text-slate-400 font-mono-num mt-1">
                  ${esc(ev.eventId)} ｜ ${esc(ev.date || '--')}${ev.location ? ` ｜ ${esc(ev.location)}` : ''}
                </div>
              </div>
            </div>
            <div class="text-xs text-slate-300 font-mono-num">
              報名人數：<strong>${regCount}</strong> 人 ｜ 分組：<strong>${(ev.categories || []).length}</strong> 組 ｜ 感應：<strong>${scanCount}</strong> 筆
            </div>
            <div class="flex flex-wrap gap-1.5 pt-1">
              <button type="button" data-select-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold ${isSelected ? 'bg-amber-500 text-slate-950' : 'bg-slate-800 hover:bg-slate-700 text-white'}">
                ${isSelected ? '✓ 管理中' : '切換管理'}
              </button>
              <button type="button" data-default-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-sky-300">
                設為公開預設
              </button>
              ${events.length > 1 ? `<button type="button" data-delete-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>` : ''}
            </div>
          </div>
        `;
      })
      .join('');

    if (currentEvt) {
      $('cfg-evt-name').value = currentEvt.name || '';
      $('cfg-evt-date').value = currentEvt.date || '';
      $('cfg-evt-location').value = currentEvt.location || '';
      $('cfg-evt-status').value = currentEvt.status || 'active';
      $('cfg-evt-debounce').value = String(currentEvt.debounceSeconds ?? 30);
    }

    $('cfg-admin-user').value = state.platform.adminUsername || 'admin';
    $('cfg-admin-pass').value = state.platform.adminPassword || 'admin1234';
  }

  function renderAdminRosterTab() {
    const evt = getActiveEvent();
    if (!evt) return;
    $('roster-import-event-name').textContent = evt.name;

    const chipMap = new Map(state.chips.map((c) => [c.idNumber, c]));
    const catMap = new Map((evt.categories || []).map((c) => [c.id, c.name]));
    const q = state.adminRegSearch.trim().toLowerCase();

    const evtRegs = state.allRegistrations
      .filter((r) => r.eventId === evt.eventId)
      .filter((r) => !q || r.bib.toLowerCase().includes(q) || r.name.toLowerCase().includes(q) || (r.idNumber || '').toLowerCase().includes(q))
      .sort((a, b) => a.bib.localeCompare(b.bib, 'zh-Hant', { numeric: true }));

    $('admin-reg-count').textContent = String(evtRegs.length);
    const tbody = $('admin-reg-tbody');
    if (evtRegs.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="py-6 text-center text-slate-500">本活動尚未匯入選手名單，請由上方貼上 CSV 或上傳 Excel 匯入</td></tr>';
      return;
    }

    tbody.innerHTML = evtRegs
      .map((r) => {
        const chip = r.idNumber ? chipMap.get(r.idNumber) : null;
        return `
          <tr class="hover:bg-slate-800/40">
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-300">${esc(r.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(r.name)}</td>
            <td class="py-2.5 pr-3 font-mono-num text-sky-300">${esc(r.idNumber || r.idMasked || '--')}</td>
            <td class="py-2.5 pr-3 text-slate-300">${esc(catMap.get(r.categoryId) || r.categoryId)}</td>
            <td class="py-2.5 pr-3 text-xs font-bold ${chip ? 'text-emerald-400' : 'text-slate-500'}">
              ${chip ? `✅ 共用晶片已建檔${chip.nfcUid ? ` (${esc(chip.nfcUid)})` : ''}` : '⏳ 尚未製卡'}
            </td>
            <td class="py-2.5 text-right space-x-1.5">
              <button type="button" data-edit-reg="${esc(r.regId)}" class="px-2.5 py-1 rounded-lg text-xs font-bold bg-sky-600/30 hover:bg-sky-600/50 text-sky-300">編輯</button>
              <button type="button" data-del-reg="${esc(r.regId)}" class="px-2.5 py-1 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">移除</button>
            </td>
          </tr>
        `;
      })
      .join('');
  }

  function renderAdminResultsTab() {
    const evt = getActiveEvent();
    const computed = state.computed;
    if (!evt || !computed) return;

    $('kpi-total').textContent = String(computed.summary.totalRunners);
    $('kpi-finished').textContent = String(computed.summary.finishedCount);
    $('kpi-oncourse').textContent = String(computed.summary.onCourseCount);
    $('kpi-notstarted').textContent = String(computed.summary.notStartedCount);

    const catFilter = $('results-category-filter');
    const cats = evt.categories || [];
    catFilter.innerHTML = '<option value="ALL">全部組別</option>' + cats.map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    if (!cats.some((c) => c.id === state.resultsCatFilter)) state.resultsCatFilter = 'ALL';
    catFilter.value = state.resultsCatFilter;
    $('results-rankby-select').value = state.rankBy;

    const checkpoints = computed.checkpoints;
    $('results-thead').innerHTML = `
      <tr>
        <th class="px-3 py-3">總名次</th>
        <th class="px-3 py-3">組名次</th>
        <th class="px-3 py-3">編號</th>
        <th class="px-3 py-3">姓名</th>
        <th class="px-3 py-3">身份證</th>
        <th class="px-3 py-3">本場組別</th>
        <th class="px-3 py-3">狀態</th>
        ${checkpoints.map((cp) => `<th class="px-3 py-3">${esc(cp.name)}</th>`).join('')}
        <th class="px-3 py-3">分段 / 配速</th>
        <th class="px-3 py-3">罰時</th>
        <th class="px-3 py-3">大會時間</th>
        <th class="px-3 py-3">淨時間</th>
        <th class="px-3 py-3">平均配速</th>
      </tr>
    `;

    const q = state.resultsSearch.trim().toLowerCase();
    const filtered = computed.rows.filter((r) => {
      if (state.resultsCatFilter !== 'ALL' && r.categoryId !== state.resultsCatFilter) return false;
      if (q) return r.bib.toLowerCase().includes(q) || r.name.toLowerCase().includes(q) || (r.idNumber || '').toLowerCase().includes(q);
      return true;
    });

    const tbody = $('results-tbody');
    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${12 + checkpoints.length}" class="py-8 text-center text-slate-500">本活動尚無成績資料</td></tr>`;
      return;
    }

    const statusMap = {
      finished: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-500/20 text-emerald-300">已完賽</span>',
      on_course: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-amber-500/20 text-amber-300">賽道上</span>',
      not_started: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-slate-800 text-slate-400">未出發</span>'
    };

    tbody.innerHTML = filtered
      .map((row) => {
        const regId = row.runner.regId || `${evt.eventId}__${row.idNumber || row.bib}`;
        const cpCells = checkpoints
          .map((cp) => {
            const s = row.cpScanMap[cp.id];
            const label = s ? TE.formatClockMs(s.timestampMs) : '+ 補登';
            const editedTag = s && s.source === 'admin' ? ' <span class="text-amber-400">*</span>' : '';
            return `
              <td class="px-3 py-2.5 font-mono-num">
                <button type="button"
                  data-edit-scan-reg="${esc(regId)}"
                  data-edit-scan-bib="${esc(row.bib)}"
                  data-edit-scan-idnum="${esc(row.idNumber || '')}"
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
            <td class="px-3 py-2.5 font-bold text-white">${esc(row.name)}</td>
            <td class="px-3 py-2.5 font-mono-num text-slate-400">${esc(row.idNumber || row.idMasked || '--')}</td>
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
              <button type="button" data-edit-penalty-reg="${esc(regId)}" data-edit-penalty-bib="${esc(row.bib)}"
                class="px-2 py-1 rounded border border-dashed border-amber-500/50 text-amber-300 font-mono-num hover:bg-amber-900/30">
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

  function renderAdminRaceTab() {
    const evt = getActiveEvent();
    if (!evt) return;

    $('categories-list').innerHTML = (evt.categories || [])
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
              <button type="button" data-gun-now="${idx}" class="px-4 py-2 rounded-xl text-xs font-black bg-emerald-500 hover:bg-emerald-400 text-slate-950 shadow">
                🔫 一鍵鳴槍起跑
              </button>
              <input type="text" placeholder="HH:mm:ss" value="${c.gunTimeMs ? TE.formatClockSec(c.gunTimeMs) : ''}" data-gun-manual-input="${idx}"
                class="w-28 font-mono-num text-xs bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-white" />
              <button type="button" data-gun-manual-set="${idx}" class="px-3 py-2 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-sky-300">手動套用</button>
              <button type="button" data-gun-clear="${idx}" class="px-2.5 py-2 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-400">清除</button>
            </div>
          </div>
        </div>
      `
      )
      .join('');

    const sortedCps = [...(evt.checkpoints || [])].sort((a, b) => a.order - b.order);
    $('checkpoints-list').innerHTML = sortedCps
      .map(
        (cp, idx) => `
        <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-center p-3 rounded-xl bg-slate-950 border border-slate-800" data-cp-row="${idx}">
          <div class="sm:col-span-1"><input type="number" value="${esc(cp.order)}" data-cp-field="order" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-sm font-mono-num text-white" /></div>
          <div class="sm:col-span-2"><input type="text" value="${esc(cp.id)}" data-cp-field="id" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-mono-num text-white" /></div>
          <div class="sm:col-span-4"><input type="text" value="${esc(cp.name)}" data-cp-field="name" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-bold text-white" /></div>
          <div class="sm:col-span-2"><input type="number" step="0.1" value="${esc(cp.distanceKm)}" data-cp-field="distanceKm" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm font-mono-num text-white" /></div>
          <div class="sm:col-span-2">
            <select data-cp-field="type" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-xs font-bold text-white">
              <option value="start" ${cp.type === 'start' ? 'selected' : ''}>起點 (Start)</option>
              <option value="checkpoint" ${cp.type === 'checkpoint' ? 'selected' : ''}>中途補給/折返</option>
              <option value="finish" ${cp.type === 'finish' ? 'selected' : ''}>終點 (Finish)</option>
            </select>
          </div>
          <div class="sm:col-span-1 text-right"><button type="button" data-del-cp="${idx}" class="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button></div>
        </div>
      `
      )
      .join('');
  }

  function renderAdminStaffTab() {
    const list = state.platform.staffAccounts || [];
    const events = state.platform.events || [];
    const evt = getActiveEvent();
    const cps = [...(evt?.checkpoints || [])].sort((a, b) => a.order - b.order);

    $('staff-accounts-list').innerHTML = list
      .map(
        (s, idx) => `
        <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-end p-3.5 rounded-xl bg-slate-950 border border-slate-800" data-staff-row="${idx}">
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">登入帳號</label>
            <input type="text" value="${esc(s.username)}" data-staff-field="username" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-mono-num font-bold text-white" />
          </div>
          <div class="sm:col-span-1">
            <label class="block text-[11px] text-slate-400 mb-1">密碼</label>
            <input type="text" value="${esc(s.password)}" data-staff-field="password" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-mono-num text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">人員名稱</label>
            <input type="text" value="${esc(s.displayName)}" data-staff-field="displayName" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-bold text-white" />
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">作業權限</label>
            <select data-staff-field="role" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-bold text-white">
              <option value="scanner" ${s.role === 'scanner' ? 'selected' : ''}>鎖定感應站點</option>
              <option value="writer" ${s.role === 'writer' ? 'selected' : ''}>僅共用晶片寫入</option>
              <option value="both" ${s.role === 'both' ? 'selected' : ''}>寫入＋鎖定站點</option>
            </select>
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">限定活動</label>
            <select data-staff-field="eventId" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-bold text-amber-300">
              <option value="ALL" ${(!s.eventId || s.eventId === 'ALL') ? 'selected' : ''}>所有活動（登入選取）</option>
              ${events.map((e) => `<option value="${esc(e.eventId)}" ${s.eventId === e.eventId ? 'selected' : ''}>僅限：${esc(e.name)}</option>`).join('')}
            </select>
          </div>
          <div class="sm:col-span-2">
            <label class="block text-[11px] text-slate-400 mb-1">鎖定站點代碼</label>
            <input list="cp-datalist" type="text" value="${esc(s.checkpointId)}" data-staff-field="checkpointId" placeholder="START/CP1/FINISH"
              class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-mono-num text-sky-300" />
          </div>
          <div class="sm:col-span-1 flex items-center justify-between gap-1 pb-1">
            <input type="checkbox" data-staff-field="enabled" ${s.enabled !== false ? 'checked' : ''} class="accent-emerald-500 w-4 h-4" title="啟用" />
            <button type="button" data-del-staff="${idx}" class="px-2 py-1 rounded text-xs font-bold bg-rose-600/20 text-rose-300">刪</button>
          </div>
        </div>
      `
      )
      .join('') +
      `<datalist id="cp-datalist">${cps.map((cp) => `<option value="${esc(cp.id)}">${esc(cp.name)}</option>`).join('')}</datalist>`;
  }

  function renderAdminSyncTab() {
    const settings = DB.getDeviceSettings();
    document.querySelectorAll('input[name="sync-mode"]').forEach((radio) => {
      radio.checked = radio.value === settings.mode;
    });
    $('input-gas-url').value = settings.gasUrl || '';
  }

  /* ========================================================================
   * 輔助：更新當前活動並同步
   * ======================================================================== */
  async function updateCurrentEvent(mutator) {
    const events = [...(state.platform.events || [])];
    const idx = events.findIndex((e) => e.eventId === state.selectedEventId);
    if (idx < 0) return;
    events[idx] = { ...events[idx], ...mutator(events[idx]), updatedAt: Date.now() };
    const updatedPlatform = await DB.savePlatform({ ...state.platform, events });
    if (Sync.isOnlineEnabled() && navigator.onLine && state.session?.mode === 'admin') {
      try {
        await Sync.pushPlatform(updatedPlatform, state.session.auth);
      } catch (err) {
        showToast(`本機已更新，雲端同步待重試：${err.message}`, 'warn');
      }
    }
    await reloadState();
  }

  async function handleUploadPendingSync(silent = false) {
    if (!navigator.onLine) return showToast('目前無網路連線，請連網後再統一上傳', 'warn');
    if (!Sync.hasUrl()) return showToast('尚未設定 Google Apps Script 雲端網址', 'warn');
    try {
      const res = await Sync.syncPendingToCloud();
      await reloadState();
      if (!silent || res.uploadedScans > 0 || res.uploadedChips > 0 || res.uploadedRegs > 0) {
        showToast(
          `☁️ 統一上傳完成！晶片 ${res.uploadedChips}、報名 ${res.uploadedRegs}、感應 ${res.uploadedScans} 筆` +
            (res.skippedDuplicate ? `（略過重複 ${res.skippedDuplicate} 筆）` : ''),
          'success'
        );
      } else {
        showToast('目前沒有待上傳的離線資料', 'info');
      }
    } catch (err) {
      showToast(`雲端上傳失敗：${err.message}`, 'error');
    }
  }

  async function downloadJsonBackup(prefix = '智慧路跑多場次完整備份') {
    const backup = await DB.exportFullBackup();
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${prefix}_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast(`已下載完整備份（${backup.platform.events.length} 場活動、${backup.chips.length} 張晶片、${backup.scans.length} 筆感應）`, 'success');
  }

  async function importRegistrationsRows(rows) {
    const evt = getActiveEvent();
    if (!evt) return showToast('請先建立或選取活動', 'warn');
    const defaultCat = evt.categories?.[0]?.id || '10K';
    const valid = [];

    for (const cols of rows) {
      if (!cols || cols.length < 3) continue;
      const bib = String(cols[0] ?? '').trim();
      const name = String(cols[1] ?? '').trim();
      const idNumber = String(cols[2] ?? '').trim().toUpperCase();
      if (!bib || !name || bib === '參賽編號' || bib.toLowerCase() === 'bib') continue;
      valid.push({
        eventId: evt.eventId,
        bib,
        name,
        idNumber,
        categoryId: String(cols[3] ?? defaultCat).trim() || defaultCat
      });
    }

    if (valid.length === 0) return showToast('未偵測到有效名單（格式需為：參賽編號,姓名,身份證字號,分組代碼）', 'warn');
    const saved = await DB.upsertRegistrationsBulk(valid, false);
    if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRegistrations(saved).then(() => updatePendingSyncBadge());
    await reloadState();
    showToast(`已成功匯入 ${valid.length} 位選手至「${evt.name}」！`, 'success');
  }

  /* ---------------- Modals ---------------- */
  function openEditScanModal(regId, bib, idNumber, cpId, scanId) {
    const evt = getActiveEvent();
    const reg = state.allRegistrations.find((r) => r.regId === regId);
    const cp = (evt?.checkpoints || []).find((c) => c.id === cpId);
    const scan = scanId ? state.allScans.find((s) => s.id === scanId) : null;

    state.editMode = 'scan';
    $('modal-edit-title').textContent = `人工修正 — ${bib} ${reg?.name || ''} @ ${cp?.name || cpId}`;
    $('edit-reg-id').value = regId;
    $('edit-bib').value = bib;
    $('edit-idnumber').value = idNumber;
    $('edit-cp-id').value = cpId;
    $('edit-scan-id').value = scanId || '';
    $('edit-time-str').value = scan ? TE.formatClockMs(scan.timestampMs) : TE.formatClockMs(Date.now());
    $('edit-note').value = scan?.note || '後台人工修正';

    $('edit-scan-fields').classList.remove('hidden');
    $('edit-penalty-fields').classList.add('hidden');
    $('btn-delete-scan').classList.toggle('hidden', !scan);
    $('modal-edit-scan').classList.remove('hidden');
  }

  function openEditPenaltyModal(regId, bib) {
    const reg = state.allRegistrations.find((r) => r.regId === regId);
    state.editMode = 'penalty';
    $('modal-edit-title').textContent = `判罰加時設定 — ${bib} ${reg?.name || ''}`;
    $('edit-reg-id').value = regId;
    $('edit-bib').value = bib;
    $('edit-penalty-sec').value = String(reg?.penaltySec || 0);

    $('edit-scan-fields').classList.add('hidden');
    $('edit-penalty-fields').classList.remove('hidden');
    $('btn-delete-scan').classList.add('hidden');
    $('modal-edit-scan').classList.remove('hidden');
  }

  function openAdminRegModal(regId = '') {
    const evt = getActiveEvent();
    if (!evt) return;
    const existing = regId ? state.allRegistrations.find((r) => r.regId === regId) : null;
    $('modal-reg-title').textContent = existing ? `編輯選手：${existing.name} (#${existing.bib})` : `新增「${evt.name}」選手`;
    $('admin-reg-id').value = existing?.regId || '';
    $('admin-reg-bib').value = existing?.bib || '';
    $('admin-reg-name').value = existing?.name || '';
    $('admin-reg-idnumber').value = existing?.idNumber || '';
    $('admin-reg-cat').innerHTML = (evt.categories || []).map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    if (existing?.categoryId) $('admin-reg-cat').value = existing.categoryId;
    $('btn-admin-delete-reg').classList.toggle('hidden', !existing);
    $('modal-edit-reg').classList.remove('hidden');
  }

  /* ========================================================================
   * 事件綁定
   * ======================================================================== */
  function bindEvents() {
    // 活動切換（大廳 & 後台）
    const onEventChange = async (eventId) => {
      state.selectedEventId = eventId;
      state.resultsCatFilter = 'ALL';
      DB.saveDeviceSettings({ selectedEventId: eventId });
      await reloadState();
    };
    $('portal-event-select').addEventListener('change', (e) => onEventChange(e.target.value));
    $('admin-event-select').addEventListener('change', (e) => onEventChange(e.target.value));

    // 大廳導覽
    $('btn-portal-staff').addEventListener('click', () => switchView('staff-login'));
    $('btn-portal-admin').addEventListener('click', () => switchView('admin-login'));
    $('btn-logout').addEventListener('click', handleLogout);
    $('header-brand-home').addEventListener('click', () => { if (!state.session) switchView('portal'); });
    document.querySelectorAll('[data-goto]').forEach((b) => b.addEventListener('click', () => switchView(b.dataset.goto)));
    $('btn-quick-sync').addEventListener('click', () => handleUploadPendingSync(false));

    // 登入表單
    document.querySelectorAll('[data-quick-staff]').forEach((btn) => {
      btn.addEventListener('click', () => {
        $('staff-login-user').value = btn.dataset.quickStaff;
        $('staff-login-pass').value = '1111';
      });
    });
    $('form-staff-login').addEventListener('submit', (e) => {
      e.preventDefault();
      handleLogin('staff', $('staff-login-user').value, $('staff-login-pass').value, $('staff-login-event').value);
    });
    $('form-admin-login').addEventListener('submit', (e) => {
      e.preventDefault();
      handleLogin('admin', $('admin-login-user').value, $('admin-login-pass').value);
    });

    // 作業員工作區控制
    $('btn-staff-toggle-mode').addEventListener('click', async () => {
      const s = DB.getDeviceSettings();
      if (s.mode === 'offline') {
        if (!s.gasUrl) return showToast('尚未設定雲端網址，請由後台設定或透過工作人員連線連結開啟', 'warn');
        DB.saveDeviceSettings({ mode: 'online' });
        updateNetworkIndicators();
        showToast('已切換為「雲端即時連線模式」', 'success');
        const pending = await DB.getPendingSyncCounts();
        if (pending.total > 0 && navigator.onLine) await handleUploadPendingSync(false);
      } else {
        DB.saveDeviceSettings({ mode: 'offline' });
        updateNetworkIndicators();
        showToast('已切換為「單機離線作業模式」', 'info');
      }
    });
    $('btn-staff-upload-sync').addEventListener('click', () => handleUploadPendingSync(false));
    $('btn-staff-export-offline-json').addEventListener('click', () => {
      const op = state.session?.account?.username || 'staff';
      downloadJsonBackup(`離線感應備份_${op}`);
    });
    document.querySelectorAll('.staff-sub-btn').forEach((b) => {
      b.addEventListener('click', () => { state.staffSubTab = b.dataset.staffSub; renderStaffWorkspace(); });
    });

    // 感應端音效、高對比、NFC、手動補登、模擬器
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
    $('btn-toggle-nfc-scan').addEventListener('click', toggleNfcScanning);

    $('form-manual-scan').addEventListener('submit', async (e) => {
      e.preventDefault();
      const input = $('manual-scan-input');
      const q = input.value.trim();
      const evt = getActiveEvent();
      if (!q || !evt) return;
      const found = await DB.findRegistrationByQuery(evt.eventId, q);
      if (found) {
        await processIncomingScan(
          { bib: found.bib, name: found.name, idNumber: found.idNumber, source: 'manual', note: '現場手動補登' },
          { bypassDebounce: true }
        );
      } else {
        await processIncomingScan({ idNumber: q.toUpperCase(), name: q, bib: q.toUpperCase(), source: 'manual' }, { bypassDebounce: true });
      }
      input.value = '';
    });

    $('btn-load-demo-shortcut').addEventListener('click', async () => {
      await DB.loadDemoData();
      if (state.session?.mode === 'staff') {
        state.session.lockedEventId = 'EVT-2026-SPRING';
        DB.saveSession(state.session);
      }
      await reloadState();
      showToast('已載入雙場次共用晶片模擬資料！可直接模擬刷卡體驗比對機制', 'success');
    });

    $('btn-mock-scan').addEventListener('click', async () => {
      if (state.allRegistrations.length === 0) {
        await DB.loadDemoData();
        if (state.session?.mode === 'staff') {
          state.session.lockedEventId = 'EVT-2026-SPRING';
          DB.saveSession(state.session);
        }
        await reloadState();
      }
      const chosenId = $('mock-chip-select').value;
      let chip = chosenId ? state.chips.find((c) => c.idNumber === chosenId) : null;
      if (!chip && chosenId) {
        const r = state.allRegistrations.find((x) => x.idNumber === chosenId);
        if (r) chip = { name: r.name, idNumber: r.idNumber, nfcUid: '' };
      }
      if (!chip) {
        const evt = getActiveEvent();
        const evtRegs = state.allRegistrations.filter((r) => evt && r.eventId === evt.eventId);
        const pick = evtRegs[Math.floor(Math.random() * evtRegs.length)] || state.chips[0];
        if (!pick) return showToast('請先匯入本活動選手名單或載入模擬資料', 'warn');
        chip = { name: pick.name, idNumber: pick.idNumber, nfcUid: pick.nfcUid || '' };
      }
      // 模擬晶片內只有「姓名＋身份證字號」（不含 bib 或組別），完整測試真實比對邏輯！
      await processIncomingScan({
        name: chip.name,
        idNumber: chip.idNumber,
        nfcUid: chip.nfcUid || '',
        source: 'mock',
        note: '模擬共用晶片刷過'
      });
    });

    // 共用晶片寫入與製卡按鈕
    $('form-nfc-write').addEventListener('submit', (e) => { e.preventDefault(); executeNfcWrite(false); });
    $('btn-bind-uid-nfc').addEventListener('click', executeBindNfcUid);
    $('btn-read-test-nfc').addEventListener('click', executeReadTestNfc);
    $('btn-cancel-nfc-op').addEventListener('click', () => NFC.cancelActiveOperation());
    $('btn-mock-write-nfc').addEventListener('click', () => executeNfcWrite(true));
    $('btn-save-chip-only').addEventListener('click', async () => {
      const p = getWriterChipInput();
      if (!p.name || !p.idNumber) return showToast('請填寫姓名與身份證字號', 'warn');
      const saved = await DB.upsertChip(p, false);
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushChips([saved]);
      await reloadState();
      showToast(`已登錄 ${p.name} (${p.idNumber}) 至共用晶片主檔`, 'success');
    });

    $('writer-event-filter').addEventListener('change', (e) => { state.writerEventFilter = e.target.value; renderWriterPanel(); });
    $('roster-search').addEventListener('input', (e) => { state.rosterSearch = e.target.value; renderWriterPanel(); });
    $('roster-tbody').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-fill-chip-id]');
      if (!btn) return;
      $('writer-name').value = btn.dataset.fillChipName;
      $('writer-idnumber').value = btn.dataset.fillChipId;
      showToast(`已帶入 ${btn.dataset.fillChipName} (${btn.dataset.fillChipId}) 準備寫入共用晶片`, 'info');
    });

    /* ==================== 後台管理區事件 ==================== */
    document.querySelectorAll('.admin-tab-btn').forEach((b) => b.addEventListener('click', () => switchAdminTab(b.dataset.adminTab)));

    // 建立新路跑活動 Modal
    const closeNewEvt = () => $('modal-new-event').classList.add('hidden');
    $('btn-open-new-event-modal').addEventListener('click', () => {
      $('new-evt-name').value = '';
      $('new-evt-date').value = new Date().toISOString().slice(0, 10);
      $('new-evt-location').value = '';
      $('modal-new-event').classList.remove('hidden');
    });
    $('btn-close-new-event').addEventListener('click', closeNewEvt);
    $('btn-cancel-new-event').addEventListener('click', closeNewEvt);

    $('form-new-event').addEventListener('submit', async (e) => {
      e.preventDefault();
      const newEvt = DB.createEventTemplate(
        $('new-evt-name').value.trim(),
        $('new-evt-date').value,
        $('new-evt-location').value.trim()
      );
      const updated = await DB.savePlatform({
        ...state.platform,
        events: [...(state.platform.events || []), newEvt]
      });
      state.selectedEventId = newEvt.eventId;
      DB.saveDeviceSettings({ selectedEventId: newEvt.eventId });
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushPlatform(updated, state.session.auth);
      closeNewEvt();
      await reloadState();
      showToast(`🏟️ 已建立新活動「${newEvt.name}」！請接著設定分組並匯入選手名單`, 'success');
    });

    // 活動卡片操作
    $('admin-events-list').addEventListener('click', async (e) => {
      const selBtn = e.target.closest('[data-select-event]');
      if (selBtn) { await onEventChange(selBtn.dataset.selectEvent); return; }

      const defBtn = e.target.closest('[data-default-event]');
      if (defBtn) {
        const updated = await DB.savePlatform({ ...state.platform, activeEventId: defBtn.dataset.defaultEvent });
        if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushPlatform(updated, state.session.auth);
        await reloadState();
        showToast('已設為公開查詢預設顯示活動', 'success');
        return;
      }

      const delBtn = e.target.closest('[data-delete-event]');
      if (delBtn) {
        const eid = delBtn.dataset.deleteEvent;
        const target = (state.platform.events || []).find((x) => x.eventId === eid);
        if (!confirm(`確定要刪除活動「${target?.name || eid}」嗎？`)) return;
        const remain = (state.platform.events || []).filter((x) => x.eventId !== eid);
        const updated = await DB.savePlatform({
          ...state.platform,
          events: remain,
          activeEventId: state.platform.activeEventId === eid ? (remain[0]?.eventId || '') : state.platform.activeEventId
        });
        if (state.selectedEventId === eid) state.selectedEventId = remain[0]?.eventId || '';
        if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushPlatform(updated, state.session.auth);
        await reloadState();
        showToast('已刪除該場路跑活動', 'info');
      }
    });

    $('form-event-basic').addEventListener('submit', async (e) => {
      e.preventDefault();
      await updateCurrentEvent(() => ({
        name: $('cfg-evt-name').value.trim(),
        date: $('cfg-evt-date').value,
        location: $('cfg-evt-location').value.trim(),
        status: $('cfg-evt-status').value,
        debounceSeconds: Number($('cfg-evt-debounce').value || 30)
      }));
      showToast('本活動基本資訊已更新', 'success');
    });

    $('form-admin-credentials').addEventListener('submit', async (e) => {
      e.preventDefault();
      const newU = $('cfg-admin-user').value.trim() || 'admin';
      const newP = $('cfg-admin-pass').value || 'admin1234';
      const oldAuth = { ...state.session.auth };
      const updated = await DB.savePlatform({ ...state.platform, adminUsername: newU, adminPassword: newP });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushPlatform(updated, oldAuth); } catch (err) { showToast(`雲端同步失敗：${err.message}`, 'warn'); }
      }
      state.session.auth = { username: newU, password: newP };
      state.session.account.username = newU;
      DB.saveSession(state.session);
      await reloadState();
      showToast('最高管理員帳密已更新', 'success');
    });

    // 分頁 2：本活動選手名單匯入與編輯
    $('btn-import-csv-text').addEventListener('click', async () => {
      const text = $('textarea-bulk-csv').value.trim();
      if (!text) return showToast('請先貼上 CSV 名單', 'warn');
      await importRegistrationsRows(text.split(/\r?\n/).map((line) => line.split(/[,\t]/)));
      $('textarea-bulk-csv').value = '';
    });

    $('input-import-file').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const wb = window.XLSX.read(await file.arrayBuffer(), { type: 'array' });
        const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
        await importRegistrationsRows(rows);
      } catch (err) {
        showToast(`檔案解析失敗：${err.message}`, 'error');
      }
      e.target.value = '';
    });

    $('btn-admin-add-reg').addEventListener('click', () => openAdminRegModal(''));
    $('admin-reg-search').addEventListener('input', (e) => { state.adminRegSearch = e.target.value; renderAdminRosterTab(); });
    $('admin-reg-tbody').addEventListener('click', async (e) => {
      const editBtn = e.target.closest('[data-edit-reg]');
      if (editBtn) { openAdminRegModal(editBtn.dataset.editReg); return; }
      const delBtn = e.target.closest('[data-del-reg]');
      if (delBtn && confirm('確定移除此選手在本活動的報名資格嗎？')) {
        await DB.deleteRegistration(delBtn.dataset.delReg);
        await reloadState();
        showToast('已移除該筆活動報名', 'info');
      }
    });

    const closeRegModal = () => $('modal-edit-reg').classList.add('hidden');
    $('btn-close-reg-modal').addEventListener('click', closeRegModal);
    $('btn-cancel-reg-modal').addEventListener('click', closeRegModal);

    $('form-admin-reg').addEventListener('submit', async (e) => {
      e.preventDefault();
      const evt = getActiveEvent();
      if (!evt) return;
      const saved = await DB.upsertRegistration({
        regId: $('admin-reg-id').value || undefined,
        eventId: evt.eventId,
        bib: $('admin-reg-bib').value.trim(),
        name: $('admin-reg-name').value.trim(),
        idNumber: $('admin-reg-idnumber').value.trim().toUpperCase(),
        categoryId: $('admin-reg-cat').value
      }, false);
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRegistrations([saved]);
      closeRegModal();
      await reloadState();
      showToast(`已儲存 ${saved.name} (#${saved.bib}) 至本活動名單`, 'success');
    });

    $('btn-admin-delete-reg').addEventListener('click', async () => {
      const regId = $('admin-reg-id').value;
      if (!regId || !confirm('確定移除此筆活動報名？')) return;
      await DB.deleteRegistration(regId);
      closeRegModal();
      await reloadState();
    });

    // 分頁 3：成績管理與修正
    $('results-category-filter').addEventListener('change', (e) => { state.resultsCatFilter = e.target.value; renderAdminResultsTab(); });
    $('results-rankby-select').addEventListener('change', async (e) => { state.rankBy = e.target.value; await reloadState(); });
    $('results-search').addEventListener('input', (e) => { state.resultsSearch = e.target.value; renderAdminResultsTab(); });
    $('btn-export-excel-quick').addEventListener('click', () => {
      const evt = getActiveEvent();
      const f = Excel.exportResultsToExcel({
        event: evt,
        computed: state.computed,
        scans: state.allScans.filter((s) => s.eventId === evt.eventId),
        maskId: false
      });
      showToast(`已匯出「${evt.name}」Excel 報表：${f}`, 'success');
    });

    $('results-tbody').addEventListener('click', (e) => {
      const scanBtn = e.target.closest('[data-edit-scan-bib]');
      if (scanBtn) {
        openEditScanModal(scanBtn.dataset.editScanReg, scanBtn.dataset.editScanBib, scanBtn.dataset.editScanIdnum, scanBtn.dataset.editScanCp, scanBtn.dataset.editScanId);
        return;
      }
      const penBtn = e.target.closest('[data-edit-penalty-reg]');
      if (penBtn) openEditPenaltyModal(penBtn.dataset.editPenaltyReg, penBtn.dataset.editPenaltyBib);
    });

    const closeEditModal = () => $('modal-edit-scan').classList.add('hidden');
    $('btn-close-edit-modal').addEventListener('click', closeEditModal);
    $('btn-cancel-edit-modal').addEventListener('click', closeEditModal);

    $('form-edit-scan').addEventListener('submit', async (e) => {
      e.preventDefault();
      const evt = getActiveEvent();
      const regId = $('edit-reg-id').value;
      const bib = $('edit-bib').value;

      if (state.editMode === 'penalty') {
        const reg = state.allRegistrations.find((r) => r.regId === regId);
        if (reg) {
          const updated = await DB.upsertRegistration({ ...reg, penaltySec: Number($('edit-penalty-sec').value || 0) }, false);
          if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRegistrations([updated]);
        }
        closeEditModal();
        await reloadState();
        showToast('判罰加時已更新並重算名次', 'success');
        return;
      }

      const cpId = $('edit-cp-id').value;
      const scanId = $('edit-scan-id').value;
      const idNumber = $('edit-idnumber').value;
      const tsMs = TE.parseDateTimeToMs(evt.date, $('edit-time-str').value.trim());
      if (tsMs === null) return showToast('時間格式錯誤，請輸入 HH:mm:ss 或 HH:mm:ss.SSS', 'error');

      let savedScan;
      if (scanId) {
        const existing = state.allScans.find((s) => s.id === scanId);
        savedScan = await DB.putScan({
          ...existing,
          timestampMs: tsMs,
          source: 'admin',
          operator: state.session.account.username,
          note: $('edit-note').value.trim() || '後台人工修正',
          deleted: false,
          synced: 0
        });
      } else {
        const res = await DB.addScan({
          eventId: evt.eventId,
          bib,
          idNumber,
          checkpointId: cpId,
          timestampMs: tsMs,
          source: 'admin',
          operator: state.session.account.username,
          note: $('edit-note').value.trim() || '後台人工補登',
          bypassDebounce: true
        });
        savedScan = res.scan;
      }

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushAdminEditScan(savedScan, state.session.auth); } catch {}
      }
      closeEditModal();
      await reloadState();
      showToast('成績時間已修正並自動重算名次！', 'success');
    });

    $('btn-delete-scan').addEventListener('click', async () => {
      const scanId = $('edit-scan-id').value;
      if (!scanId || !confirm('確定刪除此站感應紀錄？')) return;
      const existing = state.allScans.find((s) => s.id === scanId);
      if (!existing) return;
      const updated = await DB.putScan({ ...existing, deleted: true, source: 'admin', operator: state.session.account.username, note: '管理員刪除', synced: 0 });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushAdminEditScan(updated, state.session.auth); } catch {}
      }
      closeEditModal();
      await reloadState();
      showToast('已刪除此站感應並重算名次', 'info');
    });

    // 分頁 4：本活動分組鳴槍與站點
    $('btn-add-category').addEventListener('click', async () => {
      await updateCurrentEvent((ev) => {
        const n = (ev.categories || []).length + 1;
        return { categories: [...(ev.categories || []), { id: `CAT${n}`, name: `新分組 ${n}`, distanceKm: 5, gunTimeMs: null, gunTimeIso: '', color: '#a855f7' }] };
      });
    });

    $('categories-list').addEventListener('click', async (e) => {
      const evt = getActiveEvent();
      if (!evt) return;
      const cats = [...(evt.categories || [])];

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
        await updateCurrentEvent(() => ({ categories: cats }));
        showToast('分組設定已儲存', 'success');
        return;
      }

      const delBtn = e.target.closest('[data-del-cat]');
      if (delBtn) {
        const idx = Number(delBtn.dataset.delCat);
        if (!confirm(`確定刪除分組「${cats[idx]?.name}」？`)) return;
        cats.splice(idx, 1);
        await updateCurrentEvent(() => ({ categories: cats }));
        return;
      }

      const gunNow = e.target.closest('[data-gun-now]');
      if (gunNow) {
        const idx = Number(gunNow.dataset.gunNow);
        const nowSec = Math.floor(Date.now() / 1000) * 1000;
        cats[idx].gunTimeMs = nowSec;
        cats[idx].gunTimeIso = new Date(nowSec).toISOString();
        NFC.sound.beepGun();
        await updateCurrentEvent(() => ({ categories: cats }));
        showToast(`🔫 「${evt.name} - ${cats[idx].name}」已鳴槍！(${TE.formatClockSec(nowSec)})`, 'success');
        return;
      }

      const gunMan = e.target.closest('[data-gun-manual-set]');
      if (gunMan) {
        const idx = Number(gunMan.dataset.gunManualSet);
        const tsMs = TE.parseDateTimeToMs(evt.date, document.querySelector(`[data-gun-manual-input="${idx}"]`).value.trim());
        if (tsMs === null) return showToast('請輸入 HH:mm:ss 格式（如 06:30:00）', 'error');
        cats[idx].gunTimeMs = tsMs;
        cats[idx].gunTimeIso = new Date(tsMs).toISOString();
        await updateCurrentEvent(() => ({ categories: cats }));
        showToast(`已套用鳴槍時間 ${TE.formatClockSec(tsMs)}`, 'success');
        return;
      }

      const gunClr = e.target.closest('[data-gun-clear]');
      if (gunClr) {
        const idx = Number(gunClr.dataset.gunClear);
        cats[idx].gunTimeMs = null;
        cats[idx].gunTimeIso = '';
        await updateCurrentEvent(() => ({ categories: cats }));
      }
    });

    $('btn-add-checkpoint').addEventListener('click', async () => {
      await updateCurrentEvent((ev) => {
        const cps = [...(ev.checkpoints || [])];
        const maxOrd = cps.reduce((m, c) => Math.max(m, Number(c.order || 0)), 0);
        cps.push({ id: `CP${maxOrd + 1}`, name: `補給站 ${maxOrd + 1}`, order: maxOrd + 1, distanceKm: 10, type: 'checkpoint' });
        return { checkpoints: cps };
      });
    });

    $('checkpoints-list').addEventListener('click', async (e) => {
      const delBtn = e.target.closest('[data-del-cp]');
      if (!delBtn) return;
      const idx = Number(delBtn.dataset.delCp);
      await updateCurrentEvent((ev) => {
        const sorted = [...(ev.checkpoints || [])].sort((a, b) => a.order - b.order);
        sorted.splice(idx, 1);
        return { checkpoints: sorted };
      });
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
        .filter((c) => c.id && c.name)
        .sort((a, b) => a.order - b.order);
      await updateCurrentEvent(() => ({ checkpoints: cps }));
      showToast('本活動站點設定已儲存！', 'success');
    });

    // 分頁 5：作業員帳號與鎖定活動/站點管制
    $('btn-add-staff').addEventListener('click', async () => {
      const list = [...(state.platform.staffAccounts || [])];
      const n = list.length + 1;
      list.push({ username: `staff${n}`, password: '1111', displayName: `作業員 ${n}`, role: 'scanner', eventId: 'ALL', checkpointId: 'START', enabled: true, updatedAt: Date.now() });
      await DB.savePlatform({ ...state.platform, staffAccounts: list });
      await reloadState();
    });

    $('staff-accounts-list').addEventListener('click', async (e) => {
      const delBtn = e.target.closest('[data-del-staff]');
      if (!delBtn) return;
      const idx = Number(delBtn.dataset.delStaff);
      const list = [...(state.platform.staffAccounts || [])];
      if (!confirm(`確定刪除作業員帳號「${list[idx]?.username}」？`)) return;
      list.splice(idx, 1);
      const updated = await DB.savePlatform({ ...state.platform, staffAccounts: list });
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushPlatform(updated, state.session.auth);
      await reloadState();
    });

    $('btn-save-staff-accounts').addEventListener('click', async () => {
      const rows = Array.from(document.querySelectorAll('#staff-accounts-list [data-staff-row]'));
      const seen = new Set();
      const staffAccounts = [];
      for (const row of rows) {
        const username = row.querySelector('[data-staff-field="username"]').value.trim();
        if (!username) continue;
        if (seen.has(username)) return showToast(`帳號「${username}」重複`, 'error');
        seen.add(username);
        const role = row.querySelector('[data-staff-field="role"]').value;
        const cpId = row.querySelector('[data-staff-field="checkpointId"]').value.trim();
        if (role !== 'writer' && !cpId) return showToast(`帳號「${username}」屬於感應角色，必須填寫鎖定站點代碼`, 'error');
        staffAccounts.push({
          username,
          password: row.querySelector('[data-staff-field="password"]').value || '1111',
          displayName: row.querySelector('[data-staff-field="displayName"]').value.trim() || username,
          role,
          eventId: row.querySelector('[data-staff-field="eventId"]').value || 'ALL',
          checkpointId: role === 'writer' ? '' : cpId,
          enabled: row.querySelector('[data-staff-field="enabled"]').checked,
          updatedAt: Date.now()
        });
      }
      const updated = await DB.savePlatform({ ...state.platform, staffAccounts });
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushPlatform(updated, state.session.auth);
      await reloadState();
      showToast(`已儲存 ${staffAccounts.length} 組作業員帳號設定！`, 'success');
    });

    // 分頁 6：備份、Excel 匯出與雲端同步
    $('btn-export-excel-full').addEventListener('click', () => {
      const evt = getActiveEvent();
      const f = Excel.exportResultsToExcel({ event: evt, computed: state.computed, scans: state.allScans.filter((s) => s.eventId === evt.eventId), maskId: false });
      showToast(`已匯出完整 Excel：${f}`, 'success');
    });
    $('btn-export-excel-masked').addEventListener('click', () => {
      const evt = getActiveEvent();
      const f = Excel.exportResultsToExcel({ event: evt, computed: state.computed, scans: state.allScans.filter((s) => s.eventId === evt.eventId), maskId: true });
      showToast(`已匯出隱碼公開 Excel：${f}`, 'success');
    });

    $('btn-backup-download').addEventListener('click', () => downloadJsonBackup('智慧路跑多場次完整備份'));
    $('input-backup-merge').addEventListener('change', async (e) => {
      const f = e.target.files?.[0];
      if (f) {
        try {
          const r = await DB.importFullBackup(JSON.parse(await f.text()), 'merge');
          await reloadState();
          showToast(`合併完成！晶片 ${r.chipCount}、報名 ${r.regCount}、感應 ${r.scanCount} 筆`, 'success');
        } catch (err) { showToast(err.message, 'error'); }
      }
      e.target.value = '';
    });
    $('input-backup-replace').addEventListener('change', async (e) => {
      const f = e.target.files?.[0];
      if (f && confirm('確定以備份檔覆蓋還原整個本機資料庫嗎？')) {
        try {
          const r = await DB.importFullBackup(JSON.parse(await f.text()), 'replace');
          await reloadState();
          showToast(`覆蓋還原完成！晶片 ${r.chipCount}、報名 ${r.regCount}、感應 ${r.scanCount} 筆`, 'success');
        } catch (err) { showToast(err.message, 'error'); }
      }
      e.target.value = '';
    });

    document.querySelectorAll('input[name="sync-mode"]').forEach((r) => {
      r.addEventListener('change', (e) => {
        DB.saveDeviceSettings({ mode: e.target.value });
        updateNetworkIndicators();
        showToast(e.target.value === 'online' ? '已切換為多機即時連線模式' : '已切換為單機離線模式', 'info');
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
      } catch (err) { showToast(`❌ 連線測試失敗：${err.message}`, 'error'); }
    });

    $('btn-sync-upload').addEventListener('click', () => handleUploadPendingSync(false));
    $('btn-sync-pull').addEventListener('click', async () => {
      try {
        await Sync.pullFromCloud(state.session.auth);
        await reloadState();
        showToast('已從雲端拉取最新全平台活動、晶片與感應資料！', 'success');
      } catch (err) { showToast(`雲端拉取失敗：${err.message}`, 'error'); }
    });
    $('btn-sync-push-platform').addEventListener('click', async () => {
      if (!Sync.hasUrl()) return showToast('請先設定 GAS 網址', 'warn');
      try {
        await Sync.pushPlatform(state.platform, state.session.auth);
        showToast('已將所有路跑活動與作業員帳號推送至 Google 試算表！', 'success');
      } catch (err) { showToast(`推送失敗：${err.message}`, 'error'); }
    });

    const copyLink = async (path, includeEvent, label) => {
      const gasUrl = DB.getDeviceSettings().gasUrl;
      const target = new URL(path, location.href);
      if (includeEvent && state.selectedEventId) target.searchParams.set('event', state.selectedEventId);
      if (gasUrl) target.searchParams.set('api', gasUrl);
      try {
        await navigator.clipboard.writeText(target.toString());
        showToast(`已複製「${label}」！`, 'success');
      } catch { prompt(`請複製以下${label}：`, target.toString()); }
    };
    $('btn-copy-staff-link').addEventListener('click', () => copyLink('./index.html', false, '工作人員端連線網址'));
    $('btn-copy-public-link').addEventListener('click', () => copyLink('./public.html', true, '本活動公開查詢連結'));

    $('btn-load-demo-data').addEventListener('click', async () => {
      if (!confirm('載入雙場次模擬資料將會覆蓋目前本機資料，確定繼續嗎？')) return;
      await DB.loadDemoData();
      state.selectedEventId = 'EVT-2026-SPRING';
      await reloadState();
      showToast('⚡ 已載入 2 場活動與共用晶片模擬資料！同一選手已跨場報名不同組別', 'success');
    });
    $('btn-clear-all-data').addEventListener('click', async () => {
      if (!confirm('確定清空本機所有活動、晶片與感應資料嗎？')) return;
      await DB.clearAllData(true);
      await reloadState();
      showToast('本機資料庫已清空重設', 'info');
    });
  }

  document.addEventListener('DOMContentLoaded', initApp);
})();
