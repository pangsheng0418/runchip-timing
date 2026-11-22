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
    resultsCatFilter: 'ALL',
    resultsSearch: '',
    rosterSearch: '',
    writerEventFilter: 'ALL',
    adminRegSearch: '',
    adminRegCheckinFilter: 'ALL',
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

  function getLocalDateStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function getCurrentAdminUsername() {
    return state.session?.mode === 'admin' ? String(state.session.account?.username || '').trim() : '';
  }

  function isCurrentSuperAdmin() {
    return state.session?.mode === 'admin' && Boolean(state.session.account?.isSuperAdmin);
  }

  function getEventOwner(ev) {
    return String(ev?.ownerAdmin || state.platform?.adminUsername || 'admin').trim() || 'admin';
  }

  function getAdminOwnEvents() {
    const adminUser = getCurrentAdminUsername();
    if (!adminUser) return [];
    return (state.platform?.events || []).filter((e) => getEventOwner(e) === adminUser);
  }

  function isEventOwnedByCurrentAdmin(ev) {
    const adminUser = getCurrentAdminUsername();
    return Boolean(ev && adminUser && getEventOwner(ev) === adminUser);
  }

  function getAdminOwnStaffAccounts() {
    const adminUser = getCurrentAdminUsername();
    if (!adminUser) return [];
    const isSuper = isCurrentSuperAdmin();
    return (state.platform?.staffAccounts || []).filter((s) => {
      const sOwner = String(s.ownerAdmin || '').trim();
      return isSuper ? (!sOwner || sOwner === adminUser) : sOwner === adminUser;
    });
  }

  function syncNewEventModalHints(evtType) {
    const isLap = evtType === 'lap_timed';
    const year = new Date().getFullYear();
    $('new-evt-lap-km-box').classList.toggle('hidden', !isLap);
    $('new-evt-name').placeholder = isLap
      ? `例如：${year} 環湖 12 小時繞圈超馬賽`
      : `例如：${year} 高雄城市馬拉松`;
    $('new-evt-location').placeholder = isLap
      ? '例如：大安森林公園環道'
      : '例如：高雄夢時代廣場';
  }

  function getActiveEvent() {
    if (state.session?.mode === 'admin') {
      const ownEvents = getAdminOwnEvents();
      return ownEvents.find((e) => e.eventId === state.selectedEventId) || ownEvents[0] || null;
    }
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

    // 進入系統後若網路正常且已設定雲端網址，立即自動同步雲端所有最新資料，避免各裝置資料不一致！
    if (navigator.onLine && Sync.hasUrl()) {
      await autoSyncCloud({ reason: 'boot', showNotice: true });
    }

    // 網路恢復時自動雙向同步（先上傳離線感應，再拉取雲端最新）
    window.addEventListener('online', async () => {
      updateNetworkIndicators();
      if (Sync.hasUrl()) {
        showToast('網路已恢復，正在自動雙向同步雲端最新資料...', 'info');
        await autoSyncCloud({ reason: 'reconnect', showNotice: true });
      }
    });

    // 手機喚醒或切回瀏覽器分頁時，自動背景檢查最新活動與鳴槍時間（節流 15 秒）
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && navigator.onLine && Sync.hasUrl()) {
        const last = DB.getDeviceSettings().lastSyncAt || 0;
        if (Date.now() - last > 15000) {
          autoSyncCloud({ reason: 'resume', showNotice: false });
        }
      }
    });

    window.addEventListener('offline', () => {
      updateNetworkIndicators();
      showToast('已進入無網路狀態！晶片計時作業仍可離線執行，有網後再統一上傳', 'warn');
    });
  }

  let _isAutoSyncing = false;
  async function autoSyncCloud({ reason = 'manual', showNotice = false } = {}) {
    if (_isAutoSyncing || !navigator.onLine || !Sync.hasUrl()) return;
    _isAutoSyncing = true;
    const portalBanner = $('portal-net-banner');
    if (portalBanner) {
      portalBanner.className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-xs font-extrabold bg-sky-950/80 border border-sky-500/50 text-sky-300';
      portalBanner.textContent = '🔄 連線正常，正在自動同步雲端最新活動與資料...';
    }
    try {
      const uploadRes = await Sync.fullBidirectionalSync(state.session);
      await reloadState();
      const evtCount = state.session?.mode === 'admin'
        ? getAdminOwnEvents().length
        : (state.platform?.events || []).length;
      if (showNotice) {
        const uploaded = (uploadRes.uploadedChips || 0) + (uploadRes.uploadedRegs || 0) + (uploadRes.uploadedScans || 0);
        const msg = uploaded > 0 || uploadRes.platformSynced
          ? `☁️ 雲端雙向同步完成！已上傳本地異動並更新最新 ${evtCount} 場活動資料`
          : `☁️ 已同步雲端最新資料（共 ${evtCount} 場路跑活動）`;
        showToast(msg, 'success');
      }
    } catch (err) {
      console.warn(`自動同步 (${reason}) 失敗:`, err);
      if (showNotice && reason === 'manual') {
        showToast(`雲端同步失敗：${err.message}`, 'error');
      }
    } finally {
      _isAutoSyncing = false;
      updateNetworkIndicators();
    }
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

    // 若已登入後台管理員，同步最新帳號名稱與角色權限，並嚴格限制選取活動僅限本人所建活動
    if (state.session?.mode === 'admin') {
      const matchedAdmin = (platform.adminAccounts || []).find((a) => a.username === state.session.account?.username);
      if (matchedAdmin) {
        state.session.account = {
          ...state.session.account,
          displayName: matchedAdmin.displayName || matchedAdmin.username,
          isSuperAdmin: Boolean(matchedAdmin.isSuperAdmin)
        };
        DB.saveSession(state.session);
      }
      const ownEvents = getAdminOwnEvents();
      if (!state.selectedEventId || !ownEvents.some((e) => e.eventId === state.selectedEventId)) {
        const defaultId = platform.activeEventId;
        const validDefault = defaultId && ownEvents.some((e) => e.eventId === defaultId);
        state.selectedEventId = validDefault ? defaultId : (ownEvents[0]?.eventId || '');
      }
    } else {
      const events = platform.events || [];
      if (!state.selectedEventId || !events.some((e) => e.eventId === state.selectedEventId)) {
        const defaultId = platform.activeEventId;
        const validDefault = defaultId && events.some((e) => e.eventId === defaultId);
        state.selectedEventId = validDefault ? defaultId : (events[0]?.eventId || '');
        DB.saveDeviceSettings({ selectedEventId: state.selectedEventId });
      }
    }

    const evt = getActiveEvent();
    if (evt) {
      const evtRegs = regs.filter((r) => r.eventId === evt.eventId);
      const evtScans = scans.filter((s) => s.eventId === evt.eventId);
      state.computed = TE.computeEventResults({
        event: evt,
        runners: evtRegs,
        scans: evtScans
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
        ? `🛡️ ${acc.displayName || acc.username} (${acc.isSuperAdmin ? '最高管理員' : '活動管理員'})`
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
    const allEvents = state.platform?.events || [];
    const statusLabel = { active: '進行中', upcoming: '籌備中', finished: '已結束' };
    const buildOptionsHtml = (list, emptyLabel) => (
      list.length
        ? list.map((ev) => {
            const typeTag = ev.eventType === 'lap_timed' ? `🔁繞圈(${ev.lapDistanceKm}K/圈)` : '🏃距離賽';
            return `<option value="${esc(ev.eventId)}" style="background-color:#0f172a;color:#f8fafc;">${typeTag} ${esc(ev.name)} (${esc(ev.date || '')}) [${statusLabel[ev.status] || ev.status}]</option>`;
          }).join('')
        : `<option value="" style="background-color:#0f172a;color:#f8fafc;">${emptyLabel}</option>`
    );

    const allOptionsHtml = buildOptionsHtml(allEvents, '尚無活動');
    ['portal-event-select', 'staff-login-event'].forEach((id) => {
      const el = $(id);
      if (!el) return;
      el.innerHTML = allOptionsHtml;
      if (allEvents.some((e) => e.eventId === state.selectedEventId)) {
        el.value = state.selectedEventId;
      }
    });

    // 後台管理區下拉選單：嚴格僅顯示目前登入管理者本人建立的活動！
    const adminSel = $('admin-event-select');
    if (adminSel) {
      const ownEvents = getAdminOwnEvents();
      adminSel.innerHTML = buildOptionsHtml(ownEvents, '尚無您建立的活動（請點右側「＋ 建立新路跑活動」）');
      if (ownEvents.some((e) => e.eventId === state.selectedEventId)) {
        adminSel.value = state.selectedEventId;
      }
    }

    const wf = $('writer-event-filter');
    if (wf) {
      const prev = state.writerEventFilter;
      wf.innerHTML = '<option value="ALL">所有活動名單</option>' + allOptionsHtml;
      wf.value = (prev === 'ALL' || allEvents.some((e) => e.eventId === prev)) ? prev : 'ALL';
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
    const portalCloudTip = $('portal-cloud-tip');
    if (portalCloudTip) {
      portalCloudTip.classList.toggle('hidden', Boolean(settings.gasUrl));
    }
    if (portalBanner && !_isAutoSyncing) {
      if (onlineNet) {
        portalBanner.className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-xs font-extrabold bg-emerald-950/70 border border-emerald-500/40 text-emerald-300';
        const evtCount = (state.platform?.events || []).length;
        const lastStr = settings.lastSyncAt ? ` • 最後同步 ${TE.formatClockSec(settings.lastSyncAt)}` : '';
        portalBanner.textContent = settings.gasUrl
          ? `🟢 已連線雲端試算表（共 ${evtCount} 場活動）${lastStr}`
          : '🟡 尚未綁定雲端網址 • 目前僅顯示本機離線資料';
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
    let cloudLegacyFallback = false;

    if (navigator.onLine && Sync.hasUrl()) {
      try {
        const res = await Sync.loginCloud(mode, u, p);
        account = {
          ...res.account,
          password: p,
          isSuperAdmin: mode === 'admin'
            ? (res.account?.isSuperAdmin !== undefined ? Boolean(res.account.isSuperAdmin) : true)
            : false
        };
        usedCloud = true;
      } catch (err) {
        if (String(err.message).includes('錯誤') || String(err.message).includes('停用')) {
          // 若雲端 GAS 尚未升級至 v3.4 多管理者版本，允許回退驗證本機新建的管理者帳號
          if (mode === 'admin' && !err.multiAdmin) {
            cloudLegacyFallback = true;
          } else {
            NFC.sound.beepError();
            showToast(err.message, 'error');
            return;
          }
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
        // 若該作業員隸屬於特定管理者，檢查所選活動是否為該管理者建立（防止誤感應至他人活動）
        const staffOwner = String(account.ownerAdmin || '').trim();
        if (staffOwner && lockedEventId) {
          const targetEvt = events.find((e) => e.eventId === lockedEventId);
          const targetOwner = String(targetEvt?.ownerAdmin || platform.adminUsername || 'admin').trim();
          if (targetEvt && targetOwner !== staffOwner) {
            NFC.sound.beepError();
            showToast(`此作業員帳號僅限操作管理者「${staffOwner}」建立的活動，請重新選擇活動`, 'error');
            return;
          }
        }
      }
      state.selectedEventId = lockedEventId;
      DB.saveDeviceSettings({ selectedEventId: lockedEventId });
    } else if (mode === 'admin') {
      // 登入管理員時清除非本人建立的 selectedEventId
      const defaultOwner = String(platform.adminUsername || 'admin').trim();
      const ownEvents = events.filter((e) => (String(e.ownerAdmin || '').trim() || defaultOwner) === account.username);
      if (!ownEvents.some((e) => e.eventId === state.selectedEventId)) {
        state.selectedEventId = ownEvents[0]?.eventId || '';
      }
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
      const ownCount = getAdminOwnEvents().length;
      if (cloudLegacyFallback) {
        showToast(`⚠️ 已以本機管理者「${account.username}」登入（請同步更新部署 gas/Code.gs 以啟用雲端多管理者隔離）`, 'warn');
      } else {
        showToast(
          usedCloud
            ? `🛡️ 已驗證管理者「${account.displayName || account.username}」（共 ${ownCount} 場本人活動）`
            : `🛡️ 已登入管理後台：${account.displayName || account.username}（共 ${ownCount} 場本人活動）`,
          'success'
        );
      }
      switchView('admin-workspace');
      switchAdminTab('events');
    } else {
      state.staffSubTab = account.role === 'writer' ? 'writer' : 'scan';
      const evt = getActiveEvent();
      const lockedCp = getLockedCheckpoint();
      const zoneName = account.role === 'writer'
        ? '共用 NFC 晶片寫入區'
        : `${evt ? evt.name : lockedEventId} ／ ${lockedCp ? lockedCp.name : account.checkpointId}`;
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
   * VIEW 4：晶片計時作業區（鎖定活動＋報到感應／站點感應 / 共用晶片寫入）
   * ======================================================================== */
  function getLockedCheckpoint() {
    const evt = getActiveEvent();
    const acc = state.session?.account;
    if (!acc) return null;
    const accCp = String(acc.checkpointId || '').trim().toUpperCase();
    // 選手報到感應組 (CHECKIN)：同時適用「一般距離路跑」與「繞圈計時賽」，優先鎖定為報到作業！
    if (accCp === TE.CHECKIN_CP_ID) {
      return {
        id: TE.CHECKIN_CP_ID,
        name: '選手報到感應區',
        order: 0,
        distanceKm: 0,
        type: 'checkin'
      };
    }
    // 繞圈計時賽無中間管制點，非報到組的感應作業員進入繞圈賽自動鎖定「繞圈計次感應區 (LAP)」
    if (evt && evt.eventType === 'lap_timed') {
      return (evt.checkpoints && evt.checkpoints[0]) || {
        id: 'LAP',
        name: '繞圈計次感應區',
        order: 1,
        distanceKm: Number(evt.lapDistanceKm || 2.54),
        type: 'lap'
      };
    }
    if (!acc.checkpointId) return null;
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
    const isCheckIn = lockedCp?.id === TE.CHECKIN_CP_ID;
    const isLap = !isCheckIn && evt?.eventType === 'lap_timed';
    const evtRegs = state.allRegistrations.filter((r) => evt && r.eventId === evt.eventId);
    const checkedInBibSet = new Set(
      state.allScans
        .filter((s) => !s.deleted && evt && s.eventId === evt.eventId && s.checkpointId === TE.CHECKIN_CP_ID)
        .map((s) => s.bib)
    );
    const checkedInCount = evtRegs.filter((r) => checkedInBibSet.has(r.bib)).length;

    $('locked-event-title').textContent = evt
      ? (isCheckIn
          ? `📋 選手報到作業：${evt.name}（已報到 ${checkedInCount} / ${evtRegs.length} 人 • 未報到 ${Math.max(0, evtRegs.length - checkedInCount)} 人）`
          : isLap
            ? `🔁 繞圈計時賽：${evt.name}（單圈 ${evt.lapDistanceKm} km • 無管制點，每刷 1 次累積 1 圈）`
            : `🏟️ 鎖定活動：${evt.name}（${evt.date || ''}）`)
      : '尚未選定活動';
    $('locked-cp-name').textContent = lockedCp
      ? (isCheckIn ? `✅ ${lockedCp.name}` : isLap ? `🔁 ${lockedCp.name}` : `#${lockedCp.order || '-'} ${lockedCp.name}`)
      : '未指派站點';
    $('locked-cp-meta').textContent = lockedCp
      ? (isCheckIn
          ? `代碼：CHECKIN ｜ 僅登記選手到場報到（不計入起跑或繞圈成績）`
          : isLap
            ? `代碼：${lockedCp.id} ｜ 每圈計程：${evt.lapDistanceKm} km`
            : `代碼：${lockedCp.id} ｜ 累計距離：${lockedCp.distanceKm} km`)
      : '請確認本活動有對應站點代碼';

    const col3Label = $('hero-col3-label');
    if (col3Label) {
      col3Label.textContent = isCheckIn ? '報到狀態 / 本場已報到' : isLap ? '累積圈數 / 總里程' : '本場經過時間';
    }

    // 模擬器列出所有「共用晶片」（並標示該晶片是否已報名本活動／是否已報到，方便測試比對與攔截！）
    const evtRegById = new Map(
      evtRegs.map((r) => [String(r.idNumber || '').toUpperCase(), r])
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
          const reg = evtRegById.get(c.idNumber);
          let tag = '⚠️未報本場';
          if (reg) {
            tag = isCheckIn
              ? (checkedInBibSet.has(reg.bib) ? '✅已報到' : '⏳未報到')
              : '✅已報本場';
          }
          return `<option value="${esc(c.idNumber)}">${esc(c.name)} (${esc(c.idNumber)}) ${tag}</option>`;
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

    // 報到感應模式：若該選手已報到則提醒略過，避免重複報到紀錄
    if (lockedCp.id === TE.CHECKIN_CP_ID) {
      const existingCheckIn = state.allScans.find(
        (s) => !s.deleted && s.eventId === evt.eventId && s.bib === reg.bib && s.checkpointId === TE.CHECKIN_CP_ID
      );
      if (existingCheckIn) {
        NFC.sound.beepWarning();
        heroCard.classList.add('flash-warn');
        $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-amber-500 text-slate-950 mb-3';
        $('scan-status-pill').textContent = `⚠️ 已完成報到（無須重複感應）`;
        $('hero-bib-cat').textContent = `${reg.bib} • ${displayCat}`;
        $('hero-name').textContent = reg.name;
        $('hero-elapsed').textContent = `已報到`;
        $('hero-subinfo').textContent = `原報到時間：${TE.formatClockSec(existingCheckIn.timestampMs)} ｜ 活動：${evt.name}`;
        showToast(`${reg.name} (#${reg.bib}) 已於 ${TE.formatClockSec(existingCheckIn.timestampMs)} 完成報到`, 'info');
        return;
      }
    }

    // 2. 寫入本活動感應紀錄（含防連點檢查）
    const res = await DB.addScan({
      eventId: evt.eventId,
      bib: reg.bib,
      idNumber: reg.idNumber,
      checkpointId: lockedCp.id,
      timestampMs: scanTs,
      source: tagData.source || 'nfc',
      operator,
      note: tagData.note || (lockedCp.id === TE.CHECKIN_CP_ID ? '現場晶片報到' : (tagData.uidOnly ? `UID感應(${tagData.nfcUid})` : '')),
      bypassDebounce: bypassDebounce || lockedCp.id === TE.CHECKIN_CP_ID
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

    // 3. 依站點與活動性質計算顯示（報到感應 / 繞圈計時賽 / 一般距離賽）
    if (lockedCp.id === TE.CHECKIN_CP_ID) {
      const evtRegs = state.allRegistrations.filter((r) => r.eventId === evt.eventId);
      const prevCheckedBibs = new Set(
        state.allScans
          .filter((s) => !s.deleted && s.eventId === evt.eventId && s.checkpointId === TE.CHECKIN_CP_ID)
          .map((s) => s.bib)
      );
      prevCheckedBibs.add(reg.bib);

      NFC.sound.beepSuccess();
      heroCard.classList.add('flash-success');
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-sky-400 text-slate-950 mb-3';
      $('scan-status-pill').textContent = `✅ 報到成功！歡迎參加「${evt.name}」`;
      $('hero-bib-cat').textContent = `${reg.bib} • ${displayCat}`;
      $('hero-name').textContent = reg.name;
      $('hero-elapsed').textContent = `已報到 (${prevCheckedBibs.size}/${evtRegs.length})`;
      $('hero-subinfo').textContent = `報到時間：${TE.formatClockSec(scanTs)} ｜ 活動：${evt.name} ｜ 操作員：${operator}`;
    } else if (evt.eventType === 'lap_timed') {
      const lapKm = Number(evt.lapDistanceKm || 2.54);
      const gunMs = cat && cat.gunTimeMs ? Number(cat.gunTimeMs) : (evt.unifiedGunTimeMs ? Number(evt.unifiedGunTimeMs) : null);
      const hrs = cat ? Number(cat.durationHours || cat.distanceKm || 6) : 6;
      const cutMs = gunMs !== null ? gunMs + hrs * 3600 * 1000 : null;
      const isOvertime = cutMs !== null && scanTs > cutMs;

      const prevValidLaps = state.allScans
        .filter((s) => !s.deleted && s.eventId === evt.eventId && s.bib === reg.bib && s.checkpointId !== TE.CHECKIN_CP_ID && (cutMs === null || s.timestampMs <= cutMs))
        .sort((a, b) => a.timestampMs - b.timestampMs);

      const newLapCount = isOvertime ? prevValidLaps.length : prevValidLaps.length + 1;
      const totalKm = (newLapCount * lapKm).toFixed(2);
      const lastPrevTs = prevValidLaps.length ? prevValidLaps[prevValidLaps.length - 1].timestampMs : gunMs;
      const singleLapStr = lastPrevTs !== null ? TE.formatDuration(scanTs - lastPrevTs, false) : '--';

      if (isOvertime) {
        NFC.sound.beepWarning();
        heroCard.classList.add('flash-warn');
        $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-amber-500 text-slate-950 mb-3';
        $('scan-status-pill').textContent = `⏰ 已逾 ${hrs} 小時限時（已留存紀錄，不計入正式圈數）`;
      } else {
        NFC.sound.beepSuccess();
        heroCard.classList.add('flash-success');
        $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-emerald-500 text-slate-950 mb-3';
        $('scan-status-pill').textContent = `🔁 繞圈計次成功！第 ${newLapCount} 圈 (+${lapKm} km)`;
      }
      $('hero-bib-cat').textContent = `${reg.bib} • ${displayCat}`;
      $('hero-name').textContent = reg.name;
      $('hero-elapsed').textContent = `${newLapCount} 圈 • ${totalKm} km`;
      $('hero-subinfo').textContent = `感應時間：${TE.formatClockMs(scanTs)} ｜ 本圈耗時：${singleLapStr} ｜ 操作員：${operator}`;
    } else {
      const gunMs = cat && cat.gunTimeMs ? Number(cat.gunTimeMs) : (evt.unifiedGunTimeMs ? Number(evt.unifiedGunTimeMs) : null);
      const elapsedStr = gunMs !== null ? TE.formatDuration(scanTs - gunMs, true) : '尚未設定起跑時間';

      NFC.sound.beepSuccess();
      heroCard.classList.add('flash-success');
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-emerald-500 text-slate-950 mb-3';
      $('scan-status-pill').textContent = `✅ 比對成功並計時 — ${lockedCp.name}`;
      $('hero-bib-cat').textContent = `${reg.bib} • ${displayCat}`;
      $('hero-name').textContent = reg.name;
      $('hero-elapsed').textContent = elapsedStr;
      $('hero-subinfo').textContent = `感應時間：${TE.formatClockMs(scanTs)} ｜ 活動：${evt.name} ｜ 操作員：${operator}`;
    }

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

  function setWriterVerifyBox(status, title, detail, { showCancel = false, showVerifyBtn = false, showFallbackUid = false } = {}) {
    const box = $('writer-verify-box');
    const styles = {
      working: 'rounded-xl border border-sky-500/50 bg-sky-950/40 p-4 text-sm text-sky-200 space-y-2.5',
      success: 'rounded-xl border border-emerald-500/50 bg-emerald-950/40 p-4 text-sm text-emerald-200 space-y-2.5',
      warn: 'rounded-xl border border-amber-500/50 bg-amber-950/40 p-4 text-sm text-amber-200 space-y-2.5',
      error: 'rounded-xl border border-rose-500/50 bg-rose-950/40 p-4 text-sm text-rose-200 space-y-2.5'
    };
    box.className = styles[status] || styles.working;
    $('writer-verify-title').textContent = title;
    $('writer-verify-detail').textContent = detail;
    $('btn-cancel-nfc-op').classList.toggle('hidden', !showCancel);
    $('btn-verify-just-written').classList.toggle('hidden', !showVerifyBtn);
    $('btn-fallback-bind-uid').classList.toggle('hidden', !showFallbackUid);
  }

  async function executeNfcWrite(isMock = false) {
    const person = getWriterChipInput();
    if (!person.name || !person.idNumber) {
      return showToast('請填寫「選手姓名」與「身份證字號」', 'warn');
    }

    setWriterVerifyBox(
      'working',
      isMock ? '🧪 模擬寫入共用晶片...' : '📡 等待貼近 NFC 晶片寫入...',
      `準備寫入（僅姓名｜身份證）：${NFC.encodeChipPayload(person)} ｜ 請將晶片緊貼手機背面 NFC 天線並保持靜止 1～2 秒`,
      { showCancel: !isMock }
    );

    try {
      const fn = isMock ? NFC.mockWriteChip.bind(NFC) : NFC.writeChip.bind(NFC);
      const result = await fn(person, {
        onStep: (step) => setWriterVerifyBox('working', '📡 請將 NFC 晶片貼緊手機背面...', step, { showCancel: !isMock })
      });

      const finalChip = await DB.upsertChip(
        { ...person, nfcUid: result.nfcUid || 'NDEF-WRITTEN', writtenAt: Date.now() },
        false
      );

      setWriterVerifyBox(
        'success',
        `✅ 晶片寫入成功！已寫入「${result.raw}」`,
        isMock
          ? `模擬驗證完成（UID: ${result.nfcUid}），同一晶片可直接參加所有已報名活動！`
          : '資料已寫入晶片並存檔！建議點擊下方綠色按鈕「回讀驗證晶片」（移開 1 秒再貼近），同時讀取晶片 UID 與核對資料。',
        { showVerifyBtn: !isMock }
      );
      showToast(`✅ ${person.name} (${person.idNumber}) 共用晶片寫入成功！`, 'success');

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        Sync.pushChips([finalChip]).then(() => updatePendingSyncBadge());
      }
      await reloadState();
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ 共用晶片寫入失敗', err.message, { showFallbackUid: true });
      showToast('晶片寫入失敗，請查看下方詳細原因或改用 UID 綁定', 'error');
    }
  }

  /**
   * 寫入後的獨立回讀驗證（使用全新使用者點擊手勢呼叫 ndef.scan，徹底避免與 write 衝突）
   */
  async function executeVerifyWrittenChip() {
    const person = getWriterChipInput();
    setWriterVerifyBox(
      'working',
      '🔍 回讀驗證中：請將剛寫入的晶片貼近手機背面...',
      `正在核對晶片內容是否為「${person.name}|${person.idNumber}」並擷取硬體 UID...`,
      { showCancel: true }
    );

    try {
      const tag = await NFC.readSingleTag((step) => {
        setWriterVerifyBox('working', '🔍 等待貼近晶片驗證...', step, { showCancel: true });
      });

      const readName = tag.parsed?.name || '';
      const readId = tag.parsed?.idNumber || '';
      if (!readId || (person.idNumber && readId !== person.idNumber)) {
        throw new Error(`回讀驗證不符！晶片讀回「${tag.raw || '空'}」，與表單「${person.name}|${person.idNumber}」不一致。`);
      }

      const saved = await DB.upsertChip(
        { name: readName || person.name, idNumber: readId, nfcUid: tag.nfcUid || 'NDEF-VERIFIED', writtenAt: Date.now() },
        false
      );
      NFC.sound.beepWriteVerified();
      setWriterVerifyBox(
        'success',
        `✅ 回讀驗證 100% 一致！（UID: ${tag.nfcUid || '標準 NDEF'}）`,
        `確認讀回內容：${tag.raw} ｜ 姓名：${readName} ｜ 身份證：${readId}（跨活動共用晶片已就緒！）`
      );
      showToast(`✅ 回讀驗證通過！(${tag.raw})`, 'success');

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        Sync.pushChips([saved]).then(() => updatePendingSyncBadge());
      }
      await reloadState();
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ 回讀驗證失敗', err.message, { showVerifyBtn: true, showFallbackUid: true });
      showToast(err.message, 'error');
    }
  }

  async function executeBindNfcUid() {
    const person = getWriterChipInput();
    if (!person.name || !person.idNumber) {
      return showToast('請先填寫「選手姓名」與「身份證字號」再綁定晶片 UID', 'warn');
    }

    setWriterVerifyBox('working', '🔗 UID 綁定模式：請貼近 NFC 晶片...', '直接綁定晶片唯一序號 UID 至此身份證（免寫入晶片記憶體，跨活動通用）。', { showCancel: true });
    try {
      const tag = await NFC.readSingleTag((step) => setWriterVerifyBox('working', '🔗 讀取晶片 UID 中...', step, { showCancel: true }));
      if (!tag.nfcUid) throw new Error('此晶片未回傳硬體 UID，請改用標準「寫入共用晶片」');

      const saved = await DB.upsertChip({ ...person, nfcUid: tag.nfcUid, writtenAt: Date.now() }, false);
      NFC.sound.beepWriteVerified();
      setWriterVerifyBox(
        'success',
        `✅ 晶片 UID 綁定成功！（UID: ${tag.nfcUid}）`,
        `${person.name} (${person.idNumber}) 持此晶片可於本平台所有已報名活動直接感應計時！`
      );
      showToast(`已綁定晶片 UID (${tag.nfcUid}) 給 ${person.name}`, 'success');
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushChips([saved]).then(() => updatePendingSyncBadge());
      await reloadState();
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ UID 綁定失敗', err.message);
      showToast(err.message, 'error');
    }
  }

  async function executeReadTestNfc() {
    setWriterVerifyBox('working', '🔍 讀取測試晶片：請貼近手機背面感應區...', '正在檢測晶片內容、UID 與所有已報名活動...', { showCancel: true });
    try {
      const tag = await NFC.readSingleTag((step) => setWriterVerifyBox('working', '🔍 讀取測試中...', step, { showCancel: true }));
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
        `NDEF類型：${tag.recordTypes.length ? tag.recordTypes.join(',') : '無'}`,
        `原始內容：${tag.raw || '（空）'}`,
        `識別身分：${detectedName || '未知'} (${detectedId || '無身份證'})`,
        `已報名場次：${joinedText}`
      ].join(' ｜ ');

      setWriterVerifyBox('success', '✅ 晶片讀取成功！NFC 功能正常', detail);
      showToast('晶片讀取測試成功！', 'success');
    } catch (err) {
      NFC.sound.beepError();
      setWriterVerifyBox('error', '❌ 晶片讀取失敗', err.message);
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
    const adminUser = getCurrentAdminUsername();
    const isSuper = isCurrentSuperAdmin();
    const events = getAdminOwnEvents();
    const currentEvt = getActiveEvent();
    const statusBadge = {
      active: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-500/20 text-emerald-300">進行中</span>',
      upcoming: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-sky-500/20 text-sky-300">籌備中</span>',
      finished: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-slate-700 text-slate-300">已結束</span>'
    };

    const ownerBadge = $('admin-events-owner-badge');
    if (ownerBadge) {
      ownerBadge.textContent = `🔒 目前登入：${adminUser}（僅顯示本人建立的 ${events.length} 場活動）`;
    }

    if (events.length === 0) {
      $('admin-events-list').innerHTML = `
        <div class="md:col-span-2 p-6 rounded-2xl border border-dashed border-slate-700 bg-slate-950/60 text-center space-y-2">
          <div class="text-base font-black text-white">尚無您建立的路跑活動</div>
          <p class="text-xs text-slate-400">
            目前登入帳號為 <code class="text-amber-300 font-bold">${esc(adminUser)}</code>；系統已啟用活動隔離，您只會看到自己建立的活動。<br />
            請點擊上方「<strong class="text-emerald-400">＋ 建立新路跑活動</strong>」開始建立您的第一場賽事！
          </p>
        </div>
      `;
    } else {
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
                    ${ev.eventType === 'lap_timed'
                      ? `<span class="px-2 py-0.5 rounded-full text-[11px] font-black bg-purple-500/20 text-purple-300 border border-purple-500/40">🔁 繞圈計時 (${ev.lapDistanceKm}km/圈)</span>`
                      : '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-sky-500/20 text-sky-300">🏃 距離賽</span>'}
                    ${statusBadge[ev.status] || ''}
                    ${isDefault ? '<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-sky-600 text-white">公開預設</span>' : ''}
                  </div>
                  <div class="text-xs text-slate-400 font-mono-num mt-1">
                    ${esc(ev.eventId)} ｜ ${esc(ev.date || '--')}${ev.location ? ` ｜ ${esc(ev.location)}` : ''} ｜ 建立者：${esc(getEventOwner(ev))}
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
                <button type="button" data-delete-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>
              </div>
            </div>
          `;
        })
        .join('');
    }

    $('form-event-basic').classList.toggle('hidden', !currentEvt);
    $('admin-no-event-basic-hint').classList.toggle('hidden', Boolean(currentEvt));

    if (currentEvt) {
      $('cfg-evt-name').value = currentEvt.name || '';
      $('cfg-evt-date').value = currentEvt.date || '';
      $('cfg-evt-location').value = currentEvt.location || '';
      $('cfg-evt-status').value = currentEvt.status || 'active';
      $('cfg-evt-debounce').value = String(currentEvt.debounceSeconds ?? 30);
      const isLap = currentEvt.eventType === 'lap_timed';
      $('cfg-evt-type-badge').textContent = isLap ? '🔁 繞圈計時賽' : '🏃 一般距離賽';
      $('cfg-evt-type-badge').className = isLap
        ? 'px-2.5 py-2 rounded-xl text-xs font-black bg-purple-900/60 text-purple-200 border border-purple-500/40 shrink-0'
        : 'px-2.5 py-2 rounded-xl text-xs font-black bg-slate-800 text-sky-300 shrink-0';
      $('cfg-evt-lap-km').classList.toggle('hidden', !isLap);
      if (isLap) $('cfg-evt-lap-km').value = String(currentEvt.lapDistanceKm || 2.54);
    }

    const selfAcc = (state.platform.adminAccounts || []).find((a) => a.username === adminUser) || {
      username: adminUser,
      password: state.session?.auth?.password || '',
      displayName: state.session?.account?.displayName || adminUser
    };
    const userInput = $('cfg-admin-user');
    userInput.value = selfAcc.username;
    userInput.readOnly = !isSuper;
    userInput.className = isSuper
      ? 'w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-sm font-mono-num text-white'
      : 'w-full bg-slate-900/70 border border-slate-800 rounded-xl px-3 py-2 text-sm font-mono-num text-slate-400 cursor-not-allowed';
    $('cfg-admin-displayname').value = selfAcc.displayName || selfAcc.username;
    $('cfg-admin-pass').value = selfAcc.password || '';

    // 最高管理員專用：多管理者帳號管理區塊
    const manageCard = $('admin-accounts-manage-card');
    if (manageCard) {
      manageCard.classList.toggle('hidden', !isSuper);
      if (isSuper) {
        const admins = state.platform.adminAccounts || [];
        $('admin-accounts-list').innerHTML = admins
          .map((a, idx) => {
            const isSelf = a.username === adminUser;
            const ownEvtCount = (state.platform.events || []).filter((e) => getEventOwner(e) === a.username).length;
            return `
              <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-end p-3.5 rounded-xl bg-slate-950 border ${isSelf ? 'border-amber-500/40' : 'border-slate-800'}" data-admin-row="${idx}" data-admin-orig-user="${esc(a.username)}">
                <div class="sm:col-span-2">
                  <label class="block text-[11px] text-slate-400 mb-1">登入帳號</label>
                  <input type="text" value="${esc(a.username)}" data-admin-field="username" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-mono-num font-bold text-white" />
                </div>
                <div class="sm:col-span-2">
                  <label class="block text-[11px] text-slate-400 mb-1">登入密碼</label>
                  <input type="text" value="${esc(a.password)}" data-admin-field="password" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-mono-num text-white" />
                </div>
                <div class="sm:col-span-3">
                  <label class="block text-[11px] text-slate-400 mb-1">管理者名稱（名下 ${ownEvtCount} 場活動）</label>
                  <input type="text" value="${esc(a.displayName || a.username)}" data-admin-field="displayName" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-bold text-white" />
                </div>
                <div class="sm:col-span-3">
                  <label class="block text-[11px] text-slate-400 mb-1">管理權限角色</label>
                  <select data-admin-field="isSuperAdmin" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-bold text-amber-300" ${isSelf ? 'disabled' : ''}>
                    <option value="0" ${!a.isSuperAdmin ? 'selected' : ''}>活動管理員（僅管自己活動）</option>
                    <option value="1" ${a.isSuperAdmin ? 'selected' : ''}>最高管理員（可增刪管理者）</option>
                  </select>
                </div>
                <div class="sm:col-span-2 flex items-center justify-between gap-2 pb-1">
                  <label class="inline-flex items-center gap-1.5 text-xs text-slate-300 cursor-pointer">
                    <input type="checkbox" data-admin-field="enabled" ${a.enabled !== false ? 'checked' : ''} ${isSelf ? 'disabled' : ''} class="accent-emerald-500 w-4 h-4" />
                    <span>啟用</span>
                  </label>
                  ${isSelf
                    ? '<span class="px-2 py-1 rounded text-[11px] font-bold bg-amber-500/20 text-amber-300">目前登入</span>'
                    : `<button type="button" data-del-admin="${idx}" class="px-2.5 py-1 rounded text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>`}
                </div>
              </div>
            `;
          })
          .join('');
      }
    }
  }

  function renderAdminRosterTab() {
    const evt = getActiveEvent();
    const tbody = $('admin-reg-tbody');
    if (!evt) {
      $('roster-import-event-name').textContent = '尚未建立活動';
      $('admin-reg-count').textContent = '0';
      $('admin-roster-checkin-summary').textContent = '✅ 已報到 0 / 0 人';
      tbody.innerHTML = '<tr><td colspan="7" class="py-6 text-center text-slate-500">您尚未建立任何活動，請先點擊上方「＋ 建立新路跑活動」後再匯入選手名單</td></tr>';
      return;
    }
    $('roster-import-event-name').textContent = evt.name;

    const chipMap = new Map(state.chips.map((c) => [c.idNumber, c]));
    const catMap = new Map((evt.categories || []).map((c) => [c.id, c.name]));
    const checkInMap = new Map();
    state.allScans
      .filter((s) => !s.deleted && s.eventId === evt.eventId && s.checkpointId === TE.CHECKIN_CP_ID)
      .sort((a, b) => Number(a.timestampMs) - Number(b.timestampMs))
      .forEach((s) => {
        if (!checkInMap.has(s.bib) || s.source === 'admin') checkInMap.set(s.bib, s);
      });

    const allEvtRegs = state.allRegistrations.filter((r) => r.eventId === evt.eventId);
    const totalCount = allEvtRegs.length;
    const checkedInCount = allEvtRegs.filter((r) => checkInMap.has(r.bib)).length;
    $('admin-roster-checkin-summary').textContent = `✅ 已報到 ${checkedInCount} / ${totalCount} 人（未報到 ${Math.max(0, totalCount - checkedInCount)} 人）`;

    const q = state.adminRegSearch.trim().toLowerCase();
    const chkFilter = state.adminRegCheckinFilter || 'ALL';
    const filterEl = $('admin-reg-checkin-filter');
    if (filterEl) filterEl.value = chkFilter;

    const evtRegs = allEvtRegs
      .filter((r) => {
        const isChk = checkInMap.has(r.bib);
        if (chkFilter === 'checked_in' && !isChk) return false;
        if (chkFilter === 'not_checked_in' && isChk) return false;
        if (q) return r.bib.toLowerCase().includes(q) || r.name.toLowerCase().includes(q) || (r.idNumber || '').toLowerCase().includes(q);
        return true;
      })
      .sort((a, b) => a.bib.localeCompare(b.bib, 'zh-Hant', { numeric: true }));

    $('admin-reg-count').textContent = String(totalCount);
    if (evtRegs.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" class="py-6 text-center text-slate-500">查無符合條件的選手名單（請匯入名單或切換報到篩選條件）</td></tr>';
      return;
    }

    tbody.innerHTML = evtRegs
      .map((r) => {
        const chip = r.idNumber ? chipMap.get(r.idNumber) : null;
        const chkScan = checkInMap.get(r.bib);
        const checkInCell = chkScan
          ? `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-sky-500/20 text-sky-300 font-mono-num">✅ ${TE.formatClockSec(chkScan.timestampMs)}</span>`
          : `<button type="button" data-quick-checkin-bib="${esc(r.bib)}" data-quick-checkin-idnum="${esc(r.idNumber || '')}" class="px-2 py-0.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-sky-600/40 text-slate-300 hover:text-sky-200 border border-slate-700">⏳ 未報到（點此補報到）</button>`;
        return `
          <tr class="hover:bg-slate-800/40">
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-300">${esc(r.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(r.name)}</td>
            <td class="py-2.5 pr-3 font-mono-num text-sky-300">${esc(r.idNumber || r.idMasked || '--')}</td>
            <td class="py-2.5 pr-3 text-slate-300">${esc(catMap.get(r.categoryId) || r.categoryId)}</td>
            <td class="py-2.5 pr-3">${checkInCell}</td>
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
    const tbody = $('results-tbody');
    if (!evt || !computed) {
      $('kpi-total').textContent = '0';
      $('kpi-checkin').textContent = '0';
      $('kpi-finished').textContent = '0';
      $('kpi-oncourse').textContent = '0';
      $('kpi-notstarted').textContent = '0';
      $('results-category-filter').innerHTML = '<option value="ALL">全部組別</option>';
      $('results-thead').innerHTML = '';
      tbody.innerHTML = '<tr><td class="py-8 text-center text-slate-500">您尚未建立任何活動，請先點擊上方「＋ 建立新路跑活動」</td></tr>';
      return;
    }

    const isLap = evt.eventType === 'lap_timed';
    const s = computed.summary;

    const catFilter = $('results-category-filter');
    const cats = evt.categories || [];
    catFilter.innerHTML = '<option value="ALL">全部組別</option>' + cats.map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    if (!cats.some((c) => c.id === state.resultsCatFilter)) state.resultsCatFilter = 'ALL';
    catFilter.value = state.resultsCatFilter;

    const q = state.resultsSearch.trim().toLowerCase();
    const filtered = computed.rows.filter((r) => {
      if (state.resultsCatFilter !== 'ALL' && r.categoryId !== state.resultsCatFilter) return false;
      if (q) return r.bib.toLowerCase().includes(q) || r.name.toLowerCase().includes(q) || (r.idNumber || '').toLowerCase().includes(q);
      return true;
    });

    if (isLap) {
      $('kpi-label-1').textContent = '本場報名人數';
      $('kpi-label-checkin').textContent = '已報到人數';
      $('kpi-label-2').textContent = '全場累積總圈數';
      $('kpi-label-3').textContent = '全場累積總里程 (km)';
      $('kpi-label-4').textContent = '領先最高里程 (km)';

      $('kpi-total').textContent = String(s.totalRunners);
      $('kpi-checkin').textContent = String(s.checkedInCount ?? 0);
      $('kpi-finished').textContent = `${s.totalLapsAll} 圈`;
      $('kpi-oncourse').textContent = s.totalKmAll.toFixed(2);
      $('kpi-notstarted').textContent = s.topDistanceKm.toFixed(2);

      $('results-edit-banner').textContent = `🔁 繞圈計時賽模式（單圈 ${evt.lapDistanceKm} km • 依統一律定起跑時間與累積總里程排名）：點擊「報到時間」可補改報到，點「逐圈明細」可補刪圈。`;

      $('results-thead').innerHTML = `
        <tr>
          <th class="px-3 py-3">總名次</th>
          <th class="px-3 py-3">組名次</th>
          <th class="px-3 py-3">編號</th>
          <th class="px-3 py-3">姓名</th>
          <th class="px-3 py-3">身份證</th>
          <th class="px-3 py-3">限時組別</th>
          <th class="px-3 py-3">報到時間</th>
          <th class="px-3 py-3">累積圈數</th>
          <th class="px-3 py-3">累積總里程</th>
          <th class="px-3 py-3">累積耗時</th>
          <th class="px-3 py-3">最後通過時間</th>
          <th class="px-3 py-3">最速 / 平均單圈</th>
          <th class="px-3 py-3">平均配速</th>
          <th class="px-3 py-3 text-right">逐圈管理</th>
        </tr>
      `;

      if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="14" class="py-8 text-center text-slate-500">本場繞圈活動尚無成績資料</td></tr>';
        return;
      }

      tbody.innerHTML = filtered
        .map((row) => {
          const regId = row.runner.regId || `${evt.eventId}__${row.idNumber || row.bib}`;
          const chk = row.checkInScan;
          return `
            <tr class="hover:bg-slate-800/40 align-middle">
              <td class="px-3 py-2.5 font-mono-num font-black text-amber-400">${row.overallRank ?? '-'}</td>
              <td class="px-3 py-2.5 font-mono-num font-bold text-sky-300">${row.categoryRank ?? '-'}</td>
              <td class="px-3 py-2.5 font-mono-num font-bold text-white">${esc(row.bib)}</td>
              <td class="px-3 py-2.5 font-bold text-white">${esc(row.name)}</td>
              <td class="px-3 py-2.5 font-mono-num text-slate-400">${esc(row.idNumber || row.idMasked || '--')}</td>
              <td class="px-3 py-2.5">
                <span class="inline-flex items-center gap-1.5 text-xs font-bold">
                  <span class="w-2 h-2 rounded-full" style="background:${esc(row.categoryColor)}"></span>
                  ${esc(row.categoryName)} (${row.durationHours}H)
                </span>
              </td>
              <td class="px-3 py-2.5 font-mono-num">
                <button type="button"
                  data-edit-scan-reg="${esc(regId)}"
                  data-edit-scan-bib="${esc(row.bib)}"
                  data-edit-scan-idnum="${esc(row.idNumber || '')}"
                  data-edit-scan-cp="CHECKIN"
                  data-edit-scan-id="${esc(chk ? chk.id : '')}"
                  class="px-2 py-1 rounded-lg border border-dashed ${chk ? 'border-sky-500/50 text-sky-300 hover:bg-sky-900/40' : 'border-slate-700 text-slate-500 hover:border-sky-400 hover:text-sky-300'} transition text-xs">
                  ${chk ? `✅ ${TE.formatClockSec(chk.timestampMs)}` : '+ 補報到'}
                </button>
              </td>
              <td class="px-3 py-2.5 font-mono-num font-black text-purple-300 text-base">${row.lapCount} 圈${row.overtimeCount ? ` <span class="text-[10px] text-rose-400">(+${row.overtimeCount}超時)</span>` : ''}</td>
              <td class="px-3 py-2.5 font-mono-num font-black text-emerald-400 text-base">${row.totalDistanceKm.toFixed(2)} km</td>
              <td class="px-3 py-2.5 font-mono-num text-sky-300">${row.totalElapsedMs !== null ? TE.formatDuration(row.totalElapsedMs, false) : '--'}</td>
              <td class="px-3 py-2.5 font-mono-num text-slate-300">${row.lastLapTimeMs ? TE.formatClockSec(row.lastLapTimeMs) : '--'}</td>
              <td class="px-3 py-2.5 font-mono-num text-xs text-slate-300">
                ${row.bestLapMs !== null ? TE.formatDuration(row.bestLapMs, false) : '--'} / ${row.avgLapMs !== null ? TE.formatDuration(row.avgLapMs, false) : '--'}
              </td>
              <td class="px-3 py-2.5 font-mono-num text-slate-300">${esc(row.avgPace)}</td>
              <td class="px-3 py-2.5 text-right">
                <button type="button" data-open-lap-bib="${esc(row.bib)}"
                  class="px-3 py-1.5 rounded-lg text-xs font-black bg-purple-600/30 hover:bg-purple-600/50 border border-purple-500/40 text-purple-200 transition">
                  🔁 逐圈明細 / 補刪圈
                </button>
              </td>
            </tr>
          `;
        })
        .join('');
      return;
    }

    // 一般距離路跑（統一律定起跑時間計時，不另計個人淨時間）
    $('kpi-label-1').textContent = '本場報名人數';
    $('kpi-label-checkin').textContent = '已報到人數';
    $('kpi-label-2').textContent = '已完賽';
    $('kpi-label-3').textContent = '賽道進行中';
    $('kpi-label-4').textContent = '未報到';
    $('results-edit-banner').textContent = '🛡️ 管理者修正模式（統一律定起跑時間計時）：點擊下方任一選手的「報到／站點時間」或「罰時」即可補登、修改或刪除，自動重算名次。';

    $('kpi-total').textContent = String(s.totalRunners);
    $('kpi-checkin').textContent = String(s.checkedInCount ?? 0);
    $('kpi-finished').textContent = String(s.finishedCount);
    $('kpi-oncourse').textContent = String(s.onCourseCount);
    $('kpi-notstarted').textContent = String(s.notCheckedInCount ?? 0);

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
        <th class="px-3 py-3">報到時間</th>
        ${checkpoints.map((cp) => `<th class="px-3 py-3">${esc(cp.name)}</th>`).join('')}
        <th class="px-3 py-3">分段 / 配速</th>
        <th class="px-3 py-3">罰時</th>
        <th class="px-3 py-3">完賽時間</th>
        <th class="px-3 py-3">平均配速</th>
      </tr>
    `;

    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${12 + checkpoints.length}" class="py-8 text-center text-slate-500">本活動尚無成績資料</td></tr>`;
      return;
    }

    const statusMap = {
      finished: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-500/20 text-emerald-300">已完賽</span>',
      on_course: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-amber-500/20 text-amber-300">賽道上</span>',
      checked_in: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-sky-500/20 text-sky-300">已報到</span>',
      not_checked_in: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-slate-800 text-slate-400">未報到</span>'
    };

    tbody.innerHTML = filtered
      .map((row) => {
        const regId = row.runner.regId || `${evt.eventId}__${row.idNumber || row.bib}`;
        const chk = row.checkInScan;
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
            <td class="px-3 py-2.5 font-mono-num">
              <button type="button"
                data-edit-scan-reg="${esc(regId)}"
                data-edit-scan-bib="${esc(row.bib)}"
                data-edit-scan-idnum="${esc(row.idNumber || '')}"
                data-edit-scan-cp="CHECKIN"
                data-edit-scan-id="${esc(chk ? chk.id : '')}"
                class="px-2 py-1 rounded-lg border border-dashed ${chk ? 'border-sky-500/50 text-sky-300 hover:bg-sky-900/40' : 'border-slate-700 text-slate-500 hover:border-sky-400 hover:text-sky-300'} transition text-xs">
                ${chk ? `✅ ${TE.formatClockSec(chk.timestampMs)}` : '+ 補報到'}
              </button>
            </td>
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
            <td class="px-3 py-2.5 font-mono-num text-slate-300">${esc(row.avgPace)}</td>
          </tr>
        `;
      })
      .join('');
  }

  function renderAdminRaceTab() {
    const evt = getActiveEvent();
    if (!evt) {
      $('unified-gun-display').textContent = '尚未建立活動';
      $('unified-gun-input').value = '';
      $('admin-lap-no-cp-banner').classList.add('hidden');
      $('admin-checkpoints-card').classList.remove('hidden');
      $('categories-list').innerHTML = '<div class="p-4 rounded-xl bg-slate-950 border border-slate-800 text-xs text-slate-500 text-center">請先點擊上方「＋ 建立新路跑活動」後再設定起跑時間與分組</div>';
      $('checkpoints-list').innerHTML = '<div class="p-4 rounded-xl bg-slate-950 border border-slate-800 text-xs text-slate-500 text-center">尚無活動站點</div>';
      return;
    }
    const isLap = evt.eventType === 'lap_timed';
    const unifiedGunMs = evt.unifiedGunTimeMs || (evt.categories || []).find((c) => c.gunTimeMs)?.gunTimeMs || null;
    $('unified-gun-display').textContent = unifiedGunMs ? TE.formatClockSec(unifiedGunMs) : '尚未設定';
    $('unified-gun-display').className = `font-mono-num text-2xl font-black ${unifiedGunMs ? 'text-emerald-400' : 'text-slate-500'}`;
    $('unified-gun-input').value = unifiedGunMs ? TE.formatClockSec(unifiedGunMs) : '';

    $('admin-lap-no-cp-banner').classList.toggle('hidden', !isLap);
    $('admin-checkpoints-card').classList.toggle('hidden', isLap);

    $('categories-list').innerHTML = (evt.categories || [])
      .map((c, idx) => {
        const hrs = Number(c.durationHours || c.distanceKm || 6);
        const effGun = c.gunTimeMs || unifiedGunMs;
        const cutoffStr = isLap && effGun ? ` ｜ 限時截止：${TE.formatClockSec(Number(effGun) + hrs * 3600000)}` : '';
        return `
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
              <label class="block text-[11px] text-slate-400 mb-1">${isLap ? '限時小時數 (H)' : '距離 (km)'}</label>
              <input type="number" step="${isLap ? '0.5' : '0.001'}" value="${esc(isLap ? hrs : c.distanceKm)}" data-cat-field="distanceKm" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-sm font-mono-num text-white" />
            </div>
            <div class="sm:col-span-4 flex flex-wrap gap-1.5">
              <button type="button" data-save-cat="${idx}" class="px-3 py-2 rounded-lg text-xs font-bold bg-sky-600 hover:bg-sky-500 text-white">儲存分組</button>
              <button type="button" data-del-cat="${idx}" class="px-3 py-2 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>
            </div>
          </div>

          <div class="pt-2 border-t border-slate-800/80 flex flex-wrap items-center justify-between gap-2.5">
            <div class="flex items-center gap-2 flex-wrap">
              <span class="text-xs text-slate-400 font-bold">律定起跑時間：</span>
              <span class="font-mono-num text-base font-black ${effGun ? 'text-emerald-400' : 'text-slate-500'}">
                ${effGun ? TE.formatClockSec(effGun) : '尚未設定'}
              </span>
              ${cutoffStr ? `<span class="text-xs font-mono-num text-amber-300">${cutoffStr}</span>` : ''}
            </div>
            <div class="flex flex-wrap items-center gap-2">
              <button type="button" data-gun-now="${idx}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-emerald-600/30 border border-slate-700 text-emerald-300">
                🔫 個別鳴槍
              </button>
              <input type="text" placeholder="HH:mm:ss" value="${effGun ? TE.formatClockSec(effGun) : ''}" data-gun-manual-input="${idx}"
                class="w-28 font-mono-num text-xs bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white" />
              <button type="button" data-gun-manual-set="${idx}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-sky-300">套用此組</button>
              <button type="button" data-gun-clear="${idx}" class="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-400">清除</button>
            </div>
          </div>
        </div>
      `;
      })
      .join('');

    const sortedCps = [...(evt.checkpoints || [])]
      .filter((cp) => cp && cp.type !== 'start' && String(cp.id || '').toUpperCase() !== 'START')
      .sort((a, b) => a.order - b.order);
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
              <option value="checkpoint" ${cp.type !== 'finish' ? 'selected' : ''}>中途補給/折返</option>
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
    const isSuper = isCurrentSuperAdmin();
    const list = getAdminOwnStaffAccounts();
    const events = getAdminOwnEvents();
    const evt = getActiveEvent();
    const cps = [...(evt?.checkpoints || [])].sort((a, b) => a.order - b.order);

    const allOptionLabel = isSuper ? '所有活動（登入選取）' : '本人建立的所有活動';
    const listHtml = list.length === 0
      ? '<div class="p-5 rounded-xl bg-slate-950 border border-dashed border-slate-800 text-center text-xs text-slate-500">您尚未建立專屬作業員帳號，可點擊右上角「＋ 新增作業員帳號」建立（僅供您本人建立的活動使用）。</div>'
      : list
          .map(
            (s, idx) => `
            <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-end p-3.5 rounded-xl bg-slate-950 border border-slate-800" data-staff-row="${idx}" data-staff-owner="${esc(s.ownerAdmin || '')}">
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
                  <option value="ALL" ${(!s.eventId || s.eventId === 'ALL') ? 'selected' : ''}>${allOptionLabel}</option>
                  ${events.map((e) => `<option value="${esc(e.eventId)}" ${s.eventId === e.eventId ? 'selected' : ''}>僅限：${esc(e.name)}</option>`).join('')}
                </select>
              </div>
              <div class="sm:col-span-2">
                <label class="block text-[11px] text-slate-400 mb-1">鎖定站點代碼</label>
                <input list="cp-datalist" type="text" value="${esc(s.checkpointId)}" data-staff-field="checkpointId" placeholder="CHECKIN/CP1/FINISH"
                  class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-mono-num text-sky-300" />
              </div>
              <div class="sm:col-span-1 flex items-center justify-between gap-1 pb-1">
                <input type="checkbox" data-staff-field="enabled" ${s.enabled !== false ? 'checked' : ''} class="accent-emerald-500 w-4 h-4" title="啟用" />
                <button type="button" data-del-staff="${esc(s.username)}" class="px-2 py-1 rounded text-xs font-bold bg-rose-600/20 text-rose-300">刪</button>
              </div>
            </div>
          `
          )
          .join('');

    $('staff-accounts-list').innerHTML =
      listHtml +
      `<datalist id="cp-datalist"><option value="CHECKIN">選手報到感應區 (CHECKIN)</option>${cps.filter((cp) => cp.type !== 'start' && String(cp.id || '').toUpperCase() !== 'START').map((cp) => `<option value="${esc(cp.id)}">${esc(cp.name)}</option>`).join('')}</datalist>`;
  }

  function renderAdminSyncTab() {
    const settings = DB.getDeviceSettings();
    const isSuper = isCurrentSuperAdmin();
    document.querySelectorAll('input[name="sync-mode"]').forEach((radio) => {
      radio.checked = radio.value === settings.mode;
    });
    $('input-gas-url').value = settings.gasUrl || '';
    const replaceLabel = $('btn-backup-replace-label');
    if (replaceLabel) replaceLabel.classList.toggle('hidden', !isSuper);
    const maintCard = $('admin-maintenance-card');
    if (maintCard) maintCard.classList.toggle('hidden', !isSuper);
  }

  /* ========================================================================
   * 輔助：更新當前活動並同步（嚴格驗證活動歸屬權）
   * ======================================================================== */
  async function updateCurrentEvent(mutator) {
    const events = [...(state.platform.events || [])];
    const idx = events.findIndex((e) => e.eventId === state.selectedEventId);
    if (idx < 0) {
      showToast('請先建立或選取您本人的路跑活動', 'warn');
      return;
    }
    if (!isEventOwnedByCurrentAdmin(events[idx])) {
      showToast('權限不足：不可修改其他管理者建立的活動', 'error');
      return;
    }
    events[idx] = {
      ...events[idx],
      ...mutator(events[idx]),
      ownerAdmin: getCurrentAdminUsername(),
      updatedAt: Date.now()
    };
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
      const res = await Sync.fullBidirectionalSync(state.session);
      await reloadState();
      const evtCount = state.session?.mode === 'admin'
        ? getAdminOwnEvents().length
        : (state.platform?.events || []).length;
      if (!silent || res.platformSynced || res.uploadedScans > 0 || res.uploadedChips > 0 || res.uploadedRegs > 0) {
        showToast(
          `☁️ 雙向同步完成！${res.platformSynced ? '活動設定已上傳、' : ''}晶片 ${res.uploadedChips}、報名 ${res.uploadedRegs}、感應 ${res.uploadedScans} 筆（共 ${evtCount} 場活動）` +
            (res.skippedDuplicate ? `（略過重複 ${res.skippedDuplicate} 筆）` : ''),
          'success'
        );
      } else {
        showToast(`☁️ 已與雲端同步完成（共 ${evtCount} 場活動，資料完全一致）`, 'success');
      }
    } catch (err) {
      showToast(`雲端同步失敗：${err.message}`, 'error');
    }
  }

  async function downloadJsonBackup(prefix = '智慧路跑多場次完整備份') {
    const opts = state.session?.mode === 'admin'
      ? { ownerAdmin: getCurrentAdminUsername(), isSuperAdmin: isCurrentSuperAdmin() }
      : {};
    const backup = await DB.exportFullBackup(opts);
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${prefix}_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast(`已下載備份（${backup.platform.events.length} 場活動、${backup.chips.length} 張晶片、${backup.scans.length} 筆感應）`, 'success');
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
    if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRegistrations(saved, state.session?.auth).then(() => updatePendingSyncBadge());
    await reloadState();
    showToast(`已成功匯入 ${valid.length} 位選手至「${evt.name}」！`, 'success');
  }

  /* ---------------- Modals ---------------- */
  function openEditScanModal(regId, bib, idNumber, cpId, scanId) {
    const evt = getActiveEvent();
    const reg = state.allRegistrations.find((r) => r.regId === regId);
    const isCheckIn = cpId === TE.CHECKIN_CP_ID;
    const cp = isCheckIn ? { id: TE.CHECKIN_CP_ID, name: '選手報到感應區' } : (evt?.checkpoints || []).find((c) => c.id === cpId);
    const scan = scanId ? state.allScans.find((s) => s.id === scanId) : null;

    state.editMode = 'scan';
    $('modal-edit-title').textContent = isCheckIn
      ? `人工修正報到時間 — ${bib} ${reg?.name || ''}`
      : `人工修正 — ${bib} ${reg?.name || ''} @ ${cp?.name || cpId}`;
    $('edit-reg-id').value = regId;
    $('edit-bib').value = bib;
    $('edit-idnumber').value = idNumber;
    $('edit-cp-id').value = cpId;
    $('edit-scan-id').value = scanId || '';
    $('edit-time-str').value = scan ? TE.formatClockMs(scan.timestampMs) : TE.formatClockMs(Date.now());
    $('edit-note').value = scan?.note || (isCheckIn ? '後台人工補報到' : '後台人工修正');

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

    // 大廳導覽、手動同步與快速綁定雲端網址
    $('btn-portal-staff').addEventListener('click', () => switchView('staff-login'));
    $('btn-portal-admin').addEventListener('click', () => switchView('admin-login'));
    $('btn-portal-sync-now').addEventListener('click', async () => {
      if (!Sync.hasUrl()) {
        showToast('尚未設定雲端網址，請點下方「🔗 綁定雲端網址」或從後台設定', 'warn');
        return;
      }
      await autoSyncCloud({ reason: 'manual', showNotice: true });
    });
    $('btn-portal-bind-url').addEventListener('click', async () => {
      const url = prompt('請貼上主辦單位的 Google Apps Script 網頁應用程式網址 (https://script.google.com/macros/s/.../exec)：');
      if (!url || !url.trim()) return;
      DB.saveDeviceSettings({ gasUrl: url.trim(), mode: 'online' });
      renderHeader();
      updateNetworkIndicators();
      await autoSyncCloud({ reason: 'manual', showNotice: true });
    });
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
    $('btn-fallback-bind-uid').addEventListener('click', executeBindNfcUid);
    $('btn-verify-just-written').addEventListener('click', executeVerifyWrittenChip);
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

    // 建立新路跑活動 Modal（支援一般距離賽與繞圈計時賽，依活動性質動態切換範例提示）
    const closeNewEvt = () => $('modal-new-event').classList.add('hidden');
    $('new-evt-type').addEventListener('change', (e) => {
      syncNewEventModalHints(e.target.value);
    });
    $('btn-open-new-event-modal').addEventListener('click', () => {
      $('new-evt-type').value = 'distance';
      syncNewEventModalHints('distance');
      $('new-evt-lap-km').value = '2.54';
      $('new-evt-name').value = '';
      $('new-evt-date').value = getLocalDateStr();
      $('new-evt-location').value = '';
      $('modal-new-event').classList.remove('hidden');
    });
    $('btn-close-new-event').addEventListener('click', closeNewEvt);
    $('btn-cancel-new-event').addEventListener('click', closeNewEvt);

    $('form-new-event').addEventListener('submit', async (e) => {
      e.preventDefault();
      const ownerAdmin = getCurrentAdminUsername();
      if (!ownerAdmin) return showToast('請先登入管理者帳號', 'error');
      const evtType = $('new-evt-type').value;
      const lapKm = Number($('new-evt-lap-km').value || 2.54);
      const newEvt = DB.createEventTemplate(
        $('new-evt-name').value.trim(),
        $('new-evt-date').value,
        $('new-evt-location').value.trim(),
        evtType,
        lapKm,
        ownerAdmin
      );
      const updated = await DB.savePlatform({
        ...state.platform,
        events: [...(state.platform.events || []), newEvt]
      });
      state.selectedEventId = newEvt.eventId;
      DB.saveDeviceSettings({ selectedEventId: newEvt.eventId });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushPlatform(updated, state.session.auth); } catch (err) { showToast(`活動已建於本機，雲端同步待重試：${err.message}`, 'warn'); }
      }
      closeNewEvt();
      await reloadState();
      showToast(
        evtType === 'lap_timed'
          ? `🔁 已建立繞圈計時賽「${newEvt.name}」(單圈 ${lapKm} km，預設 6H/12H/24H)！`
          : `🏟️ 已建立新活動「${newEvt.name}」！請接著設定分組並匯入選手名單`,
        'success'
      );
    });

    // 活動卡片操作（嚴格檢查僅可操作本人建立的活動）
    $('admin-events-list').addEventListener('click', async (e) => {
      const selBtn = e.target.closest('[data-select-event]');
      if (selBtn) {
        const target = (state.platform.events || []).find((x) => x.eventId === selBtn.dataset.selectEvent);
        if (!isEventOwnedByCurrentAdmin(target)) return showToast('無權限切換至其他管理者的活動', 'error');
        await onEventChange(selBtn.dataset.selectEvent);
        return;
      }

      const defBtn = e.target.closest('[data-default-event]');
      if (defBtn) {
        const target = (state.platform.events || []).find((x) => x.eventId === defBtn.dataset.defaultEvent);
        if (!isEventOwnedByCurrentAdmin(target)) return showToast('無權限設定其他管理者的活動', 'error');
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
        if (!isEventOwnedByCurrentAdmin(target)) return showToast('無權限刪除其他管理者的活動', 'error');
        if (!confirm(`確定要刪除活動「${target?.name || eid}」嗎？`)) return;
        const remain = (state.platform.events || []).filter((x) => x.eventId !== eid);
        const remainOwn = remain.filter((x) => isEventOwnedByCurrentAdmin(x));
        const updated = await DB.savePlatform({
          ...state.platform,
          events: remain,
          activeEventId: state.platform.activeEventId === eid ? (remainOwn[0]?.eventId || remain[0]?.eventId || '') : state.platform.activeEventId
        });
        if (state.selectedEventId === eid) state.selectedEventId = remainOwn[0]?.eventId || '';
        if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushPlatform(updated, state.session.auth);
        await reloadState();
        showToast('已刪除該場路跑活動', 'info');
      }
    });

    $('form-event-basic').addEventListener('submit', async (e) => {
      e.preventDefault();
      await updateCurrentEvent((ev) => {
        const patch = {
          name: $('cfg-evt-name').value.trim(),
          date: $('cfg-evt-date').value,
          location: $('cfg-evt-location').value.trim(),
          status: $('cfg-evt-status').value,
          debounceSeconds: Number($('cfg-evt-debounce').value || 30)
        };
        if (ev.eventType === 'lap_timed') {
          const newLapKm = Math.max(0.05, Number($('cfg-evt-lap-km').value || 2.54));
          patch.lapDistanceKm = newLapKm;
          if (ev.checkpoints && ev.checkpoints[0]) {
            patch.checkpoints = [{ ...ev.checkpoints[0], distanceKm: newLapKm }];
          }
        }
        return patch;
      });
      showToast('本活動基本資訊已更新', 'success');
    });

    $('form-admin-credentials').addEventListener('submit', async (e) => {
      e.preventDefault();
      const oldU = getCurrentAdminUsername();
      const isSuper = isCurrentSuperAdmin();
      const newU = isSuper ? ($('cfg-admin-user').value.trim() || oldU) : oldU;
      const newDisplay = $('cfg-admin-displayname').value.trim() || newU;
      const newP = $('cfg-admin-pass').value;
      if (!newP) return showToast('密碼不可為空', 'warn');

      const existingAdmins = [...(state.platform.adminAccounts || [])];
      if (newU !== oldU && existingAdmins.some((a) => a.username === newU)) {
        return showToast(`管理者帳號「${newU}」已存在，請更換其他帳號`, 'error');
      }

      const oldAuth = { ...state.session.auth };
      const nextAdmins = existingAdmins.map((a) =>
        a.username === oldU
          ? { ...a, username: newU, displayName: newDisplay, password: newP, updatedAt: Date.now() }
          : a
      );
      if (!nextAdmins.some((a) => a.username === newU)) {
        nextAdmins.push({ username: newU, password: newP, displayName: newDisplay, isSuperAdmin: isSuper, enabled: true, updatedAt: Date.now() });
      }

      const renamedAdmins = newU !== oldU ? { [oldU]: newU } : {};
      const nextEvents = newU !== oldU
        ? (state.platform.events || []).map((ev) => (getEventOwner(ev) === oldU ? { ...ev, ownerAdmin: newU } : ev))
        : (state.platform.events || []);
      const nextStaff = newU !== oldU
        ? (state.platform.staffAccounts || []).map((s) => (s.ownerAdmin === oldU ? { ...s, ownerAdmin: newU } : s))
        : (state.platform.staffAccounts || []);

      const updated = await DB.savePlatform({
        ...state.platform,
        adminUsername: state.platform.adminUsername === oldU ? newU : state.platform.adminUsername,
        adminPassword: state.platform.adminUsername === oldU ? newP : state.platform.adminPassword,
        adminAccounts: nextAdmins,
        events: nextEvents,
        staffAccounts: nextStaff
      });

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try {
          await Sync.pushPlatform(updated, oldAuth, { targetOwner: newU, renamedAdmins });
        } catch (err) {
          showToast(`本機帳密已更新，雲端同步失敗：${err.message}`, 'warn');
        }
      }
      state.session.auth = { username: newU, password: newP };
      state.session.account.username = newU;
      state.session.account.displayName = newDisplay;
      state.session.account.password = newP;
      DB.saveSession(state.session);
      await reloadState();
      showToast('您的管理者帳密已更新', 'success');
    });

    // 最高管理員：多管理者帳號增刪改
    $('btn-add-admin-account').addEventListener('click', async () => {
      if (!isCurrentSuperAdmin()) return;
      const list = [...(state.platform.adminAccounts || [])];
      let n = list.length + 1;
      while (list.some((a) => a.username === `admin${n}`)) n++;
      list.push({
        username: `admin${n}`,
        password: `pass${n}234`,
        displayName: `活動管理者 ${n}`,
        isSuperAdmin: false,
        enabled: true,
        updatedAt: Date.now()
      });
      await DB.savePlatform({ ...state.platform, adminAccounts: list });
      await reloadState();
    });

    $('admin-accounts-list').addEventListener('click', async (e) => {
      if (!isCurrentSuperAdmin()) return;
      const delBtn = e.target.closest('[data-del-admin]');
      if (!delBtn) return;
      const idx = Number(delBtn.dataset.delAdmin);
      const list = [...(state.platform.adminAccounts || [])];
      const target = list[idx];
      if (!target) return;
      if (target.username === getCurrentAdminUsername()) {
        return showToast('不可刪除目前登入中的管理者帳號', 'error');
      }
      if (!confirm(`確定刪除管理者帳號「${target.username} (${target.displayName})」嗎？\n（該管理者原本建立的活動資料仍會保留）`)) return;
      list.splice(idx, 1);
      if (!list.some((a) => a.isSuperAdmin && a.enabled !== false)) {
        return showToast('系統至少需保留一位啟用中的最高管理員', 'error');
      }
      const updated = await DB.savePlatform({ ...state.platform, adminAccounts: list });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushPlatform(updated, state.session.auth); } catch (err) { showToast(`雲端同步失敗：${err.message}`, 'warn'); }
      }
      await reloadState();
      showToast(`已刪除管理者帳號「${target.username}」`, 'info');
    });

    $('btn-save-admin-accounts').addEventListener('click', async () => {
      if (!isCurrentSuperAdmin()) return;
      const rows = Array.from(document.querySelectorAll('#admin-accounts-list [data-admin-row]'));
      const seen = new Set();
      const nextAdmins = [];
      const renamedAdmins = {};
      const oldSelf = getCurrentAdminUsername();
      let newSelf = oldSelf;
      let newSelfPass = state.session.auth.password;
      let newSelfDisplay = state.session.account.displayName;

      for (const row of rows) {
        const origUser = row.dataset.adminOrigUser || '';
        const username = row.querySelector('[data-admin-field="username"]').value.trim();
        const password = row.querySelector('[data-admin-field="password"]').value;
        const displayName = row.querySelector('[data-admin-field="displayName"]').value.trim() || username;
        const isSuperAdmin = row.querySelector('[data-admin-field="isSuperAdmin"]').value === '1';
        const enabled = row.querySelector('[data-admin-field="enabled"]').checked;

        if (!username) continue;
        if (!password) return showToast(`管理者「${username}」密碼不可為空`, 'error');
        if (seen.has(username)) return showToast(`管理者帳號「${username}」重複`, 'error');
        seen.add(username);

        if (origUser && origUser !== username) {
          renamedAdmins[origUser] = username;
        }
        if (origUser === oldSelf) {
          newSelf = username;
          newSelfPass = password;
          newSelfDisplay = displayName;
        }
        nextAdmins.push({
          username,
          password,
          displayName,
          isSuperAdmin: origUser === oldSelf ? true : isSuperAdmin,
          enabled: origUser === oldSelf ? true : enabled,
          updatedAt: Date.now()
        });
      }

      if (!nextAdmins.some((a) => a.isSuperAdmin && a.enabled !== false)) {
        return showToast('系統至少需保留一位啟用中的最高管理員', 'error');
      }

      const oldAuth = { ...state.session.auth };
      const mapOwner = (o) => (renamedAdmins[o] ? renamedAdmins[o] : o);
      const nextEvents = Object.keys(renamedAdmins).length
        ? (state.platform.events || []).map((ev) => ({ ...ev, ownerAdmin: mapOwner(getEventOwner(ev)) }))
        : (state.platform.events || []);
      const nextStaff = Object.keys(renamedAdmins).length
        ? (state.platform.staffAccounts || []).map((s) => (s.ownerAdmin ? { ...s, ownerAdmin: mapOwner(s.ownerAdmin) } : s))
        : (state.platform.staffAccounts || []);

      const updated = await DB.savePlatform({
        ...state.platform,
        adminAccounts: nextAdmins,
        events: nextEvents,
        staffAccounts: nextStaff
      });

      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try {
          await Sync.pushPlatform(updated, oldAuth, { targetOwner: newSelf, renamedAdmins });
        } catch (err) {
          showToast(`本機已儲存，雲端同步失敗：${err.message}`, 'warn');
        }
      }

      state.session.auth = { username: newSelf, password: newSelfPass };
      state.session.account.username = newSelf;
      state.session.account.password = newSelfPass;
      state.session.account.displayName = newSelfDisplay;
      DB.saveSession(state.session);

      await reloadState();
      showToast(`已儲存 ${nextAdmins.length} 組管理者帳號設定！每位管理者登入後僅可操作自己建立的活動`, 'success');
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
    $('admin-reg-checkin-filter').addEventListener('change', (e) => { state.adminRegCheckinFilter = e.target.value; renderAdminRosterTab(); });
    $('admin-reg-tbody').addEventListener('click', async (e) => {
      const quickChk = e.target.closest('[data-quick-checkin-bib]');
      if (quickChk) {
        const evt = getActiveEvent();
        if (!evt) return;
        const bib = quickChk.dataset.quickCheckinBib;
        const idNumber = quickChk.dataset.quickCheckinIdnum || '';
        const res = await DB.addScan({
          eventId: evt.eventId,
          bib,
          idNumber,
          checkpointId: TE.CHECKIN_CP_ID,
          timestampMs: Date.now(),
          source: 'admin',
          operator: state.session.account.username,
          note: '後台人工補報到',
          bypassDebounce: true
        });
        if (Sync.isOnlineEnabled() && navigator.onLine) {
          try { await Sync.pushAdminEditScan(res.scan, state.session.auth); } catch {}
        }
        await reloadState();
        showToast(`✅ 已完成 #${bib} 人工補報到！`, 'success');
        return;
      }
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
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRegistrations([saved], state.session?.auth);
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
    $('results-search').addEventListener('input', (e) => { state.resultsSearch = e.target.value; renderAdminResultsTab(); });
    $('btn-export-excel-quick').addEventListener('click', () => {
      const evt = getActiveEvent();
      if (!evt || !state.computed) return showToast('請先建立或選取活動', 'warn');
      const f = Excel.exportResultsToExcel({
        event: evt,
        computed: state.computed,
        scans: state.allScans.filter((s) => s.eventId === evt.eventId),
        maskId: false
      });
      showToast(`已匯出「${evt.name}」Excel 報表：${f}`, 'success');
    });

    let activeLapModalBib = '';
    const renderLapDetailModal = (bib) => {
      const evt = getActiveEvent();
      if (!evt || !state.computed) return;
      const row = state.computed.rows.find((r) => r.bib === bib);
      if (!row) return;
      activeLapModalBib = bib;
      $('modal-lap-title').textContent = `🔁 逐圈明細 — ${row.bib} ${row.name} (${row.categoryName})`;
      $('modal-lap-summary').textContent = `有效 ${row.lapCount} 圈 × ${row.lapDistanceKm} km ＝ 總里程 ${row.totalDistanceKm.toFixed(2)} km ｜ 總名次 #${row.overallRank || '-'} ｜ 組名次 #${row.categoryRank || '-'}`;
      $('lap-modal-bib').value = row.bib;
      $('lap-modal-idnumber').value = row.idNumber || '';
      $('lap-manual-time').value = TE.formatClockMs(Date.now());

      const tb = $('modal-lap-tbody');
      if (!row.laps.length) {
        tb.innerHTML = '<tr><td colspan="6" class="py-5 text-center text-slate-500">尚無繞圈感應紀錄，可由上方表單手動補登</td></tr>';
      } else {
        tb.innerHTML = row.laps
          .map((l) => `
            <tr class="${l.overtime ? 'bg-rose-950/30 text-rose-300' : 'hover:bg-slate-800/40'}">
              <td class="py-2 px-3 font-mono-num font-bold">${l.overtime ? '超時' : `第 ${l.lapNumber} 圈`}</td>
              <td class="py-2 px-2 font-mono-num text-sky-300">${TE.formatClockMs(l.timestampMs)}</td>
              <td class="py-2 px-2 font-mono-num">${l.lapDurationMs !== null ? TE.formatDuration(l.lapDurationMs, false) : '--'}</td>
              <td class="py-2 px-2 font-mono-num text-slate-400">${esc(l.lapPace)}</td>
              <td class="py-2 px-2 font-mono-num font-bold text-emerald-400">${l.cumulativeKm !== null ? `${l.cumulativeKm.toFixed(2)} km` : '不計'}</td>
              <td class="py-2 px-3 text-right">
                <button type="button" data-del-lap-scan="${esc(l.scanId)}" class="px-2 py-1 rounded text-[11px] font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>
              </td>
            </tr>
          `)
          .join('');
      }
      $('modal-lap-detail').classList.remove('hidden');
    };

    $('btn-close-lap-modal').addEventListener('click', () => {
      activeLapModalBib = '';
      $('modal-lap-detail').classList.add('hidden');
    });

    $('form-add-lap-manual').addEventListener('submit', async (e) => {
      e.preventDefault();
      const evt = getActiveEvent();
      const bib = $('lap-modal-bib').value;
      const idNumber = $('lap-modal-idnumber').value;
      const tsMs = TE.parseDateTimeToMs(evt.date, $('lap-manual-time').value.trim());
      if (tsMs === null) return showToast('時間格式錯誤，請輸入 HH:mm:ss 或 HH:mm:ss.SSS', 'error');

      const lapCpId = evt.checkpoints?.[0]?.id || 'LAP';
      const res = await DB.addScan({
        eventId: evt.eventId,
        bib,
        idNumber,
        checkpointId: lapCpId,
        timestampMs: tsMs,
        source: 'admin',
        operator: state.session.account.username,
        note: $('lap-manual-note').value.trim() || '後台人工補圈',
        bypassDebounce: true
      });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushAdminEditScan(res.scan, state.session.auth); } catch {}
      }
      await reloadState();
      renderLapDetailModal(bib);
      showToast('已人工補登 1 圈並重新計算總里程與名次！', 'success');
    });

    $('modal-lap-tbody').addEventListener('click', async (e) => {
      const delBtn = e.target.closest('[data-del-lap-scan]');
      if (!delBtn || !confirm('確定刪除此圈感應紀錄嗎？')) return;
      const scanId = delBtn.dataset.delLapScan;
      const existing = state.allScans.find((s) => s.id === scanId);
      if (!existing) return;
      const updated = await DB.putScan({ ...existing, deleted: true, source: 'admin', operator: state.session.account.username, note: '管理員刪除誤刷圈', synced: 0 });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushAdminEditScan(updated, state.session.auth); } catch {}
      }
      await reloadState();
      if (activeLapModalBib) renderLapDetailModal(activeLapModalBib);
      showToast('已刪除此圈紀錄並重算總里程與名次', 'info');
    });

    $('results-tbody').addEventListener('click', (e) => {
      const lapBtn = e.target.closest('[data-open-lap-bib]');
      if (lapBtn) { renderLapDetailModal(lapBtn.dataset.openLapBib); return; }
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
        if (reg && evt && reg.eventId === evt.eventId) {
          const updated = await DB.upsertRegistration({ ...reg, penaltySec: Number($('edit-penalty-sec').value || 0) }, false);
          if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushRegistrations([updated], state.session?.auth);
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

    // 分頁 4：本活動統一律定起跑時間、分組與站點
    $('btn-unified-gun-now').addEventListener('click', async () => {
      const evt = getActiveEvent();
      if (!evt) return showToast('請先建立或選取活動', 'warn');
      const nowSec = Math.floor(Date.now() / 1000) * 1000;
      const iso = new Date(nowSec).toISOString();
      NFC.sound.beepGun();
      await updateCurrentEvent((ev) => ({
        unifiedGunTimeMs: nowSec,
        categories: (ev.categories || []).map((c) => ({ ...c, gunTimeMs: nowSec, gunTimeIso: iso }))
      }));
      showToast(`🔫 「${evt.name}」全組已統一鳴槍起跑！(${TE.formatClockSec(nowSec)})`, 'success');
    });

    $('btn-unified-gun-set').addEventListener('click', async () => {
      const evt = getActiveEvent();
      if (!evt) return showToast('請先建立或選取活動', 'warn');
      const tsMs = TE.parseDateTimeToMs(evt.date, $('unified-gun-input').value.trim());
      if (tsMs === null) return showToast('請輸入 HH:mm:ss 格式（如 06:30:00）', 'error');
      const iso = new Date(tsMs).toISOString();
      await updateCurrentEvent((ev) => ({
        unifiedGunTimeMs: tsMs,
        categories: (ev.categories || []).map((c) => ({ ...c, gunTimeMs: tsMs, gunTimeIso: iso }))
      }));
      showToast(`✅ 已套用統一律定起跑時間 ${TE.formatClockSec(tsMs)} 至全活動所有分組！`, 'success');
    });

    $('btn-unified-gun-clear').addEventListener('click', async () => {
      const evt = getActiveEvent();
      if (!evt) return;
      if (!confirm('確定清除本活動所有分組的起跑時間嗎？')) return;
      await updateCurrentEvent((ev) => ({
        unifiedGunTimeMs: null,
        categories: (ev.categories || []).map((c) => ({ ...c, gunTimeMs: null, gunTimeIso: '' }))
      }));
      showToast('已清除本活動統一律定起跑時間', 'info');
    });

    $('btn-add-category').addEventListener('click', async () => {
      await updateCurrentEvent((ev) => {
        const n = (ev.categories || []).length + 1;
        const baseGun = ev.unifiedGunTimeMs || null;
        const baseIso = baseGun ? new Date(baseGun).toISOString() : '';
        if (ev.eventType === 'lap_timed') {
          return { categories: [...(ev.categories || []), { id: `H${n}`, name: `自訂 ${n * 3} 小時組`, durationHours: n * 3, distanceKm: n * 3, gunTimeMs: baseGun, gunTimeIso: baseIso, color: '#a855f7' }] };
        }
        return { categories: [...(ev.categories || []), { id: `CAT${n}`, name: `新分組 ${n}`, distanceKm: 5, gunTimeMs: baseGun, gunTimeIso: baseIso, color: '#a855f7' }] };
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
        const numVal = Number(card.querySelector('[data-cat-field="distanceKm"]').value || 0);
        cats[idx] = {
          ...cats[idx],
          id: card.querySelector('[data-cat-field="id"]').value.trim(),
          name: card.querySelector('[data-cat-field="name"]').value.trim(),
          distanceKm: numVal,
          ...(evt.eventType === 'lap_timed' ? { durationHours: numVal } : {})
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
        const sorted = [...(ev.checkpoints || [])]
          .filter((cp) => cp && cp.type !== 'start' && String(cp.id || '').toUpperCase() !== 'START')
          .sort((a, b) => a.order - b.order);
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

    // 分頁 5：作業員帳號與鎖定活動/站點管制（僅管理本人名下作業員，不影響其他管理者帳號）
    $('btn-add-staff').addEventListener('click', async () => {
      const adminUser = getCurrentAdminUsername();
      const allStaff = [...(state.platform.staffAccounts || [])];
      const existingUsers = new Set(allStaff.map((s) => s.username));
      const prefix = isCurrentSuperAdmin() ? 'staff' : `${adminUser}_staff`;
      let n = 1;
      while (existingUsers.has(`${prefix}${n}`)) n++;
      const evt = getActiveEvent();
      allStaff.push({
        username: `${prefix}${n}`,
        password: '1111',
        displayName: `作業員 ${n}`,
        role: 'scanner',
        eventId: evt ? evt.eventId : 'ALL',
        checkpointId: evt?.eventType === 'lap_timed' ? 'LAP' : 'CHECKIN',
        ownerAdmin: adminUser,
        enabled: true,
        updatedAt: Date.now()
      });
      await DB.savePlatform({ ...state.platform, staffAccounts: allStaff });
      await reloadState();
    });

    $('staff-accounts-list').addEventListener('click', async (e) => {
      const delBtn = e.target.closest('[data-del-staff]');
      if (!delBtn) return;
      const targetUser = delBtn.dataset.delStaff;
      const adminUser = getCurrentAdminUsername();
      const isSuper = isCurrentSuperAdmin();
      const allStaff = [...(state.platform.staffAccounts || [])];
      const idx = allStaff.findIndex((s) => s.username === targetUser);
      if (idx < 0) return;
      const sOwner = String(allStaff[idx].ownerAdmin || '').trim();
      const canManage = isSuper ? (!sOwner || sOwner === adminUser) : sOwner === adminUser;
      if (!canManage) return showToast('無權限刪除此作業員帳號', 'error');
      if (!confirm(`確定刪除作業員帳號「${targetUser}」？`)) return;
      allStaff.splice(idx, 1);
      const updated = await DB.savePlatform({ ...state.platform, staffAccounts: allStaff });
      if (Sync.isOnlineEnabled() && navigator.onLine) Sync.pushPlatform(updated, state.session.auth);
      await reloadState();
    });

    $('btn-save-staff-accounts').addEventListener('click', async () => {
      const adminUser = getCurrentAdminUsername();
      const isSuper = isCurrentSuperAdmin();
      const otherStaff = (state.platform.staffAccounts || []).filter((s) => {
        const sOwner = String(s.ownerAdmin || '').trim();
        return isSuper ? (sOwner && sOwner !== adminUser) : sOwner !== adminUser;
      });
      const otherUserSet = new Set(otherStaff.map((s) => s.username));

      const rows = Array.from(document.querySelectorAll('#staff-accounts-list [data-staff-row]'));
      const seen = new Set();
      const ownStaff = [];
      for (const row of rows) {
        const username = row.querySelector('[data-staff-field="username"]').value.trim();
        if (!username) continue;
        if (seen.has(username)) return showToast(`帳號「${username}」重複`, 'error');
        if (otherUserSet.has(username)) return showToast(`帳號「${username}」已被其他管理者使用，請更換其他名稱`, 'error');
        seen.add(username);
        const role = row.querySelector('[data-staff-field="role"]').value;
        const cpId = row.querySelector('[data-staff-field="checkpointId"]').value.trim();
        if (role !== 'writer' && !cpId) return showToast(`帳號「${username}」屬於感應角色，必須填寫鎖定站點代碼`, 'error');
        const origOwner = String(row.dataset.staffOwner || '').trim();
        ownStaff.push({
          username,
          password: row.querySelector('[data-staff-field="password"]').value || '1111',
          displayName: row.querySelector('[data-staff-field="displayName"]').value.trim() || username,
          role,
          eventId: row.querySelector('[data-staff-field="eventId"]').value || 'ALL',
          checkpointId: role === 'writer' ? '' : cpId,
          ownerAdmin: isSuper ? (origOwner === '' ? '' : adminUser) : adminUser,
          enabled: row.querySelector('[data-staff-field="enabled"]').checked,
          updatedAt: Date.now()
        });
      }
      const mergedStaff = [...otherStaff, ...ownStaff];
      const updated = await DB.savePlatform({ ...state.platform, staffAccounts: mergedStaff });
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        try { await Sync.pushPlatform(updated, state.session.auth); } catch (err) { showToast(`本機已儲存，雲端同步失敗：${err.message}`, 'warn'); }
      }
      await reloadState();
      showToast(`已儲存 ${ownStaff.length} 組作業員帳號設定！`, 'success');
    });

    // 分頁 6：備份、Excel 匯出與雲端同步
    $('btn-export-excel-full').addEventListener('click', () => {
      const evt = getActiveEvent();
      if (!evt || !state.computed) return showToast('請先建立或選取活動', 'warn');
      const f = Excel.exportResultsToExcel({ event: evt, computed: state.computed, scans: state.allScans.filter((s) => s.eventId === evt.eventId), maskId: false });
      showToast(`已匯出完整 Excel：${f}`, 'success');
    });
    $('btn-export-excel-masked').addEventListener('click', () => {
      const evt = getActiveEvent();
      if (!evt || !state.computed) return showToast('請先建立或選取活動', 'warn');
      const f = Excel.exportResultsToExcel({ event: evt, computed: state.computed, scans: state.allScans.filter((s) => s.eventId === evt.eventId), maskId: true });
      showToast(`已匯出隱碼公開 Excel：${f}`, 'success');
    });

    $('btn-backup-download').addEventListener('click', () => downloadJsonBackup(`智慧路跑備份_${getCurrentAdminUsername() || 'all'}`));
    $('input-backup-merge').addEventListener('change', async (e) => {
      const f = e.target.files?.[0];
      if (f) {
        try {
          const r = await DB.importFullBackup(JSON.parse(await f.text()), 'merge', {
            ownerAdmin: getCurrentAdminUsername(),
            isSuperAdmin: isCurrentSuperAdmin()
          });
          await reloadState();
          showToast(`合併完成！晶片 ${r.chipCount}、報名 ${r.regCount}、感應 ${r.scanCount} 筆`, 'success');
        } catch (err) { showToast(err.message, 'error'); }
      }
      e.target.value = '';
    });
    $('input-backup-replace').addEventListener('change', async (e) => {
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理員可執行覆蓋還原', 'error');
      const f = e.target.files?.[0];
      if (f && confirm('確定以備份檔覆蓋還原整個本機資料庫嗎？')) {
        try {
          const r = await DB.importFullBackup(JSON.parse(await f.text()), 'replace', {
            ownerAdmin: getCurrentAdminUsername(),
            isSuperAdmin: true
          });
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
      DB.saveDeviceSettings({ gasUrl: url, mode: url ? 'online' : 'offline' });
      renderHeader();
      updateNetworkIndicators();
      if (!url) return showToast('已清除 GAS 網址並切回離線模式', 'info');
      try {
        const res = await Sync.ping(url);
        showToast(`✅ 連線成功！${res.message || ''} 正在自動同步資料...`, 'success');
        await handleUploadPendingSync(false);
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
    $('btn-copy-config-js').addEventListener('click', async () => {
      const gasUrl = DB.getDeviceSettings().gasUrl;
      if (!gasUrl) return showToast('請先於上方輸入並儲存 Google Apps Script 網址', 'warn');
      const code = `window.APP_CONFIG = {\n  DEFAULT_GAS_URL: '${gasUrl}',\n  AUTO_SYNC_ON_START: true\n};\n`;
      try {
        await navigator.clipboard.writeText(code);
        showToast('已複製 js/config.js 設定碼！貼入 GitHub 的 js/config.js 後，全部手機開啟即自動同步！', 'success');
      } catch {
        prompt('請複製以下內容並貼入 js/config.js：', code);
      }
    });

    $('btn-load-demo-data').addEventListener('click', async () => {
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理員可載入模擬資料', 'error');
      if (!confirm('載入模擬資料將會覆蓋目前本機資料（包含距離賽與 2.54km 繞圈計時賽），確定繼續嗎？')) return;
      await DB.loadDemoData(getCurrentAdminUsername());
      state.selectedEventId = 'EVT-2026-ULTRA';
      await reloadState();
      showToast('⚡ 已載入 3 場模擬活動（含 6H/12H/24H 環湖繞圈計時賽）！', 'success');
    });
    $('btn-clear-all-data').addEventListener('click', async () => {
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理員可清空本機資料庫', 'error');
      if (!confirm('確定清空本機所有活動、晶片與感應資料嗎？')) return;
      await DB.clearAllData(true);
      await reloadState();
      showToast('本機資料庫已清空重設', 'info');
    });
  }

  document.addEventListener('DOMContentLoaded', initApp);
})();
