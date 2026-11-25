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
  const Cert = window.CertificateGenerator;

  const state = {
    view: 'portal',
    session: null,      // { mode: 'staff'|'admin', account, auth, lockedEventId }
    selectedEventId: '',
    staffSubTab: 'scan',
    adminZone: null,
    adminTab: 'events',
    adminLastOpsTab: 'events',
    adminLastClosedTab: 'export',
    platform: null,
    chips: [],
    allRegistrations: [],
    allScans: [],
    historyRecords: [],
    computed: null,
    resultsCatFilter: 'ALL',
    resultsSearch: '',
    rosterSearch: '',
    writerEventFilter: 'ALL',
    adminRegSearch: '',
    adminRegCheckinFilter: 'ALL',
    adminHistoryEventFilter: 'ALL',
    adminHistorySearch: '',
    exportSelectedEventId: '',
    staffLoginSelectedGroup: '',
    adminLoginRole: 'super',
    adminLoginSelectedUser: '',
    historyPreview: {
      items: [],
      skippedCount: 0,
      sourceLabel: '',
      fromText: false,
      eventFilter: 'ALL',
      search: ''
    },
    editMode: 'scan',
    cert: {
      mode: 'official',
      selectedSourceKey: '',
      loadedEventId: '',
      catFilter: 'ALL',
      search: '',
      onlyFinished: true,
      selectedBib: '',
      orientation: 'landscape',
      format: 'classic_gold',
      theme: 'ivory_gold',
      englishName: '',
      organization: 'Taiwan Ultra Runners',
      showSplits: true,
      bgDataUrl: '',
      bgImage: null,
      showTurLogo: false,
      layout: Cert ? Cert.getDefaultCustomLayout() : [],
      selectedFieldKey: 'name',
      hitBoxes: []
    }
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

  function isEventFinished(ev) {
    return Boolean(ev && ev.status === 'finished');
  }

  function getOpenEvents() {
    return (state.platform?.events || []).filter((e) => !isEventFinished(e));
  }

  function getAdminOwnEvents() {
    const adminUser = getCurrentAdminUsername();
    if (!adminUser) return [];
    return (state.platform?.events || []).filter((e) => getEventOwner(e) === adminUser);
  }

  function getAdminOpenOwnEvents() {
    return getAdminOwnEvents().filter((e) => !isEventFinished(e));
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
      const openOwnEvents = getAdminOpenOwnEvents();
      return openOwnEvents.find((e) => e.eventId === state.selectedEventId) || openOwnEvents[0] || null;
    }
    const openEvents = getOpenEvents();
    const targetId =
      state.view === 'staff-workspace' && state.session?.lockedEventId
        ? state.session.lockedEventId
        : state.selectedEventId;
    if (state.view === 'staff-workspace' && state.session?.lockedEventId) {
      return openEvents.find((e) => e.eventId === targetId) || null;
    }
    return openEvents.find((e) => e.eventId === targetId) || openEvents[0] || null;
  }

  /* ========================================================================
   * 全系統字體大小控制與密碼防暫存保護
   * ======================================================================== */
  const FONT_SCALE_KEY = 'tur_font_scale_v1';
  const FONT_SCALE_OPTIONS = ['sm', 'md', 'lg', 'xl'];

  function getSavedFontScale() {
    try {
      const saved = localStorage.getItem(FONT_SCALE_KEY) || 'md';
      return FONT_SCALE_OPTIONS.includes(saved) ? saved : 'md';
    } catch {
      return 'md';
    }
  }

  function applyFontScale(scale, { save = true, toast = false } = {}) {
    const valid = FONT_SCALE_OPTIONS.includes(scale) ? scale : 'md';
    document.documentElement.setAttribute('data-font-scale', valid);
    if (save) {
      try { localStorage.setItem(FONT_SCALE_KEY, valid); } catch {}
    }
    document.querySelectorAll('.font-scale-btn').forEach((btn) => {
      const active = btn.dataset.fontScale === valid;
      btn.className = active
        ? 'font-scale-btn px-2.5 sm:px-3 py-1 rounded-lg text-xs font-black transition bg-sky-500 text-slate-950 shadow'
        : 'font-scale-btn px-2.5 sm:px-3 py-1 rounded-lg text-xs font-black transition text-slate-300 hover:bg-slate-800';
      btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
    if (toast) {
      const labels = { sm: '小 (100%)', md: '標準 (112%)', lg: '大 (125%)', xl: '特大 (138%)' };
      showToast(`🔤 已切換系統字體為「${labels[valid]}」`, 'info');
    }
  }

  function bindNoAutofillUnlock(el) {
    if (!el || el.dataset.noAutofillBound === '1') return;
    el.dataset.noAutofillBound = '1';
    const unlock = () => {
      if (el.hasAttribute('readonly')) el.removeAttribute('readonly');
    };
    el.addEventListener('pointerdown', unlock);
    el.addEventListener('touchstart', unlock, { passive: true });
    el.addEventListener('focus', unlock);
  }

  function initNoCachePasswordInputs(root = document) {
    const supportsTextSecurity =
      typeof CSS !== 'undefined' &&
      typeof CSS.supports === 'function' &&
      (CSS.supports('-webkit-text-security', 'disc') || CSS.supports('text-security', 'disc'));
    const randTag = Math.random().toString(36).slice(2, 8);

    root.querySelectorAll('.tur-secret-input').forEach((el, idx) => {
      if (!supportsTextSecurity) {
        el.type = 'password';
        el.setAttribute('autocomplete', 'new-password');
      } else {
        el.type = 'text';
        el.setAttribute('autocomplete', 'new-password');
      }
      el.setAttribute('name', `tur_sec_${randTag}_${idx}`);
      el.setAttribute('data-lpignore', 'true');
      el.setAttribute('data-1p-ignore', 'true');
      el.setAttribute('data-bwignore', 'true');
      el.setAttribute('readonly', 'readonly');
      el.value = '';
      bindNoAutofillUnlock(el);
    });

    root.querySelectorAll('.tur-no-autofill').forEach((el, idx) => {
      el.setAttribute('name', `tur_usr_${randTag}_${idx}`);
      el.setAttribute('autocomplete', 'off');
      el.setAttribute('data-lpignore', 'true');
      el.setAttribute('data-1p-ignore', 'true');
      el.setAttribute('data-bwignore', 'true');
      el.setAttribute('readonly', 'readonly');
      el.value = '';
      bindNoAutofillUnlock(el);
    });
  }

  function clearLoginInputs() {
    const ids = ['staff-login-user', 'staff-login-pass', 'admin-login-user', 'admin-login-pass'];
    const wipe = () => {
      ids.forEach((id) => {
        const el = $(id);
        if (!el) return;
        el.value = '';
        if (el.classList.contains('tur-secret-input') || el.classList.contains('tur-no-autofill')) {
          el.setAttribute('readonly', 'readonly');
        }
      });
    };
    wipe();
    // 防止瀏覽器延遲自動填入（Chrome / Edge / Safari 密碼管理員）
    setTimeout(wipe, 120);
  }

  /* ========================================================================
   * 初始化與資料重新載入
   * ======================================================================== */
  async function initApp() {
    applyFontScale(getSavedFontScale(), { save: false });
    initNoCachePasswordInputs();
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
        state.adminZone = 'ops';
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

    // 採「本機快取秒開＋背景非同步雲端同步」：不阻斷進入畫面，且尊重「單機離線」模式設定
    if (navigator.onLine && Sync.isOnlineEnabled()) {
      autoSyncCloud({ reason: 'boot', showNotice: false });
    }

    // 網路恢復時若處於「雲端連線模式」，自動背景雙向同步
    window.addEventListener('online', () => {
      updateNetworkIndicators();
      if (Sync.isOnlineEnabled()) {
        showToast('網路已恢復，正在背景雙向同步雲端資料...', 'info');
        autoSyncCloud({ reason: 'reconnect', showNotice: true });
      }
    });

    // 手機喚醒或切回瀏覽器分頁時，若為雲端連線模式則自動背景更新（節流 20 秒）
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && navigator.onLine && Sync.isOnlineEnabled()) {
        const last = DB.getDeviceSettings().lastSyncAt || 0;
        if (Date.now() - last > 20000) {
          autoSyncCloud({ reason: 'resume', showNotice: false });
        }
      }
    });

    window.addEventListener('offline', () => {
      updateNetworkIndicators();
      showToast('已進入無網路狀態！晶片計時作業仍可離線執行，有網後再統一上傳', 'warn');
    });

    // 防止瀏覽器上下頁快取 (bfcache) 還原先前輸入的密碼
    window.addEventListener('pageshow', () => {
      if (!state.session) clearLoginInputs();
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
    const [platform, chips, regs, scans, historyRecords] = await Promise.all([
      DB.getPlatform(),
      DB.getAllChips(),
      DB.getAllRegistrations(),
      DB.getAllScans(),
      DB.getAllHistoryRecords()
    ]);

    state.platform = platform;
    state.chips = chips;
    state.allRegistrations = regs;
    state.allScans = scans;
    state.historyRecords = historyRecords;

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
      const openOwnEvents = getAdminOpenOwnEvents();
      if (!state.selectedEventId || !openOwnEvents.some((e) => e.eventId === state.selectedEventId)) {
        const defaultId = platform.activeEventId;
        const validDefault = defaultId && openOwnEvents.some((e) => e.eventId === defaultId);
        state.selectedEventId = validDefault ? defaultId : (openOwnEvents[0]?.eventId || '');
      }
    } else {
      const openEvents = getOpenEvents();
      if (!state.selectedEventId || !openEvents.some((e) => e.eventId === state.selectedEventId)) {
        const defaultId = platform.activeEventId;
        const validDefault = defaultId && openEvents.some((e) => e.eventId === defaultId);
        state.selectedEventId = validDefault ? defaultId : (openEvents[0]?.eventId || '');
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
    if (viewName === 'staff-login') {
      state.staffLoginSelectedGroup = '';
      clearLoginInputs();
      renderStaffLoginGroups();
    }
    if (viewName === 'admin-login') {
      state.adminLoginSelectedUser = '';
      clearLoginInputs();
      renderAdminLoginView();
    }
    if (viewName === 'staff-workspace') renderStaffWorkspace();
    if (viewName === 'admin-workspace') {
      renderAdminWorkspace();
      if (state.adminZone) {
        switchAdminTab(state.adminTab || 'events');
      } else {
        switchAdminZone(null);
      }
    }
    window.scrollTo({ top: 0, behavior: 'instant' });
  }

  function updateAdminZoneNavStyles(activeZone) {
    const styleMap = {
      ops: {
        active: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border-2 border-amber-400 bg-amber-500/20 shadow-lg shadow-amber-950/40 transition flex flex-col justify-between gap-1.5 min-w-0',
        idle: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border border-amber-500/30 bg-slate-950/80 hover:bg-amber-950/35 hover:border-amber-400/70 transition flex flex-col justify-between gap-1.5 min-w-0',
        labelCls: 'text-amber-300'
      },
      closed: {
        active: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border-2 border-emerald-400 bg-emerald-500/20 shadow-lg shadow-emerald-950/40 transition flex flex-col justify-between gap-1.5 min-w-0',
        idle: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border border-emerald-500/30 bg-slate-950/80 hover:bg-emerald-950/35 hover:border-emerald-400/70 transition flex flex-col justify-between gap-1.5 min-w-0',
        labelCls: 'text-emerald-300'
      },
      history: {
        active: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border-2 border-purple-400 bg-purple-500/20 shadow-lg shadow-purple-950/40 transition flex flex-col justify-between gap-1.5 min-w-0',
        idle: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border border-purple-500/30 bg-slate-950/80 hover:bg-purple-950/35 hover:border-purple-400/70 transition flex flex-col justify-between gap-1.5 min-w-0',
        labelCls: 'text-purple-300'
      },
      db: {
        active: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border-2 border-sky-400 bg-sky-500/20 shadow-lg shadow-sky-950/40 transition flex flex-col justify-between gap-1.5 min-w-0',
        idle: 'admin-zone-nav-btn text-left rounded-xl px-3 py-2.5 border border-sky-500/30 bg-slate-950/80 hover:bg-sky-950/35 hover:border-sky-400/70 transition flex flex-col justify-between gap-1.5 min-w-0',
        labelCls: 'text-sky-300'
      }
    };

    const isSuper = isCurrentSuperAdmin();
    const navGrid = $('admin-zone-nav-grid');
    if (navGrid) {
      navGrid.className = isSuper
        ? 'grid grid-cols-2 sm:grid-cols-4 gap-2'
        : 'grid grid-cols-1 sm:grid-cols-3 gap-2';
    }
    const dbNumEl = $('admin-zone-db-num');
    if (dbNumEl) dbNumEl.textContent = isSuper ? '專區 4' : '專區 3';

    document.querySelectorAll('.admin-zone-nav-btn').forEach((btn) => {
      const z = btn.dataset.adminZone;
      const cfg = styleMap[z];
      if (!cfg) return;
      const isAct = z === activeZone;
      const hideForNonSuper = z === 'history' && !isSuper;
      btn.className = (isAct ? cfg.active : cfg.idle) + (hideForNonSuper ? ' hidden' : '');
      btn.setAttribute('aria-expanded', isAct ? 'true' : 'false');
      const chev = btn.querySelector('.admin-zone-chevron');
      if (chev) {
        chev.textContent = isAct ? '▲' : '▼';
        chev.className = `admin-zone-chevron text-[11px] font-black ${isAct ? cfg.labelCls : 'text-slate-400'}`;
      }
      const st = btn.querySelector('.admin-zone-state-label');
      if (st) {
        st.textContent = isAct ? '作業中' : '點選開啟';
        st.className = `admin-zone-state-label text-[10px] font-black ${isAct ? cfg.labelCls : 'text-slate-500'}`;
      }
    });
  }

  function switchAdminZone(zoneKey, { toggle = false } = {}) {
    if (zoneKey === 'history' && !isCurrentSuperAdmin()) {
      showToast('僅最高管理權責可進入歷年活動成績匯入', 'error');
      switchAdminZone(null);
      return;
    }
    if (toggle && zoneKey && state.adminZone === zoneKey) {
      zoneKey = null;
    }

    if (!zoneKey) {
      state.adminZone = null;
      document.querySelectorAll('.admin-zone-card').forEach((c) => c.classList.add('hidden'));
      document.querySelectorAll('.admin-panel').forEach((p) => p.classList.add('hidden'));
      if ($('admin-zone-placeholder')) $('admin-zone-placeholder').classList.remove('hidden');
      updateAdminZoneNavStyles(null);
      return;
    }

    const defaultTabMap = {
      ops: state.adminLastOpsTab || 'events',
      closed: state.adminLastClosedTab || 'export',
      history: 'history',
      db: 'sync'
    };
    switchAdminTab(defaultTabMap[zoneKey] || 'events');
  }

  function switchAdminTab(tabName) {
    if (tabName === 'history' && !isCurrentSuperAdmin()) {
      showToast('僅最高管理權責可進入歷年活動成績匯入', 'error');
      switchAdminZone(null);
      return;
    }
    state.adminTab = tabName;

    // 依所選分頁鎖定唯一顯示之專區（橫式選單切換，每次僅展開單一專區，簡化畫面）
    const activeZone = ['events', 'roster', 'results', 'race', 'staff'].includes(tabName)
      ? 'ops'
      : (tabName === 'export' || tabName === 'certificate')
        ? 'closed'
        : tabName === 'history'
          ? 'history'
          : 'db';

    state.adminZone = activeZone;
    if (activeZone === 'ops') state.adminLastOpsTab = tabName;
    if (activeZone === 'closed') state.adminLastClosedTab = tabName;

    if ($('admin-zone-placeholder')) $('admin-zone-placeholder').classList.add('hidden');

    const zoneCards = [
      ['ops', 'admin-zone-ops', 'admin-zone-ops-body'],
      ['closed', 'admin-zone-closed', 'admin-zone-closed-body'],
      ['history', 'admin-zone-history', 'admin-zone-history-body'],
      ['db', 'admin-zone-db', 'admin-zone-db-body']
    ];
    zoneCards.forEach(([key, cardId, bodyId]) => {
      const isAct = key === activeZone;
      const cardEl = $(cardId);
      if (cardEl) cardEl.classList.toggle('hidden', !isAct);
      const bodyEl = $(bodyId);
      if (bodyEl) bodyEl.classList.toggle('hidden', !isAct);
    });

    document.querySelectorAll('.admin-panel').forEach((p) => p.classList.add('hidden'));
    const target = $(`admin-tab-${tabName}`);
    if (target) target.classList.remove('hidden');

    updateAdminZoneNavStyles(activeZone);

    document.querySelectorAll('.admin-tab-btn').forEach((btn) => {
      const active = btn.dataset.adminTab === tabName;
      const group = btn.dataset.tabGroup || 'ops';
      if (group === 'closed') {
        btn.className = active
          ? 'admin-tab-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-black bg-emerald-500 text-slate-950 shadow'
          : 'admin-tab-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-bold text-emerald-200 hover:bg-emerald-900/50 border border-emerald-500/30';
      } else {
        btn.className = active
          ? 'admin-tab-btn px-3 py-2 rounded-xl text-xs sm:text-sm font-black bg-amber-500 text-slate-950 shadow'
          : 'admin-tab-btn px-3 py-2 rounded-xl text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800';
      }
    });

    if (tabName === 'export') {
      renderAdminExportTab();
    } else if (tabName === 'certificate') {
      renderAdminCertificateTab();
    } else if (tabName === 'history') {
      renderAdminHistoryTab();
    } else if (tabName === 'sync') {
      renderAdminSyncTab();
    }
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
      finalMsg = '❌ NFC 權限遭拒，請於瀏覽器權限允許 NFC';
    } else if (diag.supported) {
      finalMsg = '✅ Web NFC 就緒';
    } else {
      finalMsg = '⚠️ 此裝置不支援 Web NFC（可使用模擬或手動補登）';
    }
    if (scanHint) {
      scanHint.textContent = diag.supported
        ? (perm === 'denied' ? '❌ NFC 權限遭拒' : '✅ Web NFC 就緒')
        : '⚠️ 無 Web NFC（可用手動／模擬）';
    }
    if (writerText) writerText.textContent = finalMsg;
    if (writerBanner && (!diag.supported || perm === 'denied')) {
      writerBanner.className = 'rounded-xl border border-amber-500/40 bg-amber-950/30 px-4 py-2.5 text-xs text-amber-200 flex items-center gap-2';
    }
  }

  /* ========================================================================
   * 頂部狀態、活動選單與網路燈號
   * ======================================================================== */
  function renderHeader() {
    const settings = DB.getDeviceSettings();
    const evt = getActiveEvent();
    const totalEvents = (state.platform?.events || []).length;
    $('header-event-badge').textContent = evt
      ? `${evt.name} (${evt.date || ''})`
      : (totalEvents > 0 ? '無進行中活動' : '尚無活動');
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
    const portalAdminTag = $('portal-admin-session-tag');
    const portalStaffTag = $('portal-staff-session-tag');
    if (state.session) {
      userBox.classList.remove('hidden');
      const acc = state.session.account;
      userLabel.textContent = state.session.mode === 'admin'
        ? (acc.isSuperAdmin
          ? `🛡️ ${acc.displayName || acc.username} (最高管理)`
          : `👤 ${acc.displayName || acc.username} (活動承辦)`)
        : `👷 ${acc.displayName || acc.username}`;
      if (portalAdminTag) {
        const isAdm = state.session.mode === 'admin';
        portalAdminTag.classList.toggle('hidden', !isAdm);
        if (isAdm) {
          portalAdminTag.textContent = `✓ ${acc.username}・直接進入`;
        }
      }
      if (portalStaffTag) {
        const isStf = state.session.mode === 'staff';
        portalStaffTag.classList.toggle('hidden', !isStf);
        if (isStf) {
          portalStaffTag.textContent = `✓ ${acc.username}・直接進入`;
        }
      }
    } else {
      userBox.classList.add('hidden');
      if (portalAdminTag) portalAdminTag.classList.add('hidden');
      if (portalStaffTag) portalStaffTag.classList.add('hidden');
    }

    const soundBtn = $('btn-toggle-sound');
    const contrastBtn = $('btn-toggle-contrast');
    if (soundBtn) soundBtn.textContent = NFC.sound.enabled ? '🔊 音效' : '🔇 靜音';
    if (contrastBtn) contrastBtn.textContent = settings.highContrast ? '🌙 一般' : '☀️ 高對比';
  }

  function renderEventDropdowns() {
    // 狀態為「已結束 (finished)」的活動一律不出現在下拉選單，避免資料遭誤選或修改
    const openEvents = getOpenEvents();
    const allCount = (state.platform?.events || []).length;
    const statusLabel = { active: '進行中', upcoming: '籌備中', finished: '已結束' };
    const buildOptionsHtml = (list) => (
      list.map((ev) => {
        const typeTag = ev.eventType === 'lap_timed' ? `🔁繞圈(${ev.lapDistanceKm}K/圈)` : '🏃距離賽';
        return `<option value="${esc(ev.eventId)}" style="background-color:#0f172a;color:#f8fafc;">${typeTag} ${esc(ev.name)} (${esc(ev.date || '')}) [${statusLabel[ev.status] || ev.status}]</option>`;
      }).join('')
    );

    const openOptionsHtml = buildOptionsHtml(openEvents);
    const hasOpen = openEvents.length > 0;

    // 大廳選取活動下拉
    const portalSel = $('portal-event-select');
    const portalEmpty = $('portal-event-select-empty');
    if (portalSel) {
      portalSel.innerHTML = openOptionsHtml;
      portalSel.classList.toggle('hidden', !hasOpen);
      if (hasOpen && openEvents.some((e) => e.eventId === state.selectedEventId)) {
        portalSel.value = state.selectedEventId;
      }
    }
    if (portalEmpty) {
      portalEmpty.classList.toggle('hidden', hasOpen);
      portalEmpty.textContent = allCount > 0 ? '🔒 目前無開放活動' : '尚無活動';
    }

    // 作業員登入選取活動下拉
    const staffSel = $('staff-login-event');
    const staffEmpty = $('staff-login-event-empty');
    if (staffSel) {
      staffSel.innerHTML = openOptionsHtml;
      staffSel.classList.toggle('hidden', !hasOpen);
      if (hasOpen && openEvents.some((e) => e.eventId === state.selectedEventId)) {
        staffSel.value = state.selectedEventId;
      }
    }
    if (staffEmpty) {
      staffEmpty.classList.toggle('hidden', hasOpen);
      staffEmpty.textContent = allCount > 0 ? '🔒 目前無開放活動' : '尚無活動';
    }

    // 後台管理區下拉選單：僅顯示目前登入管理者本人建立且「非已結束」的活動
    const adminSel = $('admin-event-select');
    const adminEmpty = $('admin-event-select-empty');
    if (adminSel) {
      const ownTotal = getAdminOwnEvents().length;
      const openOwnEvents = getAdminOpenOwnEvents();
      const hasOpenOwn = openOwnEvents.length > 0;
      adminSel.innerHTML = buildOptionsHtml(openOwnEvents);
      adminSel.classList.toggle('hidden', !hasOpenOwn);
      if (hasOpenOwn && openOwnEvents.some((e) => e.eventId === state.selectedEventId)) {
        adminSel.value = state.selectedEventId;
      }
      if (adminEmpty) {
        adminEmpty.classList.toggle('hidden', hasOpenOwn);
        adminEmpty.textContent = ownTotal > 0 ? '🔒 無開放活動' : '尚無活動，請按「＋ 新增活動」';
      }
    }

    // 共用晶片寫入區活動篩選下拉：不顯示已結束活動
    const wf = $('writer-event-filter');
    if (wf) {
      const prev = state.writerEventFilter;
      wf.innerHTML = '<option value="ALL">全部活動</option>' + openOptionsHtml;
      wf.classList.toggle('hidden', !hasOpen);
      const nextVal = (prev === 'ALL' || openEvents.some((e) => e.eventId === prev)) ? prev : 'ALL';
      state.writerEventFilter = nextVal;
      wf.value = nextVal;
    }

    renderStaffLoginGroups();
    if (state.view === 'admin-login') renderAdminLoginView();
  }

  /**
   * 渲染後台管理登入畫面：極簡雙身分切換（最高管理權責 ／ 活動承辦人）
   */
  function renderAdminLoginView() {
    const role = state.adminLoginRole === 'coordinator' ? 'coordinator' : 'super';
    const isSuperRole = role === 'super';
    const allAdmins = (state.platform?.adminAccounts || []).filter((a) => a && a.enabled !== false);
    const list = allAdmins.filter((a) => (isSuperRole ? Boolean(a.isSuperAdmin) : !a.isSuperAdmin));

    // 1. 雙身分頁籤樣式切換
    document.querySelectorAll('.admin-login-role-btn').forEach((btn) => {
      const btnRole = btn.dataset.adminLoginRole;
      const isAct = btnRole === role;
      if (btnRole === 'super') {
        btn.className = isAct
          ? 'admin-login-role-btn py-2.5 px-3 rounded-xl text-xs sm:text-sm font-black bg-amber-500 text-slate-950 shadow transition text-center'
          : 'admin-login-role-btn py-2.5 px-3 rounded-xl text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition text-center';
      } else {
        btn.className = isAct
          ? 'admin-login-role-btn py-2.5 px-3 rounded-xl text-xs sm:text-sm font-black bg-sky-500 text-slate-950 shadow transition text-center'
          : 'admin-login-role-btn py-2.5 px-3 rounded-xl text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition text-center';
      }
    });

    // 2. 卡片邊框、圖示、標題與送出按鈕主題
    const cardEl = $('admin-login-card');
    if (cardEl) {
      cardEl.className = isSuperRole
        ? 'bg-slate-900 border border-amber-500/45 rounded-3xl p-5 sm:p-6 space-y-4 shadow-2xl'
        : 'bg-slate-900 border border-sky-500/45 rounded-3xl p-5 sm:p-6 space-y-4 shadow-2xl';
    }
    const iconEl = $('admin-login-icon');
    if (iconEl) {
      iconEl.textContent = isSuperRole ? '🛡️' : '👤';
      iconEl.className = isSuperRole
        ? 'w-10 h-10 rounded-xl bg-amber-500/15 border border-amber-500/35 flex items-center justify-center text-xl shrink-0'
        : 'w-10 h-10 rounded-xl bg-sky-500/15 border border-sky-500/35 flex items-center justify-center text-xl shrink-0';
    }
    if ($('admin-login-title')) {
      $('admin-login-title').textContent = isSuperRole ? '最高管理權責登入' : '活動承辦人登入';
    }
    if ($('admin-login-subtitle')) {
      $('admin-login-subtitle').textContent = isSuperRole
        ? '管理全平台帳號、歷年成績匯入與資料庫備份還原'
        : '管理本人承辦之活動、名單、成績與結案證書';
    }
    if ($('admin-login-accounts-count')) {
      $('admin-login-accounts-count').textContent = list.length ? `快選 (${list.length})` : '';
    }
    const submitBtn = $('btn-admin-login-submit');
    if (submitBtn && !submitBtn.disabled) {
      submitBtn.textContent = isSuperRole ? '🛡️ 登入最高管理權責後台' : '👤 登入活動承辦人後台';
      submitBtn.className = isSuperRole
        ? 'w-full py-3.5 rounded-xl text-base font-black bg-amber-500 hover:bg-amber-400 text-slate-950 transition shadow-lg'
        : 'w-full py-3.5 rounded-xl text-base font-black bg-sky-500 hover:bg-sky-400 text-slate-950 transition shadow-lg';
    }

    // 3. 若目前選定帳號不屬於該身分則清除
    if (state.adminLoginSelectedUser && !list.some((a) => a.username === state.adminLoginSelectedUser)) {
      state.adminLoginSelectedUser = '';
    }

    // 4. 精簡快速選取帳號膠囊（點一下即填入帳號並聚焦密碼欄，無帳號時隱藏不佔空間）
    const grid = $('admin-login-account-grid');
    if (!grid) return;
    if (list.length === 0) {
      grid.innerHTML = '';
      grid.classList.add('hidden');
      return;
    }
    grid.classList.remove('hidden');
    grid.innerHTML = list
      .map((a) => {
        const selected = state.adminLoginSelectedUser === a.username;
        const activeCls = isSuperRole
          ? 'bg-amber-500/20 border-amber-400 text-amber-200'
          : 'bg-sky-500/20 border-sky-400 text-sky-200';
        const idleCls = 'bg-slate-950 border-slate-700 text-slate-300 hover:border-slate-500';
        const label = a.displayName && a.displayName !== a.username
          ? `${a.username} (${a.displayName})`
          : a.username;
        return `
          <button type="button" data-pick-admin-user="${esc(a.username)}"
            class="px-2.5 py-1 rounded-lg border text-[11px] font-mono-num font-bold transition ${selected ? activeCls : idleCls}">
            ${selected ? '✓ ' : ''}${esc(label)}
          </button>
        `;
      })
      .join('');
  }

  /**
   * 依作業人員登入畫面所選活動，動態列出各作業組別卡片（含站點即時進度，供點選進入並登錄帳密）
   */
  function getStaffLoginGroupsForEvent(evt) {
    const allStaff = (state.platform?.staffAccounts || []).filter((s) => s && s.enabled !== false);
    const totalChips = (state.chips || []).length;

    if (!evt) {
      return allStaff
        .filter((s) => s.role === 'writer')
        .map((s) => ({
          username: s.username,
          displayName: s.displayName || s.username,
          role: s.role,
          cpId: 'WRITER',
          cpName: '跨活動共用晶片寫入',
          icon: '📡',
          colorCls: 'text-sky-300 bg-sky-500/15 border-sky-500/35',
          progressText: `已建檔 ${totalChips} 張晶片`,
          sortOrder: 0
        }));
    }

    const evtOwner = getEventOwner(evt);
    const isLap = evt.eventType === 'lap_timed';
    const cps = (evt.checkpoints || []).filter((c) => c.type !== 'start' && String(c.id || '').toUpperCase() !== 'START');
    const cpMap = new Map(cps.map((c, idx) => [String(c.id).toUpperCase(), { ...c, _order: idx + 2 }]));

    const evtRegs = (state.allRegistrations || []).filter((r) => r.eventId === evt.eventId);
    const totalRegs = evtRegs.length;
    const evtScans = (state.allScans || []).filter((s) => !s.deleted && s.eventId === evt.eventId);
    const checkedInBibs = new Set(evtScans.filter((s) => String(s.checkpointId).toUpperCase() === TE.CHECKIN_CP_ID).map((s) => s.bib));
    const lapScanCount = evtScans.filter((s) => String(s.checkpointId).toUpperCase() !== TE.CHECKIN_CP_ID).length;

    const groups = [];
    for (const s of allStaff) {
      const sOwner = String(s.ownerAdmin || '').trim();
      if (sOwner && sOwner !== evtOwner) continue;
      const boundEvt = String(s.eventId || 'ALL').trim();
      if (boundEvt && boundEvt !== 'ALL' && boundEvt !== evt.eventId) continue;

      const cpUpper = String(s.checkpointId || '').trim().toUpperCase();

      if (s.role === 'writer') {
        groups.push({
          username: s.username,
          displayName: s.displayName || '晶片寫入發卡組',
          role: s.role,
          cpId: 'WRITER',
          cpName: '共用晶片寫入 / UID 綁定',
          icon: '📡',
          colorCls: 'text-sky-300 bg-sky-500/15 border-sky-500/35',
          progressText: `已建檔 ${totalChips} 張晶片`,
          sortOrder: 0
        });
        continue;
      }

      if (cpUpper === TE.CHECKIN_CP_ID) {
        groups.push({
          username: s.username,
          displayName: s.displayName || '選手報到感應組',
          role: s.role,
          cpId: TE.CHECKIN_CP_ID,
          cpName: '選手報到感應區',
          icon: '✅',
          colorCls: 'text-cyan-300 bg-cyan-500/15 border-cyan-500/35',
          progressText: `已報到 ${checkedInBibs.size} / ${totalRegs} 人`,
          sortOrder: 1
        });
        continue;
      }

      const matchedCp = cpMap.get(cpUpper);
      if (!matchedCp && boundEvt !== evt.eventId) continue;

      if (isLap || cpUpper === 'LAP') {
        groups.push({
          username: s.username,
          displayName: s.displayName || '繞圈計次感應組',
          role: s.role,
          cpId: cpUpper || 'LAP',
          cpName: matchedCp?.name || '繞圈計次感應門',
          icon: '🔁',
          colorCls: 'text-purple-300 bg-purple-500/15 border-purple-500/35',
          progressText: `已感應 ${lapScanCount} 圈`,
          sortOrder: matchedCp ? matchedCp._order : 2
        });
      } else {
        const passedBibs = new Set(evtScans.filter((sc) => String(sc.checkpointId).toUpperCase() === cpUpper).map((sc) => sc.bib));
        const isFinish = matchedCp?.type === 'finish' || cpUpper === 'FINISH';
        groups.push({
          username: s.username,
          displayName: s.displayName || matchedCp?.name || cpUpper,
          role: s.role,
          cpId: cpUpper,
          cpName: matchedCp?.name || cpUpper,
          icon: isFinish ? '🏆' : '💧',
          colorCls: isFinish
            ? 'text-amber-300 bg-amber-500/15 border-amber-500/35'
            : 'text-emerald-300 bg-emerald-500/15 border-emerald-500/35',
          progressText: isFinish
            ? `已完賽 ${passedBibs.size} / ${totalRegs} 人`
            : `已通過 ${passedBibs.size} / ${totalRegs} 人`,
          sortOrder: matchedCp ? matchedCp._order : 50
        });
      }
    }

    return groups.sort((a, b) => a.sortOrder - b.sortOrder || a.username.localeCompare(b.username));
  }

  function renderStaffLoginGroups() {
    const grid = $('staff-login-group-grid');
    if (!grid) return;
    const staffSel = $('staff-login-event');
    const openEvents = getOpenEvents();
    const selectedEvtId = staffSel?.value || state.selectedEventId || openEvents[0]?.eventId || '';
    const evt = openEvents.find((e) => e.eventId === selectedEvtId) || null;

    const groups = getStaffLoginGroupsForEvent(evt);
    if ($('staff-login-groups-count')) {
      $('staff-login-groups-count').textContent = `${groups.length} 個作業組別`;
    }

    // 若目前選取的組別不在此活動清單內，自動重設
    if (state.staffLoginSelectedGroup && !groups.some((g) => g.username === state.staffLoginSelectedGroup)) {
      state.staffLoginSelectedGroup = '';
    }

    const selectedObj = groups.find((g) => g.username === state.staffLoginSelectedGroup) || null;
    const badgeEl = $('staff-selected-group-badge');
    if (badgeEl) {
      if (selectedObj) {
        badgeEl.textContent = `✅ 已選定：${selectedObj.icon} ${selectedObj.displayName}（${selectedObj.cpId}）`;
        badgeEl.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-black bg-emerald-500/20 border border-emerald-500/50 text-emerald-300';
      } else {
        badgeEl.textContent = '請先點選上方作業組別';
        badgeEl.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-slate-800 text-slate-300';
      }
    }

    if (!groups.length) {
      grid.innerHTML = '<div class="sm:col-span-2 p-4 rounded-2xl bg-slate-950 border border-dashed border-slate-800 text-center text-xs text-slate-500">本活動尚無對應的作業組別帳號，請由後台「專區 1 → 👷 作業員」新增</div>';
      return;
    }

    grid.innerHTML = groups.map((g) => {
      const isSelected = g.username === state.staffLoginSelectedGroup;
      return `
        <button type="button" data-pick-staff-group="${esc(g.username)}"
          class="text-left p-3.5 rounded-2xl border-2 transition space-y-2 ${isSelected
            ? 'bg-emerald-500/15 border-emerald-400 shadow-lg ring-1 ring-emerald-400/40'
            : 'bg-slate-950 border-slate-800 hover:border-emerald-500/50 hover:bg-slate-900'}">
          <div class="flex items-start justify-between gap-2">
            <div class="flex items-center gap-2 min-w-0">
              <span class="text-xl shrink-0">${g.icon}</span>
              <div class="min-w-0">
                <div class="text-sm font-black text-white truncate">${esc(g.displayName)}</div>
                <div class="text-[11px] text-slate-400 truncate">${esc(g.cpName)}</div>
              </div>
            </div>
            <span class="px-2 py-0.5 rounded-lg border text-[10px] font-mono-num font-black shrink-0 ${g.colorCls}">${esc(g.cpId)}</span>
          </div>
          <div class="flex items-center justify-between gap-2 pt-1.5 border-t border-slate-800/80 text-[11px] font-mono-num">
            <span class="text-emerald-300 font-bold">${esc(g.progressText)}</span>
            <span class="${isSelected ? 'text-emerald-300 font-black' : 'text-slate-400'}">
              ${isSelected ? '✓ 已選取' : `帳號 ${esc(g.username)} →`}
            </span>
          </div>
        </button>
      `;
    }).join('');
  }

  function updateNetworkIndicators() {
    const settings = DB.getDeviceSettings();
    const onlineNet = navigator.onLine;
    const badge = $('badge-network');

    if (settings.mode === 'online' && settings.gasUrl) {
      badge.innerHTML = onlineNet
        ? '<span class="w-2 h-2 rounded-full bg-sky-400 animate-pulse"></span><span>雲端連線</span>'
        : '<span class="w-2 h-2 rounded-full bg-amber-400"></span><span>離線暫存</span>';
      badge.className = onlineNet
        ? 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-sky-950 border border-sky-500/50 text-sky-300'
        : 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-amber-950 border border-amber-500/50 text-amber-300';
    } else {
      badge.innerHTML = '<span class="w-2 h-2 rounded-full bg-emerald-400"></span><span>單機離線</span>';
      badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-slate-800 border border-slate-700 text-slate-300';
    }

    const portalBanner = $('portal-net-banner');
    const portalCloudTip = $('portal-cloud-tip');
    if (portalCloudTip) {
      portalCloudTip.classList.toggle('hidden', Boolean(settings.gasUrl));
    }
    if (portalBanner && !_isAutoSyncing) {
      if (onlineNet) {
        portalBanner.className = 'inline-flex items-center gap-2 px-3.5 py-1 rounded-full text-xs font-bold bg-emerald-950/70 border border-emerald-500/40 text-emerald-300';
        const evtCount = (state.platform?.events || []).length;
        const lastStr = settings.lastSyncAt ? ` • ${TE.formatClockSec(settings.lastSyncAt)}` : '';
        portalBanner.textContent = settings.gasUrl
          ? `🟢 雲端已連線（${evtCount} 場活動）${lastStr}`
          : '🟡 本機離線模式';
      } else {
        portalBanner.className = 'inline-flex items-center gap-2 px-3.5 py-1 rounded-full text-xs font-bold bg-amber-950/80 border border-amber-500/50 text-amber-300';
        portalBanner.textContent = '📴 離線模式（可正常作業，連網後上傳）';
      }
    }

    const loginHint = $('staff-login-offline-hint');
    if (loginHint) {
      if (!onlineNet) {
        loginHint.className = 'inline-block rounded-full border border-amber-500/40 bg-amber-950/40 px-3 py-1 text-xs text-amber-200 font-bold';
        loginHint.textContent = '📴 離線模式';
      } else if (settings.gasUrl) {
        loginHint.className = 'inline-block rounded-full border border-sky-500/40 bg-sky-950/40 px-3 py-1 text-xs text-sky-200 font-bold';
        loginHint.textContent = '🌐 雲端連線就緒';
      } else {
        loginHint.className = 'inline-block rounded-full border border-emerald-500/30 bg-emerald-950/30 px-3 py-1 text-xs text-emerald-200 font-bold';
        loginHint.textContent = '💾 本機單機模式';
      }
    }

    const toggleBtn = $('btn-staff-toggle-mode');
    if (toggleBtn) {
      if (settings.mode === 'online' && settings.gasUrl) {
        toggleBtn.textContent = onlineNet ? '🌐 連線中' : '⚠️ 待訊號';
        toggleBtn.className = onlineNet
          ? 'px-3 py-2 rounded-xl text-xs font-bold bg-sky-600/30 hover:bg-sky-600/50 border border-sky-500/60 text-sky-300 transition'
          : 'px-3 py-2 rounded-xl text-xs font-bold bg-amber-600/30 border border-amber-500/60 text-amber-300 transition';
      } else {
        toggleBtn.textContent = '📴 離線中';
        toggleBtn.className = 'px-3 py-2 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-700 border border-slate-700 text-slate-200 transition';
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
      syncInfo.textContent = `待同步：晶片 ${pending.chips}、報名 ${pending.registrations}、感應 ${pending.scans}、歷年 ${pending.history || 0} 筆 ｜ 最後同步：${lastStr}`;
    }
  }

  /* ========================================================================
   * 登入與登出流程（本機秒驗快速進入 ＋ 背景雲端同步驗證 ＋ 自動維持登入）
   * ======================================================================== */
  let _isLoggingIn = false;

  async function syncLoggedInSessionInBackground(sessionSnapshot) {
    if (!sessionSnapshot?.auth || !navigator.onLine || !Sync.isOnlineEnabled()) return;
    try {
      const res = await Sync.pullFromCloud(sessionSnapshot.auth);
      // 確保期間未登出或換人
      if (!state.session || state.session.loginAt !== sessionSnapshot.loginAt) return;
      if (res && res.isAdmin) {
        await DB.cacheAdminAccount({
          ...state.session.account,
          username: res.adminUsername || state.session.account.username,
          isSuperAdmin: Boolean(res.isSuperAdmin),
          password: sessionSnapshot.auth.password
        });
      }
      await reloadState();
    } catch (err) {
      if (!state.session || state.session.loginAt !== sessionSnapshot.loginAt) return;
      const msg = String(err.message || '');
      if (err.multiAdmin && (msg.includes('驗證失敗') || msg.includes('錯誤') || msg.includes('停用'))) {
        NFC.sound.beepError();
        state.session = null;
        state.adminZone = null;
        DB.saveSession(null);
        clearLoginInputs();
        showToast('⚠️ 雲端帳密已變更或帳號已停用，請重新登入', 'error');
        switchView(sessionSnapshot.mode === 'admin' ? 'admin-login' : 'staff-login');
      } else {
        console.warn('登入後背景雲端同步暫緩:', err);
      }
    }
  }

  async function handleLogin(mode, username, password, chosenEventId = '', adminRoleIntent = '') {
    if (_isLoggingIn) return;
    const u = String(username || '').trim();
    const p = String(password || '');
    if (!u || !p) return showToast('請輸入帳號與密碼', 'warn');

    _isLoggingIn = true;
    const submitBtn = mode === 'admin' ? $('btn-admin-login-submit') : $('form-staff-login')?.querySelector('button[type="submit"]');
    const origBtnText = submitBtn ? submitBtn.textContent : '';
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.textContent = '⏳ 驗證登入中...';
    }

    try {
      let account = null;
      let usedCloud = false;
      let needBackgroundSync = false;
      let cloudLegacyFallback = false;

      // 1. 極速路徑：優先以本機已快取帳號秒速驗證（<10ms），不卡在雲端連線等待！
      const localRes = await DB.verifyLocalLogin(mode, u, p);
      if (localRes.ok) {
        account = { ...localRes.account, password: p };
        needBackgroundSync = navigator.onLine && Sync.isOnlineEnabled();
      } else if (navigator.onLine && Sync.hasUrl()) {
        // 2. 若本機無此帳號或雲端剛改過密碼，才以 lite 極速模式向雲端驗證
        try {
          const res = await Sync.loginCloud(mode, u, p, { lite: true });
          account = {
            ...res.account,
            password: p,
            isSuperAdmin: mode === 'admin'
              ? (res.account?.isSuperAdmin !== undefined ? Boolean(res.account.isSuperAdmin) : true)
              : false
          };
          usedCloud = true;
          needBackgroundSync = Boolean(res.lite && !res.data);
        } catch (err) {
          if (String(err.message).includes('錯誤') || String(err.message).includes('停用')) {
            if (mode === 'admin' && !err.multiAdmin) {
              cloudLegacyFallback = true;
            } else {
              NFC.sound.beepError();
              showToast(err.message, 'error');
              return;
            }
          } else {
            NFC.sound.beepError();
            showToast(`登入失敗：${err.message}`, 'error');
            return;
          }
        }
      }

      if (!account) {
        NFC.sound.beepError();
        showToast(localRes.error || '帳號或密碼錯誤', 'error');
        return;
      }

      // 後台管理登入身分核對：防止最高管理權責與活動承辦人混用入口
      if (mode === 'admin' && adminRoleIntent) {
        if (adminRoleIntent === 'super' && !account.isSuperAdmin) {
          NFC.sound.beepError();
          showToast('此帳號為「活動承辦人」權限，請改點上方「👤 活動承辦人」登入', 'warn');
          return;
        }
        if (adminRoleIntent === 'coordinator' && account.isSuperAdmin) {
          NFC.sound.beepError();
          showToast('此帳號為「最高管理權責」，請改點上方「🛡️ 最高管理權責」登入', 'warn');
          return;
        }
      }

      // 決定作業員鎖定的活動：若帳號本身有綁定特定 eventId（非 ALL）則強制鎖定該活動；否則鎖定登入時選取的活動
      const platform = await DB.getPlatform();
      const events = platform.events || [];
      const openEventsList = events.filter((e) => !isEventFinished(e));
      let lockedEventId = '';
      if (mode === 'staff') {
        const isWriterOnly = account.role === 'writer';
        if (account.eventId && account.eventId !== 'ALL') {
          lockedEventId = account.eventId;
          const boundEvt = events.find((e) => e.eventId === lockedEventId);
          if (!isWriterOnly && isEventFinished(boundEvt)) {
            NFC.sound.beepError();
            showToast(`活動「${boundEvt.name}」狀態為「已結束」，已鎖定禁止作業員登入感應以避免資料遭修改`, 'error');
            return;
          }
        } else {
          const candidateId = chosenEventId || state.selectedEventId || openEventsList[0]?.eventId || '';
          const targetEvt = openEventsList.find((e) => e.eventId === candidateId) || null;
          if (!isWriterOnly && !targetEvt) {
            NFC.sound.beepError();
            showToast('目前無進行中或籌備中的路跑活動可供作業（已結束活動已鎖定）', 'error');
            return;
          }
          lockedEventId = targetEvt ? targetEvt.eventId : '';
          const staffOwner = String(account.ownerAdmin || '').trim();
          if (staffOwner && lockedEventId && targetEvt) {
            const targetOwner = String(targetEvt.ownerAdmin || platform.adminUsername || 'admin').trim();
            if (targetOwner !== staffOwner) {
              NFC.sound.beepError();
              showToast(`此作業員帳號僅限操作管理者「${staffOwner}」建立的活動，請重新選擇活動`, 'error');
              return;
            }
          }
        }
        state.selectedEventId = lockedEventId;
        DB.saveDeviceSettings({ selectedEventId: lockedEventId });
      } else if (mode === 'admin') {
        const defaultOwner = String(platform.adminUsername || 'admin').trim();
        const openOwnEvents = openEventsList.filter((e) => (String(e.ownerAdmin || '').trim() || defaultOwner) === account.username);
        if (!openOwnEvents.some((e) => e.eventId === state.selectedEventId)) {
          state.selectedEventId = openOwnEvents[0]?.eventId || '';
        }
      }

      state.session = {
        mode,
        account,
        auth: { username: u, password: p },
        lockedEventId,
        loginAt: Date.now()
      };
      clearLoginInputs();
      DB.saveSession(state.session);
      await reloadState();

      if (mode === 'admin') {
        const ownCount = getAdminOwnEvents().length;
        const roleLabel = account.isSuperAdmin ? '最高管理權責' : '活動承辦人';
        if (cloudLegacyFallback) {
          showToast(`⚠️ 已以本機管理者「${account.username}」登入（請同步更新部署 gas/Code.gs）`, 'warn');
        } else {
          showToast(
            `${account.isSuperAdmin ? '🛡️' : '👤'} 已登入${roleLabel}：${account.displayName || account.username}（共 ${ownCount} 場本人活動）${usedCloud ? '・雲端驗證完成' : ''}`,
            'success'
          );
        }
        state.adminZone = 'ops';
        switchView('admin-workspace');
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

      // 3. 登入畫面已秒速進入後，於背景非同步拉取最新雲端資料並複驗權限
      if (needBackgroundSync) {
        syncLoggedInSessionInBackground(state.session);
      }
    } finally {
      _isLoggingIn = false;
      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = origBtnText;
      }
    }
  }

  function handleLogout() {
    NFC.cancelActiveOperation();
    if (NFC.isScanning) {
      NFC.stopScanning();
      resetScanButtonUi();
    }
    state.session = null;
    state.adminZone = null;
    state.adminLoginSelectedUser = '';
    state.staffLoginSelectedGroup = '';
    clearLoginInputs();
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
      ? '🌐 共用製卡'
      : `🏟️ ${evt ? evt.name : '未選定'}`;

    let zoneTitle = '';
    if (role === 'writer') zoneTitle = '晶片寫入 / UID 綁定';
    else if (role === 'both') zoneTitle = `寫入 ＋ ${lockedCp ? lockedCp.name : '未指派'}`;
    else zoneTitle = `${lockedCp ? lockedCp.name : '未指派站點'}`;

    $('staff-locked-zone-title').textContent = zoneTitle;
    $('staff-operator-meta').textContent = `${acc.displayName || acc.username} (${acc.username})`;

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
          ? `📋 ${evt.name} ｜ 已報到 ${checkedInCount} / ${evtRegs.length}`
          : isLap
            ? `🔁 ${evt.name}（${evt.lapDistanceKm} km/圈）`
            : `🏟️ ${evt.name}`)
      : '未選定活動';
    $('locked-cp-name').textContent = lockedCp
      ? (isCheckIn ? `✅ ${lockedCp.name}` : isLap ? `🔁 ${lockedCp.name}` : `#${lockedCp.order || '-'} ${lockedCp.name}`)
      : '未指派站點';
    $('locked-cp-meta').textContent = lockedCp
      ? (isCheckIn
          ? 'CHECKIN'
          : isLap
            ? `${lockedCp.id} • ${evt.lapDistanceKm} km/圈`
            : `${lockedCp.id} • ${lockedCp.distanceKm} km`)
      : '';

    const col3Label = $('hero-col3-label');
    if (col3Label) {
      col3Label.textContent = isCheckIn ? '報到進度' : isLap ? '圈數 / 里程' : '經過時間';
    }

    // 模擬器僅直接載入「本場活動」已報名人員，不載入其他未參加本場活動之人員
    const mockSel = $('mock-chip-select');
    const prevVal = mockSel.value;

    const sortedEvtRegs = [...evtRegs].sort((a, b) =>
      String(a.bib || '').localeCompare(String(b.bib || ''), 'zh-Hant', { numeric: true })
    );
    const validKeySet = new Set();

    mockSel.innerHTML =
      `<option value="">${sortedEvtRegs.length ? `-- 隨機刷過本場選手（共 ${sortedEvtRegs.length} 人） --` : '-- 本活動尚無報名選手 --'}</option>` +
      sortedEvtRegs
        .map((reg) => {
          const idUpper = String(reg.idNumber || '').trim().toUpperCase();
          const optKey = idUpper || `BIB:${reg.bib}`;
          validKeySet.add(optKey);
          const tag = isCheckIn
            ? (checkedInBibSet.has(reg.bib) ? '✅已報到' : '⏳未報到')
            : '✅本場選手';
          const idHint = reg.idNumber ? ` (${reg.idNumber})` : '';
          return `<option value="${esc(optKey)}">${esc(reg.bib ? `#${reg.bib} ` : '')}${esc(reg.name)}${esc(idHint)} ${tag}</option>`;
        })
        .join('');
    if (prevVal && validKeySet.has(prevVal)) mockSel.value = prevVal;

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
    if (!evt || isEventFinished(evt) || !lockedCp || !lockedCp.id) {
      NFC.sound.beepError();
      showToast('活動已結束或尚未鎖定站點，無法執行感應', 'error');
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
      $('hero-elapsed').textContent = '--:--:--';
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
      const elapsedStr = gunMs !== null ? TE.formatDuration(scanTs - gunMs, false) : '尚未設定起跑時間';

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
    btn.textContent = '📡 啟動 NFC 感應';
    btn.className = 'no-select w-full sm:w-auto sm:min-w-[280px] py-5 px-8 rounded-2xl text-xl sm:text-2xl font-black bg-emerald-500 hover:bg-emerald-400 active:scale-[0.98] text-slate-950 shadow-lg transition';
  }

  async function toggleNfcScanning() {
    const btn = $('btn-toggle-nfc-scan');
    if (NFC.isScanning) {
      NFC.stopScanning();
      resetScanButtonUi();
      $('scan-status-pill').textContent = '已暫停感應';
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-slate-800 text-slate-300 mb-4';
      return;
    }
    try {
      await NFC.startScanning(
        (tag) => processIncomingScan(tag),
        (err) => showToast(err.message, 'error')
      );
      btn.textContent = '🛑 停止感應';
      btn.className = 'no-select nfc-active-pulse w-full sm:w-auto sm:min-w-[280px] py-5 px-8 rounded-2xl text-xl sm:text-2xl font-black bg-rose-500 hover:bg-rose-400 text-white shadow-lg transition';
      $('scan-status-pill').textContent = '🟢 感應中 — 請靠近晶片';
      $('scan-status-pill').className = 'inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-extrabold bg-emerald-950 border border-emerald-500 text-emerald-300 mb-4';
      showToast('已啟動 NFC 感應', 'success');
    } catch (err) {
      showToast(err.message, 'warn');
    }
  }

  /* ---------------- 共用晶片寫入面板（僅姓名＋身份證字號） ---------------- */
  function renderWriterPanel() {
    const chipMap = new Map(state.chips.map((c) => [c.idNumber, c]));
    const peopleMap = new Map();

    // 彙整所有進行中／籌備中活動名單的選手（排除「已結束」活動名單以避免誤操作）
    const openEventIdSet = new Set(getOpenEvents().map((e) => e.eventId));
    for (const r of state.allRegistrations) {
      if (!openEventIdSet.has(r.eventId)) continue;
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
   * VIEW 5：後台管理區（8 大分頁渲染）
   * ======================================================================== */
  function updateAdminZoneNavBadges() {
    const isSuper = isCurrentSuperAdmin();
    const openCnt = getAdminOpenOwnEvents().length;
    const finishedCnt = getAdminFinishedEvents().length;
    const totalEvents = (state.platform?.events || []).length;
    const totalScans = (state.allScans || []).length;
    if ($('zone-ops-badge')) $('zone-ops-badge').textContent = `${openCnt} 場開放`;
    if ($('zone-closed-badge')) $('zone-closed-badge').textContent = `${finishedCnt} 場已結案`;
    if ($('zone-db-badge')) $('zone-db-badge').textContent = `${totalEvents} 活動・${totalScans} 感應`;
    if (!isSuper && (state.adminZone === 'history' || state.adminTab === 'history')) {
      state.adminZone = null;
      state.adminTab = 'events';
    }
    updateAdminZoneNavStyles(state.adminZone);
  }

  function renderAdminWorkspace() {
    if (!state.session || state.session.mode !== 'admin') return;
    updateAdminZoneNavBadges();
    renderAdminEventsTab();
    renderAdminRosterTab();
    renderAdminResultsTab();
    renderAdminRaceTab();
    renderAdminStaffTab();
    renderAdminHistoryTab();
    renderAdminExportTab();
    renderAdminCertificateTab();
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
      finished: '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-slate-700 text-slate-300">🔒 已結束（已鎖定）</span>'
    };

    const ownerBadge = $('admin-events-owner-badge');
    if (ownerBadge) {
      const openOwnCount = getAdminOpenOwnEvents().length;
      const finishedCount = events.length - openOwnCount;
      ownerBadge.textContent = finishedCount > 0
        ? `${adminUser} • 共 ${events.length} 場 (${finishedCount} 已結束)`
        : `${adminUser} • 共 ${events.length} 場`;
    }

    if (events.length === 0) {
      $('admin-events-list').innerHTML = `
        <div class="md:col-span-2 p-6 rounded-2xl border border-dashed border-slate-700 bg-slate-950/60 text-center text-xs text-slate-400">
          尚無活動，請點擊上方「<strong class="text-emerald-400">＋ 新增活動</strong>」
        </div>
      `;
    } else {
      $('admin-events-list').innerHTML = events
        .map((ev) => {
          const regCount = state.allRegistrations.filter((r) => r.eventId === ev.eventId).length;
          const scanCount = state.allScans.filter((s) => !s.deleted && s.eventId === ev.eventId).length;
          const finished = isEventFinished(ev);
          const isSelected = !finished && currentEvt && ev.eventId === currentEvt.eventId;
          const isDefault = ev.eventId === state.platform.activeEventId;
          return `
            <div class="p-4 rounded-2xl border ${finished ? 'border-slate-800/80 bg-slate-950/60 opacity-80' : (isSelected ? 'border-amber-400 bg-amber-950/20' : 'border-slate-800 bg-slate-950')} space-y-2">
              <div class="flex items-center gap-2 flex-wrap">
                <span class="font-black ${finished ? 'text-slate-300' : 'text-white'} text-base">${esc(ev.name)}</span>
                ${ev.eventType === 'lap_timed'
                  ? `<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-purple-500/20 text-purple-300">🔁 ${ev.lapDistanceKm}km/圈</span>`
                  : '<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-sky-500/20 text-sky-300">🏃 距離賽</span>'}
                ${statusBadge[ev.status] || ''}
                ${isDefault ? '<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-sky-600 text-white">預設</span>' : ''}
              </div>
              <div class="text-xs text-slate-400 font-mono-num">
                ${esc(ev.date || '--')}${ev.location ? ` ｜ ${esc(ev.location)}` : ''} ｜ ${regCount} 人 ｜ ${scanCount} 筆感應
              </div>
              <div class="flex flex-wrap gap-1.5 pt-1">
                ${finished
                  ? `<button type="button" data-reopen-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-amber-500/20 hover:bg-amber-500/35 text-amber-300 border border-amber-500/40">🔓 重新開啟</button>`
                  : `<button type="button" data-select-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold ${isSelected ? 'bg-amber-500 text-slate-950' : 'bg-slate-800 hover:bg-slate-700 text-white'}">
                      ${isSelected ? '✓ 管理中' : '切換'}
                    </button>
                    <button type="button" data-default-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-sky-300">
                      設為預設
                    </button>`}
                <button type="button" data-delete-event="${esc(ev.eventId)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪除</button>
              </div>
            </div>
          `;
        })
        .join('');
    }

    $('form-event-basic').classList.toggle('hidden', !currentEvt);
    const noBasicHint = $('admin-no-event-basic-hint');
    noBasicHint.classList.toggle('hidden', Boolean(currentEvt));
    if (!currentEvt) {
      noBasicHint.textContent = events.length > 0
        ? '🔒 所有活動皆已結束，如需編輯請先點「🔓 重新開啟」或新增活動。'
        : '請先點擊上方「＋ 新增活動」。';
    }

    if (currentEvt) {
      $('cfg-evt-name').value = currentEvt.name || '';
      $('cfg-evt-en-name').value = currentEvt.englishName || '';
      $('cfg-evt-org').value = currentEvt.organization || 'Taiwan Ultra Runners';
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
    // 密碼欄永遠預設空白（不回填明文密碼，避免外人窺見或瀏覽器自動填入）
    const cfgPassEl = $('cfg-admin-pass');
    if (cfgPassEl) {
      cfgPassEl.value = '';
      cfgPassEl.setAttribute('readonly', 'readonly');
      bindNoAutofillUnlock(cfgPassEl);
    }

    // 最高管理員專用：多管理者帳號管理區塊（所有管理者與承辦人員密碼欄一律預設空白，不可自動填入）
    const manageCard = $('admin-accounts-manage-card');
    if (manageCard) {
      manageCard.classList.toggle('hidden', !isSuper);
      if (isSuper) {
        const admins = state.platform.adminAccounts || [];
        $('admin-accounts-list').innerHTML = admins
          .map((a, idx) => {
            const isSelf = a.username === adminUser;
            const hasPass = Boolean(String(a.password || ''));
            const ownEvtCount = (state.platform.events || []).filter((e) => getEventOwner(e) === a.username).length;
            return `
              <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-end p-3.5 rounded-xl bg-slate-950 border ${isSelf ? 'border-amber-500/40' : 'border-slate-800'}" data-admin-row="${idx}" data-admin-orig-user="${esc(a.username)}">
                <div class="sm:col-span-2">
                  <label class="block text-[11px] text-slate-400 mb-1">登入帳號</label>
                  <input type="text" value="${esc(a.username)}" autocomplete="off" data-admin-field="username" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-mono-num font-bold text-white" />
                </div>
                <div class="sm:col-span-2">
                  <label class="block text-[11px] text-slate-400 mb-1">${hasPass ? '重設密碼（預設空白）' : '設定密碼（必填）'}</label>
                  <input type="text" value="" readonly autocomplete="new-password" data-lpignore="true" data-1p-ignore="true" data-bwignore="true" data-admin-field="password" placeholder="${hasPass ? '留空維持原密碼' : '請手動輸入密碼'}" class="tur-secret-input w-full bg-slate-900 border ${hasPass ? 'border-slate-700' : 'border-amber-500/70'} rounded-lg px-2.5 py-2 text-xs font-mono-num text-white" />
                </div>
                <div class="sm:col-span-3">
                  <label class="block text-[11px] text-slate-400 mb-1">管理者名稱（名下 ${ownEvtCount} 場活動）</label>
                  <input type="text" value="${esc(a.displayName || a.username)}" autocomplete="off" data-admin-field="displayName" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-2 text-xs font-bold text-white" />
                </div>
                <div class="sm:col-span-3">
                  <label class="block text-[11px] text-slate-400 mb-1">管理權限角色</label>
                  <select data-admin-field="isSuperAdmin" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs font-bold text-amber-300" ${isSelf ? 'disabled' : ''}>
                    <option value="0" ${!a.isSuperAdmin ? 'selected' : ''}>活動承辦人（僅管本人活動・無備份還原）</option>
                    <option value="1" ${a.isSuperAdmin ? 'selected' : ''}>最高管理權責（全權限・含備份還原）</option>
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
        initNoCachePasswordInputs($('admin-accounts-list'));
      }
    }
  }

  function renderAdminRosterTab() {
    const evt = getActiveEvent();
    const tbody = $('admin-reg-tbody');
    const contentEl = $('admin-roster-content');
    const emptyEl = $('admin-roster-empty');
    if (contentEl) contentEl.classList.toggle('hidden', !evt);
    if (emptyEl) emptyEl.classList.toggle('hidden', Boolean(evt));
    if (!evt) {
      const ownTotal = getAdminOwnEvents().length;
      if ($('admin-roster-empty-title')) {
        $('admin-roster-empty-title').textContent = ownTotal > 0 ? '無開放活動' : '尚無活動';
      }
      if ($('admin-roster-empty-desc')) {
        $('admin-roster-empty-desc').textContent = ownTotal > 0 ? '已結束活動已鎖定，請重新開啟或新增活動。' : '請先點擊上方「＋ 新增活動」。';
      }
      $('roster-import-event-name').textContent = '';
      $('admin-reg-count').textContent = '0';
      $('admin-roster-checkin-summary').textContent = '✅ 0 / 0';
      tbody.innerHTML = '';
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
    $('admin-roster-checkin-summary').textContent = `✅ 已報到 ${checkedInCount} / ${totalCount}`;

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
          : `<button type="button" data-quick-checkin-bib="${esc(r.bib)}" data-quick-checkin-idnum="${esc(r.idNumber || '')}" class="px-2 py-0.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-sky-600/40 text-slate-300 hover:text-sky-200 border border-slate-700">+ 補報到</button>`;
        const bioMeta = [r.gender, r.dob].filter(Boolean).join(' · ');
        return `
          <tr class="hover:bg-slate-800/40">
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-300">${esc(r.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(r.name)}${bioMeta ? ` <span class="text-[11px] font-mono-num font-normal text-slate-400">(${esc(bioMeta)})</span>` : ''}</td>
            <td class="py-2.5 pr-3 font-mono-num text-sky-300">${esc(r.idNumber || r.idMasked || '--')}</td>
            <td class="py-2.5 pr-3 text-slate-300">${esc(catMap.get(r.categoryId) || r.categoryId)}</td>
            <td class="py-2.5 pr-3">${checkInCell}</td>
            <td class="py-2.5 pr-3 text-xs font-bold ${chip ? 'text-emerald-400' : 'text-slate-500'}">
              ${chip ? `✅ 已建檔${chip.nfcUid ? ` (${esc(chip.nfcUid.slice(0, 11))})` : ''}` : '⏳ 未製卡'}
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
    const contentEl = $('admin-results-content');
    const emptyEl = $('admin-results-empty');
    const hasData = Boolean(evt && computed);
    if (contentEl) contentEl.classList.toggle('hidden', !hasData);
    if (emptyEl) emptyEl.classList.toggle('hidden', hasData);
    if (!hasData) {
      const ownTotal = getAdminOwnEvents().length;
      if ($('admin-results-empty-title')) {
        $('admin-results-empty-title').textContent = ownTotal > 0 ? '無開放活動' : '尚無活動';
      }
      if ($('admin-results-empty-desc')) {
        $('admin-results-empty-desc').textContent = ownTotal > 0 ? '已結束活動已鎖定，請重新開啟或新增活動。' : '請先點擊上方「＋ 新增活動」。';
      }
      $('kpi-total').textContent = '0';
      $('kpi-checkin').textContent = '0';
      $('kpi-finished').textContent = '0';
      $('kpi-oncourse').textContent = '0';
      $('kpi-notstarted').textContent = '0';
      $('results-category-filter').innerHTML = '';
      $('results-thead').innerHTML = '';
      tbody.innerHTML = '';
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
      $('kpi-label-1').textContent = '報名';
      $('kpi-label-checkin').textContent = '已報到';
      $('kpi-label-2').textContent = '總圈數';
      $('kpi-label-3').textContent = '總里程 (km)';
      $('kpi-label-4').textContent = '最高里程 (km)';

      $('kpi-total').textContent = String(s.totalRunners);
      $('kpi-checkin').textContent = String(s.checkedInCount ?? 0);
      $('kpi-finished').textContent = `${s.totalLapsAll} 圈`;
      $('kpi-oncourse').textContent = s.totalKmAll.toFixed(2);
      $('kpi-notstarted').textContent = s.topDistanceKm.toFixed(2);

      $('results-edit-banner').textContent = `單圈 ${evt.lapDistanceKm} km • 點擊報到或逐圈管理可修正`;

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
    $('kpi-label-1').textContent = '報名';
    $('kpi-label-checkin').textContent = '已報到';
    $('kpi-label-2').textContent = '已完賽';
    $('kpi-label-3').textContent = '賽道上';
    $('kpi-label-4').textContent = '未報到';
    $('results-edit-banner').textContent = '點擊報到、站點時間或罰時可修改';

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
              ${row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, false) : '--'}
            </td>
            <td class="px-3 py-2.5 font-mono-num text-slate-300">${esc(row.avgPace)}</td>
          </tr>
        `;
      })
      .join('');
  }

  function renderAdminRaceTab() {
    const evt = getActiveEvent();
    const contentEl = $('admin-race-content');
    const emptyEl = $('admin-race-empty');
    if (contentEl) contentEl.classList.toggle('hidden', !evt);
    if (emptyEl) emptyEl.classList.toggle('hidden', Boolean(evt));
    if (!evt) {
      const ownTotal = getAdminOwnEvents().length;
      if ($('admin-race-empty-title')) {
        $('admin-race-empty-title').textContent = ownTotal > 0 ? '無開放活動' : '尚無活動';
      }
      if ($('admin-race-empty-desc')) {
        $('admin-race-empty-desc').textContent = ownTotal > 0 ? '已結束活動已鎖定，請重新開啟或新增活動。' : '請先點擊上方「＋ 新增活動」。';
      }
      $('unified-gun-display').textContent = '尚未設定';
      $('unified-gun-input').value = '';
      $('admin-lap-no-cp-banner').classList.add('hidden');
      $('admin-checkpoints-card').classList.remove('hidden');
      $('categories-list').innerHTML = '';
      $('checkpoints-list').innerHTML = '';
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

          <details class="pt-2 border-t border-slate-800/80 text-xs">
            <summary class="cursor-pointer flex flex-wrap items-center justify-between gap-2 select-none text-slate-400 hover:text-slate-200">
              <span class="font-bold">
                起跑：<strong class="font-mono-num ${effGun ? 'text-emerald-400' : 'text-slate-500'}">${effGun ? TE.formatClockSec(effGun) : '未設定'}</strong>${cutoffStr ? `<span class="text-amber-300 font-mono-num">${cutoffStr}</span>` : ''}
              </span>
              <span class="text-[11px] text-sky-400">個別微調 ▾</span>
            </summary>
            <div class="flex flex-wrap items-center gap-2 mt-2.5">
              <button type="button" data-gun-now="${idx}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-emerald-600/30 border border-slate-700 text-emerald-300">
                🔫 鳴槍
              </button>
              <input type="text" placeholder="HH:mm:ss" value="${effGun ? TE.formatClockSec(effGun) : ''}" data-gun-manual-input="${idx}"
                class="w-28 font-mono-num text-xs bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white" />
              <button type="button" data-gun-manual-set="${idx}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-sky-300">套用</button>
              <button type="button" data-gun-clear="${idx}" class="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-400">清除</button>
            </div>
          </details>
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
    // 限定活動下拉選單僅列出非「已結束」的本人活動
    const events = getAdminOpenOwnEvents();
    const evt = getActiveEvent();
    const cps = [...(evt?.checkpoints || [])].sort((a, b) => a.order - b.order);

    const allOptionLabel = isSuper ? '所有活動' : '本人所有活動';
    const listHtml = list.length === 0
      ? '<div class="p-5 rounded-xl bg-slate-950 border border-dashed border-slate-800 text-center text-xs text-slate-500">尚無作業員，請點上方「＋ 新增作業員」</div>'
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

  function renderAdminHistoryTab() {
    renderHistoryImportPreview();
    const all = [...(state.historyRecords || [])].sort((a, b) =>
      String(b.eventDate || '').localeCompare(String(a.eventDate || '')) || (a.overallRank || 99999) - (b.overallRank || 99999)
    );

    // 依「日期｜活動名稱」彙整已匯入歷年活動清單
    const evtMap = new Map();
    for (const h of all) {
      const key = `${h.eventDate}__${h.eventName}`;
      if (!evtMap.has(key)) {
        evtMap.set(key, { key, eventDate: h.eventDate, eventName: h.eventName, count: 0, finishedCount: 0, totalKm: 0 });
      }
      const g = evtMap.get(key);
      g.count += 1;
      if (String(h.status || '').includes('完賽') && !String(h.status || '').includes('未')) g.finishedCount += 1;
      g.totalKm += Number(h.distanceKm || 0);
    }
    const groups = Array.from(evtMap.values()).sort((a, b) => String(b.eventDate || '').localeCompare(String(a.eventDate || '')));

    const evCountEl = $('admin-history-events-count');
    const recCountEl = $('admin-history-records-count');
    if (evCountEl) evCountEl.textContent = String(groups.length);
    if (recCountEl) recCountEl.textContent = `${all.length} 筆`;
    if ($('zone-history-badge')) $('zone-history-badge').textContent = `${groups.length} 場 / ${all.length} 筆`;

    const evListEl = $('admin-history-events-list');
    if (evListEl) {
      evListEl.innerHTML = groups.length === 0
        ? '<div class="md:col-span-2 p-5 rounded-xl bg-slate-950 border border-dashed border-slate-800 text-center text-xs text-slate-500">尚未匯入歷年活動成績</div>'
        : groups.map((g) => {
            const is800 = g.eventName.includes('八百壯士');
            return `
              <div class="p-3.5 rounded-xl bg-slate-950 border ${is800 ? 'border-amber-500/40' : 'border-slate-800'} flex items-center justify-between gap-2">
                <div class="min-w-0">
                  <div class="flex items-center gap-1.5 flex-wrap">
                    <span class="font-mono-num text-[11px] font-bold px-1.5 py-0.5 rounded bg-slate-800 text-amber-300">${esc(g.eventDate)}</span>
                    <span class="font-bold text-white text-sm truncate">${esc(g.eventName)}</span>
                    ${is800 ? '<span class="px-1.5 py-0.5 rounded text-[10px] font-black bg-amber-500/20 text-amber-300">⚔️ 八百壯士</span>' : ''}
                  </div>
                  <div class="text-[11px] text-slate-400 font-mono-num mt-1">
                    共 ${g.count} 筆（完賽 ${g.finishedCount}）｜ 累計 ${g.totalKm.toFixed(1)} km
                  </div>
                </div>
                <button type="button" data-del-history-evt-date="${esc(g.eventDate)}" data-del-history-evt-name="${esc(g.eventName)}"
                  class="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300 shrink-0">
                  刪除整場
                </button>
              </div>
            `;
          }).join('');
    }

    const filterSel = $('admin-history-filter-event');
    if (filterSel) {
      filterSel.innerHTML = '<option value="ALL">全部歷年活動</option>' + groups
        .map((g) => `<option value="${esc(g.key)}">${esc(g.eventDate)}｜${esc(g.eventName)} (${g.count})</option>`)
        .join('');
      if (state.adminHistoryEventFilter !== 'ALL' && !evtMap.has(state.adminHistoryEventFilter)) {
        state.adminHistoryEventFilter = 'ALL';
      }
      filterSel.value = state.adminHistoryEventFilter;
    }

    const q = (state.adminHistorySearch || '').trim().toLowerCase();
    const filtered = all.filter((h) => {
      if (state.adminHistoryEventFilter !== 'ALL' && `${h.eventDate}__${h.eventName}` !== state.adminHistoryEventFilter) return false;
      if (q) {
        return h.name.toLowerCase().includes(q)
          || (h.idNumber || '').toLowerCase().includes(q)
          || (h.bib || '').toLowerCase().includes(q)
          || (h.eventName || '').toLowerCase().includes(q);
      }
      return true;
    });

    const tbody = $('admin-history-tbody');
    if (!tbody) return;
    if (!filtered.length) {
      tbody.innerHTML = '<tr><td colspan="11" class="py-6 text-center text-slate-500">無符合條件的歷年成績紀錄</td></tr>';
      return;
    }

    tbody.innerHTML = filtered.slice(0, 300).map((h) => `
      <tr class="hover:bg-slate-800/40">
        <td class="py-2 pr-3 font-mono-num text-amber-300">${esc(h.eventDate)}</td>
        <td class="py-2 pr-3 font-bold text-white">${esc(h.eventName)}${h.eventName.includes('八百壯士') ? ' <span class="text-amber-400">⚔️</span>' : ''}</td>
        <td class="py-2 pr-3 text-slate-300">${esc(h.categoryName)}</td>
        <td class="py-2 pr-3 font-mono-num text-slate-300">${esc(h.bib || '--')}</td>
        <td class="py-2 pr-3 font-bold text-white">${esc(h.name)}</td>
        <td class="py-2 pr-3 font-mono-num text-sky-300">${esc(h.idNumber || h.idMasked || '--')}</td>
        <td class="py-2 pr-3 font-mono-num text-emerald-400">${esc(TE.normalizeTimeStr(h.finishTime) || '--')}</td>
        <td class="py-2 pr-3 font-mono-num text-emerald-300">${Number(h.distanceKm || 0) > 0 ? `${Number(h.distanceKm).toFixed(2)} km` : '--'}</td>
        <td class="py-2 pr-3 font-mono-num text-amber-300">${h.overallRank ? `總#${h.overallRank}` : '-'}/${h.categoryRank ? `組#${h.categoryRank}` : '-'}</td>
        <td class="py-2 pr-3 text-slate-300">${esc(h.status)}</td>
        <td class="py-2 text-right">
          <button type="button" data-del-history-id="${esc(h.id)}" class="px-2 py-1 rounded text-xs font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">刪</button>
        </td>
      </tr>
    `).join('');
  }

  function normalizePreviewDate(rawDate) {
    const s = String(rawDate || '').trim().replace(/\//g, '-');
    if (/^\d{4}$/.test(s)) return `${s}-01-01`;
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    return s;
  }

  function clearHistoryImportPreview(showNotice = false) {
    state.historyPreview = {
      items: [],
      skippedCount: 0,
      sourceLabel: '',
      fromText: false,
      eventFilter: 'ALL',
      search: ''
    };
    if ($('history-preview-search')) $('history-preview-search').value = '';
    renderHistoryImportPreview();
    if (showNotice) showToast('已取消歷年成績匯入，未寫入任何資料至資料庫', 'info');
  }

  /**
   * 步驟 1：解析檔案或文字並顯示「預覽場次與明細」，等待管理員審查（絕不直接寫入資料庫）
   */
  function stageHistoryImportPreview(rows, { sourceLabel = '匯入資料', fromText = false } = {}) {
    if (!Array.isArray(rows) || !rows.length) return showToast('未讀取到任何資料列', 'warn');

    const headerAliases = {
      eventDate: ['活動日期', '日期', '年份', 'date', 'eventdate', 'year'],
      eventName: ['活動名稱', '賽事名稱', '活動', 'eventname', 'event'],
      categoryName: ['組別', '分組', '組別名稱', 'category', 'categoryname'],
      bib: ['參賽編號', '編號', '號碼布', 'bib'],
      name: ['姓名', '選手姓名', 'name'],
      idNumber: ['身份證字號', '身分證字號', '身份證', '身分證', 'idnumber', 'id'],
      finishTime: ['完賽時間', '成績', '耗時', '時間', 'finishtime', 'time'],
      distanceKm: ['完賽里程km', '完賽里程(km)', '完賽里程', '里程km', '里程(km)', '里程', '距離', 'distancekm', 'km'],
      laps: ['圈數', '累積圈數', 'laps'],
      overallRank: ['總名次', '總排', 'overallrank'],
      categoryRank: ['組名次', '分組名次', '組排', 'categoryrank'],
      status: ['完賽狀態', '狀態', 'status'],
      note: ['備註', 'note']
    };

    const firstRow = (rows[0] || []).map((c) => String(c ?? '').trim().toLowerCase().replace(/\s+/g, ''));
    const colMap = {};
    let hasHeader = false;
    Object.entries(headerAliases).forEach(([field, keys]) => {
      const idx = firstRow.findIndex((h) => keys.includes(h));
      if (idx >= 0) {
        colMap[field] = idx;
        hasHeader = true;
      }
    });

    const startIdx = hasHeader ? 1 : 0;
    const ownerAdmin = getCurrentAdminUsername() || 'admin';
    const items = [];
    let skippedCount = 0;

    for (let i = startIdx; i < rows.length; i++) {
      const cols = rows[i];
      if (!cols || !cols.some((c) => String(c ?? '').trim() !== '')) continue;
      const getVal = (field, fallbackIdx) => {
        const idx = hasHeader ? colMap[field] : fallbackIdx;
        return idx !== undefined && idx >= 0 ? String(cols[idx] ?? '').trim() : '';
      };
      const eventDate = normalizePreviewDate(getVal('eventDate', 0));
      const eventName = getVal('eventName', 1);
      const name = getVal('name', 4);
      if (!eventName || !name) {
        skippedCount += 1;
        continue;
      }

      const finishTime = TE.normalizeTimeStr(getVal('finishTime', 6));
      const distanceKm = Number(getVal('distanceKm', 7)) || 0;
      const laps = parseInt(getVal('laps', 8), 10) || 0;
      const rawStatus = getVal('status', 11);
      const status = rawStatus || ((finishTime || distanceKm > 0 || laps > 0) ? '完賽' : '未完賽');

      items.push({
        _previewId: `PV_${Date.now()}_${i}`,
        eventDate,
        eventName,
        categoryName: getVal('categoryName', 2) || '一般組',
        bib: getVal('bib', 3),
        name,
        idNumber: getVal('idNumber', 5).toUpperCase(),
        finishTime,
        distanceKm,
        laps,
        overallRank: getVal('overallRank', 9),
        categoryRank: getVal('categoryRank', 10),
        status,
        note: getVal('note', 12),
        ownerAdmin
      });
    }

    if (!items.length) {
      return showToast('未偵測到有效歷年成績（至少需有「活動名稱」與「姓名」）', 'warn');
    }

    state.historyPreview = {
      items,
      skippedCount,
      sourceLabel,
      fromText,
      eventFilter: 'ALL',
      search: ''
    };
    if ($('history-preview-search')) $('history-preview-search').value = '';
    renderHistoryImportPreview();

    const card = $('history-import-preview-card');
    if (card) card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    const evtCount = new Set(items.map((x) => `${x.eventDate}__${x.eventName}`)).size;
    showToast(`🔍 已解析 ${evtCount} 場共 ${items.length} 筆待審查資料，確認無誤後請按「✅ 確認寫入資料庫」`, 'info');
  }

  function renderHistoryImportPreview() {
    const card = $('history-import-preview-card');
    if (!card) return;
    const pv = state.historyPreview;
    const items = pv?.items || [];
    card.classList.toggle('hidden', items.length === 0);
    if (!items.length) return;

    // 對照資料庫中既有歷年場次，標示「新場次」或「合併既有場次」
    const existingEvtKeys = new Set((state.historyRecords || []).map((h) => `${h.eventDate}__${h.eventName}`));

    const evtMap = new Map();
    let missingIdCount = 0;
    let missingDateCount = 0;

    for (const item of items) {
      if (!item.idNumber) missingIdCount += 1;
      if (!item.eventDate) missingDateCount += 1;
      const key = `${item.eventDate}__${item.eventName}`;
      if (!evtMap.has(key)) {
        evtMap.set(key, {
          key,
          eventDate: item.eventDate,
          eventName: item.eventName,
          count: 0,
          finishedCount: 0,
          totalKm: 0,
          existsInDb: existingEvtKeys.has(key)
        });
      }
      const g = evtMap.get(key);
      g.count += 1;
      if (String(item.status || '').includes('完賽') && !String(item.status || '').includes('未')) {
        g.finishedCount += 1;
      }
      g.totalKm += Number(item.distanceKm || 0);
    }

    const groups = Array.from(evtMap.values()).sort((a, b) =>
      String(b.eventDate || '').localeCompare(String(a.eventDate || ''))
      || String(a.eventName || '').localeCompare(String(b.eventName || ''), 'zh-Hant')
    );

    if (pv.eventFilter !== 'ALL' && !evtMap.has(pv.eventFilter)) {
      pv.eventFilter = 'ALL';
    }

    if ($('history-preview-meta')) {
      $('history-preview-meta').textContent = `來源：${pv.sourceLabel} ｜ 辨識 ${groups.length} 場活動、共 ${items.length} 筆明細（尚未存入資料庫）`;
    }
    if ($('history-preview-confirm-count')) $('history-preview-confirm-count').textContent = String(items.length);
    if ($('history-preview-events-count')) $('history-preview-events-count').textContent = String(groups.length);
    if ($('btn-history-preview-all-events')) {
      $('btn-history-preview-all-events').classList.toggle('hidden', pv.eventFilter === 'ALL');
    }

    // 審查提醒訊息
    const warnings = [];
    if (pv.skippedCount > 0) warnings.push(`⚠️ 已自動略過 ${pv.skippedCount} 列缺少「活動名稱」或「姓名」的資料列。`);
    if (missingIdCount > 0) warnings.push(`⚠️ 有 ${missingIdCount} 筆缺少「身份證字號」（仍可寫入，但無法依身份證自動合併個人歷年成績）。`);
    if (missingDateCount > 0) warnings.push(`⚠️ 有 ${missingDateCount} 筆缺少「活動日期」，建議補齊日期後再匯入。`);
    const warnEl = $('history-preview-warnings');
    if (warnEl) {
      warnEl.classList.toggle('hidden', warnings.length === 0);
      warnEl.innerHTML = warnings.map((w) => `<div>${esc(w)}</div>`).join('');
    }

    // 預覽場次卡片
    const evListEl = $('history-preview-events-list');
    if (evListEl) {
      evListEl.innerHTML = groups.map((g) => {
        const is800 = g.eventName.includes('八百壯士');
        const active = pv.eventFilter === g.key;
        return `
          <div class="p-3 rounded-xl border transition flex items-center justify-between gap-2 ${active ? 'bg-amber-500/15 border-amber-400' : 'bg-slate-950 border-slate-800 hover:border-slate-600'}">
            <button type="button" data-preview-evt-filter="${esc(g.key)}" class="flex-1 min-w-0 text-left">
              <div class="flex items-center gap-1.5 flex-wrap">
                <span class="font-mono-num text-[11px] font-bold px-1.5 py-0.5 rounded bg-slate-800 text-amber-300">${esc(g.eventDate || '無日期')}</span>
                <span class="font-bold text-white text-xs sm:text-sm truncate">${esc(g.eventName)}</span>
                ${is800 ? '<span class="px-1.5 py-0.5 rounded text-[10px] font-black bg-amber-500/20 text-amber-300">⚔️ 八百壯士</span>' : ''}
                ${g.existsInDb
                  ? '<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-sky-500/20 text-sky-300">🔄 合併既有場次</span>'
                  : '<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-300">🆕 新場次</span>'}
              </div>
              <div class="text-[11px] text-slate-400 font-mono-num mt-1">
                預覽 ${g.count} 筆（完賽 ${g.finishedCount}）｜ 累計 ${g.totalKm.toFixed(1)} km
              </div>
            </button>
            <button type="button" data-remove-preview-evt="${esc(g.key)}" title="剔除此場次（不匯入）"
              class="px-2 py-1.5 rounded-lg text-[11px] font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300 shrink-0">
              剔除場次
            </button>
          </div>
        `;
      }).join('');
    }

    // 預覽明細篩選與表格渲染
    const q = (pv.search || '').trim().toLowerCase();
    const filtered = items.filter((item) => {
      if (pv.eventFilter !== 'ALL' && `${item.eventDate}__${item.eventName}` !== pv.eventFilter) return false;
      if (q) {
        return item.name.toLowerCase().includes(q)
          || item.eventName.toLowerCase().includes(q)
          || (item.bib || '').toLowerCase().includes(q)
          || (item.idNumber || '').toLowerCase().includes(q)
          || (item.categoryName || '').toLowerCase().includes(q);
      }
      return true;
    });

    if ($('history-preview-rows-label')) {
      $('history-preview-rows-label').textContent = filtered.length === items.length
        ? `共 ${items.length} 筆`
        : `顯示 ${filtered.length} / 共 ${items.length} 筆`;
    }

    const tbody = $('history-preview-tbody');
    if (tbody) {
      if (!filtered.length) {
        tbody.innerHTML = '<tr><td colspan="14" class="py-5 text-center text-slate-500">篩選條件下無預覽明細</td></tr>';
      } else {
        tbody.innerHTML = filtered.slice(0, 500).map((item, idx) => `
          <tr class="hover:bg-slate-800/50">
            <td class="py-2 pl-3 pr-2 font-mono-num text-slate-500">${idx + 1}</td>
            <td class="py-2 pr-3 font-mono-num ${item.eventDate ? 'text-amber-300' : 'text-rose-400'}">${esc(item.eventDate || '缺日期')}</td>
            <td class="py-2 pr-3 font-bold text-white">${esc(item.eventName)}${item.eventName.includes('八百壯士') ? ' <span class="text-amber-400">⚔️</span>' : ''}</td>
            <td class="py-2 pr-3 text-slate-300">${esc(item.categoryName)}</td>
            <td class="py-2 pr-3 font-mono-num text-slate-300">${esc(item.bib || '--')}</td>
            <td class="py-2 pr-3 font-bold text-white">${esc(item.name)}</td>
            <td class="py-2 pr-3 font-mono-num ${item.idNumber ? 'text-sky-300' : 'text-amber-400'}">${esc(item.idNumber || '未填')}</td>
            <td class="py-2 pr-3 font-mono-num text-emerald-400">${esc(item.finishTime || '--')}</td>
            <td class="py-2 pr-3 font-mono-num text-emerald-300">${item.distanceKm > 0 ? `${item.distanceKm.toFixed(2)} km` : '--'}</td>
            <td class="py-2 pr-3 font-mono-num text-purple-300">${item.laps > 0 ? `${item.laps} 圈` : '--'}</td>
            <td class="py-2 pr-3 font-mono-num text-amber-300">${item.overallRank ? `總#${esc(item.overallRank)}` : '-'}/${item.categoryRank ? `組#${esc(item.categoryRank)}` : '-'}</td>
            <td class="py-2 pr-3 text-slate-300">${esc(item.status)}</td>
            <td class="py-2 pr-3 text-slate-400">${esc(item.note || '')}</td>
            <td class="py-2 pr-3 text-right">
              <button type="button" data-remove-preview-id="${esc(item._previewId)}"
                class="px-2 py-0.5 rounded text-[11px] font-bold bg-rose-600/20 hover:bg-rose-600/40 text-rose-300">✕</button>
            </td>
          </tr>
        `).join('');
      }
    }
  }

  /**
   * 步驟 2：管理員審查無誤並點擊「✅ 確認寫入資料庫」後，才正式存入本機與雲端
   */
  async function commitHistoryImportPreview() {
    if (!isCurrentSuperAdmin()) return showToast('僅最高管理權責可匯入歷年活動成績', 'error');
    const pv = state.historyPreview;
    const items = pv?.items || [];
    if (!items.length) return showToast('目前沒有待寫入的預覽資料', 'warn');

    const btn = $('btn-confirm-history-import');
    if (btn) btn.disabled = true;
    try {
      const cleanItems = items.map(({ _previewId, ...rest }) => rest);
      const saved = await DB.upsertHistoryRecordsBulk(cleanItems, false);
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        Sync.pushHistoryRecords(saved, state.session?.auth).then(() => updatePendingSyncBadge());
      }
      if (pv.fromText && $('textarea-history-csv')) {
        $('textarea-history-csv').value = '';
      }
      clearHistoryImportPreview(false);
      await reloadState();
      showToast(`✅ 審查完成！已正式寫入 ${saved.length} 筆歷年活動成績至資料庫`, 'success');
    } catch (err) {
      showToast(`歷年成績寫入失敗：${err.message}`, 'error');
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  /* ========================================================================
   * 分頁 7：🏅 製作證書（僅包含「已結束」與「歷年」活動；官方中英公版 ＋ 上傳底圖套印）
   * ======================================================================== */

  function parseHistoryDurationToMs(str) {
    const s = TE.normalizeTimeStr(str);
    if (!s) return null;
    const m = s.match(/^(\d+):(\d{1,2})(?::(\d{1,2}))?$/);
    if (!m) return null;
    const h = parseInt(m[1], 10) || 0;
    const min = parseInt(m[2], 10) || 0;
    const sec = m[3] !== undefined ? (parseInt(m[3], 10) || 0) : 0;
    return (h * 3600 + min * 60 + sec) * 1000;
  }

  /**
   * 取得後台「製作證書」可選活動清單：僅包含「已結束」系統活動與「歷年成績」活動（進行中與籌備中排除不列）
   */
  function getAdminCertCatalog() {
    const list = [];
    const adminUser = getCurrentAdminUsername();
    const isSuper = isCurrentSuperAdmin();

    // 1. 系統內「已結束」活動（排除進行中 active 與籌備中 upcoming）
    const finishedEvts = getAdminOwnEvents().filter(isEventFinished);
    for (const ev of finishedEvts) {
      list.push({
        sourceKey: `SYS__${ev.eventId}`,
        sourceType: 'system_finished',
        eventDate: ev.date || '',
        eventName: ev.name || '',
        event: ev,
        records: null
      });
    }

    // 2. 歷年成績活動（history_records 依日期＋活動名稱彙整）
    const histMap = new Map();
    for (const h of state.historyRecords || []) {
      if (!h.eventName || !h.name) continue;
      if (!isSuper && h.ownerAdmin && h.ownerAdmin !== adminUser) continue;
      const evDate = h.eventDate || (h.year ? `${h.year}-01-01` : '');
      const key = `IMP__${evDate}__${h.eventName}`;
      if (!histMap.has(key)) {
        histMap.set(key, {
          sourceKey: key,
          sourceType: 'imported',
          eventDate: evDate,
          eventName: h.eventName,
          event: null,
          records: []
        });
      }
      histMap.get(key).records.push(h);
    }
    histMap.forEach((item) => list.push(item));

    return list.sort((a, b) =>
      String(b.eventDate || '').localeCompare(String(a.eventDate || ''))
      || String(a.eventName || '').localeCompare(String(b.eventName || ''), 'zh-Hant')
    );
  }

  /**
   * 依目前選定的已結束／歷年活動，產出標準化的 { sourceKey, sourceType, event, rows }
   */
  function getAdminCertContext() {
    const catalog = getAdminCertCatalog();
    if (!catalog.length) return null;
    const picked = catalog.find((c) => c.sourceKey === state.cert.selectedSourceKey) || catalog[0];
    state.cert.selectedSourceKey = picked.sourceKey;

    if (picked.sourceType === 'system_finished' && picked.event) {
      const ev = picked.event;
      const eventRegs = state.allRegistrations.filter((r) => r.eventId === ev.eventId);
      const eventScans = state.allScans.filter((s) => s.eventId === ev.eventId && !s.deleted);
      const computed = TE.computeEventResults({ event: ev, runners: eventRegs, scans: eventScans });
      return {
        sourceKey: picked.sourceKey,
        sourceType: 'system_finished',
        event: ev,
        rows: computed?.rows || []
      };
    }

    // 歷年成績群組轉為標準證書活動與選手成績結構
    const recs = [...(picked.records || [])].sort((a, b) =>
      (a.overallRank || 99999) - (b.overallRank || 99999)
      || String(a.bib || '').localeCompare(String(b.bib || ''), 'zh-Hant')
    );
    const hasLaps = recs.some((r) => Number(r.laps || 0) > 0);
    const catSet = new Map();
    const rows = recs.map((h, idx) => {
      const catName = String(h.categoryName || '一般組').trim() || '一般組';
      if (!catSet.has(catName)) {
        catSet.set(catName, { id: catName, name: catName, distanceKm: Number(h.distanceKm || 0) });
      }
      const statusStr = String(h.status || '').trim() || '完賽';
      const isUnfinished = statusStr.includes('未') || statusStr.toUpperCase() === 'DNF' || statusStr.toUpperCase() === 'DNS';
      const km = Number(h.distanceKm || 0);
      const laps = Number(h.laps || 0);
      const timeText = TE.normalizeTimeStr(h.finishTime);
      const timeMs = parseHistoryDurationToMs(timeText);
      const completed = !isUnfinished && (statusStr.includes('完賽') || Boolean(timeText) || km > 0 || laps > 0);
      const avgPace = (timeMs && km > 0) ? TE.formatPace(timeMs, km) : '--';
      const bibVal = String(h.bib || `H${String(idx + 1).padStart(3, '0')}`).trim();

      if (hasLaps) {
        return {
          bib: bibVal,
          name: h.name,
          categoryId: catName,
          categoryName: catName,
          lapCount: laps,
          totalDistanceKm: km,
          totalElapsedMs: timeMs,
          overallRank: h.overallRank || null,
          categoryRank: h.categoryRank || null,
          avgPace,
          status: completed ? 'finished' : 'dnf',
          laps: []
        };
      }
      return {
        bib: bibVal,
        name: h.name,
        categoryId: catName,
        categoryName: catName,
        distanceKm: km,
        gunDurationMs: timeMs,
        finishTimeText: timeText,
        overallRank: h.overallRank || null,
        categoryRank: h.categoryRank || null,
        avgPace,
        status: completed ? 'finished' : 'dnf',
        splits: []
      };
    });

    const sampleLapRec = recs.find((r) => Number(r.laps || 0) > 0 && Number(r.distanceKm || 0) > 0);
    const estLapDist = sampleLapRec ? Number((Number(sampleLapRec.distanceKm) / Number(sampleLapRec.laps)).toFixed(2)) : 0;

    const syntheticEvent = {
      eventId: picked.sourceKey,
      name: picked.eventName,
      englishName: '',
      organization: 'Taiwan Ultra Runners',
      date: picked.eventDate,
      location: '',
      eventType: hasLaps ? 'lap_timed' : 'distance',
      lapDistanceKm: estLapDist,
      categories: Array.from(catSet.values())
    };

    return {
      sourceKey: picked.sourceKey,
      sourceType: 'imported',
      event: syntheticEvent,
      rows
    };
  }

  function getCertFilteredRows() {
    const ctx = getAdminCertContext();
    if (!ctx || !ctx.event) return [];
    const evt = ctx.event;
    const isLap = evt.eventType === 'lap_timed';
    const q = (state.cert.search || '').trim().toLowerCase();

    return (ctx.rows || []).filter((r) => {
      if (state.cert.onlyFinished) {
        if (isLap) {
          if (!(Number(r.lapCount || 0) > 0)) return false;
        } else {
          if (r.status !== 'finished') return false;
        }
      }
      if (state.cert.catFilter !== 'ALL' && r.categoryId !== state.cert.catFilter) return false;
      if (q) {
        return String(r.bib || '').toLowerCase().includes(q) || String(r.name || '').toLowerCase().includes(q);
      }
      return true;
    });
  }

  function renderCertFieldChips() {
    const wrap = $('cert-field-list');
    if (!wrap) return;
    wrap.innerHTML = state.cert.layout.map((f) => {
      const isSel = f.key === state.cert.selectedFieldKey;
      return `
        <div class="inline-flex items-center gap-1 px-2 py-1 rounded-lg border text-[11px] font-bold transition ${isSel ? 'bg-amber-500/20 border-amber-400 text-amber-200' : (f.enabled ? 'bg-slate-950 border-slate-700 text-slate-200' : 'bg-slate-950/50 border-slate-800 text-slate-500')}">
          <input type="checkbox" data-cert-field-toggle="${esc(f.key)}" ${f.enabled ? 'checked' : ''} class="accent-sky-500 cursor-pointer" />
          <button type="button" data-cert-field-select="${esc(f.key)}" class="hover:text-white">${esc(f.label)}</button>
        </div>
      `;
    }).join('');
  }

  function syncCertFieldEditor() {
    const f = state.cert.layout.find((x) => x.key === state.cert.selectedFieldKey) || state.cert.layout[0];
    if (!f) return;
    state.cert.selectedFieldKey = f.key;
    if ($('cert-f-title')) $('cert-f-title').textContent = `${f.label}${f.enabled ? '' : '（目前隱藏）'}`;
    if ($('cert-f-x')) $('cert-f-x').value = String(Number(f.x ?? 50).toFixed(1));
    if ($('cert-f-y')) $('cert-f-y').value = String(Number(f.y ?? 50).toFixed(1));
    if ($('cert-f-size')) $('cert-f-size').value = String(Math.round(f.fontSize || 48));
    if ($('cert-f-color')) $('cert-f-color').value = f.color || '#0f172a';
    if ($('cert-f-prefix')) $('cert-f-prefix').value = f.prefix || '';
    if ($('cert-f-align')) $('cert-f-align').value = f.align || 'center';
    if ($('cert-f-bold')) $('cert-f-bold').checked = Boolean(f.bold);
  }

  async function renderCertToCanvas(canvas, row, showGuides = false) {
    const ctx = getAdminCertContext();
    const evt = ctx?.event;
    if (!evt || !Cert || !canvas) return null;
    if (state.cert.mode === 'custom') {
      return Cert.renderCustomBgCertificate(canvas, {
        event: evt,
        row,
        bgImage: state.cert.bgImage,
        layout: state.cert.layout,
        options: { showTurLogo: state.cert.showTurLogo },
        selectedKey: state.cert.selectedFieldKey,
        showGuides
      });
    }
    return Cert.renderOfficialCertificate(canvas, {
      event: evt,
      row,
      options: {
        orientation: state.cert.orientation,
        format: state.cert.format,
        theme: state.cert.theme,
        englishName: state.cert.englishName,
        organization: state.cert.organization,
        showSplits: state.cert.showSplits
      }
    });
  }

  async function renderCertificatePreview() {
    const ctx = getAdminCertContext();
    const evt = ctx?.event;
    const canvas = $('cert-preview-canvas');
    if (!evt || !canvas || !Cert) return;

    const rows = getCertFilteredRows();
    let row = rows.find((r) => r.bib === state.cert.selectedBib) || rows[0] || ctx.rows?.[0] || null;
    if (!row) {
      row = {
        bib: 'A001',
        name: '示範選手',
        categoryName: evt.categories?.[0]?.name || '42K',
        distanceKm: evt.categories?.[0]?.distanceKm || 42.195,
        gunDurationMs: 13525000,
        overallRank: 1,
        categoryRank: 1,
        avgPace: "5'20\"/km",
        status: 'finished'
      };
    }

    const res = await renderCertToCanvas(canvas, row, state.cert.mode === 'custom');
    state.cert.hitBoxes = res?.hitBoxes || [];

    const caption = $('cert-preview-caption');
    if (caption) {
      const fmtLabel = Cert?.OFFICIAL_FORMATS?.[state.cert.format]?.label || '經典燙金典藏式';
      const thmLabel = Cert?.OFFICIAL_THEMES?.[state.cert.theme]?.label || '經典象牙金';
      const modeLabel = state.cert.mode === 'custom'
        ? '自訂底圖套印'
        : `官方公版・${fmtLabel}・${thmLabel}（${state.cert.orientation === 'portrait' ? 'A4 直式' : 'A4 橫式'}）`;
      caption.textContent = `${modeLabel} ｜ 目前預覽：#${row.bib} ${row.name}（${row.categoryName || '--'}）`;
    }
  }

  function renderAdminCertificateTab() {
    const catalog = getAdminCertCatalog();
    const emptyEl = $('admin-cert-empty');
    const contentEl = $('admin-cert-content');
    const hasEvt = catalog.length > 0;
    if (emptyEl) emptyEl.classList.toggle('hidden', hasEvt);
    if (contentEl) contentEl.classList.toggle('hidden', !hasEvt);
    if (!hasEvt) return;

    // 渲染證書活動來源下拉選單（僅包含已結束與歷年活動）
    const ctx = getAdminCertContext();
    const evt = ctx.event;
    const evSel = $('cert-event-select');
    if (evSel) {
      evSel.innerHTML = catalog.map((c) => {
        const tag = c.sourceType === 'imported' ? '📜 [歷年]' : '🔒 [已結束]';
        return `<option value="${esc(c.sourceKey)}">${tag} ${esc(c.eventDate || '--')}｜${esc(c.eventName)}</option>`;
      }).join('');
      evSel.value = ctx.sourceKey;
    }

    // 切換活動時自動載入該活動英文名、主辦單位與已儲存的自訂底圖設定
    if (state.cert.loadedEventId !== ctx.sourceKey) {
      state.cert.loadedEventId = ctx.sourceKey;
      state.cert.catFilter = 'ALL';
      state.cert.search = '';
      state.cert.selectedBib = '';
      state.cert.englishName = evt.englishName || '';
      state.cert.organization = evt.organization || 'Taiwan Ultra Runners';
      state.cert.bgDataUrl = '';
      state.cert.bgImage = null;
      state.cert.showTurLogo = false;
      state.cert.layout = Cert ? Cert.getDefaultCustomLayout() : [];
      state.cert.selectedFieldKey = 'name';

      const savedCfg = Cert ? Cert.loadSavedCustomConfig(evt.eventId) : null;
      if (savedCfg) {
        if (Array.isArray(savedCfg.layout) && savedCfg.layout.length) state.cert.layout = savedCfg.layout;
        if (typeof savedCfg.showTurLogo === 'boolean') state.cert.showTurLogo = savedCfg.showTurLogo;
        if (savedCfg.bgDataUrl) {
          state.cert.bgDataUrl = savedCfg.bgDataUrl;
          Cert.loadImage(savedCfg.bgDataUrl).then((img) => {
            if (state.cert.loadedEventId === ctx.sourceKey) {
              state.cert.bgImage = img;
              if ($('cert-bg-status')) $('cert-bg-status').textContent = img ? `✅ 已載入儲存底圖 (${img.naturalWidth}×${img.naturalHeight})` : '尚未上傳底圖';
              if ($('btn-cert-clear-bg')) $('btn-cert-clear-bg').classList.toggle('hidden', !img);
              if (state.adminZone === 'closed' && state.adminTab === 'certificate') renderCertificatePreview();
            }
          });
        }
      }
      if ($('cert-en-name')) $('cert-en-name').value = state.cert.englishName;
      if ($('cert-org-name')) $('cert-org-name').value = state.cert.organization;
      if ($('cert-show-splits')) $('cert-show-splits').checked = state.cert.showSplits;
      if ($('cert-orientation')) $('cert-orientation').value = state.cert.orientation;
      if ($('cert-format')) $('cert-format').value = state.cert.format;
      if ($('cert-theme')) $('cert-theme').value = state.cert.theme;
      if ($('cert-custom-show-logo')) $('cert-custom-show-logo').checked = state.cert.showTurLogo;
      if ($('cert-bg-status') && !state.cert.bgDataUrl) $('cert-bg-status').textContent = '尚未上傳底圖（支援 A4 橫式或直式高解析圖檔）';
      if ($('btn-cert-clear-bg')) $('btn-cert-clear-bg').classList.toggle('hidden', !state.cert.bgDataUrl);
    }

    if ($('cert-event-name-badge')) $('cert-event-name-badge').textContent = evt.name;
    if ($('cert-event-date-badge')) $('cert-event-date-badge').textContent = `📅 活動日期：${evt.date || '--'}`;

    // 模式按鈕與對應設定面板
    document.querySelectorAll('.cert-mode-btn').forEach((btn) => {
      const active = btn.dataset.certMode === state.cert.mode;
      btn.className = active
        ? 'cert-mode-btn px-3.5 py-2 rounded-lg text-xs sm:text-sm font-black bg-amber-500 text-slate-950 transition'
        : 'cert-mode-btn px-3.5 py-2 rounded-lg text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition';
    });
    if ($('cert-panel-official')) $('cert-panel-official').classList.toggle('hidden', state.cert.mode !== 'official');
    if ($('cert-panel-custom')) $('cert-panel-custom').classList.toggle('hidden', state.cert.mode !== 'custom');
    if ($('cert-drag-hint')) $('cert-drag-hint').classList.toggle('hidden', state.cert.mode !== 'custom');

    // 組別選單
    const catSel = $('cert-cat-filter');
    if (catSel) {
      catSel.innerHTML = '<option value="ALL">全部組別</option>' +
        (evt.categories || []).map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
      catSel.value = state.cert.catFilter;
    }
    if ($('cert-only-finished')) $('cert-only-finished').checked = state.cert.onlyFinished;

    // 篩選選手清單
    const rows = getCertFilteredRows();
    if (!rows.some((r) => r.bib === state.cert.selectedBib)) {
      state.cert.selectedBib = rows[0]?.bib || '';
    }
    if ($('cert-runner-count')) $('cert-runner-count').textContent = `${rows.length} 人`;
    if ($('cert-batch-count')) $('cert-batch-count').textContent = String(rows.length);

    const listEl = $('cert-runner-list');
    if (listEl) {
      const isLap = evt.eventType === 'lap_timed';
      listEl.innerHTML = rows.length === 0
        ? '<div class="p-4 rounded-xl bg-slate-950 border border-dashed border-slate-800 text-center text-xs text-slate-500">目前無符合條件的選手（可取消勾選「僅列出已完賽」預覽樣式）</div>'
        : rows.map((r) => {
            const active = r.bib === state.cert.selectedBib;
            const resText = isLap
              ? `${Number(r.totalDistanceKm || 0).toFixed(2)} km (${r.lapCount}圈)`
              : (r.gunDurationMs !== null && r.gunDurationMs !== undefined ? TE.formatDuration(r.gunDurationMs, false) : (r.finishTimeText || '--'));
            return `
              <button type="button" data-cert-bib="${esc(r.bib)}"
                class="w-full text-left px-3 py-2 rounded-xl border transition flex items-center justify-between gap-2 ${active ? 'bg-amber-500/15 border-amber-400 text-white' : 'bg-slate-950 border-slate-800 hover:border-slate-700 text-slate-300'}">
                <div class="min-w-0">
                  <span class="font-mono-num text-xs font-black text-amber-300 mr-1.5">#${esc(r.bib)}</span>
                  <span class="text-xs font-bold">${esc(r.name)}</span>
                  <span class="text-[10px] text-slate-400 ml-1">${esc(r.categoryName)}</span>
                </div>
                <span class="font-mono-num text-[11px] font-bold text-emerald-400 shrink-0">${esc(resText)}</span>
              </button>
            `;
          }).join('');
    }

    renderCertFieldChips();
    syncCertFieldEditor();

    // 僅在目前顯示「製作證書」分頁時即時重繪 Canvas，避免背景分頁多餘計算
    if (state.adminZone === 'closed' && state.adminTab === 'certificate') {
      renderCertificatePreview();
    }
  }

  /* ========================================================================
   * 已結案專區 1：📊 匯出 Excel / DUV 成績（僅限已結束之結案活動）
   * ======================================================================== */
  function getAdminFinishedEvents() {
    return getAdminOwnEvents()
      .filter(isEventFinished)
      .sort((a, b) =>
        String(b.date || '').localeCompare(String(a.date || ''))
        || String(a.name || '').localeCompare(String(b.name || ''), 'zh-Hant')
      );
  }

  function getSelectedExportContext() {
    const finished = getAdminFinishedEvents();
    if (!finished.length) return null;
    const evt = finished.find((e) => e.eventId === state.exportSelectedEventId) || finished[0];
    state.exportSelectedEventId = evt.eventId;
    const eventRegs = state.allRegistrations.filter((r) => r.eventId === evt.eventId);
    const eventScans = state.allScans.filter((s) => s.eventId === evt.eventId && !s.deleted);
    const computed = TE.computeEventResults({ event: evt, runners: eventRegs, scans: eventScans });
    return { event: evt, computed, scans: eventScans };
  }

  function renderAdminExportTab() {
    const finished = getAdminFinishedEvents();
    const emptyEl = $('admin-export-empty');
    const contentEl = $('admin-export-content');
    const sel = $('export-event-select');
    const selEmpty = $('export-event-select-empty');
    const hasAny = finished.length > 0;
    if (emptyEl) emptyEl.classList.toggle('hidden', hasAny);
    if (contentEl) contentEl.classList.toggle('hidden', !hasAny);
    if (sel) sel.classList.toggle('hidden', !hasAny);
    if (selEmpty) selEmpty.classList.toggle('hidden', hasAny);
    if (!hasAny) {
      if (sel) sel.innerHTML = '';
      return;
    }

    const ctx = getSelectedExportContext();
    const evt = ctx.event;
    const isLap = evt.eventType === 'lap_timed';

    if (sel) {
      sel.innerHTML = finished.map((ev) => {
        const icon = ev.eventType === 'lap_timed' ? '🔁' : '🏃';
        return `<option value="${esc(ev.eventId)}">🔒 [已結案] ${esc(ev.date || '--')}｜${icon} ${esc(ev.name)}</option>`;
      }).join('');
      sel.value = evt.eventId;
    }

    const summaryEl = $('export-event-summary');
    if (summaryEl) {
      const s = ctx.computed.summary || {};
      const metricLabel = isLap ? '有效圈數選手' : '完賽人數';
      const metricVal = isLap ? (s.activeWithLaps ?? 0) : (s.finishedCount ?? 0);
      summaryEl.innerHTML = `
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">活動日期 / 賽制</div>
          <div class="text-sm font-black font-mono-num text-amber-300 mt-1">${esc(evt.date || '--')}</div>
          <div class="text-[11px] text-slate-400 mt-0.5">${isLap ? `繞圈賽 (${evt.lapDistanceKm || 2.54} km/圈)` : '一般距離賽'}</div>
        </div>
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">報名 / 報到</div>
          <div class="text-lg font-black font-mono-num text-white mt-0.5">${s.totalRunners ?? 0} <span class="text-xs text-sky-300">/ 報到 ${s.checkedInCount ?? 0}</span></div>
        </div>
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">${metricLabel}</div>
          <div class="text-lg font-black font-mono-num text-emerald-400 mt-0.5">${metricVal} 人</div>
        </div>
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">有效感應紀錄</div>
          <div class="text-lg font-black font-mono-num text-sky-300 mt-0.5">${ctx.scans.length} 筆</div>
        </div>
      `;
    }

    const warnEl = $('export-duv-warn-banner');
    if (warnEl && window.ExcelExporter?.buildDuvCategoryBlocks) {
      const preview = Excel.buildDuvCategoryBlocks({
        event: evt,
        computed: ctx.computed,
        options: {
          organization: evt.organization || 'Taiwan Ultra Runners',
          eventName: evt.englishName || evt.name || '',
          defaultCity: 'Taipei',
          defaultCountry: 'Taiwan',
          lapMetric: isLap ? 'distance' : 'time'
        }
      });
      const warns = [];
      if (preview.missingDobCount > 0) warns.push(`${preview.missingDobCount} 位缺出生年月日 (DOB)`);
      if (preview.missingGenderCount > 0) warns.push(`${preview.missingGenderCount} 位缺性別 (Gender)`);
      if (!evt.englishName) warns.push('尚未填寫英文活動名稱（匯出 DUV 時可於設定視窗補填）');
      warnEl.classList.toggle('hidden', warns.length === 0);
      if (warns.length) warnEl.textContent = `⚠️ DUV 匯出前提醒：${warns.join('；')}`;
    }
  }

  function updateBackupRestoreModeUi() {
    const mode = document.querySelector('input[name="backup-restore-mode"]:checked')?.value || 'merge';
    const pickBtn = $('btn-backup-restore-pick');
    if (!pickBtn) return;
    const isReplace = mode === 'replace';
    pickBtn.className = isReplace
      ? 'cursor-pointer w-full py-3 px-4 rounded-xl text-sm font-black bg-amber-500 hover:bg-amber-400 text-slate-950 transition shadow text-center block'
      : 'cursor-pointer w-full py-3 px-4 rounded-xl text-sm font-black bg-indigo-600 hover:bg-indigo-500 text-white transition shadow text-center block';
    const inputEl = $('input-backup-file');
    pickBtn.childNodes[0].textContent = isReplace
      ? '♻️ 選擇備份檔 (.json) 並覆蓋還原'
      : '📂 選擇備份檔 (.json) 並合併還原';
    if (inputEl && !pickBtn.contains(inputEl)) pickBtn.appendChild(inputEl);
  }

  function showBackupRestoreResult(message, ok = true) {
    const box = $('backup-restore-result');
    if (!box) return;
    box.classList.remove('hidden');
    box.className = ok
      ? 'rounded-xl border border-emerald-500/45 bg-emerald-950/40 text-emerald-200 px-3.5 py-2.5 text-xs font-bold'
      : 'rounded-xl border border-rose-500/45 bg-rose-950/40 text-rose-200 px-3.5 py-2.5 text-xs font-bold';
    box.textContent = `${TE.formatClockSec(Date.now())} ｜ ${message}`;
  }

  function renderAdminSyncTab() {
    const settings = DB.getDeviceSettings();
    const isSuper = isCurrentSuperAdmin();
    const allEvents = state.platform?.events || [];
    const ownEvents = isSuper ? allEvents : getAdminOwnEvents();
    const finishedCount = ownEvents.filter(isEventFinished).length;
    const activeScans = (state.allScans || []).filter((s) => !s.deleted).length;

    // 渲染本機資料庫即時統計
    const statsBar = $('db-stats-bar');
    if (statsBar) {
      statsBar.innerHTML = `
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">活動總數</div>
          <div class="text-lg font-black font-mono-num text-white mt-0.5">${ownEvents.length} <span class="text-[11px] text-emerald-400 font-bold">(${finishedCount} 結案)</span></div>
        </div>
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">共用晶片</div>
          <div class="text-lg font-black font-mono-num text-sky-300 mt-0.5">${(state.chips || []).length}</div>
        </div>
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">報名明細</div>
          <div class="text-lg font-black font-mono-num text-amber-300 mt-0.5">${(state.allRegistrations || []).length}</div>
        </div>
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">有效感應</div>
          <div class="text-lg font-black font-mono-num text-emerald-400 mt-0.5">${activeScans}</div>
        </div>
        <div class="rounded-xl bg-slate-950 border border-slate-800 p-3">
          <div class="text-[11px] text-slate-400 font-bold">歷年紀錄</div>
          <div class="text-lg font-black font-mono-num text-purple-300 mt-0.5">${(state.historyRecords || []).length}</div>
        </div>
      `;
    }

    const summaryEl = $('backup-summary-text');
    if (summaryEl) {
      summaryEl.textContent = `${allEvents.length} 活動・${(state.chips || []).length} 晶片・${(state.allRegistrations || []).length} 報名・${activeScans} 感應・${(state.historyRecords || []).length} 歷年`;
    }

    document.querySelectorAll('input[name="sync-mode"]').forEach((radio) => {
      radio.checked = radio.value === settings.mode;
    });
    $('input-gas-url').value = settings.gasUrl || '';

    // 即時模式白話提示框
    const callout = $('sync-mode-callout');
    const calloutTitle = $('sync-mode-callout-title');
    const isOnlineMode = settings.mode === 'online' && Boolean(settings.gasUrl);
    if (callout && calloutTitle) {
      if (isOnlineMode) {
        callout.className = 'rounded-xl border border-sky-500/35 bg-sky-950/30 px-3.5 py-2.5 text-xs space-y-1';
        calloutTitle.className = 'font-black text-sky-300';
        calloutTitle.textContent = navigator.onLine
          ? '🌐 目前為「雲端連線模式」：有網路時會自動與 Google 試算表保持同步'
          : '⚠️ 目前為「雲端連線模式（無網路）」：異動將先安全暫存本機，連網後自動同步';
      } else {
        callout.className = 'rounded-xl border border-emerald-500/35 bg-emerald-950/30 px-3.5 py-2.5 text-xs space-y-1';
        calloutTitle.className = 'font-black text-emerald-300';
        calloutTitle.textContent = '📴 目前為「單機離線模式」：所有操作僅存本機，不自動上傳雲端（欲同步請切回雲端連線並按雙向同步）';
      }
    }

    // 資料庫專區「備份與還原」及危險重置項目僅限最高管理權責顯示，活動承辦人一律隱藏防誤觸
    const backupSection = $('admin-backup-restore-section');
    if (backupSection) backupSection.classList.toggle('hidden', !isSuper);
    const maintCard = $('admin-maintenance-card');
    if (maintCard) maintCard.classList.toggle('hidden', !isSuper);
    const pushPlatformBtn = $('btn-sync-push-platform');
    if (pushPlatformBtn) pushPlatformBtn.classList.toggle('hidden', !isSuper);
    const dbSubtitle = $('admin-zone-db-subtitle');
    if (dbSubtitle) {
      dbSubtitle.textContent = isSuper ? '備份與還原・雲端同步・測試資料與重置' : '雲端資料庫同步與分享連結';
    }
    const dbOverviewTitle = $('db-overview-title');
    if (dbOverviewTitle) {
      dbOverviewTitle.textContent = isSuper ? '🗄️ 本機資料庫總覽' : '🗄️ 本機資料庫總覽與雲端同步';
    }
    updateBackupRestoreModeUi();
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
    if (isEventFinished(events[idx])) {
      showToast('此活動狀態為「已結束」，已鎖定禁止修改資料', 'error');
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
      if (!silent || res.platformSynced || res.uploadedScans > 0 || res.uploadedChips > 0 || res.uploadedRegs > 0 || res.purgedRegistrations > 0 || res.purgedScans > 0) {
        const purgeText = (res.purgedRegistrations || res.purgedScans)
          ? `、雲端清理已刪活動明細（報名 ${res.purgedRegistrations || 0}、感應 ${res.purgedScans || 0} 筆）`
          : '';
        showToast(
          `☁️ 雙向同步完成！${res.platformSynced ? '活動設定已上傳、' : ''}晶片 ${res.uploadedChips}、報名 ${res.uploadedRegs}、感應 ${res.uploadedScans} 筆${purgeText}（共 ${evtCount} 場活動）` +
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
    const filename = `${prefix}_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    const summary = `已下載備份「${filename}」（${backup.platform.events.length} 場活動、${backup.chips.length} 張晶片、${backup.registrations.length} 筆報名、${backup.scans.length} 筆感應、${backup.historyRecords.length} 筆歷年）`;
    showToast(summary, 'success');
    showBackupRestoreResult(`✅ ${summary}`, true);
  }

  async function importRegistrationsRows(rows) {
    const evt = getActiveEvent();
    if (!evt) return showToast('請先建立或選取活動', 'warn');
    if (!Array.isArray(rows) || !rows.length) return showToast('未讀取到任何名單資料', 'warn');

    const defaultCat = evt.categories?.[0]?.id || '10K';
    const headerAliases = {
      bib: ['參賽編號', '編號', '號碼布', 'bib', 'bibno', 'bibno.'],
      name: ['姓名', '選手姓名', 'name', 'fullname'],
      lastName: ['姓', '姓氏', 'lastname'],
      firstName: ['名', '名字', 'firstname', '1stname'],
      idNumber: ['身份證字號', '身分證字號', '身份證', '身分證', 'idnumber', 'id'],
      categoryId: ['分組代碼', '分組', '組別', '組別代碼', 'category', 'categoryid'],
      gender: ['性別', 'gender', 'sex'],
      dob: ['出生年月日', '生日', '出生日期', 'dob', 'birthday', 'birthdate'],
      city: ['城市', '縣市', 'city'],
      country: ['國家', '國籍', 'country']
    };

    const firstRow = (rows[0] || []).map((c) => String(c ?? '').trim().toLowerCase().replace(/\s+/g, ''));
    const colMap = {};
    let hasHeader = false;
    Object.entries(headerAliases).forEach(([field, keys]) => {
      const idx = firstRow.findIndex((h) => keys.includes(h));
      if (idx >= 0) {
        colMap[field] = idx;
        hasHeader = true;
      }
    });

    const startIdx = hasHeader ? 1 : 0;
    const valid = [];

    for (let i = startIdx; i < rows.length; i++) {
      const cols = rows[i];
      if (!cols || !cols.some((c) => String(c ?? '').trim() !== '')) continue;
      const getVal = (field, fallbackIdx) => {
        const idx = hasHeader ? colMap[field] : fallbackIdx;
        return idx !== undefined && idx >= 0 ? String(cols[idx] ?? '').trim() : '';
      };

      const bib = getVal('bib', 0);
      const lastName = getVal('lastName', -1);
      const firstName = getVal('firstName', -1);
      const rawName = getVal('name', 1);
      const name = rawName || [lastName, firstName].filter(Boolean).join('');
      const idNumber = getVal('idNumber', 2).toUpperCase();
      if (!bib || !name || bib === '參賽編號' || bib.toLowerCase() === 'bib') continue;

      const rec = {
        eventId: evt.eventId,
        bib,
        name,
        idNumber,
        categoryId: getVal('categoryId', 3) || defaultCat
      };
      const gender = getVal('gender', 4);
      const dob = getVal('dob', 5);
      const city = getVal('city', 6);
      const country = getVal('country', 7);
      if (gender) rec.gender = gender;
      if (dob) rec.dob = dob;
      if (city) rec.city = city;
      if (country) rec.country = country;
      if (lastName) rec.lastName = lastName;
      if (firstName) rec.firstName = firstName;
      valid.push(rec);
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
    $('admin-reg-gender').value = existing?.gender || '';
    $('admin-reg-dob').value = existing?.dob || '';
    $('admin-reg-city').value = existing?.city || '';
    $('admin-reg-country').value = existing?.country || '';
    $('admin-reg-lastname').value = existing?.lastName || '';
    $('admin-reg-firstname').value = existing?.firstName || '';
    $('admin-reg-cat').innerHTML = (evt.categories || []).map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join('');
    if (existing?.categoryId) $('admin-reg-cat').value = existing.categoryId;
    $('btn-admin-delete-reg').classList.toggle('hidden', !existing);
    $('modal-edit-reg').classList.remove('hidden');
  }

  function openDuvExportModal() {
    const ctx = getSelectedExportContext();
    if (!ctx || !ctx.event || !ctx.computed) {
      return showToast('尚無已結案活動可匯出 DUV（請先將活動狀態設為「已結束」）', 'warn');
    }
    const evt = ctx.event;
    const computed = ctx.computed;

    const isLap = evt.eventType === 'lap_timed';
    $('duv-org').value = evt.organization || 'Taiwan Ultra Runners';
    $('duv-event-name').value = evt.englishName || evt.name || '';
    $('duv-default-city').value = 'Taipei';
    $('duv-default-country').value = 'Taiwan';
    $('duv-lap-metric-wrap').classList.toggle('hidden', !isLap);
    if (isLap) $('duv-lap-metric').value = 'distance';

    const preview = Excel.buildDuvCategoryBlocks({
      event: evt,
      computed,
      options: {
        organization: $('duv-org').value,
        eventName: $('duv-event-name').value,
        defaultCity: $('duv-default-city').value,
        defaultCountry: $('duv-default-country').value,
        lapMetric: isLap ? $('duv-lap-metric').value : 'time'
      }
    });

    const catDesc = preview.blocks.length
      ? preview.blocks.map((b) => `<span class="inline-block px-2 py-0.5 rounded bg-slate-800 text-amber-300 font-mono-num mr-1">${esc(b.categoryTitle)}：${b.count} 人</span>`).join('')
      : '<span class="text-rose-400">目前尚無完賽（或有效圈數）選手</span>';

    $('duv-preview-summary').innerHTML = `
      <div class="font-bold text-white">活動日期：${esc(preview.duvDate || evt.date || '--')} ｜ 共 <span class="text-emerald-400 font-mono-num">${preview.totalFinishers}</span> 位完賽選手</div>
      <div class="pt-1">${catDesc}</div>
    `;

    const warnEl = $('duv-preview-warn');
    const warns = [];
    if (preview.missingDobCount > 0) warns.push(`${preview.missingDobCount} 位缺出生年月日 (DOB)`);
    if (preview.missingGenderCount > 0) warns.push(`${preview.missingGenderCount} 位缺性別 (Gender)`);
    if (warns.length > 0 && preview.totalFinishers > 0) {
      warnEl.textContent = `⚠️ 提醒：${warns.join('、')}（仍可直接匯出，亦可先至「👥 名單」補齊後再匯出）`;
      warnEl.classList.remove('hidden');
    } else {
      warnEl.classList.add('hidden');
    }

    $('modal-duv-export').classList.remove('hidden');
  }

  /* ========================================================================
   * 事件綁定
   * ======================================================================== */
  function bindEvents() {
    // 活動切換（大廳 & 後台）
    const onEventChange = async (eventId) => {
      const target = (state.platform?.events || []).find((x) => x.eventId === eventId);
      if (isEventFinished(target)) {
        return showToast('已結束活動已鎖定保護，不可切換編輯', 'warn');
      }
      state.selectedEventId = eventId;
      state.resultsCatFilter = 'ALL';
      DB.saveDeviceSettings({ selectedEventId: eventId });
      await reloadState();
    };
    $('portal-event-select').addEventListener('change', (e) => onEventChange(e.target.value));
    $('admin-event-select').addEventListener('change', (e) => {
      onEventChange(e.target.value);
      if (!['events', 'roster', 'results', 'race', 'staff'].includes(state.adminTab)) {
        switchAdminTab('events');
      }
    });

    // 首頁字體大小即時切換按鈕
    document.querySelectorAll('.font-scale-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        applyFontScale(btn.dataset.fontScale, { save: true, toast: true });
      });
    });

    // 大廳導覽：若已登入同身分則直接進入作業區，無須重輸帳密（直到登出才換人）
    $('btn-portal-staff').addEventListener('click', () => {
      if (state.session?.mode === 'staff') {
        switchView('staff-workspace');
      } else if (state.session?.mode === 'admin') {
        showToast('目前已登入後台管理帳號，如需切換現場作業員請先點右上角「登出 / 換人」', 'warn');
      } else {
        switchView('staff-login');
      }
    });
    $('btn-portal-admin').addEventListener('click', () => {
      if (state.session?.mode === 'admin') {
        if (!state.adminZone) state.adminZone = 'ops';
        switchView('admin-workspace');
      } else if (state.session?.mode === 'staff') {
        showToast('目前已登入作業員帳號，如需進入後台管理請先點右上角「登出 / 換人」', 'warn');
      } else {
        switchView('admin-login');
      }
    });
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
    $('header-brand-home').addEventListener('click', () => switchView('portal'));
    document.querySelectorAll('[data-goto]').forEach((b) => b.addEventListener('click', () => switchView(b.dataset.goto)));
    $('btn-quick-sync').addEventListener('click', () => handleUploadPendingSync(false));

    // 登入表單：選擇活動即時更新各作業組別卡片，點選組別帶入帳號並聚焦密碼欄
    $('staff-login-event').addEventListener('change', (e) => {
      state.selectedEventId = e.target.value;
      state.staffLoginSelectedGroup = '';
      $('staff-login-user').value = '';
      $('staff-login-pass').value = '';
      renderStaffLoginGroups();
    });

    $('staff-login-group-grid').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-pick-staff-group]');
      if (!btn) return;
      const username = btn.dataset.pickStaffGroup;
      state.staffLoginSelectedGroup = username;
      $('staff-login-user').value = username;
      const passEl = $('staff-login-pass');
      passEl.removeAttribute('readonly');
      passEl.value = '';
      renderStaffLoginGroups();
      passEl.focus();
    });

    const fillTestPassBtn = $('btn-staff-fill-test-pass');
    if (fillTestPassBtn) {
      fillTestPassBtn.addEventListener('click', () => {
        if (!$('staff-login-user').value.trim()) {
          const firstBtn = $('staff-login-group-grid')?.querySelector('[data-pick-staff-group]');
          if (firstBtn) {
            const u = firstBtn.dataset.pickStaffGroup;
            state.staffLoginSelectedGroup = u;
            $('staff-login-user').value = u;
            renderStaffLoginGroups();
          }
        }
        $('staff-login-pass').value = '1111';
        $('staff-login-pass').focus();
      });
    }

    $('form-staff-login').addEventListener('submit', (e) => {
      e.preventDefault();
      handleLogin('staff', $('staff-login-user').value, $('staff-login-pass').value, $('staff-login-event').value);
    });
    // 後台管理登入：切換最高管理權責／活動承辦人身分與快選帳號
    $('btn-staff-goto-admin').addEventListener('click', () => {
      state.adminLoginRole = 'super';
      state.adminLoginSelectedUser = '';
      clearLoginInputs();
    });
    $('admin-login-role-selector').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-admin-login-role]');
      if (!btn) return;
      state.adminLoginRole = btn.dataset.adminLoginRole === 'coordinator' ? 'coordinator' : 'super';
      state.adminLoginSelectedUser = '';
      clearLoginInputs();
      renderAdminLoginView();
    });
    $('admin-login-account-grid').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-pick-admin-user]');
      if (!btn) return;
      const username = btn.dataset.pickAdminUser;
      state.adminLoginSelectedUser = username;
      const userEl = $('admin-login-user');
      const passEl = $('admin-login-pass');
      userEl.removeAttribute('readonly');
      userEl.value = username;
      passEl.removeAttribute('readonly');
      passEl.value = '';
      renderAdminLoginView();
      passEl.focus();
    });
    $('form-admin-login').addEventListener('submit', (e) => {
      e.preventDefault();
      handleLogin('admin', $('admin-login-user').value, $('admin-login-pass').value, '', state.adminLoginRole);
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
        state.session.lockedEventId = 'EVT-2026-800-EAST';
        DB.saveSession(state.session);
      }
      await reloadState();
      showToast('已載入 5 場全情境共用晶片測試資料！可直接模擬刷卡體驗比對機制', 'success');
    });

    $('btn-mock-scan').addEventListener('click', async () => {
      const evt = getActiveEvent();
      const evtRegs = evt ? state.allRegistrations.filter((r) => r.eventId === evt.eventId) : [];
      if (!evtRegs.length) {
        return showToast('本場活動尚無報名選手，請先於後台匯入名單或點右側「載入模擬資料」', 'warn');
      }

      const chosenKey = $('mock-chip-select').value;
      let reg = null;
      if (chosenKey) {
        if (chosenKey.startsWith('BIB:')) {
          const targetBib = chosenKey.slice(4);
          reg = evtRegs.find((r) => String(r.bib || '') === targetBib) || null;
        } else {
          reg = evtRegs.find((r) => String(r.idNumber || '').trim().toUpperCase() === chosenKey) || null;
        }
      }
      if (!reg) {
        reg = evtRegs[Math.floor(Math.random() * evtRegs.length)];
      }
      if (!reg) return showToast('本場活動尚無報名選手可供模擬感應', 'warn');

      const matchedChip = reg.idNumber
        ? state.chips.find((c) => String(c.idNumber || '').trim().toUpperCase() === String(reg.idNumber || '').trim().toUpperCase())
        : null;

      // 模擬晶片僅使用本場活動選手資料進行感應比對
      await processIncomingScan({
        name: reg.name,
        idNumber: reg.idNumber || '',
        bib: reg.idNumber ? '' : (reg.bib || ''),
        nfcUid: matchedChip?.nfcUid || '',
        source: 'mock',
        note: '模擬本場選手晶片刷過'
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
    document.querySelectorAll('.admin-zone-nav-btn').forEach((b) => {
      b.addEventListener('click', () => switchAdminZone(b.dataset.adminZone, { toggle: true }));
    });
    document.querySelectorAll('[data-admin-zone-collapse]').forEach((b) => {
      b.addEventListener('click', () => switchAdminZone(null));
    });
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
      const reopenBtn = e.target.closest('[data-reopen-event]');
      if (reopenBtn) {
        const eid = reopenBtn.dataset.reopenEvent;
        const target = (state.platform.events || []).find((x) => x.eventId === eid);
        if (!isEventOwnedByCurrentAdmin(target)) return showToast('無權限操作其他管理者的活動', 'error');
        if (!confirm(`確定要將已結束活動「${target?.name || eid}」重新開啟為「進行中」嗎？\n（重新開啟後將恢復顯示於下拉選單與其他管理分頁，可進行編輯）`)) return;
        const nextEvents = (state.platform.events || []).map((x) =>
          x.eventId === eid ? { ...x, status: 'active', updatedAt: Date.now() } : x
        );
        const updated = await DB.savePlatform({ ...state.platform, events: nextEvents });
        state.selectedEventId = eid;
        DB.saveDeviceSettings({ selectedEventId: eid });
        if (Sync.isOnlineEnabled() && navigator.onLine) {
          try { await Sync.pushPlatform(updated, state.session.auth); } catch (err) { showToast(`本機已更新，雲端同步待重試：${err.message}`, 'warn'); }
        }
        await reloadState();
        showToast(`🔓 已將活動「${target?.name || eid}」改回「進行中」`, 'success');
        return;
      }

      const selBtn = e.target.closest('[data-select-event]');
      if (selBtn) {
        const target = (state.platform.events || []).find((x) => x.eventId === selBtn.dataset.selectEvent);
        if (!isEventOwnedByCurrentAdmin(target)) return showToast('無權限切換至其他管理者的活動', 'error');
        if (isEventFinished(target)) return showToast('已結束活動已鎖定保護，請先點擊「重新開啟」再管理', 'warn');
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

        const regCount = state.allRegistrations.filter((r) => r.eventId === eid).length;
        const scanCount = state.allScans.filter((s) => s.eventId === eid).length;
        const confirmMsg =
          `確定要刪除活動「${target?.name || eid}」嗎？\n\n` +
          `⚠️ 將一併永久刪除資料庫中該活動的明細資料（報名名單 ${regCount} 筆、感應資料 ${scanCount} 筆），以減輕資料庫負荷。\n` +
          `（跨活動共用晶片主檔不會受到影響，此操作無法復原）`;
        if (!confirm(confirmMsg)) return;

        // 1. 刪除本機 IndexedDB 中該活動的所有參賽明細與感應資料
        const localPurge = await DB.deleteEventData(eid);

        // 2. 更新活動清單與預設顯示活動
        const remain = (state.platform.events || []).filter((x) => x.eventId !== eid);
        const remainOwn = remain.filter((x) => isEventOwnedByCurrentAdmin(x));
        const updated = await DB.savePlatform({
          ...state.platform,
          events: remain,
          activeEventId: state.platform.activeEventId === eid ? (remainOwn[0]?.eventId || remain[0]?.eventId || '') : state.platform.activeEventId
        });
        if (state.selectedEventId === eid) {
          state.selectedEventId = remainOwn[0]?.eventId || '';
          DB.saveDeviceSettings({ selectedEventId: state.selectedEventId });
        }

        // 3. 若有連網，立即同步至雲端 GAS 並連動清除試算表中該活動的 Registrations 與 Scans 明細列
        let cloudMsg = '';
        if (Sync.isOnlineEnabled() && navigator.onLine) {
          try {
            const res = await Sync.pushPlatform(updated, state.session.auth, { deletedEventIds: [eid] });
            if (res && res.purge) {
              cloudMsg = `；雲端已同步清除（報名 ${res.purge.deletedRegistrations} 筆、感應 ${res.purge.deletedScans} 筆）`;
            } else {
              cloudMsg = '；雲端活動已移除（請重新部署新版 gas/Code.gs 以連動清除雲端明細）';
            }
          } catch (err) {
            cloudMsg = `；雲端清理待同步重試（${err.message}）`;
          }
        } else if (Sync.isOnlineEnabled()) {
          cloudMsg = '（目前離線，恢復連線同步後將一併清除雲端明細）';
        }

        await reloadState();
        showToast(
          `已刪除活動「${target?.name || eid}」及明細資料（本機清除報名 ${localPurge.deletedRegistrations} 筆、感應 ${localPurge.deletedScans} 筆${cloudMsg}）`,
          'info'
        );
      }
    });

    $('form-event-basic').addEventListener('submit', async (e) => {
      e.preventDefault();
      const newStatus = $('cfg-evt-status').value;
      const newName = $('cfg-evt-name').value.trim();
      await updateCurrentEvent((ev) => {
        const patch = {
          name: newName,
          englishName: $('cfg-evt-en-name').value.trim(),
          organization: $('cfg-evt-org').value.trim() || 'Taiwan Ultra Runners',
          date: $('cfg-evt-date').value,
          location: $('cfg-evt-location').value.trim(),
          status: newStatus,
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
      showToast(
        newStatus === 'finished'
          ? `🔒 活動「${newName}」已設為「已結束」，已自下拉選單與其他管理選項隱藏以防資料遭修改`
          : '本活動基本資訊已更新',
        newStatus === 'finished' ? 'info' : 'success'
      );
    });

    $('form-admin-credentials').addEventListener('submit', async (e) => {
      e.preventDefault();
      const oldU = getCurrentAdminUsername();
      const isSuper = isCurrentSuperAdmin();
      const newU = isSuper ? ($('cfg-admin-user').value.trim() || oldU) : oldU;
      const newDisplay = $('cfg-admin-displayname').value.trim() || newU;
      const typedPass = $('cfg-admin-pass').value;
      const existingSelf = (state.platform.adminAccounts || []).find((a) => a.username === oldU);
      // 密碼欄留空則維持原密碼，有輸入才變更
      const newP = typedPass !== '' ? typedPass : (existingSelf?.password || state.session?.auth?.password || '');
      if (!newP) return showToast('請輸入密碼', 'warn');

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

    // 最高管理員：多管理者帳號增刪改（新增帳號密碼預設空白，不可自動填入預設密碼）
    $('btn-add-admin-account').addEventListener('click', async () => {
      if (!isCurrentSuperAdmin()) return;
      const list = [...(state.platform.adminAccounts || [])];
      let n = 1;
      while (list.some((a) => a.username === `host${n}`)) n++;
      list.push({
        username: `host${n}`,
        password: '',
        displayName: `活動承辦人 ${n}`,
        isSuperAdmin: false,
        enabled: true,
        updatedAt: Date.now()
      });
      await DB.savePlatform({ ...state.platform, adminAccounts: list });
      await reloadState();
      showToast(`已新增承辦人「host${n}」，請手動輸入登入密碼後按「💾 儲存管理者」`, 'info');
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

      const existingByOrig = new Map((state.platform.adminAccounts || []).map((a) => [a.username, a]));
      for (const row of rows) {
        const origUser = row.dataset.adminOrigUser || '';
        const username = row.querySelector('[data-admin-field="username"]').value.trim();
        const typedPass = row.querySelector('[data-admin-field="password"]').value;
        const existingAcc = existingByOrig.get(origUser);
        // 既有帳號密碼欄留空時維持原密碼；新帳號未填密碼則攔截要求手動設定
        const password = typedPass !== '' ? typedPass : String(existingAcc?.password || '');
        const displayName = row.querySelector('[data-admin-field="displayName"]').value.trim() || username;
        const isSuperAdmin = row.querySelector('[data-admin-field="isSuperAdmin"]').value === '1';
        const enabled = row.querySelector('[data-admin-field="enabled"]').checked;

        if (!username) continue;
        if (!password) return showToast(`管理者「${username}」尚未設定密碼，請手動輸入密碼`, 'error');
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
        const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });
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
        categoryId: $('admin-reg-cat').value,
        gender: $('admin-reg-gender').value,
        dob: $('admin-reg-dob').value.trim(),
        city: $('admin-reg-city').value.trim(),
        country: $('admin-reg-country').value.trim(),
        lastName: $('admin-reg-lastname').value.trim(),
        firstName: $('admin-reg-firstname').value.trim()
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

    // 分頁 3：成績管理與修正（即時監控與人工修正；正式匯出已獨立至「📦 已結案專區」）
    $('results-category-filter').addEventListener('change', (e) => { state.resultsCatFilter = e.target.value; renderAdminResultsTab(); });
    $('results-search').addEventListener('input', (e) => { state.resultsSearch = e.target.value; renderAdminResultsTab(); });

    // Modal 5：DUV 成績匯出（僅針對已結案專區選取之已結束活動）
    const closeDuvModal = () => $('modal-duv-export').classList.add('hidden');
    $('btn-close-duv-modal').addEventListener('click', closeDuvModal);
    $('btn-cancel-duv-modal').addEventListener('click', closeDuvModal);
    $('form-duv-export').addEventListener('submit', async (e) => {
      e.preventDefault();
      const ctx = getSelectedExportContext();
      if (!ctx || !ctx.event || !ctx.computed) {
        return showToast('尚無已結案活動可匯出 DUV（請先將活動狀態設為「已結束」）', 'warn');
      }
      const evt = ctx.event;
      try {
        const org = $('duv-org').value.trim() || 'Taiwan Ultra Runners';
        const enName = $('duv-event-name').value.trim() || evt.name;
        const res = Excel.exportDuvResults({
          event: evt,
          computed: ctx.computed,
          options: {
            organization: org,
            eventName: enName,
            defaultCity: $('duv-default-city').value.trim(),
            defaultCountry: $('duv-default-country').value.trim() || 'Taiwan',
            format: $('duv-format').value,
            lapMetric: $('duv-lap-metric').value
          }
        });
        closeDuvModal();
        showToast(`🏅 已匯出「${evt.name}」DUV 成績（${res.totalFinishers} 位完賽）：${res.fileName}`, 'success');
      } catch (err) {
        showToast(err.message, 'error');
      }
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
      if (tsMs === null) return showToast('時間格式錯誤，請輸入 HH:mm:ss', 'error');

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
      if (tsMs === null) return showToast('時間格式錯誤，請輸入 HH:mm:ss', 'error');

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

    // 分頁 6：歷年活動及成績匯入管理
    $('btn-download-history-tpl').addEventListener('click', () => {
      try {
        const f = Excel.downloadHistoryTemplate();
        showToast(`已下載歷年成績匯入範本：${f}`, 'success');
      } catch (err) {
        showToast(err.message, 'error');
      }
    });

    $('input-import-history-file').addEventListener('change', async (e) => {
      if (!isCurrentSuperAdmin()) {
        e.target.value = '';
        return showToast('僅最高管理權責可匯入歷年活動成績', 'error');
      }
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const wb = window.XLSX.read(await file.arrayBuffer(), { type: 'array' });
        const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });
        stageHistoryImportPreview(rows, { sourceLabel: file.name, fromText: false });
      } catch (err) {
        showToast(`檔案解析失敗：${err.message}`, 'error');
      }
      e.target.value = '';
    });

    $('btn-import-history-text').addEventListener('click', () => {
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理權責可匯入歷年活動成績', 'error');
      const text = $('textarea-history-csv').value.trim();
      if (!text) return showToast('請先貼上歷年成績文字資料', 'warn');
      stageHistoryImportPreview(
        text.split(/\r?\n/).map((line) => line.split(/[,\t]/)),
        { sourceLabel: '貼上文字批次輸入', fromText: true }
      );
    });

    $('btn-cancel-history-preview').addEventListener('click', () => {
      clearHistoryImportPreview(true);
    });

    $('btn-confirm-history-import').addEventListener('click', async () => {
      await commitHistoryImportPreview();
    });

    $('btn-history-preview-all-events').addEventListener('click', () => {
      state.historyPreview.eventFilter = 'ALL';
      renderHistoryImportPreview();
    });

    $('history-preview-search').addEventListener('input', (e) => {
      state.historyPreview.search = e.target.value;
      renderHistoryImportPreview();
    });

    $('history-preview-events-list').addEventListener('click', (e) => {
      const rmBtn = e.target.closest('[data-remove-preview-evt]');
      if (rmBtn) {
        const key = rmBtn.dataset.removePreviewEvt;
        state.historyPreview.items = state.historyPreview.items.filter((x) => `${x.eventDate}__${x.eventName}` !== key);
        if (state.historyPreview.eventFilter === key) state.historyPreview.eventFilter = 'ALL';
        renderHistoryImportPreview();
        showToast('已從預覽中剔除該場次（尚未寫入資料庫）', 'info');
        return;
      }
      const filterBtn = e.target.closest('[data-preview-evt-filter]');
      if (filterBtn) {
        const key = filterBtn.dataset.previewEvtFilter;
        state.historyPreview.eventFilter = state.historyPreview.eventFilter === key ? 'ALL' : key;
        renderHistoryImportPreview();
      }
    });

    $('history-preview-tbody').addEventListener('click', (e) => {
      const rmBtn = e.target.closest('[data-remove-preview-id]');
      if (!rmBtn) return;
      const pid = rmBtn.dataset.removePreviewId;
      state.historyPreview.items = state.historyPreview.items.filter((x) => x._previewId !== pid);
      renderHistoryImportPreview();
    });

    $('admin-history-filter-event').addEventListener('change', (e) => {
      state.adminHistoryEventFilter = e.target.value;
      renderAdminHistoryTab();
    });

    $('admin-history-search').addEventListener('input', (e) => {
      state.adminHistorySearch = e.target.value;
      renderAdminHistoryTab();
    });

    $('admin-history-events-list').addEventListener('click', async (e) => {
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理權責可刪除歷年活動成績', 'error');
      const btn = e.target.closest('[data-del-history-evt-name]');
      if (!btn) return;
      const evDate = btn.dataset.delHistoryEvtDate;
      const evName = btn.dataset.delHistoryEvtName;
      if (!confirm(`確定刪除歷年活動「${evDate} ${evName}」的全部匯入成績嗎？`)) return;
      const deletedIds = await DB.deleteHistoryEvent(evDate, evName);
      if (deletedIds.length && Sync.isOnlineEnabled() && navigator.onLine) {
        Sync.deleteHistoryRecordsCloud(deletedIds, state.session?.auth).catch(() => {});
      }
      await reloadState();
      showToast(`已刪除歷年活動「${evName}」（${deletedIds.length} 筆）`, 'info');
    });

    $('admin-history-tbody').addEventListener('click', async (e) => {
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理權責可刪除歷年活動成績', 'error');
      const btn = e.target.closest('[data-del-history-id]');
      if (!btn || !confirm('確定刪除此筆歷年成績紀錄嗎？')) return;
      const id = btn.dataset.delHistoryId;
      await DB.deleteHistoryRecord(id);
      if (Sync.isOnlineEnabled() && navigator.onLine) {
        Sync.deleteHistoryRecordsCloud([id], state.session?.auth).catch(() => {});
      }
      await reloadState();
      showToast('已刪除此筆歷年成績', 'info');
    });

    // 分頁 7：🏅 製作證書（僅包含已結束與歷年活動；官方中英文對照公版 ＋ 承辦人上傳自訂底圖套印）
    $('cert-event-select').addEventListener('change', (e) => {
      state.cert.selectedSourceKey = e.target.value;
      renderAdminCertificateTab();
    });

    document.querySelectorAll('.cert-mode-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.cert.mode = btn.dataset.certMode === 'custom' ? 'custom' : 'official';
        renderAdminCertificateTab();
      });
    });

    $('cert-cat-filter').addEventListener('change', (e) => {
      state.cert.catFilter = e.target.value;
      renderAdminCertificateTab();
    });
    $('cert-runner-search').addEventListener('input', (e) => {
      state.cert.search = e.target.value;
      renderAdminCertificateTab();
    });
    $('cert-only-finished').addEventListener('change', (e) => {
      state.cert.onlyFinished = e.target.checked;
      renderAdminCertificateTab();
    });
    $('cert-runner-list').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-cert-bib]');
      if (!btn) return;
      state.cert.selectedBib = btn.dataset.certBib;
      renderAdminCertificateTab();
    });

    // 官方公版選項即時預覽
    $('cert-orientation').addEventListener('change', (e) => {
      state.cert.orientation = e.target.value;
      renderCertificatePreview();
    });
    $('cert-format')?.addEventListener('change', (e) => {
      state.cert.format = e.target.value;
      renderCertificatePreview();
    });
    $('cert-theme').addEventListener('change', (e) => {
      state.cert.theme = e.target.value;
      renderCertificatePreview();
    });
    $('cert-en-name').addEventListener('input', (e) => {
      state.cert.englishName = e.target.value;
      renderCertificatePreview();
    });
    $('cert-org-name').addEventListener('input', (e) => {
      state.cert.organization = e.target.value;
      renderCertificatePreview();
    });
    $('cert-show-splits').addEventListener('change', (e) => {
      state.cert.showSplits = e.target.checked;
      renderCertificatePreview();
    });

    // 自訂底圖上傳與清除
    $('input-cert-bg-file').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const dataUrl = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result);
          reader.onerror = () => reject(new Error('圖片讀取失敗'));
          reader.readAsDataURL(file);
        });
        const img = await Cert.loadImage(dataUrl);
        if (!img) throw new Error('不支援的圖片格式');
        state.cert.bgDataUrl = dataUrl;
        state.cert.bgImage = img;
        $('cert-bg-status').textContent = `✅ 已上傳底圖：${file.name} (${img.naturalWidth}×${img.naturalHeight})`;
        $('btn-cert-clear-bg').classList.remove('hidden');
        await renderCertificatePreview();
        showToast('已載入自訂證書底圖！可直接於預覽圖拖曳文字對齊位置', 'success');
      } catch (err) {
        showToast(err.message, 'error');
      }
      e.target.value = '';
    });

    $('btn-cert-clear-bg').addEventListener('click', () => {
      state.cert.bgDataUrl = '';
      state.cert.bgImage = null;
      $('cert-bg-status').textContent = '尚未上傳底圖（支援 A4 橫式或直式高解析圖檔）';
      $('btn-cert-clear-bg').classList.add('hidden');
      renderCertificatePreview();
    });

    $('cert-custom-show-logo').addEventListener('change', (e) => {
      state.cert.showTurLogo = e.target.checked;
      renderCertificatePreview();
    });

    // 自訂欄位勾選與選取
    $('cert-field-list').addEventListener('change', (e) => {
      const chk = e.target.closest('[data-cert-field-toggle]');
      if (!chk) return;
      const key = chk.dataset.certFieldToggle;
      const item = state.cert.layout.find((f) => f.key === key);
      if (item) {
        item.enabled = chk.checked;
        state.cert.selectedFieldKey = key;
        renderCertFieldChips();
        syncCertFieldEditor();
        renderCertificatePreview();
      }
    });

    $('cert-field-list').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-cert-field-select]');
      if (!btn) return;
      state.cert.selectedFieldKey = btn.dataset.certFieldSelect;
      renderCertFieldChips();
      syncCertFieldEditor();
      renderCertificatePreview();
    });

    const updateSelectedCertField = (patch) => {
      const item = state.cert.layout.find((f) => f.key === state.cert.selectedFieldKey);
      if (!item) return;
      Object.assign(item, patch);
      renderCertificatePreview();
    };
    $('cert-f-x').addEventListener('input', (e) => updateSelectedCertField({ x: Math.max(0, Math.min(100, Number(e.target.value || 0))) }));
    $('cert-f-y').addEventListener('input', (e) => updateSelectedCertField({ y: Math.max(0, Math.min(100, Number(e.target.value || 0))) }));
    $('cert-f-size').addEventListener('input', (e) => updateSelectedCertField({ fontSize: Math.max(12, Math.min(260, Number(e.target.value || 48))) }));
    $('cert-f-color').addEventListener('input', (e) => updateSelectedCertField({ color: e.target.value }));
    $('cert-f-prefix').addEventListener('input', (e) => updateSelectedCertField({ prefix: e.target.value }));
    $('cert-f-align').addEventListener('change', (e) => updateSelectedCertField({ align: e.target.value }));
    $('cert-f-bold').addEventListener('change', (e) => updateSelectedCertField({ bold: e.target.checked }));

    $('btn-cert-save-layout').addEventListener('click', () => {
      const evt = getAdminCertContext()?.event;
      if (!evt || !Cert) return;
      const savedAll = Cert.saveCustomConfig(evt.eventId, {
        bgDataUrl: state.cert.bgDataUrl,
        showTurLogo: state.cert.showTurLogo,
        layout: state.cert.layout
      });
      showToast(
        savedAll ? '💾 已儲存此活動的自訂底圖與座標版面！' : '💾 已儲存座標版面（底圖檔案較大，重新開啟分頁時請再上傳底圖）',
        'success'
      );
    });

    $('btn-cert-reset-layout').addEventListener('click', () => {
      if (!Cert) return;
      state.cert.layout = Cert.getDefaultCustomLayout();
      state.cert.selectedFieldKey = 'name';
      renderCertFieldChips();
      syncCertFieldEditor();
      renderCertificatePreview();
      showToast('已重設自訂底圖欄位至預設座標', 'info');
    });

    // 自訂底圖模式：預覽畫布直接拖曳定位（滑鼠與觸控皆支援）
    let certDragState = null;
    const getCanvasCoords = (clientX, clientY) => {
      const canvas = $('cert-preview-canvas');
      const rect = canvas.getBoundingClientRect();
      return {
        x: ((clientX - rect.left) / rect.width) * canvas.width,
        y: ((clientY - rect.top) / rect.height) * canvas.height,
        w: canvas.width,
        h: canvas.height
      };
    };
    const onCertPointerDown = (clientX, clientY) => {
      if (state.cert.mode !== 'custom') return;
      const pt = getCanvasCoords(clientX, clientY);
      const hit = [...state.cert.hitBoxes].reverse().find(
        (b) => pt.x >= b.x && pt.x <= b.x + b.w && pt.y >= b.y && pt.y <= b.y + b.h
      );
      if (!hit) return;
      const item = state.cert.layout.find((f) => f.key === hit.key);
      if (!item) return;
      state.cert.selectedFieldKey = item.key;
      certDragState = {
        key: item.key,
        startX: pt.x,
        startY: pt.y,
        origPctX: Number(item.x || 50),
        origPctY: Number(item.y || 50),
        w: pt.w,
        h: pt.h
      };
      renderCertFieldChips();
      syncCertFieldEditor();
      renderCertificatePreview();
    };
    const onCertPointerMove = (clientX, clientY) => {
      if (!certDragState) return;
      const pt = getCanvasCoords(clientX, clientY);
      const dxPct = ((pt.x - certDragState.startX) / certDragState.w) * 100;
      const dyPct = ((pt.y - certDragState.startY) / certDragState.h) * 100;
      const item = state.cert.layout.find((f) => f.key === certDragState.key);
      if (!item) return;
      item.x = Number(Math.max(0, Math.min(100, certDragState.origPctX + dxPct)).toFixed(1));
      item.y = Number(Math.max(0, Math.min(100, certDragState.origPctY + dyPct)).toFixed(1));
      syncCertFieldEditor();
      renderCertificatePreview();
    };
    const onCertPointerUp = () => { certDragState = null; };

    const certCanvasEl = $('cert-preview-canvas');
    certCanvasEl.addEventListener('mousedown', (e) => onCertPointerDown(e.clientX, e.clientY));
    window.addEventListener('mousemove', (e) => onCertPointerMove(e.clientX, e.clientY));
    window.addEventListener('mouseup', onCertPointerUp);
    certCanvasEl.addEventListener('touchstart', (e) => {
      if (state.cert.mode === 'custom' && e.touches[0]) {
        onCertPointerDown(e.touches[0].clientX, e.touches[0].clientY);
        if (certDragState) e.preventDefault();
      }
    }, { passive: false });
    window.addEventListener('touchmove', (e) => {
      if (certDragState && e.touches[0]) {
        onCertPointerMove(e.touches[0].clientX, e.touches[0].clientY);
        e.preventDefault();
      }
    }, { passive: false });
    window.addEventListener('touchend', onCertPointerUp);

    // 單張下載（PNG 圖檔 / PDF 檔）與批次合併下載 PDF
    const getSelectedCertRow = () => {
      const rows = getCertFilteredRows();
      return rows.find((r) => r.bib === state.cert.selectedBib) || rows[0] || null;
    };

    $('btn-cert-download-png').addEventListener('click', async () => {
      const evt = getAdminCertContext()?.event;
      const row = getSelectedCertRow();
      if (!evt || !row) return showToast('請先選取一位有成績的選手', 'warn');
      if (state.cert.mode === 'custom' && !state.cert.bgImage) {
        return showToast('自訂底圖模式請先上傳底圖後再下載', 'warn');
      }
      try {
        const offscreen = document.createElement('canvas');
        await renderCertToCanvas(offscreen, row, false);
        const safeName = `${evt.name}_${row.bib}_${row.name}`.replace(/[\\/:*?"<>|]/g, '_');
        await Cert.downloadPng(offscreen, `完賽證書_${safeName}.png`);
        showToast(`🖼️ 已下載證書圖檔：${row.name} (#${row.bib})`, 'success');
      } catch (err) {
        showToast(`圖檔產生失敗：${err.message}`, 'error');
      }
    });

    $('btn-cert-download-pdf').addEventListener('click', async () => {
      const evt = getAdminCertContext()?.event;
      const row = getSelectedCertRow();
      if (!evt || !row) return showToast('請先選取一位有成績的選手', 'warn');
      if (state.cert.mode === 'custom' && !state.cert.bgImage) {
        return showToast('自訂底圖模式請先上傳底圖後再下載', 'warn');
      }
      try {
        const offscreen = document.createElement('canvas');
        await renderCertToCanvas(offscreen, row, false);
        const safeName = `${evt.name}_${row.bib}_${row.name}`.replace(/[\\/:*?"<>|]/g, '_');
        Cert.downloadPdfFromCanvas(offscreen, `完賽證書_${safeName}.pdf`);
        showToast(`📄 已下載證書 PDF：${row.name} (#${row.bib})`, 'success');
      } catch (err) {
        showToast(`PDF 產生失敗：${err.message}`, 'error');
      }
    });

    $('btn-cert-batch-pdf').addEventListener('click', async () => {
      const evt = getAdminCertContext()?.event;
      const rows = getCertFilteredRows();
      if (!evt || !rows.length) return showToast('目前篩選清單內無選手可批次產製證書', 'warn');
      if (state.cert.mode === 'custom' && !state.cert.bgImage) {
        return showToast('自訂底圖模式請先上傳底圖後再批次下載', 'warn');
      }
      if (rows.length > 20 && !confirm(`確定要一次批次產製並合併下載 ${rows.length} 張證書為單一 PDF 嗎？`)) return;
      const btn = $('btn-cert-batch-pdf');
      const origHtml = btn.innerHTML;
      btn.disabled = true;
      try {
        const offscreen = document.createElement('canvas');
        const pages = [];
        for (let i = 0; i < rows.length; i++) {
          btn.textContent = `⏳ 產製中 (${i + 1}/${rows.length})...`;
          await renderCertToCanvas(offscreen, rows[i], false);
          pages.push(Cert.canvasToJpegPage(offscreen, 0.9));
        }
        const safeEvt = (evt.name || 'TUR').replace(/[\\/:*?"<>|]/g, '_');
        const catTag = state.cert.catFilter === 'ALL' ? '全部' : state.cert.catFilter;
        Cert.downloadPdfFromPages(pages, `${safeEvt}_完賽證書_${catTag}_${rows.length}張.pdf`);
        showToast(`📚 已批次產生並下載 ${rows.length} 頁完賽證書 PDF！`, 'success');
      } catch (err) {
        showToast(`批次 PDF 產生失敗：${err.message}`, 'error');
      } finally {
        btn.disabled = false;
        btn.innerHTML = origHtml;
        if ($('cert-batch-count')) $('cert-batch-count').textContent = String(rows.length);
      }
    });

    // 專區 2：📦 已結案專區（選擇已結案活動即自動展開匯出正式 Excel / DUV）
    $('export-event-select').addEventListener('change', (e) => {
      state.exportSelectedEventId = e.target.value;
      if (state.adminTab !== 'export') {
        switchAdminTab('export');
      } else {
        renderAdminExportTab();
      }
    });
    $('btn-export-excel-full').addEventListener('click', () => {
      const ctx = getSelectedExportContext();
      if (!ctx || !ctx.event || !ctx.computed) {
        return showToast('尚無已結案活動可匯出（請先將活動狀態設為「已結束」）', 'warn');
      }
      const f = Excel.exportResultsToExcel({ event: ctx.event, computed: ctx.computed, scans: ctx.scans, maskId: false });
      showToast(`已匯出「${ctx.event.name}」完整 Excel：${f}`, 'success');
    });
    $('btn-export-excel-masked').addEventListener('click', () => {
      const ctx = getSelectedExportContext();
      if (!ctx || !ctx.event || !ctx.computed) {
        return showToast('尚無已結案活動可匯出（請先將活動狀態設為「已結束」）', 'warn');
      }
      const f = Excel.exportResultsToExcel({ event: ctx.event, computed: ctx.computed, scans: ctx.scans, maskId: true });
      showToast(`已匯出「${ctx.event.name}」隱碼公開 Excel：${f}`, 'success');
    });
    $('btn-export-duv').addEventListener('click', openDuvExportModal);

    $('btn-backup-download').addEventListener('click', () => {
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理權責可執行資料庫備份', 'error');
      downloadJsonBackup(`TUR完整備份_${getCurrentAdminUsername() || 'all'}`);
    });
    document.querySelectorAll('input[name="backup-restore-mode"]').forEach((r) => {
      r.addEventListener('change', updateBackupRestoreModeUi);
    });
    $('input-backup-file').addEventListener('change', async (e) => {
      if (!isCurrentSuperAdmin()) {
        e.target.value = '';
        return showToast('僅最高管理權責可執行資料庫還原', 'error');
      }
      const f = e.target.files?.[0];
      if (!f) return;
      const strategy = document.querySelector('input[name="backup-restore-mode"]:checked')?.value || 'merge';
      if (strategy === 'replace' && !confirm(`確定以「${f.name}」覆蓋還原整個本機資料庫嗎？\n（目前本機資料將先清空再還原）`)) {
        e.target.value = '';
        return;
      }
      try {
        const text = await f.text();
        let parsed;
        try {
          parsed = JSON.parse(text);
        } catch {
          throw new Error('JSON 檔案格式錯誤，無法解析');
        }
        const r = await DB.importFullBackup(parsed, strategy, {
          ownerAdmin: getCurrentAdminUsername(),
          isSuperAdmin: true
        });
        await reloadState();
        const modeLabel = strategy === 'replace' ? '覆蓋還原' : '合併還原';
        const msg = `${modeLabel}成功（${f.name}）！活動 ${r.eventCount} 場、晶片 ${r.chipCount}、報名 ${r.regCount}、感應 ${r.scanCount}、歷年 ${r.historyCount || 0} 筆`;
        showToast(`✅ ${msg}`, 'success');
        showBackupRestoreResult(`✅ ${msg}`, true);
      } catch (err) {
        showToast(`還原失敗：${err.message}`, 'error');
        showBackupRestoreResult(`❌ 還原失敗：${err.message}`, false);
      } finally {
        e.target.value = '';
      }
    });

    document.querySelectorAll('input[name="sync-mode"]').forEach((r) => {
      r.addEventListener('change', (e) => {
        const nextMode = e.target.value;
        DB.saveDeviceSettings({ mode: nextMode });
        updateNetworkIndicators();
        renderAdminSyncTab();
        showToast(
          nextMode === 'online'
            ? '🌐 已切換為「雲端連線模式」（可立即點「🔄 立即雙向同步」對齊資料）'
            : '📴 已切換為「單機離線模式」（後續異動僅存本機，不自動上傳）',
          'info'
        );
      });
    });

    $('btn-save-gas-url').addEventListener('click', async () => {
      const url = $('input-gas-url').value.trim();
      DB.saveDeviceSettings({ gasUrl: url, mode: url ? 'online' : 'offline' });
      renderHeader();
      updateNetworkIndicators();
      renderAdminSyncTab();
      if (!url) return showToast('已清除 GAS 網址並切回單機離線模式', 'info');
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
      if (!isCurrentSuperAdmin()) return showToast('僅最高管理員可載入測試資料', 'error');
      if (!confirm('載入「全情境測試資料」將會覆蓋目前本機資料，建立 5 場不同型態與狀態的活動（進行中超馬距離賽、進行中繞圈賽含逾時圈、籌備中賽事、2 場已結束歸檔賽事）及 2023–2025 八百壯士歷年紀錄，確定繼續嗎？')) return;
      await DB.loadDemoData(getCurrentAdminUsername());
      state.selectedEventId = 'EVT-2026-800-EAST';
      await reloadState();
      showToast('⚡ 已載入 5 場全情境測試活動＋2023–2025 歷年紀錄！可一次完整檢視所有模組狀態', 'success');
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
